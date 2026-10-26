#!/usr/bin/env python3
"""
tex2docx_v2.py - Two-way converter: LaTeX (.tex) <-> Word (.docx) using Pandoc.
The direction is chosen automatically from the input file's extension:
    paper.tex  -> paper.docx
    paper.docx -> paper_from_docx.tex  (+ paper_from_docx_media/ with images)

What it fixes compared with a plain `pandoc paper.tex -o paper.docx`:
  * PDF figures  -> converted to PNG (Word cannot embed PDF) and re-linked
  * References   -> auto-detects \\bibliography{...}, uses the .bib file + citeproc
  * \\ref/\\autoref/\\eqref -> replaced with real numbers read from the .aux file
  * Figure/Table captions   -> prefixed with "Fig. N." / "Table N." (from .aux)
  * Original .tex is never modified (a temporary copy is processed)

Install:
    pip install pymupdf pypandoc_binary
  (or install Pandoc separately from https://pandoc.org/installing.html)

Usage (run from the folder containing the .tex, .bib, .aux and figures):
    python tex2docx_v2.py Islam-etal2025Natural_Hazards20261024.tex
    python tex2docx_v2.py paper.tex -o paper.docx --csl springer-basic-author-date.csl
    python tex2docx_v2.py paper.tex --reference-doc template.docx --dpi 250
    python tex2docx_v2.py paper.docx              # Word -> LaTeX
"""

import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path

IMG_EXTS = ["", ".pdf", ".png", ".jpg", ".jpeg", ".PNG", ".JPG", ".PDF", ".tif", ".tiff"]


# --------------------------------------------------------------------------- #
# Tools
# --------------------------------------------------------------------------- #
def find_pandoc():
    exe = shutil.which("pandoc")
    if exe:
        return exe
    try:
        import pypandoc
        return pypandoc.get_pandoc_path()
    except Exception:
        return None


def pdf_to_png(pdf: Path, png: Path, dpi: int):
    """Render first page of a PDF to PNG. Uses PyMuPDF, falls back to pdftoppm."""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(str(pdf))
        doc[0].get_pixmap(dpi=dpi).save(str(png))
        doc.close()
        return
    except ImportError:
        pass
    if shutil.which("pdftoppm"):
        base = str(png.with_suffix(""))
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-singlefile",
                        str(pdf), base], check=True)
        return
    raise RuntimeError("Cannot convert PDF figures. Run:  pip install pymupdf")


# --------------------------------------------------------------------------- #
# .aux parsing  (label -> number / type)
# --------------------------------------------------------------------------- #
def read_labels(aux: Path):
    labels = {}
    if not aux.is_file():
        print(f"  [warn] {aux.name} not found - \\ref numbers/captions won't be resolved. "
              f"Compile the paper once with LaTeX to create it.")
        return labels
    text = aux.read_text(encoding="utf-8", errors="ignore")
    # \newlabel{fig:x}{{1}{5}{Caption...}{figure.1}{}}
    for line in text.splitlines():
        m = re.match(r"\\newlabel\{([^}]+)\}\{\{([^{}]*)\}", line)
        if not m:
            continue
        name, num = m.group(1), m.group(2)
        kind = ""
        k = re.search(r"\{([A-Za-z*]+)\.[^{}]*\}\{\}\}\s*$", line)
        if k:
            kind = k.group(1).lower()
        labels[name] = (num, kind)
    return labels


KIND_NAME = {"figure": "Fig.", "table": "Table", "section": "Section",
             "subsection": "Section", "subsubsection": "Section",
             "appendix": "Appendix", "equation": "Eq."}


# --------------------------------------------------------------------------- #
# Front matter (Springer Nature sn-jnl: \abstract, \author, \affil, \keywords)
# --------------------------------------------------------------------------- #
def strip_comments(s: str) -> str:
    """Remove LaTeX % comments (keeps escaped \\%)."""
    return "\n".join(re.sub(r"(?<!\\)%.*$", "", ln) for ln in s.split("\n"))


def read_group(s: str, i: int):
    """s[i] must be '{'. Return (content, index just after the matching '}')."""
    depth, j = 0, i
    while j < len(s):
        c = s[j]
        if c == "\\":
            j += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    return s[i + 1:], len(s)


def find_cmds(s: str, name: str):
    """Find \\name*[opt]{arg} occurrences; returns list of dicts with spans."""
    out = []
    for m in re.finditer(r"\\" + name + r"(?![A-Za-z])", s):
        i = m.end()
        star = i < len(s) and s[i] == "*"
        if star:
            i += 1
        while i < len(s) and s[i] in " \t":
            i += 1
        opt = None
        if i < len(s) and s[i] == "[":
            k = s.find("]", i)
            if k != -1:
                opt, i = s[i + 1:k], k + 1
        while i < len(s) and s[i] in " \t":
            i += 1
        if i < len(s) and s[i] == "{":
            arg, end = read_group(s, i)
            out.append(dict(start=m.start(), end=end, star=star, opt=opt, arg=arg))
    return out


def _field(content: str, name: str) -> str:
    f = find_cmds(content, name)
    return f[0]["arg"].strip() if f else ""


def convert_front_matter(src: str) -> str:
    """Turn sn-jnl front matter into plain LaTeX that Pandoc understands."""
    titles = find_cmds(src, "title")
    authors = find_cmds(src, "author")
    emails = find_cmds(src, "email")
    affils = find_cmds(src, "affil")
    abstracts = find_cmds(src, "abstract")
    keywords = find_cmds(src, "keywords")
    if not (authors or abstracts or keywords):
        return src

    # authors
    names, corr_email = [], ""
    for a in authors:
        parts = [_field(a["arg"], n) for n in ("fnm", "spfx", "sur", "sfx")]
        name = " ".join(p for p in parts if p) or a["arg"].strip()
        marks = (a["opt"] or "").replace(" ", "")
        if a["star"]:
            marks = (marks + ",*") if marks else "*"
        names.append(f"{name}\\textsuperscript{{{marks}}}" if marks else name)
        if a["star"]:
            em = [e for e in emails if e["start"] >= a["end"]]
            if em:
                corr_email = min(em, key=lambda e: e["start"])["arg"].strip()

    # affiliations
    aff_lines = []
    for af in affils:
        bits = [_field(af["arg"], n) for n in
                ("orgdiv", "orgname", "street", "city", "state", "postcode", "country")]
        text = ", ".join(b for b in bits if b) or af["arg"].strip()
        num = (af["opt"] or "").strip()
        aff_lines.append(f"\\textsuperscript{{{num}}}{text}" if num else text)

    block = []
    if names:
        block.append("\\textbf{" + ", ".join(names) + "}")
    block += aff_lines
    if corr_email:
        block.append(f"\\textsuperscript{{*}}Corresponding author: {corr_email}")
    if abstracts:
        block.append("\\section*{Abstract}")
        block.append(abstracts[0]["arg"].strip())
    if keywords:
        block.append("\\textbf{Keywords:} " + keywords[0]["arg"].strip())
    front = "\n\n".join(block) + "\n\n"

    # collect edits (start, end, replacement) and apply from the end
    edits = []
    for t in titles:
        edits.append((t["start"], t["end"], "\\title{" + t["arg"] + "}"))
    for group in (authors, emails, affils, abstracts, keywords):
        for g in group:
            edits.append((g["start"], g["end"], ""))
    mt = re.search(r"\\maketitle", src)
    if mt:
        edits.append((mt.start(), mt.end(), front))
    else:
        bd = re.search(r"\\begin\{document\}", src)
        if bd:
            edits.append((bd.end(), bd.end(), "\n" + front))
    for s, e, rep in sorted(edits, key=lambda x: x[0], reverse=True):
        src = src[:s] + rep + src[e:]
    return src


# --------------------------------------------------------------------------- #
# Preprocessing
# --------------------------------------------------------------------------- #
def preprocess(tex_path: Path, labels: dict, dpi: int):
    src = tex_path.read_text(encoding="utf-8", errors="ignore")
    base = tex_path.parent

    # Drop comments (also removes commented-out sample abstracts/titles)
    src = strip_comments(src)
    # Springer-specific front matter and headings
    src = convert_front_matter(src)
    src = re.sub(r"\\bmhead\{", r"\\section*{", src)
    src = re.sub(r"\\backmatter\b", "", src)
    src = re.sub(r"\\botrule\b", r"\\bottomrule", src)
    fig_dir = base / "_figs_png"

    # Search dirs from \graphicspath{{a/}{b/}}
    search_dirs = [base]
    gp = re.search(r"\\graphicspath\{((?:\{[^}]*\})+)\}", src)
    if gp:
        for d in re.findall(r"\{([^}]*)\}", gp.group(1)):
            search_dirs.append(base / d)

    missing = []

    def fix_graphic(m):
        opts, name = m.group(1) or "", m.group(2).strip()
        found = None
        for d in search_dirs:
            for ext in IMG_EXTS:
                cand = d / (name + ext)
                if cand.is_file():
                    found = cand
                    break
            if found:
                break
        if not found:
            missing.append(name)
            return m.group(0)
        if found.suffix.lower() == ".pdf":
            fig_dir.mkdir(exist_ok=True)
            safe = re.sub(r"[^A-Za-z0-9_-]", "_", found.stem)
            png = fig_dir / f"{safe}.png"
            if not png.exists() or png.stat().st_mtime < found.stat().st_mtime:
                print(f"  figure: {found.name} -> {png.name}")
                pdf_to_png(found, png, dpi)
            found = png
        rel = found.relative_to(base).as_posix()
        return f"\\includegraphics{opts}{{{rel}}}"

    src = re.sub(r"\\includegraphics(\[[^\]]*\])?\{([^}]+)\}", fix_graphic, src)
    if missing:
        print("  [warn] figures not found: " + ", ".join(missing))

    # Prefix captions with "Fig. N." / "Table N."
    def fix_float(m):
        env, block = m.group(1), m.group(0)
        lab = re.search(r"\\label\{([^}]+)\}", block)
        if not lab or lab.group(1) not in labels:
            return block
        num = labels[lab.group(1)][0]
        word = "Fig." if env.startswith("figure") else "Table"
        return re.sub(r"\\caption(\[[^\]]*\])?\{",
                      lambda c: f"\\caption{c.group(1) or ''}{{\\textbf{{{word} {num}.}} ",
                      block, count=1)

    src = re.sub(r"\\begin\{(figure\*?|table\*?)\}.*?\\end\{\1\}", fix_float,
                 src, flags=re.DOTALL)

    # Resolve cross-references using .aux numbers
    def fix_ref(m):
        cmd, keys = m.group(1), m.group(2)
        out = []
        for key in [k.strip() for k in keys.split(",")]:
            if key not in labels:
                out.append("??")
                continue
            num, kind = labels[key]
            if cmd == "eqref":
                out.append(f"({num})")
            elif cmd in ("autoref", "cref", "Cref"):
                out.append(f"{KIND_NAME.get(kind, '')} {num}".strip())
            else:
                out.append(num)
        return ", ".join(out)

    src = re.sub(r"\\(ref|autoref|eqref|cref|Cref|pageref)\{([^}]+)\}", fix_ref, src)

    # Bibliography file(s)
    bibs = []
    for m in re.finditer(r"\\(?:bibliography|addbibresource)\{([^}]+)\}", src):
        for b in m.group(1).split(","):
            b = b.strip()
            p = base / (b if b.endswith(".bib") else b + ".bib")
            if p.is_file():
                bibs.append(p)
            else:
                print(f"  [warn] bibliography file not found: {p.name}")

    tmp = base / "_pandoc_tmp.tex"
    tmp.write_text(src, encoding="utf-8")
    return tmp, bibs


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def docx_to_tex(src: Path, out: Path, pandoc: str, extra):
    """Word (.docx) -> LaTeX. Images are extracted next to the output file."""
    media = out.stem + "_media"
    cmd = [pandoc, str(src), "--from", "docx", "--to", "latex", "--standalone",
           "--output", str(out), f"--extract-media={media}",
           "--track-changes=accept"] + list(extra)
    r = subprocess.run(cmd, cwd=out.parent, capture_output=True, text=True)
    if r.stderr.strip():
        print("  pandoc messages:\n    " + r.stderr.strip().replace("\n", "\n    "))
    if r.returncode != 0:
        sys.exit("Conversion FAILED (see messages above)")
    print(f"Done -> {out}")
    print(f"  images (if any) -> {out.parent / media}")


def main():
    ap = argparse.ArgumentParser(
        description="Convert .tex -> .docx or .docx -> .tex (direction chosen by file extension)")
    ap.add_argument("tex", metavar="input", help="input file: .tex or .docx")
    ap.add_argument("-o", "--output", help="output path (default: same name, other extension)")
    ap.add_argument("--bib", help="override .bib file (default: auto-detect)")
    ap.add_argument("--csl", help="CSL citation style file (default: Chicago author-date)")
    ap.add_argument("--reference-doc", help="Word template for styles")
    ap.add_argument("--dpi", type=int, default=200, help="PDF figure resolution (default 200)")
    ap.add_argument("--toc", action="store_true")
    ap.add_argument("--keep-temp", action="store_true", help="keep temp .tex for debugging")
    args, extra = ap.parse_known_args()

    pandoc = find_pandoc()
    if not pandoc:
        sys.exit("Pandoc not found. Run: pip install pypandoc_binary  (or install from pandoc.org)")

    tex = Path(args.tex).resolve()
    if not tex.is_file():
        sys.exit(f"File not found: {tex}")

    if tex.suffix.lower() == ".docx":
        # Default name avoids overwriting an existing original .tex
        out = (Path(args.output).resolve() if args.output
               else tex.with_name(tex.stem + "_from_docx.tex"))
        print(f"Reading {tex.name} (Word -> LaTeX)")
        docx_to_tex(tex, out, pandoc, extra)
        return
    if tex.suffix.lower() != ".tex":
        sys.exit("Input must be a .tex or .docx file")

    out = Path(args.output).resolve() if args.output else tex.with_suffix(".docx")

    print(f"Reading {tex.name}")
    labels = read_labels(tex.with_suffix(".aux"))
    tmp, bibs = preprocess(tex, labels, args.dpi)
    if args.bib:
        bibs = [Path(args.bib).resolve()]

    cmd = [pandoc, tmp.name, "--from", "latex", "--to", "docx", "--output", str(out),
           "--standalone", "--resource-path", str(tex.parent),
           "--metadata", "reference-section-title=References"]
    if bibs:
        cmd.append("--citeproc")
        for b in bibs:
            cmd += ["--bibliography", str(b)]
        print("  bibliography: " + ", ".join(b.name for b in bibs))
    else:
        print("  [warn] no .bib found - citations will not be formatted")
    if args.csl:
        cmd += ["--csl", str(Path(args.csl).resolve())]
    if args.reference_doc:
        cmd += ["--reference-doc", str(Path(args.reference_doc).resolve())]
    if args.toc:
        cmd.append("--toc")
    cmd += extra

    try:
        r = subprocess.run(cmd, cwd=tex.parent, capture_output=True, text=True)
    finally:
        if not args.keep_temp and tmp.exists():
            tmp.unlink()

    if r.stderr.strip():
        print("  pandoc messages:\n    " + r.stderr.strip().replace("\n", "\n    "))
    if r.returncode != 0:
        sys.exit("Conversion FAILED (see messages above)")
    print(f"Done -> {out}")


if __name__ == "__main__":
    main()
