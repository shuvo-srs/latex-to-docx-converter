# LaTeX to Docx Converter

A small Python command-line tool that converts **LaTeX (`.tex`) to Word (`.docx`)** and **Word back to LaTeX**, built for scientific papers. It wraps [Pandoc](https://pandoc.org) and adds the preprocessing that a plain `pandoc paper.tex -o paper.docx` misses: PDF figures, bibliographies, cross-references and journal front matter.

The direction is chosen automatically from the input file's extension.

```
python tex2docx_v2.py paper.tex     ->  paper.docx
python tex2docx_v2.py paper.docx    ->  paper_from_docx.tex  (+ paper_from_docx_media/)
```

## Features

**LaTeX to Word**

- **PDF figures** are rendered to PNG (Word cannot embed PDFs) and re-linked automatically. Figure files with spaces or dots in their names are handled.
- **References**: finds `\bibliography{...}` / `\addbibresource{...}`, loads the `.bib` file and formats citations and the reference list with Pandoc's citeproc.
- **Cross-references**: `\ref`, `\autoref`, `\eqref`, `\cref` are replaced with the real numbers read from your compiled `.aux` file.
- **Captions** are prefixed with "Fig. N." / "Table N." using the same `.aux` numbers.
- **Springer Nature (`sn-jnl`) front matter**: `\title`, `\author`, `\affil`, `\email`, `\abstract` and `\keywords` are converted so the title, authors, affiliations, abstract and keywords appear in the Word file. `\bmhead{...}` becomes an unnumbered heading.
- Commented-out LaTeX (including sample abstracts left in templates) is ignored.
- Your original `.tex` is **never modified**. A temporary copy is processed and then deleted.

**Word to LaTeX**

- Headings, paragraphs, tables, lists and equations are converted; images are extracted to a `<name>_media/` folder and linked.
- Tracked changes are accepted.
- The output is named `<name>_from_docx.tex` so it never overwrites an existing `.tex`.

## Requirements

- Python 3.8+
- [Pandoc](https://pandoc.org/installing.html) (or install the bundled binary with pip, below)
- [PyMuPDF](https://pymupdf.readthedocs.io) for converting PDF figures (only needed for `.tex` to `.docx` with PDF figures)

```bash
pip install pymupdf pypandoc_binary
```

`pypandoc_binary` bundles Pandoc, so you can skip a separate Pandoc install. If `pymupdf` is not available, the script falls back to `pdftoppm` (from Poppler) when it is on your PATH.

## Usage

Run the script from the folder that contains your `.tex`, `.bib`, `.aux` file and figures.

```bash
# LaTeX -> Word
python tex2docx_v2.py paper.tex
python tex2docx_v2.py paper.tex -o manuscript.docx

# Choose a citation style and/or a Word template
python tex2docx_v2.py paper.tex --csl springer-basic-author-date.csl
python tex2docx_v2.py paper.tex --reference-doc template.docx

# Higher-resolution figures, with a table of contents
python tex2docx_v2.py paper.tex --dpi 300 --toc

# Word -> LaTeX
python tex2docx_v2.py manuscript.docx
python tex2docx_v2.py manuscript.docx -o draft.tex
```

> **Tip:** compile your paper with LaTeX at least once before converting, so the `.aux` file exists. Without it, figure/table numbers and `\ref` links cannot be resolved and will show as `??`.

### Options

| Option | Description |
| --- | --- |
| `-o`, `--output` | Output path. Default: same name with the other extension. |
| `--bib FILE` | Use this `.bib` file instead of auto-detecting it from the `.tex`. |
| `--csl FILE` | [CSL](https://citationstyles.org/) citation style. Default is Chicago author-date. Styles are available at [zotero.org/styles](https://www.zotero.org/styles). |
| `--reference-doc FILE` | A Word file whose styles (fonts, headings) are applied to the output. |
| `--dpi N` | Resolution used when rendering PDF figures to PNG (default 200). |
| `--toc` | Add a table of contents. |
| `--keep-temp` | Keep the temporary preprocessed `.tex` (`_pandoc_tmp.tex`) for debugging. |

Any other flags are passed through to Pandoc.

## Known limitations

- Pandoc does not run LaTeX, so **custom macros, TikZ drawings and some complex tables** (`\multirow`, `\multicolumn`, `sidewaystable`) may be simplified. Check tables in the output.
- **Equations** become editable Word equations, but Word does not auto-number them. References such as "Eq. (3)" in the text keep the correct number.
- Only the **first page** of each PDF figure is used.
- Journal class files (`sn-jnl`, `elsarticle`, etc.) are not interpreted. Front-matter conversion is built for the Springer Nature `sn-jnl` macros; other classes may need small additions in `convert_front_matter()`.
- **Word to LaTeX** produces a generic `article`-style file. Citations and cross-references come back as plain text because Word does not store the `\cite{}` / `\ref{}` links. Keep your original `.tex` and `.bib` as the master copy.

## How it works

1. Reads the `.tex` and strips comments.
2. Converts the front matter and journal-specific commands to plain LaTeX.
3. Renders PDF figures to PNG and rewrites `\includegraphics` paths.
4. Reads labels from the `.aux` file and replaces cross-references and caption numbers.
5. Writes a temporary `.tex`, runs Pandoc (with citeproc if a `.bib` file is found), then deletes the temporary file.

## Contributing

Issues and pull requests are welcome, especially sample `.tex` files from other journal templates that convert poorly.

## License

Add a license of your choice (for example [MIT](https://choosealicense.com/licenses/mit/)) before publishing.
