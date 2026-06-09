# citeproc-endnote-uv

`citeproc-endnote-uv` is a small Pandoc JSON filter that turns LaTeX/BibTeX citations into EndNote temporary citations in Word documents.

It is meant as a current, slim replacement for the historical Haskell `citeproc-endnote` workflow:

```bash
pandoc -f latex -t docx input.tex -o output.docx -F citeproc-endnote
```

The generated DOCX contains editable Word text such as `{Cao, 2004}`. Open the file in Word, then run EndNote's **Update Citations and Bibliography** command to convert those temporary citations into EndNote-managed citations and bibliography entries.

## Install

From this repository:

```bash
uv tool install .
```

From GitHub:

```bash
uv tool install git+https://github.com/jacob-greene/citeproc-endnote-uv
```

For project-local use:

```bash
uv venv
uv pip install -e .
```

## Use

Pass a BibTeX file to the filter through the environment. This avoids depending on Pandoc metadata details and keeps the tool explicit.

```bash
CITEPROC_ENDNOTE_BIBLIOGRAPHY=references.bib \
  pandoc -f latex -t docx input.tex -o output.docx -F citeproc-endnote
```

The filter also checks common local names if the environment variable is not set:

- `references.bib`
- `latex/references.bib`

## Notes

- This tool emits EndNote temporary citations, not final EndNote fields.
- EndNote does the final conversion inside Word.
- If multiple library records match an author-year pair, EndNote may ask you to resolve the match.
- Record-number citations can be supported later if the BibTeX source includes an EndNote record number field.
