# LaTeX to Word with EndNote

Use this workflow when TeX/BibTeX is the source of truth and the desired output is an editable Word document whose citations can be converted into EndNote fields.

## Requirements

- `pandoc`
- this package installed with `uv tool install .` or `uv tool install git+https://github.com/jacob-greene/citeproc-endnote-uv`
- a BibTeX file with stable citation keys

## Convert

```bash
CITEPROC_ENDNOTE_BIBLIOGRAPHY=references.bib \
  pandoc -f latex -t docx input.tex -o output.docx -F citeproc-endnote

bib-to-ris references.bib output.ris
docx-word-sanity output.docx
```

The DOCX contains EndNote temporary citations such as `{Cao, 2004}`. Import `output.ris` into EndNote, open `output.docx` in Word, then run EndNote's `Update Citations and Bibliography`.

## Optional Word Style Reference

Pandoc can inherit styles from an existing Word document:

```bash
CITEPROC_ENDNOTE_BIBLIOGRAPHY=references.bib \
  pandoc -f latex -t docx input.tex \
  --reference-doc=style-reference.docx \
  -o output.docx \
  -F citeproc-endnote
```

## Citation Practice

- Use BibTeX keys, not manually typed superscript numbers.
- Prefer placing citations after words rather than directly after digit-ending terms.
- Good: `trimethylation is completed\cite{Chory2019}.`
- Avoid: `H3K27me3\cite{Chory2019}` when a numeric citation style may later place citation numbers directly after `H3K27me3`.
- Let EndNote handle final citation formatting in Word.

## Static Fallback

Pandoc's built-in citeproc route produces static citations and a static bibliography:

```bash
pandoc -f latex -t docx input.tex \
  --citeproc \
  --bibliography references.bib \
  -o output-static.docx
```

Use this only when EndNote field conversion is not required.
