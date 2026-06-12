from citeproc_endnote_uv.docx_reference_list_to_ris import author_list, parse_reference, write_record
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

from citeproc_endnote_uv.docx_numeric_to_endnote_temp import Reference, convert, make_temp_citation


def test_author_list_preserves_all_named_authors():
    authors = (
        "Ferrari, Karin J., Scelfo, A., Jammula, S., Cuomo, A., Barozzi, I., "
        "Stutzer, A., Fischle, W., Bonaldi, T., and Pasini, D."
    )

    assert author_list(authors) == [
        "Ferrari, Karin J.",
        "Scelfo, A.",
        "Jammula, S.",
        "Cuomo, A.",
        "Barozzi, I.",
        "Stutzer, A.",
        "Fischle, W.",
        "Bonaldi, T.",
        "Pasini, D.",
    ]


def test_reference_record_writes_all_authors():
    reference = parse_reference(
        28,
        "Ferrari, Karin J., Scelfo, A., Jammula, S., Cuomo, A., Barozzi, I., "
        "Stutzer, A., Fischle, W., Bonaldi, T., and Pasini, D. (2014). "
        "Polycomb-Dependent H3K27me1 and H3K27me2 Regulate Active Transcription and Enhancer Fidelity. "
        "Molecular Cell 53, 49-62. 10.1016/j.molcel.2013.10.030.",
    )

    assert reference is not None
    record = "\n".join(write_record(reference))
    assert "AU  - Ferrari, Karin J." in record
    assert "AU  - Scelfo, A." in record
    assert "AU  - Pasini, D." in record
    assert "DO  - 10.1016/j.molcel.2013.10.030" in record


def test_only_true_author_year_collisions_need_titles():
    references = {
        1: Reference("McCabe", "2012", "Mutation of A677 in histone methyltransferase EZH2"),
        2: Reference("McCabe", "2012", "EZH2 inhibition as a therapeutic strategy"),
        3: Reference("Lee", "2015", "Genome-wide activities of Polycomb complexes"),
        4: Reference("Lee", "2015", "Genome-wide activities of Polycomb complexes"),
    }
    ambiguous_keys = {"McCabe, 2012"}

    assert make_temp_citation("1,2", references, ambiguous_keys) == (
        "{McCabe, 2012, Mutation of A677 in histone methyltransferase EZH2; "
        "McCabe, 2012, EZH2 inhibition as a therapeutic strategy}"
    )
    assert make_temp_citation("3", references, ambiguous_keys) == "{Lee, 2015}"


def test_numeric_citation_conversion_adds_separator_before_temporary_cite(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>H3K27me3</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>1.Cao, R., and Zhang, Y. (2004). Example title. Molecular Cell 15, 57-67. 10.1016/example.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    with ZipFile(source, "w", ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)

    convert(source, output)

    with ZipFile(output) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    converted = "".join(t.text or "" for t in root.findall(".//w:t", ns))
    assert "H3K27me3 {Cao, 2004}" in converted
    assert "H3K27me3{Cao, 2004}" not in converted
