from asta_revision_workflow.docx_reference_list_to_ris import (
    author_list,
    export_ris,
    parse_reference,
    ris_text_from_references,
    Reference as RisReference,
    split_reference_entries,
    validate_ris_matches_docx,
    write_record,
)
import pytest
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

from asta_revision_workflow.docx_numeric_to_endnote_temp import Reference, convert, make_temp_citation


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


def test_author_list_handles_ampersand_joined_authors():
    assert author_list("Cao, R. & Zhang, Y.") == ["Cao, R.", "Zhang, Y."]
    assert author_list("Pasini, D., Bracken, A.P., Jensen, M.R., Lazzerini Denchi, E. & Helin, K.") == [
        "Pasini, D.",
        "Bracken, A.P.",
        "Jensen, M.R.",
        "Lazzerini Denchi, E.",
        "Helin, K.",
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


def test_reference_split_does_not_break_h3_variant_titles():
    text = (
        "1. Kraushaar, D. et al. (2013). Genome-wide incorporation dynamics reveal "
        "distinct categories of turnover for the histone variant H3.3. Genome Biology 14, R121-R121. "
        "\n2. Tie, F., Banerjee, R., Conrad, P.A., Scacheri, P. & Harte, P. (2012). "
        "Histone Demethylase UTX and Chromatin Remodeler BRM Bind Directly to CBP and "
        "Modulate Acetylation of Histone H3 Lysine 27. Molecular and Cellular Biology 32, 2323-2334."
    )

    entries = split_reference_entries(text)

    assert len(entries) == 2
    assert entries[0][0] == 1
    assert "H3.3" in entries[0][1]
    assert entries[1][0] == 2
    assert "Lysine 27" in entries[1][1]


def test_reference_export_stops_before_appended_methods_note(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "refs.ris"
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Main text.</w:t></w:r></w:p>
    <w:p><w:r><w:t>1.Cao, R., and Zhang, Y. (2004). Example title. Molecular Cell 15, 57-67.</w:t></w:r></w:p>
    <w:p><w:r><w:t>AI use in writing: This note describes the revision workflow.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    with ZipFile(source, "w", ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)

    count = export_ris(source, output)

    assert count == 1
    ris = output.read_text(encoding="utf-8")
    assert "TI  - Example title" in ris
    assert "AI use" not in ris


def test_reference_export_uses_complete_duplicate_reference_block(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "refs.ris"
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Main text.</w:t></w:r></w:p>
    <w:p><w:r><w:t>1.Cao, R., and Zhang, Y. (2004). First title. Molecular Cell 15, 57-67.</w:t></w:r></w:p>
    <w:p><w:r><w:t>1.Cao, R., and Zhang, Y. (2004). First title. Molecular Cell 15, 57-67.</w:t></w:r></w:p>
    <w:p><w:r><w:t>2.Kaplan, T., Liu, C.L., Erkmann, J.A., Holik, J., Grunstein, M., Kaufman, P.D., Friedman, N. and Rando, O.J. (2008). Cell Cycle- and Chaperone-Mediated Regulation of H3K56ac Incorporation in Yeast. PLoS Genetics 4, e1000270. 10.1371/journal.pgen.1000270.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    with ZipFile(source, "w", ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)

    assert export_ris(source, output) == 2
    ris = output.read_text(encoding="utf-8")

    assert "TI  - First title" in ris
    assert "TI  - Cell Cycle- and Chaperone-Mediated Regulation of H3K56ac Incorporation in Yeast" in ris


def test_duplicate_ris_ids_are_suffixed_with_reference_number():
    references = [
        RisReference(43, "Tie, F.", "2012", "Histone Demethylase UTX", "", "", "", "", ""),
        RisReference(71, "Tie, F.", "2012", "Histone Demethylase UTX", "", "", "", "", ""),
    ]
    ris = ris_text_from_references(references)

    assert "ID  - Tie2012HistoneDemethylUTXRef43" in ris
    assert "ID  - Tie2012HistoneDemethylUTXRef71" in ris


def test_ris_provenance_check_requires_current_docx_metadata(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "refs.ris"
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Main text.</w:t></w:r></w:p>
    <w:p><w:r><w:t>1.Cao, R., and Zhang, Y. (2004). Example title. Molecular Cell 15, 57-67.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    with ZipFile(source, "w", ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)

    export_ris(source, output)

    assert validate_ris_matches_docx(source, output) == 1

    output.write_text(
        output.read_text(encoding="utf-8").replace("TI  - Example title", "TI  - Archive-filled title"),
        encoding="utf-8",
    )
    with pytest.raises(RuntimeError, match="RIS provenance check failed"):
        validate_ris_matches_docx(source, output)


def test_reference_export_can_overlay_pinned_complete_author_metadata(tmp_path):
    source = tmp_path / "source.docx"
    metadata = tmp_path / "metadata.ris"
    output = tmp_path / "refs.ris"
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Main text.</w:t></w:r></w:p>
    <w:p><w:r><w:t>1.Margueron, R. et al. (2009). Role of the polycomb protein EED in the propagation of repressive histone marks. Nature 461, 762-767.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    metadata.write_text(
        "\n".join(
            [
                "TY  - JOUR",
                "TI  - Role of the polycomb protein EED in the propagation of repressive histone marks",
                "AU  - Margueron, R.",
                "AU  - Justin, N.",
                "AU  - Ohno, K.",
                "PY  - 2009",
                "JO  - Nature",
                "VL  - 461",
                "SP  - 762",
                "EP  - 767",
                "DO  - 10.1038/nature08398",
                "ID  - Margueron2009RoleOfThe",
                "ER  -",
                "",
            ]
        ),
        encoding="utf-8",
    )
    with ZipFile(source, "w", ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)

    export_ris(source, output, metadata)
    ris = output.read_text(encoding="utf-8")

    assert "AU  - Margueron, R." in ris
    assert "AU  - Justin, N." in ris
    assert "AU  - Ohno, K." in ris
    assert "DO  - 10.1038/nature08398" in ris
    assert validate_ris_matches_docx(source, output, metadata) == 1


def test_reference_export_requires_metadata_match_for_abbreviated_authors(tmp_path):
    source = tmp_path / "source.docx"
    metadata = tmp_path / "metadata.ris"
    output = tmp_path / "refs.ris"
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Main text.</w:t></w:r></w:p>
    <w:p><w:r><w:t>1.Margueron, R. et al. (2009). Role of the polycomb protein EED in the propagation of repressive histone marks. Nature 461, 762-767.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    metadata.write_text(
        "\n".join(
            [
                "TY  - JOUR",
                "TI  - A different title",
                "AU  - Margueron, R.",
                "AU  - Justin, N.",
                "PY  - 2009",
                "ER  -",
                "",
            ]
        ),
        encoding="utf-8",
    )
    with ZipFile(source, "w", ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)

    with pytest.raises(RuntimeError, match="did not match the metadata RIS"):
        export_ris(source, output, metadata, require_metadata_match=True)


def test_end_year_entry_with_ampersand_author_list_parses_title():
    reference = parse_reference(
        61,
        "Allshire, R.C. & Madhani, H.D. Ten principles of heterochromatin formation and function. "
        "Nature Reviews Molecular Cell Biology 19, 229-244 (2018).",
    )

    assert reference is not None
    assert reference.authors == "Allshire, R.C. & Madhani, H.D."
    assert reference.year == "2018"
    assert reference.title == "Ten principles of heterochromatin formation and function"


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


def test_duplicate_author_year_can_be_disambiguated_with_title():
    references = {
        1: Reference("Lee", "2015", "Genome-wide activities of Polycomb complexes"),
        2: Reference("Lee", "2015", "Genome-wide activities of Polycomb complexes"),
    }

    assert make_temp_citation("1", references, {"Lee, 2015"}) == (
        "{Lee, 2015, Genome-wide activities of Polycomb complexes}"
    )


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


def test_numeric_citation_conversion_preserves_prc2_subtype_decimal(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>MTF2-containing PRC2.</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r>
      <w:r><w:t> supports nucleation</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>1.Cao, R., and Zhang, Y. (2004). Example title. Molecular Cell 15, 57-67.</w:t></w:r></w:p>
    <w:p><w:r><w:t>2.Pasini, D., and Helin, K. (2004). Second title. EMBO Journal 23, 4061-4071.</w:t></w:r></w:p>
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
    assert "PRC2.1 supports nucleation {Pasini, 2004}" in converted
    assert "PRC2. {Cao, 2004}" not in converted


def test_numeric_citation_conversion_uses_ris_and_strips_stale_endnote_parts(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"
    ris = tmp_path / "refs.ris"
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Claim</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>1.Bad, B. (2020). Truncated source title. Journal 1, 1-2.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    styles_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:styleId="EndNoteBibliography" w:type="paragraph"><w:name w:val="EndNote Bibliography"/></w:style>
</w:styles>
"""
    comments_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
"""
    ris.write_text(
        "\n".join(
            [
                "TY  - JOUR",
                "TI  - Complete RIS title",
                "AU  - Real, Author",
                "PY  - 2021",
                "ER  -",
                "",
            ]
        ),
        encoding="utf-8",
    )
    with ZipFile(source, "w", ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/comments.xml", comments_xml)

    convert(source, output, ris=ris)

    with ZipFile(output) as zf:
        names = set(zf.namelist())
        root = ET.fromstring(zf.read("word/document.xml"))
        styles = zf.read("word/styles.xml").decode("utf-8")
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    converted = "".join(t.text or "" for t in root.findall(".//w:t", ns))
    assert "{Real, 2021}" in converted
    assert "word/comments.xml" not in names
    assert "EndNote" not in styles


def test_keep_references_retains_only_complete_reference_block(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Claim</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>1.Cao, R., and Zhang, Y. (2004). Stale title. Molecular Cell 15, 57-67.</w:t></w:r></w:p>
    <w:p><w:r><w:t>1.Cao, R., and Zhang, Y. (2004). Canonical first title. Molecular Cell 15, 57-67.</w:t></w:r></w:p>
    <w:p><w:r><w:t>2.Pasini, D., and Helin, K. (2004). Canonical second title. EMBO Journal 23, 4061-4071.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    with ZipFile(source, "w", ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)

    convert(source, output, keep_references=True)

    with ZipFile(output) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    converted = "\n".join(t.text or "" for t in root.findall(".//w:t", ns))

    assert "{Pasini, 2004}" in converted
    assert "Stale title" not in converted
    assert "Canonical first title" in converted
    assert "Canonical second title" in converted
