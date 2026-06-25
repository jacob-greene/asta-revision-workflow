import base64
import html
from zipfile import ZIP_DEFLATED, ZipFile

from asta_revision_workflow.docx_endnote_to_ris import export_ris, records_from_docx


def make_docx(path, document_xml):
    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)


def endnote_xml():
    return """<EndNote><Cite><record>
<rec-number>42</rec-number>
<contributors><authors>
<author>Margueron, R.</author>
<author>Justin, N.</author>
<author>Ohno, K.</author>
</authors></contributors>
<titles>
<title>Role of the polycomb protein EED in the propagation of repressive histone marks</title>
<secondary-title>Nature</secondary-title>
</titles>
<volume>461</volume>
<pages>762-767</pages>
<dates><year>2009</year></dates>
<electronic-resource-num>10.1038/nature08398</electronic-resource-num>
</record></Cite></EndNote>"""


def endnote_xml_with_unescaped_ampersand():
    return endnote_xml().replace("Nature", "Nature Structural & Molecular Biology")


def test_extracts_complete_authors_from_instr_text(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "metadata.ris"
    escaped = html.escape(f"ADDIN EN.CITE {endnote_xml()}")
    make_docx(
        source,
        f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:instrText>{escaped}</w:instrText></w:r></w:p>
    <w:p><w:r><w:t>1. Margueron, R. et al. (2009). Role of the polycomb protein EED in the propagation of repressive histone marks. Nature 461, 762-767.</w:t></w:r></w:p>
  </w:body>
</w:document>
""",
    )

    audit = export_ris(source, output)
    ris = output.read_text(encoding="utf-8")

    assert audit["embedded_record_count"] == 1
    assert "AU  - Margueron, R." in ris
    assert "AU  - Justin, N." in ris
    assert "AU  - Ohno, K." in ris
    assert "DO  - 10.1038/nature08398" in ris


def test_extracts_records_from_base64_field_data(tmp_path):
    source = tmp_path / "source.docx"
    encoded = base64.b64encode(endnote_xml().encode("utf-8")).decode("ascii")
    make_docx(
        source,
        f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:fldData>{encoded}</w:fldData></w:r></w:p></w:body>
</w:document>
""",
    )

    records = records_from_docx(source)

    assert len(records) == 1
    assert records[0].authors == ("Margueron, R.", "Justin, N.", "Ohno, K.")


def test_sanitizes_unescaped_ampersands_in_endnote_payload(tmp_path):
    source = tmp_path / "source.docx"
    encoded = base64.b64encode(endnote_xml_with_unescaped_ampersand().encode("utf-8")).decode("ascii")
    make_docx(
        source,
        f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:fldData>{encoded}</w:fldData></w:r></w:p></w:body>
</w:document>
""",
    )

    records = records_from_docx(source)

    assert len(records) == 1
    assert records[0].journal == "Nature Structural & Molecular Biology"
