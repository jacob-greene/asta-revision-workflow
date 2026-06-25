from zipfile import ZIP_DEFLATED, ZipFile

from asta_revision_workflow.docx_extract_comments import extract_comments, format_text


def test_extract_comments_from_word_xml(tmp_path):
    docx = tmp_path / "commented.docx"
    document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Before </w:t></w:r>
      <w:commentRangeStart w:id="2"/>
      <w:r><w:t>anchored words</w:t></w:r>
      <w:commentRangeEnd w:id="2"/>
      <w:r><w:commentReference w:id="2"/></w:r>
      <w:r><w:t> after.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
    comments_xml = """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="2">
    <w:p><w:r><w:t>Revise this sentence.</w:t></w:r></w:p>
  </w:comment>
</w:comments>
"""
    with ZipFile(docx, "w", ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/comments.xml", comments_xml)

    anchors = extract_comments(docx)

    assert len(anchors) == 1
    assert anchors[0].comment_id == "2"
    assert anchors[0].comment_text == "Revise this sentence."
    assert anchors[0].anchored_text == "anchored words"
    assert anchors[0].paragraph_text == "Before anchored words after."
    assert "COMMENT 2" in format_text(anchors)
