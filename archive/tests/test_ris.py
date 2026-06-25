from citeproc_endnote_uv.ris import bibtex_to_ris


def test_bibtex_to_ris_article():
    ris = bibtex_to_ris(
        """
        @article{cao2004,
          author = {Cao, Rong and Zhang, Yi},
          title = {Example title},
          journal = {Molecular Cell},
          year = {2004},
          volume = {15},
          pages = {57--67},
          doi = {10.1016/example}
        }
        """
    )

    assert "TY  - JOUR" in ris
    assert "TI  - Example title" in ris
    assert "AU  - Cao, Rong" in ris
    assert "AU  - Zhang, Yi" in ris
    assert "PY  - 2004" in ris
    assert "JO  - Molecular Cell" in ris
    assert "SP  - 57" in ris
    assert "EP  - 67" in ris
    assert "DO  - 10.1016/example" in ris
    assert "ID  - cao2004" in ris
    assert "ER  -" in ris
