from citeproc_endnote_uv.cli import endnote_citation, parse_bibtex, walk


def test_parse_bibtex_and_format_author_year():
    entries = parse_bibtex(
        """
        @article{cao2004,
          author = {Cao, Rong and Zhang, Yi},
          title = {Example},
          year = {2004}
        }
        """
    )

    assert endnote_citation("cao2004", entries["cao2004"]) == "Cao, 2004"


def test_filter_replaces_pandoc_cite_node():
    document = {
        "pandoc-api-version": [1, 23],
        "meta": {},
        "blocks": [
            {
                "t": "Para",
                "c": [
                    {"t": "Str", "c": "Text"},
                    {"t": "Space"},
                    {
                        "t": "Cite",
                        "c": [
                            [
                                {
                                    "citationId": "cao2004",
                                    "citationPrefix": [],
                                    "citationSuffix": [],
                                    "citationMode": {"t": "NormalCitation"},
                                    "citationNoteNum": 1,
                                    "citationHash": 0,
                                }
                            ],
                            [{"t": "Str", "c": "[@cao2004]"}],
                        ],
                    },
                ],
            }
        ],
    }
    bib = {"cao2004": {"author": "Cao, Rong and Zhang, Yi", "year": "2004"}}

    assert walk(document, bib)["blocks"][0]["c"][-1] == {"t": "Str", "c": "{Cao, 2004}"}

