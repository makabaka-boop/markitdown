"""Rebuild the reviewed PDF fixture used by the annotation extraction tests.

Run from packages/markitdown with:
    python tests/test_files/generate_pdf_annotations_fixture.py

The fixture is a two-page technical document containing sticky notes,
highlights (with and without comments, including cross-line and
discontinuous selections), threaded replies (including an orphan reply),
an unresolvable annotation reference, a highlight with malformed
QuadPoints, an empty note, and an empty highlight placed off-page.
"""

from pathlib import Path

import pymupdf

FIXTURES = Path(__file__).resolve().parent


def _set_reply_to(
    doc: pymupdf.Document, annot: pymupdf.Annot, parent_xref: int
) -> None:
    """Link an annotation to its parent via the /IRT key."""
    doc.xref_set_key(annot.xref, "IRT", f"{parent_xref} 0 R")
    annot.update()


def main() -> None:
    doc = pymupdf.open()
    doc.new_page()
    doc.new_page()
    page1 = doc[0]
    page2 = doc[1]

    page1_text = (
        "Introduction to Distributed Systems\n"
        "A distributed system is a collection of independent computers that\n"
        "appears to its users as a single coherent system. The key goals are\n"
        "transparency, scalability, and reliability. Consensus protocols allow\n"
        "nodes to agree on a single value even when some nodes may fail.\n"
    )
    page2_text = (
        "Replication and Consistency\n"
        "Replication improves availability but introduces the challenge of\n"
        "keeping replicas consistent. Eventual consistency allows temporary\n"
        "divergence between copies. Strong consistency guarantees that all\n"
        "reads return the latest committed write at the cost of latency.\n"
    )
    page1.insert_text((72, 72), page1_text, fontsize=12)
    page2.insert_text((72, 72), page2_text, fontsize=12)

    # 1) Root sticky note with an author and multi-line text.
    root_note = page1.add_text_annot(
        pymupdf.Point(380, 90),
        "Please clarify the failure model.\nWhich faults are Byzantine?",
    )
    root_note.set_info(title="Alice Reviewer")
    root_note.update()

    # 2) Highlight with a comment on a single-line phrase.
    highlighted_with_comment = page1.add_highlight_annot(
        page1.search_for("collection of independent computers")
    )
    highlighted_with_comment.set_info(
        title="Bob", content="Cite the textbook definition for this claim."
    )
    highlighted_with_comment.update()

    # 3) Highlight WITHOUT a comment spanning two discontinuous selections on
    #    different lines (text between the quads must not be quoted).
    discontinuous = page1.add_highlight_annot(
        page1.search_for("appears to its users as a single coherent system.")
        + page1.search_for("transparency, scalability, and reliability.")
    )
    discontinuous.set_info(title="Bob")
    discontinuous.update()

    # 4) Highlight whose selection wraps across two lines, with a comment.
    cross_line = page1.add_highlight_annot(
        page1.search_for("Consensus protocols allow nodes to agree on a single value")
    )
    cross_line.set_info(
        title="Carol", content="Quote wraps across two lines; keep both lines."
    )
    cross_line.update()

    # 5) First reply to the root note (multi-line text).
    first_reply = page1.add_text_annot(
        pymupdf.Point(400, 110), "Crash faults only.\nWe exclude Byzantine nodes here."
    )
    first_reply.set_info(title="Charlie")
    _set_reply_to(doc, first_reply, root_note.xref)

    # 6) Nested reply to the first reply.
    nested_reply = page1.add_text_annot(
        pymupdf.Point(420, 130), "Agreed; I will spell this out."
    )
    nested_reply.set_info(title="Alice Reviewer")
    _set_reply_to(doc, nested_reply, first_reply.xref)

    # 7) Orphan reply whose parent object does not exist in the file.
    orphan = page1.add_text_annot(
        pymupdf.Point(440, 150), "My parent annotation went missing."
    )
    orphan.set_info(title="Olivia Orphan")
    doc.xref_set_key(orphan.xref, "IRT", "9999 0 R")
    orphan.update()

    # 8) Empty sticky note: carries no retrievable information and is dropped.
    empty_note = page1.add_text_annot(pymupdf.Point(460, 170), "")
    empty_note.set_info(title="Nobody")
    empty_note.update()

    # 9) Highlight with malformed QuadPoints: must fall back to its Rect and
    #    still quote the covered text instead of breaking extraction.
    malformed = page1.add_highlight_annot(page1.search_for("single coherent system"))
    malformed.set_info(title="Dana", content="QuadPoints are malformed on purpose.")
    malformed.update()
    doc.xref_set_key(malformed.xref, "QuadPoints", "[1 2 3]")

    # 10) Highlight with no comment over a region with no text: dropped.
    blank_highlight = page1.add_highlight_annot(pymupdf.Rect(360, 600, 500, 620))
    blank_highlight.set_info(title="Nobody")
    blank_highlight.update()

    # 11) Independent note on page 2.
    page2_note = page2.add_text_annot(
        pymupdf.Point(380, 90), "Define the consistency levels.\nUse a table."
    )
    page2_note.set_info(title="Dana")
    page2_note.update()

    # 12) Reply to the page 2 note.
    page2_reply = page2.add_text_annot(
        pymupdf.Point(400, 110), "Will add a CAP theorem summary."
    )
    page2_reply.set_info(title="Eve")
    _set_reply_to(doc, page2_reply, page2_note.xref)

    # 13) Append an unresolvable reference to page 1's annotation array.
    kind, value = doc.xref_get_key(page1.xref, "Annots")
    assert kind == "array", value
    doc.xref_set_key(page1.xref, "Annots", value.replace("]", " 9998 0 R]"))

    doc.save(FIXTURES / "pdf_annotations.pdf", deflate=True)


if __name__ == "__main__":
    main()
