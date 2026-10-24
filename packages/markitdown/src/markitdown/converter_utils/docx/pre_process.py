import re
import struct
import zipfile
from io import BytesIO
from typing import Any, BinaryIO
from xml.etree import ElementTree as ET

from bs4 import BeautifulSoup, Tag

from .math.omml import OMML_NS, oMath2Latex

WORD_PROCESSINGML_NAMESPACE = (
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
)

MATH_ROOT_TEMPLATE = "".join(
    (
        "<w:document ",
        'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ',
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ',
        'xmlns:o="urn:schemas-microsoft-com:office:office" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ',
        'xmlns:v="urn:schemas-microsoft-com:vml" ',
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ',
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ',
        'xmlns:w10="urn:schemas-microsoft-com:office:word" ',
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ',
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ',
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ',
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ',
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">',
        "{0}</w:document>",
    )
)


def _convert_omath_to_latex(tag: Tag) -> str:
    """
    Converts an OMML (Office Math Markup Language) tag to LaTeX format.

    Args:
        tag (Tag): A BeautifulSoup Tag object representing the OMML element.

    Returns:
        str: The LaTeX representation of the OMML element.
    """
    # Format the tag into a complete XML document string
    math_root = ET.fromstring(MATH_ROOT_TEMPLATE.format(str(tag)))
    # Find the 'oMath' element within the XML document
    math_element = math_root.find(OMML_NS + "oMath")
    if math_element is None:
        return ""
    # Convert the 'oMath' element to LaTeX using the oMath2Latex function
    latex = oMath2Latex(math_element).latex
    return latex


def _get_omath_tag_replacement(tag: Tag, block: bool = False) -> Tag:
    """
    Creates a replacement tag for an OMML (Office Math Markup Language) element.

    Args:
        tag (Tag): A BeautifulSoup Tag object representing the "oMath" element.
        block (bool, optional): If True, the LaTeX will be wrapped in double dollar signs for block mode. Defaults to False.

    Returns:
        Tag: A BeautifulSoup Tag object representing the replacement element.
    """
    t_tag = Tag(name="w:t")
    t_tag.string = (
        f"$${_convert_omath_to_latex(tag)}$$"
        if block
        else f"${_convert_omath_to_latex(tag)}$"
    )
    r_tag = Tag(name="w:r")
    r_tag.append(t_tag)
    return r_tag


def _replace_equations(tag: Tag):
    """
    Replaces OMML (Office Math Markup Language) elements with their LaTeX equivalents.

    Args:
        tag (Tag): A BeautifulSoup Tag object representing the OMML element. Could be either "oMathPara" or "oMath".

    Raises:
        ValueError: If the tag is not supported.
    """
    if tag.name == "oMathPara":
        # Create a new paragraph tag
        p_tag = Tag(name="w:p")
        # Replace each 'oMath' child tag with its LaTeX equivalent as block equations
        for child_tag in tag.find_all("oMath"):
            p_tag.append(_get_omath_tag_replacement(child_tag, block=True))
        # Replace the original 'oMathPara' tag with the new paragraph tag
        tag.replace_with(p_tag)
    elif tag.name == "oMath":
        # Replace the 'oMath' tag with its LaTeX equivalent as inline equation
        tag.replace_with(_get_omath_tag_replacement(tag, block=False))
    else:
        raise ValueError(f"Not supported tag: {tag.name}")


def _pre_process_strike(content: bytes) -> bytes:
    """
    Pre-processes the strikethrough content in a DOCX -> XML file by normalizing double
    strikethrough runs to single strikethrough runs.

    Word marks double strikethrough with "w:dstrike", which downstream converters do not
    recognize, causing those runs to lose their strikethrough entirely. Renaming the tag to
    "w:strike" preserves the strikethrough semantics.

    Args:
        content (bytes): The XML content of the DOCX file as bytes.

    Returns:
        bytes: The processed content with "dstrike" elements renamed to "strike", encoded as bytes.
    """
    # Double strikethrough is rare, and parsing/reserializing the XML is expensive
    # on large documents, so skip the round-trip when there is nothing to rename.
    if b"dstrike" not in content:
        return content

    soup = BeautifulSoup(content.decode(), features="xml")
    for tag in soup.find_all("dstrike"):
        tag.name = "strike"
    return str(soup).encode()


def _pre_process_math(content: bytes) -> bytes:
    """
    Pre-processes the math content in a DOCX -> XML file by converting OMML (Office Math Markup Language) elements to LaTeX.
    This preprocessed content can be directly replaced in the DOCX file -> XMLs.

    Args:
        content (bytes): The XML content of the DOCX file as bytes.

    Returns:
        bytes: The processed content with OMML elements replaced by their LaTeX equivalents, encoded as bytes.
    """
    soup = BeautifulSoup(content.decode(), features="xml")
    for tag in soup.find_all("oMathPara"):
        _replace_equations(tag)
    for tag in soup.find_all("oMath"):
        _replace_equations(tag)
    return str(soup).encode()


def _fix_zip_filename_casing(input_docx: BinaryIO) -> BinaryIO:
    """
    Fix ZIP files where local file header filenames differ in casing
    from the central directory filenames.

    Some document generators (e.g. certain Microsoft Word versions,
    legal document systems) produce .docx/.pptx files where the central
    directory records one casing (e.g. 'customXml/item2.xml') but
    the local file headers record another (e.g. 'customXML/item2.xml').
    Python's zipfile module raises BadZipFile when reading such files.

    This function patches local file header filenames to match the
    central directory, which is the authoritative source used by
    zipfile.ZipFile.
    """
    input_docx.seek(0)
    raw = bytearray(input_docx.read())

    # Read the central directory to get authoritative filenames
    try:
        with zipfile.ZipFile(BytesIO(raw), "r") as zf:
            cd_entries = {
                zi.header_offset: (zi.orig_filename, zi.flag_bits)
                for zi in zf.infolist()
            }
    except zipfile.BadZipFile:
        # Can't even read central directory — return as-is, let it fail later
        input_docx.seek(0)
        return input_docx

    patched = False
    for offset, (cd_name, flag_bits) in cd_entries.items():
        # Verify local file header signature
        if offset + 30 > len(raw) or raw[offset : offset + 4] != b"PK\x03\x04":
            continue

        local_fname_len = struct.unpack_from("<H", raw, offset + 26)[0]
        if offset + 30 + local_fname_len > len(raw):
            continue

        local_name = bytes(raw[offset + 30 : offset + 30 + local_fname_len])
        # ZIP filenames are cp437 unless flag bit 11 marks them as UTF-8, which is
        # how zipfile decoded orig_filename in the first place.
        central_name = cd_name.encode("utf-8" if flag_bits & 0x800 else "cp437")

        # Only patch if lengths match but content differs (casing mismatch)
        if (
            local_name != central_name
            and len(local_name) == len(central_name)
            and local_name.lower() == central_name.lower()
        ):
            raw[offset + 30 : offset + 30 + local_fname_len] = central_name
            patched = True

    if patched:
        return BytesIO(bytes(raw))
    input_docx.seek(0)
    return input_docx


def _pre_process_styles(content: bytes) -> bytes:
    """
    Repairs DOCX style definitions that Mammoth cannot read.

    Mammoth indexes ``w:type`` and ``w:styleId`` directly, so a ``w:style``
    element missing either attribute causes conversion to fail with a
    ``KeyError`` before any document text can be extracted.

    ``w:type`` is optional in OOXML: when it is absent the style type defaults
    to ``paragraph``, so the attribute is filled in rather than dropping the
    style, which would discard its formatting (a heading would be emitted as
    plain body text). A style with no ``w:styleId`` cannot be referenced by the
    document body, so it is removed. Double strikethrough is normalized to
    single strikethrough in the same namespace-aware pass.

    Match elements and attributes by namespace URI, preserving their qualified
    names when repairing the XML. Return the original bytes if no repair is needed.
    """
    from lxml import etree

    namespace = f"{{{WORD_PROCESSINGML_NAMESPACE}}}"
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    root = etree.fromstring(content, parser=parser)
    changed = False

    for style in root.findall(namespace + "style"):
        if namespace + "styleId" not in style.attrib:
            root.remove(style)
            changed = True
        elif namespace + "type" not in style.attrib:
            style.set(namespace + "type", "paragraph")
            changed = True

    for strike in root.iter(namespace + "dstrike"):
        strike.tag = namespace + "strike"
        changed = True

    if not changed:
        return content

    return etree.tostring(root.getroottree(), encoding="utf-8", xml_declaration=True)


# Opening revision tags regardless of the namespace prefix used (Word writes
# "w:", but conforming generators may bind the namespace differently). Text
# containing the escaped sequence "&lt;w:ins" does not match.
_REVISION_TAG_PATTERN = re.compile(
    rb"<(?:[A-Za-z_][\w.-]*:)?" rb"(?:ins|del|moveFrom|moveTo)(?:\s|/|>)"
)

# Revision markers that can appear between paragraphs when locating the
# paragraph that preceded an inserted paragraph mark. They carry no content of
# their own, so crossing them does not change where the merged text belongs.
_TRANSPARENT_PARAGRAPH_MARKERS = frozenset(
    (
        "bookmarkStart",
        "bookmarkEnd",
        "proofErr",
        "commentRangeStart",
        "commentRangeEnd",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
    )
)


def _wordprocessingml_tag(name: str) -> str:
    return f"{{{WORD_PROCESSINGML_NAMESPACE}}}{name}"


def _unwrap_element(element: Any) -> None:
    """Replace an element with its children, preserving document order."""
    parent = element.getparent()
    if parent is None:
        return
    index = parent.index(element)
    for child in list(element):
        parent.insert(index, child)
        index += 1
    parent.remove(element)


def _remove_moved_to_ranges(root: Any) -> None:
    """Remove content bracketed by ``w:moveToRangeStart``/``w:moveToRangeEnd``.

    A move destination may be marked with a range rather than ``w:moveTo``
    wrappers. The pre-revision document has nothing at the destination, so the
    markers and everything between them is dropped. Only ranges whose start and
    end share a parent are handled (the typical case for moved runs); moved
    paragraphs additionally carry paragraph-mark revisions handled separately.
    """
    move_to_range_start = _wordprocessingml_tag("moveToRangeStart")
    move_to_range_end = _wordprocessingml_tag("moveToRangeEnd")
    revision_id_attribute = _wordprocessingml_tag("id")

    for start in list(root.iter(move_to_range_start)):
        parent = start.getparent()
        if parent is None:
            continue
        revision_id = start.get(revision_id_attribute)
        end = None
        for sibling in start.itersiblings():
            if (
                sibling.tag == move_to_range_end
                and sibling.get(revision_id_attribute) == revision_id
            ):
                end = sibling
                break
        if end is None:
            # Malformed range: drop the orphan marker rather than leave a
            # revision element in the output.
            parent.remove(start)
            continue
        doomed = [start]
        node = start.getnext()
        while node is not None and node is not end:
            doomed.append(node)
            node = node.getnext()
        doomed.append(end)
        for node in doomed:
            node.getparent().remove(node)


def _reject_inserted_paragraph_mark(paragraph: Any) -> None:
    """Reject an inserted paragraph mark by merging the paragraph into its predecessor.

    The inserted mark split one original paragraph in two, so the paragraph's
    content is moved into the preceding paragraph and the split paragraph is
    removed. If no preceding paragraph is available (malformed input), the
    marker is stripped and the paragraph is left in place.
    """
    parent = paragraph.getparent()
    if parent is None:
        return

    paragraph_tag = _wordprocessingml_tag("p")
    transparent = {
        _wordprocessingml_tag(name) for name in _TRANSPARENT_PARAGRAPH_MARKERS
    }
    previous = paragraph.getprevious()
    while previous is not None and previous.tag in transparent:
        previous = previous.getprevious()

    if previous is None or previous.tag != paragraph_tag:
        marker = paragraph.find(
            f"{_wordprocessingml_tag('pPr')}/{_wordprocessingml_tag('rPr')}/"
            f"{_wordprocessingml_tag('ins')}"
        )
        if marker is not None:
            marker.getparent().remove(marker)
        return

    paragraph_properties = _wordprocessingml_tag("pPr")
    for child in list(paragraph):
        if child.tag != paragraph_properties:
            previous.append(child)
    parent.remove(paragraph)


def _pre_process_reject_tracked_changes(content: bytes) -> bytes:
    """Rewrite tracked revisions so the DOCX reflects its pre-revision text.

    Equivalent to rejecting all tracked changes before conversion:

    * inserted runs (``w:ins``) are removed;
    * deleted runs (``w:del``) are restored, with ``w:delText`` turned back
      into regular ``w:t`` text;
    * moved content is kept at its original location (``w:moveFrom``) and
      removed from its destination (``w:moveTo``), so it appears exactly once;
    * inserted paragraph marks merge their paragraph back into the preceding
      paragraph, while deleted paragraph marks keep both paragraphs separate;
    * inserted table rows are removed and deleted table rows are restored.

    Only text and paragraph boundaries in the main document body and table
    cells, plus whole table-row revisions, are in scope. Formatting revisions
    (``w:rPrChange``, ``w:pPrChange`` and similar) are left untouched, so the
    current heading, list, emphasis, link, table and equation formatting is
    kept. The transformation operates on an in-memory copy and never modifies
    the input file.
    """
    # Skip the XML round-trip for documents that contain no revision tags.
    if _REVISION_TAG_PATTERN.search(content) is None:
        return content

    from lxml import etree

    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    root = etree.fromstring(content, parser=parser)

    paragraph_tag = _wordprocessingml_tag("p")
    run_properties_tag = _wordprocessingml_tag("rPr")
    table_row_tag = _wordprocessingml_tag("tr")
    table_row_properties_tag = _wordprocessingml_tag("trPr")

    # Whole table-row revisions: an inserted row disappears, while a deleted
    # row is restored by stripping its revision marker.
    for row in list(root.iter(table_row_tag)):
        row_properties = row.find(table_row_properties_tag)
        if row_properties is None:
            continue
        if row_properties.find(_wordprocessingml_tag("ins")) is not None:
            row_parent = row.getparent()
            if row_parent is not None:
                row_parent.remove(row)
        else:
            deleted_row_marker = row_properties.find(_wordprocessingml_tag("del"))
            if deleted_row_marker is not None:
                row_properties.remove(deleted_row_marker)

    # Paragraph-mark revisions live in pPr/rPr.
    paragraphs_to_remove = []
    paragraphs_to_merge = []
    for paragraph in list(root.iter(paragraph_tag)):
        mark_run_properties = paragraph.find(
            f"{_wordprocessingml_tag('pPr')}/{run_properties_tag}"
        )
        if mark_run_properties is None:
            continue
        for marker_name in ("ins", "del", "moveFrom", "moveTo"):
            marker = mark_run_properties.find(_wordprocessingml_tag(marker_name))
            if marker is None:
                continue
            if marker_name == "moveTo":
                # Destination of a moved paragraph: it did not exist here
                # before the revision, so the whole paragraph is dropped.
                paragraphs_to_remove.append(paragraph)
            elif marker_name == "ins":
                # New paragraph mark: merge back into the preceding paragraph.
                paragraphs_to_merge.append(paragraph)
            else:
                # "del" joined two paragraphs in the revised view; rejecting
                # keeps them separate. "moveFrom" marks the source paragraph,
                # which stays where it was. In both cases only the marker goes.
                mark_run_properties.remove(marker)
            break

    for paragraph in paragraphs_to_remove:
        paragraph_parent = paragraph.getparent()
        if paragraph_parent is not None:
            paragraph_parent.remove(paragraph)
    for paragraph in paragraphs_to_merge:
        _reject_inserted_paragraph_mark(paragraph)

    structural_markers = {run_properties_tag, table_row_properties_tag}

    # Inserted content and move destinations are discarded. Markers nested in
    # rPr/trPr are paragraph/row revisions already handled above.
    for marker_name in ("ins", "moveTo"):
        for element in list(root.iter(_wordprocessingml_tag(marker_name))):
            parent = element.getparent()
            if parent is None or parent.tag in structural_markers:
                continue
            parent.remove(element)

    # Deleted content and move sources are restored in place.
    for marker_name in ("del", "moveFrom"):
        for element in list(root.iter(_wordprocessingml_tag(marker_name))):
            parent = element.getparent()
            if parent is None or parent.tag in structural_markers:
                continue
            _unwrap_element(element)

    # Resolve move range markers: drop move-to ranges (content already handled
    # for move-to paragraphs) and strip any remaining range boundary markers.
    _remove_moved_to_ranges(root)
    for marker_name in (
        "moveToRangeStart",
        "moveToRangeEnd",
        "moveFromRangeStart",
        "moveFromRangeEnd",
    ):
        for element in list(root.iter(_wordprocessingml_tag(marker_name))):
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)

    # Deleted and moved-from runs store their text in w:delText instead of
    # w:t; restore it as ordinary text now that the revision wrappers are gone.
    for deleted_text in list(root.iter(_wordprocessingml_tag("delText"))):
        deleted_text.tag = _wordprocessingml_tag("t")

    return etree.tostring(root.getroottree(), encoding="utf-8", xml_declaration=True)


def pre_process_docx(
    input_docx: BinaryIO, *, reject_tracked_changes: bool = False
) -> BinaryIO:
    """
    Pre-processes a DOCX file with provided steps.

    The process works by unzipping the DOCX file in memory, transforming specific XML files
    (such as converting OMML elements to LaTeX), and then zipping everything back into a
    DOCX file without writing to disk.

    Args:
        input_docx (BinaryIO): A binary input stream representing the DOCX file.
        reject_tracked_changes (bool): When True, tracked revisions in the main
            document are rejected before conversion, yielding the pre-revision
            text: insertions are removed, deletions restored, and moved content
            kept at its source. Formatting revisions are not reverted. When
            False (the default), the document is converted as-is.

    Returns:
        BinaryIO: A binary output stream representing the processed DOCX file.
    """
    # Fix ZIP filename casing mismatch before any processing
    input_docx = _fix_zip_filename_casing(input_docx)

    document_steps = [_pre_process_strike, _pre_process_math]
    if reject_tracked_changes:
        document_steps.insert(0, _pre_process_reject_tracked_changes)

    output_docx = BytesIO()
    # The pre-processing steps to apply to each file in the .docx
    pre_process_enable_files = {
        "word/document.xml": tuple(document_steps),
        "word/footnotes.xml": (_pre_process_strike, _pre_process_math),
        "word/endnotes.xml": (_pre_process_strike, _pre_process_math),
        "word/styles.xml": (_pre_process_styles,),
    }
    with zipfile.ZipFile(input_docx, mode="r") as zip_input:
        files = {name: zip_input.read(name) for name in zip_input.namelist()}
        with zipfile.ZipFile(output_docx, mode="w") as zip_output:
            zip_output.comment = zip_input.comment
            for name, content in files.items():
                updated_content = content
                # Each step is applied independently, so one failing step does
                # not discard the results of the others.
                for pre_process_step in pre_process_enable_files.get(name, ()):
                    try:
                        updated_content = pre_process_step(updated_content)
                    except Exception:
                        pass
                zip_output.writestr(name, updated_content)
    output_docx.seek(0)
    return output_docx
