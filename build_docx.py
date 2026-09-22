#!/usr/bin/env python3
"""
Build a WE/TelecomEgypt-branded .docx from parsed TestLink JSON, using the
bundled template skeleton (header/footer/logos/styles) for everything
except the body content, which is generated fresh so it can handle any
number of suites / test cases / steps.

Usage:
    python3 build_docx.py <parsed.json> [output.docx] [skeleton_dir]

When output.docx is omitted, the filename is generated from the Test Suite's
CR identifier, e.g. "Message Broker13 Cloud - CR 6894 RestartOnFTTH" becomes
"CR 6894-RestartOnFTTH.docx".
"""
import sys
import os
import json
import shutil
import hashlib
import tempfile
import zipfile
import re
from xml.sax.saxutils import escape as xml_escape

from paths import resource_path

# Bundled docx skeleton (header/footer/logos/styles). Resolved via
# resource_path() so it works both as a plain script and inside the frozen EXE.
DEFAULT_SKELETON = resource_path(os.path.join("resources", "skeleton"))


# Characters that XML 1.0 forbids (control characters, lone surrogates, U+FFFE/U+FFFF).
# One of these in the text makes document.xml not well-formed, and Word then refuses
# to open the file at all ("we found a problem with its contents"). Text pasted from
# Word or web pages into TestLink sometimes carries them (e.g. a vertical tab).
_XML_INVALID = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")


def clean_text(s):
    s = _XML_INVALID.sub(lambda m: " " if m.group(0) in "\x0b\x0c" else "", s or "")
    return s


class TemplateError(Exception):
    """The bundled Word template is incomplete or has been changed."""


# Names under which two special DOCX parts are stored inside resources/skeleton
# (their real names start with '[' or '.', which browsers, ZIP tools and GitHub
# uploads often mangle or silently skip). They are restored inside the generated file.
_RESTORE_NAMES = {
    "Content_Types.xml": "[Content_Types].xml",
    "_rels/rels_root.xml": "_rels/.rels",
}


def _sha256(data):
    return hashlib.sha256(data).hexdigest()


def prepare_package(skeleton_dir, pkg_dir):
    """Copy the template into pkg_dir, checking it against skeleton_manifest.json.

    * a missing file  -> TemplateError (never silently build a broken document)
    * a file whose only difference is CRLF line endings (Git on Windows can add these
      when it checks the repo out) is repaired automatically
    * any other change -> TemplateError
    """
    manifest_path = os.path.join(os.path.dirname(os.path.abspath(skeleton_dir)), "skeleton_manifest.json")
    if not os.path.isfile(manifest_path):
        raise TemplateError("The template checklist (skeleton_manifest.json) is missing from the app's resources folder.")
    with open(manifest_path, encoding="utf-8") as f:
        expected = json.load(f)["files"]

    missing, changed = [], []
    for rel, want in expected.items():
        src = os.path.join(skeleton_dir, *rel.split("/"))
        if not os.path.isfile(src):
            missing.append(rel)
            continue
        with open(src, "rb") as f:
            data = f.read()
        if _sha256(data) != want:
            repaired = data.replace(b"\r\n", b"\n")
            if _sha256(repaired) == want:
                data = repaired
            else:
                changed.append(rel)
                continue
        dst_rel = _RESTORE_NAMES.get(rel, rel)
        dst = os.path.join(pkg_dir, *dst_rel.split("/"))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        with open(dst, "wb") as f:
            f.write(data)

    if missing or changed:
        lines = ["The bundled Word template is incomplete or damaged, so no document was created."]
        if missing:
            lines.append("Missing: " + ", ".join(missing))
        if changed:
            lines.append("Changed: " + ", ".join(changed))
        lines.append("This usually means some files did not upload to GitHub. Upload the whole "
                     "'resources' folder again, then rebuild the app.")
        raise TemplateError("\n".join(lines))


_BM_BAD = re.compile(r"[^A-Za-z0-9_]")


def bm_name(name):
    """Make a bookmark name Word accepts: letters/digits/underscore, starts with a
    letter or underscore, at most 40 characters. Normal TestLink anchors such as
    'toc_tc123' pass through unchanged."""
    n = _BM_BAD.sub("_", (name or "").strip())
    if n and not (n[0].isalpha() or n[0] == "_"):
        n = "_" + n
    return n[:40]


def esc(s):
    return xml_escape(clean_text(s))


def esc_attr(s):
    """Escape for use inside a double-quoted XML attribute value."""
    return xml_escape(clean_text(s), {'"': "&quot;"})


def generate_output_filename(data):
    """Generate a safe DOCX filename from the CR portion of Test Suite.

    Examples:
        "Message Broker13 Cloud - CR 6894 RestartOnFTTH"
            -> "CR 6894-RestartOnFTTH.docx"
        "Message Broker13 Cloud - CR#6481-getFullAccess"
            -> "CR 6481-getFullAccess.docx"

    If no CR pattern is found, fall back to the Test Suite text.
    """
    suite = (data.get("test_suite") or "").strip()

    match = re.search(r"\bCR\s*#?\s*(\d+)\s*(?:[-–—:]\s*|\s+)(.+?)\s*$", suite, re.IGNORECASE)
    if match:
        cr_number = match.group(1)
        description = re.sub(r"^[-–—:]\s*", "", match.group(2).strip())
        filename = f"CR {cr_number}-{description}"
    else:
        filename = suite or "Test Specification"

    # Remove characters that Windows does not allow in filenames.
    filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', filename)
    filename = re.sub(r"\s+", " ", filename).strip(" .")
    return f"{filename}.docx"


# ---- paragraph / run building blocks, styles copied verbatim from the
# ---- reference WE/TelecomEgypt document ----

RPR_36 = '<w:rPr><w:rFonts w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>'
RPR_18 = '<w:rPr><w:rFonts w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
RPR_16 = '<w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
RPR_22 = '<w:rPr><w:rFonts w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
RPR_23 = '<w:rPr><w:rFonts w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="23"/><w:szCs w:val="23"/></w:rPr>'
RPR_18_BOLD = '<w:rPr><w:rFonts w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:bCs/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'


def p_normalweb_centered(text, rpr=RPR_36):
    return (f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/><w:jc w:val="center"/>{rpr}</w:pPr>'
            f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')


def p_normalweb(text, rpr=RPR_36, br_after=None):
    br = '<w:br/>' if br_after is not None else ''
    extra = f'<w:t xml:space="preserve">{esc(br_after)}</w:t>' if br_after is not None else ''
    return (f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>{rpr}</w:pPr>'
            f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
            f'<w:r>{rpr}{br}{extra}</w:r></w:p>')


def p_plain(text, rpr=RPR_18, page_break=False):
    pb = '<w:r><w:br w:type="page"/></w:r>' if page_break else ''
    return (f'<w:p><w:pPr><w:pStyle w:val="Normal"/>{rpr}</w:pPr>'
            f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>{pb}</w:p>')


def p_heading1(text, rpr, bookmark=None):
    bm = ""
    if bookmark:
        bm = (f'<w:bookmarkStart w:id="{bookmark[1]}" w:name="{esc_attr(bm_name(bookmark[0]))}"/>'
              f'<w:bookmarkEnd w:id="{bookmark[1]}"/>')
    return (f'<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:ind w:hanging="0" w:start="0"/>{rpr}</w:pPr>'
            f'{bm}<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')


def p_toc_entry(anchor, text, bold):
    rpr = RPR_18_BOLD if bold else RPR_18
    anchor = bm_name(anchor)
    if not anchor:  # no usable link target: keep the text, skip the hyperlink
        return (f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>{rpr}</w:pPr>'
                f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')
    return (f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>{rpr}</w:pPr>'
            f'<w:hyperlink w:anchor="{esc_attr(anchor)}">'
            f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:rFonts w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/>'
            f'{"<w:b/><w:bCs/>" if bold else ""}<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:hyperlink></w:p>')


def p_bookmark_placeholder(anchor, bm_id):
    anchor = bm_name(anchor)
    if not anchor:
        return p_spacer_normalweb()
    return (f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>{RPR_18}</w:pPr>'
            f'<w:bookmarkStart w:id="{bm_id}" w:name="{esc_attr(anchor)}"/><w:bookmarkEnd w:id="{bm_id}"/>'
            f'<w:r>{RPR_18}<w:t xml:space="preserve"> </w:t></w:r></w:p>')


def p_spacer_normalweb():
    return f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>{RPR_18}</w:pPr><w:r>{RPR_18}</w:r></w:p>'


# ---- table cell helpers ----

TC_BORDERS = ('<w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/>'
              '<w:start w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/>'
              '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/>'
              '<w:end w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders>')
TC_END_BORDER = '<w:tcBorders><w:end w:val="single" w:sz="6" w:space="0" w:color="DDDDDD"/></w:tcBorders>'
RPR_17 = '<w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr>'
RPR_17_BOLD = '<w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:bCs/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr>'
RPR_14_BOLD = '<w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:bCs/><w:sz w:val="14"/><w:szCs w:val="14"/></w:rPr>'
RPR_LABEL = '<w:rPr><w:rStyle w:val="label"/><w:rFonts w:eastAsia="Times New Roman" w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr>'


def tc(content, width, gridspan=None, shade=None, borders=TC_BORDERS, valign=True):
    span = f'<w:gridSpan w:val="{gridspan}"/>' if gridspan else ''
    shd = f'<w:shd w:fill="{shade}" w:val="clear"/>' if shade else ''
    va = '<w:vAlign w:val="center"/>' if valign else ''
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{span}{borders}{shd}{va}</w:tcPr>'
            f'{content}</w:tc>')


def p_label(text):
    return (f'<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:before="75" w:after="75"/>{RPR_17}</w:pPr>'
            f'<w:r>{RPR_LABEL}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')


def p_value(text):
    return (f'<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:before="75" w:after="75"/>{RPR_17}</w:pPr>'
            f'<w:r>{RPR_17}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')


def p_value_plain(text):
    return (f'<w:p><w:pPr><w:pStyle w:val="Normal"/>{RPR_17}</w:pPr>'
            f'<w:r>{RPR_17}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')


def p_multiline_value(text):
    """Summary / Preconditions style value paragraph (NormalWeb, spacing before=280)."""
    rpr = RPR_17
    return (f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/><w:spacing w:before="280" w:after="0"/>{rpr}</w:pPr>'
            f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')


def build_test_case_table(tc_data, widths):
    """widths = (col1, col2, col3, col4, col5, total) matching the target's
    gridCol layout for this test case's step table."""
    c1, c2, c3, c4, c5, total = widths
    rows = []

    # Header row
    header_content = (
        f'<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:before="75" w:after="75"/>{RPR_17_BOLD}</w:pPr>'
        f'<w:r>{RPR_17_BOLD}<w:t xml:space="preserve">Test Case {esc(tc_data["title"])} </w:t></w:r>'
        f'<w:r>{RPR_14_BOLD}<w:t xml:space="preserve">[Version : {esc(tc_data.get("version","1"))}]</w:t></w:r></w:p>'
    )
    rows.append(f'<w:tr><w:trPr/>{tc(header_content, total, gridspan=4, shade="EEEEEE")}</w:tr>')

    # Author
    rows.append('<w:tr><w:trPr/>' + tc(p_label("Author:"), c1) + tc(p_value(tc_data.get("author", "")), total - c1, gridspan=3) + '</w:tr>')

    # Summary
    summary_content = p_label("Summary:") + p_multiline_value(tc_data.get("summary", ""))
    rows.append('<w:tr><w:trPr/>' + tc(summary_content, total, gridspan=4) + '</w:tr>')

    # Preconditions (optional)
    if tc_data.get("preconditions"):
        precond_content = p_label("Preconditions:") + p_multiline_value(tc_data["preconditions"])
        rows.append('<w:tr><w:trPr/>' + tc(precond_content, total, gridspan=4) + '</w:tr>')

    # Steps header
    rows.append(
        '<w:tr><w:trPr/>'
        + tc(p_label("#:"), c1)
        + tc(p_label("Step actions:"), c2)
        + tc(p_label("Expected Results:"), c3)
        + tc('<w:p><w:pPr><w:pStyle w:val="Normal"/><w:snapToGrid w:val="false"/>'
             '<w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/>'
             '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
             '<w:r><w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/>'
             '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:r></w:p>', c4, borders=TC_END_BORDER)
        + '</w:tr>'
    )

    # Step rows
    for step in tc_data.get("steps", []):
        actions_paras = []
        lines = step.get("actions", [])
        for idx, line in enumerate(lines):
            is_last = idx == len(lines) - 1
            spacing = ' w:before="0" w:after="280"' if is_last else ''
            if line.get("rtl"):
                sp = ("<w:spacing" + spacing + "/>") if spacing else ""
                pPr = f'<w:pPr><w:pStyle w:val="NormalWeb"/><w:bidi w:val="1"/>{sp}<w:ind w:end="0"/><w:jc w:val="end"/>{RPR_17}</w:pPr>'
            else:
                pPr = f'<w:pPr><w:pStyle w:val="NormalWeb"/>{("<w:spacing" + spacing + "/>") if spacing else ""}{RPR_17}</w:pPr>'
            actions_paras.append(f'<w:p>{pPr}<w:r>{RPR_17}<w:t xml:space="preserve">{esc(line["text"])}</w:t></w:r></w:p>')
        actions_content = "".join(actions_paras) if actions_paras else p_value_plain("")

        expected_paras = []
        elines = step.get("expected", [])
        for idx, line in enumerate(elines):
            is_last = idx == len(elines) - 1
            rpr = RPR_17
            spacing = '<w:spacing w:before="280" w:after="0"/>' if is_last else ''
            expected_paras.append(f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>{spacing}{rpr}</w:pPr>'
                                   f'<w:r>{rpr}<w:t xml:space="preserve">{esc(line["text"])}</w:t></w:r></w:p>')
        expected_content = "".join(expected_paras) if expected_paras else p_value_plain("")

        num_content = (f'<w:p><w:pPr><w:pStyle w:val="Normal"/>{RPR_17}</w:pPr>'
                       f'<w:r>{RPR_17}<w:t xml:space="preserve">{esc(step.get("num",""))}</w:t></w:r></w:p>')
        blank_narrow = ('<w:p><w:pPr><w:pStyle w:val="Normal"/><w:snapToGrid w:val="false"/>'
                        '<w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/>'
                        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
                        '<w:r><w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/>'
                        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:r></w:p>')
        rows.append(
            '<w:tr><w:trPr/>'
            + tc(num_content, c1)
            + tc(actions_content, c2)
            + tc(expected_content, c3)
            + tc(blank_narrow, c4, borders=TC_END_BORDER)
            + '</w:tr>'
        )

    # Execution type / duration / importance
    rows.append('<w:tr><w:trPr/>' + tc(p_label("Execution type:"), c1) + tc(p_value_plain(tc_data.get("execution_type", "")), total - c1, gridspan=3) + '</w:tr>')
    rows.append('<w:tr><w:trPr/>' + tc(p_label("Estimated exec. duration (min):"), c1) + tc(p_value_plain(tc_data.get("estimated_duration", "")), total - c1, gridspan=3) + '</w:tr>')
    rows.append('<w:tr><w:trPr/>' + tc(p_label("Importance:"), c1) + tc(p_value_plain(tc_data.get("importance", "")), total - c1, gridspan=3) + '</w:tr>')

    # Spacer row
    rows.append('<w:tr><w:trPr/>' + tc(p_value_plain(""), total, gridspan=4) + '</w:tr>')

    # Requirements / Keywords
    rows.append('<w:tr><w:trPr/>' + tc(p_label("Requirements"), c1) + tc(p_value_plain(" " + tc_data.get("requirements", "")), total - c1, gridspan=3) + '</w:tr>')
    rows.append('<w:tr><w:trPr/>' + tc(p_label("Keywords:"), c1) + tc(p_value_plain(" " + tc_data.get("keywords", "")), total - c1, gridspan=3) + '</w:tr>')

    table_xml = (
        '<w:tbl><w:tblPr><w:tblW w:w="4750" w:type="pct"/><w:jc w:val="start"/>'
        '<w:tblInd w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:start w:w="30" w:type="dxa"/>'
        '<w:bottom w:w="30" w:type="dxa"/><w:end w:w="30" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{c1}"/><w:gridCol w:w="{c2}"/><w:gridCol w:w="{c3}"/><w:gridCol w:w="{c4}"/><w:gridCol w:w="{c5}"/></w:tblGrid>'
        + "".join(rows) +
        '</w:tbl>'
    )
    return table_xml


def build_body(data):
    parts = []

    # Title block
    parts.append(p_plain(""))
    parts.append(p_plain(""))
    parts.append(p_plain(""))
    # gray horizontal rule (drawing shape, copied verbatim from reference)
    parts.append(
        '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" '
        'w:eastAsia="Times New Roman" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Arial" w:ascii="Arial" w:hAnsi="Arial"/>'
        '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5274310" cy="19050"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name=""/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5274360" cy="19080"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="a0a0a0"/></a:solidFill>'
        '<a:ln w="0"><a:noFill/></a:ln></wps:spPr><wps:style><a:lnRef idx="0"/><a:fillRef idx="0"/>'
        '<a:effectRef idx="0"/><a:fontRef idx="minor"/></wps:style><wps:bodyPr/></wps:wsp></a:graphicData>'
        '</a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict>'
        '<v:rect id="shape_0" fillcolor="#a0a0a0" stroked="f" o:allowincell="f" '
        'style="position:absolute;margin-left:0pt;margin-top:-1.55pt;width:415.25pt;height:1.45pt;'
        'mso-wrap-style:none;v-text-anchor:middle;mso-position-vertical:top">'
        '<v:fill o:detectmouseclick="t" type="solid" color2="#5f5f5f"/>'
        '<v:stroke color="#3465a4" joinstyle="round" endcap="flat"/><w10:wrap type="square"/></v:rect>'
        '</w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>'
    )
    parts.append(p_normalweb_centered(""))
    parts.append(p_normalweb_centered(data.get("spec_label", "Test Specification")))
    parts.append(p_normalweb(f'Test Project: {data.get("test_project","")}', br_after=f'Test Suite: {data.get("test_suite","")}'))
    parts.append(f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>{RPR_18}</w:pPr><w:r>{RPR_18}<w:t xml:space="preserve">{esc(data.get("printed_by",""))}</w:t></w:r></w:p>')
    parts.append(p_plain(data.get("copyright", ""), rpr=RPR_16, page_break=True))

    # Table of Contents
    parts.append(p_heading1("Table Of Contents", RPR_22))
    for entry in data.get("toc", []):
        parts.append(p_toc_entry(entry["anchor"], entry["text"], entry["bold"]))
    # page break after last toc entry -> attach to last paragraph output;
    # simplest: add explicit page break paragraph
    if data.get("toc"):
        parts[-1] = parts[-1].replace("</w:p>", '<w:r><w:br w:type="page"/></w:r></w:p>', 1) if False else parts[-1]
    parts.append('<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>')

    # Suites
    bm_id = 10
    for suite in data.get("suites", []):
        parts.append(p_heading1(suite["title"], RPR_23, bookmark=("toc_", bm_id)))
        bm_id += 1
        for tc_data in suite.get("test_cases", []):
            parts.append(p_bookmark_placeholder(tc_data.get("anchor", ""), bm_id))
            bm_id += 1
            parts.append(build_test_case_table(tc_data, (1574, 4235, 1998, 66, 17, 7890)))
            parts.append(p_spacer_normalweb())

    parts.append('<w:p><w:pPr><w:pStyle w:val="Normal"/><w:rPr><w:rFonts w:eastAsia="Times New Roman"/></w:rPr></w:pPr>'
                 '<w:r><w:rPr><w:rFonts w:eastAsia="Times New Roman"/></w:rPr></w:r></w:p>')

    sect_pr = (
        '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/>'
        '<w:footerReference w:type="default" r:id="rId3"/><w:type w:val="nextPage"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:left="1800" w:right="1800" w:gutter="0" w:header="720" w:top="1440" w:footer="720" w:bottom="1440"/>'
        '<w:pgNumType w:fmt="decimal"/><w:formProt w:val="false"/><w:textDirection w:val="lrTb"/>'
        '<w:docGrid w:type="default" w:linePitch="299" w:charSpace="0"/></w:sectPr>'
    )

    body = "".join(parts) + sect_pr

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<w:document xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2012/wordml" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 wp14 w15">'
        f'<w:body>{body}</w:body></w:document>'
    )
    return document_xml


def build_docx_from_data(data, output_path, skeleton_dir=None):
    """Build the branded .docx from an already-parsed dict (see parse_source.parse).

    The package is assembled in a temporary folder and only written to
    output_path once it is complete, so a failure never leaves a half-built
    file next to the user's documents.
    """
    skeleton_dir = skeleton_dir or DEFAULT_SKELETON
    if not os.path.isdir(skeleton_dir):
        raise TemplateError(f"The bundled Word template folder is missing: {skeleton_dir}")

    work_dir = tempfile.mkdtemp(prefix="testlink_we_")
    try:
        pkg_dir = os.path.join(work_dir, "pkg")
        os.makedirs(pkg_dir)
        prepare_package(skeleton_dir, pkg_dir)   # checks + copies, restores special names
        docx_ct = os.path.join(pkg_dir, "[Content_Types].xml")

        document_xml = build_body(data)
        os.makedirs(os.path.join(pkg_dir, "word"), exist_ok=True)
        with open(os.path.join(pkg_dir, "word", "document.xml"), "w", encoding="utf-8", newline="") as f:
            f.write(document_xml)

        tmp_docx = os.path.join(work_dir, "out.docx")
        with zipfile.ZipFile(tmp_docx, "w", zipfile.ZIP_DEFLATED) as zf:
            # [Content_Types].xml first, as Word expects.
            zf.write(docx_ct, "[Content_Types].xml")
            for root, _, files in os.walk(pkg_dir):
                for name in sorted(files):
                    full = os.path.join(root, name)
                    rel = os.path.relpath(full, pkg_dir).replace(os.sep, "/")
                    if rel == "[Content_Types].xml":
                        continue
                    zf.write(full, rel)   # ZIP entry names always use '/' 

        if os.path.exists(output_path):
            os.remove(output_path)  # PermissionError here => file open in Word
        shutil.move(tmp_docx, output_path)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def build_docx(json_path, output_path, skeleton_dir=None):
    """CLI-compatible wrapper: build from a parsed JSON file."""
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    build_docx_from_data(data, output_path, skeleton_dir)
    print(f"Wrote {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: build_docx.py <parsed.json> [output.docx] [skeleton_dir]", file=sys.stderr)
        sys.exit(1)

    json_path = sys.argv[1]
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        output_dir = os.path.dirname(os.path.abspath(json_path))
        output_path = os.path.join(output_dir, generate_output_filename(data))

    skeleton_dir = sys.argv[3] if len(sys.argv) > 3 else None
    build_docx(json_path, output_path, skeleton_dir)
