"""Core logic for the TestLink -> WE/TelecomEgypt formatter.

No GUI imports here on purpose: this module can be run and tested on its own,
and main.py just calls run_conversion() from a background thread.

Pipeline (same as the original skill):
    1. parse_source.parse()      TestLink HTML-flavoured .doc -> structured dict
    2. build_docx.build_docx_from_data()   dict -> WE-branded .docx
"""
import os
import posixpath
import xml.etree.ElementTree as ET
import zipfile

from parse_source import parse
from build_docx import build_docx_from_data, generate_output_filename, TemplateError

OLE_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"  # real (binary) Word .doc files


class ConversionError(Exception):
    """A problem the user can understand and fix. Message is shown as-is."""


def _noop(_msg):
    pass


def _check_source(path):
    """Fail early, with a plain-language message, if the input isn't usable."""
    if not path or not os.path.isfile(path):
        raise ConversionError("The selected input file could not be found.")

    with open(path, "rb") as f:
        head = f.read(4096)

    if head.startswith(OLE_MAGIC):
        raise ConversionError(
            "This is a real binary Word (.doc) file, not a TestLink export.\n"
            "TestLink exports are actually HTML saved with a .doc extension. "
            "Please pick the file straight from TestLink's Test Specification export."
        )

    lowered = head.lower()
    if b"<html" not in lowered and b"<!doctype" not in lowered and b"<table" not in lowered:
        raise ConversionError(
            "This file doesn't look like a TestLink Test Specification export "
            "(it isn't an HTML document)."
        )


def _verify_docx(path):
    """Self-check of the finished .docx, so a file Word can't open is never handed over.

    Checks the package the way Word does at load time: required parts exist, every XML
    part is well-formed, every internal relationship points at a real part, and every
    part has a content type. Returns a list of problems (empty = fine).
    """
    problems = []
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        return ["the file is not a valid ZIP package"]
    names = set(zf.namelist())

    for req in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
        if req not in names:
            problems.append(f"required part missing: {req}")

    for n in sorted(names):
        if n.endswith((".xml", ".rels")):
            try:
                ET.fromstring(zf.read(n))
            except ET.ParseError as e:
                problems.append(f"{n} is not well-formed XML ({e})")

    if problems:
        return problems

    for n in sorted(names):
        if not n.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(n))  # folder of the part the rels belong to
        for rel in ET.fromstring(zf.read(n)):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "")
            resolved = target[1:] if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
            if resolved not in names:
                problems.append(f"{n} points to a missing part: {resolved}")

    ct = ET.fromstring(zf.read("[Content_Types].xml"))
    defaults = {e.get("Extension", "").lower() for e in ct if e.tag.endswith("Default")}
    overrides = {e.get("PartName", "") for e in ct if e.tag.endswith("Override")}
    for n in names:
        if n == "[Content_Types].xml":
            continue
        ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
        if ext not in defaults and "/" + n not in overrides:
            problems.append(f"no content type declared for: {n}")
    return problems


def _normalize_with_word(path, log=_noop):
    """Ask Microsoft Word to open and re-save the .docx so the package becomes
    a native Word file. Many enterprise upload systems only accept packages
    that Word itself has written.

    Returns True if Word successfully rewrote the file, False if Word is not
    available or the step failed (the original file is left unchanged).
    """
    try:
        import win32com.client  # type: ignore
        import pythoncom
    except ImportError:
        log("Microsoft Word automation not available (pywin32 missing). Skipping native re-save.")
        return False

    word = None
    doc = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        abs_path = os.path.abspath(path)
        # Open and immediately SaveAs as .docx (FileFormat 16 = wdFormatXMLDocument).
        # This forces Word to rewrite the entire OOXML package in its native layout.
        doc = word.Documents.Open(abs_path, ReadOnly=False, AddToRecentFiles=False)
        doc.SaveAs2(abs_path, FileFormat=16)
        doc.Close(SaveChanges=False)
        doc = None
        log("Re-saved with Microsoft Word so the file is accepted by strict upload systems.")
        return True
    except Exception as e:
        log(f"Could not re-save with Microsoft Word ({e}). The generated file is still usable in Word.")
        return False
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def run_conversion(source_path, output_dir=None, log=_noop):
    """Convert one TestLink export. Returns the full path of the .docx created.

    source_path : the TestLink .doc/.html export
    output_dir  : folder for the result (defaults to the source file's folder)
    log         : callable taking a string, for progress messages
    """
    _check_source(source_path)

    output_dir = output_dir or os.path.dirname(os.path.abspath(source_path))
    try:
        os.makedirs(output_dir, exist_ok=True)
    except OSError as e:
        raise ConversionError(f"Can't use the output folder:\n{output_dir}\n({e})")

    log(f"Reading {os.path.basename(source_path)} ...")
    data = parse(source_path)

    suites = data.get("suites", [])
    n_cases = sum(len(s.get("test_cases", [])) for s in suites)
    if not suites or n_cases == 0:
        raise ConversionError(
            "No test suites or test cases were found in this file.\n"
            "Make sure it is a TestLink 'Test Specification' export."
        )
    log(f"Found {len(suites)} test suite(s) and {n_cases} test case(s).")
    log("Removed TestLink branding (logo and 'Community [configure ...]' line).")

    out_name = generate_output_filename(data)
    out_path = os.path.join(output_dir, out_name)

    log("Building the WE / TelecomEgypt Word document ...")
    try:
        build_docx_from_data(data, out_path)
    except TemplateError as e:
        raise ConversionError(str(e))
    except PermissionError:
        raise ConversionError(
            f"Couldn't write:\n{out_path}\n\n"
            "If that file is open in Word, close it and try again."
        )

    log("Checking the finished file ...")
    problems = _verify_docx(out_path)
    if problems:
        try:
            os.remove(out_path)
        except OSError:
            pass
        raise ConversionError(
            "The Word file failed its self-check, so it was not kept.\n\n- "
            + "\n- ".join(problems[:6])
            + "\n\nPlease send this message to whoever maintains this tool."
        )

    # Final step: if Microsoft Word is installed, let it rewrite the package.
    # This is what makes the file pass strict enterprise upload validators.
    log("Normalizing the package with Microsoft Word (if available) ...")
    _normalize_with_word(out_path, log=log)

    log(f"Done. Saved to {out_path}")
    return out_path
