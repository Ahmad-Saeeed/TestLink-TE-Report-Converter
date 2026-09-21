"""Core logic for the TestLink -> WE/TelecomEgypt formatter.

No GUI imports here on purpose: this module can be run and tested on its own,
and main.py just calls run_conversion() from a background thread.

Pipeline (same as the original skill):
    1. parse_source.parse()      TestLink HTML-flavoured .doc -> structured dict
    2. build_docx.build_docx_from_data()   dict -> WE-branded .docx
"""
import os

from parse_source import parse
from build_docx import build_docx_from_data, generate_output_filename

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
    except PermissionError:
        raise ConversionError(
            f"Couldn't write:\n{out_path}\n\n"
            "If that file is open in Word, close it and try again."
        )

    log(f"Done. Saved to {out_path}")
    return out_path
