#!/usr/bin/env python3
"""
Parse a TestLink "Test Specification" export (saved as .doc, actually HTML)
into structured JSON, stripping the TestLink branding block (logo + the
"<Product> Community [configure ...]" line + horizontal rule).

Usage:
    python3 parse_source.py <source.doc-or-.html> <output.json>
"""
import sys
import json
import re
from bs4 import Tag


def cell_lines(td):
    """Extract a table cell's content as a list of {"text", "rtl"} lines,
    one per <p> child (direct-child paragraphs), preserving reading order
    and whether TestLink marked the paragraph dir="RTL" (this shows up on
    ad-hoc bullet-style parameter lists and should render right-aligned,
    matching the original)."""
    if td is None:
        return []
    lines = []
    ps = td.find_all("p", recursive=False)
    if ps:
        for p in ps:
            text = p.get_text(" ", strip=True)
            text = re.sub(r"\s+", " ", text).strip()
            if text:
                lines.append({"text": text, "rtl": p.get("dir", "").upper() == "RTL"})
    if not lines:
        text = td.get_text(" ", strip=True)
        text = re.sub(r"\s+", " ", text).strip()
        if text:
            lines.append({"text": text, "rtl": False})
    return lines


def strip_label(td, label_text):
    """Full text of td (including any malformed nested <td>), with the
    leading label (e.g. 'Summary:', 'Requirements') removed."""
    if td is None:
        return ""
    text = td.get_text(" ", strip=True)
    text = re.sub(r"\s+", " ", text).strip()
    label_core = label_text.rstrip(":").strip()
    text = re.sub(r"^" + re.escape(label_core) + r"\s*:?\s*", "", text, flags=re.IGNORECASE)
    return text.strip()


def label_value(table, label_text):
    """Find a <tr> whose first cell's label matches label_text and
    return the sibling/nested value cell (as a Tag) or None. Handles
    TestLink's malformed markup where the value <td> can end up nested
    inside the label <td> (unclosed tag) instead of a sibling."""
    for tr in table.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if not tds:
            continue
        first = tds[0]
        span = first.find("span", class_="label")
        label = (span.get_text(strip=True) if span else first.get_text(strip=True))
        if label.rstrip(":").strip().lower() == label_text.rstrip(":").strip().lower():
            if len(tds) > 1:
                return tds[1]
            nested = first.find("td")
            if nested is not None:
                return nested
            return first
    return None


def parse_test_case(table):
    tc = {}
    th = table.find("th")
    header_text = th.get_text(" ", strip=True) if th else ""
    m = re.match(r"Test Case\s+(.*?)\s*\[Version\s*:\s*(\d+)\]\s*$", header_text)
    if m:
        tc["title"] = m.group(1).strip()
        tc["version"] = m.group(2).strip()
    else:
        tc["title"] = header_text
        tc["version"] = ""

    author_cell = label_value(table, "Author:")
    tc["author"] = author_cell.get_text(strip=True) if author_cell else ""

    summary_cell = label_value(table, "Summary:")
    tc["summary"] = strip_label(summary_cell, "Summary:") if summary_cell else ""

    precond_cell = label_value(table, "Preconditions:")
    tc["preconditions"] = strip_label(precond_cell, "Preconditions:") if precond_cell else ""

    exec_type_cell = label_value(table, "Execution type:")
    tc["execution_type"] = exec_type_cell.get_text(strip=True) if exec_type_cell else ""

    duration_cell = label_value(table, "Estimated exec. duration (min):")
    tc["estimated_duration"] = duration_cell.get_text(strip=True) if duration_cell else ""

    importance_cell = label_value(table, "Importance:")
    tc["importance"] = importance_cell.get_text(strip=True) if importance_cell else ""

    req_cell = label_value(table, "Requirements")
    tc["requirements"] = strip_label(req_cell, "Requirements") if req_cell else ""

    kw_cell = label_value(table, "Keywords:")
    tc["keywords"] = strip_label(kw_cell, "Keywords:") if kw_cell else ""

    steps = []
    for tr in table.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if len(tds) == 3:
            first_span = tds[0].find("span", class_="label")
            if first_span:
                continue  # header row
            step_num = tds[0].get_text(strip=True)
            actions = cell_lines(tds[1])
            expected = cell_lines(tds[2])
            if step_num or actions or expected:
                steps.append({"num": step_num, "actions": actions, "expected": expected})
    tc["steps"] = steps
    return tc


def parse(html_path):
    from bs4 import BeautifulSoup
    with open(html_path, encoding="utf-8", errors="replace") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    data = {}

    title_tag = soup.find("title")
    data["doc_title"] = title_tag.get_text(strip=True) if title_tag else ""

    doc_titles = soup.find_all("div", class_="doc_title")
    data["spec_label"] = doc_titles[0].get_text(strip=True) if len(doc_titles) > 0 else "Test Specification"

    project = ""
    suite = ""
    if len(doc_titles) > 1:
        p = doc_titles[1].find("p")
        if p:
            html_inner = str(p)
            parts = re.split(r"<br\s*/?>", html_inner)
            texts = [BeautifulSoup(part, "html.parser").get_text(strip=True) for part in parts]
            for t in texts:
                if t.lower().startswith("test project:"):
                    project = t.split(":", 1)[1].strip()
                elif t.lower().startswith("test suite:"):
                    suite = t.split(":", 1)[1].strip()
    data["test_project"] = project
    data["test_suite"] = suite

    printedby = soup.find(id="printedby")
    data["printed_by"] = printedby.get_text(strip=True) if printedby else ""

    copyright_tag = soup.find(id="copyright")
    data["copyright"] = copyright_tag.get_text(strip=True) if copyright_tag else ""

    toc_div = soup.find("div", class_="toc")
    toc_entries = []
    if toc_div:
        for p in toc_div.find_all("p", recursive=True):
            a = p.find("a")
            if not a:
                continue
            href = a.get("href", "").lstrip("#")
            text = a.get_text(strip=True)
            is_bold = p.find_parent("b") is not None or p.find("b") is not None
            toc_entries.append({"anchor": href, "text": text, "bold": is_bold})
    data["toc"] = toc_entries

    suites = []
    doclevel_headers = soup.find_all("h1", class_="doclevel")
    for h in doclevel_headers:
        suite_data = {"title": h.get_text(strip=True), "test_cases": []}
        node = h.find_next_sibling()
        current_anchor = None
        while node is not None and not (isinstance(node, Tag) and node.name == "h1" and "doclevel" in (node.get("class") or [])):
            if isinstance(node, Tag):
                if node.name == "a" and node.get("name", "").startswith("toc_tc"):
                    current_anchor = node.get("name")
                elif node.name == "table" and "tc" in (node.get("class") or []):
                    tc = parse_test_case(node)
                    tc["anchor"] = current_anchor or ""
                    suite_data["test_cases"].append(tc)
                elif node.name == "div":
                    inner_table = node.find("table", class_="tc")
                    if inner_table:
                        tc = parse_test_case(inner_table)
                        tc["anchor"] = current_anchor or ""
                        suite_data["test_cases"].append(tc)
            node = node.find_next_sibling()
        suites.append(suite_data)
    data["suites"] = suites

    return data


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: parse_source.py <source.doc> <output.json>", file=sys.stderr)
        sys.exit(1)
    result = parse(sys.argv[1])
    with open(sys.argv[2], "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    n_tc = sum(len(s["test_cases"]) for s in result.get("suites", []))
    print(f"Parsed {len(result.get('suites', []))} suite(s), {n_tc} test case(s) -> {sys.argv[2]}")
