#!/usr/bin/env python3
"""
Sovereign Doc – Core Library

Local, dependency-free document liberation + basic reformatting core.
(See README for full description.)
"""

import csv
import time
from datetime import datetime
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

DOCX_MAIN_XML = "word/document.xml"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
LOG_FILENAME = "sovereign_doc_log.csv"


class SovereignDocError(Exception):
    pass


def _unique_path(base_path: Path) -> Path:
    candidate = base_path
    counter = 1
    while candidate.exists():
        candidate = base_path.with_name(f"{base_path.stem} ({counter}){base_path.suffix}")
        counter += 1
    return candidate


def _log_conversion(input_path, output_path, src_fmt, dst_fmt, mode, duration_ms, chars_out, omega_score):
    here = Path(__file__).resolve().parent
    log_path = here / LOG_FILENAME
    new_file = not log_path.exists()
    row = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "input_path": str(input_path),
        "output_path": str(output_path),
        "src_format": src_fmt,
        "dst_format": dst_fmt,
        "mode": mode,
        "duration_ms": f"{duration_ms:.1f}",
        "chars_out": chars_out,
        "omega_score": f"{omega_score:.3f}",
    }
    fieldnames = list(row.keys())
    try:
        with log_path.open("a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if new_file:
                writer.writeheader()
            writer.writerow(row)
    except Exception:
        pass


def extract_docx_text(docx_path: Path) -> str:
    if not docx_path.is_file():
        raise SovereignDocError(f"Input file does not exist: {docx_path}")
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            xml_bytes = zf.read(DOCX_MAIN_XML)
    except Exception as e:
        raise SovereignDocError(f"Unable to open or read {docx_path} as docx") from e
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        raise SovereignDocError(f"Unable to parse XML in {docx_path}") from e
    paragraphs = []
    for p in root.findall(".//w:p", NS):
        texts = []
        for t in p.findall(".//w:t", NS):
            if t.text:
                texts.append(t.text)
        if texts:
            txt = "".join(texts).strip()
            if txt:
                paragraphs.append(txt)
    return ("\n\n".join(paragraphs).strip() + "\n") if paragraphs else ""


def txt_to_markdown(text: str) -> str:
    lines = text.splitlines()
    md_lines = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            md_lines.append("")
            continue
        words = stripped.split()
        if 1 <= len(words) <= 8:
            upperish = sum(1 for w in words if w[:1].isupper())
            if upperish >= max(1, len(words) // 2):
                md_lines.append("# " + stripped)
                continue
        if stripped[0] in ("•", "-", "*"):
            content = stripped.lstrip("•-*").strip()
            md_lines.append(f"- {content}")
            continue
        md_lines.append(stripped)
    return "\n".join(md_lines).rstrip() + "\n"


def _html_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def text_to_html(text: str) -> str:
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    parts = ["<!DOCTYPE html>", "<html>", "<body>"]
    for p in paragraphs:
        parts.append(f"<p>{_html_escape(p)}</p>")
    parts.extend(["</body>", "</html>"])
    return "\n".join(parts) + "\n"


def markdown_to_html(md: str) -> str:
    lines = md.splitlines()
    html_lines = []
    in_list = False

    def close_list():
        nonlocal in_list
        if in_list:
            html_lines.append("</ul>")
            in_list = False

    for line in lines:
        stripped = line.strip()
        if not stripped:
            close_list()
            continue
        if stripped.startswith("### "):
            close_list()
            html_lines.append(f"<h3>{_html_escape(stripped[4:].strip())}</h3>")
        elif stripped.startswith("## "):
            close_list()
            html_lines.append(f"<h2>{_html_escape(stripped[3:].strip())}</h2>")
        elif stripped.startswith("# "):
            close_list()
            html_lines.append(f"<h1>{_html_escape(stripped[2:].strip())}</h1>")
        elif stripped.startswith(("- ", "* ", "+ ")):
            if not in_list:
                html_lines.append("<ul>")
                in_list = True
            html_lines.append(f"<li>{_html_escape(stripped[2:].strip())}</li>")
        else:
            close_list()
            html_lines.append(f"<p>{_html_escape(stripped)}</p>")
    close_list()
    return "<!DOCTYPE html>\n<html>\n<body>\n" + "\n".join(html_lines) + "\n</body>\n</html>\n"


def _create_basic_docx_xml(paragraphs):
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ET.register_namespace("w", w_ns)
    document = ET.Element(f"{{{w_ns}}}document")
    body = ET.SubElement(document, f"{{{w_ns}}}body")
    for text in paragraphs:
        p = ET.SubElement(body, f"{{{w_ns}}}p")
        r = ET.SubElement(p, f"{{{w_ns}}}r")
        t = ET.SubElement(r, f"{{{w_ns}}}t")
        t.text = text
    return ET.tostring(document, encoding="utf-8", xml_declaration=True)


def text_to_docx_bytes(text: str) -> dict:
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    if not paragraphs:
        paragraphs = [""]
    document_xml = _create_basic_docx_xml(paragraphs)
    content_types_xml = b"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels_xml = b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>
"""
    return {
        "[Content_Types].xml": content_types_xml,
        "_rels/.rels": rels_xml,
        "word/document.xml": document_xml,
    }


def markdown_to_plain_text(md: str) -> str:
    lines = md.splitlines()
    out = []
    for line in lines:
        stripped = line.lstrip()
        if stripped.startswith("#"):
            stripped = stripped.lstrip("#").strip()
        elif stripped.startswith(("- ", "* ", "+ ")):
            stripped = stripped[2:].strip()
        else:
            parts = stripped.split(maxsplit=1)
            if parts and parts[0].rstrip(".").isdigit() and len(parts) == 2:
                stripped = parts[1].strip()
        out.append(stripped)
    return "\n".join(out).strip() + "\n"


def _compute_omega_score(converted: str, dst_fmt: str) -> float:
    chars_out = len(converted)
    length_score = min(chars_out / 10000.0, 1.0)
    structure_score = 0.0
    if dst_fmt == "md":
        lines = converted.splitlines()
        headings = sum(1 for ln in lines if ln.lstrip().startswith("#"))
        lists = sum(1 for ln in lines if ln.lstrip().startswith(("- ", "* ", "1. ")))
        structure_score = min((headings + lists) / 50.0, 1.0)
    elif dst_fmt in ("html", "docx"):
        paras = converted.count("\n\n") + 1
        structure_score = min(paras / 50.0, 1.0)
    return max(0.0, min(0.4 * length_score + 0.6 * structure_score, 1.0))


def convert_any(input_path: Path, dst_format: str, mode: str = "core", enable_log: bool = True) -> Path:
    if not input_path.is_file():
        raise SovereignDocError(f"Input file does not exist: {input_path}")
    src_ext = input_path.suffix.lower()
    if src_ext not in (".docx", ".txt", ".md"):
        raise SovereignDocError("Unsupported input format. Use .docx, .txt, or .md")
    dst_format = dst_format.lower()
    if dst_format not in ("txt", "md", "html", "docx"):
        raise SovereignDocError("Unsupported output format. Use txt, md, html, or docx")

    base = input_path.with_suffix("")
    raw_out = base.with_suffix("." + dst_format)
    out_path = _unique_path(raw_out)

    start = time.perf_counter()
    if src_ext == ".docx":
        plain = extract_docx_text(input_path)
    else:
        plain = input_path.read_text(encoding="utf-8")

    if dst_format == "txt":
        converted = plain
        out_path.write_text(converted, encoding="utf-8")
    elif dst_format == "md":
        converted = plain if src_ext == ".md" else txt_to_markdown(plain)
        if not converted.endswith("\n"):
            converted += "\n"
        out_path.write_text(converted, encoding="utf-8")
    elif dst_format == "html":
        converted = markdown_to_html(plain) if src_ext == ".md" else text_to_html(plain)
        out_path.write_text(converted, encoding="utf-8")
    else:  # docx
        flat = markdown_to_plain_text(plain) if src_ext == ".md" else plain
        files = text_to_docx_bytes(flat)
        with zipfile.ZipFile(out_path, "w") as zf:
            for arcname, data in files.items():
                zf.writestr(arcname, data)
        converted = flat

    duration_ms = (time.perf_counter() - start) * 1000.0
    chars_out = len(converted)
    omega_score = _compute_omega_score(converted, dst_format)

    if enable_log:
        _log_conversion(
            input_path=input_path,
            output_path=out_path,
            src_fmt=src_ext.lstrip("."),
            dst_fmt=dst_format,
            mode=mode,
            duration_ms=duration_ms,
            chars_out=chars_out,
            omega_score=omega_score,
        )

    return out_path
