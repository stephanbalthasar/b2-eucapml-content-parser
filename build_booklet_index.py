#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EUCapML booklet parser → JSONL (RAG-ready)

Implements the agreed spec:
- Types: section, paragraph, footnote, case_note (paragraphs starting with "Case Study")
- Inline footnotes in paragraphs/lists/table cells as: (fn N: {footnote_text})
- Also emit each footnote as its own node with link to the parent paragraph
- Tables: each text cell becomes a paragraph node with table_id/row/col in `links`
- Chunking: <= 1200 chars after inlining; else soft-split with 15% overlap
- Anchors: human-friendly breadcrumb + ¶ordinal (+ fn labels for footnote nodes)
- Lang: simple heuristic ('en' or 'de')
- Optional: write citation map {node_id -> anchor}

Usage:
    python build_booklet_index.py \
        --input /path/to/booklet.docx \
        --output ./booklet_index.jsonl \
        --emit-citation-map ./citation_map.json
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

# 3rd-party
#   - python-docx parses the main document structure
#   - lxml is used to parse footnotes.xml from the DOCX ZIP for reliable footnotes
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table

try:
    import lxml.etree as ET
except ImportError:
    ET = None  # Will be validated in main()


# ---------- Config (defaults) ----------
DEFAULT_MAX_LEN = 1200
DEFAULT_OVERLAP_RATIO = 0.15
DOC_ID = "EUCapML-2026-v11-2025-06-11"  # stable ID for this booklet version


# ---------- Helpers: headings ----------
def is_heading(paragraph: Paragraph) -> bool:
    try:
        name = paragraph.style.name or ""
    except Exception:
        name = ""
    return name.startswith("Heading")


def heading_level(paragraph: Paragraph) -> Optional[int]:
    try:
        name = paragraph.style.name or ""
    except Exception:
        name = ""
    if name.startswith("Heading"):
        parts = name.split()
        if len(parts) >= 2 and parts[1].isdigit():
            return int(parts[1])
    return None


# ---------- Footnotes: load footnotes.xml from DOCX ----------
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def load_docx_and_footnotes(docx_path: str) -> Tuple[Document, Dict[int, str]]:
    """
    Return python-docx Document + {footnote_id -> text} from word/footnotes.xml
    """
    if ET is None:
        raise RuntimeError("lxml is required. Please `pip install lxml`.")

    doc = Document(docx_path)
    footnotes_map: Dict[int, str] = {}

    with zipfile.ZipFile(docx_path, "r") as z:
        try:
            xml_bytes = z.read("word/footnotes.xml")
        except KeyError:
            xml_bytes = None

    if not xml_bytes:
        return doc, footnotes_map  # no footnotes present

    root = ET.fromstring(xml_bytes)
    for fn_el in root.findall(".//w:footnote", W_NS):
        fn_id = fn_el.get("{%s}id" % W_NS["w"])
        if fn_id is None:
            continue
        # Skip separators and negative IDs
        fn_type = fn_el.get("{%s}type" % W_NS["w"])
        if fn_type in ("separator", "continuationSeparator"):
            continue
        try:
            fid = int(fn_id)
        except ValueError:
            continue
        if fid < 0:
            continue

        texts = [t.text for t in fn_el.findall(".//w:t", W_NS) if t.text]
        text = " ".join(texts)
        text = re.sub(r"\s+", " ", text).strip()
        footnotes_map[fid] = text

    return doc, footnotes_map


# ---------- Low-level XML access to paragraph runs (for footnoteReference) ----------
def paragraph_runs_xml(paragraph: Paragraph):
    """Return the raw XML children under w:p to find w:footnoteReference reliably."""
    return list(paragraph._p.iterchildren())


def run_contains_visible_text(run_el) -> bool:
    return any(
        child.tag.endswith("}t") and child.text
        for child in run_el.iter()
    )


def run_footnote_refs(run_el) -> List[int]:
    refs: List[int] = []
    for ref in run_el.findall(".//w:footnoteReference", W_NS):
        fid = ref.get("{%s}id" % W_NS["w"])
        if fid is not None:
            try:
                refs.append(int(fid))
            except ValueError:
                pass
    return refs


# ---------- Inline footnote replacement ----------
def extract_text_with_inlined_fns(
    paragraph: Paragraph,
    footnotes_map: Dict[int, str],
    inline_in_headings: bool = False,
) -> str:
    """
    Build paragraph text and inline footnotes as "(fn N: ...)".
    Multiple adjacent footnote runs with no visible text in between are merged:
        (fn 12: ...; fn 13: ...)
    """
    if is_heading(paragraph) and not inline_in_headings:
        return paragraph.text

    pieces: List[str] = []
    runs = paragraph_runs_xml(paragraph)
    i = 0
    while i < len(runs):
        r = runs[i]
        tag = r.tag
        # Skip paragraph properties
        if tag.endswith("}pPr"):
            i += 1
            continue

        refs = run_footnote_refs(r)
        if refs:
            # Merge contiguous footnote-only runs
            merged_ids = list(refs)
            j = i + 1
            while j < len(runs):
                r2 = runs[j]
                if run_contains_visible_text(r2):
                    break
                refs2 = run_footnote_refs(r2)
                if refs2:
                    merged_ids.extend(refs2)
                    j += 1
                else:
                    break

            # Build merged bracket
            parts = []
            for fid in merged_ids:
                fn_txt = footnotes_map.get(fid, "")
                fn_txt = re.sub(r"\s+", " ", fn_txt).strip()
                parts.append(f"fn {fid}: {fn_txt}")
            pieces.append("(" + "; ".join(parts) + ")")
            i = j
            continue

        # Otherwise, collect visible text
        texts = [t.text for t in r.findall(".//w:t", W_NS) if t.text]
        if texts:
            pieces.append("".join(texts))
        i += 1

    out = "".join(pieces)
    out = re.sub(r"\s+", " ", out).strip()
    return out


# ---------- Language heuristic ----------
DE_HINTS = set(
    """
    der die das und nicht mit auf aus bei durch gegen ohne unter vom zur zum gemäß
    auch sowie daher soweit darüber hierzu hiervon hierfür ist sind war waren einer
    einem einen eines denn doch oder aber noch schon sehr mehr weniger danach
    """.split()
)


def detect_lang(text: str) -> str:
    t = text.lower()
    # quick signal
    if any(ch in t for ch in "äöüß"):
        return "de"
    words = re.findall(r"[a-zA-ZäöüÄÖÜß]{2,}", t)
    if not words:
        return "en"
    de_count = sum(1 for w in words if w in DE_HINTS)
    ratio = de_count / max(1, len(words))
    return "de" if ratio > 0.08 else "en"


# ---------- Chunking (soft split with 15% overlap) ----------
SENT_SPLIT_RE = re.compile(r"(?<=[.!?])\s+(?=[A-ZÄÖÜ])")


def soft_split(text: str, max_len: int, overlap_ratio: float) -> List[str]:
    if len(text) <= max_len:
        return [text]
    sentences = SENT_SPLIT_RE.split(text)
    chunks: List[str] = []
    buf = ""
    for sent in sentences:
        if not buf:
            buf = sent
        elif len(buf) + 1 + len(sent) <= max_len:
            buf = buf + " " + sent
        else:
            chunks.append(buf)
            buf = sent
    if buf:
        chunks.append(buf)

    if len(chunks) <= 1:
        return chunks

    # Apply simple char overlap by prepending tail from previous chunk
    overlap_chars = int(max_len * overlap_ratio)
    overlapped: List[str] = []
    for idx, ch in enumerate(chunks):
        if idx == 0:
            overlapped.append(ch)
        else:
            prev = overlapped[-1]
            tail = prev[-overlap_chars:]
            merged = (tail + " " + ch).strip()
            overlapped.append(merged)
    return overlapped


# ---------- Anchors ----------
def format_anchor(
    breadcrumb_list: List[str],
    para_ord: Optional[int] = None,
    fn_ids: Optional[List[int]] = None,
) -> str:
    title = " » ".join([b for b in breadcrumb_list if b])
    anchor = title
    if para_ord is not None:
        anchor += f" » ¶{para_ord}"
    if fn_ids:
        inside = "; ".join([f"fn {fid}" for fid in fn_ids])
        anchor += f" ({inside})"
    return anchor


# ---------- Main parse ----------
def build_index(
    input_path: str,
    output_path: str,
    emit_citation_map: Optional[str] = None,
    max_len: int = DEFAULT_MAX_LEN,
    overlap_ratio: float = DEFAULT_OVERLAP_RATIO,
) -> int:
    """
    Parse the DOCX into JSONL. Returns count of nodes written.
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input DOCX not found: {input_path}")

    doc, footnotes_map = load_docx_and_footnotes(input_path)
    body = doc.element.body

    # Breadcrumb & counters
    breadcrumb: List[str] = []
    para_counters: Dict[Tuple[str, ...], int] = defaultdict(int)
    case_counters: Dict[Tuple[str, ...], int] = defaultdict(int)
    section_counters: Dict[Tuple[str, ...], int] = defaultdict(int)

    table_counter = 0
    nodes: List[dict] = []
    citation_map: Dict[str, str] = {}

    for child in body.iterchildren():
        tag = child.tag

        # Paragraphs ----------------------------------------------------------
        if tag.endswith("}p"):
            p = Paragraph(child, doc)

            # Heading → section node
            if is_heading(p):
                lvl = heading_level(p) or 1

                # Ensure breadcrumb length == level
                while len(breadcrumb) < lvl:
                    breadcrumb.append("")
                breadcrumb = breadcrumb[:lvl]

                title = (p.text or "").strip()
                breadcrumb[lvl - 1] = title

                section_counters[tuple(breadcrumb)] += 1
                node = {
                    "doc_id": DOC_ID,
                    "node_id": f"sec:{'.'.join(str(i+1) for i in range(lvl))}.{section_counters[tuple(breadcrumb)]}",
                    "type": "section",
                    "breadcrumb": breadcrumb.copy(),
                    "anchor": " » ".join([b for b in breadcrumb if b]),
                    "text": title,
                    "links": {},
                    "lang": detect_lang(title) if title else "en",
                }
                nodes.append(node)
                citation_map[node["node_id"]] = node["anchor"]
                continue

            # Normal paragraph/case_note
            text_inlined = extract_text_with_inlined_fns(
                p, footnotes_map, inline_in_headings=False
            ).strip()
            if not text_inlined:
                continue

            stripped = text_inlined.lstrip()
            is_case = stripped.lower().startswith("case study")

            key = tuple(breadcrumb)
            if is_case:
                case_counters[key] += 1
                base_id = (
                    f"{'-'.join(str(i+1) for i in range(len(breadcrumb)))}.cn{case_counters[key]}"
                    if breadcrumb else f"cn{case_counters[key]}"
                )

                chunks = soft_split(text_inlined, max_len=max_len, overlap_ratio=overlap_ratio)
                for ci, ch in enumerate(chunks, 1):
                    node_id = f"{base_id}" + (f".{ci}" if len(chunks) > 1 else "")
                    node = {
                        "doc_id": DOC_ID,
                        "node_id": node_id,
                        "type": "case_note",
                        "breadcrumb": breadcrumb.copy(),
                        "anchor": format_anchor(breadcrumb, None, None),
                        "text": ch,
                        "links": {},
                        "lang": detect_lang(ch),
                    }
                    nodes.append(node)
                    citation_map[node_id] = node["anchor"]

            else:
                # Paragraph node(s)
                para_counters[key] += 1
                para_ord = para_counters[key]
                base_id = (
                    f"{'-'.join(str(i+1) for i in range(len(breadcrumb)))}.p{para_ord}"
                    if breadcrumb else f"p{para_ord}"
                )

                chunks = soft_split(text_inlined, max_len=max_len, overlap_ratio=overlap_ratio)
                for ci, ch in enumerate(chunks, 1):
                    node_id = f"{base_id}" + (f".{ci}" if len(chunks) > 1 else "")
                    node = {
                        "doc_id": DOC_ID,
                        "node_id": node_id,
                        "type": "paragraph",
                        "breadcrumb": breadcrumb.copy(),
                        "anchor": format_anchor(breadcrumb, para_ord, None),
                        "text": ch,
                        "links": {},
                        "lang": detect_lang(ch),
                    }
                    nodes.append(node)
                    citation_map[node_id] = node["anchor"]

                # Footnote nodes referenced by THIS paragraph
                fids: List[int] = []
                for r in paragraph_runs_xml(p):
                    fids.extend(run_footnote_refs(r))
                if fids:
                    seen = set()
                    ordered = []
                    for fid in fids:
                        if fid not in seen:
                            seen.add(fid)
                            ordered.append(fid)
                    for fid in ordered:
                        fn_text = footnotes_map.get(fid, "")
                        if not fn_text:
                            continue
                        fn_node_id = f"{base_id}.fn{fid}"
                        node = {
                            "doc_id": DOC_ID,
                            "node_id": fn_node_id,
                            "type": "footnote",
                            "breadcrumb": breadcrumb.copy(),
                            "anchor": format_anchor(breadcrumb, para_ord, [fid]),
                            "text": fn_text,
                            "links": {"footnote_of": base_id},
                            "lang": detect_lang(fn_text),
                        }
                        nodes.append(node)
                        citation_map[fn_node_id] = node["anchor"]

        # Tables ---------------------------------------------------------------
        elif tag.endswith("}tbl"):
            table_counter += 1
            tbl = Table(child, doc)

            for r_i, row in enumerate(tbl.rows):
                for c_i, cell in enumerate(row.cells):
                    # Concatenate non-empty paragraph texts from the cell
                    cell_texts: List[str] = []
                    for p in cell.paragraphs:
                        if is_heading(p):
                            continue
                        t = extract_text_with_inlined_fns(
                            p, footnotes_map, inline_in_headings=False
                        ).strip()
                        if t:
                            cell_texts.append(t)
                    if not cell_texts:
                        continue

                    cell_text = " ".join(cell_texts)
                    key = tuple(breadcrumb)
                    para_counters[key] += 1
                    para_ord = para_counters[key]
                    base_id = (
                        f"{'-'.join(str(i+1) for i in range(len(breadcrumb)))}.p{para_ord}"
                        if breadcrumb else f"p{para_ord}"
                    )

                    chunks = soft_split(cell_text, max_len=max_len, overlap_ratio=overlap_ratio)
                    for ci, ch in enumerate(chunks, 1):
                        node_id = f"{base_id}" + (f".{ci}" if len(chunks) > 1 else "")
                        node = {
                            "doc_id": DOC_ID,
                            "node_id": node_id,
                            "type": "paragraph",
                            "breadcrumb": breadcrumb.copy(),
                            "anchor": format_anchor(breadcrumb, para_ord, None),
                            "text": ch,
                            "links": {"table_id": f"table{table_counter}", "row": r_i, "col": c_i},
                            "lang": detect_lang(ch),
                        }
                        nodes.append(node)
                        citation_map[node_id] = node["anchor"]

        # Ignore other body children (e.g., sectPr)
        else:
            continue

    # Write JSONL
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        for n in nodes:
            f.write(json.dumps(n, ensure_ascii=False) + "\n")

    # Optional citation map
    if emit_citation_map:
        with open(emit_citation_map, "w", encoding="utf-8") as f:
            json.dump(citation_map, f, ensure_ascii=False, indent=2)

    return len(nodes)


# ---------- CLI ----------
def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Build JSONL index for the EUCapML booklet")
    p.add_argument("--input", required=True, help="Path to booklet.docx")
    p.add_argument("--output", required=True, help="Path to write JSONL (booklet_index.jsonl)")
    p.add_argument("--emit-citation-map", default=None, help="Optional path to write citation_map.json")
    p.add_argument("--max-len", type=int, default=DEFAULT_MAX_LEN, help="Max chars per chunk (default: 1200)")
    p.add_argument("--overlap", type=float, default=DEFAULT_OVERLAP_RATIO, help="Overlap ratio (default: 0.15)")
    return p.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> None:
    ns = parse_args(argv)

    # Validate inputs
    if not os.path.exists(ns.input):
        print(f"[ERROR] Input file not found: {ns.input}", file=sys.stderr)
        sys.exit(2)

    if ET is None:
        print("[ERROR] lxml is required. Run: pip install -r requirements.txt", file=sys.stderr)
        sys.exit(3)

    count = build_index(
        input_path=ns.input,
        output_path=ns.output,
        emit_citation_map=ns.emit_citation_map,
        max_len=ns.max_len,
        overlap_ratio=ns.overlap,
    )
    print(f"[OK] Wrote {count} nodes → {ns.output}")
    if ns.emit_citation_map:
        print(f"[OK] Wrote citation map → {ns.emit_citation_map}")


if __name__ == "__main__":
    main()
