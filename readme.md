# EUCapML Booklet → JSONL (RAG-ready)

This repo contains a deterministic parser to convert the **EU Capital Markets Law** course booklet
(DOCX) into a **JSONL** index for Retrieval-Augmented Generation.

## What it does

- **Content types:** `section`, `paragraph`, `footnote`, `case_note`
  - A `case_note` is any paragraph that **starts with** `Case Study`
- **Footnotes:** Replaced inline in paragraph/list/table text as `(fn N: {...})`
  and also emitted as separate `footnote` nodes linked to the parent paragraph.
- **Tables:** Each non-empty cell becomes a `paragraph` with `links = {table_id, row, col}`.
- **Chunking:** Paragraph-level if `≤ 1200` chars after inlining; otherwise soft-split on
  sentence boundaries with **15% overlap**.
- **Anchors:** Human-friendly breadcrumb + `¶` ordinal (and `fn` labels for footnotes).
- **Language:** Node-level heuristic `en`/`de`.
- **Citation map:** Optional extra file `{node_id → anchor}` to help your prompt builder.

## Install

```bash
python -m venv .venv
source .venv/bin/activate         # Windows: .venv\Scripts\activate
pip install -r requirements.txt
