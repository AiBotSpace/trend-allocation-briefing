#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Extract plain text (+ Telegram chunks) from an edition's .docx.

Usage:
    python3 generator/extract_text.py <NN> [doc-url]

Reads editions/briefing-NN.docx, writes editions/briefing-NN.txt (full text,
headings / numbered lists / tables preserved as plain text) and
editions/briefing-NN.chunkN.txt (<=4000 chars each, split at line boundaries,
for Telegram's 4096-char message limit). If doc-url is given it is appended
as the final line.
"""
import sys
from pathlib import Path
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

REPO = Path(__file__).resolve().parent.parent


def extract(docx_path):
    doc = Document(str(docx_path))
    lines = []
    num = 0
    for block in doc.element.body.iterchildren():
        if block.tag == qn("w:p"):
            p = Paragraph(block, doc)
            text = p.text.strip()
            style = p.style.name if p.style else ""
            if not text:
                continue
            if style.startswith("Heading"):
                num = 0
                lines.append("")
                lines.append(text)
            elif style == "List Number":
                num += 1
                lines.append(f"{num}. {text}")
            elif style == "List Bullet":
                lines.append(f"• {text}")
            else:
                lines.append(text)
        elif block.tag == qn("w:tbl"):
            t = Table(block, doc)
            for row in t.rows:
                cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                lines.append(" ｜ ".join(cells))
            lines.append("")
    return lines


def chunk(lines, limit=4000):
    chunks, cur = [], ""
    for ln in lines:
        if len(cur) + len(ln) + 1 > limit:
            chunks.append(cur)
            cur = ln
        else:
            cur = cur + "\n" + ln if cur else ln
    if cur:
        chunks.append(cur)
    return chunks


def main():
    nn = sys.argv[1]
    doc_url = sys.argv[2] if len(sys.argv) > 2 else None
    editions = REPO / "editions"
    lines = extract(editions / f"briefing-{nn}.docx")
    if doc_url:
        lines += ["", f"全文链接：{doc_url}"]
    full = "\n".join(lines)
    (editions / f"briefing-{nn}.txt").write_text(full, encoding="utf-8")
    chunks = chunk(lines)
    for i, c in enumerate(chunks, 1):
        (editions / f"briefing-{nn}.chunk{i}.txt").write_text(c, encoding="utf-8")
    print(f"briefing-{nn}.txt: {len(full)} chars, {len(chunks)} telegram chunk(s)")


if __name__ == "__main__":
    main()
