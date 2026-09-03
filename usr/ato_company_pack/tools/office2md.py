#!/usr/bin/env python3
"""Convert Office files to Markdown so Agent Zero's knowledge index can read them.

Agent Zero indexes only these types automatically:
    .txt  .md  .pdf  .html  .csv  .json

Your Word manuals, PowerPoint decks and Excel registers are NOT in that list, so
without this step they sit in the knowledge folder unread.  This script turns them
into .md files (plain text with headings, lists and tables) that Agent Zero indexes
like any other note.  The original Office file is always kept.

Standard library only - no pip install, works on macOS, Linux and Windows.

    python3 office2md.py FILE.docx [FILE.pptx ...]     convert specific files
    python3 office2md.py --dir ../knowledge            convert everything in a folder
    python3 office2md.py --dir ../knowledge --dry-run  show what would be created

Output: <name>.md next to the original.  Existing .md files are skipped unless
--force is given, so re-running is cheap and safe.
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

OFFICE_EXT = {".docx", ".pptx", ".xlsx"}

# heading style ids inside a docx -> markdown level
HEADING_MAP = {
    "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6,
    "heading1": 1, "heading2": 2, "heading3": 3, "heading4": 4,
    "heading5": 5, "heading6": 6, "title": 1, "subtitle": 2,
}


# --------------------------------------------------------------------------- docx
def _para_text(p: ET.Element) -> str:
    out = []
    for node in p.iter():
        if node.tag == W + "t":
            out.append(node.text or "")
        elif node.tag == W + "tab":
            out.append("\t")
        elif node.tag == W + "br":
            out.append("\n")
    return "".join(out).strip()


def _para_style(p: ET.Element) -> str:
    pr = p.find(W + "pPr")
    if pr is None:
        return ""
    st = pr.find(W + "pStyle")
    if st is None:
        return ""
    return (st.get(W + "val") or st.get("val") or "").lower()


NUM_HEADING = re.compile(r"^(\d+(?:\.\d+)*)[.)]?\s+\S")
IGNORE_HEADING = re.compile(r"^(version|page|date|ref|figure|table)\b", re.I)


def _guess_heading(text: str) -> int | None:
    """Fall back to a numbering / shout heuristic when a docx has no heading styles.

    Headings matter: they are what the agent's search lands on, and a wall of
    undifferentiated paragraphs retrieves badly.
    """
    if len(text) > 90 or "\n" in text or text.endswith("."):
        return None
    m = NUM_HEADING.match(text)
    if m and not IGNORE_HEADING.match(text):
        return min(1 + m.group(1).count("."), 4)
    letters = [c for c in text if c.isalpha()]
    if len(letters) > 6 and text.upper() == text and not text.endswith(":"):
        return 3
    return None


def _is_list(p: ET.Element) -> str | None:
    pr = p.find(W + "pPr")
    if pr is None:
        return None
    num = pr.find(W + "numPr")
    if num is None:
        return None
    ilvl = num.find(W + "ilvl")
    lvl = 0
    if ilvl is not None:
        try:
            lvl = int(ilvl.get(W + "val") or ilvl.get("val") or 0)
        except ValueError:
            lvl = 0
    return "  " * lvl + "- "


def docx_to_md(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find(W + "body")
    if body is None:
        return ""
    lines: list[str] = []
    for child in body:
        tag = child.tag
        if tag == W + "p":
            text = _para_text(child)
            if not text:
                continue
            style = _para_style(child)
            lvl = HEADING_MAP.get(style) or _guess_heading(text)
            if lvl:
                lines.append(f"\n{'#' * lvl} {text}\n")
                continue
            bullet = _is_list(child)
            if bullet:
                lines.append(f"{bullet}{text}")
                continue
            lines.append(text)
            lines.append("")
        elif tag == W + "tbl":
            lines.extend(_docx_table(child))
            lines.append("")
    return "\n".join(lines).strip() + "\n"


def _docx_table(tbl: ET.Element) -> list[str]:
    rows: list[list[str]] = []
    for tr in tbl.findall(W + "tr"):
        cells: list[str] = []
        for tc in tr.findall(W + "tc"):
            parts = [_para_text(p) for p in tc.findall(W + "p")]
            cells.append(" ".join(x for x in parts if x).strip())
        if any(cells):
            rows.append(cells)
    if not rows:
        return []
    return _md_table(rows)


def _md_table(rows: list[list[str]]) -> list[str]:
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    header, *rest = rows
    out = ["", "| " + " | ".join(c.replace("|", "/") or " " for c in header) + " |",
           "|" + "|".join("---" for _ in header) + "|"]
    for r in rest:
        out.append("| " + " | ".join(c.replace("|", "/") or " " for c in r) + " |")
    out.append("")
    return out


# --------------------------------------------------------------------------- pptx
def _slide_sort(name: str) -> int:
    m = re.search(r"(\d+)", Path(name).stem)
    return int(m.group(1)) if m else 0


def _ph_type(sp: ET.Element) -> str:
    """Placeholder type of a shape, e.g. 'title', 'body', 'sldNum', 'ftr', ''."""
    for node in sp.iter():
        if node.tag.endswith("}ph"):
            return (node.get("type") or "").lower()
    return ""


def _slide_paragraphs(z: zipfile.ZipFile, name: str) -> list[list[str]]:
    """One list of paragraphs per shape, in document order."""
    root = ET.fromstring(z.read(name))
    shapes: list[list[str]] = []
    for sp in root.iter():
        if not sp.tag.endswith("}sp"):
            continue
        if _ph_type(sp) in ("sldnum", "ftr", "dt"):
            continue
        paras = []
        for p in sp.iter(A + "p"):
            text = "".join(t.text or "" for t in p.iter(A + "t")).strip()
            if text and not re.fullmatch(r"\d{1,2}", text):      # lone page numbers
                paras.append(text)
        if paras:
            shapes.append(paras)
    return shapes


def _boilerplate(slides: list[list[list[str]]]) -> set[str]:
    """Text that repeats on most slides (logos, footers, page numbers) - drop it."""
    if len(slides) < 3:
        return set()
    counts: dict[str, int] = {}
    for shapes in slides:
        for para in {p for shape in shapes for p in shape}:
            counts[para] = counts.get(para, 0) + 1
    limit = len(slides) * 0.6
    return {p for p, c in counts.items() if c > limit}


def pptx_to_md(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        slide_names = sorted((n for n in z.namelist()
                              if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)), key=_slide_sort)
        notes = sorted((n for n in z.namelist()
                        if re.fullmatch(r"ppt/notesSlides/notesSlide\d+\.xml", n)),
                       key=_slide_sort)
        parsed = [_slide_paragraphs(z, n) for n in slide_names]
        noise = _boilerplate(parsed)
        out: list[str] = []
        for i, shapes in enumerate(parsed, start=1):
            clean = [[p for p in shape if p not in noise] for shape in shapes]
            clean = [shape for shape in clean if shape]
            flat = [p for shape in clean for p in shape]
            title = ""
            if flat and len(flat[0]) <= 100 and not flat[0].endswith("."):
                title = flat.pop(0)
            out.append(f"\n## Slide {i}: {title or '(untitled)'}\n")
            for b in flat:
                out.append(f"- {b}")
            out.append("")
        for i, name in enumerate(notes, start=1):
            root = ET.fromstring(z.read(name))
            texts = []
            for p in root.iter(A + "p"):
                txt = "".join(t.text or "" for t in p.iter(A + "t")).strip()
                if txt:
                    texts.append(txt)
            if texts:
                out.append(f"### Slide {i} speaker notes\n")
                out.extend(texts)
                out.append("")
    return "\n".join(out).strip() + "\n"


# --------------------------------------------------------------------------- xlsx
def xlsx_to_md(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(S + "si"):
                shared.append("".join(t.text or "" for t in si.iter(S + "t")))
        names = _sheet_names(z)
        sheets = sorted((n for n in z.namelist()
                         if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)), key=_slide_sort)
        out: list[str] = []
        for idx, sname in enumerate(sheets):
            root = ET.fromstring(z.read(sname))
            label = names[idx] if idx < len(names) else Path(sname).stem
            out.append(f"\n## {label}\n")
            rows: list[list[str]] = []
            for row in root.iter(S + "row"):
                cells: list[str] = []
                for c in row.findall(S + "c"):
                    cells.append(_cell_text(c, shared))
                while cells and cells[-1] == "":
                    cells.pop()
                if cells:
                    rows.append(cells)
            rows = rows[:2000]          # guard against runaway sheets
            if not rows:
                out.append("_empty sheet_")
                continue
            out.extend(_md_table(rows))
    return "\n".join(out).strip() + "\n"


def _sheet_names(z: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(z.read("xl/workbook.xml"))
        return [sh.get("name") or f"Sheet{i + 1}"
                for i, sh in enumerate(root.iter(S + "sheet"))]
    except KeyError:
        return []


def _cell_text(c: ET.Element, shared: list[str]) -> str:
    t = c.get("t")
    if t == "inlineStr":
        is_el = c.find(S + "is")
        return "".join(x.text or "" for x in is_el.iter(S + "t")) if is_el is not None else ""
    v = c.find(S + "v")
    if v is None or v.text is None:
        return ""
    if t == "s":
        try:
            return shared[int(v.text)]
        except (ValueError, IndexError):
            return ""
    return v.text


# --------------------------------------------------------------------------- driver
def convert_file(path: Path, force: bool = False) -> Path | None:
    """Write <path>.md next to the Office file. Returns the new path or None."""
    ext = path.suffix.lower()
    if ext not in OFFICE_EXT:
        return None
    dest = path.with_suffix(".md")
    if dest.exists() and not force:
        return None
    try:
        if ext == ".docx":
            text = docx_to_md(path)
        elif ext == ".pptx":
            text = pptx_to_md(path)
        else:
            text = xlsx_to_md(path)
    except Exception as exc:                                # never break the run
        print(f"  ! {path.name}: could not convert ({exc})", file=sys.stderr)
        return None
    if not text.strip():
        return None
    header = (f"<!-- auto-converted from {path.name} by office2md.py - "
              f"edit this file, not a re-conversion -->\n\n")
    dest.write_text(header + text, encoding="utf-8")
    return dest


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*", help="Office files to convert")
    ap.add_argument("--dir", help="convert every Office file under this folder")
    ap.add_argument("--force", action="store_true", help="overwrite existing .md output")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    targets: list[Path] = [Path(f) for f in args.files]
    if args.dir:
        base = Path(args.dir)
        targets += [p for p in sorted(base.rglob("*"))
                    if p.is_file() and p.suffix.lower() in OFFICE_EXT]
    if not targets:
        ap.print_help()
        return 1

    made = skipped = 0
    for t in targets:
        if not t.exists():
            print(f"  ! missing: {t}", file=sys.stderr)
            continue
        dest = t.with_suffix(".md")
        if dest.exists() and not args.force:
            skipped += 1
            continue
        print(f"  {t.name}  →  {dest.name}")
        if args.dry_run:
            continue
        out = convert_file(t, force=args.force)
        if out:
            made += 1
    print(f"\nConverted {made} file(s)" + (f", {skipped} already had a .md" if skipped else ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
