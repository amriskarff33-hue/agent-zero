#!/usr/bin/env python3
"""AM Risk company pack manager.

Keeps the company knowledge base tidy and gets it into the place Agent Zero indexes
(usr/knowledge/main/ato_company/).

    python3 pack_manage.py status              what is in the pack and the inbox
    python3 pack_manage.py ingest              show the filing plan for new inbox files
    python3 pack_manage.py ingest --apply      file them (copy into knowledge/, keep originals)
    python3 pack_manage.py ingest --apply --move   file them and clear the inbox
    python3 pack_manage.py convert --apply     Word/PowerPoint/Excel -> .md so they can be indexed
    python3 pack_manage.py index               rebuild knowledge/00_index.md
    python3 pack_manage.py sync                copy knowledge/ -> usr/knowledge/main/ato_company/
    python3 pack_manage.py checklist           ATO knowledge checklist with what is present

Safe by default: ingest and sync only ever add or overwrite files that came from this pack.
They never delete anything in the sync target, and ingest keeps the originals unless --move.
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
PACK_DIR = TOOLS_DIR.parent                      # .../usr/ato_company_pack
ROOT = PACK_DIR.parents[1]                       # repository root  (/a0 on the Mac)
KNOWLEDGE = PACK_DIR / "knowledge"
INBOX = PACK_DIR / "inbox"
DEFAULT_SYNC_TARGET = ROOT / "usr" / "knowledge" / "main" / "ato_company"

sys.path.insert(0, str(TOOLS_DIR))
try:
    from office2md import OFFICE_EXT, convert_file as office_to_md
except ImportError:                                   # office2md.py sits beside this file
    OFFICE_EXT, office_to_md = set(), None

SKIP = {".gitkeep", ".ds_store", "thumbs.db",
        "readme.md", "naming.md", "knowledge_checklist.md", "00_index.md"}

# (folder, keywords found in the file name, description)
RULES = [
    ("company", ("company", "profile", "about us", "organogram", "org chart", "brand",
                 "service catalogue", "services", "ato certificate", "approval", "team",
                 "staff", "contact"),
     "company profile, brand, service catalogue, approvals"),
    ("airports", ("aip", "ad 2", "ad2", "aerodrome", "chart", "notam", "airport data"),
     "aerodrome data packs, AIP pages, aerodrome charts"),
    ("regulatory", ("car ", "car-", "sacaa", "sacat", "icao", "annex", "nfpa", "aci", "iata",
                    "gazette", "regulation", "standard", "sa-cats", "ohs", "act"),
     "standards and regulations"),
    ("alliance", ("alliance", "training centre alliance", "training center alliance", "partner",
                  "consortium", "joint venture", "jv", "network", "member", "accredited centre"),
     "Training Centre Alliance material shared with partner centres"),
    ("clients", ("client", "customer", "contract", "appointment", "quote", "proposal", "sla"),
     "clients, contracts, appointments and proposals"),
    ("courses", ("course", "lesson", "manual", "workbook", "handout", "exam", "assessment",
                 "avop", "arff", "syllabus", "curriculum", "module", "slide", "firefighter",
                 "marshalling", "emergency"),
     "past courses, lesson plans, manuals, exams"),
    ("operations", ("sop", "procedure", "policy", "checklist", "form", "template", "register",
                    "report", "plan", "brief"),
     "SOPs, policies, forms, checklists"),
    ("media", ("photo", "image", "logo", "figure", "diagram", "brand"),
     "photos, logos, figures"),
]

FOLDERS = list(dict.fromkeys([r[0] for r in RULES] + ["company"]))

TEXT_EXT = {".md", ".txt"}
DOC_EXT = {".pdf", ".docx", ".doc", ".pptx", ".ppt", ".xlsx", ".xls", ".html", ".htm", ".csv"}
IMG_EXT = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".tif", ".tiff"}

# What Agent Zero's knowledge index reads on its own (plugins/_memory/helpers/knowledge_import.py).
# Everything else needs converting first - pack_manage.py convert does that.
INDEXABLE = {".md", ".txt", ".pdf", ".html", ".htm", ".csv", ".json"}


# --------------------------------------------------------------------------- utils
def slugify(name: str) -> str:
    stem, ext = Path(name).stem, Path(name).suffix.lower()
    stem = stem.strip().lower()
    stem = re.sub(r"[^\w\s.-]", "", stem)
    stem = re.sub(r"[\s_]+", "-", stem)
    stem = re.sub(r"-{2,}", "-", stem).strip("-")
    return f"{stem}{ext}" if stem else f"untitled{ext}"


def classify(path: Path) -> str:
    name = path.name.lower()
    if path.suffix.lower() in IMG_EXT and not any(k in name for k in ("logo", "diagram", "figure")):
        return "media"
    for folder, keywords, _ in RULES:
        if any(k in name for k in keywords):
            return folder
    if path.suffix.lower() in IMG_EXT:
        return "media"
    return "unsorted"


def human(size: int) -> str:
    return f"{size/1024/1024:.1f} MB" if size > 1024 * 1024 else f"{max(1, size//1024)} kB"


def iter_files(folder: Path):
    if not folder.exists():
        return []
    return sorted(p for p in folder.rglob("*")
                  if p.is_file() and p.name.lower() not in SKIP and not p.name.startswith("."))


def ensure_dirs():
    for f in FOLDERS + ["unsorted"]:
        (KNOWLEDGE / f).mkdir(parents=True, exist_ok=True)
    INBOX.mkdir(parents=True, exist_ok=True)


# --------------------------------------------------------------------------- actions
def convert(args) -> int:
    """Turn Word / PowerPoint / Excel files into .md so Agent Zero can index them."""
    ensure_dirs()
    if office_to_md is None:
        print("office2md.py is missing next to pack_manage.py - cannot convert.")
        return 1
    sources = iter_files(KNOWLEDGE) + iter_files(INBOX)
    todo = [f for f in sources if f.suffix.lower() in OFFICE_EXT
            and (args.force or not f.with_suffix(".md").exists())]
    if not todo:
        print("Nothing to convert - every Office file already has a .md beside it.\n"
              "Use --force to regenerate them.")
        return 0
    print("CONVERTING:" if args.apply else "CONVERTING (dry run - add --apply):")
    print()
    done = 0
    for f in todo:
        print(f"  {f.relative_to(PACK_DIR)}  ->  {f.with_suffix('.md').name}")
        if args.apply:
            if office_to_md(f, force=args.force):
                done += 1
    if args.apply:
        print(f"\nConverted {done} file(s). Originals untouched.\n"
              "Next: python3 pack_manage.py index   then   python3 pack_manage.py sync")
    else:
        print("\nNothing was changed. Re-run with --apply.")
    return 0


def status(_args) -> int:
    ensure_dirs()
    print(f"Company pack : {PACK_DIR}")
    print(f"Sync target  : {DEFAULT_SYNC_TARGET} "
          f"({'exists' if DEFAULT_SYNC_TARGET.exists() else 'not synced yet'})")
    print("\nKnowledge:")
    for folder in FOLDERS + ["unsorted"]:
        files = iter_files(KNOWLEDGE / folder)
        total = sum(f.stat().st_size for f in files)
        print(f"  {folder:<12} {len(files):>3} file(s)  {human(total):>9}")
    unreadable = [f for f in iter_files(KNOWLEDGE)
                  if f.suffix.lower() not in INDEXABLE
                  and not f.with_suffix(".md").exists()]
    if unreadable:
        print(f"\nNot readable by Agent Zero yet ({len(unreadable)} file(s)) - run "
              f"'pack_manage.py convert --apply':")
        for f in unreadable[:10]:
            print(f"  ! {f.relative_to(KNOWLEDGE)}")
    inbox = iter_files(INBOX)
    print(f"\nInbox: {len(inbox)} file(s) waiting to be filed")
    for f in inbox[:20]:
        print(f"  → {f.name}   ({classify(f)})")
    if len(inbox) > 20:
        print(f"  … {len(inbox) - 20} more")
    return 0


def ingest(args) -> int:
    ensure_dirs()
    inbox = iter_files(INBOX)
    if not inbox:
        print("Inbox is empty. Drop files into:")
        print(f"  {INBOX}")
        return 0
    print(f"FILING PLAN ({'APPLYING' if args.apply else 'dry run — add --apply to file them'}):\n")
    moved = 0
    for src in inbox:
        folder = classify(src)
        target_dir = KNOWLEDGE / folder
        target_dir.mkdir(parents=True, exist_ok=True)
        dest = target_dir / slugify(src.name)
        n = 1
        while dest.exists():
            dest = target_dir / f"{Path(slugify(src.name)).stem}-{n}{src.suffix.lower()}"
            n += 1
        action = "COPY" if not args.move else "MOVE"
        print(f"  {folder:<11} {src.name}  →  {dest.name}   [{action}]")
        if args.apply:
            if args.move:
                shutil.move(str(src), str(dest))
            else:
                shutil.copy2(src, dest)
            moved += 1
            if src.suffix.lower() in OFFICE_EXT and office_to_md:
                if office_to_md(dest, force=False):
                    print(f"               + {dest.with_suffix('.md').name} "
                          f"(converted so it can be indexed)")
    print()
    if args.apply:
        print(f"Filed {moved} file(s). Originals {'moved' if args.move else 'kept in the inbox'}.")
        print("Next: python3 pack_manage.py index   then   python3 pack_manage.py sync")
    else:
        print("Nothing was changed. Re-run with --apply to file them.")
    return 0


def index(_args) -> int:
    ensure_dirs()
    lines = ["# Company knowledge index", "",
             f"Generated: {datetime.now():%Y-%m-%d %H:%M}", "",
             "This file is regenerated by `tools/pack_manage.py index`.",
             "It tells the agent what exists so it can ask for the right document.", "",
             "Folder README files are instructions for people, not knowledge - they are not indexed.",
             ""]
    purpose = {f: d for f, _, d in RULES}
    purpose.setdefault("unsorted", "rename or move these into the right folder")
    for folder in FOLDERS:
        files = iter_files(KNOWLEDGE / folder)
        lines.append(f"## {folder.capitalize()}")
        if folder in purpose:
            lines.append("")
            lines.append(f"_{purpose[folder]}._")
        if not files:
            lines.append("")
            lines.append("_Nothing here yet._")
            lines.append("")
            continue
        lines.append("")
        for f in files:
            rel = f.relative_to(KNOWLEDGE)
            lines.append(f"- `{rel}` — {human(f.stat().st_size)}")
        lines.append("")
    unsorted = iter_files(KNOWLEDGE / "unsorted")
    if unsorted:
        lines.append("## Needs sorting")
        lines.append("")
        for f in unsorted:
            lines.append(f"- `{f.relative_to(KNOWLEDGE)}` — rename or move it to the right folder")
        lines.append("")
    out = KNOWLEDGE / "00_index.md"
    out.write_text("\n".join(lines), encoding="utf-8")
    total = len(iter_files(KNOWLEDGE))
    print(f"Index written: {out}  ({total} files indexed)")
    return 0


def sync(args) -> int:
    ensure_dirs()
    target = Path(args.target) if args.target else DEFAULT_SYNC_TARGET
    target.mkdir(parents=True, exist_ok=True)
    copied = updated = skipped = 0
    for src in iter_files(KNOWLEDGE):
        dest = target / src.relative_to(KNOWLEDGE)
        dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.exists() and dest.stat().st_size == src.stat().st_size:
            skipped += 1
            continue
        shutil.copy2(src, dest)
        if dest.exists():
            updated += 1
        else:
            copied += 1
    print(f"Synced into {target}")
    print(f"  new: {copied}  updated: {updated}  unchanged: {skipped}")
    print("Agent Zero indexes usr/knowledge/ — ask him what he knows and he will quote these.")
    return 0


def checklist(_args) -> int:
    items = [
        ("company", "Company profile: legal name, ATO approval, services, key people"),
        ("company", "Brand and document style: logo, colours, templates, writing voice"),
        ("company", "Service catalogue with course codes and durations"),
        ("clients", "Client and airport list with contacts and authorities"),
        ("clients", "Contracts, SLAs and appointment letters"),
        ("courses", "One complete past course per type (the best exemplar of each)"),
        ("courses", "Lesson plans and facilitator notes"),
        ("courses", "Exam papers and memoranda"),
        ("courses", "Learner manuals and workbooks"),
        ("operations", "Training delivery SOP (booking → delivery → records)"),
        ("operations", "Assessment, moderation and re-sit policy"),
        ("operations", "Record-keeping and POPIA procedure"),
        ("operations", "Quotation and pricing templates"),
        ("regulatory", "SACAA CARs / SA-CATS relevant parts"),
        ("regulatory", "ICAO Annexes and Docs in scope"),
        ("regulatory", "NFPA standards in scope"),
        ("regulatory", "ACI / IATA material in scope"),
        ("airports", "An aerodrome data pack for every airport you train at"),
        ("media", "Logo and brand assets"),
        ("media", "Photo library of equipment, vehicles, aprons, incidents"),
        ("alliance", "Training Centre Alliance: shared courses, standards and brand rules"),
    ]
    print("ATO KNOWLEDGE CHECKLIST\n")
    present = 0
    for folder, item in items:
        has = bool(iter_files(KNOWLEDGE / folder))
        present += has
        print(f"  [{'x' if has else ' '}] {item}   ({folder})")
    print(f"\n{present}/{len(items)} areas have content. "
          "Drop files in the inbox and run: python3 pack_manage.py ingest --apply")
    return 0


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="action", required=True)
    sub.add_parser("status", help="what is in the pack and in the inbox").set_defaults(func=status)
    ig = sub.add_parser("ingest", help="file inbox documents into knowledge/")
    ig.add_argument("--apply", action="store_true", help="actually file them")
    ig.add_argument("--move", action="store_true", help="move instead of copy")
    ig.set_defaults(func=ingest)
    sub.add_parser("index", help="rebuild knowledge/00_index.md").set_defaults(func=index)
    sy = sub.add_parser("sync", help="copy knowledge/ into usr/knowledge/main/ato_company/")
    sy.add_argument("--target", help="override the destination")
    sy.set_defaults(func=sync)
    cv = sub.add_parser("convert", help="turn Word/PowerPoint/Excel files into .md")
    cv.add_argument("--apply", action="store_true", help="actually write the .md files")
    cv.add_argument("--force", action="store_true", help="regenerate .md files that already exist")
    cv.set_defaults(func=convert)
    sub.add_parser("checklist", help="ATO knowledge checklist").set_defaults(func=checklist)
    args = p.parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
