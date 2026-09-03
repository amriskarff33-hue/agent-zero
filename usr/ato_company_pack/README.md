# AM Risk company pack — his memory of how you work

This folder is **your Agent Zero's company brain**. It is plain files: documents you drop in,
filed automatically, indexed so he can recall them while working.

It is deliberately outside the plugin and outside core, so:

- **Updates never wipe it** (it lives in `usr/`, which every Agent Zero update preserves).
- **It is portable** — copy this one folder to any machine and it is the same agent.
- **You own it** — no code required to grow it.

## Install

```bash
# inside your Agent Zero folder (in Docker this is /a0)
cp -R ato_company_pack /a0/usr/
```

Then, after adding documents:

```bash
cd /a0/usr/ato_company_pack
python3 tools/pack_manage.py ingest --apply   # file whatever is in the inbox
                                              # (Word/PowerPoint files are converted to .md)
python3 tools/pack_manage.py index            # rebuild the knowledge index
python3 tools/pack_manage.py sync             # copy into usr/knowledge/main/ato_company/
```

`sync` puts everything where Agent Zero indexes it (`usr/knowledge/`). Ask him afterwards —
*"what do you know about our AVOP courses?"* — and he should answer from your documents.

## The commands

| Command | What it does |
|---|---|
| `python3 tools/pack_manage.py status` | what is in the pack, what is waiting in the inbox |
| `python3 tools/pack_manage.py ingest --apply` | files inbox documents into the right knowledge folder |
| `python3 tools/pack_manage.py convert --apply` | turns Word/PowerPoint/Excel into `.md` he can index |
| `python3 tools/pack_manage.py index` | rebuilds `knowledge/00_index.md` so he knows what exists |
| `python3 tools/pack_manage.py sync` | copies `knowledge/` into `usr/knowledge/main/ato_company/` |
| `python3 tools/pack_manage.py checklist` | the ATO knowledge checklist, with what is already present |

Safe defaults: ingest **copies** (add `--move` to clear the inbox), and sync never deletes
anything in the destination.

## What he can actually read

Agent Zero indexes these types on its own:

| Read directly | Needs `convert` first | Not indexed (keep as assets) |
|---|---|---|
| `.md` `.txt` `.pdf` `.html` `.csv` `.json` | `.docx` `.pptx` `.xlsx` | `.png` `.jpg` — images |

Your past courses will mostly be Word, PowerPoint and Excel — **those are invisible to him until
they are converted.** `ingest` converts them automatically and writes a `.md` beside the
original (the original is always kept). If you add Office files by hand, run `convert --apply`.

`convert` uses only the Python standard library — nothing to install. It pulls the text, tables,
slide text and speaker notes out of the file, keeps headings as headings, and drops repeated
footers and page numbers.

## Layout

```
ato_company_pack/
├── knowledge/
│   ├── company/      who you are — profile, brand, service catalogue
│   ├── alliance/     Training Centre Alliance: shared courses, standards, brand rules
│   ├── clients/      clients, contracts, appointments, quotations
│   ├── courses/      past courses, lesson plans, exams, manuals   ← highest value
│   ├── operations/   SOPs, policies, forms, assessment rules
│   ├── regulatory/   CARs, ICAO, NFPA, ACI, IATA
│   ├── airports/     one verified data pack per aerodrome
│   └── media/        logo, brand masters, photo library index
├── inbox/            ← drop new documents here
├── tools/pack_manage.py
├── tools/office2md.py    Office → Markdown converter (stdlib only)
├── KNOWLEDGE_CHECKLIST.md
└── NAMING.md
```

## Golden rules

1. **Only put in what you are happy for him to quote.** Everything here is retrievable by the agent.
2. **No secrets** in this folder — API keys belong in Settings or `.env`, never in knowledge.
3. **Mark verified aerodrome facts with their source.** Unverified values stay blank; that is a
   feature, not a gap.
4. **Do not paste licensed text** (NFPA, ACI, IATA) into knowledge you will regenerate courses
   from — store the reference and cite it.
5. **Alliance material stays in `alliance/`.** It is not yours alone: keep it in that one folder
   so it can be shared, held back, or removed cleanly if the arrangement changes.

*AM Risk and Training (Pty) Ltd — "Semper Paratus"*
