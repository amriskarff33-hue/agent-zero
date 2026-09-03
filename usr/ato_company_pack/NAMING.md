# Naming convention

He sorts documents by **file name**, so a good name files itself.

## Courses

```
<course-type>-<airport-code>-<year>-<document>.<ext>
```

Examples:

```
avop-faxx-2026-learner-manual.docx
avop-faxx-2026-exam-paper-a.docx
arff-faxx-2026-lesson-plan-03.docx
firefighter-1-client-site-2024-progress-report.pdf
```

## Aerodromes

One markdown file per aerodrome, named by ICAO code:

```
knowledge/airports/FAXX.md
knowledge/airports/FAYY.md
```

## Regulatory instruments

```
<instrument>-<part>.<ext>
```

```
car-part-139.pdf
sa-cats-139.pdf
icao-annex-14-vol1.pdf
nfpa-1003.pdf
```

## Rules of thumb

- Lowercase, hyphens instead of spaces
- Keep the year in course filenames
- Put the document *type* last — that is what he reads first
- Never use "final", "v2", "new" — version numbers make him guess

## If he files something wrongly

Rename it and run `ingest --apply` again, or just move the file in Finder.
