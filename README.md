# Content Auditor

A local browser-based content auditor built from `Guidelines for Authors.docx`.

## What it checks

- Banned wording such as `also`, opinion phrases, weak certainty, and back-references.
- Sentences that start with `if` or `because`.
- Long sentences, passive voice, half-sentence list introductions, and possibility modals.
- Central entity coverage across the intro and sections.
- Central search intent coverage in the intro.
- Direct answers under question headings.
- Research, numeric detail, examples, units, percentages, and first-use abbreviations.
- Featured snippet length, image introductions, table comparison language, and anchor placement.

## Run

```powershell
python app.py 8000
```

Then open `http://127.0.0.1:8000`.

## Supported input

Paste content directly, or upload `.txt`, `.md`, `.docx`, `.pdf`, `.html`, `.csv`, or `.rtf`.
