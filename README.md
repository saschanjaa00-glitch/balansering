# Novaschem TXT Viewer

Web app for exploring Novaschem TXT exports before SATS import.

## What It Shows

- Student list with searchable names and student numbers
- Subject choices per student
- Group and block assignments per subject
- Subject and group breakdown by block
- Subject overview with student counts and block coverage

## Supported Novaschem Tables

The parser reads these tables and joins them:

- `Student`
- `Subject`
- `TA`
- `Group`
- `Group_Student`

## Run Locally

```bash
npm install
npm run dev
```

Open `http://localhost:5173/` and upload a Novaschem TXT export.

## Validate Build

```bash
npm run build
npm run lint
```

## Notes

- TXT files are parsed client-side in the browser.
- Encoding is auto-detected between UTF-8 and Windows-1252 to keep Norwegian characters readable.
- If `TA.Blockname` is empty, blocks are inferred from suffixes in long group/subject codes:
	- `A` -> `Blokk 1`
	- `B` -> `Blokk 2`
	- `C` -> `Blokk 3`
	- `D` -> `Blokk 4`
