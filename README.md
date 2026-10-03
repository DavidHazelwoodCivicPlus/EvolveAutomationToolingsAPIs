# CivicPlus Evolve API Automation Tooling

A local Node.js tool with a small web GUI for bulk-importing website content into CivicPlus Evolve sites via the Evolve REST API — instead of manually creating hundreds or thousands of articles, staff records, FAQs, etc. through the CMS UI.

## What it does

1. You prep an Excel workbook (`EvolveUploads.xlsx`) with one sheet per content type.
2. Start the local server (`node index.js`, or the included `.bat` on Windows), it serves a dashboard at `http://localhost:4000`.
3. Enter the target Evolve site's app name and an access token. Optionally click **Test connection** first — it's read-only and reports whether the site is reachable, the token works, which content types actually exist on that site, and the full list of schemas on the site (flagging which ones this tool already uses).
4. Choose your workbook. The tool reads its sheet names and matches each one to a content type.
5. Click the button for the content type you want to upload. The server reads the matching sheet, maps each row into the JSON payload shape the Evolve API expects, and POSTs the records in batches of 10, reporting how many succeeded.

## Sheets are matched by name

Each content type looks for a sheet by **name**, not by position:

| Upload type | Sheet name | Notes |
|---|---|---|
| Articles | `Articles` | Title, content, content type, tags, categories |
| Resources | `Resources` | Department contact info, address, hours, website link |
| FAQs | `FAQs` | Question/answer pairs |
| Staff | `Staff` | Name, title, contact info, bio, department reference |
| Departments | `Departments` | Contact info, hours, parent department, staff directory |
| Quick Links | `QuickLinks` | Link records |
| News | `News` | Title, text, date, asset reference |
| Calendar | `Calendar` | Event name, date, time, details, attachments |
| Facilities | `Facility` | Facility name |

Matching ignores case and punctuation, so `Facility`, `facility`, and `FACILITY ` all work. If a sheet has been renamed to something unrecognizable, the tool falls back to the old fixed position — and either way you can override the choice per content type with the **Sheet** dropdown next to each button once a workbook is loaded.

## Custom content types

Two ways to give a content type its own upload button:

- **Register a content type** — a button in the sidebar, always available, no workbook required. Enter a display name, the schema slug on the site, and the sheet name it should read. Run **Test connection** first and the schema slug field autocompletes from the real schemas on that site, so you don't have to look one up or type it from memory. Every column in that sheet is uploaded as a plain text field.
- **Unrecognized sheets** — load a workbook and any sheet that doesn't match a known type appears under *Unrecognized sheets in this workbook*, with its columns already listed, for one-click registration. Columns added this way are pinned to what the sheet had at that moment.

Custom definitions are saved to `customTypes.json` next to `index.js`, so they survive the server restart the `.bat` file performs each run. Remove one from the **Saved custom types** list in the sidebar.

**The schema must already exist on the site.** This tool never creates or modifies schemas — which content types a given site should have is a decision for whoever administers that CMS. The tool only learns how to send rows to a schema that's already there.

## Project structure

| File | Purpose |
|---|---|
| `index.js` | Express server: serves the GUI, exposes `POST /upload/:type` plus the connection-check, workbook-inspection, and custom-type routes; reads the relevant sheet, maps rows to Evolve API payloads, and batches the upload |
| `app.js` | Front-end dashboard logic (buttons, sheet pickers, custom imports, connection check) |
| `index.html` | Dashboard markup |
| `style.css` | Dashboard styling |
| `EvolveUploads.xlsx` | The source data workbook, one sheet per content type |
| `customTypes.json` | Saved custom type definitions (created automatically; absent until you add one) |
| `tests/` | Automated test suite — see `TESTING.md` |
| `CivicPlus Evolve Article Import.bat` | Windows convenience script: kills any existing `node.exe`, starts the server, opens the dashboard |

## Requirements

```
npm install
```

## Usage

**Windows:** double-click `CivicPlus Evolve Article Import.bat` — it starts the server and opens `http://localhost:4000` automatically.

**Manual:**
```
node index.js
```
Then open `http://localhost:4000` in a browser.

## Running the tests

```
npm test
```

Runs the full suite (293 tests). Nothing touches a real site — every test runs against mocked responses and synthetic workbooks. Run this before publishing new builds; see `TESTING.md` for what's covered and how it's structured.

## Adding a new built-in content type

Most new types should be added through the GUI as custom imports (above) — no code change needed. Add one in code only when it needs real per-field logic (date ranges, reference resolution, compound fields) that a plain text mapping can't express: add an entry to `uploadConfig` in `index.js` with a `sheetName`, `label`, `description`, `endpoint`, and a `mapPayload` function. The button appears automatically; `index.html` doesn't need touching.

## Known limitations

- **Not yet validated against a live site.** The reference-resolution and custom-type features are built against Squidex's documented API behavior (Evolve's content API is built on Squidex) but haven't been run against a real CivicPlus site yet.
- **"Agendas and Minutes"** has no built-in content type. Add it through the GUI as a custom import.
- **Excel-native date cells** come through as raw serial numbers rather than date strings. Format date columns as text if a date field rejects the value.
- One npm advisory exists in the `xlsx` dependency; it predates this tooling work.
