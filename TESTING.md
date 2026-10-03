# Tests

## Running them

```
npm install
npm test              # run once
npm run test:watch    # re-run on file changes
npm run test:coverage # with a coverage summary
```

No config beyond `package.json`'s `test` script — Jest works out of the box
on this plain CommonJS project.

## Layout

```
tests/
  unit/
    helpers.test.js      # normalizeSlug, resolveId, parseTags/parseCategories,
                          # resolveEndpoint, conformFields, etc. — pure, no I/O
    mapPayload.test.js    # each content type's spreadsheet-row -> API-payload mapping
    envConfig.test.js     # ENDPOINT_ALIASES_*/FIELD_ALIASES_* env-var overrides
    frontend.test.js      # the dashboard's own JS (app.js), in jsdom, against
                          # the real index.html markup
  integration/
    network.test.js       # fetchReferenceData, fetchSchemaNames/Fields, postAssetFile
    processUpload.test.js # the engine: schema resolution, batching, publish, etc.
    routes.test.js         # POST /upload/:type, POST /upload/assets, via supertest
    checkConnection.test.js # the Test Connection feature (checkSiteConnection,
                            # checkConnection, POST /check-connection)
    customTypes.test.js     # sheet-name matching, custom-type persistence, and
                            # /content-types, /inspect-workbook, /custom-types
  fixtures/
    workbook.js            # builds small in-memory .xlsx buffers with the real xlsx package
    schemaResponses.js     # a configurable mock for the global fetch calls hit content.civicplus.com
```

Nothing here touches the real `content.civicplus.com` API or the real
`EvolveUploads.xlsx` — every test runs against small synthetic fixtures.
`global.fetch` is mocked per test; workbooks are built for real with the
`xlsx` package (not mocked), so the actual spreadsheet-parsing path gets
exercised too.

## What changed in `index.js` to make this possible

1. **`uploadConfig.facility.sheetIndex` was `8`, now `9`.** The real
   workbook's sheet order is Articles(0) → Resources(1) → FAQs(2) →
   Staff(3) → Departments(4) → QuickLinks(5) → News(6) → Calendar(7) →
   *Agendas and Minutes*(8) → *Facility*(9). The Facilities button was
   reading the Agendas and Minutes sheet. See
   `tests/integration/processUpload.test.js` → `sheetIndex wiring` and
   `facility sheetIndex regression` for the tests that pin the fix (and
   would catch a repeat for any type, not just facility).
2. **More functions are exported** (`normalizeSlug`, `resolveSheetName`,
   `resolveId`, `parseTags`, `parseCategories`, `permissionSetByName`,
   `getEndpointCandidates`, `resolveEndpoint`, `suggestClosestSchemas`,
   `getFieldCandidates`, `resolveField`, `conformPartition`,
   `conformFields`, `fetchSchemaNames`, `fetchSchemaFieldMap`,
   `getSchemaNames`, `getSchemaFields`, `resolveReferenceIds`,
   `fetchContentNameLookup`, `fetchAssetNameLookup`, `loadCustomTypes`,
   `saveCustomTypes`, `buildCustomMapPayload`, `buildCustomTypeConfig`,
   `registerCustomTypes`). Purely additive — nothing else changed to rely
   on this, it just makes these testable directly instead of only
   indirectly through `mapPayload`/`processUpload`.
3. **Every content type now carries `label`, `description`, and
   `sheetName`** on its own `uploadConfig` entry. The separate
   `CONTENT_TYPE_LABELS` map is gone (one source of truth), and the GUI
   builds its buttons from `GET /content-types` rather than hardcoding
   nine of them in `index.html`.
4. **`package.json`'s `"main"`** pointed at a non-existent `app.js`; it
   now points at `index.js`. (Unrelated to the new front-end file, which
   is also called `app.js` — the entry point is and always was `index.js`.)
5. **The token field is masked** (`type="password"`), so it isn't readable
   over someone's shoulder.

"Agendas and Minutes" (sheet index 8) still isn't wired to any
`uploadConfig` entry or button — that's a separate, bigger addition
(a new content type, endpoint slug, mapPayload, and UI button) that wasn't
part of this pass. `tests/integration/processUpload.test.js` has one test
documenting that it's still unconfigured, so it won't go unnoticed.

## The Test Connection feature

A new "Test connection" button on the dashboard (and a `POST /check-connection`
route behind it) checks, using the same Site app name / Token already in the
form:

1. Whether the site is reachable and the token is accepted (`checkSiteConnection`
   — a fresh, uncached request to `/schemas`, unlike `fetchSchemaNames`/
   `getSchemaNames` above, which are best-effort and collapse every failure to
   `null` on purpose; this one keeps the specific status code so it can tell a
   bad token apart from a wrong site name apart from the site being down).
2. Which of the 9 configured content types have a matching schema on this site
   (reusing `getEndpointCandidates`/`resolveEndpoint` — the exact same
   resolution a real upload does, just run for all 9 up front instead of one
   at a time, mid-upload).
3. Whether the PermissionSet and Categories reference lists load, with a count.
4. The full list of schemas on the site, each flagged with whether a content
   type in this tool already points at it. The unflagged ones are the useful
   part — they're the candidates when registering a new content type, so the
   same list also feeds an autocomplete on every schema-slug field.

It's read-only end to end — it never creates, updates, or publishes anything,
so it's safe to click at any time. Covered by
`tests/integration/checkConnection.test.js`.

## Tests that pin CURRENT behavior, not necessarily CORRECT behavior

A few tests exist specifically to document quirks found while writing this
suite, so a future change to any of them is a deliberate decision instead
of an accident:

- **Publish is a strict `=== 'Yes'` string check** — `"yes"`, `"YES"`, and
  boolean `true` in the sheet do *not* trigger the publish step.
  `processUpload.test.js` → `publish behavior`.
- **A failed publish PATCH doesn't fail the row** — if content creation
  succeeds but the publish step then fails (bad response or network
  error), that failure is only logged; the row still counts as a
  **success** in the upload response. Same section.
- **`conformFields`'s anti-clobber guard is order-dependent** — whether a
  renamed field or a literal field "wins" when both exist depends on which
  key appears first in the mapper's output object, not on any explicit
  precedence rule. `helpers.test.js` → `conformFields`, two tests with the
  same two values in reversed key order.
- **An Excel-native date cell becomes a raw serial number, not a date
  string** — `sheet_to_json` is called with no options in `index.js`, so a
  `NewsDate` column formatted as a real date in Excel (as opposed to typed
  as text) comes through as a number like `45444`, not an ISO string.
  `processUpload.test.js` → `Excel date-typed cells`.
- **`resources.mapPayload`** builds two overlapping sets of keys
  (lowercase + PascalCase, e.g. `phonenumber` *and* `PhoneNumber`) from the
  same source values. Possibly deliberate, possibly leftover — either way
  it's pinned with a Jest snapshot (`mapPayload.test.js`) so a future
  cleanup can't silently change the payload shape. If you do change it on
  purpose, update the snapshot with `npx jest -u`.

## What this is built on (worth knowing before you touch any of the below)

`content.civicplus.com` is built on **Squidex**, an open-source headless
CMS — confirmed by the "Empty arrays are usually accepted by Squidex" note
that was already in this file (in `departments`) before any of today's
changes. That's not a guess: Squidex's own public docs
([docs.squidex.io](https://docs.squidex.io)) confirm Reference fields and
Asset fields are both arrays of raw id strings (`{iv: ["id1", "id2"]}`),
content can be listed via `GET /api/content/{app}/{schema}`, assets via
`GET /api/apps/{app}/assets`, and schemas can be created via
`POST /api/apps/{app}/schemas`. Everything below is grounded in that —
but **none of it is verified against CivicPlus's specific deployment or
your token's permissions**. CivicPlus may restrict any of this regardless
of what the underlying engine supports; a real Squidex user has hit a
permissions wall on schema creation even on Squidex's own cloud product
(github.com/Squidex/squidex/issues/110), so a 403 doesn't necessarily mean
something's configured wrong.

Separately: "CivicPlus Engage" (CivicEngage) appears, from CivicPlus's own
public documentation, to be the historical/predecessor name for what's now
branded Evolve — the same Municipal Websites platform, not a separate
product — so this is very likely already "the Engage API," just under
current branding.

## The reference-resolution fixes (the former `//FIX LATER` stubs)

All six are now real, not stubbed:

| Field | Type | Was | Now |
|---|---|---|---|
| `staff.department` | Reference | always `{en: []}` | resolves the sheet's department name against existing Department content |
| `departments.parentdepartment` | Reference | always `{iv: []}` | resolves a department name (self-referential) |
| `departments.staffdirectory` | Reference | always `{iv: []}` | resolves a comma-separated list of staff names |
| `news.newsasset` | Asset | always `{iv: []}` | resolves a filename against existing uploaded assets |
| `calendar.attachments` | Asset | always `{iv: []}` | resolves a comma-separated list of filenames |
| `calendar.submission-pdf` | Asset | always `{iv: []}` | resolves a filename |
| `news.newsdate` fallback | — | hardcoded `2019-08-24` | current date/time when the sheet is blank |
| `calendar.url-link` fallback | — | placeholder `"test.com"` | `null` when the sheet is blank |

The Reference/Asset resolution works by listing what already exists on the
site (once per run, before the batch loop, the same "load once" pattern as
permissionSet/categories) and matching by name — `fetchContentNameLookup`
for Department/Staff content, `fetchAssetNameLookup` for uploaded assets,
both new, plus `resolveReferenceIds` (like `resolveId`, but returns an
array, since Reference/Asset fields always are one). A name/filename with
no match is left out and warned about rather than failing the row —
consistent with every other soft-fail already in this file. Covered in
`helpers.test.js` (`resolveReferenceIds`), `network.test.js`
(`fetchContentNameLookup`/`fetchAssetNameLookup`), `mapPayload.test.js`,
and `processUpload.test.js` (`content & asset reference lookups`).

**Partitioning — now settled by a live run, not inference.** A real Staff
import returned `department: Not a known language.` when `department` was
sent as `{iv: [...]}`. That field is **language-partitioned (`en`)**, as
the original code had it before reference resolution was added. It's back
to `en`, with a regression test. `departments.parentdepartment` and
`staffdirectory` genuinely are `iv` — one does not imply the other, so
don't "tidy" them to match.

**Unmatched categories are dropped, not sent.** The same live run returned
`Categories: One or more Categories do not exist in App` and rejected
entire rows because a sheet value ("Fire Marshal" — a valid *permission
set* on that site, but not a category) was sent with an empty id. Rows now
upload without the bad category, and `processUpload` returns
`missingCategories` so the dashboard names exactly what was skipped
instead of changing the data silently.

**Still unverified:** the Reference/Asset array *shape* is confirmed from
Squidex's docs, but not which field Department/Staff content stores its
display name under, and which field Department/Staff content
actually stores its display name under — this assumes `department` (own
name) is `en`, matching how `departments.mapPayload` already sends it, and
that staff records store `firstname`/`lastname` the same way `staff.mapPayload`
sends them. If a real upload comes back with a field-partition mismatch,
`conformPartitions: true` on the type (see `field-name conforming` in
`processUpload.test.js`) will fix it automatically once schema fields are
readable — no code change needed first.

## Sheet matching by name (replaces matching by position)

A content type used to find its sheet purely by position (`sheetIndex: 9`),
which is exactly how the facility/"Agendas and Minutes" bug happened. Now
every type also declares a `sheetName`, and `resolveSheetName()` checks, in
order:

1. an explicit override — the per-type dropdown in the GUI, sent as a
   `sheetOverride` form field;
2. the type's `sheetName`, fuzzy-matched with the same normalization used
   for schema slugs (so `Facility`, `facility`, and `FACILITY ` all match);
3. only if neither matches, the old `sheetIndex` position — a safety net
   for a workbook whose sheet was renamed to something unrecognizable.

One deliberate asymmetry, found by a test rather than by reading the code:
`/inspect-workbook` uses **name matching only, no positional fallback**. In
a partial workbook (say just Articles and FAQs), the fallback would let
unrelated types claim whatever sheet sits at their index — `resources`
(index 1) would "match" the FAQs sheet — which in turn hides genuinely
unmatched sheets from the custom-import list. The fallback still applies at
upload time, where the user has deliberately chosen one type.

## Custom content types

Defined from the GUI, not by editing code. Two ways in:

- **Register a content type** — always available in the sidebar, with or
  without a workbook loaded. Give it a display name, the schema slug, and
  the sheet name to read. Columns aren't asked for, because with no
  workbook there are no headers to read; they're discovered from each row
  at upload time.
- **Unrecognized sheets** in a loaded workbook get a one-click shortcut
  with their columns already filled in. Columns registered this way are
  *pinned*, so a column added to that sheet later doesn't silently start
  uploading.

Either way every column becomes a plain text field (`{iv: value}` — the same shape
`processUpload` already uses for unrecognized columns on a built-in type,
so this is consistent with existing behavior rather than a new convention).
Blank cells are skipped; `PermissionSet`, `Categories`, `Publish`, `Tags`,
and `Name` are handled exactly as they are for built-in types.

Definitions are saved to `customTypes.json` next to `index.js` and loaded
at startup, which matters because the `.bat` file kills and restarts the
server every run. A corrupt or non-object file is ignored with a warning
rather than stopping the server. Once registered, a custom type flows
through the same upload route, connection check, and sheet matching as a
built-in one, with no special-casing — `customTypes.test.js` includes an
end-to-end test that defines one and uploads through it.

Custom types get `sheetIndex: -1` so they are **only** ever matched by
name; there is no meaningful fixed position for one, and -1 turns a
name-match failure into a clear error instead of silently grabbing an
arbitrary sheet.

**This tool does not create schemas on the site.** An earlier pass added a
Create Schema button that POSTed to the CMS's schema-creation endpoint;
it has been removed entirely, along with its route and tests. Which content
types a site should have is a deliberate, site-by-site decision made by
whoever administers the CMS — the uploader only learns how to send rows to
a schema that already exists. `frontend.test.js` asserts the connection
report offers no such action, so it can't quietly come back.

## Front-end tests

`app.js` used to be an inline `<script>` in `index.html`, which made it
untestable. It's now a separate file served at `/app.js`, and
`tests/unit/frontend.test.js` runs it in jsdom against the **real**
`index.html` parsed off disk — not a hand-written fixture — so renaming or
deleting an element breaks a test instead of silently breaking the page.
`fetch` is mocked throughout; nothing contacts a server.

Two mechanics worth knowing if you extend these:
- Each test reloads the page by compiling `app.js` with `new Function(...)`.
  A plain re-run would collide on its top-level `const`s, and
  `vm.runInThisContext` would run the code outside jsdom's globals (no
  `document`). This mirrors a browser giving each page load a fresh scope.
- `flushPromises()` pumps several macrotask turns, because the startup path
  chains multiple awaits before the buttons exist.

## A couple of testability notes for anyone extending this suite

- `getSchemaNames`/`getSchemaFields` cache results in module-level `Map`s
  that live for the whole test run. Every test that calls them (directly
  or via `processUpload`) uses a **never-reused site string** so tests
  can't see each other's cached results — there's a `uniqueSite()` helper
  at the top of each affected file for this.
- `uploadConfig` is exported and mutable, which a few tests lean on
  deliberately: injecting a temporary synthetic content type (to test the
  "no endpoint configured" guard) or flipping `conformPartitions` on a real
  type for one test, always restored in a `finally`/`afterEach`.
- `customTypes.json` is real on-disk state shared by the whole test run.
  `customTypes.test.js` snapshots whatever is there in `beforeAll` and
  restores it in `afterAll`, deletes the file between tests, and strips
  any `isCustom` entry it registered out of `uploadConfig` in `afterEach`
  — otherwise a custom type defined by one test leaks into every later
  test's view of the config, and running the suite would clobber a real
  `customTypes.json` sitting next to `index.js`.
