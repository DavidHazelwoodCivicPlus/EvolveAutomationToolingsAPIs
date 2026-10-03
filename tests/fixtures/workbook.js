// Builds small, in-memory .xlsx workbooks for tests, using the same `xlsx`
// package index.js already depends on. Nothing here mocks XLSX — a real
// workbook is written to a real Buffer and read back through the exact
// same XLSX.read(buffer, {type:'buffer'}) + XLSX.utils.sheet_to_json(sheet)
// call shape index.js uses, so tests exercise the real parsing path.

const XLSX = require('xlsx');

// The real EvolveUploads.xlsx sheet order. Content types in uploadConfig
// are wired to these sheets BY INDEX, so getting this order right (and
// keeping it right) is what the regression test in
// tests/integration/processUpload.test.js depends on.
const REAL_SHEET_ORDER = [
    'Articles',        // 0
    'Resources',       // 1
    'FAQs',             // 2
    'Staff',            // 3
    'Departments',     // 4
    'QuickLinks',       // 5
    'News',             // 6
    'Calendar',         // 7
    'Agendas and Minutes', // 8 — not wired to any uploadConfig entry
    'Facility'          // 9
];

// sheets: an object of { SheetName: [rowObject, ...] }, in the order you
// want them to appear in the workbook (insertion order of object keys).
// `opts` is passed through to json_to_sheet (e.g. { cellDates: true } to
// force real Date objects to serialize as date-typed cells rather than
// numbers — see the News/Calendar date tests, which deliberately do NOT
// pass this, to match how index.js actually reads uploaded files).
function buildWorkbook(sheets, opts = {}) {
    const wb = XLSX.utils.book_new();
    for (const [name, rows] of Object.entries(sheets)) {
        // json_to_sheet needs at least an empty array; an empty sheet still
        // gets a valid (headerless) worksheet so SheetNames stays correct.
        const ws = XLSX.utils.json_to_sheet(rows || [], opts);
        XLSX.utils.book_append_sheet(wb, ws, name);
    }
    return wb;
}

function workbookToBuffer(wb) {
    return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

function buildWorkbookBuffer(sheets, opts) {
    return workbookToBuffer(buildWorkbook(sheets, opts));
}

// A single-sheet workbook at a given position, for tests that only care
// about one content type. Cheaper than building all 10 sheets every time.
// `sheetIndex` pads with empty sheets before the real one so the target
// sheet lands at exactly the index the test wants to exercise.
function buildSingleSheetWorkbook(sheetName, rows, sheetIndex = 0) {
    const sheets = {};
    for (let i = 0; i < sheetIndex; i++) {
        sheets[`Padding${i}`] = [];
    }
    sheets[sheetName] = rows;
    return buildWorkbook(sheets);
}

// Minimal realistic row data for all 10 real sheets, in the real order —
// used by the full end-to-end wiring regression test. Every row has just
// enough fields to satisfy each type's mapPayload without throwing.
function buildFullWorkbook(rowOverrides = {}) {
    const defaults = {
        Articles: [{ Title: 'July Test', ContentType: 'article', Content: '<p>hi</p>', Publish: 'Yes', PermissionSet: 'General', Categories: 'Fire', Name: 'Tester' }],
        Resources: [{ Department: 'Fire Dept', Name: 'Fire Dept', PermissionSet: 'General', Categories: 'Fire', Publish: 'Yes' }],
        FAQs: [{ Question: 'Q1', Answer: 'A1', PermissionSet: 'General', Categories: 'Fire', Publish: 'Yes' }],
        Staff: [{ FirstName: 'Jo', LastName: 'Schmoe', PermissionSet: 'General', Categories: 'Fire', Publish: 'Yes' }],
        Departments: [{ Department: 'Parks', PermissionSet: 'General', Categories: 'Fire', Publish: 'Yes' }],
        QuickLinks: [{ Link: 'https://example.com', Name: 'Example', PermissionSet: 'General', Categories: 'Fire', Publish: 'Yes' }],
        News: [{ NewsTitle: 'Big News', PermissionSet: 'General', Categories: 'Fire', Publish: 'Yes' }],
        Calendar: [{ TitleOfEvent: 'Town Hall', PermissionSet: 'General', Categories: 'Fire', Publish: 'Yes' }],
        'Agendas and Minutes': [{ PermissionSet: 'General', Categories: 'Fire', Publish: 'Yes' }],
        Facility: [{ FacilityName: 'Rec Center', PermissionSet: 'General', Categories: 'Fire', Publish: 'Yes' }]
    };
    const merged = { ...defaults, ...rowOverrides };
    // Preserve REAL_SHEET_ORDER regardless of key order in the overrides.
    const ordered = {};
    for (const name of REAL_SHEET_ORDER) ordered[name] = merged[name];
    return buildWorkbook(ordered);
}

module.exports = {
    REAL_SHEET_ORDER,
    buildWorkbook,
    workbookToBuffer,
    buildWorkbookBuffer,
    buildSingleSheetWorkbook,
    buildFullWorkbook
};
