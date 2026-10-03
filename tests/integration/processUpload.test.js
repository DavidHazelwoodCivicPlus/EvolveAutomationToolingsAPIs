const { processUpload, uploadConfig } = require('../../index.js');
const { buildFetchMock, jsonResponse, textResponse } = require('../fixtures/schemaResponses.js');
const { buildWorkbook, buildFullWorkbook, buildSingleSheetWorkbook, REAL_SHEET_ORDER } = require('../fixtures/workbook.js');

const originalFetch = global.fetch;
let warnSpy, logSpy, errorSpy;

beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    logSpy = jest.spyOn(console, 'log').mockImplementation(() => {});
    errorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
});
afterEach(() => {
    warnSpy.mockRestore();
    logSpy.mockRestore();
    errorSpy.mockRestore();
});
afterAll(() => {
    global.fetch = originalFetch;
});

// getSchemaNames/getSchemaFields cache per site in module-level Maps that
// outlive any single test — a fresh, never-reused site per test sidesteps
// that without touching index.js. See tests/integration/network.test.js for
// tests of the caching behavior itself.
let siteCounter = 0;
const uniqueSite = () => `process-upload-test-${siteCounter++}.example.com`;

// Row-creation POSTs (as opposed to /schemas, reference-data, or publish
// PUT calls) are what carry the actual mapped payload — most assertions
// below inspect exactly one of these.
function creationCalls(fetchMock) {
    return fetchMock.mock.calls.filter(([url, opts]) =>
        (opts?.method || 'GET').toUpperCase() === 'POST' &&
        /\/api\/content\/[^/]+\/[^/]+$/.test(new URL(String(url)).pathname)
    );
}

describe('sheetIndex wiring (static — no fetch needed)', () => {
    const expectedSheetNameByType = {
        articles: 'Articles',
        resources: 'Resources',
        faqs: 'FAQs',
        staff: 'Staff',
        departments: 'Departments',
        quicklinks: 'QuickLinks',
        news: 'News',
        calendar: 'Calendar',
        facility: 'Facility'
    };

    test.each(Object.entries(expectedSheetNameByType))(
        '%s.sheetIndex points at "%s" in the real workbook\'s sheet order',
        (type, expectedName) => {
            expect(REAL_SHEET_ORDER[uploadConfig[type].sheetIndex]).toBe(expectedName);
        }
    );

    test('"Agendas and Minutes" (index 8) is not wired to any content type yet', () => {
        const wiredIndexes = Object.values(uploadConfig).map(c => c.sheetIndex);
        expect(wiredIndexes).not.toContain(REAL_SHEET_ORDER.indexOf('Agendas and Minutes'));
    });
});

describe('facility sheetIndex regression (end-to-end against a full, realistically-ordered workbook)', () => {
    test('uploading "facility" reads the real Facility sheet, not Agendas and Minutes', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock();
        const workbook = buildFullWorkbook({
            Facility: [{ FacilityName: 'Rec Center', Publish: 'Yes', PermissionSet: 'General', Categories: 'Fire' }],
            'Agendas and Minutes': [{ PermissionSet: 'General', Categories: 'Fire' }] // no FacilityName here on purpose
        });

        const { successes, failures } = await processUpload('facility', site, 'token', workbook, 10);

        expect(failures).toEqual([]);
        expect(successes).toHaveLength(1);
        const [, options] = creationCalls(global.fetch)[0];
        expect(JSON.parse(options.body).data.facilityname).toEqual({ en: 'Rec Center' });
    });
});

describe('endpoint resolution', () => {
    test('resolves via an exact slug match', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['article'] });
        await processUpload('articles', site, 'token', buildSingleSheetWorkbook('Articles', [{ Title: 'T', Content: 'C' }], 0), 10);
        const [url] = creationCalls(global.fetch)[0];
        expect(new URL(String(url)).pathname).toBe(`/api/content/${site}/article`);
    });

    test('falls back to a normalized match when the real slug differs only by punctuation/casing', () => {
        return (async () => {
            const site = uniqueSite();
            global.fetch = buildFetchMock({ schemaNames: ['BP-Quick-Link'] });
            await processUpload('quicklinks', site, 'token', buildSingleSheetWorkbook('QuickLinks', [{ Link: 'https://x.com', Name: 'X' }], 5), 10);
            const [url] = creationCalls(global.fetch)[0];
            expect(new URL(String(url)).pathname).toBe(`/api/content/${site}/BP-Quick-Link`);
        })();
    });

    test('throws a descriptive error naming the tried candidates when no schema matches', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['something-unrelated'] });
        await expect(
            processUpload('staff', site, 'token', buildSingleSheetWorkbook('Staff', [{ FirstName: 'A' }], 3), 10)
        ).rejects.toThrow(/employee/);
    });

    test('falls back to the first candidate (with a warning) when /schemas itself is unavailable', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNamesStatus: 500 });
        const { successes, failures } = await processUpload(
            'articles', site, 'token', buildSingleSheetWorkbook('Articles', [{ Title: 'T', Content: 'C' }], 0), 10
        );
        expect(failures).toEqual([]);
        expect(successes).toHaveLength(1);
        expect(warnSpy).toHaveBeenCalled();
    });
});

describe('field-name conforming', () => {
    test('renames a field to the schema\'s real spelling via fieldAliases', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({
            schemaNames: ['employee'],
            schemaFields: { employee: [{ name: 'phone', partitioning: 'invariant' }, { name: 'firstname', partitioning: 'invariant' }] }
        });
        await processUpload('staff', site, 'token', buildSingleSheetWorkbook('Staff', [{ FirstName: 'Jo', PhoneNumber: '555' }], 3), 10);

        const body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        expect(body.data.phone).toEqual({ en: '555' });
        expect(body.data.phonenumber).toBeUndefined();
    });

    test('leaves field names untouched when the schema\'s fields cannot be read', async () => {
        const site = uniqueSite();
        // schemaNames resolves fine, but no schemaFields.employee entry is
        // configured, so fetchSchemaFieldMap -> 404 -> getSchemaFields -> null.
        global.fetch = buildFetchMock({ schemaNames: ['employee'] });
        await processUpload('staff', site, 'token', buildSingleSheetWorkbook('Staff', [{ FirstName: 'Jo', PhoneNumber: '555' }], 3), 10);

        const body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        expect(body.data.phonenumber).toEqual({ en: '555' });
    });

    test('conformPartitions (opt-in per type) also fixes the {en}/{iv} wrapper to match the schema', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({
            schemaNames: ['employee'],
            schemaFields: { employee: [{ name: 'firstname', partitioning: 'invariant' }] }
        });
        const original = uploadConfig.staff.conformPartitions;
        uploadConfig.staff.conformPartitions = true;
        try {
            await processUpload('staff', site, 'token', buildSingleSheetWorkbook('Staff', [{ FirstName: 'Jo' }], 3), 10);
            const body = JSON.parse(creationCalls(global.fetch)[0][1].body);
            // mapPayload emits firstname as {en:'Jo'}; the schema says invariant.
            expect(body.data.firstname).toEqual({ iv: 'Jo' });
        } finally {
            uploadConfig.staff.conformPartitions = original;
        }
    });
});

describe('batching and row numbering', () => {
    test('processes every row across a batch boundary; Excel row numbers use the +2 header offset', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['faq'] });
        const rows = Array.from({ length: 5 }, (_, i) => ({ Question: `Q${i}`, Answer: `A${i}` }));

        const { successes, failures } = await processUpload('faqs', site, 'token', buildSingleSheetWorkbook('FAQs', rows, 2), 2);

        expect(failures).toEqual([]);
        expect(successes.map(s => s.row).sort((a, b) => a - b)).toEqual([2, 3, 4, 5, 6]);
        expect(creationCalls(global.fetch)).toHaveLength(5);
    });

    test('reference data (permissionSet, categories) is fetched once per run, not once per row', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['faq'] });
        const rows = [{ Question: 'Q1', Answer: 'A1' }, { Question: 'Q2', Answer: 'A2' }, { Question: 'Q3', Answer: 'A3' }];

        await processUpload('faqs', site, 'token', buildSingleSheetWorkbook('FAQs', rows, 2), 1); // batchSize 1 -> 3 batches

        const refDataCalls = global.fetch.mock.calls.filter(([url, opts]) => {
            const u = new URL(String(url));
            return (opts?.method || 'GET').toUpperCase() === 'GET' && /\/api\/apps\/[^/]+\/(permissionSet|categories)$/.test(u.pathname);
        });
        expect(refDataCalls).toHaveLength(2); // permissionSet once, categories once
    });
});

describe('success/failure partitioning', () => {
    test('a row whose creation fails ends up in failures (with the API error message); others still succeed', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({
            schemaNames: ['faq'],
            createHandler: (body) =>
                body.data.question.en === 'FAIL_ME'
                    ? textResponse('validation error', 422)
                    : jsonResponse({ id: 'ok-id' })
        });
        const rows = [{ Question: 'Q1', Answer: 'A1' }, { Question: 'FAIL_ME', Answer: 'A2' }, { Question: 'Q3', Answer: 'A3' }];

        const { successes, failures } = await processUpload('faqs', site, 'token', buildSingleSheetWorkbook('FAQs', rows, 2), 10);

        expect(successes).toHaveLength(2);
        expect(failures).toHaveLength(1);
        expect(failures[0].row).toBe(3); // the 2nd data row: i=0, j=1 -> 0+1+2
        expect(failures[0].error).toMatch(/422/);
    });
});

describe('row-creation error messages', () => {
    test('a 404 gets a schema-slug-specific hint, distinct from other error codes', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['faq'], createHandler: () => textResponse('not found', 404) });

        const { failures } = await processUpload('faqs', site, 'token', buildSingleSheetWorkbook('FAQs', [{ Question: 'Q', Answer: 'A' }], 2), 10);

        expect(failures[0].error).toMatch(/Verify the schema slug/);
    });

    test('a non-404 error surfaces the raw status and response body instead', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['faq'], createHandler: () => textResponse('field "answer" is required', 422) });

        const { failures } = await processUpload('faqs', site, 'token', buildSingleSheetWorkbook('FAQs', [{ Question: 'Q', Answer: 'A' }], 2), 10);

        expect(failures[0].error).toMatch(/422/);
        expect(failures[0].error).toMatch(/field "answer" is required/);
    });
});

describe('publish behavior', () => {
    test('only an exact "Yes" string triggers the publish PATCH — "yes" and boolean true do not', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['faq'] });
        const rows = [
            { Question: 'Q1', Answer: 'A1', Publish: 'Yes' },
            { Question: 'Q2', Answer: 'A2', Publish: 'yes' },
            { Question: 'Q3', Answer: 'A3', Publish: true }
        ];

        await processUpload('faqs', site, 'token', buildSingleSheetWorkbook('FAQs', rows, 2), 10);

        const publishCalls = global.fetch.mock.calls.filter(([, opts]) => (opts?.method || '').toUpperCase() === 'PUT');
        expect(publishCalls).toHaveLength(1);
    });

    test('if the publish PATCH itself fails (non-2xx), the row still counts as a success', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['faq'], publishStatus: 500 });

        const { successes, failures } = await processUpload(
            'faqs', site, 'token', buildSingleSheetWorkbook('FAQs', [{ Question: 'Q1', Answer: 'A1', Publish: 'Yes' }], 2), 10
        );

        expect(failures).toEqual([]);
        expect(successes).toHaveLength(1);
        expect(errorSpy).toHaveBeenCalled(); // logged, never thrown or counted as a failure
    });

    test('a network error during the publish PATCH also still counts the row as a success', async () => {
        const site = uniqueSite();
        const base = buildFetchMock({ schemaNames: ['faq'] });
        global.fetch = jest.fn(async (url, opts) => {
            if ((opts?.method || '').toUpperCase() === 'PUT') throw new Error('socket hang up');
            return base(url, opts);
        });

        const { successes, failures } = await processUpload(
            'faqs', site, 'token', buildSingleSheetWorkbook('FAQs', [{ Question: 'Q1', Answer: 'A1', Publish: 'Yes' }], 2), 10
        );

        expect(failures).toEqual([]);
        expect(successes).toHaveLength(1);
    });
});

describe('dynamic custom-field injection', () => {
    test('a column not in standardColumns is injected into data as {iv: value}', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['faq'] });
        await processUpload('faqs', site, 'token', buildSingleSheetWorkbook('FAQs', [{ Question: 'Q1', Answer: 'A1', Notes: 'extra note' }], 2), 10);

        const body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        expect(body.data.Notes).toEqual({ iv: 'extra note' });
    });

    test('an empty-string value in a non-standard column is not injected at all', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['faq'] });
        await processUpload('faqs', site, 'token', buildSingleSheetWorkbook('FAQs', [{ Question: 'Q1', Answer: 'A1', Notes: '' }], 2), 10);

        const body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        expect(body.data.Notes).toBeUndefined();
    });
});

describe('content & asset reference lookups (staff.department, departments.parentdepartment/staffdirectory, news.newsasset, calendar.attachments/submission-pdf)', () => {
    test('staff.department resolves against existing Department content, loaded once before the batch', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({
            schemaNames: ['employee', 'enhanced-department'],
            contentListings: { 'enhanced-department': [{ id: 'dept-1', data: { department: { en: 'Fire Department' } } }] }
        });

        await processUpload('staff', site, 'token', buildSingleSheetWorkbook('Staff', [{ FirstName: 'Jo', Department: 'Fire Department' }], 3), 10);

        const body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        expect(body.data.department).toEqual({ en: ['dept-1'] });
    });

    test('departments.parentdepartment (self-referential) and staffdirectory both resolve in one run', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({
            schemaNames: ['enhanced-department', 'employee'],
            contentListings: {
                'enhanced-department': [{ id: 'dept-parent-1', data: { department: { en: 'City Hall' } } }],
                employee: [{ id: 'staff-1', data: { firstname: { en: 'Joe' }, lastname: { en: 'Schmoe' } } }]
            }
        });

        const rows = [{ Department: 'Fire', ParentDepartment: 'City Hall', StaffDirectory: 'Joe Schmoe' }];
        await processUpload('departments', site, 'token', buildSingleSheetWorkbook('Departments', rows, 4), 10);

        const body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        expect(body.data.parentdepartment).toEqual({ iv: ['dept-parent-1'] });
        expect(body.data.staffdirectory).toEqual({ iv: ['staff-1'] });
    });

    test('news.newsasset and calendar.attachments/submission-pdf resolve against existing assets', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({
            schemaNames: ['newsflash'],
            referenceData: { assets: [{ id: 'asset-1', fileName: 'flyer.pdf' }] }
        });
        await processUpload('news', site, 'token', buildSingleSheetWorkbook('News', [{ NewsTitle: 'T', NewsAsset: 'flyer.pdf' }], 6), 10);
        let body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        expect(body.data.newsasset).toEqual({ iv: ['asset-1'] });

        const site2 = uniqueSite();
        global.fetch = buildFetchMock({
            schemaNames: ['event'],
            referenceData: { assets: [{ id: 'asset-1', fileName: 'flyer.pdf' }] }
        });
        await processUpload('calendar', site2, 'token', buildSingleSheetWorkbook('Calendar', [{ TitleOfEvent: 'T', Attachments: 'flyer.pdf', SubmissionPDF: 'flyer.pdf' }], 7), 10);
        body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        expect(body.data.attachments).toEqual({ iv: ['asset-1'] });
        expect(body.data['submission-pdf']).toEqual({ iv: ['asset-1'] });
    });

    test('a reference lookup that fails to load (e.g. 403) leaves that reference empty instead of failing the whole upload', async () => {
        const site = uniqueSite();
        global.fetch = jest.fn(async (url, opts) => {
            const u = new URL(String(url));
            if (u.pathname.endsWith('/schemas')) return jsonResponse(['employee', 'enhanced-department']);
            if (/\/api\/content\/[^/]+\/enhanced-department$/.test(u.pathname)) return textResponse('forbidden', 403);
            if ((opts.method || 'GET').toUpperCase() === 'GET') return jsonResponse({ items: [] });
            if ((opts.method || '').toUpperCase() === 'POST') return jsonResponse({ id: 'ok' });
            throw new Error('unexpected: ' + url);
        });

        const { successes, failures } = await processUpload('staff', site, 'token', buildSingleSheetWorkbook('Staff', [{ FirstName: 'Jo', Department: 'Fire Department' }], 3), 10);

        expect(failures).toEqual([]);
        expect(successes).toHaveLength(1); // the row still uploads — department is just empty
    });
});

// Reproduces a real failure from a live site (server log, Staff import):
// three rows rejected with "Categories: One or more Categories do not
// exist in App" because the sheet's "Fire Marshal" category existed as a
// PERMISSION SET but not as a category, and one row rejected with
// "department: Not a known language" because department was sent as `iv`.
describe('regression: the live Staff import failure', () => {
    test('a category that exists as a permission set but not a category no longer fails the row', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({
            schemaNames: ['employee', 'enhanced-department'],
            referenceData: {
                // "Fire Marshal" IS a permission set...
                permissionSet: [{ id: 'ps-fm', name: 'Fire Marshal' }],
                // ...but is NOT a category on this site.
                categories: [{ id: 'cat-1', name: 'Something Else' }]
            }
        });
        const rows = [{ FirstName: 'Albert', LastName: 'Scott III', PermissionSet: 'Fire Marshal', Categories: 'Fire Marshal', Publish: 'Yes' }];

        const { successes, failures, missingCategories } = await processUpload('staff', site, 'token', buildSingleSheetWorkbook('Staff', rows, 3), 10);

        expect(failures).toEqual([]);
        expect(successes).toHaveLength(1);

        const body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        // The bad category is gone entirely, rather than sent with no id.
        expect(body.categories).toBeUndefined();
        // The permission set still resolves — it really does exist.
        expect(body.permissionSet).toEqual({ id: 'ps-fm', name: 'Fire Marshal' });
        // And the caller is told what was dropped.
        expect(missingCategories).toEqual(['Fire Marshal']);
    });

    test('department is sent as `en`, the partition the live API accepts', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({
            schemaNames: ['employee', 'enhanced-department'],
            contentListings: { 'enhanced-department': [{ id: 'dept-1', data: { department: { en: 'Jamie Parks and Recreation' } } }] }
        });
        const rows = [{ FirstName: 'Fred', LastName: 'Migliaccio', Department: 'Jamie Parks and Recreation', PermissionSet: 'Admin' }];

        await processUpload('staff', site, 'token', buildSingleSheetWorkbook('Staff', rows, 3), 10);

        const body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        expect(body.data.department).toEqual({ en: ['dept-1'] });
        expect(body.data.department.iv).toBeUndefined();
    });

    test('missingCategories lists each distinct unmatched name once, across all rows', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['employee', 'enhanced-department'] });
        const rows = [
            { FirstName: 'A', Categories: 'Fire Marshal' },
            { FirstName: 'B', Categories: 'Fire Marshal' },
            { FirstName: 'C', Categories: 'Another Missing' }
        ];

        const { missingCategories } = await processUpload('staff', site, 'token', buildSingleSheetWorkbook('Staff', rows, 3), 10);

        expect(missingCategories.sort()).toEqual(['Another Missing', 'Fire Marshal']);
    });
});

describe('missing-sheet guard (what a single-sheet .csv upload hits for any non-Articles type)', () => {
    test('throws a clear, specific error instead of a low-level XLSX error', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock();
        // A .csv can only ever produce one sheet — this is what
        // XLSX.read(...) on a CSV buffer looks like to processUpload.
        const workbook = buildWorkbook({ Sheet1: [{ Question: 'Q', Answer: 'A' }] });

        await expect(processUpload('faqs', site, 'token', workbook, 10)).rejects.toThrow(/no sheet matching "FAQs"/);
    });
});

describe('defensive: a content type with no endpoint/endpoints configured at all', () => {
    afterEach(() => {
        delete uploadConfig.__test_no_endpoint__;
    });

    test('throws a clear configuration error rather than attempting to resolve anything (uploadConfig is exported, so tests can inject a synthetic type)', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock();
        uploadConfig.__test_no_endpoint__ = { sheetIndex: 0, standardColumns: [], mapPayload: () => ({ data: {} }) };

        const workbook = buildWorkbook({ Sheet1: [{ A: 1 }] });

        await expect(processUpload('__test_no_endpoint__', site, 'token', workbook, 10))
            .rejects.toThrow(/No endpoint\(s\) configured/);
    });
});

describe('Excel date-typed cells (worth knowing about, not necessarily a bug to fix here)', () => {
    test('a NewsDate cell stored as a real Excel date comes through as a raw serial number, not an ISO string', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['newsflash'] });
        const workbook = buildSingleSheetWorkbook(
            'News',
            [{ NewsTitle: 'T', NewsDate: new Date('2024-06-01T00:00:00Z'), Publish: 'Yes' }],
            6
        );

        await processUpload('news', site, 'token', workbook, 10);

        const body = JSON.parse(creationCalls(global.fetch)[0][1].body);
        // sheet_to_json is called with no options in index.js, so an
        // Excel-native date cell comes back as a number, not a string. If
        // the Evolve API expects an ISO date string for this field, a real
        // date-formatted NewsDate column would send something it likely
        // can't use — worth knowing, independent of whether it's "fixed".
        expect(typeof body.data.newsdate.iv.startDate).toBe('number');
    });
});
