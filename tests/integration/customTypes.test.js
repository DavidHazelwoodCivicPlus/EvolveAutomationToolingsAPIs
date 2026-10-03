const fs = require('node:fs');
const request = require('supertest');
const {
    app,
    uploadConfig,
    resolveSheetName,
    buildCustomMapPayload,
    buildCustomTypeConfig,
    loadCustomTypes,
    saveCustomTypes,
    registerCustomTypes,
    processUpload,
    CUSTOM_TYPES_FILE
} = require('../../index.js');
const { buildFetchMock } = require('../fixtures/schemaResponses.js');
const { buildWorkbook, workbookToBuffer, buildSingleSheetWorkbook } = require('../fixtures/workbook.js');

const originalFetch = global.fetch;
let warnSpy, logSpy, errorSpy;

// The custom-types file is real, on-disk state shared by every test in
// this file. Snapshot whatever is there before touching it and restore it
// afterwards, so running the suite never destroys a real customTypes.json
// sitting next to index.js.
let originalFileContents = null;
let fileExistedBefore = false;

beforeAll(() => {
    fileExistedBefore = fs.existsSync(CUSTOM_TYPES_FILE);
    if (fileExistedBefore) originalFileContents = fs.readFileSync(CUSTOM_TYPES_FILE, 'utf8');
});

afterAll(() => {
    if (fileExistedBefore) fs.writeFileSync(CUSTOM_TYPES_FILE, originalFileContents, 'utf8');
    else if (fs.existsSync(CUSTOM_TYPES_FILE)) fs.unlinkSync(CUSTOM_TYPES_FILE);
    global.fetch = originalFetch;
});

beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    logSpy = jest.spyOn(console, 'log').mockImplementation(() => {});
    errorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
    if (fs.existsSync(CUSTOM_TYPES_FILE)) fs.unlinkSync(CUSTOM_TYPES_FILE);
});

afterEach(() => {
    warnSpy.mockRestore();
    logSpy.mockRestore();
    errorSpy.mockRestore();
    // Any custom type a test registered would otherwise leak into every
    // later test's view of uploadConfig.
    for (const key of Object.keys(uploadConfig)) {
        if (uploadConfig[key].isCustom) delete uploadConfig[key];
    }
});

let siteCounter = 0;
const uniqueSite = () => `custom-types-test-${siteCounter++}.example.com`;

describe('resolveSheetName', () => {
    const sheetNames = ['Articles', 'Resources', 'FAQs', 'Staff', 'Departments', 'QuickLinks', 'News', 'Calendar', 'Agendas and Minutes', 'Facility'];

    test('matches by the configured sheetName, not by position', () => {
        expect(resolveSheetName(sheetNames, uploadConfig.facility)).toBe('Facility');
    });

    test('matching is case- and punctuation-insensitive', () => {
        expect(resolveSheetName(['FACILITY '], uploadConfig.facility)).toBe('FACILITY ');
        expect(resolveSheetName(['quick-links'], uploadConfig.quicklinks)).toBe('quick-links');
    });

    test('an explicit override beats the configured sheetName', () => {
        expect(resolveSheetName(sheetNames, uploadConfig.facility, 'Calendar')).toBe('Calendar');
    });

    test('an override that matches no sheet falls through to the name match rather than failing', () => {
        expect(resolveSheetName(sheetNames, uploadConfig.facility, 'Nonexistent')).toBe('Facility');
    });

    test('falls back to sheetIndex only when no name matches', () => {
        // A workbook with renamed sheets: facility's sheetIndex is 9.
        const renamed = ['s0', 's1', 's2', 's3', 's4', 's5', 's6', 's7', 's8', 's9'];
        expect(resolveSheetName(renamed, uploadConfig.facility)).toBe('s9');
    });

    test('returns undefined when neither name nor index matches', () => {
        expect(resolveSheetName(['OnlySheet'], uploadConfig.facility)).toBeUndefined();
    });

    test('a sheet order that would have caused the old facility bug now resolves correctly', () => {
        // The real workbook order, where index 8 is Agendas and Minutes.
        expect(resolveSheetName(sheetNames, uploadConfig.facility)).not.toBe('Agendas and Minutes');
    });

    test('a custom type (sheetIndex -1) matches by name only, never by position', () => {
        const config = buildCustomTypeConfig({ label: 'Agendas', endpoint: 'gh-agenda', sheetName: 'Agendas and Minutes', columns: ['A'] });
        expect(resolveSheetName(sheetNames, config)).toBe('Agendas and Minutes');
        expect(resolveSheetName(['Nothing', 'Matching'], config)).toBeUndefined();
    });
});

describe('processUpload honours the sheet override', () => {
    test('reads the overridden sheet instead of the name-matched one', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['faq'] });
        const workbook = buildWorkbook({
            FAQs: [{ Question: 'FromFAQs', Answer: 'A' }],
            OtherSheet: [{ Question: 'FromOther', Answer: 'B' }]
        });

        await processUpload('faqs', site, 'token', workbook, 10, 'OtherSheet');

        const createCall = global.fetch.mock.calls.find(([url, opts]) =>
            (opts?.method || '').toUpperCase() === 'POST' && /\/api\/content\//.test(String(url))
        );
        expect(JSON.parse(createCall[1].body).data.question).toEqual({ en: 'FromOther' });
    });
});

describe('buildCustomMapPayload', () => {
    const lookups = { permissionSet: new Map([['general', 'ps-1']]), categories: new Map([['fire', 'cat-1']]) };

    test('every data column becomes an invariant field', () => {
        const mapPayload = buildCustomMapPayload(['AgendaTitle', 'MeetingDate']);
        const result = mapPayload({ AgendaTitle: 'Jan Meeting', MeetingDate: '2026-01-05' }, lookups);
        expect(result.data).toEqual({
            AgendaTitle: { iv: 'Jan Meeting' },
            MeetingDate: { iv: '2026-01-05' }
        });
    });

    test('blank and missing cells are skipped, not sent as empty strings', () => {
        const mapPayload = buildCustomMapPayload(['A', 'B', 'C']);
        const result = mapPayload({ A: 'value', B: '' }, lookups);
        expect(result.data).toEqual({ A: { iv: 'value' } });
    });

    test('shared columns (PermissionSet, Categories, Publish, Tags, Name) are handled like every other type, not injected as data', () => {
        const mapPayload = buildCustomMapPayload(['Title', 'PermissionSet', 'Categories', 'Publish', 'Tags', 'Name']);
        const result = mapPayload({ Title: 'T', PermissionSet: 'General', Categories: 'Fire', Publish: 'Yes', Tags: 'a,b' }, lookups);

        expect(result.data).toEqual({ Title: { iv: 'T' } });
        expect(result.permissionSet).toEqual({ id: 'ps-1', name: 'General' });
        expect(result.categories).toEqual([{ id: 'cat-1', name: 'Fire' }]);
        expect(result.publish).toBe('Yes');
        expect(result.tags).toEqual(['a', 'b']);
    });

    test('tolerates an empty row without throwing', () => {
        const mapPayload = buildCustomMapPayload(['A']);
        expect(() => mapPayload({}, lookups)).not.toThrow();
    });

    test('an empty column list discovers columns from the row itself', () => {
        const mapPayload = buildCustomMapPayload([]);
        const result = mapPayload({ Whatever: 'x', Another: 'y', PermissionSet: 'General' }, lookups);
        expect(result.data).toEqual({ Whatever: { iv: 'x' }, Another: { iv: 'y' } });
        expect(result.permissionSet).toEqual({ id: 'ps-1', name: 'General' });
    });

    test('an undefined column list behaves the same as an empty one', () => {
        const mapPayload = buildCustomMapPayload(undefined);
        expect(mapPayload({ A: '1' }, lookups).data).toEqual({ A: { iv: '1' } });
    });

    test('a pinned column list ignores a column the sheet gained later', () => {
        const mapPayload = buildCustomMapPayload(['Known']);
        const result = mapPayload({ Known: 'yes', SneakedIn: 'should not upload' }, lookups);
        expect(result.data).toEqual({ Known: { iv: 'yes' } });
    });
});

describe('buildCustomTypeConfig', () => {
    test('produces a config shaped like a built-in one, matched by name only', () => {
        const config = buildCustomTypeConfig({ label: 'Agendas', endpoint: 'gh-agenda', sheetName: 'Agendas and Minutes', columns: ['AgendaTitle'] });
        expect(config.sheetIndex).toBe(-1);
        expect(config.sheetName).toBe('Agendas and Minutes');
        expect(config.endpoint).toBe('gh-agenda');
        expect(config.isCustom).toBe(true);
        expect(typeof config.mapPayload).toBe('function');
    });

    test('gets a default description when none is given', () => {
        const config = buildCustomTypeConfig({ label: 'Agendas', endpoint: 'gh-agenda', sheetName: 'Agendas and Minutes', columns: ['A'] });
        expect(config.description).toMatch(/Agendas and Minutes/);
    });
});

describe('persistence', () => {
    test('saveCustomTypes then loadCustomTypes round-trips a definition', () => {
        const definition = { agendas: { label: 'Agendas', endpoint: 'gh-agenda', sheetName: 'Agendas and Minutes', columns: ['A'] } };
        saveCustomTypes(definition);
        expect(loadCustomTypes()).toEqual(definition);
    });

    test('loadCustomTypes returns an empty object when the file does not exist', () => {
        expect(loadCustomTypes()).toEqual({});
    });

    test('a corrupt file does not throw — it just yields no custom types', () => {
        fs.writeFileSync(CUSTOM_TYPES_FILE, 'this is not json {{{', 'utf8');
        expect(loadCustomTypes()).toEqual({});
        expect(warnSpy).toHaveBeenCalled();
    });

    test('a JSON file that is an array (not an object) is ignored rather than mangling uploadConfig', () => {
        fs.writeFileSync(CUSTOM_TYPES_FILE, '["not", "an", "object"]', 'utf8');
        expect(loadCustomTypes()).toEqual({});
    });

    test('registerCustomTypes makes a saved type usable through uploadConfig', () => {
        registerCustomTypes({ custom_agendas: { label: 'Agendas', endpoint: 'gh-agenda', sheetName: 'Agendas and Minutes', columns: ['AgendaTitle'] } });
        expect(uploadConfig.custom_agendas).toBeDefined();
        expect(uploadConfig.custom_agendas.endpoint).toBe('gh-agenda');
    });
});

describe('GET /content-types', () => {
    test('lists all built-in types with labels and descriptions', async () => {
        const res = await request(app).get('/content-types');
        expect(res.status).toBe(200);
        expect(res.body.types.length).toBeGreaterThanOrEqual(9);
        const facility = res.body.types.find(t => t.type === 'facility');
        expect(facility).toMatchObject({ label: 'Facilities', sheetName: 'Facility', isCustom: false });
    });

    test('includes a registered custom type, flagged as custom', async () => {
        registerCustomTypes({ custom_agendas: { label: 'Agendas', endpoint: 'gh-agenda', sheetName: 'Agendas and Minutes', columns: ['A'] } });
        const res = await request(app).get('/content-types');
        const agendas = res.body.types.find(t => t.type === 'custom_agendas');
        expect(agendas).toMatchObject({ label: 'Agendas', isCustom: true });
    });
});

describe('POST /inspect-workbook', () => {
    test('returns each sheet with its columns, row count, and matched type', async () => {
        const buf = workbookToBuffer(buildWorkbook({
            Articles: [{ Title: 'T', Content: 'C' }],
            FAQs: [{ Question: 'Q', Answer: 'A' }, { Question: 'Q2', Answer: 'A2' }]
        }));

        const res = await request(app).post('/inspect-workbook').attach('file', buf, 'test.xlsx');

        expect(res.status).toBe(200);
        const faqs = res.body.sheets.find(s => s.name === 'FAQs');
        expect(faqs.columns).toEqual(['Question', 'Answer']);
        expect(faqs.rowCount).toBe(2);
        expect(faqs.matchedType).toBe('faqs');
        expect(faqs.unmatched).toBe(false);
    });

    test('flags a sheet that matches no known content type', async () => {
        const buf = workbookToBuffer(buildWorkbook({
            Articles: [{ Title: 'T' }],
            'Agendas and Minutes': [{ AgendaTitle: 'Jan', MeetingDate: '2026-01-05' }]
        }));

        const res = await request(app).post('/inspect-workbook').attach('file', buf, 'test.xlsx');

        const agendas = res.body.sheets.find(s => s.name === 'Agendas and Minutes');
        expect(agendas.unmatched).toBe(true);
        expect(agendas.matchedType).toBeNull();
        expect(agendas.columns).toEqual(['AgendaTitle', 'MeetingDate']);
    });

    test('reports columns even when every cell beneath them is blank', async () => {
        // sheet_to_json alone would drop these entirely — the route reads
        // the header row separately for exactly this case.
        const buf = workbookToBuffer(buildWorkbook({ Mystery: [{ ColA: '', ColB: '' }] }));
        const res = await request(app).post('/inspect-workbook').attach('file', buf, 'test.xlsx');
        const mystery = res.body.sheets.find(s => s.name === 'Mystery');
        expect(mystery.columns).toEqual(['ColA', 'ColB']);
    });

    test('no file returns 400', async () => {
        const res = await request(app).post('/inspect-workbook');
        expect(res.status).toBe(400);
    });

    test('does not upload anything — inspection is read-only', async () => {
        global.fetch = jest.fn();
        const buf = workbookToBuffer(buildWorkbook({ Articles: [{ Title: 'T' }] }));
        await request(app).post('/inspect-workbook').attach('file', buf, 'test.xlsx');
        expect(global.fetch).not.toHaveBeenCalled();
    });
});

describe('POST /custom-types', () => {
    test('saves a definition and makes it immediately usable', async () => {
        const res = await request(app).post('/custom-types').send({
            label: 'Agendas',
            endpoint: 'gh-agenda',
            sheetName: 'Agendas and Minutes',
            columns: ['AgendaTitle', 'MeetingDate']
        });

        expect(res.status).toBe(200);
        expect(res.body.type).toBe('custom_agendas');
        expect(uploadConfig.custom_agendas).toBeDefined();
        expect(loadCustomTypes().custom_agendas.endpoint).toBe('gh-agenda');
    });

    test('a saved custom type can actually upload, end to end', async () => {
        await request(app).post('/custom-types').send({
            label: 'Agendas', endpoint: 'gh-agenda', sheetName: 'Agendas and Minutes', columns: ['AgendaTitle']
        });

        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['gh-agenda'] });
        const workbook = buildWorkbook({ 'Agendas and Minutes': [{ AgendaTitle: 'January meeting' }] });

        const { successes, failures } = await processUpload('custom_agendas', site, 'token', workbook, 10);

        expect(failures).toEqual([]);
        expect(successes).toHaveLength(1);
        const createCall = global.fetch.mock.calls.find(([url, opts]) =>
            (opts?.method || '').toUpperCase() === 'POST' && /\/api\/content\//.test(String(url))
        );
        expect(new URL(String(createCall[0])).pathname).toBe(`/api/content/${site}/gh-agenda`);
        expect(JSON.parse(createCall[1].body).data.AgendaTitle).toEqual({ iv: 'January meeting' });
    });

    test('missing required fields return 400', async () => {
        expect((await request(app).post('/custom-types').send({ endpoint: 'x', sheetName: 'y', columns: ['A'] })).status).toBe(400);
        expect((await request(app).post('/custom-types').send({ label: 'x', sheetName: 'y', columns: ['A'] })).status).toBe(400);
        expect((await request(app).post('/custom-types').send({ label: 'x', endpoint: 'y', columns: ['A'] })).status).toBe(400);
    });

    test('columns are optional — a hand-registered type saves without them', async () => {
        const res = await request(app).post('/custom-types').send({ label: 'Agendas', endpoint: 'gh-agenda', sheetName: 'Agendas and Minutes' });
        expect(res.status).toBe(200);
        expect(loadCustomTypes().custom_agendas.columns).toEqual([]);
    });

    test('a non-array columns value is rejected', async () => {
        const res = await request(app).post('/custom-types').send({ label: 'x', endpoint: 'y', sheetName: 'z', columns: 'not-a-list' });
        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/list/);
    });

    test('a type registered without columns discovers them from the sheet at upload time', async () => {
        await request(app).post('/custom-types').send({ label: 'Agendas', endpoint: 'gh-agenda', sheetName: 'Agendas and Minutes' });

        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['gh-agenda'] });
        const workbook = buildWorkbook({
            'Agendas and Minutes': [{ AgendaTitle: 'January', MeetingDate: '2026-01-05', PermissionSet: 'General' }]
        });

        const { successes, failures } = await processUpload('custom_agendas', site, 'token', workbook, 10);

        expect(failures).toEqual([]);
        expect(successes).toHaveLength(1);
        const createCall = global.fetch.mock.calls.find(([url, opts]) =>
            (opts?.method || '').toUpperCase() === 'POST' && /\/api\/content\//.test(String(url))
        );
        const body = JSON.parse(createCall[1].body);
        // Discovered from the row itself, with PermissionSet handled
        // separately rather than sent as a data field.
        expect(body.data).toEqual({
            AgendaTitle: { iv: 'January' },
            MeetingDate: { iv: '2026-01-05' }
        });
    });

    test('a label with no letters or digits returns 400 rather than creating an unusable key', async () => {
        const res = await request(app).post('/custom-types').send({ label: '!!!', endpoint: 'y', sheetName: 'z', columns: ['A'] });
        expect(res.status).toBe(400);
    });

    test('a label that would collide with a built-in type is rejected', async () => {
        // uploadConfig has no "custom_..." built-ins, so simulate the
        // collision guard directly by adding a non-custom entry.
        uploadConfig.custom_collide = { sheetIndex: 0, endpoint: 'x', mapPayload: () => ({ data: {} }) };
        try {
            const res = await request(app).post('/custom-types').send({ label: 'collide', endpoint: 'y', sheetName: 'z', columns: ['A'] });
            expect(res.status).toBe(400);
            expect(res.body.message).toMatch(/built-in/);
        } finally {
            delete uploadConfig.custom_collide;
        }
    });

    test('saving the same label twice updates rather than duplicating', async () => {
        await request(app).post('/custom-types').send({ label: 'Agendas', endpoint: 'first-slug', sheetName: 'S', columns: ['A'] });
        await request(app).post('/custom-types').send({ label: 'Agendas', endpoint: 'second-slug', sheetName: 'S', columns: ['A'] });

        expect(Object.keys(loadCustomTypes())).toEqual(['custom_agendas']);
        expect(uploadConfig.custom_agendas.endpoint).toBe('second-slug');
    });
});

describe('DELETE /custom-types/:key', () => {
    test('removes a saved type from both disk and uploadConfig', async () => {
        await request(app).post('/custom-types').send({ label: 'Agendas', endpoint: 'gh-agenda', sheetName: 'S', columns: ['A'] });

        const res = await request(app).delete('/custom-types/custom_agendas');

        expect(res.status).toBe(200);
        expect(uploadConfig.custom_agendas).toBeUndefined();
        expect(loadCustomTypes().custom_agendas).toBeUndefined();
    });

    test('an unknown key returns 404', async () => {
        const res = await request(app).delete('/custom-types/custom_nothing');
        expect(res.status).toBe(404);
    });
});
