const request = require('supertest');
const { app } = require('../../index.js');
const { buildFetchMock, jsonResponse, textResponse } = require('../fixtures/schemaResponses.js');
const { buildSingleSheetWorkbook, workbookToBuffer } = require('../fixtures/workbook.js');

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

let siteCounter = 0;
const uniqueSite = () => `routes-test-${siteCounter++}.example.com`;

describe('static routes', () => {
    test('GET / serves the dashboard HTML', async () => {
        const res = await request(app).get('/');
        expect(res.status).toBe(200);
        expect(res.text).toContain('Evolve Content Uploader');
    });

    test('GET /style.css serves the stylesheet', async () => {
        const res = await request(app).get('/style.css');
        expect(res.status).toBe(200);
    });

    test('GET /civicplus-mark.png serves the brand mark', async () => {
        const res = await request(app).get('/civicplus-mark.png');
        expect(res.status).toBe(200);
    });
});

describe('POST /upload/:type', () => {
    test('an unknown type returns 400', async () => {
        const buf = workbookToBuffer(buildSingleSheetWorkbook('Articles', [{ Title: 'T' }], 0));
        const res = await request(app)
            .post('/upload/not-a-real-type')
            .field('url', 'site.example.com')
            .field('apiKey', 'token')
            .attach('file', buf, 'test.xlsx');

        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/Invalid upload type/);
    });

    test('no file attached returns 400', async () => {
        const res = await request(app)
            .post('/upload/articles')
            .field('url', 'site.example.com')
            .field('apiKey', 'token');

        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/No Excel file/);
    });

    test('a disallowed file extension is rejected by multer\'s fileFilter with a 400', async () => {
        const res = await request(app)
            .post('/upload/articles')
            .field('url', 'site.example.com')
            .field('apiKey', 'token')
            .attach('file', Buffer.from('not a spreadsheet'), 'test.exe');

        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/xlsx/);
    });

    test('a full successful upload returns 200 with a processed count', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['article'] });
        const buf = workbookToBuffer(buildSingleSheetWorkbook(
            'Articles',
            [{ Title: 'T', Content: 'C', ContentType: 'article', PermissionSet: 'General' }],
            0
        ));

        const res = await request(app)
            .post('/upload/articles')
            .field('url', site)
            .field('apiKey', 'token')
            .attach('file', buf, 'test.xlsx');

        expect(res.status).toBe(200);
        expect(res.body.message).toMatch(/Processed 1 articles/);
    });

    test('partial failure returns 207 with an errors list capped at 10', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['faq'], createHandler: () => textResponse('nope', 422) });
        const rows = Array.from({ length: 12 }, (_, i) => ({ Question: `Q${i}`, Answer: `A${i}` }));
        const buf = workbookToBuffer(buildSingleSheetWorkbook('FAQs', rows, 2));

        const res = await request(app)
            .post('/upload/faqs')
            .field('url', site)
            .field('apiKey', 'token')
            .attach('file', buf, 'test.xlsx');

        expect(res.status).toBe(207);
        expect(res.body.errors).toHaveLength(10); // 12 rows failed, but the response caps the list at 10
        expect(res.body.message).toMatch(/12 failed/);
    });

    test('the response names categories that were skipped, so a silent data change is visible', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({
            schemaNames: ['faq'],
            referenceData: { categories: [{ id: 'cat-1', name: 'Real Category' }] }
        });
        const buf = workbookToBuffer(buildSingleSheetWorkbook(
            'FAQs', [{ Question: 'Q', Answer: 'A', Categories: 'Fire Marshal' }], 2
        ));

        const res = await request(app)
            .post('/upload/faqs')
            .field('url', site)
            .field('apiKey', 'token')
            .attach('file', buf, 'test.xlsx');

        expect(res.status).toBe(200); // the row still uploads
        expect(res.body.message).toMatch(/Fire Marshal/);
        expect(res.body.message).toMatch(/skipped/i);
    });

    test('an error outside the row loop (e.g. an unresolvable schema slug) returns 500', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({ schemaNames: ['totally-unrelated'] });
        const buf = workbookToBuffer(buildSingleSheetWorkbook('Staff', [{ FirstName: 'A' }], 3));

        const res = await request(app)
            .post('/upload/staff')
            .field('url', site)
            .field('apiKey', 'token')
            .attach('file', buf, 'test.xlsx');

        expect(res.status).toBe(500);
        expect(res.body.message).toMatch(/employee/);
    });
});

describe('POST /upload/assets', () => {
    test('no files selected returns 400', async () => {
        const res = await request(app).post('/upload/assets').field('url', 'site').field('apiKey', 'token');
        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/No files selected/);
    });

    test('missing url/apiKey returns 400', async () => {
        const res = await request(app).post('/upload/assets').attach('files', Buffer.from('hi'), 'a.txt');
        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/required/);
    });

    test('a full successful upload returns 200 with an uploaded count', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock();

        const res = await request(app)
            .post('/upload/assets')
            .field('url', site)
            .field('apiKey', 'token')
            .attach('files', Buffer.from('file one'), 'one.txt')
            .attach('files', Buffer.from('file two'), 'two.txt');

        expect(res.status).toBe(200);
        expect(res.body.message).toMatch(/Uploaded 2 file\(s\)/);
    });

    test('partial failure returns 207', async () => {
        const site = uniqueSite();
        let assetCallCount = 0;
        const base = buildFetchMock();
        global.fetch = jest.fn(async (url, opts) => {
            const u = new URL(String(url));
            if ((opts.method || '').toUpperCase() === 'POST' && /\/assets$/.test(u.pathname)) {
                assetCallCount++;
                return assetCallCount === 1 ? textResponse('server error', 500) : jsonResponse({ id: 'ok' });
            }
            return base(url, opts);
        });

        const res = await request(app)
            .post('/upload/assets')
            .field('url', site)
            .field('apiKey', 'token')
            .attach('files', Buffer.from('file one'), 'one.txt')
            .attach('files', Buffer.from('file two'), 'two.txt');

        expect(res.status).toBe(207);
        expect(res.body.errors).toHaveLength(1);
    });

    test('exceeding multer\'s file-count limit (200) surfaces as a 400, not a crash', async () => {
        let req = request(app).post('/upload/assets').field('url', 'site').field('apiKey', 'token');
        for (let i = 0; i < 201; i++) {
            req = req.attach('files', Buffer.from('x'), `f${i}.txt`);
        }
        const res = await req;
        expect(res.status).toBe(400);
    }, 15000);

    test('the upload still proceeds when reference-data lookups fail (categories/permissionSet just stay unresolved)', async () => {
        const site = uniqueSite();
        global.fetch = jest.fn(async (url, opts) => {
            const u = new URL(String(url));
            if ((opts.method || 'GET').toUpperCase() === 'GET' && /\/api\/apps\/[^/]+\/(permissionSet|categories)$/.test(u.pathname)) {
                return textResponse('server error', 500);
            }
            if ((opts.method || '').toUpperCase() === 'POST' && /\/assets$/.test(u.pathname)) {
                return jsonResponse({ id: 'ok' });
            }
            throw new Error('unexpected call: ' + url);
        });

        const res = await request(app)
            .post('/upload/assets')
            .field('url', site)
            .field('apiKey', 'token')
            .field('permissionSet', 'General')
            .attach('files', Buffer.from('file'), 'one.txt');

        expect(res.status).toBe(200); // the file itself still uploads fine
    });

    test('category/permissionSet names are resolved against reference data, with a raw-GUID fallback', async () => {
        const site = uniqueSite();
        global.fetch = buildFetchMock({
            referenceData: {
                permissionSet: [{ id: 'ps-1', name: 'General' }],
                categories: [{ id: 'cat-1', name: 'Fire' }]
            }
        });

        const res = await request(app)
            .post('/upload/assets')
            .field('url', site)
            .field('apiKey', 'token')
            .field('permissionSet', 'General')
            .field('categories', 'Fire,00000000-0000-0000-0000-000000000000')
            .attach('files', Buffer.from('file'), 'one.txt');

        expect(res.status).toBe(200);
        const assetCall = global.fetch.mock.calls.find(([url, opts]) =>
            (opts.method || '').toUpperCase() === 'POST' && /\/assets\?/.test(String(url))
        );
        expect(assetCall).toBeDefined();
        const form = assetCall[1].body;
        expect(form.get('permissionSet')).toBe('ps-1');
        expect(form.getAll('categories')).toEqual(['cat-1', '00000000-0000-0000-0000-000000000000']);
    });
});
