// Bulk asset upload: the Assets sheet, per-file settings, alt text, retries
// and the downloadable sheet template. Nothing here talks to a real site.

const request = require('supertest');
const XLSX = require('xlsx');
const {
    app,
    postAssetFile,
    ASSET_RETRY,
    parsePublishFlag,
    parseAssetRows,
    buildAssetTemplateWorkbook,
    assetLookupCache
} = require('../../index.js');
const { buildFetchMock, jsonResponse, textResponse } = require('../fixtures/schemaResponses.js');
const { buildWorkbook, workbookToBuffer } = require('../fixtures/workbook.js');

const originalFetch = global.fetch;
const originalRetry = { ...ASSET_RETRY };
let spies = [];

beforeEach(() => {
    spies = ['warn', 'log', 'error'].map(m => jest.spyOn(console, m).mockImplementation(() => {}));
    ASSET_RETRY.baseDelayMs = 1;
});
afterEach(() => {
    spies.forEach(s => s.mockRestore());
    Object.assign(ASSET_RETRY, originalRetry);
});
afterAll(() => {
    global.fetch = originalFetch;
});

let siteCounter = 0;
const uniqueSite = () => `assets-test-${siteCounter++}.example.com`;

const assetPosts = () => global.fetch.mock.calls.filter(([url, opts]) =>
    (opts.method || '').toUpperCase() === 'POST' && /\/assets\?/.test(String(url)));
const assetPuts = () => global.fetch.mock.calls.filter(([url, opts]) =>
    (opts.method || '').toUpperCase() === 'PUT' && /\/assets\/[^/?]+$/.test(String(url)));

describe('parsePublishFlag', () => {
    test.each([
        ['Yes', true], ['TRUE', true], ['1', true], [true, true],
        ['No', false], ['false', false], ['0', false], [false, false],
        ['', undefined], [undefined, undefined], ['maybe', undefined]
    ])('%p -> %p', (input, expected) => {
        expect(parsePublishFlag(input)).toBe(expected);
    });
});

describe('parseAssetRows', () => {
    test('reads the Assets sheet with loose header names and drops blanks', () => {
        const wb = buildWorkbook({
            Articles: [{ Title: 'x' }],
            assets: [
                { 'File Name': 'a.jpg', 'Alt Text': 'A dog', PermissionSet: 'General', Categories: 'Parks', Tags: 't1, t2', Publish: 'Yes', 'Parent Folder Id': '' },
                { 'File Name': 'b.pdf', 'Alt Text': '', Publish: '' },
                { 'File Name': '', 'Alt Text': 'orphan row' }
            ]
        });
        const parsed = parseAssetRows(wb);
        expect(parsed.sheetName).toBe('assets');
        expect(parsed.rows).toEqual([
            { fileName: 'a.jpg', altText: 'A dog', permissionSet: 'General', categories: 'Parks', tags: 't1, t2', publish: true },
            { fileName: 'b.pdf' }
        ]);
    });

    test('returns null when the workbook has no Assets sheet', () => {
        expect(parseAssetRows(buildWorkbook({ Articles: [{ Title: 'x' }] }))).toBeNull();
    });
});

describe('buildAssetTemplateWorkbook', () => {
    test('one row per file, pre-filled with the defaults, in the column order the parser reads', () => {
        const wb = buildAssetTemplateWorkbook(['a.jpg', 'sub/b.pdf'], { permissionSet: 'General', altText: 'Photo', publish: true });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets.Assets, { header: 1 });
        expect(rows[0]).toEqual(['FileName', 'ParentFolderId', 'PermissionSet', 'Categories', 'AltText', 'Tags', 'Publish']);
        expect(parseAssetRows(wb).rows).toEqual([
            { fileName: 'a.jpg', permissionSet: 'General', altText: 'Photo', publish: true },
            { fileName: 'sub/b.pdf', permissionSet: 'General', altText: 'Photo', publish: true }
        ]);
    });
});

describe('postAssetFile alt text, tags and retries', () => {
    test('sends altText and one tags entry per tag in the multipart form', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ id: 'a1' }));
        await postAssetFile({ siteURL: 's', accessToken: 't', fileBuffer: Buffer.from('x'), fileName: 'x.jpg', altText: 'A cat', tags: ['one', 'two'] });
        const form = global.fetch.mock.calls[0][1].body;
        expect(form.get('altText')).toBe('A cat');
        expect(form.getAll('tags')).toEqual(['one', 'two']);
    });

    test('retries a 429 / 503 and succeeds, with a fresh form each attempt', async () => {
        global.fetch = jest.fn()
            .mockResolvedValueOnce(textResponse('slow down', 429))
            .mockResolvedValueOnce(textResponse('unavailable', 503))
            .mockResolvedValueOnce(jsonResponse({ id: 'a2' }));
        const result = await postAssetFile({ siteURL: 's', accessToken: 't', fileBuffer: Buffer.from('x'), fileName: 'x.jpg' });
        expect(result).toEqual({ id: 'a2' });
        expect(global.fetch).toHaveBeenCalledTimes(3);
        const bodies = global.fetch.mock.calls.map(c => c[1].body);
        expect(new Set(bodies).size).toBe(3);
    });

    test('retries a network error', async () => {
        global.fetch = jest.fn()
            .mockRejectedValueOnce(new Error('ECONNRESET'))
            .mockResolvedValueOnce(jsonResponse({ id: 'a3' }));
        await expect(postAssetFile({ siteURL: 's', accessToken: 't', fileBuffer: Buffer.from('x'), fileName: 'x.jpg' }))
            .resolves.toEqual({ id: 'a3' });
    });

    test('does not retry a 400 — it would only fail the same way', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('bad file', 400));
        await expect(postAssetFile({ siteURL: 's', accessToken: 't', fileBuffer: Buffer.from('x'), fileName: 'x.jpg' }))
            .rejects.toThrow(/400/);
        expect(global.fetch).toHaveBeenCalledTimes(1);
    });

    test('a call that never answers times out and is retried instead of hanging the batch', async () => {
        ASSET_RETRY.timeoutMs = 20;
        global.fetch = jest.fn()
            .mockImplementationOnce((url, init) => new Promise((resolve, reject) => {
                init.signal.addEventListener('abort', () => reject(init.signal.reason));
            }))
            .mockResolvedValueOnce(jsonResponse({ id: 'a4' }));
        await expect(postAssetFile({ siteURL: 's', accessToken: 't', fileBuffer: Buffer.from('x'), fileName: 'x.jpg' }))
            .resolves.toEqual({ id: 'a4' });
        expect(global.fetch).toHaveBeenCalledTimes(2);
    });

    test('a call that keeps timing out fails with a readable message', async () => {
        ASSET_RETRY.timeoutMs = 10;
        global.fetch = jest.fn((url, init) => new Promise((resolve, reject) => {
            init.signal.addEventListener('abort', () => reject(init.signal.reason));
        }));
        await expect(postAssetFile({ siteURL: 's', accessToken: 't', fileBuffer: Buffer.from('x'), fileName: 'x.jpg' }))
            .rejects.toThrow(/No response from the site/);
    });

    test('a huge Retry-After is capped rather than stalling the run', async () => {
        ASSET_RETRY.maxDelayMs = 5;
        global.fetch = jest.fn()
            .mockResolvedValueOnce(new Response('slow down', { status: 429, headers: { 'Retry-After': '3600' } }))
            .mockResolvedValueOnce(jsonResponse({ id: 'a5' }));
        const started = Date.now();
        await postAssetFile({ siteURL: 's', accessToken: 't', fileBuffer: Buffer.from('x'), fileName: 'x.jpg' });
        expect(Date.now() - started).toBeLessThan(1000);
    });

    test('duplicate: true asks the site for a new copy even of a file it already has', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ id: 'a6' }));
        await postAssetFile({ siteURL: 's', accessToken: 't', fileBuffer: Buffer.from('x'), fileName: 'x.jpg', duplicate: true });
        expect(new URL(String(global.fetch.mock.calls[0][0])).searchParams.get('duplicate')).toBe('true');
    });

    test('gives up after the configured number of attempts', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('busy', 503));
        await expect(postAssetFile({ siteURL: 's', accessToken: 't', fileBuffer: Buffer.from('x'), fileName: 'x.jpg' }))
            .rejects.toThrow(/503/);
        expect(global.fetch).toHaveBeenCalledTimes(ASSET_RETRY.attempts);
    });
});

describe('POST /upload/assets with per-file settings', () => {
    const referenceData = {
        permissionSet: [{ id: 'ps-general', name: 'General' }, { id: 'ps-staff', name: 'Staff Only' }],
        categories: [{ id: 'cat-parks', name: 'Parks' }, { id: 'cat-fire', name: 'Fire' }]
    };

    test('a fileSettings row overrides the form defaults for its file only; blanks keep the default', async () => {
        global.fetch = buildFetchMock({ referenceData });

        const res = await request(app)
            .post('/upload/assets')
            .field('url', uniqueSite())
            .field('apiKey', 'token')
            .field('permissionSet', 'General')
            .field('categories', 'Fire')
            .field('publish', 'false')
            .field('fileSettings', JSON.stringify([
                { permissionSet: 'Staff Only', categories: 'Parks', altText: 'Splash pad', publish: true, parentId: 'folder-1' },
                { permissionSet: '', altText: '' }
            ]))
            .attach('files', Buffer.from('one'), 'one.jpg')
            .attach('files', Buffer.from('two'), 'two.jpg');

        expect(res.status).toBe(200);
        expect(res.body.uploaded).toBe(2);

        const byName = Object.fromEntries(assetPosts().map(([url, opts]) => [opts.body.get('file').name, { url: new URL(String(url)), form: opts.body }]));

        expect(byName['one.jpg'].form.get('permissionSet')).toBe('ps-staff');
        expect(byName['one.jpg'].form.getAll('categories')).toEqual(['cat-parks']);
        expect(byName['one.jpg'].form.get('altText')).toBe('Splash pad');
        expect(byName['one.jpg'].url.searchParams.get('publish')).toBe('true');
        expect(byName['one.jpg'].url.searchParams.get('parentId')).toBe('folder-1');

        expect(byName['two.jpg'].form.get('permissionSet')).toBe('ps-general');
        expect(byName['two.jpg'].form.getAll('categories')).toEqual(['cat-fire']);
        expect(byName['two.jpg'].form.get('altText')).toBeNull();
        expect(byName['two.jpg'].url.searchParams.get('publish')).toBe('false');
    });

    test('alt text the create call did not echo back is set with a follow-up PUT', async () => {
        global.fetch = buildFetchMock({ referenceData, assetResponse: () => ({ id: 'asset-xyz' }) });

        const res = await request(app)
            .post('/upload/assets')
            .field('url', uniqueSite())
            .field('apiKey', 'token')
            .field('altText', 'Default alt')
            .field('tags', 'a, b')
            .attach('files', Buffer.from('one'), 'one.jpg');

        expect(res.status).toBe(200);
        const puts = assetPuts();
        expect(puts).toHaveLength(1);
        expect(String(puts[0][0])).toMatch(/\/assets\/asset-xyz$/);
        expect(JSON.parse(puts[0][1].body)).toEqual({ altText: 'Default alt', tags: ['a', 'b'] });
    });

    test('no follow-up PUT when the create response already carries the alt text', async () => {
        global.fetch = buildFetchMock({ referenceData, assetResponse: () => ({ id: 'asset-1', altText: 'Default alt' }) });

        await request(app)
            .post('/upload/assets')
            .field('url', uniqueSite())
            .field('apiKey', 'token')
            .field('altText', 'Default alt')
            .attach('files', Buffer.from('one'), 'one.jpg');

        expect(assetPuts()).toHaveLength(0);
    });

    test('a failed alt-text follow-up is a warning, not a failed upload', async () => {
        global.fetch = buildFetchMock({ referenceData, annotateStatus: 400 });

        const res = await request(app)
            .post('/upload/assets')
            .field('url', uniqueSite())
            .field('apiKey', 'token')
            .field('altText', 'Alt')
            .attach('files', Buffer.from('one'), 'one.jpg');

        expect(res.status).toBe(200);
        expect(res.body.uploaded).toBe(1);
        expect(res.body.warnings.join(' ')).toMatch(/alt text/);
    });

    test('unknown category / permission set names come back as warnings', async () => {
        global.fetch = buildFetchMock({ referenceData });

        const res = await request(app)
            .post('/upload/assets')
            .field('url', uniqueSite())
            .field('apiKey', 'token')
            .field('permissionSet', 'Nope')
            .field('categories', 'Parks, Unknown Cat')
            .attach('files', Buffer.from('one'), 'one.jpg');

        expect(res.status).toBe(200);
        expect(res.body.warnings).toEqual([
            'one.jpg: permission set "Nope" not found on the site',
            'one.jpg: category "Unknown Cat" not found on the site'
        ]);
    });

    test('a partial failure lists exactly which files failed, so the dashboard can retry them', async () => {
        const base = buildFetchMock({ referenceData });
        global.fetch = jest.fn(async (url, opts) => {
            if ((opts.method || '').toUpperCase() === 'POST' && /\/assets\?/.test(String(url)) && opts.body.get('file').name === 'bad.jpg') {
                return textResponse('nope', 400);
            }
            return base(url, opts);
        });

        const res = await request(app)
            .post('/upload/assets')
            .field('url', uniqueSite())
            .field('apiKey', 'token')
            .attach('files', Buffer.from('1'), 'good.jpg')
            .attach('files', Buffer.from('2'), 'bad.jpg');

        expect(res.status).toBe(207);
        expect(res.body.uploaded).toBe(1);
        expect(res.body.failedFiles).toEqual(['bad.jpg']);
    });

    test('files the site already had are counted as duplicates, not as uploaded', async () => {
        let n = 0;
        global.fetch = buildFetchMock({
            referenceData,
            assetResponse: () => (n++ === 0 ? { id: 'existing', _meta: { isDuplicate: 'true' } } : { id: 'new', _meta: { isDuplicate: 'false' } })
        });

        const res = await request(app)
            .post('/upload/assets')
            .field('url', uniqueSite())
            .field('apiKey', 'token')
            .attach('files', Buffer.from('1'), 'old.jpg')
            .attach('files', Buffer.from('2'), 'new.jpg');

        expect(res.status).toBe(200);
        expect(res.body.uploaded).toBe(1);
        expect(res.body.duplicates).toBe(1);
        expect(res.body.duplicateFiles).toHaveLength(1);
        expect(res.body.message).toMatch(/1 already existed/);
    });

    test('allowDuplicates is passed through to the site', async () => {
        global.fetch = buildFetchMock({ referenceData });
        await request(app)
            .post('/upload/assets')
            .field('url', uniqueSite())
            .field('apiKey', 'token')
            .field('allowDuplicates', 'true')
            .attach('files', Buffer.from('1'), 'a.jpg');
        expect(new URL(String(assetPosts()[0][0])).searchParams.get('duplicate')).toBe('true');
    });

    test('malformed fileSettings is rejected with a 400', async () => {
        const res = await request(app)
            .post('/upload/assets')
            .field('url', 'site')
            .field('apiKey', 'token')
            .field('fileSettings', '{not json')
            .attach('files', Buffer.from('1'), 'a.jpg');
        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/fileSettings/);
    });

    test('lookups are fetched once per site across batches, not once per request', async () => {
        assetLookupCache.clear();
        const site = uniqueSite();
        global.fetch = buildFetchMock({ referenceData });

        for (let i = 0; i < 3; i++) {
            await request(app).post('/upload/assets').field('url', site).field('apiKey', 'token')
                .attach('files', Buffer.from('x'), `f${i}.jpg`);
        }

        const lookupCalls = global.fetch.mock.calls.filter(([url, opts]) =>
            (opts.method || 'GET').toUpperCase() === 'GET' && /\/(permissionSet|categories)$/.test(new URL(String(url)).pathname));
        expect(lookupCalls).toHaveLength(2);
    });
});

describe('POST /inspect-assets-sheet', () => {
    test('returns the parsed rows when the workbook has an Assets sheet', async () => {
        const buffer = workbookToBuffer(buildWorkbook({ Assets: [{ FileName: 'a.jpg', AltText: 'Alt', Publish: 'yes' }] }));
        const res = await request(app).post('/inspect-assets-sheet').attach('file', buffer, 'book.xlsx');
        expect(res.status).toBe(200);
        expect(res.body).toEqual({ found: true, sheetName: 'Assets', rows: [{ fileName: 'a.jpg', altText: 'Alt', publish: true }] });
    });

    test('reports found: false when there is no Assets sheet', async () => {
        const buffer = workbookToBuffer(buildWorkbook({ Articles: [{ Title: 'x' }] }));
        const res = await request(app).post('/inspect-assets-sheet').attach('file', buffer, 'book.xlsx');
        expect(res.body).toEqual({ found: false, rows: [] });
    });
});

describe('POST /inspect-workbook', () => {
    test('does not list the Assets sheet as an unrecognized content sheet', async () => {
        const buffer = workbookToBuffer(buildWorkbook({ Articles: [{ Title: 'x' }], Assets: [{ FileName: 'a.jpg' }] }));
        const res = await request(app).post('/inspect-workbook').attach('file', buffer, 'book.xlsx');
        const assets = res.body.sheets.find(s => s.name === 'Assets');
        expect(assets.unmatched).toBe(false);
    });
});

describe('POST /assets/template', () => {
    test('returns an .xlsx with one row per file name', async () => {
        const res = await request(app)
            .post('/assets/template')
            .send({ fileNames: ['a.jpg', 'b.pdf'], defaults: { permissionSet: 'General' } })
            .buffer(true)
            .parse((r, cb) => { const chunks = []; r.on('data', c => chunks.push(c)); r.on('end', () => cb(null, Buffer.concat(chunks))); });

        expect(res.status).toBe(200);
        expect(res.headers['content-disposition']).toMatch(/AssetsSheet\.xlsx/);
        const parsed = parseAssetRows(XLSX.read(res.body, { type: 'buffer' }));
        expect(parsed.rows.map(r => r.fileName)).toEqual(['a.jpg', 'b.pdf']);
        expect(parsed.rows[0].permissionSet).toBe('General');
    });

    test('400 with no file names', async () => {
        const res = await request(app).post('/assets/template').send({ fileNames: [] });
        expect(res.status).toBe(400);
    });
});
