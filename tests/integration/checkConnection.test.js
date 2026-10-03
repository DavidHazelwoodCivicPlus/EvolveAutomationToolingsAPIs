const { app, uploadConfig, checkSiteConnection, checkConnection } = require('../../index.js');
const request = require('supertest');
const { jsonResponse, textResponse } = require('../fixtures/schemaResponses.js');

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

describe('checkSiteConnection', () => {
    test('a network failure (site name wrong, no internet, etc.) is reported distinctly', async () => {
        global.fetch = jest.fn().mockRejectedValue(new Error('getaddrinfo ENOTFOUND'));
        const result = await checkSiteConnection('bad-site', 'token');
        expect(result).toMatchObject({ ok: false, status: null, reason: 'network' });
        expect(result.message).toMatch(/Could not reach/);
    });

    test.each([401, 403])('a %i response is reported as an auth problem, not a generic error', async (status) => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('nope', status));
        const result = await checkSiteConnection('site', 'bad-token');
        expect(result).toMatchObject({ ok: false, status, reason: 'auth' });
        expect(result.message).toMatch(/Token was rejected/);
    });

    test('a 404 is reported as the site app name being wrong, not an auth problem', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('nope', 404));
        const result = await checkSiteConnection('typo-site', 'token');
        expect(result).toMatchObject({ ok: false, status: 404, reason: 'not-found' });
        expect(result.message).toMatch(/was not found/);
    });

    test('any other non-2xx status is still surfaced, with the raw status and body', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('upstream is down', 503));
        const result = await checkSiteConnection('site', 'token');
        expect(result.ok).toBe(false);
        expect(result.reason).toBe('error');
        expect(result.message).toMatch(/503/);
        expect(result.message).toMatch(/upstream is down/);
    });

    test('success extracts schema names from a plain array response', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse(['article', 'faq']));
        const result = await checkSiteConnection('site', 'token');
        expect(result).toMatchObject({ ok: true, status: 200 });
        expect(result.schemaNames).toEqual(['article', 'faq']);
    });

    test('success extracts schema names from a {items:[{name}]} response, and 0 schemas still counts as ok', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ items: [] }));
        const result = await checkSiteConnection('site', 'token');
        expect(result.ok).toBe(true);
        expect(result.schemaNames).toEqual([]);
        expect(result.message).toMatch(/found 0 schema/);
    });

    test('never caches — two calls in a row always hit fetch twice', async () => {
        global.fetch = jest.fn().mockImplementation(async () => jsonResponse(['article']));
        await checkSiteConnection('same-site', 'token');
        await checkSiteConnection('same-site', 'token');
        expect(global.fetch).toHaveBeenCalledTimes(2);
    });
});

describe('checkConnection', () => {
    test('when the connection itself fails, contentTypes and referenceData are left empty rather than attempted', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('nope', 401));
        const report = await checkConnection('site', 'bad-token');
        expect(report.connection.ok).toBe(false);
        expect(report.contentTypes).toEqual([]);
        expect(report.referenceData).toEqual([]);
    });

    test('when every real content type\'s schema exists, every entry reports ok, with a friendly label', async () => {
        // Every uploadConfig type's PREFERRED candidate, so every type should resolve.
        const allSchemas = Object.values(uploadConfig).map(cfg => (cfg.endpoints || [cfg.endpoint])[0]);
        global.fetch = jest.fn(async (url) => {
            if (String(url).endsWith('/schemas')) return jsonResponse(allSchemas);
            return jsonResponse({ items: [] }); // permissionSet / categories
        });

        const report = await checkConnection('site', 'token');

        expect(report.contentTypes).toHaveLength(Object.keys(uploadConfig).length);
        expect(report.contentTypes.every(ct => ct.ok)).toBe(true);
        const staffEntry = report.contentTypes.find(ct => ct.type === 'staff');
        expect(staffEntry.label).toBe('Staff');
    });

    test('a content type whose schema is missing on this site is reported as not ok, without failing the others', async () => {
        // Everything except staff's candidates ("employee"/"enhanced-employee").
        const partialSchemas = Object.entries(uploadConfig)
            .filter(([type]) => type !== 'staff')
            .map(([, cfg]) => (cfg.endpoints || [cfg.endpoint])[0]);
        global.fetch = jest.fn(async (url) => {
            if (String(url).endsWith('/schemas')) return jsonResponse(partialSchemas);
            return jsonResponse({ items: [] });
        });

        const report = await checkConnection('site', 'token');

        const staffEntry = report.contentTypes.find(ct => ct.type === 'staff');
        const articlesEntry = report.contentTypes.find(ct => ct.type === 'articles');
        expect(staffEntry.ok).toBe(false);
        expect(staffEntry.resolvedEndpoint).toBeNull();
        expect(articlesEntry.ok).toBe(true);
    });

    test('reference data reports a count on success, and a message (not a thrown error) on failure', async () => {
        global.fetch = jest.fn(async (url) => {
            const u = new URL(String(url));
            if (u.pathname.endsWith('/schemas')) return jsonResponse(['article']);
            if (u.pathname.endsWith('/permissionSet')) return jsonResponse({ items: [{ id: '1', name: 'General' }, { id: '2', name: 'Admin' }] });
            if (u.pathname.endsWith('/categories')) return textResponse('server error', 500);
            throw new Error('unexpected call: ' + url);
        });

        const report = await checkConnection('site', 'token');

        const permissionSet = report.referenceData.find(r => r.endpoint === 'permissionSet');
        const categories = report.referenceData.find(r => r.endpoint === 'categories');
        expect(permissionSet).toEqual({ endpoint: 'permissionSet', ok: true, count: 2 });
        expect(categories.ok).toBe(false);
        expect(categories.message).toMatch(/500/);
    });
});

describe('the schemas list in the report', () => {
    test('lists every schema on the site, flagging which are already wired to a content type', async () => {
        global.fetch = jest.fn(async (url) => {
            const u = new URL(String(url));
            if (u.pathname.endsWith('/schemas')) return jsonResponse(['article', 'faq', 'gh-agenda', 'some-other-thing']);
            return jsonResponse({ items: [] });
        });

        const report = await checkConnection('site', 'token');

        expect(report.schemas).toEqual([
            { name: 'article', used: true },
            { name: 'faq', used: true },
            { name: 'gh-agenda', used: false },
            { name: 'some-other-thing', used: false }
        ]);
    });

    test('is an empty array when the connection failed, keeping the report shape consistent', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('nope', 401));
        const report = await checkConnection('site', 'bad-token');
        expect(report.schemas).toEqual([]);
    });

    test('a site with no schemas at all yields an empty list, not an error', async () => {
        global.fetch = jest.fn(async (url) => {
            const u = new URL(String(url));
            if (u.pathname.endsWith('/schemas')) return jsonResponse([]);
            return jsonResponse({ items: [] });
        });
        const report = await checkConnection('site', 'token');
        expect(report.schemas).toEqual([]);
    });
});

describe('POST /check-connection', () => {
    test('missing url/apiKey returns 400', async () => {
        const res = await request(app).post('/check-connection').send({});
        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/required/);
    });

    test('a fully successful check returns 200 with the full report shape', async () => {
        global.fetch = jest.fn(async (url) => {
            const u = new URL(String(url));
            if (u.pathname.endsWith('/schemas')) return jsonResponse(['article', 'faq', 'employee', 'enhanced-department', 'resource-directory', 'bpquicklink', 'newsflash', 'event', 'facility']);
            return jsonResponse({ items: [{ id: '1', name: 'General' }] });
        });

        const res = await request(app).post('/check-connection').send({ url: 'site.example.com', apiKey: 'tok' });

        expect(res.status).toBe(200);
        expect(res.body.connection.ok).toBe(true);
        expect(res.body.contentTypes).toHaveLength(9);
        expect(res.body.referenceData).toHaveLength(2);
    });

    // A bad token/site is a SUCCESSFUL check (it correctly found the
    // problem) — the route still returns 200, not a 4xx/5xx of its own.
    test('a rejected token still returns 200 — the report itself carries the failure', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('nope', 401));

        const res = await request(app).post('/check-connection').send({ url: 'site.example.com', apiKey: 'bad-tok' });

        expect(res.status).toBe(200);
        expect(res.body.connection.ok).toBe(false);
        expect(res.body.connection.reason).toBe('auth');
    });
});
