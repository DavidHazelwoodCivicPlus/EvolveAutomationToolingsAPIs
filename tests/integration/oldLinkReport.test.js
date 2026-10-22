// Old-site link report: finds links in the site's content that still point
// at the old website's files and says whether Evolve has a matching asset.
// Read-only — these tests also check it never writes to the site.

const request = require('supertest');
const XLSX = require('xlsx');
const {
    app,
    parseOldDomains,
    findOldLinks,
    collectStrings,
    itemTitle,
    buildOldLinkReport,
    ASSET_RETRY
} = require('../../index.js');
const { jsonResponse, textResponse } = require('../fixtures/schemaResponses.js');

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

// The real article HTML from cp-examplesites that prompted this report.
const FORM_G_HTML = '<p><strong>FYI:</strong> For assistance with completing <strong>Form G,&nbsp;</strong><strong><strong>Judgment of Incapacity,</strong></strong>we suggest referring to <a href="https://co.monmouth.nj.us/documents/14/Form-G-SAMPLE_Judgement-Of-Incapacity_Person-And-Estate_REV1-17-25.doc.pdf" target="_blank"><strong>Sample Form G</strong> (PDF)</a>. This resource offers helpful guidance.</p>';

describe('parseOldDomains', () => {
    test('strips scheme, path and www., splits on commas/spaces', () => {
        expect(parseOldDomains('https://www.co.monmouth.nj.us/Home, monmouthcounty.gov')).toEqual(['co.monmouth.nj.us', 'monmouthcounty.gov']);
        expect(parseOldDomains('')).toEqual([]);
    });
});

describe('findOldLinks', () => {
    const domains = ['co.monmouth.nj.us'];

    test('finds the Sample Form G link in the real article HTML, with its link text and file name', () => {
        expect(findOldLinks(FORM_G_HTML, domains)).toEqual([{
            url: 'https://co.monmouth.nj.us/documents/14/Form-G-SAMPLE_Judgement-Of-Incapacity_Person-And-Estate_REV1-17-25.doc.pdf',
            linkText: 'Sample Form G (PDF)',
            isFile: true,
            fileName: 'Form-G-SAMPLE_Judgement-Of-Incapacity_Person-And-Estate_REV1-17-25.doc.pdf'
        }]);
    });

    test('tells files from pages, including extensionless DocumentCenter links', () => {
        const html = '<a href="https://co.monmouth.nj.us/page.aspx?ID=5">page</a>'
            + '<a href="https://www.co.monmouth.nj.us/DocumentCenter/View/55">doc</a>'
            + '<img src="https://co.monmouth.nj.us/images/seal.png">';
        const links = findOldLinks(html, domains);
        expect(links.map(l => [l.url, l.isFile])).toEqual([
            ['https://co.monmouth.nj.us/page.aspx?ID=5', false],
            ['https://www.co.monmouth.nj.us/DocumentCenter/View/55', true],
            ['https://co.monmouth.nj.us/images/seal.png', true]
        ]);
    });

    test('decodes &amp; and %20 and catches bare URLs in plain text', () => {
        const links = findOldLinks('See https://co.monmouth.nj.us/files/Budget%202024.pdf. Or <a href="https://co.monmouth.nj.us/x.pdf?a=1&amp;b=2">x</a>', domains);
        expect(links.map(l => l.url)).toEqual(['https://co.monmouth.nj.us/x.pdf?a=1&b=2', 'https://co.monmouth.nj.us/files/Budget%202024.pdf']);
        expect(links[1].fileName).toBe('Budget 2024.pdf');
    });

    test('ignores other domains, including look-alikes', () => {
        const html = '<a href="https://example.com/a.pdf">a</a><a href="https://notco.monmouth.nj.us.evil.com/b.pdf">b</a>';
        expect(findOldLinks(html, domains)).toEqual([]);
    });

    test('a link appearing twice in one field is reported once', () => {
        const html = '<a href="https://co.monmouth.nj.us/a.pdf">A</a> https://co.monmouth.nj.us/a.pdf';
        expect(findOldLinks(html, domains)).toHaveLength(1);
    });
});

describe('collectStrings / itemTitle', () => {
    test('walks partitions, arrays and nested components, naming the field', () => {
        const data = { content: { iv: '<p>x</p>' }, blocks: { iv: [{ body: 'y' }] } };
        expect(collectStrings(data, '', [])).toEqual([
            { field: 'content', text: '<p>x</p>' },
            { field: 'blocks[0].body', text: 'y' }
        ]);
    });

    test('picks a title-like field for the report', () => {
        expect(itemTitle({ data: { content: { iv: 'body' }, title: { iv: 'Guardianship Forms' } } })).toBe('Guardianship Forms');
        expect(itemTitle({ data: { NewsTitle: { iv: '<b>Big</b> News' } } })).toBe('Big News');
    });
});

// A site with two content types, a paged article list, and a few assets.
function siteMock({ articles, faqs = [], assets = [], failSchema } = {}) {
    return jest.fn(async (rawUrl, opts = {}) => {
        const url = new URL(String(rawUrl));
        const method = (opts.method || 'GET').toUpperCase();
        if (method !== 'GET') throw new Error(`report must not write: ${method} ${url}`);

        const top = Number(url.searchParams.get('$top') || 200);
        const skip = Number(url.searchParams.get('$skip') || 0);
        const page = (items) => jsonResponse({ total: items.length, items: items.slice(skip, skip + top) });

        if (/\/api\/apps\/[^/]+\/schemas$/.test(url.pathname)) return jsonResponse([{ name: 'article' }, { name: 'faq' }]);
        if (/\/api\/apps\/[^/]+\/assets$/.test(url.pathname)) return page(assets);
        const content = url.pathname.match(/\/api\/content\/[^/]+\/([^/]+)$/);
        if (content) {
            if (content[1] === failSchema) return textResponse('forbidden', 403);
            expect(opts.headers['X-Unpublished']).toBe('1');
            return page(content[1] === 'article' ? articles : faqs);
        }
        throw new Error('unexpected call: ' + url);
    });
}

describe('buildOldLinkReport', () => {
    test('scans every page of every content type and matches file links to assets', async () => {
        const filler = Array.from({ length: 250 }, (_, i) => ({ id: `filler-${i}`, data: { title: { iv: `Filler ${i}` } } }));
        global.fetch = siteMock({
            articles: [
                ...filler,
                { id: 'a1', status: 'Published', data: { title: { iv: 'Guardianship Forms' }, content: { iv: FORM_G_HTML } } },
                { id: 'a2', status: 'Draft', data: { title: { iv: 'Planning' }, content: { iv: '<a href="https://co.monmouth.nj.us/documents/24/DRC_Agenda_101524.pdf">Agenda</a> <a href="https://co.monmouth.nj.us/documents/24/Form_F.pdf">F</a> <a href="https://co.monmouth.nj.us/Planning.aspx">Planning page</a>' } } }
            ],
            faqs: [{ id: 'f1', data: { question: { iv: 'Q?' }, answer: { iv: 'Nothing old here' } } }],
            assets: [
                { id: 'asset-drc', fileName: 'DRC_Agenda_101524.pdf', _links: { content: { href: 'https://content.civicplus.com/api/assets/x/asset-drc' } } },
                { id: 'asset-f', fileName: 'Form F.pdf' }
            ]
        });

        const { fileLinks, pageLinks, summary } = await buildOldLinkReport({ siteURL: 'cp-examplesites', accessToken: 't', oldDomains: 'co.monmouth.nj.us' });

        expect(summary.itemsScanned).toBe(253); // paging reached item 251+
        expect(fileLinks).toHaveLength(3);
        expect(pageLinks.map(l => l.oldUrl)).toEqual(['https://co.monmouth.nj.us/Planning.aspx']);

        const byFile = Object.fromEntries(fileLinks.map(l => [l.fileName, l]));
        const formG = byFile['Form-G-SAMPLE_Judgement-Of-Incapacity_Person-And-Estate_REV1-17-25.doc.pdf'];
        expect(formG.matchType).toBe('none');
        expect(formG.result).toMatch(/NO matching file/);
        expect(formG.title).toBe('Guardianship Forms');
        expect(formG.linkText).toBe('Sample Form G (PDF)');
        expect(formG.field).toBe('content');
        expect(formG.editLink).toBe('https://content.civicplus.com/app/cp-examplesites/content/article/a1');

        expect(byFile['DRC_Agenda_101524.pdf'].matchType).toBe('exact');
        expect(byFile['DRC_Agenda_101524.pdf'].matchedAssetIds).toBe('asset-drc');
        expect(byFile['DRC_Agenda_101524.pdf'].matchedAssetUrls).toBe('https://content.civicplus.com/api/assets/x/asset-drc');
        expect(byFile['DRC_Agenda_101524.pdf'].status).toBe('Draft');

        expect(byFile['Form_F.pdf'].matchType).toBe('similar');
        expect(byFile['Form_F.pdf'].matchedAssets).toBe('Form F.pdf');

        expect(summary).toMatchObject({ fileLinks: 3, fileLinksWithNoMatch: 1, fileLinksWithSimilarMatch: 1, fileLinksWithExactMatch: 1, itemsWithFileLinks: 2, pageLinks: 1 });
    });

    test('a content type that cannot be read is reported, not fatal', async () => {
        global.fetch = siteMock({ articles: [{ id: 'a1', data: { content: { iv: FORM_G_HTML } } }], failSchema: 'faq' });
        const { summary, fileLinks } = await buildOldLinkReport({ siteURL: 's', accessToken: 't', oldDomains: 'co.monmouth.nj.us' });
        expect(fileLinks).toHaveLength(1);
        expect(summary.schemaErrors[0]).toMatch(/^faq: .*403/);
    });
});

describe('POST /old-link-report', () => {
    test('returns an .xlsx with the no-match links first, plus a summary header', async () => {
        global.fetch = siteMock({
            articles: [
                { id: 'a2', data: { title: { iv: 'Planning' }, content: { iv: '<a href="https://co.monmouth.nj.us/documents/24/DRC.pdf">Agenda</a>' } } },
                { id: 'a1', data: { title: { iv: 'Guardianship Forms' }, content: { iv: FORM_G_HTML } } }
            ],
            assets: [{ id: 'asset-drc', fileName: 'DRC.pdf' }]
        });

        const res = await request(app)
            .post('/old-link-report')
            .send({ url: 'cp-examplesites', apiKey: 'token', oldDomains: 'co.monmouth.nj.us' })
            .buffer(true)
            .parse((r, cb) => { const chunks = []; r.on('data', c => chunks.push(c)); r.on('end', () => cb(null, Buffer.concat(chunks))); });

        expect(res.status).toBe(200);
        expect(res.headers['content-disposition']).toMatch(/OldLinkReport-cp-examplesites\.xlsx/);
        const summary = JSON.parse(decodeURIComponent(res.headers['x-report-summary']));
        expect(summary.fileLinksWithNoMatch).toBe(1);

        const wb = XLSX.read(res.body, { type: 'buffer' });
        expect(wb.SheetNames).toEqual(['Summary', 'File links', 'Other old-site links']);
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['File links']);
        expect(rows.map(r => r['File name'])).toEqual(['Form-G-SAMPLE_Judgement-Of-Incapacity_Person-And-Estate_REV1-17-25.doc.pdf', 'DRC.pdf']);
        expect(rows[0]['Edit in Evolve']).toBe('https://content.civicplus.com/app/cp-examplesites/content/article/a1');
        expect(rows[1]['Matching Evolve asset(s)']).toBe('DRC.pdf');
    });

    test('400 without connection details or a domain', async () => {
        expect((await request(app).post('/old-link-report').send({ oldDomains: 'x.gov' })).status).toBe(400);
        const res = await request(app).post('/old-link-report').send({ url: 's', apiKey: 't', oldDomains: ' ' });
        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/domain/);
    });

    test('a site whose content types cannot be listed gives a clear 500', async () => {
        global.fetch = jest.fn(async () => textResponse('nope', 401));
        const res = await request(app).post('/old-link-report').send({ url: 's', apiKey: 't', oldDomains: 'co.monmouth.nj.us' });
        expect(res.status).toBe(500);
        expect(res.body.message).toMatch(/Test connection/);
    });
});
