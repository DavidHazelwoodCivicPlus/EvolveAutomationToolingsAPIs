const {
    fetchReferenceData,
    fetchSchemaNames,
    fetchSchemaFieldMap,
    getSchemaNames,
    getSchemaFields,
    fetchContentNameLookup,
    fetchAssetNameLookup,
    CONTENT_NAME_EXTRACTORS,
    postAssetFile
} = require('../../index.js');

function jsonResponse(body, status = 200) {
    return new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });
}
function textResponse(text, status = 200) {
    return new Response(text, { status });
}

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

describe('fetchReferenceData', () => {
    test('builds a lowercase-name -> id map from {items: [...]}', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ items: [{ id: 'id-1', name: 'Fire' }, { id: 'id-2', name: 'Social' }] }));

        const map = await fetchReferenceData('site.example.com', 'categories', 'token123');

        expect(map.get('fire')).toBe('id-1');
        expect(map.get('social')).toBe('id-2');
        expect(global.fetch).toHaveBeenCalledWith(
            'https://content.civicplus.com/api/apps/site.example.com/categories',
            expect.objectContaining({
                method: 'GET',
                headers: expect.objectContaining({ Authorization: 'Bearer token123' })
            })
        );
    });

    test('trims and lowercases names for the lookup key', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ items: [{ id: 'id-1', name: '  Fire Dept  ' }] }));
        const map = await fetchReferenceData('site', 'categories', 'token');
        expect(map.get('fire dept')).toBe('id-1');
    });

    test('skips items with no name', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ items: [{ id: 'id-1' }, { id: 'id-2', name: 'Fire' }] }));
        const map = await fetchReferenceData('site', 'categories', 'token');
        expect(map.size).toBe(1);
    });

    test('defaults to an empty map when `items` is missing from the response', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({}));
        const map = await fetchReferenceData('site', 'categories', 'token');
        expect(map.size).toBe(0);
    });

    test('throws, including the response body, on a non-2xx response', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('permission denied', 403));
        await expect(fetchReferenceData('site', 'categories', 'bad-token')).rejects.toThrow(/403/);
    });
});

describe('fetchSchemaNames', () => {
    test('extracts names from a plain array response', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse(['article', 'faq']));
        expect(await fetchSchemaNames('site', 'token')).toEqual(['article', 'faq']);
    });

    test('extracts names from a {items:[{name|slug}]} response', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ items: [{ name: 'article' }, { slug: 'faq' }] }));
        expect(await fetchSchemaNames('site', 'token')).toEqual(['article', 'faq']);
    });

    test('returns null (does not throw) on a non-2xx response', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('nope', 401));
        expect(await fetchSchemaNames('site', 'bad-token')).toBeNull();
    });

    test('returns null (does not throw) when fetch itself rejects', async () => {
        global.fetch = jest.fn().mockRejectedValue(new Error('network down'));
        expect(await fetchSchemaNames('site', 'token')).toBeNull();
    });

    test('returns null when the list comes back empty', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse([]));
        expect(await fetchSchemaNames('site', 'token')).toBeNull();
    });
});

describe('fetchSchemaFieldMap', () => {
    test('reads fields from json.fields', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ fields: [{ name: 'phone', partitioning: 'invariant' }] }));
        const map = await fetchSchemaFieldMap('site', 'employee', 'token');
        expect(map.get('phone')).toEqual({ name: 'phone', partitioning: 'invariant' });
    });

    test('reads fields from json.schema.fields', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ schema: { fields: [{ name: 'title', partitioning: 'language' }] } }));
        const map = await fetchSchemaFieldMap('site', 'article', 'token');
        expect(map.get('title')).toEqual({ name: 'title', partitioning: 'language' });
    });

    test('reads fields from json.data.fields', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ data: { fields: [{ name: 'title', partitioning: 'language' }] } }));
        const map = await fetchSchemaFieldMap('site', 'article', 'token');
        expect(map.get('title')).toEqual({ name: 'title', partitioning: 'language' });
    });

    test('defaults partitioning to "invariant" when the field omits it', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ fields: [{ name: 'foo' }] }));
        const map = await fetchSchemaFieldMap('site', 'article', 'token');
        expect(map.get('foo').partitioning).toBe('invariant');
    });

    test('keys the map by the NORMALIZED field name, but keeps the real name in the value', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ fields: [{ name: 'Fax-Number', partitioning: 'invariant' }] }));
        const map = await fetchSchemaFieldMap('site', 'department', 'token');
        expect(map.get('faxnumber')).toEqual({ name: 'Fax-Number', partitioning: 'invariant' });
    });

    test('returns null on a non-2xx response', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('nope', 404));
        expect(await fetchSchemaFieldMap('site', 'article', 'token')).toBeNull();
    });

    test('returns null when there is no usable fields array', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({}));
        expect(await fetchSchemaFieldMap('site', 'article', 'token')).toBeNull();
    });

    test('returns null on a network error', async () => {
        global.fetch = jest.fn().mockRejectedValue(new Error('down'));
        expect(await fetchSchemaFieldMap('site', 'article', 'token')).toBeNull();
    });
});

// getSchemaNames/getSchemaFields cache per site (and per site+schema) in
// module-level Maps that live for the whole test run. Every test below
// uses a never-reused site string so tests can't see each other's cache
// entries — this is the "unique site per test" strategy from the review,
// and it means these tests need no changes to index.js itself.
let siteCounter = 0;
const uniqueSite = () => `cache-test-${siteCounter++}.example.com`;

describe('getSchemaNames (cached wrapper around fetchSchemaNames)', () => {
    test('a second call for the same site does not call fetch again', async () => {
        const site = uniqueSite();
        global.fetch = jest.fn().mockResolvedValue(jsonResponse(['article', 'faq']));

        expect(await getSchemaNames(site, 'token')).toEqual(['article', 'faq']);
        expect(await getSchemaNames(site, 'token')).toEqual(['article', 'faq']);
        expect(global.fetch).toHaveBeenCalledTimes(1);
    });

    test('a null (failed) result is NOT cached — the next call retries', async () => {
        const site = uniqueSite();
        global.fetch = jest.fn().mockResolvedValue(textResponse('down', 500));

        await getSchemaNames(site, 'token');
        await getSchemaNames(site, 'token');

        expect(global.fetch).toHaveBeenCalledTimes(2);
    });

    test('different sites are cached independently', async () => {
        const siteA = uniqueSite();
        const siteB = uniqueSite();
        global.fetch = jest.fn().mockResolvedValue(jsonResponse(['article']));

        await getSchemaNames(siteA, 'token');
        await getSchemaNames(siteB, 'token');

        expect(global.fetch).toHaveBeenCalledTimes(2);
    });
});

describe('getSchemaFields (cached wrapper around fetchSchemaFieldMap)', () => {
    test('a second call for the same site+schema does not call fetch again', async () => {
        const site = uniqueSite();
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ fields: [{ name: 'title', partitioning: 'language' }] }));

        await getSchemaFields(site, 'article', 'token');
        await getSchemaFields(site, 'article', 'token');

        expect(global.fetch).toHaveBeenCalledTimes(1);
    });

    test('the same site with a different schema is a separate cache entry', async () => {
        const site = uniqueSite();
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ fields: [{ name: 'title', partitioning: 'language' }] }));

        await getSchemaFields(site, 'article', 'token');
        await getSchemaFields(site, 'faq', 'token');

        expect(global.fetch).toHaveBeenCalledTimes(2);
    });
});

describe('fetchContentNameLookup', () => {
    test('builds a name -> id map using the given extractName function', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({
            items: [
                { id: 'dept-1', data: { department: { en: 'Fire Department' } } },
                { id: 'dept-2', data: { department: { en: 'Parks and Recreation' } } }
            ]
        }));

        const lookup = await fetchContentNameLookup('site', 'enhanced-department', 'token', CONTENT_NAME_EXTRACTORS.departments);

        expect(lookup.get('fire department')).toBe('dept-1');
        expect(lookup.get('parks and recreation')).toBe('dept-2');
    });

    test('CONTENT_NAME_EXTRACTORS.staff combines first + last name, trimmed', () => {
        const item = { id: 'staff-1', data: { firstname: { en: 'Joe' }, lastname: { en: 'Schmoe' } } };
        expect(CONTENT_NAME_EXTRACTORS.staff(item)).toBe('Joe Schmoe');
    });

    test('CONTENT_NAME_EXTRACTORS.staff falls back to iv when a field has no en value', () => {
        const item = { id: 'staff-1', data: { firstname: { iv: 'Joe' }, lastname: { iv: 'Schmoe' } } };
        expect(CONTENT_NAME_EXTRACTORS.staff(item)).toBe('Joe Schmoe');
    });

    test('items with no extractable name are skipped, not added with an undefined key', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({
            items: [{ id: 'dept-1', data: {} }, { id: 'dept-2', data: { department: { en: 'Fire Department' } } }]
        }));
        const lookup = await fetchContentNameLookup('site', 'enhanced-department', 'token', CONTENT_NAME_EXTRACTORS.departments);
        expect(lookup.size).toBe(1);
    });

    test('returns an empty map (not throw) on a non-2xx response', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('forbidden', 403));
        const lookup = await fetchContentNameLookup('site', 'enhanced-department', 'token', CONTENT_NAME_EXTRACTORS.departments);
        expect(lookup.size).toBe(0);
    });
});

describe('fetchAssetNameLookup', () => {
    test('builds a fileName -> id map', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({
            items: [{ id: 'asset-1', fileName: 'Flyer.pdf' }, { id: 'asset-2', fileName: 'photo.jpg' }]
        }));

        const lookup = await fetchAssetNameLookup('site', 'token');

        expect(lookup.get('flyer.pdf')).toBe('asset-1'); // case-insensitive key
        expect(lookup.get('photo.jpg')).toBe('asset-2');
    });

    test('returns an empty map (not throw) on a non-2xx response', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('server error', 500));
        const lookup = await fetchAssetNameLookup('site', 'token');
        expect(lookup.size).toBe(0);
    });
});

describe('postAssetFile', () => {
    test('posts multipart form data with the file, categories, and permissionSet', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ id: 'asset-1' }));

        const result = await postAssetFile({
            siteURL: 'site.example.com',
            accessToken: 'tok',
            parentId: 'parent-123',
            publish: true,
            fileBuffer: Buffer.from('hello'),
            fileName: 'hello.txt',
            mimeType: 'text/plain',
            categoryIds: ['cat-1', 'cat-2'],
            permissionSetId: 'ps-1'
        });

        expect(result).toEqual({ id: 'asset-1' });
        expect(global.fetch).toHaveBeenCalledTimes(1);

        const [calledUrl, options] = global.fetch.mock.calls[0];
        expect(String(calledUrl)).toBe('https://content.civicplus.com/api/apps/site.example.com/assets?parentId=parent-123&publish=true');
        expect(options.headers.Authorization).toBe('Bearer tok');
        expect(options.headers['Content-Type']).toBeUndefined(); // must be left for fetch/FormData to set

        const form = options.body;
        expect(form.getAll('categories')).toEqual(['cat-1', 'cat-2']);
        expect(form.get('permissionSet')).toBe('ps-1');
        expect(form.get('file').name).toBe('hello.txt');
        expect(form.get('file').type).toBe('text/plain');
    });

    test('defaults parentId to the all-zero GUID and publish to false when not provided', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ id: 'asset-2' }));
        await postAssetFile({ siteURL: 'site', accessToken: 'tok', fileBuffer: Buffer.from('x'), fileName: 'x.txt' });

        const [calledUrl] = global.fetch.mock.calls[0];
        expect(String(calledUrl)).toBe('https://content.civicplus.com/api/apps/site/assets?parentId=00000000-0000-0000-0000-000000000000&publish=false');
    });

    test('omits permissionSet from the form when none is given, and skips falsy category ids', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ id: 'asset-3' }));
        await postAssetFile({ siteURL: 'site', accessToken: 'tok', fileBuffer: Buffer.from('x'), fileName: 'x.txt', categoryIds: ['cat-1', undefined, ''] });

        const [, options] = global.fetch.mock.calls[0];
        expect(options.body.get('permissionSet')).toBeNull();
        expect(options.body.getAll('categories')).toEqual(['cat-1']);
    });

    test('throws, including the response body, on a non-2xx response', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('too large', 413));
        await expect(postAssetFile({ siteURL: 'site', accessToken: 'tok', fileBuffer: Buffer.from('x'), fileName: 'x.txt' }))
            .rejects.toThrow(/413/);
    });
});
