// A configurable mock for the global `fetch` that index.js calls. Node 22's
// global Response/URL are used to build real Response instances (so
// `.ok`/`.json()`/`.text()` all behave exactly like a real fetch would),
// rather than ad hoc plain objects that only coincidentally look right.
//
// This dispatcher covers the full set of endpoints processUpload/the routes
// call (schemas, schema fields, reference data, row creation, publish,
// asset upload) so processUpload/routes tests can mock `fetch` ONCE per
// test and only override the piece they actually care about. The lower-
// level network-function tests (tests/integration/network.test.js) use
// plain one-off jest.fn()s instead, since they want to assert on a single
// call in isolation.

function jsonResponse(body, status = 200) {
    return new Response(JSON.stringify(body), {
        status,
        headers: { 'Content-Type': 'application/json' }
    });
}

function textResponse(text, status = 200) {
    return new Response(text, { status });
}

// Slugs that match uploadConfig's *preferred* candidate for every type, so
// a happy-path test doesn't need to specify schemaNames unless it's
// specifically exercising resolution (exact vs normalized vs missing).
const DEFAULT_SCHEMA_NAMES = [
    'article', 'resource-directory', 'faq', 'employee', 'enhanced-department',
    'bpquicklink', 'newsflash', 'event', 'facility'
];

function buildFetchMock(overrides = {}) {
    const {
        schemaNames = DEFAULT_SCHEMA_NAMES,
        schemaNamesStatus = 200, // set e.g. 500 to simulate "/schemas unreachable"
        schemaFields = {},       // { [schemaSlug]: [{name, partitioning}, ...] }
        referenceData = {},      // { permissionSet: [{id,name}], categories: [...] }
        createStatus = 200,
        createResponse = () => ({ id: 'generated-' + Math.random().toString(36).slice(2, 8) }),
        createErrorBody = 'create failed',
        // Full override for the row-creation POST: (parsedBody, url) => Response.
        // Use this when different rows in the same test need different
        // outcomes (e.g. one row succeeds, another fails) — createStatus/
        // createResponse only support one fixed outcome for every row.
        createHandler = null,
        publishStatus = 200,
        assetStatus = 200,
        assetResponse = () => ({ id: 'asset-' + Math.random().toString(36).slice(2, 8) }),
        assetErrorBody = 'asset upload failed',
        // { [schemaSlug]: [contentItem, ...] } — answers GET
        // /api/content/:site/:schema?$top=200, used by
        // fetchContentNameLookup (department/staff reference lookups).
        contentListings = {}
    } = overrides;

    return jest.fn(async (rawUrl, options = {}) => {
        const url = new URL(String(rawUrl));
        const method = (options.method || 'GET').toUpperCase();
        const path = url.pathname;

        // GET /api/apps/<site>/schemas
        if (method === 'GET' && /\/api\/apps\/[^/]+\/schemas$/.test(path)) {
            return schemaNamesStatus >= 200 && schemaNamesStatus < 300
                ? jsonResponse(schemaNames)
                : textResponse('schemas unavailable', schemaNamesStatus);
        }

        // GET /api/apps/<site>/schemas/<schema>
        const fieldsMatch = path.match(/\/api\/apps\/[^/]+\/schemas\/([^/]+)$/);
        if (method === 'GET' && fieldsMatch) {
            const fields = schemaFields[fieldsMatch[1]];
            if (!fields) return textResponse('schema not found', 404);
            return jsonResponse({ fields });
        }

        // POST /api/apps/<site>/assets
        if (method === 'POST' && /\/api\/apps\/[^/]+\/assets$/.test(path)) {
            return assetStatus >= 200 && assetStatus < 300
                ? jsonResponse(assetResponse(url))
                : textResponse(assetErrorBody, assetStatus);
        }

        // GET /api/apps/<site>/<endpoint>  (reference data — permissionSet/categories)
        const refMatch = path.match(/\/api\/apps\/[^/]+\/([^/]+)$/);
        if (method === 'GET' && refMatch) {
            return jsonResponse({ items: referenceData[refMatch[1]] || [] });
        }

        // PUT /api/content/<site>/<endpoint>/<id>/status/
        if (method === 'PUT' && /\/api\/content\/[^/]+\/[^/]+\/[^/]+\/status\/?$/.test(path)) {
            return publishStatus >= 200 && publishStatus < 300
                ? jsonResponse({ status: 'Published' })
                : textResponse('publish failed', publishStatus);
        }

        // GET /api/content/<site>/<schema>  (list existing content — the
        // fetchContentNameLookup case, distinct from the POST case below)
        const listMatch = path.match(/\/api\/content\/[^/]+\/([^/]+)$/);
        if (method === 'GET' && listMatch) {
            return jsonResponse({ items: contentListings[listMatch[1]] || [] });
        }

        // POST /api/content/<site>/<endpoint>  (row creation)
        const createMatch = /\/api\/content\/[^/]+\/[^/]+$/.test(path);
        if (method === 'POST' && createMatch) {
            const body = options.body ? JSON.parse(options.body) : {};
            if (createHandler) return createHandler(body, url);
            return createStatus >= 200 && createStatus < 300
                ? jsonResponse(createResponse(body))
                : textResponse(createErrorBody, createStatus);
        }

        throw new Error(`buildFetchMock: no handler configured for ${method} ${rawUrl}`);
    });
}

module.exports = { jsonResponse, textResponse, buildFetchMock, DEFAULT_SCHEMA_NAMES };
