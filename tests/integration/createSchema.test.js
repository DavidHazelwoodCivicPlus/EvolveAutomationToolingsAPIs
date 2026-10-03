const request = require('supertest');
const { app, uploadConfig, buildSchemaBlueprint, createSchema } = require('../../index.js');
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

describe('buildSchemaBlueprint', () => {
    test('derives one field per real mapPayload output key, not a guess at the column names', () => {
        const fields = buildSchemaBlueprint(uploadConfig.calendar, 'calendar');
        const names = fields.map(f => f.name);
        // These don't follow any simple transformation of the sheet's column
        // names (DateOfEvent, TimeTest) — only calling mapPayload for real
        // gets them right.
        expect(names).toContain('date-of-event');
        expect(names).toContain('TimeTest');
    });

    test('every field defaults to String, not required', () => {
        const fields = buildSchemaBlueprint(uploadConfig.facility, 'facility');
        expect(fields).toEqual([{ name: 'facilityname', properties: { fieldType: 'String', isRequired: false } }]);
    });

    test('overrides are scoped per type — "department" is a References field on staff, but plain String on resources', () => {
        const staffFields = buildSchemaBlueprint(uploadConfig.staff, 'staff');
        const resourcesFields = buildSchemaBlueprint(uploadConfig.resources, 'resources');

        expect(staffFields.find(f => f.name === 'department').properties.fieldType).toBe('References');
        expect(resourcesFields.find(f => f.name === 'department').properties.fieldType).toBe('String');
    });

    test('asset-reference fields are typed Assets', () => {
        const fields = buildSchemaBlueprint(uploadConfig.calendar, 'calendar');
        expect(fields.find(f => f.name === 'attachments').properties.fieldType).toBe('Assets');
        expect(fields.find(f => f.name === 'submission-pdf').properties.fieldType).toBe('Assets');
    });

    test('compound/custom fields (quicklinks\' link, news\' newsdate) fall back to Json rather than a wrong guess', () => {
        expect(buildSchemaBlueprint(uploadConfig.quicklinks, 'quicklinks')[0].properties.fieldType).toBe('Json');
        const newsFields = buildSchemaBlueprint(uploadConfig.news, 'news');
        expect(newsFields.find(f => f.name === 'newsdate').properties.fieldType).toBe('Json');
    });

    test.each(Object.keys(uploadConfig))('%s produces at least one field and never throws', (type) => {
        const fields = buildSchemaBlueprint(uploadConfig[type], type);
        expect(fields.length).toBeGreaterThan(0);
    });
});

describe('createSchema', () => {
    test('posts the Squidex schema-creation shape: {name, fields, isPublished}', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ id: 'schema-1', name: 'facility' }));
        const fields = [{ name: 'facilityname', properties: { fieldType: 'String', isRequired: false } }];

        await createSchema('site.example.com', 'tok', 'facility', fields);

        expect(global.fetch).toHaveBeenCalledWith(
            'https://content.civicplus.com/api/apps/site.example.com/schemas',
            expect.objectContaining({
                method: 'POST',
                headers: expect.objectContaining({ Authorization: 'Bearer tok' })
            })
        );
        const body = JSON.parse(global.fetch.mock.calls[0][1].body);
        expect(body).toEqual({ name: 'facility', fields, isPublished: true });
    });

    test('a 403 is reported as a likely permissions restriction, not a generic failure', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('forbidden', 403));
        await expect(createSchema('site', 'tok', 'facility', [])).rejects.toThrow(/permission/);
    });

    test('a 409 is reported as "already exists"', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('conflict', 409));
        await expect(createSchema('site', 'tok', 'facility', [])).rejects.toThrow(/already exists/);
    });

    test('any other non-2xx surfaces the raw status and body', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('bad request: invalid field type', 400));
        await expect(createSchema('site', 'tok', 'facility', [])).rejects.toThrow(/400/);
    });
});

describe('POST /create-schema', () => {
    test('missing url/apiKey returns 400', async () => {
        const res = await request(app).post('/create-schema').send({ type: 'facility' });
        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/required/);
    });

    test('an unknown content type returns 400', async () => {
        const res = await request(app).post('/create-schema').send({ url: 'site', apiKey: 'tok', type: 'not-a-real-type' });
        expect(res.status).toBe(400);
        expect(res.body.message).toMatch(/Invalid content type/);
    });

    test('a successful creation returns 200 with a review-it-yourself reminder', async () => {
        global.fetch = jest.fn().mockResolvedValue(jsonResponse({ id: 'schema-1' }));
        const res = await request(app).post('/create-schema').send({ url: 'site', apiKey: 'tok', type: 'facility' });

        expect(res.status).toBe(200);
        expect(res.body.message).toMatch(/Created schema "facility"/);
        expect(res.body.message).toMatch(/Review it/);
    });

    test('a 403 (likely a permissions restriction on this site) surfaces as a 500 with a clear message, not a silent failure', async () => {
        global.fetch = jest.fn().mockResolvedValue(textResponse('forbidden', 403));
        const res = await request(app).post('/create-schema').send({ url: 'site', apiKey: 'tok', type: 'facility' });

        expect(res.status).toBe(500);
        expect(res.body.message).toMatch(/permission/);
    });
});
