const { uploadConfig } = require('../../index.js');

let warnSpy;
beforeEach(() => { warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {}); });
afterEach(() => { warnSpy.mockRestore(); });

const lookups = {
    permissionSet: new Map([['general', 'ps-general-id'], ['admin', 'ps-admin-id']]),
    categories: new Map([['fire', 'cat-fire-id'], ['social', 'cat-social-id']]),
    departments: new Map([['fire department', 'dept-fire-id'], ['parks and recreation', 'dept-parks-id']]),
    staff: new Map([['joe schmoe', 'staff-joe-id'], ['jane doe', 'staff-jane-id']]),
    assets: new Map([['flyer.pdf', 'asset-flyer-id'], ['photo.jpg', 'asset-photo-id']])
};

describe('articles.mapPayload', () => {
    const mapPayload = uploadConfig.articles.mapPayload;

    test('maps a full row', () => {
        const entry = {
            Title: 'July Test',
            Content: '<p>hi</p>',
            ContentType: 'article',
            contentTypeDisplayName: 'Article',
            Publish: 'Yes',
            PermissionSet: 'General',
            Categories: 'Fire, Social',
            Tags: 'a, b'
        };
        expect(mapPayload(entry, lookups)).toEqual({
            data: { name: { en: 'July Test' }, article: { en: '<p>hi</p>' } },
            contentTypeName: 'article',
            contentTypeDisplayName: 'Article',
            permissionSet: { id: 'ps-general-id', name: 'General' },
            tags: ['a', 'b'],
            categories: [
                { id: 'cat-fire-id', name: 'Fire' },
                { id: 'cat-social-id', name: 'Social' }
            ],
            publish: 'Yes'
        });
    });

    test('a sparse row (the real sheet has rows with no Tags) leaves Tags/Categories/Publish undefined, not empty', () => {
        const entry = { Title: 'T', Content: 'C', ContentType: 'article', PermissionSet: 'General', Categories: 'Fire' };
        const result = mapPayload(entry, lookups);
        expect(result.tags).toBeUndefined();
    });
});

describe('resources.mapPayload', () => {
    const mapPayload = uploadConfig.resources.mapPayload;
    const fullEntry = {
        Department: 'Putnam County Chamber', MaskedEmail: 'Test mask email', PhoneNumber: '123-456-7890',
        AdditionalNumber: '555-444-5555', FaxNumber: '333-444-8888', Address1: 'Test address 1', Address2: 'test address 2',
        City: 'fake city', State: 'MO', Zip: '66502', HoursOfOperation: '5 am to 6 am', AdditionalInformation: 'More info',
        Name: 'Name', Email: 'email@email.com', Fax: '555-666-8888', Hours: '9 pm to 5 pm', WebsiteUrl: 'test.com',
        WebsiteDisplayName: 'Fake website name', Description: 'description goes here', Publish: 'Yes',
        PermissionSet: 'General', Categories: 'Fire'
    };

    // resources builds TWO overlapping sets of keys (lowercase + PascalCase)
    // from the same source values in both `data` and `searchFields`. This
    // may be deliberate (covering two schema-naming conventions) or leftover
    // duplication — either way, freeze it so a future cleanup can't change
    // the payload shape without the snapshot flagging it. Run `npx jest -u`
    // to intentionally update this snapshot after a real change.
    test('matches the known output shape, including the lowercase+PascalCase duplication (snapshot)', () => {
        expect(mapPayload(fullEntry, lookups)).toMatchSnapshot();
    });

    test('titles is keyed off Department', () => {
        expect(mapPayload(fullEntry, lookups).titles).toEqual({ department: { en: 'Putnam County Chamber' } });
    });
});

describe('faqs.mapPayload', () => {
    const mapPayload = uploadConfig.faqs.mapPayload;

    test('maps Question/Answer', () => {
        const result = mapPayload({ Question: 'Test Question', Answer: 'Test Awnser', Publish: 'Yes', PermissionSet: 'General', Categories: 'Fire' }, lookups);
        expect(result.data).toEqual({ question: { en: 'Test Question' }, answer: { en: 'Test Awnser' } });
    });

    test('a row missing Publish/PermissionSet (the real sheet has such rows) does not throw', () => {
        const result = mapPayload({ Question: 'Q2', Answer: 'A2', Categories: 'Fire' }, lookups);
        expect(result.publish).toBeUndefined();
        expect(result.permissionSet).toEqual({ id: undefined, name: undefined });
    });
});

describe('staff.mapPayload', () => {
    const mapPayload = uploadConfig.staff.mapPayload;

    test('maps a full row', () => {
        const entry = {
            FirstName: 'Joe', LastName: 'Schmoe', Title: 'Fake title', Department: 'test',
            PhoneNumber: '555-555-5555', FaxNumber: '', EmailAddress: 'fakeEmail@gmail.com', Biography: 'Fake bio',
            Publish: 'Yes', PermissionSet: 'General', Categories: 'Fire'
        };
        const result = mapPayload(entry, lookups);
        expect(result.data.firstname).toEqual({ en: 'Joe' });
        expect(result.data.lastname).toEqual({ en: 'Schmoe' });
        expect(result.data.emailaddress).toEqual({ en: 'fakeEmail@gmail.com' });
    });

    test('department resolves a name against the departments lookup', () => {
        const result = mapPayload({ FirstName: 'Joe', LastName: 'Schmoe', Department: 'Fire Department' }, lookups);
        // `en`, not `iv` — a live run returned "department: Not a known language" for iv.
        expect(result.data.department).toEqual({ en: ['dept-fire-id'] });
    });

    test('a department name with no match is left out (not a crash, not a bad id)', () => {
        const result = mapPayload({ FirstName: 'Joe', LastName: 'Schmoe', Department: 'Made Up Department' }, lookups);
        expect(result.data.department).toEqual({ en: [] });
    });

    test('a blank Department is an empty reference array, not undefined', () => {
        const result = mapPayload({ FirstName: 'Joe', LastName: 'Schmoe' }, lookups);
        expect(result.data.department).toEqual({ en: [] });
    });
});

describe('departments.mapPayload', () => {
    const mapPayload = uploadConfig.departments.mapPayload;

    test('maps a full row', () => {
        const entry = {
            Department: 'Test Department', PhoneNumber: '111-111-1111', EmergencyNumber: '222-222-2222',
            FaxNumber: '333-333-3333', HoursOfOperation: 'Open never', AdditionalInformation: 'more info here',
            Publish: 'Yes', PermissionSet: 'General', Categories: 'Fire, Social'
        };
        const result = mapPayload(entry, lookups);
        expect(result.data.department).toEqual({ en: 'Test Department' });
        expect(result.categories).toEqual([
            { id: 'cat-fire-id', name: 'Fire' },
            { id: 'cat-social-id', name: 'Social' }
        ]);
    });

    test('parentdepartment resolves a name against the (self-referential) departments lookup', () => {
        const result = mapPayload({ Department: 'D', ParentDepartment: 'Parks and Recreation' }, lookups);
        expect(result.data.parentdepartment).toEqual({ iv: ['dept-parks-id'] });
    });

    test('staffdirectory resolves a comma-separated list of names against the staff lookup', () => {
        const result = mapPayload({ Department: 'D', StaffDirectory: 'Joe Schmoe, Jane Doe' }, lookups);
        expect(result.data.staffdirectory).toEqual({ iv: ['staff-joe-id', 'staff-jane-id'] });
    });

    test('blank ParentDepartment/StaffDirectory are empty reference arrays, not undefined', () => {
        const result = mapPayload({ Department: 'D' }, lookups);
        expect(result.data.parentdepartment).toEqual({ iv: [] });
        expect(result.data.staffdirectory).toEqual({ iv: [] });
    });
});

describe('quicklinks.mapPayload', () => {
    const mapPayload = uploadConfig.quicklinks.mapPayload;

    test('maps Link/Name and always opens in a new window', () => {
        const entry = { Link: 'www.youtube.com/watch?v=hPtVP_r7wng&t=2s', Name: 'David Hazelwood', Publish: 'Yes', PermissionSet: 'General', Categories: 'Fire' };
        const result = mapPayload(entry, lookups);
        expect(result.data.link).toEqual({
            en: { url: 'www.youtube.com/watch?v=hPtVP_r7wng&t=2s', displayName: 'David Hazelwood', openInNewWindow: true }
        });
    });
});

describe('news.mapPayload', () => {
    const mapPayload = uploadConfig.news.mapPayload;

    test('uses the sheet\'s NewsDate for both ends of the range when provided', () => {
        const result = mapPayload({ NewsTitle: 'T', NewsDate: '2024-06-01T00:00:00Z', NewsText: 'text', Publish: 'Yes' }, lookups);
        expect(result.data.newsdate).toEqual({
            iv: { startDate: '2024-06-01T00:00:00Z', endDate: '2024-06-01T00:00:00Z' }
        });
    });

    test('falls back to the current date/time (not a hardcoded past date) when NewsDate is blank — this is the exact shape of the real sample sheet\'s row', () => {
        const before = Date.now();
        const result = mapPayload({ NewsTitle: 'Test News', NewsText: 'NewsText Test', Publish: 'Yes' }, lookups);
        const after = Date.now();

        expect(result.data.newsdate.iv.startDate).toBe(result.data.newsdate.iv.endDate);
        const parsed = new Date(result.data.newsdate.iv.startDate).getTime();
        expect(parsed).toBeGreaterThanOrEqual(before);
        expect(parsed).toBeLessThanOrEqual(after);
    });

    test('newsasset resolves a filename against the assets lookup', () => {
        const result = mapPayload({ NewsTitle: 'T', NewsAsset: 'flyer.pdf' }, lookups);
        expect(result.data.newsasset).toEqual({ iv: ['asset-flyer-id'] });
    });

    test('a blank NewsAsset is an empty reference array, not undefined', () => {
        expect(mapPayload({ NewsTitle: 'T' }, lookups).data.newsasset).toEqual({ iv: [] });
    });
});

describe('calendar.mapPayload', () => {
    const mapPayload = uploadConfig.calendar.mapPayload;

    test('maps event fields, including the invariant-partitioned "name-of-event" field', () => {
        const entry = { TitleOfEvent: 'Town Hall', DateOfEvent: '2024-06-01', StartTimeOfEvent: '17:00', Details: 'details', UrlLink: 'https://civicplus.com', Publish: 'Yes' };
        const result = mapPayload(entry, lookups);
        expect(result.data['name-of-event']).toEqual({ iv: 'Town Hall' });
        expect(result.data['date-of-event']).toEqual({ iv: '2024-06-01' });
        expect(result.data['url-link']).toEqual({ iv: 'https://civicplus.com' });
    });

    test('url-link is null (not the old "test.com" placeholder) when blank — this is the exact shape of the real sample sheet\'s row', () => {
        const result = mapPayload({ TitleOfEvent: 'TestCalendar', Publish: 'Yes' }, lookups);
        expect(result.data['url-link']).toEqual({ iv: null });
    });

    test('attachments and submission-pdf resolve filenames against the assets lookup', () => {
        const result = mapPayload({ TitleOfEvent: 'T', Attachments: 'flyer.pdf, photo.jpg', SubmissionPDF: 'flyer.pdf' }, lookups);
        expect(result.data.attachments).toEqual({ iv: ['asset-flyer-id', 'asset-photo-id'] });
        expect(result.data['submission-pdf']).toEqual({ iv: ['asset-flyer-id'] });
    });

    test('blank Attachments/SubmissionPDF are empty reference arrays, not undefined', () => {
        const result = mapPayload({ TitleOfEvent: 'T' }, lookups);
        expect(result.data.attachments).toEqual({ iv: [] });
        expect(result.data['submission-pdf']).toEqual({ iv: [] });
    });
});

describe('facility.mapPayload', () => {
    const mapPayload = uploadConfig.facility.mapPayload;

    test('maps FacilityName', () => {
        const result = mapPayload({ FacilityName: 'TestFacility', Publish: 'Yes', PermissionSet: 'General' }, lookups);
        expect(result.data.facilityname).toEqual({ en: 'TestFacility' });
    });

    test('mapPayload itself has no bug — it just transforms whatever entry it is handed; the real sheetIndex issue is covered in processUpload.test.js', () => {
        expect(mapPayload({}, lookups).data.facilityname).toEqual({ en: undefined });
    });
});

// A lightweight smoke test across every configured content type: none of
// them should throw when given a nearly-empty row. The real workbook has
// plenty of sparse/incomplete rows (QuickLinks, News, Calendar all have
// rows with only PermissionSet + Categories filled in), so this is a
// realistic case, not a hypothetical one.
describe('every mapPayload tolerates a nearly-empty row', () => {
    test.each(Object.keys(uploadConfig))('%s.mapPayload does not throw on {}', (type) => {
        expect(() => uploadConfig[type].mapPayload({}, lookups)).not.toThrow();
    });

    test.each(Object.keys(uploadConfig))('%s.mapPayload always returns a `data` object', (type) => {
        const result = uploadConfig[type].mapPayload({}, lookups);
        expect(result.data).toBeDefined();
        expect(typeof result.data).toBe('object');
    });
});
