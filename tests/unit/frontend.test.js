/**
 * @jest-environment jsdom
 */

// Tests the dashboard's own JavaScript — the gap every earlier pass left
// open. app.js was extracted out of index.html specifically so it could
// be loaded here.
//
// The DOM comes from the REAL index.html (parsed off disk, not a
// hand-written fixture), so a renamed or deleted element breaks these
// tests instead of silently breaking the page. fetch is mocked, so
// nothing here talks to a server.

const fs = require('node:fs');
const path = require('node:path');


const PROJECT_ROOT = path.join(__dirname, '..', '..');
const html = fs.readFileSync(path.join(PROJECT_ROOT, 'index.html'), 'utf8');
const appJs = fs.readFileSync(path.join(PROJECT_ROOT, 'app.js'), 'utf8');

const BUILT_IN_TYPES = [
    { type: 'articles', label: 'Articles', description: 'General pages and blog-style posts', sheetName: 'Articles', isCustom: false },
    { type: 'faqs', label: 'FAQs', description: 'Frequently asked questions', sheetName: 'FAQs', isCustom: false }
];

// Loads the page + script fresh, and resolves once the initial
// loadContentTypes() call has settled — otherwise assertions race the
// buttons being rendered.
//
// app.js is compiled into a function per load because its top-level
// `const`s would otherwise collide on a second run. A real browser gives
// every page load a fresh global; this reproduces that without changing
// app.js itself. (`new Function` rather than `vm.runInThisContext` so the
// code sees jsdom's `document`/`window` rather than the bare Node realm.)
async function loadPage() {
    document.documentElement.innerHTML = html;
    // eslint-disable-next-line no-new-func
    new Function(appJs)();
    await flushPromises();
}

// app.js chains a few awaits; several macrotask turns clear them all.
async function flushPromises() {
    for (let i = 0; i < 6; i++) await new Promise(resolve => setTimeout(resolve, 0));
}

function jsonOk(body) {
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });
}
function jsonErr(body, status = 400) {
    return Promise.resolve({ ok: false, status, statusText: 'Error', json: () => Promise.resolve(body) });
}

// A File the change handler can read. jsdom supports FileList poorly, so
// the input's `files` is redefined directly.
function attachFile(input, name = 'book.xlsx') {
    const file = new File(['x'], name, { type: 'application/vnd.ms-excel' });
    Object.defineProperty(input, 'files', { value: [file], configurable: true });
    return file;
}

beforeEach(() => {
    jest.useRealTimers();
    global.fetch = jest.fn((url) => {
        if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
        return jsonOk({});
    });
});

afterEach(() => {
    jest.restoreAllMocks();
});

describe('content type buttons', () => {
    test('are built from /content-types, not hardcoded in the HTML', async () => {
        // The markup ships an empty container — proving the buttons below
        // genuinely came from the server response.
        expect(html).toContain('<div id="buttonsDiv"></div>');

        await loadPage();

        const buttons = document.querySelectorAll('.upload-btn');
        expect(buttons).toHaveLength(2);
        expect(buttons[0].dataset.name).toBe('Articles');
        expect(buttons[0].dataset.endpoint).toBe('/upload/articles');
    });

    test('a custom type renders as a button just like a built-in one', async () => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) {
                return jsonOk({ types: [...BUILT_IN_TYPES, { type: 'custom_agendas', label: 'Agendas', description: 'Custom', sheetName: 'Agendas', isCustom: true }] });
            }
            return jsonOk({});
        });

        await loadPage();

        const labels = [...document.querySelectorAll('.upload-btn')].map(b => b.dataset.name);
        expect(labels).toContain('Agendas');
    });

    test('carry an accessible label describing what they upload', async () => {
        await loadPage();
        const button = document.querySelector('.upload-btn');
        expect(button.getAttribute('aria-label')).toMatch(/Upload Articles/);
    });

    test('clicking with no file selected warns instead of uploading', async () => {
        await loadPage();
        const callsBefore = global.fetch.mock.calls.length;

        document.querySelector('.upload-btn').click();
        await flushPromises();

        expect(document.getElementById('results').innerText).toMatch(/choose an Excel file/i);
        expect(global.fetch.mock.calls.length).toBe(callsBefore);
    });
});

describe('workbook inspection', () => {
    const INSPECT_RESPONSE = {
        sheets: [
            { name: 'Articles', columns: ['Title'], rowCount: 3, matchedType: 'articles', unmatched: false },
            { name: 'FAQs', columns: ['Question', 'Answer'], rowCount: 2, matchedType: 'faqs', unmatched: false },
            { name: 'Agendas and Minutes', columns: ['AgendaTitle', 'MeetingDate'], rowCount: 5, matchedType: null, unmatched: true }
        ]
    };

    beforeEach(() => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/inspect-workbook')) return jsonOk(INSPECT_RESPONSE);
            return jsonOk({});
        });
    });

    test('choosing a file shows its name and inspects it', async () => {
        await loadPage();
        const input = document.getElementById('inputFile');
        attachFile(input, 'EvolveUploads.xlsx');

        input.dispatchEvent(new Event('change'));
        await flushPromises();

        expect(document.getElementById('fileName').innerText).toMatch(/EvolveUploads.xlsx/);
        expect(global.fetch.mock.calls.some(([u]) => String(u).includes('/inspect-workbook'))).toBe(true);
    });

    test('each content type gets a sheet picker, preselected to the matched sheet', async () => {
        await loadPage();
        const input = document.getElementById('inputFile');
        attachFile(input);
        input.dispatchEvent(new Event('change'));
        await flushPromises();

        const pickers = document.querySelectorAll('.sheet-picker select');
        expect(pickers).toHaveLength(2);
        // The FAQs type should have preselected the FAQs sheet.
        const faqsRow = [...document.querySelectorAll('.type-row')].find(r => r.querySelector('.upload-btn').dataset.type === 'faqs');
        expect(faqsRow.querySelector('select').value).toBe('FAQs');
    });

    test('a sheet matching no known type appears in the custom imports panel', async () => {
        await loadPage();
        const input = document.getElementById('inputFile');
        attachFile(input);
        input.dispatchEvent(new Event('change'));
        await flushPromises();

        expect(document.getElementById('customImportsPanel').hidden).toBe(false);
        const unmatched = document.getElementById('unmatchedSheets').textContent;
        expect(unmatched).toMatch(/Agendas and Minutes/);
        expect(unmatched).toMatch(/AgendaTitle/); // its columns are shown
    });

    test('the custom imports panel stays hidden when every sheet matches', async () => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/inspect-workbook')) {
                return jsonOk({ sheets: INSPECT_RESPONSE.sheets.filter(s => !s.unmatched) });
            }
            return jsonOk({});
        });

        await loadPage();
        const input = document.getElementById('inputFile');
        attachFile(input);
        input.dispatchEvent(new Event('change'));
        await flushPromises();

        expect(document.getElementById('customImportsPanel').hidden).toBe(true);
    });
});

describe('sheet override', () => {
    beforeEach(() => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/inspect-workbook')) {
                return jsonOk({
                    sheets: [
                        { name: 'FAQs', columns: ['Question'], rowCount: 2, matchedType: 'faqs', unmatched: false },
                        { name: 'OtherSheet', columns: ['Question'], rowCount: 1, matchedType: null, unmatched: true }
                    ]
                });
            }
            return jsonOk({ message: 'Upload complete. Processed 1 faqs.' });
        });
    });

    test('changing the dropdown sends sheetOverride with the upload', async () => {
        await loadPage();
        const input = document.getElementById('inputFile');
        attachFile(input);
        input.dispatchEvent(new Event('change'));
        await flushPromises();

        document.getElementById('inputURL').value = 'site.example.com';
        document.getElementById('inputToken').value = 'tok';

        const faqsRow = [...document.querySelectorAll('.type-row')].find(r => r.querySelector('.upload-btn').dataset.type === 'faqs');
        const select = faqsRow.querySelector('select');
        select.value = 'OtherSheet';
        select.dispatchEvent(new Event('change'));

        faqsRow.querySelector('.upload-btn').click();
        await flushPromises();

        const uploadCall = global.fetch.mock.calls.find(([u]) => String(u) === '/upload/faqs');
        expect(uploadCall[1].body.get('sheetOverride')).toBe('OtherSheet');
    });

    test('an untouched dropdown sends no sheetOverride, leaving the server to match by name', async () => {
        await loadPage();
        const input = document.getElementById('inputFile');
        attachFile(input);
        input.dispatchEvent(new Event('change'));
        await flushPromises();

        document.getElementById('inputURL').value = 'site.example.com';
        document.getElementById('inputToken').value = 'tok';

        const faqsRow = [...document.querySelectorAll('.type-row')].find(r => r.querySelector('.upload-btn').dataset.type === 'faqs');
        faqsRow.querySelector('.upload-btn').click();
        await flushPromises();

        const uploadCall = global.fetch.mock.calls.find(([u]) => String(u) === '/upload/faqs');
        expect(uploadCall[1].body.get('sheetOverride')).toBeNull();
    });
});

describe('defining a custom import', () => {
    beforeEach(() => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/inspect-workbook')) {
                return jsonOk({ sheets: [{ name: 'Agendas and Minutes', columns: ['AgendaTitle'], rowCount: 5, matchedType: null, unmatched: true }] });
            }
            if (String(url).includes('/custom-types')) return jsonOk({ message: 'Saved "Agendas"', type: 'custom_agendas' });
            return jsonOk({});
        });
    });

    async function openCustomForm() {
        await loadPage();
        const input = document.getElementById('inputFile');
        attachFile(input);
        input.dispatchEvent(new Event('change'));
        await flushPromises();
        // Scoped to the unrecognized-sheets container: the hand-registration
        // form shares the .custom-form class and appears earlier in the DOM.
        return document.querySelector('#unmatchedSheets .custom-form');
    }

    test('the form is prefilled with the sheet name as a starting label', async () => {
        const form = await openCustomForm();
        expect(form.querySelectorAll('input[type="text"]')[0].value).toBe('Agendas and Minutes');
    });

    test('submitting with an empty schema slug shows an inline error and sends nothing', async () => {
        const form = await openCustomForm();
        const callsBefore = global.fetch.mock.calls.filter(([u]) => String(u).includes('/custom-types')).length;

        form.querySelectorAll('input[type="text"]')[1].value = '';
        form.querySelector('.custom-save-btn').click();
        await flushPromises();

        const error = form.querySelector('.custom-error');
        expect(error.hidden).toBe(false);
        expect(error.textContent).toMatch(/required/i);
        expect(global.fetch.mock.calls.filter(([u]) => String(u).includes('/custom-types')).length).toBe(callsBefore);
    });

    test('a valid submission posts the label, slug, sheet name, and columns', async () => {
        const form = await openCustomForm();
        const [labelInput, endpointInput] = form.querySelectorAll('input[type="text"]');
        labelInput.value = 'Agendas';
        endpointInput.value = 'gh-agenda';

        form.querySelector('.custom-save-btn').click();
        await flushPromises();

        const saveCall = global.fetch.mock.calls.find(([u]) => String(u).includes('/custom-types'));
        expect(JSON.parse(saveCall[1].body)).toEqual({
            label: 'Agendas',
            endpoint: 'gh-agenda',
            sheetName: 'Agendas and Minutes',
            columns: ['AgendaTitle']
        });
    });

    test('a server-side rejection is shown inline and the button becomes usable again', async () => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/inspect-workbook')) {
                return jsonOk({ sheets: [{ name: 'Agendas and Minutes', columns: ['AgendaTitle'], rowCount: 5, matchedType: null, unmatched: true }] });
            }
            return jsonErr({ message: 'collides with a built-in content type' });
        });

        const form = await openCustomForm();
        form.querySelectorAll('input[type="text"]')[1].value = 'gh-agenda';
        const saveBtn = form.querySelector('.custom-save-btn');
        saveBtn.click();
        await flushPromises();

        expect(form.querySelector('.custom-error').textContent).toMatch(/collides/);
        expect(saveBtn.disabled).toBe(false);
    });
});

describe('test connection', () => {
    test('requires a URL and token before calling the server', async () => {
        await loadPage();
        const callsBefore = global.fetch.mock.calls.length;

        document.getElementById('testConnectionBtn').click();
        await flushPromises();

        expect(document.getElementById('results').innerText).toMatch(/required/i);
        expect(global.fetch.mock.calls.length).toBe(callsBefore);
    });

    test('renders a pass/fail row per content type, with no create-schema action', async () => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/check-connection')) {
                return jsonOk({
                    connection: { ok: true, message: 'Connected' },
                    contentTypes: [
                        { type: 'articles', label: 'Articles', ok: true },
                        { type: 'faqs', label: 'FAQs', ok: false }
                    ],
                    referenceData: [{ endpoint: 'permissionSet', ok: true, count: 4 }]
                });
            }
            return jsonOk({});
        });

        await loadPage();
        document.getElementById('inputURL').value = 'site.example.com';
        document.getElementById('inputToken').value = 'tok';
        document.getElementById('testConnectionBtn').click();
        await flushPromises();

        const result = document.getElementById('connectionCheckResult');
        expect(result.querySelectorAll('.check-ok').length).toBeGreaterThan(0);
        expect(result.querySelectorAll('.check-fail').length).toBe(1);
        // The Create Schema feature was removed — nothing in the report
        // should offer to change the site.
        expect(result.querySelector('.create-schema-btn')).toBeNull();
        expect(result.textContent).not.toMatch(/create schema/i);
    });

    test('a failed connection shows the reason and skips the per-type detail', async () => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/check-connection')) {
                return jsonOk({ connection: { ok: false, message: 'Token was rejected (HTTP 401)' }, contentTypes: [], referenceData: [] });
            }
            return jsonOk({});
        });

        await loadPage();
        document.getElementById('inputURL').value = 'site.example.com';
        document.getElementById('inputToken').value = 'bad';
        document.getElementById('testConnectionBtn').click();
        await flushPromises();

        const result = document.getElementById('connectionCheckResult');
        expect(result.textContent).toMatch(/Token was rejected/);
        expect(result.querySelector('.check-subheading')).toBeNull();
    });
});

describe('the site schema list', () => {
    function withSchemas(schemas) {
        return jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/check-connection')) {
                return jsonOk({
                    connection: { ok: true, message: 'Connected' },
                    contentTypes: [{ type: 'articles', label: 'Articles', ok: true }],
                    referenceData: [],
                    schemas
                });
            }
            return jsonOk({});
        });
    }

    async function runCheck(schemas) {
        global.fetch = withSchemas(schemas);
        await loadPage();
        document.getElementById('inputURL').value = 'site.example.com';
        document.getElementById('inputToken').value = 'tok';
        document.getElementById('testConnectionBtn').click();
        await flushPromises();
    }

    test('shows every schema after a connection check, with a count', async () => {
        await runCheck([
            { name: 'article', used: true },
            { name: 'gh-agenda', used: false }
        ]);

        const list = document.querySelector('.schema-list');
        expect(list).not.toBeNull();
        expect(list.querySelector('summary').textContent).toMatch(/Schemas on this site \(2\)/);
        expect(list.textContent).toMatch(/gh-agenda/);
    });

    test('counts the ones not wired to a content type, since those are the ones to register', async () => {
        await runCheck([
            { name: 'article', used: true },
            { name: 'gh-agenda', used: false },
            { name: 'spare', used: false }
        ]);

        expect(document.querySelector('.schema-list summary').textContent).toMatch(/2 not wired up/);
    });

    test('marks in-use schemas so they stand apart from available ones', async () => {
        await runCheck([
            { name: 'article', used: true },
            { name: 'gh-agenda', used: false }
        ]);

        const rows = [...document.querySelectorAll('.schema-row')];
        const used = rows.find(r => r.textContent.includes('article'));
        const free = rows.find(r => r.textContent.includes('gh-agenda'));
        expect(used.classList.contains('schema-used')).toBe(true);
        expect(free.classList.contains('schema-used')).toBe(false);
    });

    test('feeds the schema-slug autocomplete, so the field offers real slugs', async () => {
        await runCheck([
            { name: 'article', used: true },
            { name: 'gh-agenda', used: false }
        ]);

        const options = [...document.querySelectorAll('#schemaOptions option')].map(o => o.value);
        expect(options).toEqual(['article', 'gh-agenda']);
        // The manual form's slug field points at that list.
        expect(document.getElementById('manualTypeEndpoint').getAttribute('list')).toBe('schemaOptions');
    });

    test('renders nothing extra when the site reports no schemas', async () => {
        await runCheck([]);
        expect(document.querySelector('.schema-list')).toBeNull();
    });

    test('a failed connection shows no schema list', async () => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/check-connection')) {
                return jsonOk({ connection: { ok: false, message: 'Token rejected' }, contentTypes: [], referenceData: [], schemas: [] });
            }
            return jsonOk({});
        });
        await loadPage();
        document.getElementById('inputURL').value = 's';
        document.getElementById('inputToken').value = 't';
        document.getElementById('testConnectionBtn').click();
        await flushPromises();

        expect(document.querySelector('.schema-list')).toBeNull();
    });
});

describe('registering a content type by hand (no workbook needed)', () => {
    test('the register button is visible on a fresh page, with no file chosen', async () => {
        await loadPage();
        const btn = document.getElementById('addContentTypeBtn');
        expect(btn).not.toBeNull();
        expect(btn.offsetParent === null && btn.hidden).toBe(false);
        // The form itself stays collapsed until asked for.
        expect(document.getElementById('manualTypeForm').hidden).toBe(true);
    });

    test('clicking it reveals the form and flips the button to Cancel', async () => {
        await loadPage();
        document.getElementById('addContentTypeBtn').click();

        expect(document.getElementById('manualTypeForm').hidden).toBe(false);
        expect(document.getElementById('addContentTypeBtn').textContent).toMatch(/cancel/i);
    });

    test('Cancel collapses the form again', async () => {
        await loadPage();
        document.getElementById('addContentTypeBtn').click();
        document.getElementById('manualTypeCancel').click();

        expect(document.getElementById('manualTypeForm').hidden).toBe(true);
        expect(document.getElementById('addContentTypeBtn').textContent).toMatch(/register/i);
    });

    test('missing fields show an inline error and send nothing', async () => {
        await loadPage();
        document.getElementById('addContentTypeBtn').click();
        const before = global.fetch.mock.calls.length;

        document.getElementById('manualTypeLabel').value = 'Agendas';
        // slug and sheet left blank
        document.getElementById('manualTypeSave').click();
        await flushPromises();

        const error = document.getElementById('manualTypeError');
        expect(error.hidden).toBe(false);
        expect(error.textContent).toMatch(/required/i);
        expect(global.fetch.mock.calls.length).toBe(before);
    });

    test('a valid submission posts the type with no workbook loaded, and no columns', async () => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/custom-types')) return jsonOk({ message: 'Saved "Agendas"', type: 'custom_agendas' });
            return jsonOk({});
        });

        await loadPage();
        document.getElementById('addContentTypeBtn').click();
        document.getElementById('manualTypeLabel').value = 'Agendas';
        document.getElementById('manualTypeEndpoint').value = 'gh-agenda';
        document.getElementById('manualTypeSheet').value = 'Agendas and Minutes';
        document.getElementById('manualTypeSave').click();
        await flushPromises();

        const saveCall = global.fetch.mock.calls.find(([u]) => String(u).includes('/custom-types'));
        expect(JSON.parse(saveCall[1].body)).toEqual({
            label: 'Agendas',
            endpoint: 'gh-agenda',
            sheetName: 'Agendas and Minutes'
        });
        // Form clears and collapses on success.
        expect(document.getElementById('manualTypeForm').hidden).toBe(true);
        expect(document.getElementById('manualTypeLabel').value).toBe('');
    });

    test('a server rejection stays inline and leaves the form open to fix', async () => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (String(url).includes('/custom-types')) return jsonErr({ message: 'collides with a built-in content type' });
            return jsonOk({});
        });

        await loadPage();
        document.getElementById('addContentTypeBtn').click();
        document.getElementById('manualTypeLabel').value = 'Articles';
        document.getElementById('manualTypeEndpoint').value = 'x';
        document.getElementById('manualTypeSheet').value = 'y';
        document.getElementById('manualTypeSave').click();
        await flushPromises();

        expect(document.getElementById('manualTypeError').textContent).toMatch(/collides/);
        expect(document.getElementById('manualTypeForm').hidden).toBe(false);
        expect(document.getElementById('manualTypeSave').disabled).toBe(false);
    });

    test('a registered type appears in the Registered list with the sheet it reads', async () => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) {
                return jsonOk({ types: [...BUILT_IN_TYPES, { type: 'custom_agendas', label: 'Agendas', description: 'Custom', sheetName: 'Agendas and Minutes', isCustom: true }] });
            }
            return jsonOk({});
        });

        await loadPage();

        expect(document.getElementById('savedCustomTypes').hidden).toBe(false);
        expect(document.getElementById('savedCustomList').textContent).toMatch(/Agendas → Agendas and Minutes/);
    });

    test('removing a registered type calls DELETE and refreshes', async () => {
        const deleted = [];
        global.fetch = jest.fn((url, opts) => {
            if (String(url).includes('/content-types')) {
                return jsonOk({ types: [...BUILT_IN_TYPES, { type: 'custom_agendas', label: 'Agendas', description: 'C', sheetName: 'S', isCustom: true }] });
            }
            if (String(url).includes('/custom-types/') && opts?.method === 'DELETE') {
                deleted.push(String(url));
                return jsonOk({ message: 'Removed' });
            }
            return jsonOk({});
        });

        await loadPage();
        document.querySelector('.remove-custom-btn').click();
        await flushPromises();

        expect(deleted).toEqual(['/custom-types/custom_agendas']);
    });
});

describe('the token field', () => {
    test('is masked, so it is not readable over someone\'s shoulder', () => {
        document.documentElement.innerHTML = html;
        expect(document.getElementById('inputToken').getAttribute('type')).toBe('password');
    });
});
