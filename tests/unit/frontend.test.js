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

describe('folder (assets) upload in batches', () => {
    // A folder's FileList, with webkitRelativePath set the way the
    // directory picker sets it ("Folder/sub/name").
    function attachFolder(names) {
        const files = names.map(rel => {
            const file = new File(['x'], rel.split('/').pop(), { type: 'image/jpeg' });
            Object.defineProperty(file, 'webkitRelativePath', { value: `Photos/${rel}` });
            return file;
        });
        const input = document.getElementById('inputFolder');
        Object.defineProperty(input, 'files', { value: files, configurable: true });
        input.dispatchEvent(new Event('change'));
        return files;
    }

    function fillConnection() {
        document.getElementById('inputURL').value = 'site';
        document.getElementById('inputToken').value = 'token';
    }

    const isUpload = (url) => String(url).startsWith('/upload/assets?');
    const isStatus = (url) => String(url).startsWith('/upload/assets/status/');

    // Answers every batch as fully uploaded unless `onUpload` says
    // otherwise; status polls say "unknown" unless `onStatus` says
    // otherwise; `extra` handles any other URL the test cares about.
    function mockFetch({ onUpload, onStatus, extra } = {}) {
        global.fetch = jest.fn((url, opts) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (isUpload(url)) {
                const custom = onUpload && onUpload(opts, url);
                if (custom) return custom;
                return jsonOk({ uploaded: opts.body.getAll('files').length, failedFiles: [], errors: [], warnings: [] });
            }
            if (isStatus(url)) {
                return (onStatus && onStatus(url)) || jsonOk({ state: 'unknown' });
            }
            const handled = extra && extra(url);
            return handled || jsonOk({});
        });
    }

    // Promise that never settles until the abort signal fires — a request
    // the browser never delivers.
    const neverDelivered = (opts) => new Promise((resolve, reject) => {
        opts.signal.addEventListener('abort', () => reject(new DOMException('aborted', 'AbortError')));
    });

    async function loadAssetSheet() {
        attachFile(document.getElementById('inputFile'));
        document.getElementById('inputFile').dispatchEvent(new Event('change'));
        await flushPromises();
    }

    // The upload loop polls on timers, so runs are driven with fake time.
    async function startUpload(buttonId = 'uploadAssetsBtn') {
        jest.useFakeTimers();
        document.getElementById(buttonId).click();
    }
    const advance = (ms) => jest.advanceTimersByTimeAsync(ms);

    const uploadCalls = () => global.fetch.mock.calls.filter(([url]) => isUpload(url));
    const results = () => document.getElementById('results').innerText;

    test('splits a large folder into sequential batches of the chosen size', async () => {
        mockFetch();
        await loadPage();
        fillConnection();
        attachFolder(Array.from({ length: 60 }, (_, i) => `f${i}.jpg`));
        document.getElementById('inputBatchSize').value = '25';

        await startUpload();
        await advance(3000);

        const sizes = uploadCalls().map(([, opts]) => opts.body.getAll('files').length);
        expect(sizes).toEqual([25, 25, 10]);
        expect(results()).toMatch(/Uploaded 60 of 60/);
    });

    test('every batch gets its own batchId', async () => {
        mockFetch();
        await loadPage();
        fillConnection();
        attachFolder(['a.jpg', 'b.jpg']);
        document.getElementById('inputBatchSize').value = '1';

        await startUpload();
        await advance(2000);

        const ids = uploadCalls().map(([url]) => new URL(url, 'http://x').searchParams.get('batchId'));
        expect(ids).toHaveLength(2);
        expect(ids[0]).toBeTruthy();
        expect(ids[0]).not.toBe(ids[1]);
    });

    test('batch size is capped at the server limit of 50', async () => {
        mockFetch();
        await loadPage();
        fillConnection();
        attachFolder(Array.from({ length: 120 }, (_, i) => `f${i}.jpg`));
        document.getElementById('inputBatchSize').value = '500';

        await startUpload();
        await advance(3000);

        expect(uploadCalls().map(([, opts]) => opts.body.getAll('files').length)).toEqual([50, 50, 20]);
    });

    test('Assets-sheet rows are matched to files (by sub-path, then name) and sent as fileSettings', async () => {
        mockFetch({
            extra: (url) => url === '/inspect-assets-sheet' && jsonOk({ found: true, sheetName: 'Assets', rows: [
                { fileName: 'sub/a.jpg', altText: 'Sub A' },
                { fileName: 'B.JPG', altText: 'Bee', publish: true }
            ] })
        });
        await loadPage();
        fillConnection();
        await loadAssetSheet();

        attachFolder(['sub/a.jpg', 'a.jpg', 'b.jpg', 'c.jpg']);
        expect(document.getElementById('folderFileCount').innerText).toMatch(/2 have a row/);

        await startUpload();
        await advance(1000);

        const settings = JSON.parse(uploadCalls()[0][1].body.get('fileSettings'));
        expect(settings).toEqual([
            { fileName: 'sub/a.jpg', altText: 'Sub A' },
            {},
            { fileName: 'B.JPG', altText: 'Bee', publish: true },
            {}
        ]);
    });

    test('"only listed files" skips files with no Assets-sheet row', async () => {
        mockFetch({
            extra: (url) => url === '/inspect-assets-sheet' && jsonOk({ found: true, rows: [{ fileName: 'b.jpg' }] })
        });
        await loadPage();
        fillConnection();
        await loadAssetSheet();
        attachFolder(['a.jpg', 'b.jpg']);
        document.getElementById('inputOnlyListed').checked = true;

        await startUpload();
        await advance(1000);

        expect(uploadCalls()[0][1].body.getAll('files').map(f => f.name)).toEqual(['b.jpg']);
    });

    test('failed files can be retried on their own', async () => {
        let attempt = 0;
        mockFetch({
            onUpload: () => {
                attempt++;
                if (attempt !== 1) return null;
                return Promise.resolve({ ok: false, status: 207, json: () => Promise.resolve({ uploaded: 1, failedFiles: ['b.jpg'], errors: ['b.jpg: HTTP 500'], warnings: [] }) });
            }
        });
        await loadPage();
        fillConnection();
        attachFolder(['a.jpg', 'b.jpg']);

        await startUpload();
        await advance(1000);

        const retryBtn = document.getElementById('retryAssetsBtn');
        expect(retryBtn.hidden).toBe(false);
        expect(document.getElementById('assetIssues').textContent).toMatch(/b\.jpg: HTTP 500/);

        retryBtn.click();
        await advance(1000);

        expect(uploadCalls()[1][1].body.getAll('files').map(f => f.name)).toEqual(['b.jpg']);
        expect(retryBtn.hidden).toBe(true);
    });

    test('files the site already had are reported as duplicates, not uploads', async () => {
        mockFetch({ onUpload: () => jsonOk({ uploaded: 1, duplicates: 2, failedFiles: [], errors: [], warnings: [] }) });
        await loadPage();
        fillConnection();
        attachFolder(['a.jpg', 'b.jpg', 'c.jpg']);
        document.getElementById('inputAllowDuplicates').checked = true;

        await startUpload();
        await advance(1000);

        expect(uploadCalls()[0][1].body.get('allowDuplicates')).toBe('true');
        expect(uploadCalls()[0][1].body.get('batchLabel')).toBe('Batch 1 of 1');
        expect(results()).toMatch(/Uploaded 1 of 3 file\(s\)\. 2 already existed/);
    });

    test('a slow batch shows a ticking timer and the server\'s own progress', async () => {
        let release;
        mockFetch({
            onUpload: () => new Promise(resolve => {
                release = () => resolve({ ok: true, status: 200, json: () => Promise.resolve({ uploaded: 1, failedFiles: [] }) });
            }),
            onStatus: () => jsonOk({ state: 'working', done: 3, total: 25 })
        });
        await loadPage();
        fillConnection();
        attachFolder(['a.jpg']);

        await startUpload();
        await advance(5000);
        expect(document.getElementById('assetProgressText').innerText)
            .toMatch(/working on this batch for 5s \(server has handled 3 of 25\)/);

        release();
        await advance(1000);
        expect(results()).toMatch(/Uploaded 1 of 1/);
    });

    test('a batch the server never receives is cancelled and sent again', async () => {
        let sends = 0;
        mockFetch({ onUpload: (opts) => (++sends === 1 ? neverDelivered(opts) : null) });
        await loadPage();
        fillConnection();
        attachFolder(['a.jpg', 'b.jpg']);

        await startUpload();
        await advance(59 * 1000);
        expect(uploadCalls()).toHaveLength(1);   // still waiting

        await advance(5 * 1000);
        expect(uploadCalls()).toHaveLength(2);   // gave up on the first, resent
        expect(uploadCalls()[0][1].signal.aborted).toBe(true);

        await advance(1000);
        expect(results()).toMatch(/Uploaded 2 of 2/);
        expect(document.getElementById('assetIssues').textContent).toMatch(/Batch 1 of 1: the server never received it \(send 1 of 3\) — sent again/);
    });

    test('a batch that can never be delivered fails after 3 sends and the run moves on', async () => {
        let sends = 0;
        mockFetch({ onUpload: (opts) => (++sends <= 3 ? neverDelivered(opts) : null) });
        await loadPage();
        fillConnection();
        attachFolder(['a.jpg', 'b.jpg']);
        document.getElementById('inputBatchSize').value = '1';

        await startUpload();
        await advance(4 * 60 * 1000);

        expect(uploadCalls()).toHaveLength(4); // 3 tries at batch 1, then batch 2
        expect(results()).toMatch(/Uploaded 1 of 2 file\(s\)\. 1 failed/);
        expect(document.getElementById('retryAssetsBtn').hidden).toBe(false);
    }, 30000); // four simulated minutes of polling

    test('a batch the server finished, whose response got lost, is recovered from the status', async () => {
        mockFetch({
            onUpload: (opts) => neverDelivered(opts),
            onStatus: () => jsonOk({ state: 'finished', statusCode: 200, result: { uploaded: 2, failedFiles: [] } })
        });
        await loadPage();
        fillConnection();
        attachFolder(['a.jpg', 'b.jpg']);

        await startUpload();
        await advance(5000);

        expect(uploadCalls()).toHaveLength(1); // not resent
        expect(results()).toMatch(/Uploaded 2 of 2/);
    });

    test('Stop cancels a batch the server never received instead of waiting on it', async () => {
        mockFetch({ onUpload: (opts) => neverDelivered(opts) });
        await loadPage();
        fillConnection();
        attachFolder(['a.jpg', 'b.jpg']);
        document.getElementById('inputBatchSize').value = '1';

        await startUpload();
        await advance(2000);
        document.getElementById('stopAssetsBtn').click();
        await advance(1000);

        expect(uploadCalls()).toHaveLength(1);
        expect(results()).toMatch(/Uploaded 0 of 2 file\(s\)\. Stopped — 2 not sent/);
    });

    test('losing the uploader server entirely fails the batch with a clear message', async () => {
        mockFetch({
            onUpload: () => Promise.reject(new TypeError('Failed to fetch')),
            onStatus: () => Promise.reject(new TypeError('Failed to fetch'))
        });
        await loadPage();
        fillConnection();
        attachFolder(['a.jpg']);

        await startUpload();
        await advance(5000);

        expect(document.getElementById('assetIssues').textContent).toMatch(/lost contact with the uploader server/);
    });
});

describe('old-site link report', () => {
    function fill() {
        document.getElementById('inputURL').value = 'cp-examplesites';
        document.getElementById('inputToken').value = 'token';
    }

    test('asks for the old domain before scanning', async () => {
        await loadPage();
        fill();
        const before = global.fetch.mock.calls.length;
        document.getElementById('linkReportBtn').click();
        await flushPromises();
        expect(document.getElementById('results').innerText).toMatch(/old website's domain/);
        expect(global.fetch.mock.calls.length).toBe(before);
    });

    test('sends the domain, saves the .xlsx and shows the counts', async () => {
        const summary = { fileLinks: 12, itemsWithFileLinks: 5, itemsScanned: 400, fileLinksWithNoMatch: 4, fileLinksWithSimilarMatch: 3, fileLinksWithExactMatch: 5, pageLinks: 9, schemaErrors: [] };
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (url === '/old-link-report') {
                return Promise.resolve({
                    ok: true,
                    status: 200,
                    headers: new Headers({
                        'X-Report-Summary': encodeURIComponent(JSON.stringify(summary)),
                        'Content-Disposition': 'attachment; filename="OldLinkReport-cp-examplesites.xlsx"'
                    }),
                    blob: () => Promise.resolve(new Blob(['x']))
                });
            }
            return jsonOk({});
        });
        URL.createObjectURL = jest.fn(() => 'blob:x');
        URL.revokeObjectURL = jest.fn();
        const clicks = jest.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});

        await loadPage();
        fill();
        document.getElementById('inputOldDomains').value = 'co.monmouth.nj.us';
        document.getElementById('linkReportBtn').click();
        await flushPromises();

        const [, opts] = global.fetch.mock.calls.find(([url]) => url === '/old-link-report');
        expect(JSON.parse(opts.body)).toEqual({ url: 'cp-examplesites', apiKey: 'token', oldDomains: 'co.monmouth.nj.us' });
        expect(clicks).toHaveBeenCalled();
        expect(document.getElementById('results').innerText).toMatch(/OldLinkReport-cp-examplesites\.xlsx/);
        const text = document.getElementById('linkReportSummary').innerText;
        expect(text).toMatch(/12 link\(s\) to old-site files in 5 content item\(s\)/);
        expect(text).toMatch(/4 have no matching file in Evolve/);
    });
});

describe('old-site link report against an outdated server', () => {
    test('a 404 says to restart the server', async () => {
        global.fetch = jest.fn((url) => {
            if (String(url).includes('/content-types')) return jsonOk({ types: BUILT_IN_TYPES });
            if (url === '/old-link-report') return Promise.resolve({ ok: false, status: 404, statusText: 'Not Found', json: () => Promise.reject(new Error('html')) });
            return jsonOk({});
        });
        await loadPage();
        document.getElementById('inputURL').value = 's';
        document.getElementById('inputToken').value = 't';
        document.getElementById('inputOldDomains').value = 'co.monmouth.nj.us';
        document.getElementById('linkReportBtn').click();
        await flushPromises();
        expect(document.getElementById('results').innerText).toMatch(/Restart it/);
    });
});
