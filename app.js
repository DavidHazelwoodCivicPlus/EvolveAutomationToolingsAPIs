// Front-end for the Evolve Content Uploader dashboard.
//
// Extracted from an inline <script> in index.html so it can be unit
// tested (see tests/unit/frontend.test.js, which loads this file into a
// jsdom document). Nothing here is bundled or transpiled — it's plain
// browser JS loaded with a <script src>, same as before.

const resultsField = document.getElementById('results');
const inputURL = document.getElementById('inputURL');
const inputToken = document.getElementById('inputToken');
const inputFile = document.getElementById('inputFile');
const fileName = document.getElementById('fileName');
const buttonsDiv = document.getElementById('buttonsDiv');
const customImportsPanel = document.getElementById('customImportsPanel');
const unmatchedSheets = document.getElementById('unmatchedSheets');
const savedCustomTypes = document.getElementById('savedCustomTypes');
const savedCustomList = document.getElementById('savedCustomList');
const addContentTypeBtn = document.getElementById('addContentTypeBtn');
const manualTypeForm = document.getElementById('manualTypeForm');
const manualTypeLabel = document.getElementById('manualTypeLabel');
const manualTypeEndpoint = document.getElementById('manualTypeEndpoint');
const manualTypeSheet = document.getElementById('manualTypeSheet');
const manualTypeSave = document.getElementById('manualTypeSave');
const manualTypeCancel = document.getElementById('manualTypeCancel');
const manualTypeError = document.getElementById('manualTypeError');

// Everything the last workbook inspection told us, so the sheet pickers
// and the custom-import panel can be rebuilt without re-reading the file.
let workbookSheets = [];
// Schema slugs that actually exist on the site, learned from the last
// Test connection. Empty until then — the slug fields just don't
// autocomplete in that case.
let siteSchemas = [];
// type -> sheet name the user explicitly picked, overriding the automatic
// name match. Only populated when someone changes a dropdown.
const sheetOverrides = {};

// --- Content type buttons (built from the server, not hardcoded) ---

async function loadContentTypes() {
    try {
        const response = await fetch('/content-types');
        const data = await response.json();
        renderContentTypes(data.types || []);
        renderSavedCustomTypes((data.types || []).filter(t => t.isCustom));
    } catch (error) {
        resultsField.innerText = 'Could not load content types.';
    }
}

function renderContentTypes(types) {
    buttonsDiv.innerHTML = '';
    types.forEach(t => {
        const row = document.createElement('div');
        row.className = 'type-row';

        const button = document.createElement('button');
        button.className = 'upload-btn';
        button.dataset.endpoint = `/upload/${t.type}`;
        button.dataset.type = t.type;
        button.dataset.name = t.label;
        button.dataset.desc = t.description;
        button.setAttribute('aria-label', `Upload ${t.label} — ${t.description}`);
        button.addEventListener('click', () => uploadContentType(button));
        row.appendChild(button);

        // The sheet picker only makes sense once we know what sheets the
        // chosen workbook actually has.
        if (workbookSheets.length) {
            row.appendChild(buildSheetPicker(t));
        }
        buttonsDiv.appendChild(row);
    });
}

function buildSheetPicker(t) {
    const wrapper = document.createElement('label');
    wrapper.className = 'sheet-picker';
    wrapper.textContent = 'Sheet: ';

    const select = document.createElement('select');
    select.setAttribute('aria-label', `Sheet to read for ${t.label}`);

    const noneOption = document.createElement('option');
    noneOption.value = '';
    noneOption.textContent = '— none —';
    select.appendChild(noneOption);

    workbookSheets.forEach(sheet => {
        const option = document.createElement('option');
        option.value = sheet.name;
        option.textContent = `${sheet.name} (${sheet.rowCount} rows)`;
        select.appendChild(option);
    });

    // Preselect whatever the server's own matching resolved to, so the
    // dropdown always agrees with what an upload would actually do.
    const matched = workbookSheets.find(s => s.matchedType === t.type);
    select.value = sheetOverrides[t.type] || (matched ? matched.name : '');

    select.addEventListener('change', () => {
        sheetOverrides[t.type] = select.value;
    });

    wrapper.appendChild(select);
    return wrapper;
}

// --- Workbook inspection ---

inputFile.addEventListener('change', async () => {
    fileName.innerText = inputFile.files.length
        ? `Selected: ${inputFile.files[0].name}`
        : '';
    if (!inputFile.files.length) {
        workbookSheets = [];
        customImportsPanel.hidden = true;
        await loadContentTypes();
        return;
    }
    await inspectWorkbook();
});

async function inspectWorkbook() {
    const formData = new FormData();
    formData.append('file', inputFile.files[0]);

    try {
        const response = await fetch('/inspect-workbook', { method: 'POST', body: formData });
        const data = await response.json();
        if (!response.ok) {
            resultsField.innerText = `Error: ${data.message}`;
            return;
        }
        workbookSheets = data.sheets || [];
        await loadContentTypes();
        renderUnmatchedSheets(workbookSheets.filter(s => s.unmatched));
    } catch (error) {
        resultsField.innerText = 'Network error while reading the workbook.';
    }
}

// --- Registering a content type by hand ---
//
// Always available, with or without a workbook loaded — the
// unrecognized-sheet shortcut further down is a convenience on top of
// this, not the only way in. Columns aren't asked for here: without a
// workbook there are no headers to read, so the server discovers them
// from each row at upload time.

addContentTypeBtn.addEventListener('click', () => {
    const opening = manualTypeForm.hidden;
    manualTypeForm.hidden = !opening;
    addContentTypeBtn.textContent = opening ? 'Cancel' : '+ Register a content type';
    if (opening) manualTypeLabel.focus();
});

manualTypeCancel.addEventListener('click', () => {
    manualTypeForm.hidden = true;
    manualTypeError.hidden = true;
    addContentTypeBtn.textContent = '+ Register a content type';
});

manualTypeSave.addEventListener('click', async () => {
    manualTypeError.hidden = true;

    const label = manualTypeLabel.value.trim();
    const endpoint = manualTypeEndpoint.value.trim();
    const sheetName = manualTypeSheet.value.trim();

    if (!label || !endpoint || !sheetName) {
        manualTypeError.textContent = 'A display name, schema slug, and sheet name are all required.';
        manualTypeError.hidden = false;
        return;
    }

    manualTypeSave.disabled = true;
    manualTypeSave.textContent = 'Saving…';

    const result = await saveCustomType({ label, endpoint, sheetName });

    manualTypeSave.disabled = false;
    manualTypeSave.textContent = 'Save';

    if (!result.ok) {
        manualTypeError.textContent = result.message;
        manualTypeError.hidden = false;
        return;
    }

    resultsField.innerText = result.message;
    manualTypeLabel.value = '';
    manualTypeEndpoint.value = '';
    manualTypeSheet.value = '';
    manualTypeForm.hidden = true;
    addContentTypeBtn.textContent = '+ Register a content type';
    await refreshAfterTypeChange();
});

// One place that talks to /custom-types, shared by the manual form and
// the unrecognized-sheet shortcut.
async function saveCustomType(definition) {
    try {
        const response = await fetch('/custom-types', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(definition)
        });
        const result = await response.json();
        return { ok: response.ok, message: result.message };
    } catch (error) {
        return { ok: false, message: 'Network error while saving.' };
    }
}

// After adding or removing a type: re-inspect if a workbook is loaded (so
// sheet pickers and the unrecognized list both update), otherwise just
// rebuild the buttons.
async function refreshAfterTypeChange() {
    if (inputFile.files.length) {
        await inspectWorkbook();
    } else {
        await loadContentTypes();
    }
}

// --- Custom imports (sheets that match no known type) ---

function renderUnmatchedSheets(sheets) {
    unmatchedSheets.innerHTML = '';
    customImportsPanel.hidden = sheets.length === 0;
    if (!sheets.length) return;

    sheets.forEach(sheet => {
        const item = document.createElement('div');
        item.className = 'unmatched-sheet';

        const title = document.createElement('div');
        title.className = 'unmatched-name';
        title.textContent = `${sheet.name} (${sheet.rowCount} rows)`;
        item.appendChild(title);

        const cols = document.createElement('div');
        cols.className = 'unmatched-cols';
        cols.textContent = sheet.columns.length
            ? `Columns: ${sheet.columns.join(', ')}`
            : 'No column headers found in this sheet.';
        item.appendChild(cols);

        if (sheet.columns.length) {
            item.appendChild(buildCustomTypeForm(sheet));
        }
        unmatchedSheets.appendChild(item);
    });
}

function buildCustomTypeForm(sheet) {
    const form = document.createElement('div');
    form.className = 'custom-form';

    const labelInput = document.createElement('input');
    labelInput.type = 'text';
    labelInput.placeholder = 'Display name, e.g. Agendas';
    labelInput.setAttribute('aria-label', `Display name for the ${sheet.name} import`);
    labelInput.value = sheet.name;

    const endpointInput = document.createElement('input');
    endpointInput.type = 'text';
    endpointInput.placeholder = 'Schema slug on the site';
    endpointInput.setAttribute('aria-label', `Schema slug for the ${sheet.name} import`);
    // Autocompletes from the real schemas found by Test connection.
    endpointInput.setAttribute('list', 'schemaOptions');

    const saveBtn = document.createElement('button');
    saveBtn.type = 'button';
    saveBtn.className = 'custom-save-btn';
    saveBtn.textContent = 'Add import';

    const error = document.createElement('div');
    error.className = 'custom-error';
    error.hidden = true;

    saveBtn.addEventListener('click', async () => {
        error.hidden = true;
        if (!labelInput.value.trim() || !endpointInput.value.trim()) {
            error.textContent = 'A display name and schema slug are both required.';
            error.hidden = false;
            return;
        }

        saveBtn.disabled = true;
        saveBtn.textContent = 'Saving…';

        // Columns are pinned here: they were just read from the sheet.
        const result = await saveCustomType({
            label: labelInput.value.trim(),
            endpoint: endpointInput.value.trim(),
            sheetName: sheet.name,
            columns: sheet.columns
        });

        if (!result.ok) {
            error.textContent = result.message;
            error.hidden = false;
            saveBtn.disabled = false;
            saveBtn.textContent = 'Add import';
            return;
        }

        resultsField.innerText = result.message;
        // Re-inspect so the new type appears as a button and this sheet
        // drops out of the unrecognized list.
        await inspectWorkbook();
    });

    form.append(labelInput, endpointInput, saveBtn, error);
    return form;
}

function renderSavedCustomTypes(customTypes) {
    savedCustomList.innerHTML = '';
    savedCustomTypes.hidden = customTypes.length === 0;
    if (!customTypes.length) return;

    customTypes.forEach(t => {
        const row = document.createElement('div');
        row.className = 'saved-custom-row';

        const name = document.createElement('span');
        name.textContent = `${t.label} → ${t.sheetName}`;
        row.appendChild(name);

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'remove-custom-btn';
        removeBtn.textContent = 'Remove';
        removeBtn.setAttribute('aria-label', `Remove the ${t.label} custom import`);
        removeBtn.addEventListener('click', async () => {
            removeBtn.disabled = true;
            try {
                const response = await fetch(`/custom-types/${t.type}`, { method: 'DELETE' });
                const result = await response.json();
                resultsField.innerText = result.message;
                await refreshAfterTypeChange();
            } catch (error) {
                removeBtn.disabled = false;
                resultsField.innerText = 'Network error while removing the custom import.';
            }
        });
        row.appendChild(removeBtn);

        savedCustomList.appendChild(row);
    });
}

// --- Folder-of-documents (assets) upload ---

const inputFolder = document.getElementById('inputFolder');
const folderFileCount = document.getElementById('folderFileCount');
const inputParentId = document.getElementById('inputParentId');
const inputCategories = document.getElementById('inputCategories');
const inputPermissionSet = document.getElementById('inputPermissionSet');
const inputPublish = document.getElementById('inputPublish');
const uploadAssetsBtn = document.getElementById('uploadAssetsBtn');

// webkitdirectory hands back every file in the folder (and subfolders)
// as a flat FileList — just show how many were found.
inputFolder.addEventListener('change', () => {
    folderFileCount.innerText = inputFolder.files.length
        ? `Selected: ${inputFolder.files.length} file(s)`
        : '';
});

// --- Test connection ---
// Read-only: checks the site is reachable, the token is accepted, which
// content types have a matching schema on this site, and whether the
// PermissionSet/Categories lists load. Never changes anything.

const testConnectionBtn = document.getElementById('testConnectionBtn');
const connectionCheckResult = document.getElementById('connectionCheckResult');

function checkRow(ok, label, detail) {
    const row = document.createElement('div');
    row.className = 'check-item ' + (ok ? 'check-ok' : 'check-fail');
    const icon = document.createElement('span');
    icon.className = 'check-icon';
    icon.textContent = ok ? '\u2713' : '\u2717';
    const text = document.createElement('span');
    text.className = 'check-label';
    text.textContent = label;
    row.append(icon, text);
    if (detail) {
        const detailSpan = document.createElement('span');
        detailSpan.className = 'check-detail';
        detailSpan.textContent = detail;
        row.append(detailSpan);
    }
    return row;
}

function subheading(text) {
    const el = document.createElement('div');
    el.className = 'check-subheading';
    el.textContent = text;
    return el;
}

function renderConnectionReport(report) {
    connectionCheckResult.innerHTML = '';
    connectionCheckResult.appendChild(checkRow(report.connection.ok, report.connection.message));

    if (!report.connection.ok) return;

    connectionCheckResult.appendChild(subheading('Content types'));
    report.contentTypes.forEach(ct => {
        connectionCheckResult.appendChild(
            checkRow(ct.ok, ct.label, ct.ok ? '' : 'no matching schema found on this site')
        );
    });

    connectionCheckResult.appendChild(subheading('Reference data'));
    report.referenceData.forEach(rd => {
        const label = rd.endpoint === 'permissionSet' ? 'Permission sets' : 'Categories';
        connectionCheckResult.appendChild(
            checkRow(rd.ok, label, rd.ok ? `${rd.count} found` : rd.message)
        );
    });

    renderSchemaList(report.schemas || []);
}

// Every schema that exists on the site, in a collapsed <details> so it
// doesn't bury the pass/fail summary above it. Also feeds the schema-slug
// autocomplete, which is the main reason to want this list.
function renderSchemaList(schemas) {
    siteSchemas = schemas.map(s => s.name);
    updateSchemaOptions();

    if (!schemas.length) return;

    const details = document.createElement('details');
    details.className = 'schema-list';

    const summary = document.createElement('summary');
    const unusedCount = schemas.filter(s => !s.used).length;
    summary.textContent = `Schemas on this site (${schemas.length})`;
    if (unusedCount) summary.textContent += ` — ${unusedCount} not wired up`;
    details.appendChild(summary);

    schemas.forEach(s => {
        const row = document.createElement('div');
        row.className = 'schema-row' + (s.used ? ' schema-used' : '');

        const slug = document.createElement('code');
        slug.textContent = s.name;
        row.appendChild(slug);

        if (s.used) {
            const tag = document.createElement('span');
            tag.className = 'schema-tag';
            tag.textContent = 'in use';
            row.appendChild(tag);
        }
        details.appendChild(row);
    });

    connectionCheckResult.appendChild(details);
}

// Fills the shared <datalist> that every schema-slug input points at, so
// after a Test connection you pick a real slug instead of typing one from
// memory. Created once, on demand.
function updateSchemaOptions() {
    let list = document.getElementById('schemaOptions');
    if (!list) {
        list = document.createElement('datalist');
        list.id = 'schemaOptions';
        document.body.appendChild(list);
    }
    list.innerHTML = '';
    siteSchemas.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        list.appendChild(option);
    });
}

testConnectionBtn.addEventListener('click', async () => {
    if (!inputURL.value || !inputToken.value) {
        resultsField.innerText = 'Website URL and token are required.';
        return;
    }

    testConnectionBtn.disabled = true;
    connectionCheckResult.hidden = false;
    connectionCheckResult.innerHTML = '<p class="checking">Checking connection\u2026</p>';

    try {
        const response = await fetch('/check-connection', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url: inputURL.value, apiKey: inputToken.value })
        });
        const report = await response.json();
        if (response.ok) {
            renderConnectionReport(report);
        } else {
            connectionCheckResult.innerHTML = '';
            connectionCheckResult.appendChild(checkRow(false, report.message || 'Connection check failed.'));
        }
    } catch (error) {
        connectionCheckResult.innerHTML = '';
        connectionCheckResult.appendChild(checkRow(false, 'Network error while checking connection.'));
    } finally {
        testConnectionBtn.disabled = false;
    }
});

// --- Content upload ---

async function uploadContentType(button) {
    const endpoint = button.dataset.endpoint;
    const type = button.dataset.type;

    if (!inputFile.files.length) {
        resultsField.innerText = 'Please choose an Excel file first.';
        return;
    }

    button.disabled = true;
    resultsField.innerText = 'Please wait...';

    const formData = new FormData();
    formData.append('file', inputFile.files[0]);
    formData.append('url', inputURL.value);
    formData.append('apiKey', inputToken.value);
    // Only sent when the user actually picked a different sheet; the
    // server falls back to matching by name when this is absent.
    if (sheetOverrides[type]) {
        formData.append('sheetOverride', sheetOverrides[type]);
    }

    try {
        const response = await fetch(endpoint, {
            method: 'POST',
            // NOTE: do NOT set Content-Type here — the browser adds the
            // correct multipart boundary automatically.
            body: formData
        });

        if (response.ok) {
            const data = await response.json();
            resultsField.innerText = data.message;
        } else {
            let message = response.statusText;
            try {
                const errData = await response.json();
                if (errData && errData.message) message = errData.message;
            } catch (_) { /* response wasn't JSON */ }
            resultsField.innerText = `Error: ${message}`;
        }
    } catch (error) {
        resultsField.innerText = 'Network error occurred.';
    } finally {
        setTimeout(() => {
            button.disabled = false;
        }, 3000);
    }
}

uploadAssetsBtn.addEventListener('click', async () => {
    if (!inputFolder.files.length) {
        resultsField.innerText = 'Please choose a folder first.';
        return;
    }
    if (!inputURL.value || !inputToken.value) {
        resultsField.innerText = 'Website URL and token are required.';
        return;
    }

    uploadAssetsBtn.disabled = true;
    resultsField.innerText = `Uploading ${inputFolder.files.length} file(s)...`;

    // Every file the folder picker found goes in under the same "files"
    // key — multer collects repeated keys into req.files.
    const formData = new FormData();
    for (const file of inputFolder.files) {
        formData.append('files', file);
    }
    formData.append('url', inputURL.value);
    formData.append('apiKey', inputToken.value);
    formData.append('parentId', inputParentId.value);
    formData.append('categories', inputCategories.value);
    formData.append('permissionSet', inputPermissionSet.value);
    formData.append('publish', inputPublish.checked);

    try {
        const response = await fetch('/upload/assets', {
            method: 'POST',
            body: formData
        });

        if (response.ok) {
            const data = await response.json();
            resultsField.innerText = data.message;
        } else {
            let message = response.statusText;
            try {
                const errData = await response.json();
                if (errData && errData.message) message = errData.message;
            } catch (_) { /* response wasn't JSON */ }
            resultsField.innerText = `Error: ${message}`;
        }
    } catch (error) {
        resultsField.innerText = 'Network error occurred.';
    } finally {
        setTimeout(() => {
            uploadAssetsBtn.disabled = false;
        }, 3000);
    }
});

loadContentTypes();
