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
//
// A folder can hold thousands of files, so it's sent in batches: one
// request per batch, one batch at a time, so the server only ever holds
// one batch in memory and a failure only costs that batch. Per-file
// settings come from the workbook's Assets sheet when there is one.

const inputFolder = document.getElementById('inputFolder');
const folderFileCount = document.getElementById('folderFileCount');
const inputParentId = document.getElementById('inputParentId');
const inputCategories = document.getElementById('inputCategories');
const inputPermissionSet = document.getElementById('inputPermissionSet');
const inputAltText = document.getElementById('inputAltText');
const inputTags = document.getElementById('inputTags');
const inputBatchSize = document.getElementById('inputBatchSize');
const inputPublish = document.getElementById('inputPublish');
const inputOnlyListed = document.getElementById('inputOnlyListed');
const inputAllowDuplicates = document.getElementById('inputAllowDuplicates');
const onlyListedLabel = document.getElementById('onlyListedLabel');
const assetSheetStatus = document.getElementById('assetSheetStatus');
const uploadAssetsBtn = document.getElementById('uploadAssetsBtn');
const stopAssetsBtn = document.getElementById('stopAssetsBtn');
const retryAssetsBtn = document.getElementById('retryAssetsBtn');
const assetTemplateBtn = document.getElementById('assetTemplateBtn');
const assetProgress = document.getElementById('assetProgress');
const assetProgressBar = document.getElementById('assetProgressBar');
const assetProgressText = document.getElementById('assetProgressText');
const assetIssues = document.getElementById('assetIssues');

// Server-side cap on files per request (ASSET_FILES_PER_REQUEST in index.js),
// and a soft cap on bytes per request so a batch of large videos doesn't
// get buffered all at once — a batch closes at whichever comes first.
const MAX_FILES_PER_BATCH = 50;
const MAX_BYTES_PER_BATCH = 150 * 1024 * 1024;

// Rows from the workbook's Assets sheet, or null when there isn't one.
let assetSheetRows = null;
let stopRequested = false;
let failedAssetFiles = [];

// Path of a file inside the chosen folder, without the folder's own name
// ("Photos/sub/a.jpg" -> "sub/a.jpg"), lowercased for matching.
function relativeAssetPath(file) {
    const rel = file.webkitRelativePath || file.name;
    const parts = rel.split('/');
    return (parts.length > 1 ? parts.slice(1).join('/') : rel).toLowerCase();
}

// Matches a file to its Assets-sheet row: by path within the folder first
// (so "sub/a.jpg" can be told apart from "a.jpg"), then by bare file name.
function buildAssetRowIndex(rows) {
    const index = new Map();
    for (const row of rows || []) {
        index.set(String(row.fileName).trim().replace(/\\/g, '/').toLowerCase(), row);
    }
    return index;
}

function findAssetRow(index, file) {
    return index.get(relativeAssetPath(file)) || index.get(file.name.toLowerCase()) || null;
}

function updateFolderSummary() {
    const files = inputFolder.files || [];
    if (!files.length) {
        folderFileCount.innerText = '';
        return;
    }
    let text = `Selected: ${files.length} file(s)`;
    if (assetSheetRows) {
        const index = buildAssetRowIndex(assetSheetRows);
        const matched = Array.from(files).filter(f => findAssetRow(index, f)).length;
        text += ` — ${matched} have a row in the Assets sheet`;
    }
    folderFileCount.innerText = text;
}

inputFolder.addEventListener('change', updateFolderSummary);

async function loadAssetSheet() {
    assetSheetRows = null;
    assetSheetStatus.hidden = true;
    onlyListedLabel.hidden = true;

    if (inputFile.files && inputFile.files.length) {
        const formData = new FormData();
        formData.append('file', inputFile.files[0]);
        try {
            const response = await fetch('/inspect-assets-sheet', { method: 'POST', body: formData });
            const data = await response.json();
            if (response.ok && data.found) {
                assetSheetRows = data.rows || [];
                assetSheetStatus.innerText = `Assets sheet found: ${assetSheetRows.length} file row(s). Matching files use those settings.`;
                assetSheetStatus.hidden = false;
                onlyListedLabel.hidden = false;
            }
        } catch (_) { /* no Assets sheet support for this workbook — defaults only */ }
    }
    updateFolderSummary();
}

inputFile.addEventListener('change', loadAssetSheet);

// Splits files into batches by count and total size.
function chunkAssetFiles(files, batchSize) {
    const batches = [];
    let current = [];
    let bytes = 0;
    for (const item of files) {
        const size = item.file.size || 0;
        if (current.length && (current.length >= batchSize || bytes + size > MAX_BYTES_PER_BATCH)) {
            batches.push(current);
            current = [];
            bytes = 0;
        }
        current.push(item);
        bytes += size;
    }
    if (current.length) batches.push(current);
    return batches;
}

function setAssetProgress(done, total, text) {
    assetProgress.hidden = false;
    assetProgressBar.style.width = total ? `${Math.round((done / total) * 100)}%` : '0%';
    assetProgressText.innerText = text;
}

function showAssetIssues(lines) {
    assetIssues.innerHTML = '';
    const shown = lines.slice(0, 50);
    for (const line of shown) {
        const li = document.createElement('li');
        li.textContent = line;
        assetIssues.appendChild(li);
    }
    if (lines.length > shown.length) {
        const li = document.createElement('li');
        li.textContent = `…and ${lines.length - shown.length} more (see the server console).`;
        assetIssues.appendChild(li);
    }
    assetIssues.hidden = lines.length === 0;
}

// While a batch is in flight the dashboard asks the server every few seconds
// what it knows about it. If the server still hasn't seen the request after
// notReceivedMs, the browser never delivered it — it's cancelled and sent
// again (up to maxSends times in all) instead of waiting forever. If the
// server finished the batch but the response never made it back, the
// result is taken from the status instead.
const ASSET_WATCHDOG = { pollMs: 3000, notReceivedMs: 60 * 1000, maxSends: 3 };

const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function fetchBatchStatus(batchId) {
    try {
        const response = await fetch(`/upload/assets/status/${encodeURIComponent(batchId)}`);
        if (!response.ok) return { state: 'unreachable' };
        return await response.json();
    } catch (_) {
        return { state: 'unreachable' };
    }
}

// One attempt at sending a batch, watched by polling. Resolves to one of
//   { kind: 'response', status, data }  — a real answer for this batch
//   { kind: 'notReceived' }              — the server never got it; resend
//   { kind: 'stopped' }                  — Stop pressed before it arrived
//   { kind: 'lost' }                     — the server can't be reached at all
async function watchAssetBatch(formData, onServerProgress) {
    const batchId = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
    const controller = new AbortController();

    let settled = null;
    fetch(`/upload/assets?batchId=${encodeURIComponent(batchId)}`, { method: 'POST', body: formData, signal: controller.signal })
        .then(async response => {
            let data = {};
            try { data = await response.json(); } catch (_) { /* not JSON */ }
            settled = { kind: 'response', status: response.status, data };
        })
        .catch(() => { settled = settled || { kind: 'fetchFailed' }; });

    const sentAt = Date.now();
    let lastPoll = sentAt;
    for (;;) {
        await wait(250);
        if (settled && settled.kind === 'response') return settled;
        if (Date.now() - lastPoll < ASSET_WATCHDOG.pollMs && !settled && !stopRequested) continue;
        lastPoll = Date.now();

        const status = await fetchBatchStatus(batchId);
        if (settled && settled.kind === 'response') return settled;

        if (status.state === 'finished' && status.result) {
            controller.abort();
            return { kind: 'response', status: status.statusCode || 200, data: status.result };
        }
        if (status.state === 'working') {
            onServerProgress(status.done || 0, status.total || 0);
            continue;
        }
        if (status.state === 'receiving') continue;

        // Not on the server (or the server can't be reached).
        if (settled && settled.kind === 'fetchFailed') {
            return status.state === 'unreachable' ? { kind: 'lost' } : { kind: 'notReceived' };
        }
        if (stopRequested) {
            controller.abort();
            return { kind: 'stopped' };
        }
        if (Date.now() - sentAt >= ASSET_WATCHDOG.notReceivedMs) {
            controller.abort();
            return status.state === 'unreachable' ? { kind: 'lost' } : { kind: 'notReceived' };
        }
    }
}

// Sends one batch. Resolves to { uploaded, duplicates, failedFiles, errors,
// warnings, notes } — or { stopped: true } if Stop was pressed before the
// server received it. A batch that can't be delivered counts every file in
// it as failed.
async function sendAssetBatch(batch, batchLabel, onServerProgress = () => {}) {
    const formData = new FormData();
    for (const item of batch) {
        formData.append('files', item.file, item.file.name);
    }
    formData.append('fileSettings', JSON.stringify(batch.map(item => item.row || {})));
    formData.append('url', inputURL.value);
    formData.append('apiKey', inputToken.value);
    formData.append('parentId', inputParentId.value);
    formData.append('categories', inputCategories.value);
    formData.append('permissionSet', inputPermissionSet.value);
    formData.append('altText', inputAltText.value);
    formData.append('tags', inputTags.value);
    formData.append('publish', inputPublish.checked);
    formData.append('allowDuplicates', inputAllowDuplicates.checked);
    formData.append('batchLabel', batchLabel);

    const names = batch.map(item => item.file.name);
    const failAll = (message, notes = []) =>
        ({ uploaded: 0, duplicates: 0, failedFiles: names, errors: [`${batchLabel} failed: ${message}`], warnings: [], notes });
    const notes = [];

    for (let send = 1; send <= ASSET_WATCHDOG.maxSends; send++) {
        const outcome = await watchAssetBatch(formData, onServerProgress);

        if (outcome.kind === 'stopped') return { stopped: true };
        if (outcome.kind === 'lost') {
            return failAll('lost contact with the uploader server (is it still running?)', notes);
        }
        if (outcome.kind === 'notReceived') {
            notes.push(`${batchLabel}: the server never received it (send ${send} of ${ASSET_WATCHDOG.maxSends})${send < ASSET_WATCHDOG.maxSends ? ' — sent again' : ''}`);
            continue;
        }

        const { status, data } = outcome;
        if (status === 200 || status === 207) {
            return {
                uploaded: data.uploaded ?? (status === 200 ? batch.length : 0),
                duplicates: data.duplicates || 0,
                failedFiles: data.failedFiles || [],
                errors: data.errors || [],
                warnings: data.warnings || [],
                notes
            };
        }
        return failAll(data.message || `HTTP ${status}`, notes);
    }
    return failAll(`the browser could not deliver it after ${ASSET_WATCHDOG.maxSends} tries`, notes);
}

async function runAssetUpload(items) {
    const requested = parseInt(inputBatchSize.value, 10);
    const batchSize = Math.min(Math.max(Number.isFinite(requested) ? requested : 25, 1), MAX_FILES_PER_BATCH);
    const batches = chunkAssetFiles(items, batchSize);

    stopRequested = false;
    uploadAssetsBtn.disabled = true;
    retryAssetsBtn.hidden = true;
    stopAssetsBtn.hidden = false;
    stopAssetsBtn.disabled = false;
    showAssetIssues([]);

    let done = 0;
    let uploaded = 0;
    let duplicates = 0;
    const failedNames = new Set();
    const errors = [];
    const warnings = [];
    const notes = [];

    resultsField.innerText = `Uploading ${items.length} file(s) in ${batches.length} batch(es)...`;
    setAssetProgress(0, items.length, `0 of ${items.length}`);

    for (let b = 0; b < batches.length; b++) {
        if (stopRequested) break;
        const batch = batches[b];
        const batchLabel = `Batch ${b + 1} of ${batches.length}`;
        const status = () => `${batchLabel} — ${done} of ${items.length} processed, ${uploaded} uploaded`
            + (duplicates ? `, ${duplicates} already on site` : '');

        // A ticking clock on the batch in flight, so a slow batch is
        // visibly still working rather than looking frozen.
        const batchStart = Date.now();
        let serverProgress = '';
        setAssetProgress(done, items.length, status());
        const ticker = setInterval(() => {
            const secs = Math.round((Date.now() - batchStart) / 1000);
            setAssetProgress(done, items.length, `${status()} — working on this batch for ${secs}s${serverProgress}`);
        }, 1000);

        let result;
        try {
            result = await sendAssetBatch(batch, batchLabel, (fileDone, fileTotal) => {
                serverProgress = ` (server has handled ${fileDone} of ${fileTotal})`;
            });
        } finally {
            clearInterval(ticker);
        }
        if (result.stopped) break;
        notes.push(...(result.notes || []));
        done += batch.length;
        uploaded += result.uploaded;
        duplicates += result.duplicates || 0;
        result.failedFiles.forEach(name => failedNames.add(name));
        errors.push(...result.errors);
        warnings.push(...result.warnings);

        setAssetProgress(done, items.length, status());
    }

    // Keep the actual File objects for a one-click retry of just the failures.
    failedAssetFiles = items.filter(item => failedNames.has(item.file.name));
    const skipped = items.length - done;

    let summary = `Uploaded ${uploaded} of ${items.length} file(s).`;
    if (duplicates) summary += ` ${duplicates} already existed on the site, so no new copy was made (tick "Upload duplicates as new copies" to force them).`;
    if (failedAssetFiles.length) summary += ` ${failedAssetFiles.length} failed.`;
    if (skipped) summary += ` Stopped — ${skipped} not sent.`;
    if (warnings.length) summary += ` ${warnings.length} warning(s).`;
    resultsField.innerText = summary;
    setAssetProgress(done, items.length, summary);
    showAssetIssues([...errors, ...notes, ...warnings]);

    stopAssetsBtn.hidden = true;
    uploadAssetsBtn.disabled = false;
    if (failedAssetFiles.length) {
        retryAssetsBtn.innerText = `Retry ${failedAssetFiles.length} failed`;
        retryAssetsBtn.hidden = false;
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

    const index = buildAssetRowIndex(assetSheetRows);
    let items = Array.from(inputFolder.files).map(file => ({ file, row: findAssetRow(index, file) }));
    if (assetSheetRows && inputOnlyListed.checked) {
        items = items.filter(item => item.row);
    }
    if (!items.length) {
        resultsField.innerText = 'None of the files in this folder are listed in the Assets sheet.';
        return;
    }

    await runAssetUpload(items);
});

// --- Old-site link report ---
//
// Read-only scan of the site's content for links that still point at the
// old website's files. The server returns the report as an .xlsx (saved
// straight to Downloads) with the headline counts in a header.

const inputOldDomains = document.getElementById('inputOldDomains');
const linkReportBtn = document.getElementById('linkReportBtn');
const linkReportSummary = document.getElementById('linkReportSummary');

function saveBlob(blob, name) {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
}

linkReportBtn.addEventListener('click', async () => {
    if (!inputURL.value || !inputToken.value) {
        resultsField.innerText = 'Website URL and token are required.';
        return;
    }
    if (!inputOldDomains.value.trim()) {
        resultsField.innerText = 'Enter the old website\'s domain first (for example co.monmouth.nj.us).';
        return;
    }

    linkReportBtn.disabled = true;
    linkReportSummary.hidden = true;
    const started = Date.now();
    const ticker = setInterval(() => {
        resultsField.innerText = `Scanning the site's content for old-site links... ${Math.round((Date.now() - started) / 1000)}s`;
    }, 1000);
    resultsField.innerText = 'Scanning the site\'s content for old-site links...';

    try {
        const response = await fetch('/old-link-report', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url: inputURL.value, apiKey: inputToken.value, oldDomains: inputOldDomains.value })
        });
        if (response.status === 404) {
            // The page is newer than the running server (index.js changed
            // but the server wasn't restarted), so the route doesn't exist yet.
            resultsField.innerText = 'Error: the running server is an older version without the link report. Restart it (run the .bat again) and try again.';
            return;
        }
        if (!response.ok) {
            let message = response.statusText;
            try { message = (await response.json()).message || message; } catch (_) { /* not JSON */ }
            resultsField.innerText = `Error: ${message}`;
            return;
        }

        let summary = null;
        try { summary = JSON.parse(decodeURIComponent(response.headers.get('X-Report-Summary') || '')); } catch (_) { /* no summary */ }
        const disposition = response.headers.get('Content-Disposition') || '';
        const name = (disposition.match(/filename="([^"]+)"/) || [])[1] || 'OldLinkReport.xlsx';
        saveBlob(await response.blob(), name);

        if (summary) {
            linkReportSummary.innerText =
                `${summary.fileLinks} link(s) to old-site files in ${summary.itemsWithFileLinks} content item(s) ` +
                `(${summary.itemsScanned} items scanned). ` +
                `${summary.fileLinksWithNoMatch} have no matching file in Evolve, ` +
                `${summary.fileLinksWithSimilarMatch} have a similarly named file, ` +
                `${summary.fileLinksWithExactMatch} have a file with the same name. ` +
                `${summary.pageLinks} link(s) to old-site pages are on a separate sheet.` +
                (summary.schemaErrors && summary.schemaErrors.length ? ` ${summary.schemaErrors.length} content type(s) could not be scanned — see the Summary sheet.` : '');
            linkReportSummary.hidden = false;
        }
        resultsField.innerText = `Report saved as ${name}.`;
    } catch (error) {
        resultsField.innerText = 'Network error while building the report.';
    } finally {
        clearInterval(ticker);
        linkReportBtn.disabled = false;
    }
});

stopAssetsBtn.addEventListener('click', () => {
    stopRequested = true;
    stopAssetsBtn.disabled = true;
    resultsField.innerText = 'Stopping — the batch in progress finishes first (or is cancelled if the server never received it)...';
});

retryAssetsBtn.addEventListener('click', async () => {
    if (!failedAssetFiles.length) return;
    await runAssetUpload(failedAssetFiles);
});

// Downloads an Assets sheet with one row per file in the chosen folder,
// pre-filled with the form's current values, ready to edit and paste into
// the workbook.
assetTemplateBtn.addEventListener('click', async () => {
    if (!inputFolder.files.length) {
        resultsField.innerText = 'Choose a folder first, then download its Assets sheet.';
        return;
    }
    const fileNames = Array.from(inputFolder.files).map(file => {
        const rel = file.webkitRelativePath || file.name;
        const parts = rel.split('/');
        return parts.length > 1 ? parts.slice(1).join('/') : rel;
    });

    try {
        const response = await fetch('/assets/template', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                fileNames,
                defaults: {
                    parentId: inputParentId.value,
                    permissionSet: inputPermissionSet.value,
                    categories: inputCategories.value,
                    altText: inputAltText.value,
                    tags: inputTags.value,
                    publish: inputPublish.checked
                }
            })
        });
        if (!response.ok) {
            let message = response.statusText;
            try { message = (await response.json()).message || message; } catch (_) { /* not JSON */ }
            resultsField.innerText = `Error: ${message}`;
            return;
        }
        const blob = await response.blob();
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'AssetsSheet.xlsx';
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(link.href);
        resultsField.innerText = `Downloaded an Assets sheet with ${fileNames.length} row(s).`;
    } catch (error) {
        resultsField.innerText = 'Network error while building the Assets sheet.';
    }
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


loadContentTypes();
