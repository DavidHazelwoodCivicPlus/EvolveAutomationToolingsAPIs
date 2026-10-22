const XLSX = require('xlsx');
const express = require('express');
const path = require('node:path');
const fs = require('node:fs');
const multer = require('multer');

const app = express();
// 10 MB so the Assets-sheet template request can carry a folder of thousands
// of file names.
app.use(express.json({ limit: '10mb' }));

// Excel files are received in memory (never written to disk) and parsed
// straight from the buffer. 25 MB cap + a light filter for spreadsheet types.
const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 25 * 1024 * 1024 },
    fileFilter: (req, file, cb) => {
        const ok = /\.(xlsx|xlsm|xls|csv)$/i.test(file.originalname);
        cb(ok ? null : new Error('Please upload a .xlsx, .xlsm, .xls or .csv file.'), ok);
    }
});
const uploadSingle = upload.single('file');

app.get('/style.css', (req, res) => res.sendFile(path.join(__dirname, 'style.css')));
app.get('/app.js', (req, res) => res.sendFile(path.join(__dirname, 'app.js')));
app.get('/civicplus-mark.png', (req, res) => res.sendFile(path.join(__dirname, 'civicplus-mark.png')));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

const PORT = process.env.PORT || 4000;

// --- Helper Functions ---

const GUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const parseTags = (tags) => tags ? String(tags).split(',').map(tag => tag.trim()) : undefined;

// Resolve a name (from the spreadsheet) to an id using a name->id lookup Map.
// Falls back to the raw value if it already looks like a GUID, so a sheet can
// mix plain names and raw IDs. Returns undefined (and warns) if nothing matches.
const resolveId = (value, lookup, label = 'value') => {
    const name = String(value ?? '').trim();
    if (!name) return undefined;

    const match = lookup?.get(name.toLowerCase());
    if (match) return match;

    if (GUID_REGEX.test(name)) return name; // already an id, use as-is
    console.warn(`No ${label} match for "${name}" on this site.`);
    return undefined;
};

// Categories can be a comma-separated list of names; map each to { id, name }.
//
// A name with no match on the site is DROPPED, not sent with an empty id.
// Sending one makes the API reject the whole row ("Categories: One or more
// Categories do not exist in App"), losing every other field in it — so a
// row with one bad category now uploads without that category instead of
// not uploading at all. collectMissingCategories() below reports what was
// dropped, so this is visible rather than silent.
const parseCategories = (categories, lookup) => {
    if (!categories) return undefined;
    const resolved = String(categories).split(',')
        .map(cat => cat.trim())
        .filter(Boolean)
        .map(name => ({ id: resolveId(name, lookup, 'category'), name }))
        .filter(cat => cat.id);
    return resolved.length ? resolved : undefined;
};

// Every category name in the sheet that doesn't exist on the site, so the
// upload result can name them instead of leaving the user to infer it
// from a silently missing category.
function collectMissingCategories(rows, lookup) {
    const missing = new Set();
    for (const row of rows) {
        if (!row.Categories) continue;
        String(row.Categories).split(',')
            .map(c => c.trim())
            .filter(Boolean)
            .forEach(name => {
                if (!lookup?.get(name.toLowerCase())) missing.add(name);
            });
    }
    return [...missing];
}

// Turns a (possibly comma-separated) string of NAMES into an array of
// resolved content/asset ids — for Squidex Reference and Asset fields,
// which are always arrays of raw id strings (never { id, name } objects
// like categories above), even for a field that only ever holds one.
// A name that doesn't match anything existing is skipped (warned, not
// included as undefined/null), because Squidex rejects a reference array
// containing anything that isn't a real id.
const resolveReferenceIds = (namesString, lookup, label = 'reference') => {
    if (!namesString) return [];
    return String(namesString).split(',')
        .map(name => name.trim())
        .filter(Boolean)
        .map(name => {
            const id = lookup?.get(name.toLowerCase());
            if (!id) console.warn(`No ${label} found matching "${name}" — leaving it out of this reference (it needs to already exist on the site).`);
            return id;
        })
        .filter(Boolean);
};

// Every sheet carries a PermissionSet NAME column (no id columns anymore),
// so permissionSet is resolved the same way for every content type.
const permissionSetByName = (entry, lookups) => ({
    id: resolveId(entry.PermissionSet, lookups.permissionSet, 'permissionSet'),
    name: entry.PermissionSet
});

// --- Endpoint (schema slug) resolution helpers ---
//
// The same content type is spelled differently per site: staff might be
// "employee", "enhanced-employee" or "enhancedemployee". Rather than hardcode
// one slug, each config lists the slugs it could be and we pick the one that
// actually exists on the target site (from /schemas). Punctuation/casing
// variants are matched automatically, so you only ever list genuinely different
// WORDS in the candidate array.

// Loose comparison used for BOTH slugs and field names: lowercase and strip
// everything but a-z/0-9, so "enhanced-employee"/"enhancedEmployee" collapse to
// one key, and "faxNumber"/"fax_number" do too. Per-site punctuation/casing
// stops mattering; you only list genuinely different WORDS.
const normalizeSlug = (s) => String(s ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');

// Which of this workbook's actual sheets a content type should read from.
// Checked in order: an explicit override (the GUI's per-type sheet picker,
// once a workbook is selected) always wins; then the type's configured
// sheetName, fuzzy-matched the same way schema slugs are (so "Facility",
// "facility", and "FACILITY " all match); then, only if nothing matches by
// name, the old index-based position — a safety net for a workbook whose
// sheet was renamed to something unrecognizable, not the primary path.
// Matching by name instead of position is what would have caught the
// facility/"Agendas and Minutes" mismatch automatically, for any type.
function resolveSheetName(sheetNames, config, override) {
    if (override) {
        const normalized = normalizeSlug(override);
        const match = sheetNames.find(name => normalizeSlug(name) === normalized);
        if (match) return match;
    }
    if (config.sheetName) {
        const normalized = normalizeSlug(config.sheetName);
        const match = sheetNames.find(name => normalizeSlug(name) === normalized);
        if (match) return match;
    }
    return sheetNames[config.sheetIndex];
}

// A type may list several candidate slugs (order = preference).
// Back-compat: a single `endpoint: "..."` string still works.
const getEndpointCandidates = (config) =>
    config.endpoints?.length ? config.endpoints
    : config.endpoint ? [config.endpoint]
    : [];

// Pick the real slug for THIS site from the schemas that actually exist.
// Exact match wins; otherwise fall back to a normalized match. Returns the
// slug as the site spells it, or null if none of the candidates exist. Only
// ever looks up OUR candidate keys in the schema set — it never scans for
// substrings, so it can't silently grab an unrelated schema.
function resolveEndpoint(candidates, schemaNames) {
    for (const c of candidates) {
        if (schemaNames.includes(c)) return c;            // exact
    }
    const byNorm = new Map(schemaNames.map(n => [normalizeSlug(n), n]));
    for (const c of candidates) {
        const hit = byNorm.get(normalizeSlug(c));         // ignore dashes/case
        if (hit) return hit;
    }
    return null;
}

// For error messages ONLY: schemas whose name looks related to what we wanted,
// so a brand-new variant tells you exactly what to add to endpoints[]. This is
// the only place substring matching happens, and it never picks anything — it
// just builds a helpful hint.
function suggestClosestSchemas(candidates, schemaNames, limit = 6) {
    const cores = candidates.map(normalizeSlug);
    return schemaNames
        .map(name => {
            const n = normalizeSlug(name);
            let score = 0;
            for (const core of cores) {
                if (n.includes(core) || core.includes(n)) {
                    score = Math.max(score, Math.min(n.length, core.length));
                }
            }
            return { name, score };
        })
        .filter(s => s.score > 0)
        .sort((a, b) => b.score - a.score)
        .slice(0, limit)
        .map(s => s.name);
}

// --- Field-name resolution helpers ---
//
// Same idea as endpoints, one level deeper: the SAME logical field is named
// differently per site (departments' fax number is "faxnumber" on one site and
// "fax" on another). Each config can declare `fieldAliases`, mapping the name
// the mapper emits to the alternate names other sites use. At upload time we
// read the resolved schema's real field names and rename each data key to
// whatever THIS site actually has. Punctuation/casing variants are matched
// automatically (same normalizeSlug as endpoints), so you only list genuinely
// different WORDS.

// Candidate field names for a data key = the key itself, then its declared alts.
const getFieldCandidates = (config, key) =>
    [key, ...((config.fieldAliases && config.fieldAliases[key]) || [])];

// Pick the real field descriptor for THIS site from the schema's field map
// (keyed by normalized name). Returns { name, partitioning } or null.
function resolveField(candidates, fieldMap) {
    for (const c of candidates) {
        const hit = fieldMap.get(normalizeSlug(c));
        if (hit) return hit;
    }
    return null;
}

// Re-wrap a single-partition value under the partition the schema expects.
// { en: X } <-> { iv: X } only; multi-key or non-standard wrappers are left
// alone so complex/localized fields are never mangled.
function conformPartition(value, partitioning) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return value;
    const keys = Object.keys(value);
    if (keys.length !== 1 || (keys[0] !== 'en' && keys[0] !== 'iv')) return value;

    const wanted = partitioning === 'invariant' ? 'iv' : 'en';
    if (keys[0] === wanted) return value;
    return { [wanted]: value[keys[0]] };
}

// Rename each key of a fields object (data / searchFields / titles) to the field
// name THIS site's schema actually uses, based on the config's fieldAliases and
// the live schema. Optionally also fix the { iv } vs { en } wrapper to match the
// field's partitioning. Keys with no schema match are left untouched (they may
// be injected custom fields, or the API will report them). Returns a new object.
//
// SAFE BY DESIGN: it only renames a key when the schema has a DIFFERENT spelling
// for it. For any site where the upload already succeeds, every key already
// matches, so this is a no-op there. It only changes the failing case.
function conformFields(fieldsObj, fieldMap, config, { conformPartitions } = {}) {
    if (!fieldsObj || !fieldMap) return fieldsObj;
    const out = {};

    for (const [key, value] of Object.entries(fieldsObj)) {
        const match = resolveField(getFieldCandidates(config, key), fieldMap);
        let targetKey = match ? match.name : key;

        // Never clobber a key already present under its real name.
        if (targetKey !== key && Object.prototype.hasOwnProperty.call(out, targetKey)) {
            console.warn(`Field "${key}" resolves to "${targetKey}" but that key already exists — keeping "${key}" as-is.`);
            targetKey = key;
        }
        if (targetKey !== key) {
            console.log(`Field "${key}" -> "${targetKey}" (matched this site's schema).`);
        }

        out[targetKey] = (conformPartitions && match)
            ? conformPartition(value, match.partitioning)
            : value;
    }
    return out;
}

// Reference lists pulled from /api/apps/<site>/<endpoint> for EVERY type. Names
// in each sheet's PermissionSet / Categories columns are matched against these
// to fill in ids. A type can override this (e.g. `referenceData: []` to skip),
// but by default every type loads both lists so a new sheet "just works".
const DEFAULT_REFERENCE_DATA = ["permissionSet", "categories"];

// Each config maps a sheet to an endpoint + payload shape. `standardColumns`
// lists the columns the mapper already handles; ANY column not in that list is
// auto-injected into `data` as a custom field wrapped in { iv: ... }.
//
// ENDPOINT vs ENDPOINTS: use `endpoint: "slug"` when the slug is the same on
// every site. Use `endpoints: ["a", "b"]` (preference order) when different
// sites use different WORDS for the same type. Punctuation/casing variants of a
// single slug are matched automatically.
//
// FIELDALIASES: optional per-type map of { nameTheMapperEmits: ["altName", ...] }
// for when sites name the SAME field differently (e.g. faxnumber vs fax). At
// upload time the data key is renamed to whichever name the resolved schema
// actually has. Only list genuinely different words; casing/punctuation is
// handled for you. Leave it off entirely for types whose fields don't vary.
//
// IMPORTANT: an injected column only works if a field with that EXACT name
// exists in the Squidex/CivicPlus schema AND is invariant (iv-partitioned).
// If it doesn't, the API rejects the whole record. That is exactly what the
// leftover test columns ("Ack" on Resources, "new" on Staff) were doing — they
// aren't real schema fields, so those two content types failed on every row.
// They are now listed in standardColumns so they are ignored instead of sent.
// => The rule is: leave a column OUT of standardColumns ONLY if it is a real
//    invariant custom field in your schema.
const uploadConfig = {
    articles: {
        sheetIndex: 0,
        sheetName: "Articles",
        label: "Articles",
        description: "General pages and blog-style posts",
        endpoint: "article",
        standardColumns: ["Title", "ContentType", "ContentName", "Content", "Publish", "PermissionSet", "Categories", "Tags", "Name"],
        mapPayload: (entry, lookups) => ({
            data: { name: { en: entry.Title }, article: { en: entry.Content } },
            contentTypeName: entry.ContentType,
            contentTypeDisplayName: entry.contentTypeDisplayName,
            permissionSet: permissionSetByName(entry, lookups),
            tags: parseTags(entry.Tags),
            categories: parseCategories(entry.Categories, lookups.categories),
            publish: entry.Publish
        })
    },
    resources: {
        sheetIndex: 1,
        sheetName: "Resources",
        label: "Resources",
        description: "Resource directory entries",
        // Was "resourcedirectory" — that URL 404s. Squidex slugs are kebab-case
        // (see your own "gh-agenda" / "name-of-event"), so the real slug is almost
        // certainly "resource-directory". Resolution confirms it against /schemas,
        // and a site that spells it "resourcedirectory" still matches (normalized).
        endpoint: "resource-directory",
        // "Ack" added: it is a test column, not a schema field. Without this it
        // was injected as { iv: "david" } and the API rejected every row.
        standardColumns: ["Department", "MaskedEmail", "PhoneNumber", "AdditionalNumber", "FaxNumber", "Address1", "Address2", "City", "State", "Zip", "Latitude", "Longitude", "HoursOfOperation", "AdditionalInformation", "Name", "Email", "Fax", "Hours", "WebsiteUrl", "WebsiteDisplayName", "Description", "Publish", "UserName", "PermissionSet", "Categories", "Ack"],
        // Add fieldAliases here the same way departments does if resource field
        // names vary across sites (e.g. faxnumber: ["fax"]).
        mapPayload: (entry, lookups) => ({
            data: {
                department: { en: entry.Department },
                maskedemail: { en: entry.MaskedEmail },
                phonenumber: { en: entry.PhoneNumber },
                additionalnumber: { en: entry.AdditionalNumber },
                faxnumber: { en: entry.FaxNumber },
                location: {
                    iv: {
                        address1: entry.Address1,
                        address2: entry.Address2,
                        city: entry.City,
                        state: entry.State,
                        zip: entry.Zip,
                    }
                },
                hoursofoperation: { en: entry.HoursOfOperation },
                additionalinformation: { en: entry.AdditionalInformation },
                Name: { en: entry.Name },
                email: { en: entry.Email },
                PhoneNumber: { en: entry.PhoneNumber },
                Fax: { en: entry.Fax },
                Address: {
                    iv: {
                        address1: entry.Address1,
                        address2: entry.Address2,
                        city: entry.City,
                        state: entry.State,
                        zip: entry.Zip,
                    }
                },
                Hours: { en: entry.Hours },
                Website: {
                    en: {
                        openInNewWindow: true,
                        url: entry.WebsiteUrl,
                        displayName: entry.WebsiteDisplayName
                    }
                },
                Description: { en: entry.Description }
            },
            titles: { department: { en: entry.Department } },
            searchFields: {
                department: { en: entry.Department },
                maskedemail: { en: entry.MaskedEmail },
                phonenumber: { en: entry.PhoneNumber },
                additionalnumber: { en: entry.AdditionalNumber },
                faxnumber: { en: entry.FaxNumber },
                location: {
                    iv: {
                        address1: entry.Address1,
                        address2: entry.Address2,
                        city: entry.City,
                        state: entry.State,
                        zip: entry.Zip,
                    }
                },
                hoursofoperation: { en: entry.HoursOfOperation },
                additionalinformation: { en: entry.AdditionalInformation },
                Name: { en: entry.Name },
                email: { en: entry.Email },
                PhoneNumber: { en: entry.PhoneNumber },
                Fax: { en: entry.Fax },
                Address: {
                    iv: {
                        address1: entry.Address1,
                        address2: entry.Address2,
                        city: entry.City,
                        state: entry.State,
                        zip: entry.Zip,
                    }
                },
                Hours: { en: entry.Hours },
                Website: {
                    en: {
                        openInNewWindow: true,
                        url: entry.WebsiteUrl,
                        displayName: entry.WebsiteDisplayName
                    }
                },
                Description: { en: entry.Description }
            },
            permissionSet: permissionSetByName(entry, lookups),
            tags: parseTags(entry.Tags),
            categories: parseCategories(entry.Categories, lookups.categories),
            publish: entry.Publish
        })
    },
    faqs: {
        sheetIndex: 2,
        sheetName: "FAQs",
        label: "FAQs",
        description: "Frequently asked questions",
        endpoint: "faq",
        standardColumns: ["Question", "Answer", "PermissionSet", "Categories", "Publish", "Name"],
        mapPayload: (entry, lookups) => ({
            data: {
                question: { en: entry.Question },
                answer: { en: entry.Answer }
            },
            permissionSet: permissionSetByName(entry, lookups),
            tags: parseTags(entry.Tags),
            categories: parseCategories(entry.Categories, lookups.categories),
            publish: entry.Publish
        })
    },
    staff: {
        sheetIndex: 3,
        sheetName: "Staff",
        label: "Staff",
        description: "Employee directory profiles",
        // Different sites spell this differently. Only list genuinely different
        // NAMES; punctuation/casing variants (enhancedemployee, enhanced_employee,
        // ...) are matched automatically by the resolver. Order = preference.
        endpoints: ["employee", "enhanced-employee"],
        // "new" added: test column, not a schema field (was injected as { iv: "adsfasdf" }).
        standardColumns: ["PermissionSet", "Categories", "FirstName", "LastName", "Title", "Department", "PhoneNumber", "FaxNumber", "EmailAddress", "Biography", "Publish", "Name", "new"],
        // Field names that vary site to site. Left side = what mapPayload emits.
        fieldAliases: {
            phonenumber: ["phone"],
            faxnumber: ["fax"],
            emailaddress: ["email"],
        },
        mapPayload: (entry, lookups) => ({
            data: {
                firstname: { en: entry.FirstName },
                lastname: { en: entry.LastName },
                title: { en: entry.Title },
                // Reference field (an array of the target Department item's
                // id) — the sheet gives a name, resolved against existing
                // Department content loaded once before the batch loop. A
                // name that isn't found is skipped with a warning rather
                // than failing the row.
                //
                // Partition is `en`, NOT `iv`. A live run against a real
                // site returned "department: Not a known language" for
                // `iv` — this field is language-partitioned, matching the
                // original code here before reference resolution was added.
                // (departments.parentdepartment/staffdirectory below really
                // are `iv`; don't assume one implies the other.)
                department: { en: resolveReferenceIds(entry.Department, lookups.departments, 'department') },
                phonenumber: { en: entry.PhoneNumber },
                faxnumber: { en: entry.FaxNumber },
                emailaddress: { en: entry.EmailAddress },
                Biography: { en: entry.Biography },
            },
            permissionSet: permissionSetByName(entry, lookups),
            tags: parseTags(entry.Tags),
            categories: parseCategories(entry.Categories, lookups.categories),
            publish: entry.Publish
        })
    },
    departments: {
        sheetIndex: 4,
        sheetName: "Departments",
        label: "Departments",
        description: "Department directory entries",
        // Was "enhanceddepartment" (404). Kebab-case slug is the likely default,
        // but some sites drop the "enhanced" prefix — list both, most-specific first.
        endpoints: ["enhanced-department", "department"],
        standardColumns: ["Department", "PhoneNumber", "EmergencyNumber", "FaxNumber", "HoursOfOperation", "AdditionalInformation", "ParentDepartment", "StaffDirectory", "Publish", "Name", "PermissionSet", "Categories"],
        // The field-name variants you hit (faxnumber vs fax, etc.). Add/adjust as
        // you discover more; an alias that no site actually has is simply ignored.
        fieldAliases: {
            phonenumber: ["phone"],
            faxnumber: ["fax", "faxnumber"],
            emergencynumber: ["emergencyphonenumber", "emergency"],
            hoursofoperation: ["hours"],
            additionalinformation: ["additionalinfo"],
        },
        mapPayload: (entry, lookups) => ({
            data: {
                department: { en: entry.Department },
                phonenumber: { en: entry.PhoneNumber },
                emergencynumber: { en: entry.EmergencyNumber },
                faxnumber: { en: entry.FaxNumber },
                hoursofoperation: { en: entry.HoursOfOperation },
                additionalinformation: { en: entry.AdditionalInformation },
                // Both were hardcoded to []. Same as staff.department above:
                // these are Squidex Reference fields, resolved by name against
                // existing Department/Staff content (parentdepartment against
                // this SAME schema, self-referentially).
                parentdepartment: { iv: resolveReferenceIds(entry.ParentDepartment, lookups.departments, 'parent department') },
                staffdirectory: { iv: resolveReferenceIds(entry.StaffDirectory, lookups.staff, 'staff member') },
            },
            permissionSet: permissionSetByName(entry, lookups),
            tags: parseTags(entry.Tags),
            categories: parseCategories(entry.Categories, lookups.categories),
            publish: entry.Publish
        })
    },
    quicklinks: {
        sheetIndex: 5,
        sheetName: "QuickLinks",
        label: "Quick links",
        description: "Homepage shortcut links",
        endpoint: "bpquicklink",
        standardColumns: ["Link", "Publish", "Name", "PermissionSet", "Categories"],
        // Was a hardcoded placeholder ({ url: "#" }, hardcoded permissionSet) that
        // ignored the sheet entirely. Now wired to the real Link/Name columns and
        // consistent with every other type.
        //
        // NOTE on the "title not defined" error you saw: nothing was being
        // injected here (this sheet has no extra columns), so that error came
        // from the API rejecting the placeholder — the link component needs a
        // display label, which { url: "#" } didn't provide. This sends url +
        // displayName. If your schema calls that sub-field `title` rather than
        // `displayName`, rename it below; if the link field is invariant, change
        // `en` to `iv`.
        mapPayload: (entry, lookups) => ({
            data: {
                link: {
                    en: {
                        url: entry.Link,
                        displayName: entry.Name,
                        openInNewWindow: true
                    }
                }
            },
            permissionSet: permissionSetByName(entry, lookups),
            tags: parseTags(entry.Tags),
            categories: parseCategories(entry.Categories, lookups.categories),
            publish: entry.Publish
        })
    },
    news: {
        sheetIndex: 6,
        sheetName: "News",
        label: "News",
        description: "Newsflash announcements",
        endpoint: "newsflash",
        standardColumns: ["NewsTitle", "NewsDate", "NewsText", "NewsAsset", "Publish", "Name", "PermissionSet", "Categories"],
        mapPayload: (entry, lookups) => ({
            data: {
                newstitle: { en: entry.NewsTitle },
                // Was a hardcoded 2019 date. Still needs SOME value when the
                // sheet is blank (this is likely a required field), so this
                // now falls back to the current date/time at upload time
                // instead of a fixed date years in the past.
                newsdate: {
                    iv: {
                        startDate: entry.NewsDate || new Date().toISOString(),
                        endDate: entry.NewsDate || new Date().toISOString()
                    }
                },
                newstext: { en: entry.NewsText },
                // Was hardcoded to []. Assets field — the sheet gives an
                // uploaded file's name (upload it via the Documents tab
                // first), resolved against existing assets on the site.
                newsasset: { iv: resolveReferenceIds(entry.NewsAsset, lookups.assets, 'asset') },
            },
            permissionSet: permissionSetByName(entry, lookups),
            tags: parseTags(entry.Tags),
            categories: parseCategories(entry.Categories, lookups.categories),
            publish: entry.Publish
        })
    },
    calendar: {
        sheetIndex: 7,
        sheetName: "Calendar",
        label: "Calendar",
        description: "Calendar events",
        endpoint: "event",
        standardColumns: ["TitleOfEvent", "TimeTest", "DateOfEvent", "StartTimeOfEvent", "Details", "Attachments", "UrlLink", "SubmissionPDF", "Publish", "Name", "PermissionSet", "Categories"],
        mapPayload: (entry, lookups) => ({
            data: {
                "name-of-event": { iv: entry.TitleOfEvent },
                TimeTest: { iv: entry.TimeTest },
                "date-of-event": { iv: entry.DateOfEvent },
                "start-time-of-event": { iv: entry.StartTimeOfEvent },
                details: { iv: entry.Details },
                // Both were hardcoded to []. Assets fields, same as
                // news.newsasset — sheet values are uploaded-file names,
                // resolved against existing assets. Attachments can be a
                // comma-separated list; SubmissionPDF is expected to be one.
                attachments: { iv: resolveReferenceIds(entry.Attachments, lookups.assets, 'asset') },
                // Was `entry.UrlLink || "test.com"` — a placeholder string is
                // a worse "no value" than just not sending one.
                "url-link": { iv: entry.UrlLink || null },
                "submission-pdf": { iv: resolveReferenceIds(entry.SubmissionPDF, lookups.assets, 'asset') },
            },
            permissionSet: permissionSetByName(entry, lookups),
            tags: parseTags(entry.Tags),
            categories: parseCategories(entry.Categories, lookups.categories),
            publish: entry.Publish
        })
    },
    facility: {
        // Matched by sheetName below now, primarily — sheetIndex is only a
        // fallback for a workbook whose sheet was renamed. Kept correct
        // (was 8, "Agendas and Minutes" in the real workbook; the Facility
        // sheet is index 9: Articles=0, Resources=1, FAQs=2, Staff=3,
        // Departments=4, QuickLinks=5, News=6, Calendar=7, Agendas and
        // Minutes=8, Facility=9) so the fallback path is also right if the
        // sheet is ever found by index instead of name.
        sheetIndex: 9,
        sheetName: "Facility",
        label: "Facilities",
        description: "Facility listings",
        endpoint: "facility",
        standardColumns: ["FacilityName", "Publish", "Name", "PermissionSet", "Categories"],
        mapPayload: (entry, lookups) => ({
            data: { facilityname: { en: entry.FacilityName } },
            permissionSet: permissionSetByName(entry, lookups),
            tags: parseTags(entry.Tags),
            categories: parseCategories(entry.Categories, lookups.categories),
            publish: entry.Publish
        })
    }
}

// --- Optional: extend endpoint candidates without redeploying ---
//
// Set an env var per type to append site-specific slugs discovered in the
// field, e.g.  ENDPOINT_ALIASES_STAFF="employee,enhanced-employee,employe-records"
// These are appended AFTER the built-in candidates (lower preference) and still
// benefit from normalized matching.
for (const [type, cfg] of Object.entries(uploadConfig)) {
    const extra = process.env[`ENDPOINT_ALIASES_${type.toUpperCase()}`];
    if (extra) {
        const extras = extra.split(',').map(s => s.trim()).filter(Boolean);
        if (extras.length) {
            cfg.endpoints = [...getEndpointCandidates(cfg), ...extras];
            console.log(`[${type}] endpoint candidates extended via env: ${cfg.endpoints.join(', ')}`);
        }
    }
}

// --- Optional: extend field aliases without redeploying ---
//
// Set an env var per type as JSON to merge extra field aliases discovered in the
// field, e.g.  FIELD_ALIASES_DEPARTMENTS='{"faxnumber":["fax","faxno"]}'
// Merged onto (and appended to) whatever the config already declares.
for (const [type, cfg] of Object.entries(uploadConfig)) {
    const raw = process.env[`FIELD_ALIASES_${type.toUpperCase()}`];
    if (!raw) continue;
    try {
        const extra = JSON.parse(raw);
        cfg.fieldAliases = cfg.fieldAliases || {};
        for (const [key, alts] of Object.entries(extra)) {
            const existing = cfg.fieldAliases[key] || [];
            cfg.fieldAliases[key] = [...new Set([...existing, ...[].concat(alts)])];
        }
        console.log(`[${type}] field aliases extended via env.`);
    } catch (e) {
        console.warn(`Ignoring FIELD_ALIASES_${type.toUpperCase()} — not valid JSON: ${e.message}`);
    }
}


// --- Custom content types ---
//
// The 9 types above are hand-written because each needs real logic (date
// ranges, reference resolution, compound link fields). A CUSTOM type is
// defined from the GUI instead: pick an unrecognized sheet, give it a
// schema slug, and every column becomes a plain text field. No code edit,
// no restart-safety problem — definitions are saved to customTypes.json
// next to index.js, which matters because the .bat file kills and
// restarts the server every time it runs.
//
// This tool never creates the schema on the site. A custom type must
// already exist in the CMS (that's a deliberate, site-by-site decision
// made by whoever administers it); this only teaches the uploader how to
// send rows to it.

const CUSTOM_TYPES_FILE = path.join(__dirname, 'customTypes.json');

function loadCustomTypes() {
    try {
        if (!fs.existsSync(CUSTOM_TYPES_FILE)) return {};
        const parsed = JSON.parse(fs.readFileSync(CUSTOM_TYPES_FILE, 'utf8'));
        return (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) ? parsed : {};
    } catch (e) {
        // A corrupt file must not stop the server from starting — the 9
        // built-in types still work without it.
        console.warn(`Could not read ${CUSTOM_TYPES_FILE} (${e.message}) — starting with no custom types.`);
        return {};
    }
}

function saveCustomTypes(customTypes) {
    fs.writeFileSync(CUSTOM_TYPES_FILE, JSON.stringify(customTypes, null, 2), 'utf8');
}

// Every column becomes { iv: value } — the same shape processUpload
// already uses for columns it doesn't recognize on a built-in type, so
// this is consistent with existing behavior rather than a new convention.
// Blank cells are skipped entirely (not sent as empty strings), matching
// the dynamic-injection rule already in processUpload.
// PermissionSet/Categories/Publish are handled like every other type.
//
// An EMPTY `columns` list means "discover them from each row" — that's
// how a type registered by hand works, since there's no workbook on hand
// to read headers from at registration time. Passing an explicit list
// (which the unrecognized-sheet shortcut does, having just read them)
// pins the fields instead, so a stray column added to the sheet later
// doesn't silently start uploading.
const SHARED_COLUMNS = ['PermissionSet', 'Categories', 'Publish', 'Tags', 'Name'];

function buildCustomMapPayload(columns) {
    const pinned = (columns || []).filter(c => !SHARED_COLUMNS.includes(c));
    return (entry, lookups) => {
        const sourceColumns = pinned.length
            ? pinned
            : Object.keys(entry).filter(c => !SHARED_COLUMNS.includes(c));

        const data = {};
        for (const column of sourceColumns) {
            const value = entry[column];
            if (value !== undefined && value !== '') data[column] = { iv: value };
        }
        return {
            data,
            permissionSet: permissionSetByName(entry, lookups),
            tags: parseTags(entry.Tags),
            categories: parseCategories(entry.Categories, lookups.categories),
            publish: entry.Publish
        };
    };
}

// Turns a saved definition into a live uploadConfig entry. sheetIndex is
// -1 because a custom type is ALWAYS matched by sheet name (or an explicit
// override) — there's no meaningful fixed position for it, and -1 makes a
// name-match failure produce a clear error instead of silently grabbing
// whatever sheet happens to sit at some index.
function buildCustomTypeConfig(definition) {
    const columns = definition.columns || [];
    return {
        sheetIndex: -1,
        sheetName: definition.sheetName,
        label: definition.label,
        description: definition.description || `Custom import from the "${definition.sheetName}" sheet`,
        endpoint: definition.endpoint,
        // The shared columns must be listed as "standard" even though the
        // custom mapper handles them, or processUpload's dynamic
        // injection (which adds any column NOT in standardColumns) would
        // re-add PermissionSet/Categories/Publish as plain data fields on
        // top of the proper handling.
        standardColumns: [...new Set([...columns, ...SHARED_COLUMNS])],
        isCustom: true,
        columns,
        mapPayload: buildCustomMapPayload(columns)
    };
}

// Merge whatever is on disk into uploadConfig at startup, so custom types
// behave exactly like built-in ones everywhere downstream (the upload
// route, the connection check, sheet matching) with no special-casing.
function registerCustomTypes(customTypes) {
    for (const [key, definition] of Object.entries(customTypes)) {
        uploadConfig[key] = buildCustomTypeConfig(definition);
    }
    return uploadConfig;
}

registerCustomTypes(loadCustomTypes());

// --- Asset (raw file) upload ---
//
// Separate from the Excel-driven content upload below: this posts ONE binary
// file (image, PDF, doc, whatever) straight to the assets endpoint as
// multipart/form-data — the same shape the site's own upload UI sends.
// `categories` and `tags` are repeated once per value (that's how this API
// represents an array in form-data); `permissionSet` is a single id.
// Node's built-in fetch/FormData/Blob handle the multipart body; no extra
// package needed.

// Transient failures worth another try on a long run: rate limiting and
// gateway hiccups. Anything else (400 validation, 401 token, 500) fails
// straight away, since retrying would just fail the same way.
// Every attempt is cut off after timeoutMs, so an API call that never
// answers fails that file instead of stalling its whole batch (Node's own
// fetch would otherwise wait 5 minutes per attempt). A Retry-After longer
// than maxDelayMs is capped — the file fails and can be retried later
// rather than the run silently sleeping for however long the site asked.
// Exported and mutable so tests can shrink the delays.
const ASSET_RETRY = { attempts: 3, baseDelayMs: 1000, timeoutMs: 90 * 1000, maxDelayMs: 30 * 1000 };
const RETRYABLE_STATUSES = new Set([429, 502, 503, 504]);

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Asset-upload events also go to asset-upload.log next to index.js, so a
// long run can be diagnosed afterwards even once the server window is
// closed. Never logs the token.
const ASSET_LOG_FILE = path.join(__dirname, 'asset-upload.log');

function logAsset(level, message) {
    const line = `${new Date().toISOString()} [${level}] ${message}`;
    (level === 'error' ? console.error : level === 'warn' ? console.warn : console.log)(line);
    if (process.env.JEST_WORKER_ID) return; // keep test runs out of the real log
    fs.appendFile(ASSET_LOG_FILE, line + '\n', () => { /* logging must never break an upload */ });
}

// fetch() with retries on RETRYABLE_STATUSES, network errors and timeouts,
// backing off 1s, 2s, 4s… (or whatever Retry-After asks for, up to
// maxDelayMs). `buildInit` is called per attempt because a FormData body
// can't safely be reused once sent.
async function fetchWithRetry(url, buildInit, label = 'request') {
    let lastError;
    for (let attempt = 1; attempt <= ASSET_RETRY.attempts; attempt++) {
        try {
            const response = await fetch(url, { ...buildInit(), signal: AbortSignal.timeout(ASSET_RETRY.timeoutMs) });
            if (!RETRYABLE_STATUSES.has(response.status) || attempt === ASSET_RETRY.attempts) {
                return response;
            }
            const retryAfter = Number(response.headers?.get?.('retry-after'));
            const wanted = retryAfter > 0 ? retryAfter * 1000 : ASSET_RETRY.baseDelayMs * 2 ** (attempt - 1);
            const delay = Math.min(wanted, ASSET_RETRY.maxDelayMs);
            logAsset('warn', `${label}: HTTP ${response.status}, retrying in ${delay}ms (attempt ${attempt + 1} of ${ASSET_RETRY.attempts})`);
            await sleep(delay);
        } catch (error) {
            lastError = error.name === 'TimeoutError'
                ? new Error(`No response from the site after ${Math.round(ASSET_RETRY.timeoutMs / 1000)}s`)
                : error;
            if (attempt === ASSET_RETRY.attempts) break;
            const delay = ASSET_RETRY.baseDelayMs * 2 ** (attempt - 1);
            logAsset('warn', `${label}: ${lastError.message}, retrying in ${delay}ms (attempt ${attempt + 1} of ${ASSET_RETRY.attempts})`);
            await sleep(delay);
        }
    }
    throw lastError;
}

async function postAssetFile({ siteURL, accessToken, parentId, publish, fileBuffer, fileName, mimeType, categoryIds = [], permissionSetId, altText, tags = [], duplicate = false }) {
    const buildForm = () => {
        const form = new FormData();
        form.append('file', new Blob([fileBuffer], { type: mimeType || 'application/octet-stream' }), fileName);

        for (const id of categoryIds) {
            if (id) form.append('categories', id);
        }
        if (permissionSetId) {
            form.append('permissionSet', permissionSetId);
        }
        if (altText) {
            form.append('altText', altText);
        }
        for (const tag of tags) {
            if (tag) form.append('tags', tag);
        }
        return form;
    };

    const url = new URL(`https://content.civicplus.com/api/apps/${siteURL}/assets`);
    url.searchParams.set('parentId', parentId || '00000000-0000-0000-0000-000000000000');
    url.searchParams.set('publish', String(!!publish));
    // Without this the site recognises a file it already has (same content
    // and name) and hands back the existing asset instead of a new one.
    if (duplicate) url.searchParams.set('duplicate', 'true');

    const response = await fetchWithRetry(url, () => ({
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${accessToken}`
            // No Content-Type here — fetch generates its own multipart
            // boundary, and setting one manually will break the upload.
        },
        body: buildForm()
    }), `Asset "${fileName}"`);

    if (!response.ok) {
        const errBody = await response.text();
        throw new Error(`HTTP ${response.status}: ${errBody}`);
    }
    return response.json();
}

// Sets alt text / tags on an asset that already exists. Used as a follow-up
// when the create call doesn't echo back the alt text it was given — the
// asset DTO has an `altText` field, but whether the multipart create accepts
// it isn't documented, so this makes sure it lands either way. PUT on an
// asset in Squidex only changes the fields that are sent.
async function annotateAsset({ siteURL, accessToken, assetId, altText, tags }) {
    const body = {};
    if (altText) body.altText = altText;
    if (tags && tags.length) body.tags = tags;

    const response = await fetchWithRetry(
        `https://content.civicplus.com/api/apps/${siteURL}/assets/${assetId}`,
        () => ({
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(body)
        }),
        `Annotate asset ${assetId}`
    );

    if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${await response.text()}`);
    }
    return response.json();
}

// --- Assets sheet ---
//
// Optional per-file settings for the folder upload, read from a sheet named
// "Assets" in the same workbook as everything else. One row per file;
// FileName is matched against the files in the chosen folder. Any blank cell
// falls back to the value in the dashboard's Documents form.
const ASSET_SHEET_NAME = 'Assets';
const ASSET_SHEET_COLUMNS = ['FileName', 'ParentFolderId', 'PermissionSet', 'Categories', 'AltText', 'Tags', 'Publish'];

// Loose header lookup ("Alt Text", "alt_text", "ALTTEXT" all work), same
// normalization used for sheet and schema names.
const ASSET_COLUMN_KEYS = {
    filename: 'fileName',
    file: 'fileName',
    parentfolderid: 'parentId',
    parentid: 'parentId',
    folderid: 'parentId',
    permissionset: 'permissionSet',
    categories: 'categories',
    category: 'categories',
    alttext: 'altText',
    alt: 'altText',
    tags: 'tags',
    publish: 'publish'
};

// Yes/No-style cell -> boolean. Blank (or anything unrecognised) -> undefined,
// meaning "use the form default".
function parsePublishFlag(value) {
    if (typeof value === 'boolean') return value;
    const v = String(value ?? '').trim().toLowerCase();
    if (['true', 'yes', 'y', '1', 'publish', 'published'].includes(v)) return true;
    if (['false', 'no', 'n', '0', 'draft'].includes(v)) return false;
    return undefined;
}

function findAssetSheetName(sheetNames) {
    return sheetNames.find(name => normalizeSlug(name) === normalizeSlug(ASSET_SHEET_NAME));
}

// Rows of the Assets sheet as { fileName, parentId, permissionSet,
// categories, altText, tags, publish } with blanks dropped. Rows with no
// FileName are skipped. Returns null when the workbook has no Assets sheet.
function parseAssetRows(workbook) {
    const sheetName = findAssetSheetName(workbook.SheetNames);
    if (!sheetName) return null;

    const raw = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });
    const rows = [];
    for (const record of raw) {
        const row = {};
        for (const [header, value] of Object.entries(record)) {
            const key = ASSET_COLUMN_KEYS[normalizeSlug(header)];
            if (!key) continue;
            if (key === 'publish') {
                const flag = parsePublishFlag(value);
                if (flag !== undefined) row.publish = flag;
            } else {
                const text = String(value ?? '').trim();
                if (text) row[key] = text;
            }
        }
        if (row.fileName) rows.push(row);
    }
    return { sheetName, rows };
}

// A ready-to-fill Assets sheet listing the given file names, with the
// form's current values pre-filled in every row.
function buildAssetTemplateWorkbook(fileNames, defaults = {}) {
    const rows = fileNames.map(fileName => ({
        FileName: fileName,
        ParentFolderId: defaults.parentId || '',
        PermissionSet: defaults.permissionSet || '',
        Categories: defaults.categories || '',
        AltText: defaults.altText || '',
        Tags: defaults.tags || '',
        Publish: defaults.publish ? 'Yes' : 'No'
    }));
    const sheet = XLSX.utils.json_to_sheet(rows, { header: ASSET_SHEET_COLUMNS });
    sheet['!cols'] = [{ wch: 40 }, { wch: 38 }, { wch: 18 }, { wch: 24 }, { wch: 40 }, { wch: 20 }, { wch: 8 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet, ASSET_SHEET_NAME);
    return wb;
}

const splitList = (value) => String(value ?? '').split(',').map(s => s.trim()).filter(Boolean);

// Permission set + category lookups for one site, cached for a few minutes
// so a large run split into many requests doesn't refetch them per batch.
// Only fully successful lookups are cached; a failed one is retried next
// batch rather than leaving names unresolved for the whole run.
const ASSET_LOOKUP_TTL_MS = 5 * 60 * 1000;
const assetLookupCache = new Map();

async function getAssetLookups(siteURL, accessToken) {
    const key = `${siteURL}\u0000${accessToken}`;
    const cached = assetLookupCache.get(key);
    if (cached && Date.now() - cached.at < ASSET_LOOKUP_TTL_MS) return cached.lookups;

    let ok = true;
    const [permissionSet, categories] = await Promise.all([
        fetchReferenceData(siteURL, 'permissionSet', accessToken).catch(() => { ok = false; return new Map(); }),
        fetchReferenceData(siteURL, 'categories', accessToken).catch(() => { ok = false; return new Map(); })
    ]);
    const lookups = { permissionSet, categories };
    if (ok) assetLookupCache.set(key, { at: Date.now(), lookups });
    return lookups;
}

// The site flags a re-upload of a file it already has with
// _meta.isDuplicate (a "true"/"false" string) and returns the existing asset
// instead of creating a new one.
const isDuplicateResponse = (data) => String(data?._meta?.isDuplicate ?? '').toLowerCase() === 'true';

// Uploads one file with its effective settings (sheet row merged over the
// form defaults). Resolves to { file, id, duplicate } or { file, error },
// plus any non-fatal warnings (unknown category names, alt text that didn't
// stick).
async function uploadOneAsset({ file, settings, siteURL, accessToken, lookups, allowDuplicates }) {
    const warnings = [];

    const permissionSetId = settings.permissionSet
        ? resolveId(settings.permissionSet, lookups.permissionSet, 'permissionSet')
        : undefined;
    if (settings.permissionSet && !permissionSetId) {
        warnings.push(`${file.originalname}: permission set "${settings.permissionSet}" not found on the site`);
    }

    const categoryIds = [];
    for (const name of splitList(settings.categories)) {
        const id = resolveId(name, lookups.categories, 'category');
        if (id) categoryIds.push(id);
        else warnings.push(`${file.originalname}: category "${name}" not found on the site`);
    }

    const tags = splitList(settings.tags);
    const altText = settings.altText ? String(settings.altText) : undefined;

    try {
        const data = await postAssetFile({
            siteURL,
            accessToken,
            parentId: settings.parentId,
            publish: settings.publish,
            fileBuffer: file.buffer,
            fileName: file.originalname,
            mimeType: file.mimetype,
            categoryIds,
            permissionSetId,
            altText,
            tags,
            duplicate: allowDuplicates
        });

        const needsAltText = altText && data?.altText !== altText;
        const needsTags = tags.length && !(Array.isArray(data?.tags) && tags.every(t => data.tags.includes(t)));
        if (data?.id && (needsAltText || needsTags)) {
            try {
                await annotateAsset({ siteURL, accessToken, assetId: data.id, altText, tags });
            } catch (error) {
                warnings.push(`${file.originalname}: uploaded, but alt text/tags could not be set (${error.message})`);
            }
        }
        return { file: file.originalname, id: data?.id, duplicate: isDuplicateResponse(data), warnings };
    } catch (error) {
        logAsset('error', `Asset "${file.originalname}" failed: ${error.message}`);
        return { file: file.originalname, error: error.message, warnings };
    }
}

// Most files a single request may carry. The dashboard splits a big folder
// into batches below this and sends them one after another, so a run of
// thousands of files never has more than one batch buffered in memory.
const ASSET_FILES_PER_REQUEST = 50;

const uploadAssetFiles = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 100 * 1024 * 1024, files: ASSET_FILES_PER_REQUEST }
}).array('files', ASSET_FILES_PER_REQUEST);

// What the server knows about each batch the dashboard has sent, keyed by
// the batchId the dashboard puts in the query string. The dashboard polls
// this while a batch is in flight, so it can tell "the server is working on
// it" apart from "the request never arrived" (and resend), and can pick up
// the result of a batch whose response got lost on the way back.
// state: receiving -> working -> finished. Entries expire after an hour.
const assetBatchStatus = new Map();
const ASSET_BATCH_STATUS_TTL_MS = 60 * 60 * 1000;

function setAssetBatchStatus(batchId, update) {
    if (!batchId) return;
    const now = Date.now();
    for (const [id, entry] of assetBatchStatus) {
        if (now - entry.updatedAt > ASSET_BATCH_STATUS_TTL_MS) assetBatchStatus.delete(id);
    }
    assetBatchStatus.set(batchId, { ...assetBatchStatus.get(batchId), ...update, updatedAt: now });
}

app.get('/upload/assets/status/:batchId', (req, res) => {
    const entry = assetBatchStatus.get(req.params.batchId);
    if (!entry) return res.json({ state: 'unknown' });
    const { state, done, total, result, statusCode } = entry;
    res.json({ state, done, total, result, statusCode });
});

// One batch of the folder upload. The assets API takes one file per request,
// so each file becomes its own POST, a few at a time.
//
// Body fields: url, apiKey, and the form-level defaults (parentId,
// permissionSet, categories, altText, tags, publish). Optional `fileSettings`
// is a JSON array lined up with the files — one object per file holding that
// file's Assets-sheet row — whose non-blank values override the defaults.
// Category / permission set values can be NAMES (resolved against the site,
// same as the Excel sheets) or raw ids.
app.post('/upload/assets', (req, res) => {
    // In the query string (not the body) so it's known before the files
    // have finished arriving.
    const batchId = req.query.batchId ? String(req.query.batchId) : null;
    setAssetBatchStatus(batchId, { state: 'receiving' });

    // Every response from here on also lands in the batch status, so the
    // dashboard can recover it if the response itself never arrives.
    const reply = (statusCode, body) => {
        setAssetBatchStatus(batchId, { state: 'finished', statusCode, result: body });
        return res.status(statusCode).json(body);
    };

    uploadAssetFiles(req, res, async (uploadErr) => {
        if (uploadErr) {
            const message = uploadErr.code === 'LIMIT_FILE_COUNT'
                ? `Too many files in one request (max ${ASSET_FILES_PER_REQUEST}). The dashboard sends large folders in batches automatically.`
                : uploadErr.message;
            return reply(400, { message });
        }

        const { url, apiKey } = req.body;

        if (!req.files || !req.files.length) {
            return reply(400, { message: 'No files selected. Choose a folder and try again.' });
        }
        if (!url || !apiKey) {
            return reply(400, { message: 'Website URL and token are required.' });
        }

        let perFile = [];
        if (req.body.fileSettings) {
            try {
                perFile = JSON.parse(req.body.fileSettings);
                if (!Array.isArray(perFile)) throw new Error('not an array');
            } catch (_) {
                return reply(400, { message: 'fileSettings must be a JSON array.' });
            }
        }

        const defaults = {
            parentId: req.body.parentId,
            permissionSet: req.body.permissionSet,
            categories: req.body.categories,
            altText: req.body.altText,
            tags: req.body.tags,
            publish: req.body.publish === 'true' || req.body.publish === true
        };
        const allowDuplicates = req.body.allowDuplicates === 'true';
        const batchLabel = req.body.batchLabel ? ` (${req.body.batchLabel})` : '';
        const startedAt = Date.now();

        setAssetBatchStatus(batchId, { state: 'working', done: 0, total: req.files.length });
        logAsset('info', `Batch${batchLabel} started - site ${url} - ${req.files.length} file(s), first "${req.files[0].originalname}"`);

        try {
            const lookups = await getAssetLookups(url, apiKey);

            const successes = [];
            const failures = [];
            const warnings = [];
            const concurrency = 5;

            for (let i = 0; i < req.files.length; i += concurrency) {
                const slice = req.files.slice(i, i + concurrency);
                const results = await Promise.all(slice.map((file, j) => {
                    const row = perFile[i + j] || {};
                    const settings = { ...defaults };
                    for (const [key, value] of Object.entries(row)) {
                        if (key === 'publish') {
                            if (typeof value === 'boolean') settings.publish = value;
                        } else if (value !== undefined && value !== null && String(value).trim() !== '') {
                            settings[key] = value;
                        }
                    }
                    return uploadOneAsset({ file, settings, siteURL: url, accessToken: apiKey, lookups, allowDuplicates });
                }));

                for (const r of results) {
                    warnings.push(...r.warnings);
                    if (r.error) failures.push(r);
                    else successes.push(r);
                }
                setAssetBatchStatus(batchId, { done: successes.length + failures.length });
            }

            // A duplicate succeeded as a request but created nothing new, so it
            // is reported separately rather than counted as uploaded.
            const duplicates = successes.filter(s => s.duplicate);
            const created = successes.length - duplicates.length;

            const body = {
                uploaded: created,
                duplicates: duplicates.length,
                duplicateFiles: duplicates.map(d => d.file),
                failed: failures.length,
                failedFiles: failures.map(f => f.file),
                errors: failures.map(f => `${f.file}: ${f.error}`),
                warnings
            };

            logAsset('info', `Batch${batchLabel} finished in ${Math.round((Date.now() - startedAt) / 1000)}s - ${created} uploaded, ${duplicates.length} already on the site, ${failures.length} failed`);

            const dupNote = duplicates.length ? ` ${duplicates.length} already existed on the site.` : '';
            if (failures.length === 0) {
                return reply(200, { ...body, message: `Upload complete. Uploaded ${created} file(s).${dupNote}` });
            }
            return reply(207, {
                ...body,
                message: `Uploaded ${created} of ${successes.length + failures.length} file(s). ${failures.length} failed.${dupNote}`
            });
        } catch (error) {
            logAsset('error', `Batch${batchLabel} crashed: ${error.stack || error.message}`);
            reply(500, { message: error.message || 'Upload failed. Check server logs.' });
        }
    });
});

// Reads the Assets sheet out of a workbook so the dashboard can match its
// rows to the files in the chosen folder before uploading.
app.post('/inspect-assets-sheet', (req, res) => {
    uploadSingle(req, res, (uploadErr) => {
        if (uploadErr) return res.status(400).json({ message: uploadErr.message });
        if (!req.file) return res.status(400).json({ message: 'No Excel file uploaded. Choose a file and try again.' });

        try {
            const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
            const parsed = parseAssetRows(workbook);
            if (!parsed) {
                return res.json({ found: false, rows: [] });
            }
            res.json({ found: true, sheetName: parsed.sheetName, rows: parsed.rows });
        } catch (error) {
            console.error(error);
            res.status(500).json({ message: error.message || 'Could not read the workbook.' });
        }
    });
});

// Download a pre-filled Assets sheet for the files in a folder — the quick
// way to get one row per file without typing thousands of names.
app.post('/assets/template', (req, res) => {
    const { fileNames, defaults } = req.body || {};
    if (!Array.isArray(fileNames) || !fileNames.length) {
        return res.status(400).json({ message: 'Choose a folder first so there are file names to list.' });
    }
    const wb = buildAssetTemplateWorkbook(fileNames.map(String), defaults || {});
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="AssetsSheet.xlsx"');
    res.send(buffer);
});


// --- Old-site link report ---
//
// Read-only. Scans every content item on the site for links that still
// point at the old website (by domain), and lists the ones that point at a
// FILE — those need repointing to the asset in Evolve, or a manual look.
// Each file link is matched by file name against the site's asset library,
// so the report says which ones have a likely replacement already uploaded
// and which have nothing to point at yet. Nothing on the site is changed.

// Extensions that make a link a file download rather than a page.
const FILE_EXTENSIONS = new Set([
    'pdf', 'doc', 'docx', 'xls', 'xlsx', 'xlsm', 'csv', 'ppt', 'pptx', 'txt', 'rtf', 'odt', 'ods',
    'jpg', 'jpeg', 'png', 'gif', 'bmp', 'tif', 'tiff', 'svg', 'webp', 'heic',
    'mp3', 'wav', 'mp4', 'mov', 'avi', 'wmv', 'm4v',
    'zip', 'kmz', 'kml', 'dwg', 'shp', 'xml', 'json', 'eps', 'ai', 'psd'
]);

// Path segments CMSs commonly serve documents from even without an extension
// (CivicPlus CivicEngage's /DocumentCenter/View/123, generic /documents/...).
const FILE_PATH_HINTS = /\/(documentcenter\/view|documents|document|files|uploads|media|assets|archive\.aspx|imagerepository)\//i;

const OLD_LINK_PAGE_SIZE = 200;

// Every string inside a content item's data, with the field it came from.
// Squidex data is { field: { iv|en: value } }, and values can be nested
// (components, arrays, objects), so this walks all of it.
function collectStrings(value, fieldPath, out) {
    if (typeof value === 'string') {
        out.push({ field: fieldPath, text: value });
    } else if (Array.isArray(value)) {
        value.forEach((v, i) => collectStrings(v, `${fieldPath}[${i}]`, out));
    } else if (value && typeof value === 'object') {
        for (const [k, v] of Object.entries(value)) {
            // Partition keys (iv, en, …) aren't useful in the report.
            const next = /^(iv|[a-z]{2}(-[a-z]{2})?)$/i.test(k) && fieldPath ? fieldPath : (fieldPath ? `${fieldPath}.${k}` : k);
            collectStrings(v, next, out);
        }
    }
    return out;
}

const decodeHtml = (s) => String(s)
    .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ');

const stripTags = (s) => decodeHtml(String(s).replace(/<[^>]*>/g, '')).replace(/\s+/g, ' ').trim();

// Domain list from the form: "co.monmouth.nj.us, www.monmouthcounty.gov"
// -> ['co.monmouth.nj.us', 'monmouthcounty.gov'] (scheme, path and www. dropped).
function parseOldDomains(input) {
    return String(input ?? '')
        .split(/[,\s]+/)
        .map(d => d.trim().toLowerCase().replace(/^https?:\/\//, '').replace(/\/.*$/, '').replace(/^www\./, ''))
        .filter(Boolean);
}

function hostMatches(host, domains) {
    const h = host.toLowerCase().replace(/^www\./, '');
    return domains.some(d => h === d || h.endsWith(`.${d}`));
}

// The file name a link points at: last path segment, URL-decoded.
function fileNameFromUrl(u) {
    const last = u.pathname.split('/').filter(Boolean).pop() || '';
    try { return decodeURIComponent(last); } catch (_) { return last; }
}

function isFileLink(u) {
    const name = fileNameFromUrl(u);
    const ext = name.includes('.') ? name.split('.').pop().toLowerCase() : '';
    return FILE_EXTENSIONS.has(ext) || FILE_PATH_HINTS.test(u.pathname + '/');
}

// Every link to one of the old domains in a string: href/src attributes
// (with their link text) plus bare URLs in plain-text fields. Relative
// links can't be attributed to a domain, so they're left out.
function findOldLinks(text, domains) {
    const found = [];
    const seen = new Set();
    const add = (raw, linkText) => {
        const href = decodeHtml(raw).trim();
        let u;
        try { u = new URL(href.startsWith('//') ? `https:${href}` : href); } catch (_) { return; }
        if (!/^https?:$/.test(u.protocol) || !hostMatches(u.hostname, domains)) return;
        if (seen.has(href)) return;
        seen.add(href);
        found.push({ url: href, linkText: linkText || '', isFile: isFileLink(u), fileName: fileNameFromUrl(u) });
    };

    const anchor = /<a\b[^>]*?\bhref\s*=\s*(["'])(.*?)\1[^>]*>([\s\S]*?)<\/a>/gi;
    let m;
    while ((m = anchor.exec(text))) add(m[2], stripTags(m[3]));

    const attr = /\b(?:href|src)\s*=\s*(["'])(.*?)\1/gi;
    while ((m = attr.exec(text))) add(m[2], '');

    const bare = /(?:https?:)?\/\/[^\s"'<>()]+/gi;
    while ((m = bare.exec(text))) add(m[0].replace(/[.,;:]+$/, ''), '');

    return found;
}

// A readable title for an item, from whichever title-like field it has.
function itemTitle(item) {
    const data = item?.data || {};
    const preferred = ['title', 'name', 'heading', 'question', 'newstitle', 'titleofevent', 'facilityname', 'department', 'link'];
    const keys = Object.keys(data);
    const ordered = [
        ...preferred.flatMap(p => keys.filter(k => normalizeSlug(k) === p)),
        ...keys.filter(k => /title|name/i.test(k))
    ];
    for (const key of ordered) {
        const strings = collectStrings(data[key], key, []);
        const text = strings.map(s => stripTags(s.text)).find(Boolean);
        if (text) return text.slice(0, 200);
    }
    return '';
}

// Loose file-name key so "Form G (Sample).pdf" and "Form_G_Sample.pdf" meet.
const looseFileKey = (name) => String(name).toLowerCase().replace(/\.[a-z0-9]+$/, '').replace(/[^a-z0-9]/g, '');

// Pages through a list endpoint ({ total, items }) until everything is read.
async function fetchAllPages(baseUrl, accessToken, label, extraHeaders = {}) {
    const items = [];
    for (let skip = 0; ; skip += OLD_LINK_PAGE_SIZE) {
        const url = `${baseUrl}${baseUrl.includes('?') ? '&' : '?'}$top=${OLD_LINK_PAGE_SIZE}&$skip=${skip}`;
        const response = await fetchWithRetry(url, () => ({
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}`, ...extraHeaders }
        }), label);
        if (!response.ok) {
            throw new Error(`${label}: HTTP ${response.status}: ${(await response.text()).slice(0, 300)}`);
        }
        const json = await response.json();
        const page = json.items || (Array.isArray(json) ? json : []);
        items.push(...page);
        const total = typeof json.total === 'number' ? json.total : null;
        if (!page.length || page.length < OLD_LINK_PAGE_SIZE || (total !== null && items.length >= total)) break;
    }
    return items;
}

// The site's assets as exact and loose file-name lookups.
async function fetchAssetIndex(siteURL, accessToken) {
    const assets = await fetchAllPages(`https://content.civicplus.com/api/apps/${siteURL}/assets`, accessToken, 'List assets');
    const exact = new Map();
    const loose = new Map();
    for (const a of assets) {
        const name = a.fileName || a.slug;
        if (!name) continue;
        const entry = { id: a.id, fileName: name, url: a?._links?.content?.href || '' };
        const e = exact.get(name.toLowerCase()) || [];
        e.push(entry);
        exact.set(name.toLowerCase(), e);
        const l = loose.get(looseFileKey(name)) || [];
        l.push(entry);
        loose.set(looseFileKey(name), l);
    }
    return { exact, loose, count: assets.length };
}

function matchAsset(fileName, assetIndex) {
    const exact = assetIndex.exact.get(String(fileName).toLowerCase());
    if (exact) return { how: 'exact', assets: exact };
    const key = looseFileKey(fileName);
    const loose = key ? assetIndex.loose.get(key) : null;
    if (loose) return { how: 'similar', assets: loose };
    return { how: 'none', assets: [] };
}

// Scans the whole site. Returns { fileLinks, pageLinks, summary }.
async function buildOldLinkReport({ siteURL, accessToken, oldDomains, schemas }) {
    const domains = parseOldDomains(oldDomains);
    if (!domains.length) throw new Error('Enter the old website\'s domain (for example co.monmouth.nj.us).');

    const schemaNames = schemas && schemas.length ? schemas : await fetchSchemaNames(siteURL, accessToken);
    if (!schemaNames) throw new Error('Could not list the content types on this site. Check the site name and token with Test connection.');

    const assetIndex = await fetchAssetIndex(siteURL, accessToken);

    const fileLinks = [];
    const pageLinks = [];
    const schemaErrors = [];
    let itemsScanned = 0;

    for (const schema of schemaNames) {
        let items;
        try {
            // X-Unpublished so drafts are scanned too, not just published items.
            items = await fetchAllPages(
                `https://content.civicplus.com/api/content/${siteURL}/${schema}`,
                accessToken,
                `List "${schema}" content`,
                { 'X-Unpublished': '1' }
            );
        } catch (error) {
            schemaErrors.push(`${schema}: ${error.message}`);
            continue;
        }

        for (const item of items) {
            itemsScanned++;
            const base = {
                contentType: schema,
                title: itemTitle(item),
                itemId: item.id,
                status: item.status || '',
                editLink: `https://content.civicplus.com/app/${siteURL}/content/${schema}/${item.id}`
            };
            for (const { field, text } of collectStrings(item.data || {}, '', [])) {
                for (const link of findOldLinks(text, domains)) {
                    const row = { ...base, field, oldUrl: link.url, linkText: link.linkText, fileName: link.fileName };
                    if (!link.isFile) {
                        pageLinks.push(row);
                        continue;
                    }
                    const match = matchAsset(link.fileName, assetIndex);
                    fileLinks.push({
                        ...row,
                        result: match.how === 'exact'
                            ? 'Not changed - a file with this name IS in Evolve assets, repoint the link to it'
                            : match.how === 'similar'
                                ? 'Not changed - a file with a SIMILAR name is in Evolve assets, check it is the same file and repoint'
                                : 'Not changed - NO matching file in Evolve assets, find/upload the file or remove the link',
                        matchType: match.how,
                        matchedAssets: match.assets.map(a => a.fileName).join(', '),
                        matchedAssetIds: match.assets.map(a => a.id).join(', '),
                        matchedAssetUrls: match.assets.map(a => a.url).filter(Boolean).join(', ')
                    });
                }
            }
        }
    }

    const summary = {
        domains,
        contentTypesScanned: schemaNames.length - schemaErrors.length,
        itemsScanned,
        assetsOnSite: assetIndex.count,
        fileLinks: fileLinks.length,
        fileLinksWithExactMatch: fileLinks.filter(l => l.matchType === 'exact').length,
        fileLinksWithSimilarMatch: fileLinks.filter(l => l.matchType === 'similar').length,
        fileLinksWithNoMatch: fileLinks.filter(l => l.matchType === 'none').length,
        itemsWithFileLinks: new Set(fileLinks.map(l => `${l.contentType}/${l.itemId}`)).size,
        pageLinks: pageLinks.length,
        schemaErrors
    };
    return { fileLinks, pageLinks, summary };
}

// The report as a workbook: Summary, File links (sorted so the ones with
// nothing to point at come first), and Other old-site links (pages).
function buildOldLinkReportWorkbook({ fileLinks, pageLinks, summary }) {
    const wb = XLSX.utils.book_new();

    const summaryRows = [
        ['Old-site link report', ''],
        ['Old domain(s)', summary.domains.join(', ')],
        ['Content types scanned', summary.contentTypesScanned],
        ['Content items scanned', summary.itemsScanned],
        ['Assets in Evolve', summary.assetsOnSite],
        ['', ''],
        ['Links to old-site FILES (not changed)', summary.fileLinks],
        ['  ...no matching file in Evolve (needs the file found/uploaded)', summary.fileLinksWithNoMatch],
        ['  ...a similar file name in Evolve (check, then repoint)', summary.fileLinksWithSimilarMatch],
        ['  ...same file name in Evolve (repoint the link)', summary.fileLinksWithExactMatch],
        ['Content items containing these links', summary.itemsWithFileLinks],
        ['Links to old-site PAGES (separate sheet)', summary.pageLinks],
        ...summary.schemaErrors.map(e => ['Could not scan', e])
    ];
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    summarySheet['!cols'] = [{ wch: 62 }, { wch: 60 }];
    XLSX.utils.book_append_sheet(wb, summarySheet, 'Summary');

    const order = { none: 0, similar: 1, exact: 2 };
    const fileRows = [...fileLinks]
        .sort((a, b) => order[a.matchType] - order[b.matchType] || a.contentType.localeCompare(b.contentType) || a.title.localeCompare(b.title))
        .map(l => ({
            'Result': l.result,
            'Content type': l.contentType,
            'Title': l.title,
            'Status': l.status,
            'Field': l.field,
            'Link text': l.linkText,
            'Old URL': l.oldUrl,
            'File name': l.fileName,
            'Matching Evolve asset(s)': l.matchedAssets,
            'Asset ID(s)': l.matchedAssetIds,
            'Asset URL(s)': l.matchedAssetUrls,
            'Edit in Evolve': l.editLink,
            'Item ID': l.itemId
        }));
    const fileSheet = XLSX.utils.json_to_sheet(fileRows, { header: ['Result', 'Content type', 'Title', 'Status', 'Field', 'Link text', 'Old URL', 'File name', 'Matching Evolve asset(s)', 'Asset ID(s)', 'Asset URL(s)', 'Edit in Evolve', 'Item ID'] });
    fileSheet['!cols'] = [{ wch: 50 }, { wch: 18 }, { wch: 40 }, { wch: 10 }, { wch: 16 }, { wch: 30 }, { wch: 60 }, { wch: 40 }, { wch: 40 }, { wch: 38 }, { wch: 40 }, { wch: 60 }, { wch: 38 }];
    XLSX.utils.book_append_sheet(wb, fileSheet, 'File links');

    const pageRows = pageLinks.map(l => ({
        'Content type': l.contentType,
        'Title': l.title,
        'Status': l.status,
        'Field': l.field,
        'Link text': l.linkText,
        'Old URL': l.oldUrl,
        'Edit in Evolve': l.editLink,
        'Item ID': l.itemId
    }));
    const pageSheet = XLSX.utils.json_to_sheet(pageRows, { header: ['Content type', 'Title', 'Status', 'Field', 'Link text', 'Old URL', 'Edit in Evolve', 'Item ID'] });
    pageSheet['!cols'] = [{ wch: 18 }, { wch: 40 }, { wch: 10 }, { wch: 16 }, { wch: 30 }, { wch: 60 }, { wch: 60 }, { wch: 38 }];
    XLSX.utils.book_append_sheet(wb, pageSheet, 'Other old-site links');

    return wb;
}

// Builds the report and sends it back as an .xlsx download, with the
// summary in a header so the dashboard can show the counts too.
app.post('/old-link-report', async (req, res) => {
    const { url, apiKey, oldDomains } = req.body || {};
    if (!url || !apiKey) {
        return res.status(400).json({ message: 'Website URL and token are required.' });
    }
    if (!parseOldDomains(oldDomains).length) {
        return res.status(400).json({ message: 'Enter the old website\'s domain (for example co.monmouth.nj.us).' });
    }

    const startedAt = Date.now();
    logAsset('info', `Old-link report started - site ${url} - old domain(s) ${oldDomains}`);
    try {
        const report = await buildOldLinkReport({ siteURL: url, accessToken: apiKey, oldDomains });
        const wb = buildOldLinkReportWorkbook(report);
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        logAsset('info', `Old-link report finished in ${Math.round((Date.now() - startedAt) / 1000)}s - ${report.summary.itemsScanned} items, ${report.summary.fileLinks} file links (${report.summary.fileLinksWithNoMatch} with no matching asset)`);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="OldLinkReport-${url.replace(/[^a-z0-9.-]/gi, '_')}.xlsx"`);
        res.setHeader('X-Report-Summary', encodeURIComponent(JSON.stringify(report.summary)));
        res.send(buffer);
    } catch (error) {
        logAsset('error', `Old-link report failed: ${error.message}`);
        res.status(500).json({ message: error.message || 'Could not build the report.' });
    }
});


// --- Single Dynamic Route ---

// The Excel file arrives as multipart/form-data under the field name "file".
// url + apiKey travel as ordinary text fields alongside it (req.body).
app.post('/upload/:type', (req, res) => {
    uploadSingle(req, res, async (uploadErr) => {
        // multer errors (wrong type, file too large, etc.) surface here.
        if (uploadErr) {
            return res.status(400).json({ message: uploadErr.message });
        }

        const { url, apiKey, sheetOverride } = req.body;
        const { type } = req.params;

        if (!uploadConfig[type]) {
            return res.status(400).json({ message: `Invalid upload type: ${type}` });
        }

        if (!req.file) {
            return res.status(400).json({ message: 'No Excel file uploaded. Choose a file and try again.' });
        }

        console.log(`Processing [${type}] - URL: ${url} - file: ${req.file.originalname}`);

        try {
            // Parse the workbook straight from the uploaded buffer (no disk read).
            const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
            const { successes, failures, missingCategories } = await processUpload(type, url, apiKey, workbook, 10, sheetOverride);

            // Rows still upload when a category doesn't exist on the site
            // (the category is dropped, not sent), but that's a silent data
            // change unless it's said out loud.
            const categoryNote = missingCategories && missingCategories.length
                ? ` Note: these categories don't exist on this site and were skipped: ${missingCategories.join(', ')}.`
                : '';

            if (failures.length === 0) {
                return res.json({ message: `Upload complete. Processed ${successes.length} ${type}.${categoryNote}` });
            }

            // Partial (or total) failure: report counts AND the actual API errors
            // so you can see exactly which field each content type is rejecting,
            // instead of a silent "Processed 0".
            return res.status(207).json({
                message: `Processed ${successes.length} of ${successes.length + failures.length} ${type}. ${failures.length} failed.${categoryNote}`,
                errors: failures.slice(0, 10).map(f => `Row ${f.row}: ${f.error}`)
            });
        } catch (error) {
            console.error(error);
            res.status(500).json({ message: error.message || "Upload failed. Check server logs." });
        }
    });
});


// --- Reference Data Loader ---

// GET https://content.civicplus.com/api/apps/<site>/<endpoint>
// Returns a Map of lowercased name -> id for quick matching.
async function fetchReferenceData(siteURL, endpoint, accessToken) {
    const apiURL = `https://content.civicplus.com/api/apps/${siteURL}/${endpoint}`;

    const response = await fetch(apiURL, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        }
    });

    if (!response.ok) {
        throw new Error(`Failed to load ${endpoint}: HTTP ${response.status}: ${await response.text()}`);
    }

    const json = await response.json();
    const lookup = new Map();

    (json.items || []).forEach(item => {
        if (item?.name) lookup.set(String(item.name).trim().toLowerCase(), item.id);
    });

    console.log(`Loaded ${lookup.size} ${endpoint} entries for ${siteURL}`);
    return lookup;
}

// --- Content & asset name lookups (for reference-resolving fields) ---
//
// permissionSet/categories above are CivicPlus's own simple taxonomy lists
// ({items:[{id,name}]} at /api/apps/:site/:endpoint). staff.department,
// departments.parentdepartment/staffdirectory, and the asset fields
// (news.newsasset, calendar.attachments/submission-pdf) are different: they
// reference other CONTENT items or uploaded ASSETS, which are Squidex's own
// content-list and asset-list endpoints, not that taxonomy shape. Grounded
// in Squidex's public docs (docs.squidex.io) — content.civicplus.com is
// confirmed built on Squidex (see the "Empty arrays are usually accepted by
// Squidex" note in uploadConfig.departments) — but not verified against
// CivicPlus's specific deployment, so these fail soft: a lookup that can't
// be loaded just leaves that reference type unresolved, the same as an
// unresolved schema/field name elsewhere in this file.

// Each content type stores its own "name" under different data fields, so
// there's no single generic key to read one back out.
const CONTENT_NAME_EXTRACTORS = {
    departments: (item) => item.data?.department?.en ?? item.data?.department?.iv,
    staff: (item) => {
        const first = item.data?.firstname?.en ?? item.data?.firstname?.iv ?? '';
        const last = item.data?.lastname?.en ?? item.data?.lastname?.iv ?? '';
        return `${first} ${last}`.trim() || undefined;
    }
};

// Which extra lookups each content type's reference fields need, loaded
// once per processUpload run (see the "Content & asset reference lookups"
// step below), the same way permissionSet/categories already are.
const CONTENT_REFERENCE_NEEDS = {
    staff: ['departments'],
    departments: ['departments', 'staff'],
    news: ['assets'],
    calendar: ['assets']
};

// Lists existing content of a schema and builds a name -> id lookup via
// `extractName`. Mirrors the Content API's list endpoint (the same base
// path this file already POSTs new records to).
async function fetchContentNameLookup(siteURL, schemaSlug, accessToken, extractName) {
    const apiURL = `https://content.civicplus.com/api/content/${siteURL}/${schemaSlug}?$top=200`;
    const response = await fetch(apiURL, {
        method: 'GET',
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });
    if (!response.ok) {
        console.warn(`Could not list existing "${schemaSlug}" content for "${siteURL}" (HTTP ${response.status}) — references to it will be left unresolved.`);
        return new Map();
    }
    const json = await response.json();
    const items = json.items || (Array.isArray(json) ? json : []);
    const lookup = new Map();
    items.forEach(item => {
        const name = extractName(item);
        if (name) lookup.set(String(name).trim().toLowerCase(), item.id);
    });
    return lookup;
}

// Lists existing assets and builds a fileName -> id lookup. Mirrors
// postAssetFile's own /assets endpoint (POST there, GET here) — the
// Documents tab is how a file gets uploaded and named in the first place.
async function fetchAssetNameLookup(siteURL, accessToken) {
    const apiURL = `https://content.civicplus.com/api/apps/${siteURL}/assets?$top=200`;
    const response = await fetch(apiURL, {
        method: 'GET',
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });
    if (!response.ok) {
        console.warn(`Could not list existing assets for "${siteURL}" (HTTP ${response.status}) — asset references will be left unresolved.`);
        return new Map();
    }
    const json = await response.json();
    const items = json.items || (Array.isArray(json) ? json : []);
    const lookup = new Map();
    items.forEach(item => {
        const fileName = item.fileName || item.slug;
        if (fileName) lookup.set(String(fileName).trim().toLowerCase(), item.id);
    });
    return lookup;
}


// Best-effort list of the schema (content-type) slugs that actually exist in an
// app. Used both to resolve which candidate slug this site uses and to turn an
// opaque 404 into "here are the names that exist." Returns null if the endpoint
// isn't reachable with this token, in which case we fall back to the first
// candidate and let the POST report its own 404.
async function fetchSchemaNames(siteURL, accessToken) {
    try {
        const response = await fetch(`https://content.civicplus.com/api/apps/${siteURL}/schemas`, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' }
        });
        if (!response.ok) return null;
        const json = await response.json();
        const list = Array.isArray(json) ? json : (json.items || []);
        const names = list
            .map(s => (typeof s === 'string' ? s : (s?.name || s?.slug)))
            .filter(Boolean);
        return names.length ? names : null;
    } catch {
        return null;
    }
}


// Read the resolved schema's fields as a map: normalizedName -> { name, partitioning }.
// "partitioning" is "invariant" (=> { iv }) or a language (=> { en }); we only
// need the binary distinction. Returns null if the schema/fields can't be read,
// in which case field conforming is skipped and the mapper's output is sent as-is
// (so nothing that currently works can break).
async function fetchSchemaFieldMap(siteURL, schema, accessToken) {
    try {
        const response = await fetch(`https://content.civicplus.com/api/apps/${siteURL}/schemas/${schema}`, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' }
        });
        if (!response.ok) return null;
        const json = await response.json();

        // Be liberal about where the field array lives across API shapes.
        const fields = json.fields || json?.schema?.fields || json?.data?.fields;
        if (!Array.isArray(fields) || !fields.length) return null;

        const map = new Map();
        for (const f of fields) {
            const name = f?.name || f?.slug;
            if (!name) continue;
            const partitioning = String(f?.partitioning || 'invariant').toLowerCase();
            map.set(normalizeSlug(name), { name, partitioning });
        }
        return map.size ? map : null;
    } catch {
        return null;
    }
}


// Schemas are app-wide, so fetch the list once per site and reuse it across all
// content types in the same run (uploading 11 types to one site => 1 /schemas
// call instead of 11). null results are NOT cached, so a transient failure is
// retried on the next request.
const SCHEMA_CACHE_TTL_MS = 5 * 60 * 1000;
const schemaCache = new Map();      // siteURL -> { names, expires }
const schemaFieldCache = new Map(); // `${site}::${schema}` -> { map, expires }

async function getSchemaNames(siteURL, accessToken) {
    const cached = schemaCache.get(siteURL);
    if (cached && cached.expires > Date.now()) return cached.names;

    const names = await fetchSchemaNames(siteURL, accessToken);
    if (names) schemaCache.set(siteURL, { names, expires: Date.now() + SCHEMA_CACHE_TTL_MS });
    return names; // don't cache null — retry next time
}

async function getSchemaFields(siteURL, schema, accessToken) {
    const cacheKey = `${siteURL}::${schema}`;
    const cached = schemaFieldCache.get(cacheKey);
    if (cached && cached.expires > Date.now()) return cached.map;

    const map = await fetchSchemaFieldMap(siteURL, schema, accessToken);
    if (map) schemaFieldCache.set(cacheKey, { map, expires: Date.now() + SCHEMA_CACHE_TTL_MS });
    return map; // don't cache null — retry next time
}


// --- Connection check (for the "Test Connection" button) ---
//
// Deliberately separate from fetchSchemaNames/getSchemaNames above: those
// are best-effort and collapse every failure to `null` on purpose, because
// their job is "quietly fall back so an upload can still try." This is the
// opposite job — surface EXACTLY what went wrong, in a sentence a
// non-technical person can act on — so it makes its own uncached request
// rather than reusing them. Never caches, since the whole point is "is it
// working right now."

// One request to /schemas — reachability and the token are the only two
// things this can tell apart (this endpoint doesn't distinguish "site
// exists but token is wrong" from "site doesn't exist" any more granularly
// than that), but that's already the two most common real-world problems.
async function checkSiteConnection(siteURL, accessToken) {
    const apiURL = `https://content.civicplus.com/api/apps/${siteURL}/schemas`;
    let response;
    try {
        response = await fetch(apiURL, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' }
        });
    } catch (networkErr) {
        return {
            ok: false,
            status: null,
            reason: 'network',
            message: `Could not reach "${siteURL}" — check the site app name and the network connection. (${networkErr.message})`
        };
    }

    if (response.status === 401 || response.status === 403) {
        return {
            ok: false,
            status: response.status,
            reason: 'auth',
            message: `Token was rejected (HTTP ${response.status}) — check that the token is correct, current, and has permission for this site.`
        };
    }
    if (response.status === 404) {
        return {
            ok: false,
            status: 404,
            reason: 'not-found',
            message: `Site "${siteURL}" was not found (HTTP 404) — check the site app name.`
        };
    }
    if (!response.ok) {
        const body = await response.text();
        return {
            ok: false,
            status: response.status,
            reason: 'error',
            message: `Unexpected response (HTTP ${response.status}): ${body.slice(0, 300)}`
        };
    }

    const json = await response.json();
    const list = Array.isArray(json) ? json : (json.items || []);
    const schemaNames = list.map(s => (typeof s === 'string' ? s : (s?.name || s?.slug))).filter(Boolean);

    return {
        ok: true,
        status: response.status,
        reason: 'ok',
        message: `Connected — token accepted, found ${schemaNames.length} schema(s) on this site.`,
        schemaNames
    };
}

// The full report the /check-connection route returns. Read-only start to
// finish — this never creates, updates, or publishes anything, so it's
// safe to click at any time with no side effects on the real site.
async function checkConnection(siteURL, accessToken) {
    const connection = await checkSiteConnection(siteURL, accessToken);

    // Can't meaningfully check content types or reference data without a
    // working connection, so stop here and let the caller show just that.
    if (!connection.ok) {
        return { connection, contentTypes: [], referenceData: [], schemas: [] };
    }

    const contentTypes = Object.entries(uploadConfig).map(([type, config]) => {
        const candidates = getEndpointCandidates(config);
        const resolvedEndpoint = resolveEndpoint(candidates, connection.schemaNames);
        return {
            type,
            label: config.label || type,
            ok: !!resolvedEndpoint,
            resolvedEndpoint
        };
    });

    // Every schema on the site, flagged by whether a content type in this
    // tool already points at it. The unclaimed ones are the useful part:
    // they're the candidates when registering a new content type, so the
    // GUI offers them as autocomplete on the schema slug field.
    const claimed = new Set(contentTypes.filter(ct => ct.ok).map(ct => ct.resolvedEndpoint));
    const schemas = (connection.schemaNames || []).map(name => ({
        name,
        used: claimed.has(name)
    }));

    const referenceData = await Promise.all(
        ['permissionSet', 'categories'].map(async (endpoint) => {
            try {
                const lookup = await fetchReferenceData(siteURL, endpoint, accessToken);
                return { endpoint, ok: true, count: lookup.size };
            } catch (error) {
                return { endpoint, ok: false, message: error.message };
            }
        })
    );

    return { connection, contentTypes, referenceData, schemas };
}

app.post('/check-connection', async (req, res) => {
    const { url, apiKey } = req.body;
    if (!url || !apiKey) {
        return res.status(400).json({ message: 'Website URL and token are required.' });
    }

    try {
        const report = await checkConnection(url, apiKey);
        res.json(report);
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: error.message || 'Connection check failed. Check server logs.' });
    }
});


// --- Content types & workbook inspection (for the dynamic GUI) ---

// The GUI builds its content-type list from this rather than hardcoding
// nine buttons in the HTML, so a newly-defined custom type appears
// without touching index.html.
app.get('/content-types', (req, res) => {
    const types = Object.entries(uploadConfig).map(([type, config]) => ({
        type,
        label: config.label || type,
        description: config.description || '',
        sheetName: config.sheetName || null,
        isCustom: !!config.isCustom
    }));
    res.json({ types });
});

// Reads a workbook WITHOUT uploading anything: returns each sheet's name,
// its column headers, and its row count, plus which content type (if any)
// each sheet matches. This is what drives the per-type sheet pickers and
// the "unmatched sheets" list in the GUI.
app.post('/inspect-workbook', (req, res) => {
    uploadSingle(req, res, (uploadErr) => {
        if (uploadErr) return res.status(400).json({ message: uploadErr.message });
        if (!req.file) return res.status(400).json({ message: 'No Excel file uploaded. Choose a file and try again.' });

        try {
            const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });

            // Which sheet each known type resolves to, so the GUI can
            // preselect the right option in every dropdown.
            //
            // Name matching ONLY here (no sheetIndex fallback): in a
            // partial workbook — say just Articles and FAQs — the
            // positional fallback would make unrelated types claim
            // whatever sheet happens to sit at their index (resources,
            // index 1, would "match" the FAQs sheet), which in turn hides
            // genuinely unmatched sheets from the custom-import list. The
            // fallback still applies at upload time, where the user has
            // deliberately picked one type.
            const matchedByType = {};
            for (const [type, config] of Object.entries(uploadConfig)) {
                if (!config.sheetName) continue;
                const normalized = normalizeSlug(config.sheetName);
                const matched = workbook.SheetNames.find(name => normalizeSlug(name) === normalized);
                if (matched) matchedByType[type] = matched;
            }
            const claimedSheets = new Set(Object.values(matchedByType));
            // The Assets sheet drives the Documents upload, not a content
            // type — keep it out of the "unrecognized sheets" list.
            const assetSheet = findAssetSheetName(workbook.SheetNames);
            if (assetSheet) claimedSheets.add(assetSheet);

            const sheets = workbook.SheetNames.map(name => {
                const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
                // header:1 gives the raw first row, so a column stays
                // visible even when every cell under it is blank (which
                // sheet_to_json would otherwise drop entirely).
                const headerRows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
                const columns = (headerRows[0] || []).map(h => String(h).trim()).filter(Boolean);
                return {
                    name,
                    columns,
                    rowCount: rows.length,
                    matchedType: Object.keys(matchedByType).find(t => matchedByType[t] === name) || null,
                    unmatched: !claimedSheets.has(name)
                };
            });

            res.json({ sheets, matchedByType });
        } catch (error) {
            console.error(error);
            res.status(500).json({ message: error.message || 'Could not read the workbook.' });
        }
    });
});

// Saves a new custom content type (or updates one by the same key) and
// registers it immediately, so it works without a restart — and persists,
// so it also survives the restart the .bat file does.
app.post('/custom-types', (req, res) => {
    const { label, endpoint, sheetName, columns } = req.body;

    if (!label || !endpoint || !sheetName) {
        return res.status(400).json({ message: 'A label, schema slug, and sheet name are all required.' });
    }
    // columns is optional: an empty list means "use whatever columns each
    // row has", which is how a hand-registered type works (no workbook on
    // hand to read headers from). It must still be an ARRAY if sent.
    if (columns !== undefined && !Array.isArray(columns)) {
        return res.status(400).json({ message: 'Columns must be a list.' });
    }

    const key = `custom_${normalizeSlug(label)}`;
    if (!normalizeSlug(label)) {
        return res.status(400).json({ message: 'The label needs at least one letter or number.' });
    }
    // Never let a custom type shadow a built-in one — that would silently
    // change where an existing button uploads to.
    if (uploadConfig[key] && !uploadConfig[key].isCustom) {
        return res.status(400).json({ message: `"${label}" collides with a built-in content type. Pick a different name.` });
    }

    try {
        const customTypes = loadCustomTypes();
        customTypes[key] = { label, endpoint, sheetName, columns: columns || [], description: req.body.description };
        saveCustomTypes(customTypes);
        registerCustomTypes(customTypes);

        res.json({
            message: `Saved "${label}" — it now reads the "${sheetName}" sheet and uploads to the "${endpoint}" schema.`,
            type: key
        });
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: error.message || 'Could not save the custom type.' });
    }
});

app.delete('/custom-types/:key', (req, res) => {
    const { key } = req.params;
    try {
        const customTypes = loadCustomTypes();
        if (!customTypes[key]) {
            return res.status(404).json({ message: `No custom type named "${key}".` });
        }
        delete customTypes[key];
        saveCustomTypes(customTypes);
        delete uploadConfig[key];
        res.json({ message: `Removed custom type "${key}".` });
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: error.message || 'Could not remove the custom type.' });
    }
});


// --- Generic Engine ---

async function processUpload(type, siteURL, accessToken, workbook, batchSize = 10, sheetOverride) {


    const config = uploadConfig[type];

    const sheetName = resolveSheetName(workbook.SheetNames, config, sheetOverride);
    if (!sheetName) {
        throw new Error(`The uploaded workbook has no sheet matching "${config.sheetName || type}" (and no sheet at fallback index ${config.sheetIndex}) for "${type}". It has ${workbook.SheetNames.length} sheet(s): ${workbook.SheetNames.join(', ')}.`);
    }
    const excelData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // --- Resolve the schema slug THIS site uses ---
    // Instead of trusting a single hardcoded slug, resolve the configured
    // candidate(s) against the schemas that actually exist on this site. This is
    // what makes one codebase work across sites with different slugs, and turns a
    // wrong slug into a precise, actionable error instead of a blind 404.
    const candidates = getEndpointCandidates(config);
    if (!candidates.length) {
        throw new Error(`No endpoint(s) configured for "${type}". Add an "endpoint" string or "endpoints" array in uploadConfig.`);
    }

    const schemaNames = await getSchemaNames(siteURL, accessToken);
    let endpoint;
    if (schemaNames) {
        endpoint = resolveEndpoint(candidates, schemaNames);
        if (!endpoint) {
            const close = suggestClosestSchemas(candidates, schemaNames);
            throw new Error(
                `None of the known slugs for "${type}" (${candidates.join(', ')}) exist in app "${siteURL}". ` +
                (close.length
                    ? `Closest schemas that DO exist: ${close.join(', ')}. `
                    : `Existing schemas: ${schemaNames.join(', ')}. `) +
                `Add the correct slug to endpoints[] for "${type}" in uploadConfig ` +
                `(or set ENDPOINT_ALIASES_${type.toUpperCase()}).`
            );
        }
        if (endpoint !== candidates[0]) {
            console.log(`[${type}] resolved to schema "${endpoint}" for ${siteURL} (preferred "${candidates[0]}").`);
        }
    } else {
        // Couldn't list schemas (token can't reach /schemas). Fall back to the
        // first candidate and let the POST report its own 404 if it's wrong.
        endpoint = candidates[0];
        console.warn(`Could not list schemas for "${siteURL}" — trying "${endpoint}" for "${type}". A 404 below means it's the wrong slug for this site.`);
    }

    const apiURL = `https://content.civicplus.com/api/content/${siteURL}/${endpoint}`;

    // --- Resolve field names THIS site uses ---
    // Read the resolved schema's real field names once, then rename each row's
    // data keys per fieldAliases. If the fields can't be read, conforming is
    // skipped and the mapper's field names are sent as-is (current behavior).
    // Partition conforming (fixing { iv } vs { en } to match the field) is
    // opt-in: set conformPartitions:true on the type, or CONFORM_PARTITIONS=1.
    const conformPartitions = config.conformPartitions === true || !!process.env.CONFORM_PARTITIONS;
    const fieldMap = process.env.DISABLE_SCHEMA_CONFORM
        ? null
        : await getSchemaFields(siteURL, endpoint, accessToken);
    if (!fieldMap && config.fieldAliases && !process.env.DISABLE_SCHEMA_CONFORM) {
        console.warn(`Could not read fields for schema "${endpoint}" on "${siteURL}" — field-name aliases won't be applied (sending the mapper's field names as-is).`);
    }

    // Load every reference list this type needs, once, before uploading.
    const lookups = {};
    for (const refEndpoint of (config.referenceData || DEFAULT_REFERENCE_DATA)) {
        lookups[refEndpoint] = await fetchReferenceData(siteURL, refEndpoint, accessToken);
    }

    // Checked up front so the result can name every category that doesn't
    // exist on the site, rather than the user noticing later that content
    // uploaded without its categories.
    const missingCategories = collectMissingCategories(excelData, lookups.categories);
    if (missingCategories.length) {
        console.warn(`[${type}] ${missingCategories.length} category name(s) not found on ${siteURL} and will be skipped: ${missingCategories.join(', ')}`);
    }

    // --- Content & asset reference lookups ---
    // See the comment above fetchContentNameLookup/fetchAssetNameLookup for
    // why these need a different mechanism than permissionSet/categories.
    for (const need of (CONTENT_REFERENCE_NEEDS[type] || [])) {
        if (lookups[need]) continue; // e.g. departments already loaded its own self-reference lookup
        if (need === 'assets') {
            lookups.assets = await fetchAssetNameLookup(siteURL, accessToken);
        } else {
            const targetConfig = uploadConfig[need];
            const targetCandidates = getEndpointCandidates(targetConfig);
            const targetEndpoint = schemaNames ? resolveEndpoint(targetCandidates, schemaNames) : targetCandidates[0];
            lookups[need] = targetEndpoint
                ? await fetchContentNameLookup(siteURL, targetEndpoint, accessToken, CONTENT_NAME_EXTRACTORS[need])
                : new Map();
        }
    }

    const successes = [];
    const failures = [];

    for (let i = 0; i < excelData.length; i += batchSize) {
        const batch = excelData.slice(i, i + batchSize);
        console.log(`Processing batch ${Math.floor(i / batchSize) + 1} of ${Math.ceil(excelData.length / batchSize)}`);

        const batchPromises = batch.map((entry, j) => {
            // Excel row number as the user sees it (1-based + header row).
            const rowNumber = i + j + 2;
            const requestData = config.mapPayload(entry, lookups);

            // --- DYNAMIC EXTRA FIELD INJECTION ---
            // Any column not in standardColumns is treated as an invariant custom
            // field. Empty cells are skipped so a blank column never sends { iv: "" }
            // (or, worse, an empty partition the API can reject).
            Object.keys(entry).forEach(key => {
                if (config.standardColumns && !config.standardColumns.includes(key)) {
                    const value = entry[key];
                    if (value !== undefined && value !== null && value !== '') {
                        requestData.data[key] = { iv: value };
                    }
                }
            });

            // --- FIELD-NAME CONFORMING ---
            // Rename data (and searchFields / titles, if present) to the field
            // names this site's schema actually uses. No-op when a key already
            // matches, so working sites are unaffected.
            if (fieldMap) {
                requestData.data = conformFields(requestData.data, fieldMap, config, { conformPartitions });
                if (requestData.searchFields) requestData.searchFields = conformFields(requestData.searchFields, fieldMap, config, { conformPartitions });
                if (requestData.titles) requestData.titles = conformFields(requestData.titles, fieldMap, config, { conformPartitions });
            }

            // Set DEBUG_PAYLOAD=1 to see exactly what is sent for each row.
            if (process.env.DEBUG_PAYLOAD) {
                console.log(`Row ${rowNumber} payload:`, JSON.stringify(requestData, null, 2));
            }

            return fetch(apiURL, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(requestData)
            })
            .then(async response => {
                if (!response.ok) {
                    const errBody = await response.text();
                    if (response.status === 404) {
                        // "Not Found" = the schema slug in this URL doesn't exist.
                        // Resolution should have caught this already; if we get here
                        // the /schemas list was unavailable, so surface the slug we tried.
                        throw new Error(`HTTP 404 — schema "${endpoint}" not found at ${apiURL}. Verify the schema slug (it is a URL problem, not a field problem).`);
                    }
                    throw new Error(`HTTP ${response.status}: ${errBody}`);
                }
                return response.json();
            })
            .then(async data => {
                if (entry.Publish === 'Yes') {
                    const patchUrl = `https://content.civicplus.com/api/content/${siteURL}/${endpoint}/${data.id}/status/?=`;
                    try {
                        const patchRes = await fetch(patchUrl, {
                            method: 'PUT',
                            headers: {
                                'Authorization': `Bearer ${accessToken}`,
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify({ status: "Published" })
                        });

                        if (!patchRes.ok) {
                            console.error(`Publish failed for ${data.id}:`, await patchRes.text());
                        } else {
                            console.log(`Successfully published ${data.id}`);
                        }
                    } catch (err) {
                        console.error(`Network error publishing ${data.id}:`, err);
                    }
                }

                return { row: rowNumber, data };
            })
            .catch(error => {
                console.error(`Row ${rowNumber} failed:`, error.message);
                return { row: rowNumber, error: error.message };
            });
        });

        const batchResults = await Promise.all(batchPromises);
        for (const result of batchResults) {
            if (result.error) failures.push(result);
            else successes.push(result);
        }
    }

    return { successes, failures, missingCategories };
}

if (require.main === module) {
    const server = app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
    // Node cuts off any request still running after 5 minutes by default.
    // A batch of large files against a slow site can legitimately take
    // longer, and being cut off would make the dashboard count files the
    // server is still uploading as failed.
    server.requestTimeout = 30 * 60 * 1000;
}

module.exports = {
    app,
    uploadConfig,
    processUpload,
    fetchReferenceData,
    postAssetFile,
    annotateAsset,
    fetchWithRetry,
    ASSET_RETRY,
    ASSET_FILES_PER_REQUEST,
    assetBatchStatus,
    parseOldDomains,
    findOldLinks,
    collectStrings,
    itemTitle,
    buildOldLinkReport,
    buildOldLinkReportWorkbook,
    parsePublishFlag,
    parseAssetRows,
    buildAssetTemplateWorkbook,
    assetLookupCache,

    // Exported for direct unit testing below. Purely additive — nothing
    // above relies on these being exported, so this changes no behavior.
    parseTags,
    parseCategories,
    resolveId,
    permissionSetByName,
    normalizeSlug,
    resolveSheetName,
    loadCustomTypes,
    saveCustomTypes,
    buildCustomMapPayload,
    buildCustomTypeConfig,
    registerCustomTypes,
    CUSTOM_TYPES_FILE,
    getEndpointCandidates,
    resolveEndpoint,
    suggestClosestSchemas,
    getFieldCandidates,
    resolveField,
    conformPartition,
    conformFields,
    fetchSchemaNames,
    fetchSchemaFieldMap,
    getSchemaNames,
    getSchemaFields,
    resolveReferenceIds,
    fetchContentNameLookup,
    fetchAssetNameLookup,
    CONTENT_NAME_EXTRACTORS,
    CONTENT_REFERENCE_NEEDS,
    checkSiteConnection,
    checkConnection
};