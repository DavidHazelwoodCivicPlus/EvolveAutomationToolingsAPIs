const {
    parseTags,
    parseCategories,
    resolveId,
    resolveReferenceIds,
    permissionSetByName,
    normalizeSlug,
    getEndpointCandidates,
    resolveEndpoint,
    suggestClosestSchemas,
    getFieldCandidates,
    resolveField,
    conformPartition,
    conformFields
} = require('../../index.js');

// These helpers log warnings/info by design (resolveId warns on a miss,
// conformFields logs on a rename). Silence them here so test output stays
// readable; individual tests that care whether a warning fired assert on
// the spy directly, never on the message text.
let warnSpy, logSpy;
beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    logSpy = jest.spyOn(console, 'log').mockImplementation(() => {});
});
afterEach(() => {
    warnSpy.mockRestore();
    logSpy.mockRestore();
});

describe('normalizeSlug', () => {
    test.each([
        ['Enhanced-Employee', 'enhancedemployee'],
        ['enhanced_employee', 'enhancedemployee'],
        ['enhancedemployee', 'enhancedemployee'],
        ['FAX Number', 'faxnumber'],
        ['fax_number', 'faxnumber'],
        [null, ''],
        [undefined, ''],
        ['', ''],
        [123, '123']
    ])('normalizeSlug(%p) -> %p', (input, expected) => {
        expect(normalizeSlug(input)).toBe(expected);
    });
});

describe('resolveEndpoint', () => {
    test('exact match wins over normalized match', () => {
        expect(resolveEndpoint(['employee', 'enhanced-employee'], ['employee', 'other'])).toBe('employee');
    });

    test('first candidate with an exact match wins, even if a later candidate matches too', () => {
        expect(resolveEndpoint(['a', 'b'], ['b', 'a'])).toBe('a');
    });

    test('falls back to a normalized match, returning the schema\'s own spelling', () => {
        expect(resolveEndpoint(['enhanced-employee'], ['EnhancedEmployee'])).toBe('EnhancedEmployee');
    });

    test('normalized match ignores punctuation/casing differences both ways', () => {
        expect(resolveEndpoint(['bpquicklink'], ['bp-quick-link'])).toBe('bp-quick-link');
    });

    test('returns null when nothing matches, exactly or normalized', () => {
        expect(resolveEndpoint(['facility'], ['article', 'faq'])).toBeNull();
    });

    test('returns null for an empty candidate list', () => {
        expect(resolveEndpoint([], ['article'])).toBeNull();
    });
});

describe('suggestClosestSchemas', () => {
    test('returns schemas whose normalized name contains (or is contained by) a candidate, ordered by match length', () => {
        const result = suggestClosestSchemas(
            ['employee'],
            ['enhanced-employee', 'article', 'employee-record']
        );
        expect(result).toEqual(['enhanced-employee', 'employee-record']);
        expect(result).not.toContain('article');
    });

    test('returns an empty array when nothing is close', () => {
        expect(suggestClosestSchemas(['employee'], ['article', 'faq'])).toEqual([]);
    });

    test('respects the limit', () => {
        const schemaNames = ['employee1', 'employee2', 'employee3', 'employee4'];
        expect(suggestClosestSchemas(['employee'], schemaNames, 2)).toHaveLength(2);
    });
});

describe('getEndpointCandidates', () => {
    test('prefers `endpoints` array when present', () => {
        expect(getEndpointCandidates({ endpoints: ['employee', 'enhanced-employee'], endpoint: 'ignored' }))
            .toEqual(['employee', 'enhanced-employee']);
    });

    test('falls back to a single `endpoint` string', () => {
        expect(getEndpointCandidates({ endpoint: 'article' })).toEqual(['article']);
    });

    test('falls back to `endpoint` when `endpoints` is an empty array', () => {
        expect(getEndpointCandidates({ endpoints: [], endpoint: 'article' })).toEqual(['article']);
    });

    test('returns an empty array when neither is configured', () => {
        expect(getEndpointCandidates({})).toEqual([]);
    });
});

describe('getFieldCandidates', () => {
    test('includes the key itself plus any declared aliases', () => {
        const config = { fieldAliases: { phonenumber: ['phone'] } };
        expect(getFieldCandidates(config, 'phonenumber')).toEqual(['phonenumber', 'phone']);
    });

    test('is just the key when there are no aliases for it', () => {
        expect(getFieldCandidates({ fieldAliases: { other: ['x'] } }, 'phonenumber')).toEqual(['phonenumber']);
    });

    test('is just the key when the config has no fieldAliases at all', () => {
        expect(getFieldCandidates({}, 'phonenumber')).toEqual(['phonenumber']);
    });
});

describe('resolveField', () => {
    test('finds a field by an alias candidate when the exact key is not in the schema', () => {
        const fieldMap = new Map([['phone', { name: 'phone', partitioning: 'invariant' }]]);
        expect(resolveField(['phonenumber', 'phone'], fieldMap)).toEqual({ name: 'phone', partitioning: 'invariant' });
    });

    test('returns null when no candidate is in the field map', () => {
        const fieldMap = new Map([['email', { name: 'email', partitioning: 'language' }]]);
        expect(resolveField(['phonenumber', 'phone'], fieldMap)).toBeNull();
    });
});

describe('conformPartition', () => {
    test('converts a single {en} wrapper to {iv} when the field is invariant', () => {
        expect(conformPartition({ en: 'X' }, 'invariant')).toEqual({ iv: 'X' });
    });

    test('converts a single {iv} wrapper to {en} when the field is language-partitioned', () => {
        expect(conformPartition({ iv: 'X' }, 'language')).toEqual({ en: 'X' });
    });

    test('leaves the value alone when it already matches the wanted partition', () => {
        const value = { iv: 'X' };
        expect(conformPartition(value, 'invariant')).toBe(value);
    });

    test('leaves arrays alone (they are never re-wrapped)', () => {
        const value = ['a', 'b'];
        expect(conformPartition(value, 'invariant')).toBe(value);
    });

    test('leaves non-object values alone', () => {
        expect(conformPartition('plain string', 'invariant')).toBe('plain string');
        expect(conformPartition(null, 'invariant')).toBeNull();
        expect(conformPartition(undefined, 'invariant')).toBeUndefined();
    });

    test('leaves multi-key objects alone (not a simple {en}/{iv} wrapper)', () => {
        const value = { en: 'X', extra: 'Y' };
        expect(conformPartition(value, 'invariant')).toBe(value);
    });

    test('leaves an object alone whose single key is not en/iv', () => {
        const value = { foo: 'X' };
        expect(conformPartition(value, 'invariant')).toBe(value);
    });
});

describe('conformFields', () => {
    test('returns the input unchanged when there is no fieldMap (schema fields unreadable)', () => {
        const fields = { phonenumber: { en: '555' } };
        expect(conformFields(fields, null, {})).toBe(fields);
    });

    test('returns the input unchanged (including null/undefined) when there is no fieldsObj', () => {
        expect(conformFields(null, new Map(), {})).toBeNull();
        expect(conformFields(undefined, new Map(), {})).toBeUndefined();
    });

    test('renames a key to the schema\'s real spelling via fieldAliases', () => {
        const config = { fieldAliases: { phonenumber: ['phone'] } };
        const fieldMap = new Map([['phone', { name: 'phone', partitioning: 'invariant' }]]);
        const result = conformFields({ phonenumber: { en: '555' } }, fieldMap, config);
        expect(result).toEqual({ phone: { en: '555' } });
    });

    test('is a no-op when every key already matches the schema (nothing to rename)', () => {
        const config = {};
        const fieldMap = new Map([['title', { name: 'title', partitioning: 'language' }]]);
        const fields = { title: { en: 'X' } };
        expect(conformFields(fields, fieldMap, config)).toEqual({ title: { en: 'X' } });
    });

    test('only fixes the {en}/{iv} wrapper when conformPartitions is explicitly true', () => {
        const config = { fieldAliases: { phonenumber: ['phone'] } };
        const fieldMap = new Map([['phone', { name: 'phone', partitioning: 'invariant' }]]);
        const fields = { phonenumber: { en: '555' } };

        expect(conformFields(fields, fieldMap, config)).toEqual({ phone: { en: '555' } });
        expect(conformFields(fields, fieldMap, config, { conformPartitions: true })).toEqual({ phone: { iv: '555' } });
    });

    // The anti-clobber guard ("never overwrite a key already present under
    // its real name") only protects one direction. Both tests below use the
    // same two field values; only the KEY ORDER of the input object differs.
    test('anti-clobber guard is order-dependent: real key first -> both keys survive', () => {
        const config = { fieldAliases: { phonenumber: ['phone'] } };
        const fieldMap = new Map([['phone', { name: 'phone', partitioning: 'invariant' }]]);
        // "phone" (the real key) appears before "phonenumber" (the alias).
        const fields = { phone: { en: 'literal-phone-value' }, phonenumber: { en: 'aliased-value' } };

        const result = conformFields(fields, fieldMap, config);
        expect(result).toEqual({
            phone: { en: 'literal-phone-value' },
            phonenumber: { en: 'aliased-value' }
        });
        expect(warnSpy).toHaveBeenCalled();
    });

    test('anti-clobber guard is order-dependent: alias first -> the literal key silently overwrites it', () => {
        const config = { fieldAliases: { phonenumber: ['phone'] } };
        const fieldMap = new Map([['phone', { name: 'phone', partitioning: 'invariant' }]]);
        // "phonenumber" (the alias) appears before "phone" (the real key) this time.
        const fields = { phonenumber: { en: 'aliased-value' }, phone: { en: 'literal-phone-value' } };

        const result = conformFields(fields, fieldMap, config);
        // Only one "phone" key survives, and it's the literal one processed
        // last — the renamed "phonenumber" value is lost. This documents
        // CURRENT behavior; it is not necessarily desired behavior.
        expect(result).toEqual({ phone: { en: 'literal-phone-value' } });
    });
});

describe('parseTags', () => {
    test('splits and trims a comma-separated list', () => {
        expect(parseTags('a, b ,c')).toEqual(['a', 'b', 'c']);
    });

    test('returns undefined for empty/falsy input (not an empty array)', () => {
        expect(parseTags('')).toBeUndefined();
        expect(parseTags(undefined)).toBeUndefined();
        expect(parseTags(null)).toBeUndefined();
    });

    test('a single value with no commas is a one-element array', () => {
        expect(parseTags('solo')).toEqual(['solo']);
    });

    test('does NOT filter out empty entries from stray commas (contrast with parseCategories)', () => {
        expect(parseTags('a,,b')).toEqual(['a', '', 'b']);
    });
});

describe('parseCategories', () => {
    const lookup = new Map([['fire', 'cat-fire-id'], ['social', 'cat-social-id']]);

    test('splits, trims, and resolves each name to an id', () => {
        expect(parseCategories('Fire, Social', lookup)).toEqual([
            { id: 'cat-fire-id', name: 'Fire' },
            { id: 'cat-social-id', name: 'Social' }
        ]);
    });

    test('returns undefined for empty/falsy input', () => {
        expect(parseCategories('', lookup)).toBeUndefined();
        expect(parseCategories(undefined, lookup)).toBeUndefined();
    });

    test('DOES filter out empty entries from stray/double commas (contrast with parseTags)', () => {
        expect(parseCategories('Fire,,Social', lookup)).toEqual([
            { id: 'cat-fire-id', name: 'Fire' },
            { id: 'cat-social-id', name: 'Social' }
        ]);
    });

    // A live run against a real site returned "Categories: One or more
    // Categories do not exist in App" and rejected the ENTIRE row when an
    // unmatched category was sent with an empty id. Dropping it instead
    // means the row uploads with its other fields intact.
    test('a name with no lookup match is dropped rather than sent with an empty id', () => {
        expect(parseCategories('Unknown Category', lookup)).toBeUndefined();
    });

    test('keeps the matched categories and drops only the unmatched ones', () => {
        expect(parseCategories('Fire, Unknown Category, Social', lookup)).toEqual([
            { id: 'cat-fire-id', name: 'Fire' },
            { id: 'cat-social-id', name: 'Social' }
        ]);
    });

    test('is undefined (not an empty array) when nothing matches, so the field is omitted entirely', () => {
        expect(parseCategories('City Council,  Allison', new Map())).toBeUndefined();
    });

    test('handles real-world irregular spacing, e.g. "City Council,  Allison" (double space)', () => {
        const spacedLookup = new Map([['city council', 'cat-cc'], ['allison', 'cat-al']]);
        expect(parseCategories('City Council,  Allison', spacedLookup)).toEqual([
            { id: 'cat-cc', name: 'City Council' },
            { id: 'cat-al', name: 'Allison' }
        ]);
    });
});

describe('resolveId', () => {
    const lookup = new Map([['fire dept', 'dept-123']]);

    test('resolves a name via the lookup, case- and whitespace-insensitively', () => {
        expect(resolveId('  Fire Dept  ', lookup)).toBe('dept-123');
    });

    test('returns undefined for empty/falsy input without touching the lookup', () => {
        expect(resolveId('', lookup)).toBeUndefined();
        expect(resolveId(undefined, lookup)).toBeUndefined();
        expect(resolveId(null, lookup)).toBeUndefined();
    });

    test('passes a raw GUID through as-is when the lookup has no match', () => {
        expect(resolveId('00000000-0000-0000-0000-000000000000', new Map())).toBe('00000000-0000-0000-0000-000000000000');
    });

    test('a GUID that DOES match the lookup by name still returns the id from the lookup', () => {
        const guidLikeName = '11111111-1111-1111-1111-111111111111';
        const lookupWithGuidName = new Map([[guidLikeName.toLowerCase(), 'resolved-id']]);
        expect(resolveId(guidLikeName, lookupWithGuidName)).toBe('resolved-id');
    });

    test('returns undefined and warns when there is no match and the value is not a GUID', () => {
        expect(resolveId('Nonexistent Name', lookup)).toBeUndefined();
        expect(warnSpy).toHaveBeenCalledTimes(1);
    });

    test('works when no lookup is provided at all (undefined lookup)', () => {
        expect(resolveId('Anything', undefined)).toBeUndefined();
    });
});

describe('resolveReferenceIds', () => {
    const lookup = new Map([['fire department', 'dept-fire-id'], ['parks and recreation', 'dept-parks-id']]);

    test('resolves a single name to a one-element array (Reference/Asset fields are always arrays)', () => {
        expect(resolveReferenceIds('Fire Department', lookup)).toEqual(['dept-fire-id']);
    });

    test('resolves a comma-separated list of names to an array of ids, in order', () => {
        expect(resolveReferenceIds('Fire Department, Parks and Recreation', lookup)).toEqual(['dept-fire-id', 'dept-parks-id']);
    });

    test('is case- and whitespace-insensitive, same as resolveId', () => {
        expect(resolveReferenceIds('  fire department  ', lookup)).toEqual(['dept-fire-id']);
    });

    test('returns an empty array (not undefined) for empty/falsy input', () => {
        expect(resolveReferenceIds('', lookup)).toEqual([]);
        expect(resolveReferenceIds(undefined, lookup)).toEqual([]);
        expect(resolveReferenceIds(null, lookup)).toEqual([]);
    });

    test('drops a name with no match rather than including undefined, and warns', () => {
        expect(resolveReferenceIds('Fire Department, Made Up Place', lookup)).toEqual(['dept-fire-id']);
        expect(warnSpy).toHaveBeenCalledTimes(1);
    });

    test('an empty array when every name is unmatched (not a crash)', () => {
        expect(resolveReferenceIds('Nowhere, Nothing', lookup)).toEqual([]);
    });
});

describe('permissionSetByName', () => {
    test('resolves the id from lookups.permissionSet and keeps the raw name', () => {
        const lookups = { permissionSet: new Map([['general', 'ps-general']]) };
        expect(permissionSetByName({ PermissionSet: 'General' }, lookups)).toEqual({ id: 'ps-general', name: 'General' });
    });

    test('keeps the name even when it does not resolve to an id', () => {
        const lookups = { permissionSet: new Map() };
        expect(permissionSetByName({ PermissionSet: 'Mystery' }, lookups)).toEqual({ id: undefined, name: 'Mystery' });
    });
});
