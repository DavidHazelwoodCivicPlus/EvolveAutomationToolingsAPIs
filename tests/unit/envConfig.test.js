// uploadConfig's endpoint/field-alias lists can be extended via env vars
// (ENDPOINT_ALIASES_<TYPE>, FIELD_ALIASES_<TYPE>) — but that extension runs
// in a plain top-level `for` loop in index.js, executed once when the
// module is first required, not inside a function. To observe its effect,
// the env var has to be set BEFORE index.js is (re-)required, so every test
// here uses jest.resetModules() + a fresh require() rather than importing
// index.js once at the top of the file.

let originalEnv, warnSpy, logSpy;
const ENV_KEYS = ['ENDPOINT_ALIASES_STAFF', 'FIELD_ALIASES_DEPARTMENTS'];

beforeEach(() => {
    originalEnv = {};
    for (const k of ENV_KEYS) originalEnv[k] = process.env[k];
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
    logSpy = jest.spyOn(console, 'log').mockImplementation(() => {});
});

afterEach(() => {
    for (const k of ENV_KEYS) {
        if (originalEnv[k] === undefined) delete process.env[k];
        else process.env[k] = originalEnv[k];
    }
    jest.resetModules();
    warnSpy.mockRestore();
    logSpy.mockRestore();
});

test('without either env var set, endpoints/fieldAliases are just the built-in values (baseline)', () => {
    delete process.env.ENDPOINT_ALIASES_STAFF;
    delete process.env.FIELD_ALIASES_DEPARTMENTS;
    jest.resetModules();

    const { uploadConfig } = require('../../index.js');
    expect(uploadConfig.staff.endpoints).toEqual(['employee', 'enhanced-employee']);
    expect(uploadConfig.departments.fieldAliases.faxnumber).toEqual(['fax', 'faxnumber']);
});

test('ENDPOINT_ALIASES_STAFF appends extra candidates after the built-in ones (lower priority)', () => {
    process.env.ENDPOINT_ALIASES_STAFF = 'employe-records, staff-members';
    jest.resetModules();

    const { uploadConfig } = require('../../index.js');
    expect(uploadConfig.staff.endpoints).toEqual([
        'employee', 'enhanced-employee', 'employe-records', 'staff-members'
    ]);
});

test('FIELD_ALIASES_DEPARTMENTS merges extra aliases onto the existing list instead of replacing it', () => {
    process.env.FIELD_ALIASES_DEPARTMENTS = JSON.stringify({ faxnumber: ['faxno'] });
    jest.resetModules();

    const { uploadConfig } = require('../../index.js');
    expect(uploadConfig.departments.fieldAliases.faxnumber).toEqual(
        expect.arrayContaining(['fax', 'faxnumber', 'faxno'])
    );
});

test('invalid JSON in FIELD_ALIASES_* is ignored (warns) rather than crashing module load', () => {
    process.env.FIELD_ALIASES_DEPARTMENTS = 'not valid json {{{';
    jest.resetModules();

    expect(() => require('../../index.js')).not.toThrow();
});
