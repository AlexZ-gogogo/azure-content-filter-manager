const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createApp() {
    const selectedModels = [];
    const targets = new Map();
    const document = {
        addEventListener() {},
        querySelectorAll(selector) {
            if (selector === '.apply-model-cb:checked') return selectedModels;
            return [];
        },
        querySelector(selector) {
            const idx = Number(selector.match(/data-res-index="(\d+)"/)?.[1]);
            return { value: targets.get(idx) || '' };
        }
    };
    const context = vm.createContext({ document, Set, Map });
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../public/app.js'), 'utf8'), context);
    vm.runInContext(`
        selectedSubscriptions = new Set(['sub-1', 'sub-2']);
        allResources = [
            { name: 'res-1', _subId: 'sub-1', _idx: 0 },
            { name: 'res-2', _subId: 'sub-2', _idx: 1 }
        ];
        selectedResources = new Set([0, 1]);
        applyScope = new Set([0, 1]);
        applyData = new Map([
            [0, { deployments: [{ name: 'new', properties: { raiPolicyName: 'Microsoft.Default' } }], policies: [{ name: 'custom' }] }],
            [1, { deployments: [{ name: 'old', properties: { raiPolicyName: 'custom' } }], policies: [{ name: 'custom' }] }]
        ]);
    `, context);
    selectedModels.push({ dataset: { resIndex: '0', depIndex: '0' } }, { dataset: { resIndex: '1', depIndex: '0' } });
    return { context, targets };
}

test('deselected subscriptions and resources never enter the apply scope', () => {
    const { context } = createApp();
    vm.runInContext(`selectedSubscriptions.delete('sub-2')`, context);
    assert.deepEqual(Array.from(vm.runInContext('getSelectedResourceIndices()', context)), [0]);
    vm.runInContext(`selectedResources.delete(0)`, context);
    assert.deepEqual(Array.from(vm.runInContext('getSelectedResourceIndices()', context)), []);
});

test('only valid changes are planned; unchanged deployments are skipped', () => {
    const { context, targets } = createApp();
    targets.set(0, 'custom');
    targets.set(1, 'custom');
    const result = vm.runInContext('getApplyChanges()', context);
    assert.equal(result.changes.length, 1);
    assert.equal(result.changes[0].dep.name, 'new');
    assert.equal(result.unchanged, 1);
    assert.equal(result.missingTarget, 0);
});

test('missing targets block the plan and stale subscription scope is ignored', () => {
    const { context, targets } = createApp();
    targets.set(1, 'custom');
    vm.runInContext(`selectedSubscriptions.delete('sub-2')`, context);
    const result = vm.runInContext('getApplyChanges()', context);
    assert.equal(result.changes.length, 0);
    assert.equal(result.missingTarget, 1);
    assert.equal(result.unchanged, 0);
});
