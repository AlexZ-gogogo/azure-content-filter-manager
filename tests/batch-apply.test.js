const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createApp() {
    const selectedModels = [];
    const targets = new Map();
    const modelList = { innerHTML: '', querySelectorAll() { return []; } };
    const document = {
        addEventListener() {},
        getElementById(id) {
            if (id === 'apply-models-list') return modelList;
            if (id === 'apply-model-search') return { value: '' };
            if (id === 'apply-only-default') return { checked: false };
            return null;
        },
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
            { name: 'res-1', _subId: 'sub-1', _idx: 0, id: '/subscriptions/sub-1/resourceGroups/rg/providers/Microsoft.CognitiveServices/accounts/res-1' },
            { name: 'res-2', _subId: 'sub-2', _idx: 1, id: '/subscriptions/sub-2/resourceGroups/rg/providers/Microsoft.CognitiveServices/accounts/res-2' }
        ];
        selectedResources = new Set([0, 1]);
        applyScope = new Set([0, 1]);
        applyData = new Map([
            [0, { idx: 0, deployments: [{ name: 'new', properties: { raiPolicyName: 'Microsoft.Default' } }], policies: [{ name: 'custom' }] }],
            [1, { idx: 1, deployments: [{ name: 'old', properties: { raiPolicyName: 'custom' } }], policies: [{ name: 'custom' }] }]
        ]);
    `, context);
    selectedModels.push({ dataset: { resIndex: '0', depIndex: '0' } }, { dataset: { resIndex: '1', depIndex: '0' } });
    return { context, targets, modelList };
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
    targets.set(0, 'local:custom');
    targets.set(1, 'local:custom');
    const result = vm.runInContext('getApplyChanges()', context);
    assert.equal(result.changes.length, 1);
    assert.equal(result.changes[0].dep.name, 'new');
    assert.equal(result.unchanged, 1);
    assert.equal(result.missingTarget, 0);
});

test('missing targets block the plan and stale subscription scope is ignored', () => {
    const { context, targets } = createApp();
    targets.set(1, 'local:custom');
    vm.runInContext(`selectedSubscriptions.delete('sub-2')`, context);
    const result = vm.runInContext('getApplyChanges()', context);
    assert.equal(result.changes.length, 0);
    assert.equal(result.missingTarget, 1);
    assert.equal(result.unchanged, 0);
});

test('another selected resource can supply a copyable policy to a resource with no policy', () => {
    const { context, targets, modelList } = createApp();
    vm.runInContext(`
        applyData.get(0).policies = [];
        applyData.get(1).policies = [{
            name: 'custom',
            properties: { basePolicyName: 'Microsoft.Default', mode: 'Blocking',
                contentFilters: [{ name: 'Hate', source: 'Prompt', enabled: true, blocking: true, severityThreshold: 'High' }] }
        }];
    `, context);
    targets.set(0, 'copy:1:custom');
    const result = vm.runInContext('getApplyChanges()', context);
    assert.equal(result.changes.length, 1);
    assert.equal(result.changes[0].target.copied, true);
    assert.equal(result.changes[0].target.name, 'custom');
    const body = vm.runInContext('copyApplyPolicyBody(applyData.get(1).policies[0])', context);
    assert.equal(body.properties.contentFilters[0].severityThreshold, 'High');
    assert.equal(body.properties.mode, 'Blocking');
    vm.runInContext('renderApplyModels()', context);
    assert.match(modelList.innerHTML, /value="copy:1:custom"/);
    assert.match(modelList.innerHTML, /复制自/);
    assert.match(modelList.innerHTML, /data-res-index="0" ><option/);
});

test('copy is rejected if source is out of scope or has unsupported dependencies', () => {
    const { context, targets } = createApp();
    vm.runInContext(`
        applyData.get(0).policies = [];
        applyData.get(1).policies = [{ name: 'custom', properties: {
            contentFilters: [{ name: 'Hate', source: 'Prompt', enabled: true, blocking: true }],
            customBlocklists: [{ blocklistName: 'local-only' }]
        }}];
    `, context);
    targets.set(0, 'copy:1:custom');
    assert.equal(vm.runInContext('getApplyChanges().missingTarget', context), 2);
    assert.equal(vm.runInContext('getApplyChanges().changes.length', context), 0);
    vm.runInContext('applyData.get(1).policies[0].properties.customBlocklists = []; applyScope.delete(1)', context);
    assert.equal(vm.runInContext('getApplyChanges().missingTarget', context), 1);
});
