// Main Application Logic
let selectedSubscriptions = new Set();
let allResources = [];
let selectedResources = new Set();
let currentWizStep = 1;
const MAX_WIZ_STEPS = 5;
let activityLog = [];
let subscriptionMap = {};
let resSortState = { key: null, dir: 1 };
let allSubscriptions = [];
let subSortState = { key: null, dir: 1 };
let wizardExecutionCompleted = false;
let wizardExecuting = false;
let wizardCancelRequested = false;
let resViewList = [];
let resPageState = { page: 1, size: 50 };
let filtersViewData = [];
let filtersPageState = { page: 1, size: 10 };
const REQUEST_CONCURRENCY = 8;
let applyScope = new Set();
let applyData = new Map();
let applyLoadVersion = 0;
let applyLoading = false;
let applyExecuting = false;

// Run async tasks with a bounded number of parallel requests
async function mapWithConcurrency(items, limit, task) {
    const results = new Array(items.length);
    let cursor = 0;
    const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
        while (cursor < items.length) {
            const current = cursor++;
            results[current] = await task(items[current], current);
        }
    });
    await Promise.all(workers);
    return results;
}

document.addEventListener('DOMContentLoaded', () => {
    initMsal();
    setupEventListeners();
    initFilterTables();
    if (isLoggedIn()) onLoginSuccess();
});

// ============ Event Listeners ============
function setupEventListeners() {
    // Login
    document.getElementById('btn-login').addEventListener('click', handleLogin);
    document.getElementById('btn-login-main').addEventListener('click', handleLogin);
    document.getElementById('btn-logout')?.addEventListener('click', handleLogout);
    
    // Navigation
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            navigateTo(item.dataset.page);
        });
    });
    document.querySelectorAll('[data-goto]').forEach(el => {
        el.addEventListener('click', () => navigateTo(el.dataset.goto));
    });
    document.getElementById('menu-toggle').addEventListener('click', () => {
        document.getElementById('sidebar').classList.toggle('collapsed');
    });
    
    // Subscriptions
    document.getElementById('btn-refresh-subs').addEventListener('click', loadSubscriptions);
    document.getElementById('btn-select-all-subs').addEventListener('click', () => toggleAllSubs(true));
    document.getElementById('btn-deselect-all-subs').addEventListener('click', () => toggleAllSubs(false));
    document.getElementById('subs-check-all').addEventListener('change', (e) => toggleAllSubs(e.target.checked));
    document.getElementById('sub-search').addEventListener('input', filterSubsList);
    document.querySelectorAll('#subs-table thead th.sortable').forEach(th => {
        th.addEventListener('click', () => sortSubscriptions(th.dataset.sort));
    });
    
    // Resources
    document.getElementById('btn-load-resources').addEventListener('click', loadResources);
    document.getElementById('btn-res-to-apply').addEventListener('click', () => navigateTo('batch-apply'));
    document.getElementById('btn-sel-all-res').addEventListener('click', () => toggleAllRes(true));
    document.getElementById('btn-desel-all-res').addEventListener('click', () => toggleAllRes(false));
    document.getElementById('res-check-all').addEventListener('change', (e) => toggleAllRes(e.target.checked));
    document.getElementById('res-search').addEventListener('input', filterResourcesList);
    document.getElementById('res-type-filter').addEventListener('change', filterResourcesList);
    document.querySelectorAll('#res-table thead th.sortable').forEach(th => {
        th.addEventListener('click', () => sortResources(th.dataset.sort));
    });
    document.getElementById('res-page-size').addEventListener('change', (e) => {
        resPageState.size = parseInt(e.target.value);
        resPageState.page = 1;
        applyResourceView();
    });
    document.getElementById('res-page-first').addEventListener('click', () => gotoResPage(1));
    document.getElementById('res-page-prev').addEventListener('click', () => gotoResPage(resPageState.page - 1));
    document.getElementById('res-page-next').addEventListener('click', () => gotoResPage(resPageState.page + 1));
    document.getElementById('res-page-last').addEventListener('click', () => gotoResPage(resTotalPages()));
    
    // Filters View
    document.getElementById('btn-load-filters').addEventListener('click', loadExistingFilters);
    document.getElementById('filters-show-system').addEventListener('change', () => {
        filtersPageState.page = 1;
        renderFiltersView();
    });
    document.getElementById('filters-page-size').addEventListener('change', (e) => {
        filtersPageState.size = parseInt(e.target.value);
        filtersPageState.page = 1;
        renderFiltersView();
    });
    document.getElementById('filters-page-first').addEventListener('click', () => gotoFiltersPage(1));
    document.getElementById('filters-page-prev').addEventListener('click', () => gotoFiltersPage(filtersPageState.page - 1));
    document.getElementById('filters-page-next').addEventListener('click', () => gotoFiltersPage(filtersPageState.page + 1));
    document.getElementById('filters-page-last').addEventListener('click', () => gotoFiltersPage(filtersTotalPages()));
    
    // Wizard
    document.getElementById('wiz-prev').addEventListener('click', wizPrev);
    document.getElementById('wiz-next').addEventListener('click', wizNext);
    document.getElementById('wiz-exec').addEventListener('click', () => executeWizard());
    document.getElementById('wiz-cancel').addEventListener('click', requestWizardCancel);
    
    // Presets
    document.querySelectorAll('.preset-btn').forEach(btn => {
        btn.addEventListener('click', () => applyPreset(btn.dataset.preset));
    });
    
    // Batch Apply
    document.getElementById('btn-apply-edit-resources').addEventListener('click', () => navigateTo('resources'));
    document.getElementById('btn-apply-load-models').addEventListener('click', loadApplyModels);
    document.getElementById('btn-apply-exec').addEventListener('click', executeBatchApply);
    document.getElementById('apply-model-search').addEventListener('input', filterApplyModels);
    document.getElementById('apply-only-default').addEventListener('change', filterApplyModels);
    document.getElementById('btn-apply-select-visible').addEventListener('click', () => toggleVisibleApplyModels(true));
    document.getElementById('btn-apply-clear-models').addEventListener('click', () => toggleVisibleApplyModels(false, true));
    
    // Batch Delete
    document.getElementById('btn-delete-load-filters').addEventListener('click', loadDeleteFilters);
    document.getElementById('btn-delete-exec').addEventListener('click', executeBatchDelete);
    
    // Activity
    document.getElementById('btn-clear-activity').addEventListener('click', () => {
        activityLog = [];
        document.getElementById('activity-log').innerHTML = '<p class="empty-hint">暂无操作记录</p>';
    });
    
    // Global search
    document.getElementById('global-search').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            const q = e.target.value.trim().toLowerCase();
            if (q) {
                // Simple routing based on keywords
                if (q.includes('订阅') || q.includes('sub')) navigateTo('subscriptions');
                else if (q.includes('资源') || q.includes('resource')) navigateTo('resources');
                else if (q.includes('创建') || q.includes('create')) navigateTo('batch-create');
                else if (q.includes('筛选') || q.includes('filter')) navigateTo('filters-view');
                else navigateTo('resources');
            }
        }
    });
}

// ============ Auth ============
async function handleLogin() {
    try {
        await login();
        onLoginSuccess();
    } catch (err) { alert('登录失败: ' + err.message); }
}

function handleLogout() {
    logout();
    location.reload();
}

function onLoginSuccess() {
    const user = getCurrentUser();
    document.getElementById('login-overlay').classList.add('hidden');
    document.getElementById('user-section').classList.remove('hidden');
    document.getElementById('btn-login').classList.add('hidden');
    document.getElementById('user-initial').textContent = (user.username || user.name || 'U')[0].toUpperCase();
    loadSubscriptions();
}

// ============ Navigation ============
function navigateTo(page) {
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    const pageEl = document.getElementById('page-' + page);
    if (pageEl) pageEl.classList.add('active');
    const navEl = document.querySelector(`[data-page="${page}"]`);
    if (navEl) navEl.classList.add('active');
    if (page === 'batch-apply') renderApplyResources();
}

// ============ Subscriptions ============
async function loadSubscriptions() {
    const loading = document.getElementById('subs-loading');
    const tbody = document.getElementById('subs-tbody');
    loading.classList.remove('hidden');
    tbody.innerHTML = '';
    
    try {
        allSubscriptions = await listSubscriptions();
        loading.classList.add('hidden');
        allSubscriptions.forEach(sub => {
            subscriptionMap[sub.subscriptionId] = sub.displayName;
        });
        subSortState = { key: null, dir: 1 };
        renderSubscriptions(allSubscriptions);
        updateSubSummary();
        updateSubSortIcons();
        updateStats();
    } catch (err) {
        loading.classList.add('hidden');
        tbody.innerHTML = `<tr><td colspan="4" style="color:var(--danger);padding:16px;">加载失败: ${err.message}</td></tr>`;
        document.getElementById('subs-summary').classList.add('hidden');
    }
}

function renderSubscriptions(subscriptions) {
    const tbody = document.getElementById('subs-tbody');
    tbody.innerHTML = '';
    if (subscriptions.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" class="empty-hint">未找到订阅</td></tr>';
        return;
    }
    subscriptions.forEach(sub => {
            const tr = document.createElement('tr');
            tr.dataset.id = sub.subscriptionId;
            tr.innerHTML = `
                <td class="w40"><input type="checkbox" class="sub-cb" data-id="${sub.subscriptionId}" ${selectedSubscriptions.has(sub.subscriptionId) ? 'checked' : ''}></td>
                <td><strong>${sub.displayName}</strong></td>
                <td style="font-family:monospace;font-size:12px;color:var(--text-muted)">${sub.subscriptionId}</td>
                <td><span class="sub-status">${sub.state || 'Enabled'}</span></td>
            `;
            tr.querySelector('.sub-cb').addEventListener('change', (e) => {
                if (e.target.checked) selectedSubscriptions.add(sub.subscriptionId);
                else selectedSubscriptions.delete(sub.subscriptionId);
                onSubscriptionScopeChanged();
                updateSubSummary();
                updateStats();
            });
            tbody.appendChild(tr);
    });
}

function subSortValue(sub, key) {
    switch (key) {
        case 'name': return (sub.displayName || '').toLowerCase();
        case 'id': return (sub.subscriptionId || '').toLowerCase();
        case 'state': return (sub.state || 'Enabled').toLowerCase();
        default: return '';
    }
}

function sortSubscriptions(key) {
    if (!key) return;
    if (subSortState.key === key) subSortState.dir *= -1;
    else subSortState = { key, dir: 1 };
    const sorted = [...allSubscriptions].sort((a, b) => {
        const valueA = subSortValue(a, key);
        const valueB = subSortValue(b, key);
        if (valueA < valueB) return -1 * subSortState.dir;
        if (valueA > valueB) return 1 * subSortState.dir;
        return 0;
    });
    renderSubscriptions(sorted);
    updateSubSortIcons();
    filterSubsList();
}

function updateSubSortIcons() {
    document.querySelectorAll('#subs-table thead th.sortable').forEach(th => {
        const icon = th.querySelector('.sort-icon');
        const active = th.dataset.sort === subSortState.key;
        th.classList.toggle('sorted', active);
        if (!icon) return;
        icon.className = 'fas sort-icon ' + (active ? (subSortState.dir === 1 ? 'fa-sort-up' : 'fa-sort-down') : 'fa-sort');
    });
}

function updateSubSummary() {
    const summary = document.getElementById('subs-summary');
    const visibleCount = [...document.querySelectorAll('#subs-tbody tr')]
        .filter(tr => tr.dataset.id && tr.style.display !== 'none').length;
    summary.innerHTML = `共 <strong>${allSubscriptions.length}</strong> 个订阅 · 已选择 <strong>${selectedSubscriptions.size}</strong> 个${visibleCount !== allSubscriptions.length ? ` · 当前显示 <strong>${visibleCount}</strong> 个` : ''}`;
    summary.classList.toggle('hidden', allSubscriptions.length === 0);
}

function toggleAllSubs(checked) {
    document.querySelectorAll('.sub-cb').forEach(cb => {
        cb.checked = checked;
        if (checked) selectedSubscriptions.add(cb.dataset.id);
        else selectedSubscriptions.delete(cb.dataset.id);
    });
    document.getElementById('subs-check-all').checked = checked;
    onSubscriptionScopeChanged();
    updateSubSummary();
    updateStats();
}

function filterSubsList() {
    const q = document.getElementById('sub-search').value.toLowerCase();
    document.querySelectorAll('#subs-tbody tr').forEach(tr => {
        const text = tr.textContent.toLowerCase();
        tr.style.display = text.includes(q) ? '' : 'none';
    });
    updateSubSummary();
}

// ============ Resources ============
async function loadResources() {
    if (selectedSubscriptions.size === 0) {
        alert('请先在"订阅管理"中选择至少一个订阅');
        navigateTo('subscriptions');
        return;
    }
    const loading = document.getElementById('res-loading');
    const tbody = document.getElementById('res-tbody');
    const warning = document.getElementById('res-load-warning');
    warning.classList.add('hidden');
    loading.classList.remove('hidden');
    tbody.innerHTML = '';
    allResources = [];
    
    try {
        const subIds = Array.from(selectedSubscriptions);
        const failedSubscriptions = [];
        const perSubscription = await mapWithConcurrency(subIds, REQUEST_CONCURRENCY, async (subId) => {
            try {
                const resources = await listCognitiveServicesAccounts(subId);
                resources.forEach(r => { r._subId = subId; });
                return resources;
            } catch (e) {
                failedSubscriptions.push(subscriptionMap[subId] || subId);
                return [];
            }
        });
        perSubscription.forEach(list => allResources.push(...list));
        // A bulk operation must never silently target every newly loaded resource.
        selectedResources.clear();
        allResources.forEach((r, i) => { r._idx = i; });
        resetApplyScope();
        resSortState = { key: null, dir: 1 };
        resPageState.page = 1;
        loading.classList.add('hidden');
        applyResourceView();
        if (failedSubscriptions.length) {
            warning.textContent = `${failedSubscriptions.length} 个订阅的资源加载失败：${failedSubscriptions.join('、')}。这些订阅不会进入本次操作范围，请检查权限或稍后重试。`;
            warning.classList.remove('hidden');
        }
        updateSortIcons();
        updateStats();
    } catch (err) {
        loading.classList.add('hidden');
        tbody.innerHTML = `<tr><td colspan="6" style="color:var(--danger);padding:16px;">加载失败: ${err.message}</td></tr>`;
    }
}

// Get sortable value for a resource by column key
function resSortValue(r, key) {
    const p = parseResourceId(r.id);
    switch (key) {
        case 'name': return (r.name || '').toLowerCase();
        case 'kind': return (r.kind || '').toLowerCase();
        case 'resourceGroup': return (p.resourceGroup || '').toLowerCase();
        case 'location': return (r.location || '').toLowerCase();
        case 'subscription': return (subscriptionMap[r._subId] || r._subId || '').toLowerCase();
        default: return '';
    }
}

function sortResources(key) {
    if (!key) return;
    if (resSortState.key === key) resSortState.dir *= -1;
    else { resSortState.key = key; resSortState.dir = 1; }
    resPageState.page = 1;
    applyResourceView();
    updateSortIcons();
}

function resTotalPages() {
    return Math.max(1, Math.ceil(resViewList.length / resPageState.size));
}

function gotoResPage(page) {
    const target = Math.min(Math.max(1, page), resTotalPages());
    if (target === resPageState.page) return;
    resPageState.page = target;
    applyResourceView();
}

// Recompute the filtered + sorted view, then render only the current page
function applyResourceView() {
    const query = document.getElementById('res-search').value.toLowerCase();
    const typeFilter = document.getElementById('res-type-filter').value;
    resViewList = allResources.filter(r => {
        if (!selectedSubscriptions.has(r._subId)) return false;
        const p = parseResourceId(r.id);
        const subName = subscriptionMap[r._subId] || r._subId || '';
        const haystack = `${r.name} ${r.kind || ''} ${p.resourceGroup} ${r.location} ${subName}`.toLowerCase();
        const matchText = !query || haystack.includes(query);
        const matchType = !typeFilter || (r.kind || '').includes(typeFilter);
        return matchText && matchType;
    });
    if (resSortState.key) {
        resViewList.sort((a, b) => {
            const va = resSortValue(a, resSortState.key), vb = resSortValue(b, resSortState.key);
            if (va < vb) return -1 * resSortState.dir;
            if (va > vb) return 1 * resSortState.dir;
            return 0;
        });
    }
    resPageState.page = Math.min(resPageState.page, resTotalPages());
    const start = (resPageState.page - 1) * resPageState.size;
    renderResources(resViewList.slice(start, start + resPageState.size));
    updateResSummary();
    updateResPager();
}

function updateResPager() {
    const pager = document.getElementById('res-pager');
    pager.classList.toggle('hidden', allResources.length === 0);
    document.getElementById('res-page-info').textContent = `第 ${resPageState.page} / ${resTotalPages()} 页`;
    const atFirst = resPageState.page <= 1;
    const atLast = resPageState.page >= resTotalPages();
    document.getElementById('res-page-first').disabled = atFirst;
    document.getElementById('res-page-prev').disabled = atFirst;
    document.getElementById('res-page-next').disabled = atLast;
    document.getElementById('res-page-last').disabled = atLast;
}

function updateResSummary() {
    const el = document.getElementById('res-summary');
    if (!el) return;
    const subCount = new Set(allResources.map(r => r._subId)).size;
    const filteredNote = resViewList.length !== allResources.length ? ` · 当前筛选 <strong>${resViewList.length}</strong> 个` : '';
    el.innerHTML = `已加载 <strong>${allResources.length}</strong> 个资源 · 来自 <strong>${subCount}</strong> 个订阅${filteredNote} · 已选择 <strong>${getSelectedResourceIndices().length}</strong> 个（仅已选订阅可操作）`;
    el.classList.toggle('hidden', allResources.length === 0);
}

function renderResources(resources) {
    const tbody = document.getElementById('res-tbody');
    tbody.innerHTML = '';
    if (resources.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-hint">未找到 OpenAI/AI Services 资源</td></tr>';
        return;
    }
    resources.forEach((r) => {
        const i = r._idx;
        const p = parseResourceId(r.id);
        const subName = subscriptionMap[r._subId] || `${r._subId.substring(0, 8)}...`;
        const isChecked = selectedResources.has(i);
        const tr = document.createElement('tr');
        tr.dataset.index = i;
        tr.innerHTML = `
            <td class="w40"><input type="checkbox" class="res-cb" data-index="${i}" ${isChecked ? 'checked' : ''}></td>
            <td><strong>${r.name}</strong></td>
            <td><span class="media-tag">${r.kind || '-'}</span></td>
            <td>${p.resourceGroup}</td>
            <td>${r.location}</td>
            <td>${subName}</td>
        `;
        const cb = tr.querySelector('.res-cb');
        cb.addEventListener('change', (e) => {
            if (e.target.checked) selectedResources.add(i);
            else selectedResources.delete(i);
            resetApplyScope();
            updateResSummary();
            updateStats();
        });
        tbody.appendChild(tr);
    });
}

// Applies to every resource in the current filtered view, not just the visible page
function toggleAllRes(checked) {
    resViewList.forEach(r => {
        if (checked) selectedResources.add(r._idx);
        else selectedResources.delete(r._idx);
    });
    document.querySelectorAll('.res-cb').forEach(cb => { cb.checked = checked; });
    document.getElementById('res-check-all').checked = checked;
    resetApplyScope();
    updateResSummary();
    updateStats();
}

function filterResourcesList() {
    resPageState.page = 1;
    applyResourceView();
}

function updateSortIcons() {
    document.querySelectorAll('#res-table thead th.sortable').forEach(th => {
        const icon = th.querySelector('.sort-icon');
        const active = th.dataset.sort === resSortState.key;
        th.classList.toggle('sorted', active);
        if (!icon) return;
        if (active) icon.className = 'fas sort-icon ' + (resSortState.dir === 1 ? 'fa-sort-up' : 'fa-sort-down');
        else icon.className = 'fas fa-sort sort-icon';
    });
}

// ============ Existing Filters View ============
async function loadExistingFilters() {
    const indices = getSelectedResourceIndices();
    if (indices.length === 0) {
        alert('请先在"资源浏览"中加载并勾选资源');
        navigateTo('resources');
        return;
    }
    const container = document.getElementById('filters-view-content');
    const loading = document.getElementById('filters-loading');
    loading.classList.remove('hidden');
    container.innerHTML = '';
    filtersViewData = [];
    filtersPageState.page = 1;
    
    try {
        filtersViewData = await mapWithConcurrency(indices, REQUEST_CONCURRENCY, async (idx) => {
            const r = allResources[idx];
            const p = parseResourceId(r.id);
            try {
                const [policies, deployments] = await Promise.all([
                    listRaiPolicies(p.subscriptionId, p.resourceGroup, p.accountName),
                    listDeployments(p.subscriptionId, p.resourceGroup, p.accountName).catch(() => [])
                ]);
                const appliedModels = {};
                deployments.forEach(dep => {
                    const policyName = dep.properties?.raiPolicyName;
                    if (!policyName) return;
                    if (!appliedModels[policyName]) appliedModels[policyName] = [];
                    appliedModels[policyName].push(dep.name);
                });
                return {
                    resourceName: r.name,
                    location: r.location,
                    error: null,
                    policies: policies.map(pol => ({
                        name: pol.name,
                        isSystem: isSystemPolicy(pol),
                        appliedModels: appliedModels[pol.name] || []
                    }))
                };
            } catch (e) {
                return { resourceName: r.name, location: r.location, error: e.message, policies: [] };
            }
        });
        loading.classList.add('hidden');
        renderFiltersView();
    } catch (err) {
        loading.classList.add('hidden');
        container.innerHTML = `<p style="color:var(--danger)">加载失败: ${err.message}</p>`;
    }
}

// Azure marks built-in policies as SystemManaged; older API responses only expose the Microsoft.* name
function isSystemPolicy(policy) {
    if (policy.properties?.type === 'SystemManaged') return true;
    return /^microsoft\./i.test(policy.name || '');
}

function getFiltersViewGroups() {
    const showSystem = document.getElementById('filters-show-system').checked;
    return filtersViewData
        .map(group => ({
            ...group,
            visiblePolicies: showSystem ? group.policies : group.policies.filter(pol => !pol.isSystem)
        }))
        .filter(group => group.error || group.visiblePolicies.length > 0);
}

function filtersTotalPages() {
    return Math.max(1, Math.ceil(getFiltersViewGroups().length / filtersPageState.size));
}

function gotoFiltersPage(page) {
    const target = Math.min(Math.max(1, page), filtersTotalPages());
    if (target === filtersPageState.page) return;
    filtersPageState.page = target;
    renderFiltersView();
}

function renderFiltersView() {
    const container = document.getElementById('filters-view-content');
    const summary = document.getElementById('filters-summary');
    const pager = document.getElementById('filters-pager');
    if (filtersViewData.length === 0) {
        summary.classList.add('hidden');
        pager.classList.add('hidden');
        return;
    }
    
    const groups = getFiltersViewGroups();
    const customCount = filtersViewData.reduce((sum, g) => sum + g.policies.filter(p => !p.isSystem).length, 0);
    const systemCount = filtersViewData.reduce((sum, g) => sum + g.policies.filter(p => p.isSystem).length, 0);
    const shownCount = groups.reduce((sum, g) => sum + (g.visiblePolicies?.length || 0), 0);
    document.getElementById('stat-filters').textContent = customCount;
    
    summary.innerHTML = `共 <strong>${filtersViewData.length}</strong> 个资源 · 自定义 <strong>${customCount}</strong> 个 · 系统 <strong>${systemCount}</strong> 个 · 当前显示 <strong>${shownCount}</strong> 个筛选器`;
    summary.classList.remove('hidden');
    
    if (groups.length === 0) {
        container.innerHTML = '<p class="empty-hint">已选资源中未找到符合条件的筛选器</p>';
        pager.classList.add('hidden');
        return;
    }
    
    filtersPageState.page = Math.min(filtersPageState.page, filtersTotalPages());
    const start = (filtersPageState.page - 1) * filtersPageState.size;
    const pageGroups = groups.slice(start, start + filtersPageState.size);
    
    container.innerHTML = pageGroups.map(group => {
        if (group.error) {
            return `<div class="filter-group"><div class="filter-group-header"><span>${group.resourceName}</span><span style="color:var(--danger)">加载失败: ${group.error.substring(0, 120)}</span></div></div>`;
        }
        const items = group.visiblePolicies.map(pol => {
            const applied = pol.appliedModels.length > 0
                ? `<span class="applied-models">${pol.appliedModels.map(m => `<span class="applied-model-tag">${m}</span>`).join('')}</span>`
                : '<span class="applied-none">未应用到模型</span>';
            return `<div class="filter-item"><span class="fmeta"><span>${pol.name}</span><span class="filter-badge ${pol.isSystem ? 'system' : 'custom'}">${pol.isSystem ? '系统' : '自定义'}</span></span>${applied}</div>`;
        }).join('');
        return `<div class="filter-group"><div class="filter-group-header"><span>${group.resourceName} (${group.location})</span><span>${group.visiblePolicies.length} 个筛选器</span></div>${items}</div>`;
    }).join('');
    
    pager.classList.remove('hidden');
    document.getElementById('filters-page-info').textContent = `第 ${filtersPageState.page} / ${filtersTotalPages()} 页`;
    const atFirst = filtersPageState.page <= 1;
    const atLast = filtersPageState.page >= filtersTotalPages();
    document.getElementById('filters-page-first').disabled = atFirst;
    document.getElementById('filters-page-prev').disabled = atFirst;
    document.getElementById('filters-page-next').disabled = atLast;
    document.getElementById('filters-page-last').disabled = atLast;
}

// ============ Filter Table Init ============
function initFilterTables() {
    const categories = [
        { name: 'Hate', label: 'Hate (仇恨)', media: 'Text / Image' },
        { name: 'Sexual', label: 'Sexual (性内容)', media: 'Text / Image' },
        { name: 'Selfharm', label: 'Self-harm (自伤)', media: 'Text / Image' },
        { name: 'Violence', label: 'Violence (暴力)', media: 'Text / Image' }
    ];
    const others = [
        { name: 'Jailbreak', label: 'Jailbreak (越狱攻击)', source: 'Prompt' },
        { name: 'Protected Material Text', label: 'Protected Material (文本)', source: 'Completion' },
        { name: 'Protected Material Code', label: 'Protected Material (代码)', source: 'Completion' },
        { name: 'Profanity', label: 'Profanity (亵渎)', source: 'Prompt' }
    ];
    
    renderFilterTable('input-filters-body', categories);
    renderFilterTable('output-filters-body', categories);
    renderOtherFilterTable('other-filters-body', others);
}

// Slider value mapping: 1=Lowest blocking (most lenient, API High), 2=Medium, 3=Highest blocking (most strict, API Low)
// This matches the Azure portal slider direction: left=lenient, right=strict
const SLIDER_TO_API = { 1: 'High', 2: 'Medium', 3: 'Low' };
const SLIDER_LABELS = {
    1: { level: 'Lowest blocking', desc: 'Blocks only the most severe unwanted content' },
    2: { level: 'Medium blocking', desc: 'Blocks both moderate and highly severe unwanted content' },
    3: { level: 'Highest blocking', desc: 'Blocks all severity levels of unwanted content' }
};

function updateSliderLabel(tr, value) {
    const info = tr.querySelector('.severity-info');
    const data = SLIDER_LABELS[value];
    info.innerHTML = `<span class="severity-level">${data.level}</span><span class="severity-desc">${data.desc}</span>`;
}

function renderFilterTable(tbodyId, categories) {
    const tbody = document.getElementById(tbodyId);
    const source = tbodyId.includes('input') ? 'input' : 'output';
    categories.forEach(cat => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${cat.label}</strong></td>
            <td><div class="media-tags"><span class="media-tag">Text</span><span class="media-tag">Image</span></div></td>
            <td class="action-cell">
                <select class="filter-action" data-source="${source}" data-name="${cat.name}">
                    <option value="annotate_and_block">Annotate and block</option>
                    <option value="annotate_only">Annotate only *</option>
                    <option value="off">Off (关闭) *</option>
                </select>
            </td>
            <td class="severity-cell">
                <div class="severity-slider-wrap">
                    <input type="range" class="severity-slider" data-source="${source}" data-name="${cat.name}" min="1" max="3" value="2">
                    <div class="severity-info">
                        <span class="severity-level">Medium blocking</span>
                        <span class="severity-desc">Blocks both moderate and highly severe unwanted content</span>
                    </div>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
    
    // Events
    tbody.querySelectorAll('.filter-action').forEach(sel => {
        sel.addEventListener('change', (e) => {
            const row = e.target.closest('tr');
            const slider = row.querySelector('.severity-slider');
            slider.disabled = (e.target.value !== 'annotate_and_block');
            updateFallbackVisibility();
        });
    });
    tbody.querySelectorAll('.severity-slider').forEach(slider => {
        slider.addEventListener('input', (e) => {
            updateSliderLabel(e.target.closest('tr'), e.target.value);
        });
    });
}

function renderOtherFilterTable(tbodyId, others) {
    const tbody = document.getElementById(tbodyId);
    others.forEach(f => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${f.label}</strong></td>
            <td>${f.source === 'Prompt' ? '输入 (Prompt)' : '输出 (Completion)'}</td>
            <td class="action-cell">
                <select class="other-filter-action" data-name="${f.name}" data-source="${f.source}">
                    <option value="annotate_and_block">Annotate and block</option>
                    <option value="annotate_only">Annotate only *</option>
                    <option value="off">Off (关闭) *</option>
                </select>
            </td>
        `;
        tbody.appendChild(tr);
    });
    // Events for other filters
    tbody.querySelectorAll('.other-filter-action').forEach(sel => {
        sel.addEventListener('change', () => updateFallbackVisibility());
    });
}

// Show/hide fallback section based on whether any filter uses off/annotate_only
function updateFallbackVisibility() {
    const fallbackSection = document.getElementById('fallback-section');
    fallbackSection.classList.toggle('hidden', !configNeedsApproval());
}

// ============ Presets ============
function applyPreset(preset) {
    document.querySelectorAll('.preset-btn').forEach(b => b.classList.remove('active'));
    const active = document.querySelector(`[data-preset="${preset}"]`);
    if (active) active.classList.add('active');
    
    const actions = document.querySelectorAll('.filter-action, .other-filter-action');
    const sliders = document.querySelectorAll('.severity-slider');
    
    // Show/hide fallback section for permission-requiring presets
    const fallbackSection = document.getElementById('fallback-section');
    const needsApproval = (preset === 'off' || preset === 'annotate');
    fallbackSection.classList.toggle('hidden', !needsApproval);
    
    switch (preset) {
        case 'low': // Lowest blocking = most lenient = slider 1
            actions.forEach(s => s.value = 'annotate_and_block');
            sliders.forEach(s => { s.value = 1; s.disabled = false; updateSliderLabel(s.closest('tr'), 1); });
            break;
        case 'medium':
            actions.forEach(s => s.value = 'annotate_and_block');
            sliders.forEach(s => { s.value = 2; s.disabled = false; updateSliderLabel(s.closest('tr'), 2); });
            break;
        case 'high': // Highest blocking = most strict = slider 3
            actions.forEach(s => s.value = 'annotate_and_block');
            sliders.forEach(s => { s.value = 3; s.disabled = false; updateSliderLabel(s.closest('tr'), 3); });
            break;
        case 'annotate':
            actions.forEach(s => s.value = 'annotate_only');
            sliders.forEach(s => s.disabled = true);
            break;
        case 'off':
            actions.forEach(s => s.value = 'off');
            sliders.forEach(s => s.disabled = true);
            break;
    }
}

// Check if current config requires modified content filter approval
function configNeedsApproval() {
    const actions = document.querySelectorAll('.filter-action, .other-filter-action');
    for (const sel of actions) {
        if (sel.value === 'off' || sel.value === 'annotate_only') return true;
    }
    return false;
}

// Build fallback policy body (for resources without approval)
function buildFallbackPolicyBody(mode) {
    const fallbackAction = document.getElementById('fallback-action').value;
    if (fallbackAction === 'skip') return null;
    
    const severityMap = {
        'annotate_and_block_high': 'High',
        'annotate_and_block_medium': 'Medium',
        'annotate_and_block_low': 'Low'
    };
    const severity = severityMap[fallbackAction] || 'High';
    
    const categories = ['Hate', 'Sexual', 'Selfharm', 'Violence'];
    const contentFilters = [];
    
    categories.forEach(cat => {
        contentFilters.push({ name: cat, enabled: true, blocking: true, severityThreshold: severity, source: 'Prompt' });
        contentFilters.push({ name: cat, enabled: true, blocking: true, severityThreshold: severity, source: 'Completion' });
    });
    contentFilters.push({ name: 'Jailbreak', enabled: true, blocking: true, source: 'Prompt' });
    contentFilters.push({ name: 'Protected Material Text', enabled: true, blocking: true, source: 'Completion' });
    contentFilters.push({ name: 'Protected Material Code', enabled: true, blocking: true, source: 'Completion' });
    contentFilters.push({ name: 'Profanity', enabled: true, blocking: true, source: 'Prompt' });
    
    return {
        properties: {
            basePolicyName: 'Microsoft.Default',
            mode: mode || 'Asynchronous_filter',
            contentFilters
        }
    };
}

// Check if an API error is a permission/approval error for modified content filters
function isApprovalError(errorMsg) {
    const lower = errorMsg.toLowerCase();
    // Match errors specifically related to content filter approval/policy restrictions
    const approvalPatterns = [
        'requestdisallowedbypolicy',
        'contentfilterblocklistnotavailable',
        'invalidcontentfilterpolicy',
        'not approved for modified',
        'content filtering configuration is not allowed',
        'disablecontentfilter',
        'content filter policy',
        'not authorized to disable',
        'modified content filter',
        'permission to override base policy',
        'override base policy',
        'oai/rai/exceptions',
        'does not have necessary permission'
    ];
    // Also check for 403 Forbidden which indicates permission denied
    if (lower.includes('403') && (lower.includes('forbidden') || lower.includes('not authorized'))) return true;
    return approvalPatterns.some(p => lower.includes(p.toLowerCase()));
}

// Validate Azure resource name format
function isValidResourceName(name) {
    return /^[a-zA-Z0-9][a-zA-Z0-9_.\-]*$/.test(name);
}

// ============ Wizard ============
function wizPrev() {
    if (currentWizStep > 1) {
        currentWizStep--;
        if (currentWizStep < MAX_WIZ_STEPS) resetWizardExecutionState();
        updateWizardUI();
    }
}

function wizNext() {
    if (currentWizStep === 1) {
        const name = document.getElementById('filter-name').value.trim();
        if (!name) { alert('请输入筛选器名称'); return; }
        if (!isValidResourceName(name)) {
            alert('筛选器名称格式无效\n\n只允许字母、数字、下划线(_)、点(.)、短横线(-)，且必须以字母或数字开头\n\n例如: custom-filter-off');
            return;
        }
        // Also validate fallback name if visible
        const fbSection = document.getElementById('fallback-section');
        if (!fbSection.classList.contains('hidden')) {
            const fbName = document.getElementById('fallback-filter-name').value.trim();
            if (fbName && !isValidResourceName(fbName)) {
                alert('降级筛选器名称格式无效\n\n只允许字母、数字、下划线(_)、点(.)、短横线(-)，且必须以字母或数字开头');
                return;
            }
        }
    }
    if (currentWizStep < MAX_WIZ_STEPS) {
        currentWizStep++;
        if (currentWizStep === MAX_WIZ_STEPS) renderSummary();
        updateWizardUI();
    }
}

function updateWizardUI() {
    document.querySelectorAll('.wiz-page').forEach(p => p.classList.remove('active'));
    document.getElementById('wiz-' + currentWizStep).classList.add('active');
    
    document.querySelectorAll('.wstep').forEach(s => {
        const step = parseInt(s.dataset.step);
        s.classList.remove('active', 'done');
        if (step === currentWizStep) s.classList.add('active');
        else if (step < currentWizStep) s.classList.add('done');
    });
    
    document.getElementById('wiz-prev').disabled = (currentWizStep === 1);
    document.getElementById('wiz-next').classList.toggle('hidden', currentWizStep === MAX_WIZ_STEPS);
    const executeButton = document.getElementById('wiz-exec');
    executeButton.classList.toggle('hidden', currentWizStep !== MAX_WIZ_STEPS);
    executeButton.disabled = wizardExecutionCompleted || wizardExecuting;
    executeButton.innerHTML = wizardExecutionCompleted
        ? '<i class="fas fa-check"></i> 已完成执行'
        : '<i class="fas fa-rocket"></i> 开始执行';
    const cancelButton = document.getElementById('wiz-cancel');
    cancelButton.classList.toggle('hidden', !wizardExecuting);
}

function requestWizardCancel() {
    if (!wizardExecuting) return;
    wizardCancelRequested = true;
    const cancelButton = document.getElementById('wiz-cancel');
    cancelButton.disabled = true;
    cancelButton.innerHTML = '<i class="fas fa-hourglass-half"></i> 取消中...';
    log(document.getElementById('exec-log'), '已请求取消，当前资源处理完后停止', 'w');
}

function resetWizardExecutionState() {
    wizardExecutionCompleted = false;
    const result = document.getElementById('exec-result');
    result.classList.add('hidden');
    result.innerHTML = '';
}

function renderSummary() {
    const name = document.getElementById('filter-name').value;
    const mode = document.getElementById('filter-mode').value;
    const resourceCount = getSelectedResourceIndices().length;
    const applyToModels = document.getElementById('apply-to-models').checked;
    const needsApproval = configNeedsApproval();
    const fallbackEnabled = needsApproval && document.getElementById('fallback-enabled').checked;
    const fbAction = document.getElementById('fallback-action').value;
    const fbName = document.getElementById('fallback-filter-name').value.trim() || name;
    
    let html = `
        <table class="data-table" style="margin-bottom:16px">
            <tr><td style="font-weight:600;width:160px">筛选器名称</td><td>${name}</td></tr>
            <tr><td style="font-weight:600">筛选模式</td><td>${mode}</td></tr>
            <tr><td style="font-weight:600">目标资源数</td><td>${resourceCount} 个资源</td></tr>
            <tr><td style="font-weight:600">自动应用到部署</td><td>${applyToModels ? '是' : '否'}</td></tr>
            <tr><td style="font-weight:600">需要审批权限</td><td>${needsApproval ? '<span style="color:var(--warning)">是（包含关闭/仅批注配置）</span>' : '否'}</td></tr>`;
    
    if (needsApproval) {
        html += `<tr><td style="font-weight:600">自动降级</td><td>${fallbackEnabled ? '已启用' : '<span style="color:var(--danger)">未启用（权限不足将报错）</span>'}</td></tr>`;
        if (fallbackEnabled && fbAction !== 'skip') {
            html += `<tr><td style="font-weight:600">降级配置</td><td>${fbAction.replace('annotate_and_block_', 'Annotate and block - ')} (名称: ${fbName})</td></tr>`;
        } else if (fallbackEnabled && fbAction === 'skip') {
            html += `<tr><td style="font-weight:600">降级配置</td><td>跳过不创建</td></tr>`;
        }
    }
    
    html += '</table>';
    document.getElementById('wiz-summary').innerHTML = html;
}

// ============ Execute Operations ============
async function executeWizard(retryIndices) {
    const isRetry = Array.isArray(retryIndices) && retryIndices.length > 0;
    if (wizardExecuting) return;
    if (wizardExecutionCompleted && !isRetry) return;
    const filterName = document.getElementById('filter-name').value.trim();
    const applyToModels = document.getElementById('apply-to-models').checked;
    const indices = isRetry ? retryIndices : getSelectedResourceIndices();
    
    if (indices.length === 0) { alert('没有选中的资源'); return; }
    
    const config = gatherFilterConfig();
    const policyBody = buildRaiPolicyBody(config);
    const mode = document.getElementById('filter-mode').value;
    
    // Fallback settings
    const needsApproval = configNeedsApproval();
    const fallbackEnabled = needsApproval && document.getElementById('fallback-enabled').checked;
    const fallbackBody = fallbackEnabled ? buildFallbackPolicyBody(mode) : null;
    const fallbackFilterName = (document.getElementById('fallback-filter-name').value.trim()) || filterName;
    
    const section = document.getElementById('exec-progress-section');
    section.classList.remove('hidden');
    const result = document.getElementById('exec-result');
    result.classList.add('hidden');
    result.innerHTML = '';
    const logEl = document.getElementById('exec-log');
    if (!isRetry) logEl.innerHTML = '';
    
    let done = 0, errors = 0, fallbackUsed = 0, deploymentsApplied = 0, deploymentErrors = 0;
    const failures = [];
    const subName = (parsed) => subscriptionMap[parsed.subscriptionId] || parsed.subscriptionId;
    function recordFailure(resource, parsed, stage, target, message) {
        failures.push({
            idx: resource._idx,
            resource: resource.name,
            subscription: subName(parsed),
            subscriptionId: parsed.subscriptionId,
            resourceGroup: parsed.resourceGroup,
            stage,
            target: target || '-',
            message: (message || '').toString().substring(0, 300)
        });
    }
    const total = indices.length;
    
    log(logEl, isRetry ? `── 重试开始 ──` : `开始批量创建筛选器: ${filterName}`, 'i');
    log(logEl, isRetry ? `重试失败资源: ${total} 个` : `目标资源: ${total} 个`, 'i');
    if (needsApproval && fallbackEnabled) {
        const fbAction = document.getElementById('fallback-action').value;
        if (fbAction === 'skip') {
            log(logEl, `降级策略: 跳过（权限不足时不创建）`, 'w');
        } else {
            log(logEl, `降级策略: 已启用 → ${fbAction.replace('annotate_and_block_', 'Annotate and block - ')} (名称: ${fallbackFilterName})`, 'w');
        }
    }
    log(logEl, '─'.repeat(50), 'i');
    
    wizardExecuting = true;
    wizardCancelRequested = false;
    const cancelButton = document.getElementById('wiz-cancel');
    cancelButton.disabled = false;
    cancelButton.innerHTML = '<i class="fas fa-stop"></i> 取消任务';
    cancelButton.classList.remove('hidden');
    document.getElementById('wiz-exec').disabled = true;
    
    let cancelled = false;
    for (const idx of indices) {
        if (wizardCancelRequested) {
            cancelled = true;
            log(logEl, '⏹ 任务已取消，剩余资源未处理', 'w');
            break;
        }
        const r = allResources[idx];
        const p = parseResourceId(r.id);
        let usedFallback = false;
        let actualFilterName = filterName;
        
        try {
            log(logEl, `[${r.name}] 创建筛选器 "${filterName}"...`, 'i');
            await createOrUpdateRaiPolicy(p.subscriptionId, p.resourceGroup, p.accountName, filterName, policyBody);
            log(logEl, `[${r.name}] ✓ 筛选器创建成功`, 's');
        } catch (err) {
            // Check if this is a permission error and fallback is available
            const errMsg = err.message || String(err);
            if (needsApproval && fallbackEnabled && isApprovalError(errMsg)) {
                log(logEl, `[${r.name}] ⚠ 权限不足，该资源不支持关闭/仅批注筛选器`, 'w');
                log(logEl, `[${r.name}] 原始错误: ${errMsg.substring(0, 200)}`, 'w');
                
                if (fallbackBody) {
                    try {
                        actualFilterName = fallbackFilterName;
                        log(logEl, `[${r.name}] ↻ 自动降级: 使用备选配置创建 "${actualFilterName}"...`, 'w');
                        await createOrUpdateRaiPolicy(p.subscriptionId, p.resourceGroup, p.accountName, actualFilterName, fallbackBody);
                        log(logEl, `[${r.name}] ✓ 降级筛选器创建成功`, 's');
                        usedFallback = true;
                        fallbackUsed++;
                    } catch (fbErr) {
                        errors++;
                        recordFailure(r, p, '创建降级筛选器', actualFilterName, fbErr.message);
                        log(logEl, `[${subName(p)} / ${r.name}] ✗ 降级也失败: ${fbErr.message}`, 'e');
                        updateProgress('exec-progress', 'exec-progress-text', done + errors + fallbackUsed, total);
                        continue;
                    }
                } else {
                    // fallback = skip
                    log(logEl, `[${r.name}] ⏭ 跳过（降级策略: 不创建）`, 'w');
                    fallbackUsed++;
                    updateProgress('exec-progress', 'exec-progress-text', done + errors + fallbackUsed, total);
                    continue;
                }
            } else {
                errors++;
                recordFailure(r, p, '创建筛选器', filterName, err.message || String(err));
                log(logEl, `[${subName(p)} / ${r.name}] ✗ 失败: ${(err.message || String(err)).substring(0, 300)}`, 'e');
                updateProgress('exec-progress', 'exec-progress-text', done + errors + fallbackUsed, total);
                continue;
            }
        }
        
        // Apply to deployments
        if (applyToModels) {
            try {
                const deployments = await listDeployments(p.subscriptionId, p.resourceGroup, p.accountName);
                for (const dep of deployments) {
                    try {
                        await updateDeploymentRaiPolicy(p.subscriptionId, p.resourceGroup, p.accountName, dep.name, actualFilterName, dep);
                        log(logEl, `[${r.name}] ✓ 已应用到: ${dep.name}${usedFallback ? ' (降级配置)' : ''}`, 's');
                        deploymentsApplied++;
                    } catch (e) {
                        recordFailure(r, p, '应用到模型', dep.name, e.message);
                        log(logEl, `[${subName(p)} / ${r.name}] ✗ 应用失败 ${dep.name}: ${e.message}`, 'e');
                        deploymentErrors++;
                    }
                }
            } catch (e) {
                recordFailure(r, p, '获取部署列表', '-', e.message);
                log(logEl, `[${subName(p)} / ${r.name}] ✗ 获取部署列表失败: ${e.message}`, 'e');
                deploymentErrors++;
            }
        }
        
        if (usedFallback) {
            // fallbackUsed already incremented
        } else {
            done++;
        }
        updateProgress('exec-progress', 'exec-progress-text', done + errors + fallbackUsed, total);
    }
    
    log(logEl, '─'.repeat(50), 'i');
    log(logEl, `${cancelled ? '已取消' : '完成'}! 成功: ${done}, 降级: ${fallbackUsed}, 失败: ${errors}, 共 ${total}`, 
        errors + deploymentErrors === 0 ? 's' : 'w');
    if (applyToModels) {
        log(logEl, `模型应用: 成功 ${deploymentsApplied}, 失败 ${deploymentErrors}`, deploymentErrors === 0 ? 's' : 'w');
    }
    if (fallbackUsed > 0) {
        log(logEl, `⚠ ${fallbackUsed} 个资源使用了降级配置（权限不足）`, 'w');
    }
    addActivity(`批量创建 "${filterName}"${cancelled ? '（已取消）' : ''} - 成功${done} 降级${fallbackUsed} 失败${errors}/${total}${applyToModels ? `，模型应用成功${deploymentsApplied}失败${deploymentErrors}` : ''}`);
    wizardExecuting = false;
    wizardCancelRequested = false;
    cancelButton.classList.add('hidden');
    wizardExecutionCompleted = !cancelled;
    const executeButton = document.getElementById('wiz-exec');
    executeButton.disabled = wizardExecutionCompleted;
    executeButton.innerHTML = wizardExecutionCompleted
        ? '<i class="fas fa-check"></i> 已完成执行'
        : '<i class="fas fa-rocket"></i> 开始执行';
    const hasErrors = errors + deploymentErrors > 0;
    const processed = done + fallbackUsed + errors;
    result.className = `execution-result ${cancelled || hasErrors ? 'warning' : 'success'}`;
    const headline = cancelled
        ? `任务已取消，已处理 ${processed}/${total} 个资源，可修改配置后重新执行。`
        : (isRetry ? '重试已完成。' : '任务已完成，已锁定重复执行。');
    const failedResourceCount = new Set(failures.map(f => f.idx)).size;
    const failureHtml = failures.length === 0 ? '' : `
        <details class="failure-details" open>
            <summary>失败明细（${failures.length} 条 · ${failedResourceCount} 个资源）</summary>
            <table class="data-table failure-table">
                <thead><tr><th>订阅</th><th>资源组</th><th>资源</th><th>阶段</th><th>对象</th><th>错误</th></tr></thead>
                <tbody>${failures.map(f => `<tr><td>${escapeHtml(f.subscription)}</td><td>${escapeHtml(f.resourceGroup)}</td><td>${escapeHtml(f.resource)}</td><td>${escapeHtml(f.stage)}</td><td>${escapeHtml(f.target)}</td><td class="failure-msg">${escapeHtml(f.message)}</td></tr>`).join('')}</tbody>
            </table>
            <div class="failure-actions">
                <button class="btn btn-sm btn-primary" id="retry-failures"><i class="fas fa-redo"></i> 重试失败的 ${failedResourceCount} 个资源</button>
                <button class="btn btn-sm btn-outline" id="copy-failures"><i class="fas fa-copy"></i> 复制失败明细</button>
            </div>
        </details>`;
    result.innerHTML = `<i class="fas ${cancelled || hasErrors ? 'fa-exclamation-triangle' : 'fa-check-circle'}"></i><div><strong>${headline}</strong><span>筛选器：成功 ${done} 个${fallbackUsed ? `，降级 ${fallbackUsed} 个` : ''}${errors ? `，失败 ${errors} 个` : ''}，共 ${total} 个资源。${applyToModels ? ` 模型应用：成功 ${deploymentsApplied} 个，失败 ${deploymentErrors} 个。` : ''}</span>${failureHtml}</div>`;
    result.classList.remove('hidden');
    const retryButton = document.getElementById('retry-failures');
    if (retryButton) {
        retryButton.addEventListener('click', () => {
            executeWizard([...new Set(failures.map(f => f.idx))]);
        });
    }
    const copyButton = document.getElementById('copy-failures');
    if (copyButton) {
        copyButton.addEventListener('click', () => {
            const text = failures.map(f => `${f.subscription} (${f.subscriptionId}) | ${f.resourceGroup} | ${f.resource} | ${f.stage} | ${f.target} | ${f.message}`).join('\n');
            navigator.clipboard.writeText(text).then(() => {
                copyButton.innerHTML = '<i class="fas fa-check"></i> 已复制';
            });
        });
    }
}

function escapeHtml(text) {
    return String(text).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function onSubscriptionScopeChanged() {
    selectedResources = new Set(getSelectedResourceIndices());
    if (allResources.length) applyResourceView();
    resetApplyScope();
}

function resetApplyScope() {
    applyScope = new Set(getSelectedResourceIndices());
    invalidateApplyData();
}

function groupApplyResources(indices) {
    const groups = new Map();
    indices.forEach(idx => {
        const subId = allResources[idx]._subId;
        if (!groups.has(subId)) groups.set(subId, []);
        groups.get(subId).push(idx);
    });
    return [...groups].sort(([a], [b]) => (subscriptionMap[a] || a).localeCompare(subscriptionMap[b] || b, 'zh-CN'));
}

function renderApplyResources() {
    const candidates = getSelectedResourceIndices();
    const groups = groupApplyResources(candidates);
    const container = document.getElementById('apply-resources-list');
    document.getElementById('apply-scope-summary').textContent = `${groups.length} 个订阅 · 资源浏览已选 ${candidates.length} 个 · 本次操作 ${applyScope.size} 个`;
    if (!candidates.length) {
        container.innerHTML = '<div class="apply-empty">尚无操作资源。请先选择订阅，在“资源浏览”中加载并勾选需要操作的资源。</div>';
    } else {
        container.innerHTML = groups.map(([subId, indices]) => `
            <section class="apply-sub-group">
                <div class="apply-sub-heading"><label><input type="checkbox" class="apply-sub-cb" data-sub-id="${escapeHtml(subId)}"> <strong>${escapeHtml(subscriptionMap[subId] || subId)}</strong></label><span>${indices.length} 个已选资源</span></div>
                <div class="apply-resource-grid">${indices.sort((a, b) => allResources[a].name.localeCompare(allResources[b].name)).map(idx => {
                    const r = allResources[idx], p = parseResourceId(r.id);
                    return `<label class="apply-resource-item"><input type="checkbox" class="apply-resource-cb" data-index="${idx}" ${applyScope.has(idx) ? 'checked' : ''}><span><strong>${escapeHtml(r.name)}</strong><small>${escapeHtml(p.resourceGroup)} · ${escapeHtml(r.location || '-')} · ${escapeHtml(r.kind || '-')}</small></span></label>`;
                }).join('')}</div>
            </section>`).join('');
        container.querySelectorAll('.apply-sub-cb').forEach(cb => {
            const children = [...container.querySelectorAll('.apply-resource-cb')].filter(item => allResources[Number(item.dataset.index)]._subId === cb.dataset.subId);
            cb.checked = children.every(item => item.checked);
            cb.indeterminate = !cb.checked && children.some(item => item.checked);
            cb.addEventListener('change', () => {
                children.forEach(item => cb.checked ? applyScope.add(Number(item.dataset.index)) : applyScope.delete(Number(item.dataset.index)));
                invalidateApplyData();
            });
        });
        container.querySelectorAll('.apply-resource-cb').forEach(cb => cb.addEventListener('change', () => {
            const idx = Number(cb.dataset.index);
            cb.checked ? applyScope.add(idx) : applyScope.delete(idx);
            invalidateApplyData();
        }));
    }
    document.getElementById('btn-apply-load-models').disabled = !applyScope.size || applyLoading || applyExecuting;
}

function invalidateApplyData() {
    applyLoadVersion++;
    applyData.clear();
    applyLoading = false;
    document.getElementById('apply-models-loading').classList.add('hidden');
    document.getElementById('apply-model-toolbar').classList.add('hidden');
    document.getElementById('apply-models-list').innerHTML = '<p class="empty-hint">资源范围已改变，请重新加载模型部署。</p>';
    document.getElementById('apply-progress-section').classList.add('hidden');
    renderApplyResources();
    updateApplySummary();
}

async function loadApplyModels() {
    if (applyLoading || applyExecuting) return;
    const indices = [...applyScope].filter(idx => getSelectedResourceIndices().includes(idx)).sort((a, b) => a - b);
    if (!indices.length) { renderApplyResources(); return; }
    const version = ++applyLoadVersion;
    applyLoading = true;
    applyData.clear();
    document.getElementById('apply-models-loading').classList.remove('hidden');
    document.getElementById('apply-models-list').innerHTML = '';
    document.getElementById('apply-model-toolbar').classList.add('hidden');
    document.getElementById('apply-progress-section').classList.add('hidden');
    document.getElementById('btn-apply-load-models').disabled = true;
    updateApplySummary();

    try {
        const results = await mapWithConcurrency(indices, REQUEST_CONCURRENCY, async idx => {
            const p = parseResourceId(allResources[idx].id);
            const [deployments, policies] = await Promise.allSettled([
                listDeployments(p.subscriptionId, p.resourceGroup, p.accountName),
                listRaiPolicies(p.subscriptionId, p.resourceGroup, p.accountName)
            ]);
            return {
                idx,
                deployments: deployments.status === 'fulfilled' ? deployments.value.sort((a, b) => a.name.localeCompare(b.name)) : [],
                policies: policies.status === 'fulfilled' ? policies.value.filter(pol => pol.properties?.type !== 'SystemManaged').sort((a, b) => a.name.localeCompare(b.name)) : [],
                deploymentError: deployments.status === 'rejected' ? String(deployments.reason?.message || deployments.reason) : '',
                policyError: policies.status === 'rejected' ? String(policies.reason?.message || policies.reason) : ''
            };
        });
        if (version !== applyLoadVersion) return; // Scope changed while requests were in flight.
        applyData = new Map(results.map(item => [item.idx, item]));
        document.getElementById('apply-model-toolbar').classList.remove('hidden');
        document.getElementById('apply-model-search').value = '';
        document.getElementById('apply-only-default').checked = false;
        renderApplyModels();
    } catch (err) {
        if (version === applyLoadVersion) document.getElementById('apply-models-list').innerHTML = `<p class="apply-error">加载失败：${escapeHtml(err.message || String(err))}</p>`;
    } finally {
        if (version === applyLoadVersion) {
            applyLoading = false;
            document.getElementById('apply-models-loading').classList.add('hidden');
            renderApplyResources();
            updateApplySummary();
        }
    }
}

function renderApplyModels() {
    const container = document.getElementById('apply-models-list');
    const groups = groupApplyResources([...applyScope].filter(idx => applyData.has(idx)));
    container.innerHTML = groups.map(([subId, indices]) => `
        <section class="apply-sub-group apply-model-sub">
            <div class="apply-sub-heading"><strong>${escapeHtml(subscriptionMap[subId] || subId)}</strong><span>${indices.length} 个资源</span></div>
            ${indices.map(idx => {
                const r = allResources[idx], data = applyData.get(idx);
                const options = data.policies.map(pol => `<option value="${escapeHtml(pol.name)}">${escapeHtml(pol.name)}</option>`).join('');
                const rows = data.deployments.map((dep, depIndex) => {
                    const current = dep.properties?.raiPolicyName;
                    const model = dep.properties?.model || {};
                    const system = !current || current.startsWith('Microsoft.');
                    return `<label class="apply-model-row" data-search="${escapeHtml(`${r.name} ${dep.name} ${model.name || ''} ${model.version || ''} ${current || ''}`.toLowerCase())}" data-system="${system}">
                        <input type="checkbox" class="apply-model-cb" data-res-index="${idx}" data-dep-index="${depIndex}">
                        <span class="apply-model-identity"><strong>${escapeHtml(dep.name)}</strong><small>${escapeHtml(model.name || '未知模型')}${model.version ? ` · ${escapeHtml(model.version)}` : ''}</small></span>
                        <span class="apply-current-filter ${system ? 'is-default' : ''}">当前：${escapeHtml(current || '未指定（默认）')}</span>
                    </label>`;
                }).join('');
                return `<div class="apply-model-resource" data-res-index="${idx}">
                    <div class="apply-model-resource-heading"><div><strong>${escapeHtml(r.name)}</strong><small>${escapeHtml(r.location || '-')} · ${data.deployments.length} 个部署</small></div>
                        <label class="apply-target-label">目标筛选器 <select class="toolbar-select apply-target-filter" data-res-index="${idx}" ${data.policyError || !data.policies.length ? 'disabled' : ''}><option value="">请选择…</option>${options}</select></label>
                    </div>
                    ${data.deploymentError ? `<p class="apply-error">部署加载失败：${escapeHtml(data.deploymentError)}</p>` : ''}
                    ${data.policyError ? `<p class="apply-error">筛选器加载失败：${escapeHtml(data.policyError)}</p>` : !data.policies.length ? '<p class="apply-help apply-no-policy">该资源没有自定义筛选器，需先创建后才能应用。</p>' : ''}
                    ${rows || (!data.deploymentError ? '<p class="empty-hint">该资源没有模型部署</p>' : '')}
                </div>`;
            }).join('')}
        </section>`).join('') || '<p class="empty-hint">没有可加载的资源。</p>';
    container.querySelectorAll('.apply-model-cb, .apply-target-filter').forEach(el => el.addEventListener('change', updateApplySummary));
    filterApplyModels();
}

function filterApplyModels() {
    const query = document.getElementById('apply-model-search').value.trim().toLowerCase();
    const onlyDefault = document.getElementById('apply-only-default').checked;
    document.querySelectorAll('.apply-model-row').forEach(row => {
        row.classList.toggle('hidden', !!query && !row.dataset.search.includes(query) || onlyDefault && row.dataset.system !== 'true');
    });
    document.querySelectorAll('.apply-model-resource').forEach(group => {
        const hasRows = group.querySelectorAll('.apply-model-row').length > 0;
        group.classList.toggle('hidden', hasRows && !group.querySelector('.apply-model-row:not(.hidden)'));
    });
    document.querySelectorAll('.apply-model-sub').forEach(group => group.classList.toggle('hidden', !group.querySelector('.apply-model-resource:not(.hidden)')));
}

function toggleVisibleApplyModels(checked, all = false) {
    document.querySelectorAll(all ? '.apply-model-cb' : '.apply-model-row:not(.hidden) .apply-model-cb').forEach(cb => { cb.checked = checked; });
    updateApplySummary();
}

function getApplyChanges() {
    const changes = [];
    let missingTarget = 0, unchanged = 0;
    document.querySelectorAll('.apply-model-cb:checked').forEach(cb => {
        const idx = Number(cb.dataset.resIndex), data = applyData.get(idx);
        if (!applyScope.has(idx) || !getSelectedResourceIndices().includes(idx) || !data) return;
        const dep = data.deployments[Number(cb.dataset.depIndex)];
        const target = document.querySelector(`.apply-target-filter[data-res-index="${idx}"]`)?.value;
        if (!target || !data.policies.some(pol => pol.name === target)) { missingTarget++; return; }
        if (dep.properties?.raiPolicyName === target) { unchanged++; return; }
        changes.push({ idx, dep, target, cb });
    });
    return { changes, missingTarget, unchanged };
}

function updateApplySummary() {
    const { changes, missingTarget, unchanged } = getApplyChanges();
    const selected = document.querySelectorAll('.apply-model-cb:checked').length;
    const summary = document.getElementById('apply-summary');
    const btn = document.getElementById('btn-apply-exec');
    if (!selected) {
        summary.innerHTML = '<p class="empty-hint">请加载模型，勾选要修改的部署，并为其资源选择目标筛选器。</p>';
    } else {
        const details = changes.map(({ idx, dep, target }) => `<tr><td>${escapeHtml(subscriptionMap[allResources[idx]._subId] || allResources[idx]._subId)}</td><td>${escapeHtml(allResources[idx].name)} / ${escapeHtml(dep.name)}</td><td>${escapeHtml(dep.properties?.raiPolicyName || '未指定（默认）')}</td><td>${escapeHtml(target)}</td></tr>`).join('');
        summary.innerHTML = `<strong>已勾选 ${selected} 个部署 · 待修改 ${changes.length} 个</strong>${missingTarget ? `<p class="apply-warning">${missingTarget} 个部署尚未选择目标筛选器，不会执行。</p>` : ''}${unchanged ? `<p class="apply-help">${unchanged} 个部署已应用目标筛选器，无需重复提交。</p>` : ''}${details ? `<div class="apply-preview-scroll"><table class="data-table"><thead><tr><th>订阅</th><th>资源 / 部署</th><th>当前筛选器</th><th>目标筛选器</th></tr></thead><tbody>${details}</tbody></table></div>` : ''}`;
    }
    btn.disabled = !changes.length || !!missingTarget || applyLoading || applyExecuting;
}

async function executeBatchApply() {
    if (applyLoading || applyExecuting) return;
    const { changes, missingTarget } = getApplyChanges();
    if (missingTarget || !changes.length) { updateApplySummary(); return; }
    const resourceCount = new Set(changes.map(item => item.idx)).size;
    if (!window.confirm(`即将修改 ${resourceCount} 个资源中的 ${changes.length} 个模型部署的内容筛选器。确认执行吗？`)) return;
    applyExecuting = true;
    document.getElementById('btn-apply-exec').disabled = true;
    document.getElementById('btn-apply-load-models').disabled = true;
    document.querySelectorAll('.apply-resource-cb, .apply-sub-cb, .apply-model-cb, .apply-target-filter').forEach(el => { el.disabled = true; });
    const section = document.getElementById('apply-progress-section');
    section.classList.remove('hidden');
    const logEl = document.getElementById('apply-log');
    logEl.innerHTML = '';
    updateProgress('apply-progress', 'apply-progress-text', 0, changes.length);
    let done = 0, errors = 0;
    for (const { idx, dep, target, cb } of changes) {
        const r = allResources[idx], p = parseResourceId(r.id);
        try {
            await updateDeploymentRaiPolicy(p.subscriptionId, p.resourceGroup, p.accountName, dep.name, target, dep);
            dep.properties = { ...dep.properties, raiPolicyName: target };
            cb.checked = false;
            cb.closest('.apply-model-row').querySelector('.apply-current-filter').textContent = `当前：${target}`;
            cb.closest('.apply-model-row').querySelector('.apply-current-filter').classList.remove('is-default');
            cb.closest('.apply-model-row').dataset.system = String(target.startsWith('Microsoft.'));
            log(logEl, `[${r.name}/${dep.name}] 已应用 ${target}`, 's');
            done++;
        } catch (e) {
            errors++;
            log(logEl, `[${r.name}/${dep.name}] 失败：${e.message}`, 'e');
        }
        updateProgress('apply-progress', 'apply-progress-text', done + errors, changes.length);
    }
    log(logEl, `完成：成功 ${done}，失败 ${errors}。失败项仍保持勾选，可核查后重试。`, errors ? 'w' : 's');
    addActivity(`批量应用筛选器 - 成功 ${done}，失败 ${errors}，共 ${changes.length}`);
    applyExecuting = false;
    document.querySelectorAll('.apply-resource-cb, .apply-sub-cb, .apply-model-cb').forEach(el => { el.disabled = false; });
    document.querySelectorAll('.apply-target-filter').forEach(el => { el.disabled = !applyData.get(Number(el.dataset.resIndex))?.policies.length; });
    renderApplyResources();
    filterApplyModels();
    updateApplySummary();
}

// Load filters for batch-delete page
async function loadDeleteFilters() {
    const indices = getSelectedResourceIndices();
    if (indices.length === 0) { alert('请先在"资源浏览"中选择资源'); navigateTo('resources'); return; }
    
    const container = document.getElementById('delete-filters-list');
    const loading = document.getElementById('delete-filters-loading');
    loading.classList.remove('hidden');
    container.innerHTML = '';
    
    try {
        for (const idx of indices) {
            const r = allResources[idx];
            const p = parseResourceId(r.id);
            const policies = await listRaiPolicies(p.subscriptionId, p.resourceGroup, p.accountName);
            const customPolicies = policies.filter(pol => pol.properties?.type !== 'SystemManaged');
            if (customPolicies.length > 0) {
                let html = `<div class="filter-select-group"><div class="filter-select-group-header"><span class="res-name">${r.name}</span><span class="res-loc">${r.location}</span></div>`;
                customPolicies.forEach(pol => {
                    html += `<label class="filter-select-item"><input type="checkbox" class="delete-filter-cb" data-res-index="${idx}" data-filter-name="${pol.name}" checked><span class="fname">${pol.name}</span><span class="fbadge custom">自定义</span></label>`;
                });
                html += '</div>';
                container.innerHTML += html;
            }
        }
        loading.classList.add('hidden');
        if (!container.innerHTML) container.innerHTML = '<p class="empty-hint">已选资源中没有自定义筛选器</p>';
        container.querySelectorAll('.delete-filter-cb').forEach(cb => {
            cb.addEventListener('change', () => {
                document.getElementById('btn-delete-exec').disabled = document.querySelectorAll('.delete-filter-cb:checked').length === 0;
            });
        });
        document.getElementById('btn-delete-exec').disabled = document.querySelectorAll('.delete-filter-cb:checked').length === 0;
    } catch (err) {
        loading.classList.add('hidden');
        container.innerHTML = `<p style="color:var(--danger)">${err.message}</p>`;
    }
}

async function executeBatchDelete() {
    const checked = document.querySelectorAll('.delete-filter-cb:checked');
    if (checked.length === 0) { alert('请选择要删除的筛选器'); return; }
    
    const items = [];
    checked.forEach(cb => items.push({ resIndex: parseInt(cb.dataset.resIndex), filterName: cb.dataset.filterName }));
    if (!confirm(`确定要删除 ${items.length} 个筛选器吗？`)) return;
    
    const section = document.getElementById('delete-progress-section');
    section.classList.remove('hidden');
    const logEl = document.getElementById('delete-log');
    logEl.innerHTML = '';
    let done = 0;
    const total = items.length;
    
    document.getElementById('btn-delete-exec').disabled = true;
    
    for (const item of items) {
        const r = allResources[item.resIndex];
        const p = parseResourceId(r.id);
        try {
            await deleteRaiPolicy(p.subscriptionId, p.resourceGroup, p.accountName, item.filterName);
            log(logEl, `[${r.name}] ✓ 已删除 "${item.filterName}"`, 's');
            done++;
        } catch (err) {
            log(logEl, `[${r.name}] ✗ 删除 "${item.filterName}" 失败: ${err.message}`, 'e');
        }
        updateProgress('delete-progress', null, done, total);
    }
    log(logEl, `完成: ${done}/${total}`, done === total ? 's' : 'w');
    addActivity(`批量删除筛选器 - ${done}/${total}`);
    document.getElementById('btn-delete-exec').disabled = false;
}

// ============ Helpers ============
function getSelectedResourceIndices() {
    return Array.from(selectedResources)
        .filter(idx => allResources[idx] && selectedSubscriptions.has(allResources[idx]._subId))
        .sort((a, b) => a - b);
}

function gatherFilterConfig() {
    const config = { inputFilters: {}, outputFilters: {}, otherFilters: [], mode: document.getElementById('filter-mode').value };
    
    document.querySelectorAll('#input-filters-body .filter-action').forEach(sel => {
        const name = sel.dataset.name;
        const row = sel.closest('tr');
        const slider = row.querySelector('.severity-slider');
        const action = sel.value;
        config.inputFilters[name] = {
            enabled: action !== 'off',
            blocking: action === 'annotate_and_block',
            severityThreshold: action === 'annotate_and_block' ? SLIDER_TO_API[slider.value] : 'High'
        };
    });
    
    document.querySelectorAll('#output-filters-body .filter-action').forEach(sel => {
        const name = sel.dataset.name;
        const row = sel.closest('tr');
        const slider = row.querySelector('.severity-slider');
        const action = sel.value;
        config.outputFilters[name] = {
            enabled: action !== 'off',
            blocking: action === 'annotate_and_block',
            severityThreshold: action === 'annotate_and_block' ? SLIDER_TO_API[slider.value] : 'High'
        };
    });
    
    document.querySelectorAll('#other-filters-body .other-filter-action').forEach(sel => {
        const action = sel.value;
        config.otherFilters.push({
            name: sel.dataset.name,
            source: sel.dataset.source,
            enabled: action !== 'off',
            blocking: action === 'annotate_and_block'
        });
    });
    
    return config;
}

function updateProgress(barId, textId, current, total) {
    const pct = total > 0 ? Math.round(current / total * 100) : 0;
    document.getElementById(barId).style.width = pct + '%';
    if (textId) document.getElementById(textId).textContent = `${pct}% (${current}/${total})`;
}

function log(container, msg, type) {
    const div = document.createElement('div');
    div.className = 'log-' + type;
    div.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
}

function addActivity(msg) {
    activityLog.unshift({ time: new Date().toLocaleString(), msg });
    const el = document.getElementById('activity-log');
    el.innerHTML = activityLog.map(a => `<div class="log-i">[${a.time}] ${a.msg}</div>`).join('');
}

function updateStats() {
    document.getElementById('stat-subs').textContent = selectedSubscriptions.size;
    document.getElementById('stat-resources').textContent = allResources.length;
}
