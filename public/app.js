// Main Application Logic
let selectedSubscriptions = new Set();
let allResources = [];
let selectedResources = new Set();
let currentWizStep = 1;
const MAX_WIZ_STEPS = 5;
let activityLog = [];

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
    
    // Resources
    document.getElementById('btn-load-resources').addEventListener('click', loadResources);
    document.getElementById('btn-sel-all-res').addEventListener('click', () => toggleAllRes(true));
    document.getElementById('btn-desel-all-res').addEventListener('click', () => toggleAllRes(false));
    document.getElementById('res-check-all').addEventListener('change', (e) => toggleAllRes(e.target.checked));
    document.getElementById('res-search').addEventListener('input', filterResourcesList);
    document.getElementById('res-type-filter').addEventListener('change', filterResourcesList);
    
    // Filters View
    document.getElementById('btn-load-filters').addEventListener('click', loadExistingFilters);
    
    // Wizard
    document.getElementById('wiz-prev').addEventListener('click', wizPrev);
    document.getElementById('wiz-next').addEventListener('click', wizNext);
    document.getElementById('wiz-exec').addEventListener('click', executeWizard);
    
    // Presets
    document.querySelectorAll('.preset-btn').forEach(btn => {
        btn.addEventListener('click', () => applyPreset(btn.dataset.preset));
    });
    
    // Batch Apply
    document.getElementById('btn-apply-load-filters').addEventListener('click', loadApplyFilters);
    document.getElementById('btn-apply-load-models').addEventListener('click', loadApplyModels);
    document.getElementById('btn-apply-exec').addEventListener('click', executeBatchApply);
    
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
}

// ============ Subscriptions ============
async function loadSubscriptions() {
    const loading = document.getElementById('subs-loading');
    const tbody = document.getElementById('subs-tbody');
    loading.classList.remove('hidden');
    tbody.innerHTML = '';
    
    try {
        const subs = await listSubscriptions();
        loading.classList.add('hidden');
        subs.forEach(sub => {
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
                updateStats();
            });
            tbody.appendChild(tr);
        });
        updateStats();
    } catch (err) {
        loading.classList.add('hidden');
        tbody.innerHTML = `<tr><td colspan="4" style="color:var(--danger);padding:16px;">加载失败: ${err.message}</td></tr>`;
    }
}

function toggleAllSubs(checked) {
    document.querySelectorAll('.sub-cb').forEach(cb => {
        cb.checked = checked;
        if (checked) selectedSubscriptions.add(cb.dataset.id);
        else selectedSubscriptions.delete(cb.dataset.id);
    });
    document.getElementById('subs-check-all').checked = checked;
    updateStats();
}

function filterSubsList() {
    const q = document.getElementById('sub-search').value.toLowerCase();
    document.querySelectorAll('#subs-tbody tr').forEach(tr => {
        const text = tr.textContent.toLowerCase();
        tr.style.display = text.includes(q) ? '' : 'none';
    });
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
    loading.classList.remove('hidden');
    tbody.innerHTML = '';
    allResources = [];
    
    try {
        for (const subId of selectedSubscriptions) {
            const resources = await listCognitiveServicesAccounts(subId);
            resources.forEach(r => { r._subId = subId; allResources.push(r); });
        }
        loading.classList.add('hidden');
        renderResources(allResources);
        updateStats();
    } catch (err) {
        loading.classList.add('hidden');
        tbody.innerHTML = `<tr><td colspan="6" style="color:var(--danger);padding:16px;">加载失败: ${err.message}</td></tr>`;
    }
}

function renderResources(resources) {
    const tbody = document.getElementById('res-tbody');
    tbody.innerHTML = '';
    if (resources.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-hint">未找到 OpenAI/AI Services 资源</td></tr>';
        return;
    }
    resources.forEach((r, i) => {
        const p = parseResourceId(r.id);
        const tr = document.createElement('tr');
        tr.dataset.index = i;
        tr.innerHTML = `
            <td class="w40"><input type="checkbox" class="res-cb" data-index="${i}" ${selectedResources.has(i) ? 'checked' : 'checked'}></td>
            <td><strong>${r.name}</strong></td>
            <td><span class="media-tag">${r.kind || '-'}</span></td>
            <td>${p.resourceGroup}</td>
            <td>${r.location}</td>
            <td style="font-size:12px;color:var(--text-muted)">${r._subId.substring(0,8)}...</td>
        `;
        const cb = tr.querySelector('.res-cb');
        if (cb.checked) selectedResources.add(i);
        cb.addEventListener('change', (e) => {
            if (e.target.checked) selectedResources.add(i);
            else selectedResources.delete(i);
        });
        tbody.appendChild(tr);
    });
}

function toggleAllRes(checked) {
    document.querySelectorAll('.res-cb').forEach(cb => {
        cb.checked = checked;
        const idx = parseInt(cb.dataset.index);
        if (checked) selectedResources.add(idx);
        else selectedResources.delete(idx);
    });
    document.getElementById('res-check-all').checked = checked;
}

function filterResourcesList() {
    const q = document.getElementById('res-search').value.toLowerCase();
    const typeFilter = document.getElementById('res-type-filter').value;
    document.querySelectorAll('#res-tbody tr').forEach(tr => {
        const idx = parseInt(tr.dataset.index);
        if (isNaN(idx)) return;
        const r = allResources[idx];
        const matchText = tr.textContent.toLowerCase().includes(q);
        const matchType = !typeFilter || (r.kind || '').includes(typeFilter);
        tr.style.display = (matchText && matchType) ? '' : 'none';
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
    
    let totalFilters = 0;
    const showSystem = document.getElementById('filters-show-system').checked;
    try {
        for (const idx of indices) {
            const r = allResources[idx];
            const p = parseResourceId(r.id);
            try {
                const policies = await listRaiPolicies(p.subscriptionId, p.resourceGroup, p.accountName);
                const filteredPolicies = showSystem ? policies : policies.filter(pol => pol.properties?.type !== 'SystemManaged');
                if (filteredPolicies.length > 0) {
                    let html = `<div class="filter-group"><div class="filter-group-header"><span>${r.name} (${r.location})</span><span>${filteredPolicies.length} 个筛选器</span></div>`;
                    filteredPolicies.forEach(pol => {
                        const isSystem = pol.properties?.type === 'SystemManaged';
                        totalFilters += isSystem ? 0 : 1;
                        html += `<div class="filter-item"><span>${pol.name}</span><span class="filter-badge ${isSystem ? 'system' : 'custom'}">${isSystem ? '系统' : '自定义'}</span></div>`;
                    });
                    html += '</div>';
                    container.innerHTML += html;
                }
            } catch (e) {
                container.innerHTML += `<div class="filter-group"><div class="filter-group-header"><span>${r.name}</span><span style="color:var(--danger)">加载失败</span></div></div>`;
            }
        }
        if (container.innerHTML === '') container.innerHTML = '<p class="empty-hint">已选资源中未找到筛选器</p>';
        loading.classList.add('hidden');
        document.getElementById('stat-filters').textContent = totalFilters;
    } catch (err) {
        loading.classList.add('hidden');
        container.innerHTML = `<p style="color:var(--danger)">加载失败: ${err.message}</p>`;
    }
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
    document.getElementById('wiz-exec').classList.toggle('hidden', currentWizStep !== MAX_WIZ_STEPS);
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
async function executeWizard() {
    const filterName = document.getElementById('filter-name').value.trim();
    const applyToModels = document.getElementById('apply-to-models').checked;
    const indices = getSelectedResourceIndices();
    
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
    const logEl = document.getElementById('exec-log');
    logEl.innerHTML = '';
    
    let done = 0, errors = 0, fallbackUsed = 0;
    const total = indices.length;
    
    log(logEl, `开始批量创建筛选器: ${filterName}`, 'i');
    log(logEl, `目标资源: ${total} 个`, 'i');
    if (needsApproval && fallbackEnabled) {
        const fbAction = document.getElementById('fallback-action').value;
        if (fbAction === 'skip') {
            log(logEl, `降级策略: 跳过（权限不足时不创建）`, 'w');
        } else {
            log(logEl, `降级策略: 已启用 → ${fbAction.replace('annotate_and_block_', 'Annotate and block - ')} (名称: ${fallbackFilterName})`, 'w');
        }
    }
    log(logEl, '─'.repeat(50), 'i');
    
    document.getElementById('wiz-exec').disabled = true;
    
    for (const idx of indices) {
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
                        log(logEl, `[${r.name}] ✗ 降级也失败: ${fbErr.message}`, 'e');
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
                log(logEl, `[${r.name}] ✗ 失败: ${(err.message || String(err)).substring(0, 300)}`, 'e');
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
                    } catch (e) {
                        log(logEl, `[${r.name}] ✗ 应用失败 ${dep.name}: ${e.message}`, 'e');
                    }
                }
            } catch (e) {
                log(logEl, `[${r.name}] ✗ 获取部署列表失败: ${e.message}`, 'e');
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
    log(logEl, `完成! 成功: ${done}, 降级: ${fallbackUsed}, 失败: ${errors}, 共 ${total}`, 
        errors === 0 ? 's' : 'w');
    if (fallbackUsed > 0) {
        log(logEl, `⚠ ${fallbackUsed} 个资源使用了降级配置（权限不足）`, 'w');
    }
    addActivity(`批量创建 "${filterName}" - 成功${done} 降级${fallbackUsed} 失败${errors}/${total}`);
    document.getElementById('wiz-exec').disabled = false;
}

async function executeBatchApply() {
    const checkedFilters = document.querySelectorAll('.apply-filter-cb:checked');
    if (checkedFilters.length === 0) { alert('请先选择筛选器'); return; }
    
    const checkedModels = document.querySelectorAll('.apply-model-cb:checked');
    if (checkedModels.length === 0) { alert('请先选择要应用的模型部署'); return; }
    
    // Build a map: resIndex -> filterName
    const filterMap = {};
    checkedFilters.forEach(cb => {
        const idx = cb.dataset.resIndex;
        filterMap[idx] = cb.value;
    });
    
    // Only apply to models whose resource has a selected filter
    const modelsToApply = [];
    checkedModels.forEach(cb => {
        const resIdx = cb.dataset.resIndex;
        if (filterMap[resIdx]) {
            modelsToApply.push({ resIdx: parseInt(resIdx), depName: cb.dataset.depName, depData: JSON.parse(cb.dataset.depData), filterName: filterMap[resIdx] });
        }
    });
    
    if (modelsToApply.length === 0) { alert('请确保选择了筛选器对应资源下的模型部署'); return; }
    
    const section = document.getElementById('apply-progress-section');
    section.classList.remove('hidden');
    const logEl = document.getElementById('apply-log');
    logEl.innerHTML = '';
    let done = 0, errors = 0;
    const total = modelsToApply.length;
    
    document.getElementById('btn-apply-exec').disabled = true;
    log(logEl, `开始批量应用，共 ${total} 个部署`, 'i');
    
    for (const item of modelsToApply) {
        const r = allResources[item.resIdx];
        const p = parseResourceId(r.id);
        try {
            await updateDeploymentRaiPolicy(p.subscriptionId, p.resourceGroup, p.accountName, item.depName, item.filterName, item.depData);
            log(logEl, `[${r.name}/${item.depName}] ✓ 应用 "${item.filterName}" 成功`, 's');
            done++;
        } catch (e) {
            errors++;
            log(logEl, `[${r.name}/${item.depName}] ✗ ${e.message}`, 'e');
        }
        updateProgress('apply-progress', 'apply-progress-text', done + errors, total);
    }
    log(logEl, `完成: 成功 ${done}, 失败 ${errors}`, done === total ? 's' : 'w');
    addActivity(`批量应用筛选器 - ${done}/${total}`);
    document.getElementById('btn-apply-exec').disabled = false;
}

// Load filters for batch-apply page
async function loadApplyFilters() {
    const indices = getSelectedResourceIndices();
    if (indices.length === 0) { alert('请先在"资源浏览"中选择资源'); navigateTo('resources'); return; }
    
    const container = document.getElementById('apply-filters-list');
    const loading = document.getElementById('apply-filters-loading');
    loading.classList.remove('hidden');
    container.innerHTML = '';
    
    try {
        for (const idx of indices) {
            const r = allResources[idx];
            const p = parseResourceId(r.id);
            const policies = await listRaiPolicies(p.subscriptionId, p.resourceGroup, p.accountName);
            const customPolicies = policies.filter(pol => pol.properties?.type !== 'SystemManaged');
            if (customPolicies.length > 0) {
                let html = `<div class="filter-select-group"><div class="filter-select-group-header"><span class="res-name">${r.name}</span><span class="res-loc">${r.location} · ${customPolicies.length} 个筛选器</span></div>`;
                customPolicies.forEach(pol => {
                    html += `<label class="filter-select-item"><input type="checkbox" class="apply-filter-cb" value="${pol.name}" data-res-index="${idx}"><span class="fname">${pol.name}</span><span class="fbadge custom">自定义</span></label>`;
                });
                html += '</div>';
                container.innerHTML += html;
            }
        }
        loading.classList.add('hidden');
        if (!container.innerHTML) container.innerHTML = '<p class="empty-hint"><i class="fas fa-info-circle"></i> 已选资源中没有自定义筛选器，请先在"批量创建"中创建</p>';
        
        container.querySelectorAll('.apply-filter-cb').forEach(cb => {
            cb.addEventListener('change', updateApplySummary);
        });
    } catch (err) {
        loading.classList.add('hidden');
        container.innerHTML = `<p style="color:var(--danger)">${err.message}</p>`;
    }
}

// Load models for batch-apply page
async function loadApplyModels() {
    const indices = getSelectedResourceIndices();
    if (indices.length === 0) { alert('请先在"资源浏览"中选择资源'); return; }
    
    const container = document.getElementById('apply-models-list');
    const loading = document.getElementById('apply-models-loading');
    loading.classList.remove('hidden');
    container.innerHTML = '';
    
    try {
        for (const idx of indices) {
            const r = allResources[idx];
            const p = parseResourceId(r.id);
            const deployments = await listDeployments(p.subscriptionId, p.resourceGroup, p.accountName);
            if (deployments.length > 0) {
                let html = `<div class="filter-select-group"><div class="filter-select-group-header"><input type="checkbox" class="apply-res-all" data-res-index="${idx}" checked> <span class="res-name">${r.name}</span><span class="res-loc">${r.location} · ${deployments.length} 个部署</span></div>`;
                deployments.forEach(dep => {
                    const currentFilter = dep.properties?.raiPolicyName || 'Microsoft.Default';
                    const modelName = dep.properties?.model?.name || '';
                    const modelVer = dep.properties?.model?.version || '';
                    html += `<label class="filter-select-item"><input type="checkbox" class="apply-model-cb" data-res-index="${idx}" data-dep-name="${dep.name}" data-dep-data='${JSON.stringify(dep).replace(/'/g, "&#39;")}' checked><span class="fname">${dep.name}<span class="model-current-filter">(${modelName} ${modelVer})</span></span><span class="fbadge system">当前: ${currentFilter}</span></label>`;
                });
                html += '</div>';
                container.innerHTML += html;
            }
        }
        loading.classList.add('hidden');
        if (!container.innerHTML) container.innerHTML = '<p class="empty-hint">已选资源中没有模型部署</p>';
        
        // Select all toggle per resource
        container.querySelectorAll('.apply-res-all').forEach(cb => {
            cb.addEventListener('change', (e) => {
                const idx = e.target.dataset.resIndex;
                container.querySelectorAll(`.apply-model-cb[data-res-index="${idx}"]`).forEach(mcb => mcb.checked = e.target.checked);
                updateApplySummary();
            });
        });
        container.querySelectorAll('.apply-model-cb').forEach(cb => cb.addEventListener('change', updateApplySummary));
        updateApplySummary();
    } catch (err) {
        loading.classList.add('hidden');
        container.innerHTML = `<p style="color:var(--danger)">${err.message}</p>`;
    }
}

function updateApplySummary() {
    const checkedFilters = document.querySelectorAll('.apply-filter-cb:checked');
    const checkedModels = document.querySelectorAll('.apply-model-cb:checked');
    const summary = document.getElementById('apply-summary');
    const btn = document.getElementById('btn-apply-exec');
    
    if (checkedFilters.length > 0 && checkedModels.length > 0) {
        // Count how many models match selected filter resources
        const filterResIndices = new Set();
        checkedFilters.forEach(cb => filterResIndices.add(cb.dataset.resIndex));
        let matchCount = 0;
        checkedModels.forEach(cb => { if (filterResIndices.has(cb.dataset.resIndex)) matchCount++; });
        
        let details = '';
        checkedFilters.forEach(cb => {
            const r = allResources[parseInt(cb.dataset.resIndex)];
            const modelCount = document.querySelectorAll(`.apply-model-cb[data-res-index="${cb.dataset.resIndex}"]:checked`).length;
            if (modelCount > 0) details += `<div style="font-size:12px;color:var(--text-secondary)">${r.name}: "${cb.value}" → ${modelCount} 个部署</div>`;
        });
        
        summary.innerHTML = `<div style="padding:8px 0"><strong>${checkedFilters.length}</strong> 个筛选器 → <strong>${matchCount}</strong> 个模型部署</div>${details}`;
        btn.disabled = matchCount === 0;
    } else {
        summary.innerHTML = '<p class="empty-hint"><i class="fas fa-info-circle"></i> 请选择筛选器和模型部署</p>';
        btn.disabled = true;
    }
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
    const indices = [];
    document.querySelectorAll('.res-cb:checked').forEach(cb => indices.push(parseInt(cb.dataset.index)));
    return indices.length > 0 ? indices : Array.from(selectedResources);
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
