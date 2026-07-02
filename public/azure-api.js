// Azure Management API calls
const API_VERSION = '2024-10-01';
const ARM_BASE = 'https://management.azure.com';

async function azureRequest(url, method = 'GET', body = null) {
    const token = await getAccessToken();
    const options = {
        method,
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
        }
    };
    if (body) {
        options.body = JSON.stringify(body);
    }
    
    const response = await fetch(url, options);
    if (!response.ok) {
        const errorBody = await response.text();
        throw new Error(`API Error ${response.status}: ${errorBody}`);
    }
    return response.json();
}

// List all subscriptions
async function listSubscriptions() {
    const url = `${ARM_BASE}/subscriptions?api-version=2022-12-01`;
    const result = await azureRequest(url);
    return result.value || [];
}

// List Cognitive Services accounts (OpenAI) in a subscription
async function listCognitiveServicesAccounts(subscriptionId) {
    const url = `${ARM_BASE}/subscriptions/${subscriptionId}/providers/Microsoft.CognitiveServices/accounts?api-version=${API_VERSION}`;
    const result = await azureRequest(url);
    return (result.value || []).filter(account => 
        account.kind === 'OpenAI' || 
        account.kind === 'AIServices' ||
        account.kind === 'CognitiveServices'
    );
}

// List deployments for a Cognitive Services account
async function listDeployments(subscriptionId, resourceGroup, accountName) {
    const url = `${ARM_BASE}/subscriptions/${subscriptionId}/resourceGroups/${resourceGroup}/providers/Microsoft.CognitiveServices/accounts/${accountName}/deployments?api-version=${API_VERSION}`;
    const result = await azureRequest(url);
    return result.value || [];
}

// List RAI policies for a Cognitive Services account
async function listRaiPolicies(subscriptionId, resourceGroup, accountName) {
    const url = `${ARM_BASE}/subscriptions/${subscriptionId}/resourceGroups/${resourceGroup}/providers/Microsoft.CognitiveServices/accounts/${accountName}/raiPolicies?api-version=${API_VERSION}`;
    const result = await azureRequest(url);
    return result.value || [];
}

// Create or update RAI policy
async function createOrUpdateRaiPolicy(subscriptionId, resourceGroup, accountName, policyName, policyBody) {
    const url = `${ARM_BASE}/subscriptions/${subscriptionId}/resourceGroups/${resourceGroup}/providers/Microsoft.CognitiveServices/accounts/${accountName}/raiPolicies/${policyName}?api-version=${API_VERSION}`;
    return await azureRequest(url, 'PUT', policyBody);
}

// Delete RAI policy
async function deleteRaiPolicy(subscriptionId, resourceGroup, accountName, policyName) {
    const url = `${ARM_BASE}/subscriptions/${subscriptionId}/resourceGroups/${resourceGroup}/providers/Microsoft.CognitiveServices/accounts/${accountName}/raiPolicies/${policyName}?api-version=${API_VERSION}`;
    const token = await getAccessToken();
    const response = await fetch(url, {
        method: 'DELETE',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
    });
    if (!response.ok && response.status !== 204) {
        const errorBody = await response.text();
        throw new Error(`Delete failed ${response.status}: ${errorBody}`);
    }
    return true;
}

// Update deployment to apply content filter
async function updateDeploymentRaiPolicy(subscriptionId, resourceGroup, accountName, deploymentName, raiPolicyName, deploymentData) {
    const url = `${ARM_BASE}/subscriptions/${subscriptionId}/resourceGroups/${resourceGroup}/providers/Microsoft.CognitiveServices/accounts/${accountName}/deployments/${deploymentName}?api-version=${API_VERSION}`;
    
    // Update the deployment with the new RAI policy
    const body = {
        ...deploymentData,
        properties: {
            ...deploymentData.properties,
            raiPolicyName: raiPolicyName
        }
    };
    
    return await azureRequest(url, 'PUT', body);
}

// Build RAI policy body from configuration
function buildRaiPolicyBody(config) {
    const contentFilters = [];
    
    // Input (Prompt) filters
    const inputCategories = ['Hate', 'Sexual', 'Selfharm', 'Violence'];
    inputCategories.forEach(category => {
        const filterConfig = config.inputFilters[category];
        if (filterConfig) {
            contentFilters.push({
                name: category,
                enabled: filterConfig.enabled,
                blocking: filterConfig.blocking,
                severityThreshold: filterConfig.severityThreshold || 'High',
                source: 'Prompt'
            });
        }
    });
    
    // Output (Completion) filters
    const outputCategories = ['Hate', 'Sexual', 'Selfharm', 'Violence'];
    outputCategories.forEach(category => {
        const filterConfig = config.outputFilters[category];
        if (filterConfig) {
            contentFilters.push({
                name: category,
                enabled: filterConfig.enabled,
                blocking: filterConfig.blocking,
                severityThreshold: filterConfig.severityThreshold || 'High',
                source: 'Completion'
            });
        }
    });
    
    // Other filters
    if (config.otherFilters) {
        config.otherFilters.forEach(filter => {
            contentFilters.push({
                name: filter.name,
                enabled: filter.enabled,
                blocking: filter.blocking,
                source: filter.source
            });
        });
    }
    
    return {
        properties: {
            basePolicyName: 'Microsoft.Default',
            mode: config.mode || 'Asynchronous_filter',
            contentFilters: contentFilters
        }
    };
}

// Parse resource ID to extract subscription, resource group, and account name
function parseResourceId(resourceId) {
    const parts = resourceId.split('/');
    return {
        subscriptionId: parts[2],
        resourceGroup: parts[4],
        accountName: parts[8]
    };
}
