# Azure OpenAI 内容筛选器批量管理工具 - 实现文档

## 项目概述

本工具通过调用 Azure Management REST API，实现对多个订阅、多个 OpenAI/Foundry 资源的内容筛选器(RAI Policy)进行批量创建、应用和删除操作。

**线上地址**: http://content-filter-manager.azurewebsites.net

---

## 技术架构

```
┌────────────────────────────────────────────────────────────┐
│                    浏览器 (前端 SPA)                         │
│                                                            │
│  ┌──────────┐   ┌──────────────┐   ┌────────────────────┐ │
│  │ MSAL.js  │   │  Azure API   │   │   UI (HTML/CSS/JS) │ │
│  │ 认证模块  │──▶│  调用模块     │──▶│   展示和交互       │ │
│  └──────────┘   └──────────────┘   └────────────────────┘ │
│       │                │                                    │
└───────│────────────────│────────────────────────────────────┘
        │                │
        ▼                ▼
┌──────────────┐  ┌─────────────────────────────────────┐
│ Azure AD     │  │ Azure Resource Manager REST API      │
│ (login.ms)   │  │                                     │
│              │  │ • /subscriptions                     │
│ OAuth 2.0    │  │ • /providers/Microsoft.Cognitive...  │
│ PKCE Flow    │  │ • /raiPolicies (创建/删除/列表)      │
└──────────────┘  │ • /deployments (更新应用筛选器)      │
                  └─────────────────────────────────────┘
```

## 认证流程

### 使用 MSAL.js 进行 Azure AD 认证

```javascript
// 使用 Azure CLI 的公共 Client ID (无需注册应用)
const msalConfig = {
    auth: {
        clientId: '04b07795-a71b-4346-935c-9e63cf67e57b',  // Azure CLI public client
        authority: 'https://login.microsoftonline.com/common',
        redirectUri: window.location.origin
    }
};

// 请求 Azure Management API 的访问令牌
const loginRequest = {
    scopes: ['https://management.azure.com/.default']
};
```

**认证步骤**：
1. 用户点击登录 → 弹出 Microsoft 登录窗口
2. MSAL.js 使用 OAuth 2.0 PKCE 授权码流程获取 access_token
3. Token 具有 Azure Resource Manager 的访问权限
4. 后续所有 API 调用携带 `Authorization: Bearer {token}` 请求头

---

## 核心 API 调用

### 1. 获取订阅列表

```
GET https://management.azure.com/subscriptions?api-version=2022-12-01
Authorization: Bearer {access_token}
```

**响应**：返回用户有权访问的所有 Azure 订阅列表。

### 2. 获取 OpenAI/AI Services 资源

```
GET https://management.azure.com/subscriptions/{subscriptionId}/providers/Microsoft.CognitiveServices/accounts?api-version=2024-10-01
Authorization: Bearer {access_token}
```

**过滤条件**：`kind === 'OpenAI' || kind === 'AIServices'`

### 3. 创建/更新内容筛选器 (RAI Policy)

```
PUT https://management.azure.com/subscriptions/{subscriptionId}/resourceGroups/{resourceGroup}/providers/Microsoft.CognitiveServices/accounts/{accountName}/raiPolicies/{policyName}?api-version=2024-10-01
Authorization: Bearer {access_token}
Content-Type: application/json

{
  "properties": {
    "basePolicyName": "Microsoft.Default",
    "mode": "Asynchronous_filter",
    "contentFilters": [...]
  }
}
```

### 4. 获取模型部署列表

```
GET https://management.azure.com/subscriptions/{subscriptionId}/resourceGroups/{resourceGroup}/providers/Microsoft.CognitiveServices/accounts/{accountName}/deployments?api-version=2024-10-01
```

### 5. 应用筛选器到模型部署

```
PUT https://management.azure.com/subscriptions/{subscriptionId}/resourceGroups/{resourceGroup}/providers/Microsoft.CognitiveServices/accounts/{accountName}/deployments/{deploymentName}?api-version=2024-10-01

{
  ...原有部署属性,
  "properties": {
    ...原有属性,
    "raiPolicyName": "custom-filter-open"  // 关联筛选器名称
  }
}
```

### 6. 删除筛选器

```
DELETE https://management.azure.com/subscriptions/{subscriptionId}/resourceGroups/{resourceGroup}/providers/Microsoft.CognitiveServices/accounts/{accountName}/raiPolicies/{policyName}?api-version=2024-10-01
```

---

## 内容筛选器配置详解

### 筛选器 Body 结构

```json
{
  "properties": {
    "basePolicyName": "Microsoft.Default",
    "mode": "Asynchronous_filter",
    "contentFilters": [
      {
        "name": "Hate",
        "enabled": true,
        "blocking": true,
        "severityThreshold": "Medium",
        "source": "Prompt"
      },
      {
        "name": "Hate",
        "enabled": true,
        "blocking": true,
        "severityThreshold": "Medium",
        "source": "Completion"
      }
      // ... 其他类别
    ]
  }
}
```

### 操作模式对照表

| UI 选项 | enabled | blocking | severityThreshold | 说明 |
|---------|---------|----------|-------------------|------|
| Annotate and block + Low | true | true | Low | 最严格：筛选 Low+Medium+High |
| Annotate and block + Medium | true | true | Medium | 默认：筛选 Medium+High |
| Annotate and block + High | true | true | High | 宽松：仅筛选 High |
| Annotate only * | true | false | - | 仅标注不阻止（需审批） |
| Off (关闭) * | false | false | - | 完全关闭（需审批） |

> **注意**：标 * 选项需要订阅已通过 [Modified Content Filters](https://ncv.microsoft.com/uEfCgnITdR) 审批。

### 四大核心内容类别

每个类别分别配置 Prompt（输入）和 Completion（输出）：

| 类别 | 名称 | 说明 |
|------|------|------|
| Violence | 暴力 | 暴力相关内容 |
| Hate | 仇恨 | 歧视和仇恨言论 |
| Sexual | 性内容 | 色情相关内容 |
| Selfharm | 自伤 | 自我伤害相关内容 |

### 其他筛选器

| 名称 | 来源 | 说明 |
|------|------|------|
| Jailbreak | Prompt | 越狱攻击检测 |
| Protected Material Text | Completion | 受保护文本检测 |
| Protected Material Code | Completion | 受保护代码检测 |
| Profanity | Prompt | 亵渎/脏话检测 |

---

## 批量操作流程

### 批量创建筛选器流程

```
用户配置筛选器参数
        │
        ▼
选择目标资源 (可跨多个订阅)
        │
        ▼
┌───────────────────────────────────┐
│  FOR each selected resource:      │
│                                   │
│  1. PUT /raiPolicies/{name}       │
│     → 创建自定义筛选器             │
│                                   │
│  2. IF apply_to_models:           │
│     a. GET /deployments           │
│        → 获取所有模型部署          │
│     b. FOR each deployment:       │
│        PUT /deployments/{name}    │
│        → 更新 raiPolicyName 字段   │
│                                   │
│  3. 记录结果 (成功/失败)           │
└───────────────────────────────────┘
        │
        ▼
显示执行结果和日志
```

### 核心实现代码

```javascript
async function executeWizard() {
    const filterName = 'custom-filter-open';
    const policyBody = buildRaiPolicyBody(config);
    
    for (const resource of selectedResources) {
        const { subscriptionId, resourceGroup, accountName } = parseResourceId(resource.id);
        
        // Step 1: 创建筛选器
        await createOrUpdateRaiPolicy(subscriptionId, resourceGroup, accountName, filterName, policyBody);
        
        // Step 2: 应用到所有部署
        const deployments = await listDeployments(subscriptionId, resourceGroup, accountName);
        for (const dep of deployments) {
            await updateDeploymentRaiPolicy(
                subscriptionId, resourceGroup, accountName,
                dep.name, filterName, dep  // dep 包含原始部署数据
            );
        }
    }
}
```

### 应用筛选器到部署的关键逻辑

```javascript
async function updateDeploymentRaiPolicy(subscriptionId, resourceGroup, accountName, deploymentName, raiPolicyName, deploymentData) {
    const url = `${ARM_BASE}/subscriptions/${subscriptionId}/resourceGroups/${resourceGroup}/providers/Microsoft.CognitiveServices/accounts/${accountName}/deployments/${deploymentName}?api-version=2024-10-01`;
    
    // 保持原有部署配置，只更新 raiPolicyName
    const body = {
        ...deploymentData,
        properties: {
            ...deploymentData.properties,
            raiPolicyName: raiPolicyName  // ← 关键字段
        }
    };
    
    return await fetch(url, {
        method: 'PUT',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
    });
}
```

---

## 预设模式说明

| 预设 | 效果 | 适用场景 |
|------|------|----------|
| **默认 Medium** | 四类别 Medium 阈值 + 其他筛选器开启 | 标准生产环境 |
| **宽松 High** | 四类别 High 阈值（仅筛选最严重内容） | 内容创作类应用 |
| **严格 Low** | 四类别 Low 阈值（筛选所有级别） | 面向儿童/高度合规场景 |
| **仅批注** | enabled=true, blocking=false | 测试/监控（不阻止只标注） |
| **全部关闭** | enabled=false | 内部测试/已有外部审查系统 |

---

## 部署信息

| 项目 | 值 |
|------|-----|
| 网站 URL | http://content-filter-manager.azurewebsites.net |
| 订阅 | Microsoft Azure 赞助-pengcheng |
| 资源组 | alex-MAzure-RG |
| 区域 | East Asia |
| App Service Plan | B1 (Basic) |
| 运行时 | Node.js 22 LTS (Linux) |
| 认证方式 | MSAL.js + Azure CLI Public Client ID |

---

## 文件结构

```
Azure批量创建内容筛选器/
├── server.js           # Express 静态文件服务器
├── package.json        # Node.js 项目配置
├── public/
│   ├── index.html      # 主页面 (Azure Portal 风格 SPA)
│   ├── styles.css      # 样式表 (Azure Fluent Design)
│   ├── auth.js         # MSAL.js 认证模块
│   ├── azure-api.js    # Azure REST API 调用封装
│   └── app.js          # 主应用逻辑 (导航/向导/批量操作)
└── IMPLEMENTATION.md   # 本文档
```

---

## 使用步骤

1. **登录** → 点击"使用 Microsoft 账号登录"按钮
2. **选择订阅** → 在"订阅管理"页面勾选目标订阅
3. **加载资源** → 在"资源浏览"页面点击"加载资源"
4. **配置筛选器** → 在"批量创建"页面：
   - 设置名称和模式
   - 选择预设或手动配置每个类别的操作和阈值
   - 选择是否应用到模型部署
5. **执行** → 查看进度日志确认结果

---

## 关键 API 版本

- Azure Resource Manager: `2022-12-01`
- Cognitive Services (RAI Policies): `2024-10-01`

## 权限要求

用户账号需要对目标资源拥有以下权限：
- `Microsoft.CognitiveServices/accounts/raiPolicies/write` - 创建/更新筛选器
- `Microsoft.CognitiveServices/accounts/raiPolicies/delete` - 删除筛选器
- `Microsoft.CognitiveServices/accounts/deployments/write` - 更新部署（应用筛选器）
- `Microsoft.CognitiveServices/accounts/deployments/read` - 读取部署列表

通常 **Cognitive Services Contributor** 或 **Owner** 角色具备以上权限。
