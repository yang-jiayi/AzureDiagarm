// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { FABRIC_ICON_CATALOG } from './fabricIconCatalog';
import { MICROSOFT_PRODUCT_ICON_CATALOG } from './microsoftProductIconCatalog';

/**
 * Service to Icon Mapping
 * Maps Azure service names to their icon files and indicates pricing data availability
 * This helps AI-generated diagrams use the correct icons and validates pricing support
 */

export interface ServiceIconMapping {
  /** Display name of the service */
  displayName: string;
  /** Service name variations that AI might use */
  aliases: string[];
  /** Icon filename (without path or extension) */
  iconFile: string;
  /** Category/folder in icon library */
  category: 'ai + machine learning' | 'app services' | 'compute' | 'databases' | 'storage' | 'networking' | 'web' | 'analytics' | 'containers' | 'devops' | 'integration' | 'identity' | 'management + governance' | 'iot' | 'monitor' | 'security' | 'fabric' | 'power platform' | 'dynamics 365' | 'other';
  /** Whether we have real pricing data for this service */
  hasPricingData: boolean;
  /** Service name used in pricing files (if hasPricingData is true) */
  pricingServiceName?: string;
  /** Whether this is a usage-based service */
  isUsageBased?: boolean;
  /** Typical monthly cost range (for reference) */
  costRange?: string;
}

export interface ResolvedServiceIconMapping {
  serviceName: string;
  mapping: ServiceIconMapping;
}

const FABRIC_SERVICE_ICON_MAP = Object.fromEntries(
  FABRIC_ICON_CATALOG.filter(definition => definition.includeInServiceMap).map(definition => [
    definition.serviceName,
    {
      displayName: definition.displayName,
      aliases: definition.aliases,
      iconFile: definition.fileName,
      category: 'fabric',
      hasPricingData: definition.hasPricingData,
      pricingServiceName: definition.pricingServiceName,
      isUsageBased: definition.isUsageBased,
      costRange: definition.costRange,
    } satisfies ServiceIconMapping,
  ]),
) as Record<string, ServiceIconMapping>;

// The editorial exemplars (Reference / Blueprint few-shots) emit Microsoft
// Fabric item names in shorthand that the base icon catalog does not list as
// exact aliases (e.g. "OneLake Shortcuts", "Copy Jobs", "ML Model",
// "Semantic Model (Direct Lake)"). Without an exact/alias match the STRICT
// resolver used by ReferenceArchitectureCanvas and name normalization returns
// null, so those tiles render icon-less. Map each shorthand to the correct
// Fabric icon here. `aliases` references the shared catalog array, so we
// REASSIGN a fresh array (never push) to avoid mutating FABRIC_ICON_CATALOG.
const FABRIC_EXEMPLAR_ALIASES: Record<string, string[]> = {
  OneLake: ['OneLake Shortcuts', 'OneLake Shortcut', 'Shortcuts'],
  'Fabric Copy Job': ['Copy Jobs'],
  'Fabric Notebook': ['Spark Notebook', 'Spark Notebooks'],
  'Fabric Experiments': ['ML Experiment', 'ML Experiments'],
  'Fabric Model': ['ML Model', 'ML Models'],
  'Mirrored Database': ['Mirroring'],
  'Semantic Model': ['Semantic Model (Direct Lake)', 'Semantic Model (Direct Query)'],
  'Power BI Report': ['Interactive Report'],
  'Power BI Paginated Report': ['Paginated Report'],
  // Data Activator ships as part of the Real-Time Intelligence workload and has
  // no standalone icon, so route it to the RTI workload glyph.
  'Fabric Real-Time Intelligence': ['Activator', 'Data Activator', 'Fabric Activator', 'Reflex'],
  // "API for GraphQL" is a real Fabric item but ships no dedicated icon in the
  // catalog, so route it to the Fabric generic-item glyph rather than leave the
  // reference tile icon-less.
  'Fabric Generic Placeholder': ['API for GraphQL', 'GraphQL API', 'Fabric API for GraphQL'],
};

for (const [serviceName, extraAliases] of Object.entries(FABRIC_EXEMPLAR_ALIASES)) {
  const mapping = FABRIC_SERVICE_ICON_MAP[serviceName];
  if (!mapping) continue;
  mapping.aliases = [...new Set([...mapping.aliases, ...extraAliases])];
}

// Power Platform, Copilot Studio, and Dynamics 365 are licensed per user or per
// app rather than through Azure meters, so they never carry Azure pricing data.
const MICROSOFT_PRODUCT_SERVICE_ICON_MAP = Object.fromEntries(
  MICROSOFT_PRODUCT_ICON_CATALOG.filter(definition => definition.includeInServiceMap).map(
    definition => [
      definition.serviceName,
      {
        displayName: definition.displayName,
        aliases: definition.aliases,
        iconFile: definition.fileName,
        category: definition.category,
        hasPricingData: definition.hasPricingData,
        isUsageBased: definition.isUsageBased,
        costRange: definition.costRange,
      } satisfies ServiceIconMapping,
    ],
  ),
) as Record<string, ServiceIconMapping>;

/**
 * Comprehensive service-to-icon mapping with pricing data availability
 */
export const SERVICE_ICON_MAP: Record<string, ServiceIconMapping> = {
  // ========================================
  // AI & Machine Learning Services
  // ========================================
  'Azure OpenAI': {
    displayName: 'Azure OpenAI',
    aliases: ['OpenAI', 'Azure OpenAI Service', 'GPT', 'ChatGPT'],
    iconFile: 'azure-openai',
    category: 'ai + machine learning',
    hasPricingData: true,
    pricingServiceName: 'Azure OpenAI',
    isUsageBased: true,
    costRange: '$1-200/mo (token-based)'
  },

  // Microsoft Foundry — the agent/model platform formerly branded Azure AI
  // Foundry and, before that, Azure AI Studio. Listed immediately after Azure
  // OpenAI so its aliases win over any later entry. Reported missing by a user:
  // without an entry here it is absent from the KNOWN SERVICES list in the
  // generation prompt, so the model never emits it and any hand-placed node
  // renders without an icon.
  'Microsoft Foundry': {
    displayName: 'Microsoft Foundry',
    aliases: [
      'Foundry',
      'AI Foundry',
      'Azure Foundry',
      'Azure AI Foundry',
      'Microsoft AI Foundry',
      'Foundry Project',
      'Azure AI Foundry Project',
      'Foundry Agent Service',
      'Azure AI Foundry Agent Service',
      'Microsoft Foundry Agent Service',
      'Azure AI Studio',
      'AI Studio',
      'Azure AI Hub',
      'AI Hub'
    ],
    iconFile: '035746832-icon-service-AI-Foundry',
    category: 'ai + machine learning',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Usage-based (model, agent and tool consumption)'
  },
  
  'Foundry Tools': {
    displayName: 'Foundry Tools',
    aliases: ['Azure AI Services', 'Azure Cognitive Services', 'Cognitive Services', 'Cognitive Service'],
    iconFile: '10162-icon-service-Cognitive-Services',
    category: 'ai + machine learning',
    hasPricingData: true,
    pricingServiceName: 'Cognitive Services',
    isUsageBased: true,
    costRange: '$0-500/mo (varies by service)'
  },
  
  'Computer Vision': {
    displayName: 'Computer Vision',
    aliases: ['Vision', 'Azure Vision', 'Azure AI Vision', 'Image Analysis'],
    iconFile: 'computer-vision',
    category: 'ai + machine learning',
    hasPricingData: true,
    pricingServiceName: 'Vision',
    isUsageBased: true,
    costRange: '$150-1500/mo'
  },
  
  'Custom Vision': {
    displayName: 'Custom Vision',
    aliases: ['Azure Custom Vision', 'Custom Vision Service'],
    iconFile: 'custom-vision',
    category: 'ai + machine learning',
    hasPricingData: true,
    pricingServiceName: 'Custom Vision',
    isUsageBased: true,
    costRange: '$0-300/mo'
  },
  
  'Speech Services': {
    displayName: 'Speech Services',
    aliases: ['Speech', 'Azure Speech', 'Azure AI Speech', 'Speech-to-Text', 'Text-to-Speech'],
    iconFile: 'azure-speech',
    category: 'ai + machine learning',
    hasPricingData: true,
    pricingServiceName: 'Speech',
    isUsageBased: true,
    costRange: '$100-1000/mo'
  },
  
  'Translator': {
    displayName: 'Translator',
    aliases: ['Translator Text', 'Azure Translator', 'Azure AI Translator', 'Translation'],
    iconFile: 'translator',
    category: 'ai + machine learning',
    hasPricingData: true,
    pricingServiceName: 'Translator',
    isUsageBased: true,
    costRange: '$100-1000/mo'
  },
  
  'Language': {
    displayName: 'Language',
    aliases: ['Azure Language', 'Azure AI Language', 'Text Analytics', 'NLP'],
    iconFile: 'language',
    category: 'ai + machine learning',
    hasPricingData: true,
    pricingServiceName: 'Language',
    isUsageBased: true,
    costRange: '$25-250/mo'
  },
  
  'Azure AI Document Intelligence': {
    displayName: 'Azure AI Document Intelligence',
    aliases: ['Document Intelligence', 'Form Recognizer', 'Azure Form Recognizer', 'Azure Document Intelligence', 'Form Processing'],
    iconFile: '00819-icon-service-Form-Recognizers',
    category: 'ai + machine learning',
    hasPricingData: true,
    pricingServiceName: 'Document Intelligence',
    isUsageBased: true,
    costRange: '$0-500/mo'
  },
  
  'Azure Machine Learning': {
    displayName: 'Azure Machine Learning',
    aliases: [
      'Machine Learning', 
      'ML', 
      'AML', 
      'Azure ML',
      'AML Workspace',
      'Azure Machine Learning Workspace',
      'Machine Learning Workspace',
      'ML Workspace',
      'Azure ML Workspace',
      'Machine Learning Service',
      'Azure Machine Learning Service',
      'AzureML'
    ],
    iconFile: 'azure-machine-learning',
    category: 'ai + machine learning',
    hasPricingData: true,
    pricingServiceName: 'Azure Machine Learning',
    isUsageBased: true,
    costRange: '$0-5000/mo (varies greatly)'
  },
  
  // AML Sub-components (granular architecture support)
  // These allow architects to break down AML into logical components
  // with accurate cost attribution (endpoints/deployments are $0, compute has cost)
  
  'AML Online Endpoint': {
    displayName: 'AML Online Endpoint',
    aliases: ['Online Endpoint', 'AML Endpoint', 'Managed Online Endpoint', 'Real-time Endpoint'],
    iconFile: 'azure-machine-learning',
    category: 'ai + machine learning',
    hasPricingData: false, // No direct cost - routing construct only
    isUsageBased: false,
    costRange: '$0 (routing only)'
  },
  
  'AML Batch Endpoint': {
    displayName: 'AML Batch Endpoint',
    aliases: ['Batch Endpoint', 'AML Batch', 'Batch Inference Endpoint'],
    iconFile: 'azure-machine-learning',
    category: 'ai + machine learning',
    hasPricingData: false, // No direct cost - routing construct only
    isUsageBased: false,
    costRange: '$0 (routing only)'
  },
  
  'AML Deployment': {
    displayName: 'AML Deployment',
    aliases: ['Online Deployment', 'Batch Deployment', 'Model Deployment', 'Managed Deployment', 'Shared Deployment', 'Dedicated Deployment'],
    iconFile: 'azure-machine-learning',
    category: 'ai + machine learning',
    hasPricingData: false, // Configuration only - compute cost is separate
    isUsageBased: false,
    costRange: '$0 (config only)'
  },
  
  'AML Managed Compute': {
    displayName: 'AML Managed Compute',
    aliases: ['AML Compute', 'ML Compute', 'Managed Compute', 'AML Compute Instance', 'Compute Instance', 'AML Managed Compute (CPU/GPU)', 'Managed Compute (CPU/GPU)'],
    iconFile: 'virtual-machines',
    category: 'compute',
    hasPricingData: true,
    pricingServiceName: 'Virtual Machines',
    isUsageBased: false,
    costRange: '$50-2000/mo (per instance)'
  },
  
  'Batch Compute Pool': {
    displayName: 'Batch Compute Pool',
    aliases: ['Batch Pool', 'AML Batch Compute', 'Batch Compute', 'Dedicated Batch Compute', 'Batch Compute (auto-scale)'],
    iconFile: '10031-icon-service-Batch-Accounts',
    category: 'compute',
    hasPricingData: true,
    pricingServiceName: 'Batch',
    isUsageBased: true,
    costRange: '$0-2000/mo (scale-to-zero capable)'
  },
  
  'Azure AI Search': {
    displayName: 'Azure AI Search',
    aliases: ['Azure Cognitive Search', 'Cognitive Search', 'Azure Search', 'AI Search'],
    iconFile: '10044-icon-service-Cognitive-Search',
    category: 'ai + machine learning',
    hasPricingData: false,
    isUsageBased: false,
    costRange: '$75-2500/mo'
  },
  
  // ========================================
  // Compute Services
  // ========================================
  'Virtual Machines': {
    displayName: 'Virtual Machines',
    aliases: ['VM', 'VMs', 'Virtual Machine', 'Azure VM'],
    iconFile: 'virtual-machines',
    category: 'compute',
    hasPricingData: true,
    pricingServiceName: 'Virtual Machines',
    isUsageBased: false,
    costRange: '$13-17340/mo (per instance)'
  },
  
  'App Service': {
    displayName: 'App Service',
    aliases: ['Azure App Service', 'Web App', 'App Services'],
    iconFile: 'app-service',
    category: 'app services',
    hasPricingData: true,
    pricingServiceName: 'Azure App Service',
    isUsageBased: false,
    costRange: '$13-730/mo (per instance)'
  },
  
  'Functions': {
    displayName: 'Azure Functions',
    aliases: ['Function App', 'Function Apps', 'Functions', 'Serverless Functions'],
    iconFile: 'azure-functions',
    category: 'compute',
    hasPricingData: true,
    pricingServiceName: 'Functions',
    isUsageBased: true,
    costRange: '$0-160/mo (consumption-based)'
  },
  
  'Container Instances': {
    displayName: 'Container Instances',
    aliases: ['ACI', 'Azure Container Instances', 'Container Instance'],
    iconFile: 'container-instances',
    category: 'containers',
    hasPricingData: true,
    pricingServiceName: 'Container Instances',
    isUsageBased: true,
    costRange: '$0-500/mo (per-second billing)'
  },
  
  'Kubernetes Service': {
    displayName: 'Azure Kubernetes Service',
    aliases: ['AKS', 'Kubernetes', 'K8s', 'Kubernetes Services'],
    iconFile: 'azure-kubernetes-service',
    category: 'containers',
    hasPricingData: true,
    pricingServiceName: 'Azure Kubernetes Service',
    isUsageBased: false,
    costRange: '$73/mo + node costs'
  },
  
  'Container Registry': {
    displayName: 'Container Registry',
    aliases: ['ACR', 'Azure Container Registry', 'Container Registries'],
    iconFile: 'container-registry',
    category: 'containers',
    hasPricingData: true,
    pricingServiceName: 'Container Registry',
    isUsageBased: false,
    costRange: '$5-1000/mo'
  },
  
  // ========================================
  // Databases
  // ========================================
  'Azure Cosmos DB': {
    displayName: 'Azure Cosmos DB',
    aliases: ['Cosmos DB', 'CosmosDB', 'Cosmos'],
    iconFile: 'azure-cosmos-db',
    category: 'databases',
    hasPricingData: true,
    pricingServiceName: 'Azure Cosmos DB',
    isUsageBased: true,
    costRange: '$24-29185/mo'
  },
  
  'SQL Database': {
    displayName: 'SQL Database',
    aliases: ['Azure SQL', 'Azure SQL Database', 'SQL DB'],
    iconFile: 'sql-database',
    category: 'databases',
    hasPricingData: true,
    pricingServiceName: 'SQL Database',
    isUsageBased: false,
    costRange: '$5-43800/mo'
  },
  
  'PostgreSQL': {
    displayName: 'Azure Database for PostgreSQL',
    aliases: ['PostgreSQL', 'Postgres', 'Azure PostgreSQL', 'Azure Database for PostgreSQL', 'Azure Database for PostgreSQL Flexible Server', 'PostgreSQL Flexible Server', 'Azure PostgreSQL Flexible Server', 'PostgreSQL Server'],
    iconFile: '02827-icon-service-Azure-Database-PostgreSQL-Server-Group',
    category: 'databases',
    hasPricingData: true,
    pricingServiceName: 'Azure Database for PostgreSQL',
    isUsageBased: false,
    costRange: '$5-11240/mo'
  },
  
  'MySQL': {
    displayName: 'Azure Database for MySQL',
    aliases: ['MySQL', 'Azure MySQL', 'Azure Database for MySQL'],
    iconFile: 'azure-database-mysql',
    category: 'databases',
    hasPricingData: true,
    pricingServiceName: 'Azure Database for MySQL',
    isUsageBased: false,
    costRange: '$5-9800/mo'
  },
  
  'Redis Cache': {
    displayName: 'Azure Cache for Redis',
    aliases: ['Redis', 'Redis Cache', 'Cache'],
    iconFile: 'redis-cache',
    category: 'databases',
    hasPricingData: true,
    pricingServiceName: 'Redis Cache',
    isUsageBased: false,
    costRange: '$16-13600/mo'
  },
  
  // ========================================
  // Storage
  // ========================================
  'Storage Account': {
    displayName: 'Storage Account',
    aliases: ['Storage', 'Blob Storage', 'Azure Blob Storage', 'Azure Storage', 'Storage Accounts', 'Storage Accounts (Classic)'],
    iconFile: 'storage-account',
    category: 'storage',
    hasPricingData: true,
    pricingServiceName: 'Storage',
    isUsageBased: true,
    costRange: '$0.02-184/mo (per GB)'
  },
  
  // ========================================
  // Networking
  // ========================================
  'Application Gateway': {
    displayName: 'Application Gateway',
    aliases: ['App Gateway', 'Azure Application Gateway', 'Application Gateways'],
    iconFile: 'application-gateway',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'Application Gateway',
    isUsageBased: false,
    costRange: '$125-1200/mo'
  },
  
  'Azure Front Door': {
    displayName: 'Azure Front Door',
    aliases: ['Front Door', 'AFD', 'Azure Front Door Service'],
    iconFile: 'azure-front-door',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'Azure Front Door Service',
    isUsageBased: true,
    costRange: '$35-412/mo + traffic'
  },
  
  'CDN': {
    displayName: 'Content Delivery Network',
    aliases: ['Azure CDN', 'CDN', 'Content Delivery'],
    iconFile: '00056-icon-service-CDN-Profiles',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'Content Delivery Network',
    isUsageBased: true,
    costRange: '$0.081-0.20 per GB'
  },
  
  // ========================================
  // Analytics & Data
  // ========================================
  'Data Factory': {
    displayName: 'Azure Data Factory',
    aliases: ['Data Factory', 'ADF'],
    iconFile: 'data-factory',
    category: 'analytics',
    hasPricingData: true,
    pricingServiceName: 'Azure Data Factory',
    isUsageBased: true,
    costRange: '$0.50-2.50 per 1000 activities'
  },
  
  'Azure Synapse Analytics': {
    displayName: 'Azure Synapse Analytics',
    aliases: ['Synapse', 'Synapse Analytics', 'Azure Synapse'],
    iconFile: 'azure-synapse-analytics',
    category: 'analytics',
    hasPricingData: true,
    pricingServiceName: 'Azure Synapse Analytics',
    isUsageBased: true,
    costRange: '$5-8000/mo + compute'
  },
  
  'Stream Analytics': {
    displayName: 'Azure Stream Analytics',
    aliases: ['Stream Analytics', 'ASA', 'Azure Stream Analytics'],
    iconFile: 'stream-analytics',
    category: 'analytics',
    hasPricingData: true,
    pricingServiceName: 'Stream Analytics',
    isUsageBased: true,
    costRange: '$0.11 per streaming unit/hour'
  },
  
  'Event Hubs': {
    displayName: 'Event Hubs',
    aliases: ['Azure Event Hubs', 'Event Hub'],
    iconFile: 'event-hubs',
    category: 'analytics',
    hasPricingData: true,
    pricingServiceName: 'Event Hubs',
    isUsageBased: true,
    costRange: '$11-6849/mo'
  },
  
  // ========================================
  // Integration
  // ========================================
  'Service Bus': {
    displayName: 'Service Bus',
    aliases: ['Azure Service Bus', 'Message Queue'],
    iconFile: 'service-bus',
    category: 'integration',
    hasPricingData: true,
    pricingServiceName: 'Service Bus',
    isUsageBased: true,
    costRange: '$0-10/mo + messages'
  },
  
  'Logic Apps': {
    displayName: 'Logic Apps',
    aliases: ['Azure Logic Apps', 'Logic App'],
    iconFile: 'logic-apps',
    category: 'integration',
    hasPricingData: true,
    pricingServiceName: 'Logic Apps',
    isUsageBased: true,
    costRange: '$0-1000/mo (per action)'
  },
  
  // ========================================
  // Management & Security
  // ========================================
  'Key Vault': {
    displayName: 'Key Vault',
    aliases: ['Azure Key Vault', 'Secrets Management'],
    iconFile: 'key-vault',
    category: 'security',
    hasPricingData: true,
    pricingServiceName: 'Key Vault',
    isUsageBased: true,
    costRange: '$0.03 per 10K operations'
  },
  
  'Application Insights': {
    displayName: 'Application Insights',
    aliases: ['App Insights', 'Azure Application Insights', 'Monitoring'],
    iconFile: 'application-insights',
    category: 'monitor',
    hasPricingData: true,
    pricingServiceName: 'Application Insights',
    isUsageBased: true,
    costRange: '$2.30 per GB ingested'
  },
    'Log Analytics': {
    displayName: 'Log Analytics',
    aliases: ['Azure Log Analytics', 'LA', 'Log Analytics Workspace'],
    iconFile: 'log-analytics',
    category: 'monitor',
    hasPricingData: true,
    pricingServiceName: 'Log Analytics',
    isUsageBased: true,
    costRange: '$2.76 per GB ingested'
  },
  'API Management': {
    displayName: 'API Management',
    aliases: ['APIM', 'Azure API Management', 'API Gateway', 'API Management Services'],
    iconFile: 'api-management',
    category: 'integration',
    hasPricingData: true,
    pricingServiceName: 'API Management',
    isUsageBased: false,
    costRange: '$50-2800/mo'
  },
  
  // ========================================
  // Dashboard & Visualization Services
  // ========================================
  'Azure Managed Grafana': {
    displayName: 'Azure Managed Grafana',
    aliases: ['Managed Grafana', 'Grafana', 'Azure Grafana'],
    iconFile: '02905-icon-service-Azure-Managed-Grafana',
    category: 'other',
    hasPricingData: true,
    pricingServiceName: 'Azure Managed Grafana',
    isUsageBased: false,
    costRange: '$10-300/mo'
  },
  'Power BI Embedded': {
    displayName: 'Power BI Embedded',
    aliases: ['Power BI', 'PowerBI', 'Power BI Dashboard', 'PBI'],
    iconFile: '03332-icon-service-Power-BI-Embedded',
    category: 'analytics',
    hasPricingData: true,
    pricingServiceName: 'Power BI Embedded',
    isUsageBased: false,
    costRange: '$735-4,700/mo'
  },
  'Azure Dashboard': {
    displayName: 'Azure Dashboard',
    aliases: ['Azure Portal Dashboard', 'Dashboard', 'Azure Monitor Dashboard'],
    iconFile: '02488-icon-service-Azure-Monitor-Dashboard',
    category: 'other',
    hasPricingData: false,
    pricingServiceName: 'Azure Dashboard',
    isUsageBased: false,
    costRange: 'Free'
  },
  'Azure Workbooks': {
    displayName: 'Azure Workbooks',
    aliases: ['Workbooks', 'Monitor Workbooks', 'Azure Monitor Workbooks'],
    iconFile: '02189-icon-service-Azure-Workbooks',
    category: 'analytics',
    hasPricingData: false,
    pricingServiceName: 'Azure Workbooks',
    isUsageBased: false,
    costRange: 'Free (data costs apply)'
  },
  
  // ========================================
  // IoT Services
  // ========================================
  'IoT Hub': {
    displayName: 'Azure IoT Hub',
    aliases: ['Azure IoT Hub', 'IoT', 'IoT Hub'],
    iconFile: '10182-icon-service-IoT-Hub',
    category: 'iot',
    hasPricingData: true,
    pricingServiceName: 'IoT Hub',
    isUsageBased: true,
    costRange: '$0-5000/mo'
  },
  
  'IoT Central': {
    displayName: 'Azure IoT Central',
    aliases: ['Azure IoT Central', 'IoT Central'],
    iconFile: '10184-icon-service-IoT-Central-Applications',
    category: 'iot',
    hasPricingData: false,
    isUsageBased: true,
    costRange: '$0-250/mo'
  },
  
  'Digital Twins': {
    displayName: 'Azure Digital Twins',
    aliases: ['Azure Digital Twins', 'Digital Twin'],
    iconFile: '01030-icon-service-Digital-Twins',
    category: 'iot',
    hasPricingData: false,
    isUsageBased: true,
    costRange: '$0-1000/mo'
  },
  
  // ========================================
  // Container Services (additional)
  // ========================================
  'Container Apps': {
    displayName: 'Azure Container Apps',
    aliases: ['Azure Container Apps', 'Container App', 'ACA'],
    iconFile: '02989-icon-service-Container-Apps-Environments',
    category: 'containers',
    hasPricingData: true,
    pricingServiceName: 'Azure Container Apps',
    isUsageBased: true,
    costRange: '$0-500/mo (consumption-based)'
  },
  
  // ========================================
  // Networking (additional)
  // ========================================
  'Virtual Network': {
    displayName: 'Virtual Network',
    aliases: ['VNet', 'Azure Virtual Network', 'VNET'],
    iconFile: '10061-icon-service-Virtual-Networks',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'Virtual Network',
    isUsageBased: false,
    costRange: '$0-7.30/mo (peering)'
  },
  
  'Load Balancer': {
    displayName: 'Azure Load Balancer',
    aliases: ['Azure Load Balancer', 'Load Balancers', 'LB', 'ロードバランサー', '負荷分散'],
    iconFile: '10062-icon-service-Load-Balancers',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'Load Balancer',
    isUsageBased: false,
    costRange: '$18-730/mo'
  },
  
  'Azure Firewall': {
    displayName: 'Azure Firewall',
    aliases: ['Firewall'],
    iconFile: '10084-icon-service-Firewalls',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'Azure Firewall',
    isUsageBased: false,
    costRange: '$438-1095/mo'
  },
  
  'VPN Gateway': {
    displayName: 'VPN Gateway',
    aliases: ['Azure VPN Gateway', 'VPN', 'Virtual Network Gateway'],
    iconFile: '10063-icon-service-Virtual-Network-Gateways',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'VPN Gateway',
    isUsageBased: false,
    costRange: '$26-361/mo'
  },
  
  'ExpressRoute': {
    displayName: 'ExpressRoute',
    aliases: ['Azure ExpressRoute', 'Express Route'],
    iconFile: '10079-icon-service-ExpressRoute-Circuits',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'ExpressRoute',
    isUsageBased: false,
    costRange: '$55-580/mo'
  },
  
  'Traffic Manager': {
    displayName: 'Azure Traffic Manager',
    aliases: ['Azure Traffic Manager'],
    iconFile: '10065-icon-service-Traffic-Manager-Profiles',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'Traffic Manager',
    isUsageBased: true,
    costRange: '$0.54 per million queries'
  },
  
  // ========================================
  // Integration (additional)
  // ========================================
  'Event Grid': {
    displayName: 'Azure Event Grid',
    aliases: ['Azure Event Grid'],
    iconFile: '10206-icon-service-Event-Grid-Topics',
    category: 'integration',
    hasPricingData: true,
    pricingServiceName: 'Azure Event Grid',
    isUsageBased: true,
    costRange: '$0.30 per million operations'
  },
  
  'SignalR Service': {
    displayName: 'Azure SignalR Service',
    aliases: ['SignalR', 'Azure SignalR', 'Azure SignalR Service'],
    iconFile: '10052-icon-service-SignalR',
    category: 'web',
    hasPricingData: true,
    pricingServiceName: 'SignalR',
    isUsageBased: true,
    costRange: '$0-49/mo per unit'
  },
  
  'Notification Hubs': {
    displayName: 'Azure Notification Hubs',
    aliases: ['Notification Hub', 'Azure Notification Hubs', 'Push Notifications'],
    iconFile: '10045-icon-service-Notification-Hubs',
    category: 'iot',
    hasPricingData: true,
    pricingServiceName: 'Notification Hubs',
    isUsageBased: true,
    costRange: '$0-200/mo'
  },
  
  // ========================================
  // Web (additional)
  // ========================================
  'Static Web Apps': {
    displayName: 'Azure Static Web Apps',
    aliases: ['Static Web App', 'Azure Static Web Apps', 'Azure Static Web App', 'SWA'],
    iconFile: '01007-icon-service-Static-Apps',
    category: 'web',
    hasPricingData: false,
    isUsageBased: false,
    costRange: '$0-9/mo'
  },
  
  // ========================================
  // Management, Monitoring & Security (additional)
  // ========================================
  'Azure Monitor': {
    displayName: 'Azure Monitor',
    aliases: ['Monitor'],
    iconFile: '00001-icon-service-Monitor',
    category: 'monitor',
    hasPricingData: true,
    pricingServiceName: 'Azure Monitor',
    isUsageBased: true,
    costRange: '$2.30 per GB ingested'
  },
  
  'Microsoft Defender for Cloud': {
    displayName: 'Microsoft Defender for Cloud',
    aliases: ['Defender for Cloud', 'Azure Defender', 'Security Center'],
    iconFile: '10241-icon-service-Microsoft-Defender-for-Cloud',
    category: 'security',
    hasPricingData: true,
    pricingServiceName: 'Microsoft Defender for Cloud',
    isUsageBased: false,
    costRange: '$0-15/mo per server'
  },
  
  'Microsoft Entra ID': {
    displayName: 'Microsoft Entra ID',
    aliases: ['Entra ID', 'Azure AD', 'Azure Active Directory', 'Active Directory'],
    iconFile: 'microsoft-entra-id',
    category: 'identity',
    hasPricingData: false,
    isUsageBased: false,
    costRange: '$0-9/mo per user'
  },
  
  'Backup': {
    displayName: 'Azure Backup',
    aliases: ['Azure Backup', 'Recovery Services', 'Recovery Services Vault'],
    iconFile: '00017-icon-service-Recovery-Services-Vaults',
    category: 'management + governance',
    hasPricingData: true,
    pricingServiceName: 'Backup',
    isUsageBased: true,
    costRange: '$5-25/mo per instance'
  },
  
  'Network Watcher': {
    displayName: 'Network Watcher',
    aliases: ['Azure Network Watcher'],
    iconFile: '10066-icon-service-Network-Watcher',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'Network Watcher',
    isUsageBased: true,
    costRange: '$0-10/mo per resource'
  },
  
  // ========================================
  // Security (additional)
  // ========================================
  'Azure Bastion': {
    displayName: 'Azure Bastion',
    aliases: ['Bastion', 'Bastion Host'],
    iconFile: '02422-icon-service-Bastions',
    category: 'networking',
    hasPricingData: false,
    isUsageBased: false,
    costRange: '$138-876/mo'
  },
  
  'Azure DDoS Protection': {
    displayName: 'Azure DDoS Protection',
    aliases: ['DDoS Protection', 'DDoS', 'DDoS Protection Plan'],
    iconFile: '10072-icon-service-DDoS-Protection-Plans',
    category: 'networking',
    hasPricingData: false,
    isUsageBased: false,
    costRange: '$2944/mo'
  },
  
  'Azure Policy': {
    displayName: 'Azure Policy',
    aliases: ['Policy', 'Azure Policies', 'Governance Policy'],
    iconFile: '10316-icon-service-Policy',
    category: 'management + governance',
    hasPricingData: false,
    isUsageBased: false,
    costRange: '$0 (included)'
  },
  
  'Microsoft Sentinel': {
    displayName: 'Microsoft Sentinel',
    aliases: ['Sentinel', 'Azure Sentinel', 'SIEM', 'Azure SIEM'],
    iconFile: '10248-icon-service-Azure-Sentinel',
    category: 'security',
    hasPricingData: false,
    isUsageBased: true,
    costRange: '$2.46 per GB ingested'
  },
  
  'Web Application Firewall': {
    displayName: 'Web Application Firewall',
    aliases: ['WAF', 'Azure WAF', 'Web Application Firewall Policy'],
    iconFile: '10362-icon-service-Web-Application-Firewall-Policies(WAF)',
    category: 'networking',
    hasPricingData: false,
    isUsageBased: false,
    costRange: '$100-500/mo'
  },
  
  'Private Link': {
    displayName: 'Azure Private Link',
    aliases: ['Azure Private Link', 'Private Endpoint', 'Private Endpoints', 'Azure Private Endpoint'],
    iconFile: '00427-icon-service-Private-Link',
    category: 'networking',
    hasPricingData: false,
    isUsageBased: false,
    costRange: '$7.30/mo per endpoint'
  },
  
  'Azure DNS': {
    displayName: 'Azure DNS',
    aliases: ['DNS', 'DNS Zone', 'DNS Zones', 'Private DNS'],
    iconFile: '10064-icon-service-DNS-Zones',
    category: 'networking',
    hasPricingData: false,
    isUsageBased: true,
    costRange: '$0.50/mo per zone'
  },
  
  // ========================================
  // Healthcare
  // ========================================
  'Azure Health Data Services FHIR service': {
    displayName: 'Azure Health Data Services FHIR service',
    aliases: ['Azure API for FHIR', 'FHIR', 'FHIR Service', 'Azure Health Data Services FHIR', 'Azure FHIR service'],
    iconFile: '02658-icon-service-FHIR-Service',
    category: 'other',
    hasPricingData: false,
    isUsageBased: true,
    costRange: '$0-3000/mo'
  },
  
  // ========================================
  // Data Governance
  // ========================================
  'Microsoft Purview': {
    displayName: 'Microsoft Purview',
    aliases: ['Purview', 'Azure Purview', 'Data Governance', 'Data Catalog'],
    iconFile: '10314-icon-service-Azure-Purview-Accounts',
    category: 'management + governance',
    hasPricingData: true,
    pricingServiceName: 'Microsoft Purview',
    costRange: '$0.50-240/mo'
  },
  
  // ========================================
  // Storage (additional)
  // ========================================
  'Data Lake Storage': {
    displayName: 'Azure Data Lake Storage',
    aliases: ['Data Lake', 'Azure Data Lake', 'Data Lake Storage Gen2', 'ADLS', 'Azure Data Lake Storage Gen2'],
    iconFile: 'storage-account',
    category: 'storage',
    hasPricingData: false,
    isUsageBased: true,
    costRange: '$0.02-0.15 per GB/mo'
  },

  // ========================================
  // Canonical services with dedicated icons
  // ========================================
  'App Service Certificates': {
    displayName: 'App Service Certificates',
    aliases: ['App Service Certificate', 'Azure App Service Certificates'],
    iconFile: '00049-icon-service-App-Service-Certificates',
    category: 'app services',
    hasPricingData: false,
    isUsageBased: false,
    costRange: 'Certificate purchase and renewal'
  },
  'Virtual Machine Scale Sets': {
    displayName: 'Virtual Machine Scale Sets',
    aliases: ['VM Scale Sets', 'VMSS', 'Azure Virtual Machine Scale Sets'],
    iconFile: '10034-icon-service-VM-Scale-Sets',
    category: 'compute',
    hasPricingData: true,
    pricingServiceName: 'Virtual Machines',
    isUsageBased: false,
    costRange: 'Based on member VM instances'
  },
  'Azure Batch': {
    displayName: 'Azure Batch',
    aliases: ['Batch', 'Batch Accounts', 'Azure Batch Account'],
    iconFile: '10031-icon-service-Batch-Accounts',
    category: 'compute',
    hasPricingData: true,
    pricingServiceName: 'Virtual Machines',
    isUsageBased: true,
    costRange: 'Based on pool VM usage'
  },
  'Azure Bot Service': {
    displayName: 'Azure Bot Service',
    aliases: ['Bot Service', 'Bot Services', 'Azure AI Bot Service'],
    iconFile: '10165-icon-service-Bot-Services',
    category: 'ai + machine learning',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Channel and hosting dependent'
  },
  'Face API': {
    displayName: 'Face API',
    aliases: ['Azure Face API', 'Face APIs', 'Azure AI Face'],
    iconFile: '00794-icon-service-Face-APIs',
    category: 'ai + machine learning',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Transaction based'
  },
  'Azure AI Content Safety': {
    displayName: 'Azure AI Content Safety',
    aliases: ['Content Safety', 'Azure Content Safety'],
    iconFile: '03390-icon-service-Content-Safety',
    category: 'ai + machine learning',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Transaction based'
  },
  'Managed Identity': {
    displayName: 'Managed Identity',
    aliases: ['Managed Identities', 'Azure Managed Identity', 'Microsoft Entra Managed Identity'],
    iconFile: '10227-icon-service-Entra-Managed-Identities',
    category: 'identity',
    hasPricingData: false,
    isUsageBased: false,
    costRange: '$0'
  },
  'Application Gateway for Containers': {
    displayName: 'Application Gateway for Containers',
    aliases: ['AGC', 'Application Gateway Containers'],
    iconFile: '03328-icon-service-Application-Gateway-Containers',
    category: 'networking',
    hasPricingData: true,
    pricingServiceName: 'Application Gateway',
    isUsageBased: true,
    costRange: 'Gateway and capacity-unit based'
  },
  'Azure Files': {
    displayName: 'Azure Files',
    aliases: ['Azure File Shares', 'File Shares', 'Azure Fileshares'],
    iconFile: '10400-icon-service-Azure-Fileshares',
    category: 'storage',
    hasPricingData: true,
    pricingServiceName: 'Storage',
    isUsageBased: true,
    costRange: 'Capacity and transaction based'
  },
  'Queue Storage': {
    displayName: 'Queue Storage',
    aliases: ['Azure Queue Storage', 'Storage Queue'],
    iconFile: '10840-icon-service-Storage-Queue',
    category: 'storage',
    hasPricingData: true,
    pricingServiceName: 'Storage',
    isUsageBased: true,
    costRange: 'Capacity and transaction based'
  },
  'Table Storage': {
    displayName: 'Table Storage',
    aliases: ['Azure Table Storage', 'Storage Table'],
    iconFile: '10841-icon-service-Table',
    category: 'storage',
    hasPricingData: true,
    pricingServiceName: 'Storage',
    isUsageBased: true,
    costRange: 'Capacity and transaction based'
  },
  'Azure Health Data Services': {
    displayName: 'Azure Health Data Services',
    aliases: ['Health Data Services', 'FHIR', 'FHIR Service', 'Azure FHIR Service'],
    iconFile: '02658-icon-service-FHIR-Service',
    category: 'other',
    hasPricingData: true,
    pricingServiceName: 'Azure API for FHIR',
    isUsageBased: true,
    costRange: 'Service and storage based'
  },

  // ========================================
  // Widely-used services whose official icons already ship in the library.
  // No regional pricing dataset exists for these yet, so they declare
  // hasPricingData: false and carry a qualitative costRange.
  // ========================================
  'Azure Databricks': {
    displayName: 'Azure Databricks',
    aliases: ['Databricks', 'Databricks Workspace', 'Azure Databricks Workspace'],
    iconFile: '10787-icon-service-Azure-Databricks',
    category: 'analytics',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'DBU and underlying VM based'
  },
  'Azure Virtual Desktop': {
    displayName: 'Azure Virtual Desktop',
    aliases: ['Virtual Desktop', 'AVD', 'Azure Virtual Desktop Host Pool', 'Windows Virtual Desktop'],
    iconFile: '00327-icon-service-Azure-Virtual-Desktop',
    category: 'other',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Session host VM and licensing based'
  },
  'Azure Arc': {
    displayName: 'Azure Arc',
    aliases: ['Arc', 'Arc-enabled servers', 'Azure Arc enabled servers', 'Arc-enabled Kubernetes'],
    iconFile: '00756-icon-service-Azure-Arc',
    category: 'management + governance',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Free for core inventory; per-resource for add-ons'
  },
  'Azure Virtual WAN': {
    displayName: 'Azure Virtual WAN',
    aliases: ['Virtual WAN', 'vWAN', 'Virtual WANs', 'Azure vWAN'],
    iconFile: '10353-icon-service-Virtual-WANs',
    category: 'networking',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Hub, scale unit and data transfer based'
  },
  'NAT Gateway': {
    displayName: 'NAT Gateway',
    aliases: ['Azure NAT Gateway', 'Virtual Network NAT', 'VNet NAT'],
    iconFile: '10310-icon-service-NAT',
    category: 'networking',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Hourly plus data processed'
  },
  'Azure Communication Services': {
    displayName: 'Azure Communication Services',
    aliases: ['Communication Services', 'ACS', 'Azure Communication Service'],
    iconFile: '00968-icon-service-Azure-Communication-Services',
    category: 'other',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Per message, minute and channel'
  },
  'Azure Managed Redis': {
    displayName: 'Azure Managed Redis',
    aliases: ['Managed Redis', 'AMR', 'Azure Managed Redis Cache'],
    iconFile: '03675-icon-service-Azure-Managed-Redis',
    category: 'databases',
    hasPricingData: false,
    isUsageBased: false,
    costRange: 'Instance size and tier based'
  },
  'Azure Data Explorer': {
    displayName: 'Azure Data Explorer',
    aliases: ['Data Explorer', 'ADX', 'Kusto', 'Azure Data Explorer Cluster'],
    iconFile: '10145-icon-service-Azure-Data-Explorer-Clusters',
    category: 'analytics',
    hasPricingData: false,
    isUsageBased: false,
    costRange: 'Cluster size and storage based'
  },
  'Azure NetApp Files': {
    displayName: 'Azure NetApp Files',
    aliases: ['NetApp Files', 'ANF', 'Azure NetApp'],
    iconFile: '10096-icon-service-Azure-NetApp-Files',
    category: 'storage',
    hasPricingData: false,
    isUsageBased: false,
    costRange: 'Provisioned capacity and service level based'
  },
  'Azure Container Storage': {
    displayName: 'Azure Container Storage',
    aliases: ['Container Storage', 'ACStor'],
    iconFile: '03401-icon-service-Azure-Container-Storage',
    category: 'containers',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Backing storage and capacity based'
  },
  'Azure Load Testing': {
    displayName: 'Azure Load Testing',
    aliases: ['Load Testing', 'Azure Load Test'],
    iconFile: '02423-icon-service-Load-Testing',
    category: 'devops',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Virtual user hours based'
  },
  'Foundry Models': {
    displayName: 'Foundry Models',
    aliases: ['Azure AI Foundry Models', 'Model Catalog', 'Foundry Model Catalog'],
    iconFile: '038470614-icon-service-Foundry-Models',
    category: 'ai + machine learning',
    hasPricingData: true,
    pricingServiceName: 'Azure OpenAI',
    isUsageBased: true,
    costRange: 'Token and deployment based'
  },
  'Foundry Agent Service': {
    displayName: 'Foundry Agent Service',
    aliases: ['Azure AI Agent Service', 'AI Agent Service', 'Foundry Agents'],
    iconFile: '038470523-icon-service-Foundry-Agent-Service',
    category: 'ai + machine learning',
    hasPricingData: false,
    isUsageBased: true,
    costRange: 'Agent run and token based'
  },

  // Microsoft Fabric service components. Navigation and sample-only symbols
  // remain available in the icon palette but are not advertised to AI/MCP as services.
  ...FABRIC_SERVICE_ICON_MAP,

  // Microsoft Power Platform (including Copilot Studio and Agent 365) and
  // Dynamics 365 business applications.
  ...MICROSOFT_PRODUCT_SERVICE_ICON_MAP,
};

/**
 * Get icon mapping for a service by name (case-insensitive, checks aliases)
 */
export function resolveServiceIconMapping(serviceName: string): ResolvedServiceIconMapping | null {
  const normalizedName = serviceName.trim();
  
  // Direct match
  if (SERVICE_ICON_MAP[normalizedName]) {
    return {
      serviceName: normalizedName,
      mapping: SERVICE_ICON_MAP[normalizedName],
    };
  }
  
  // Search by canonical key, display name, or alias (case-insensitive).
  const lowerName = normalizedName.toLowerCase();
  for (const [canonicalName, mapping] of Object.entries(SERVICE_ICON_MAP)) {
    if (
      canonicalName.toLowerCase() === lowerName
      || mapping.displayName.toLowerCase() === lowerName
      || mapping.aliases.some(alias => alias.toLowerCase() === lowerName)
    ) {
      return { serviceName: canonicalName, mapping };
    }
  }
  
  return null;
}

export function getServiceIconMapping(serviceName: string): ServiceIconMapping | null {
  return resolveServiceIconMapping(serviceName)?.mapping ?? null;
}

/**
 * Whether a service is a Microsoft Fabric workload item whose cost is
 * included in the workspace's Fabric Capacity (it consumes Capacity Units
 * rather than billing separately). Used to show an "included in capacity"
 * indicator instead of a blank/zero cost. Excludes Fabric Capacity itself
 * and OneLake (which has its own storage billing).
 */
export function isCapacityConsumed(serviceName: string): boolean {
  const m = getServiceIconMapping(serviceName);
  return !!m
    && m.category === 'fabric'
    && !m.hasPricingData
    && /consumes Fabric capacity/i.test(m.costRange || '');
}

/**
 * Get all services with pricing data
 */
export function getServicesWithPricing(): ServiceIconMapping[] {
  return Object.values(SERVICE_ICON_MAP).filter(m => m.hasPricingData);
}

/**
 * Get all AI/ML services
 */
export function getAIServices(): ServiceIconMapping[] {
  return Object.values(SERVICE_ICON_MAP).filter(m => m.category === 'ai + machine learning');
}

/**
 * Get services by category
 */
export function getServicesByCategory(category: ServiceIconMapping['category']): ServiceIconMapping[] {
  return Object.values(SERVICE_ICON_MAP).filter(m => m.category === category);
}

/**
 * Validate if a service has pricing support
 */
export function hasRealPricingData(serviceName: string): boolean {
  const mapping = getServiceIconMapping(serviceName);
  return mapping?.hasPricingData || false;
}

/**
 * Get the correct icon filename for a service
 */
export function getIconFilename(serviceName: string): string | null {
  const mapping = getServiceIconMapping(serviceName);
  return mapping?.iconFile || null;
}
