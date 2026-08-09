// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { MarkerType, type Edge, type Node } from 'reactflow';
import type { LocalizedText } from '../i18n/localization';
import { getServiceIconMapping } from './serviceIconMapping';

export interface ArchitectureTemplate {
  id: string;
  name: LocalizedText;
  description: LocalizedText;
  tags: LocalizedText[];
  accent: string;
  diagram: {
    nodes: Node[];
    edges: Edge[];
    titleBlockData: {
      architectureName: string;
      author: string;
      version: string;
      date: string;
    };
    workflow: Array<{
      step: number;
      title: string;
      description: string;
      services: string[];
    }>;
  };
}

const node = (
  id: string,
  label: string,
  serviceName: string,
  category: string,
  x: number,
  y: number,
): Node => {
  const mapping = getServiceIconMapping(serviceName) ?? getServiceIconMapping(label);
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    data: {
      label,
      serviceName,
      category: mapping?.category ?? category,
      iconPath: mapping
        ? `/Azure_Public_Service_Icons/Icons/${mapping.category}/${mapping.iconFile}.svg`
        : undefined,
    },
  };
};

const edge = (id: string, source: string, target: string, label: string): Edge => ({
  id,
  source,
  target,
  sourceHandle: 'right',
  targetHandle: 'left',
  type: 'editableEdge',
  label,
  markerEnd: { type: MarkerType.ArrowClosed },
  data: {
    connectionType: label.toLowerCase().includes('event') ? 'async' : 'sync',
    direction: 'forward',
    baseFlowAnimated: true,
    flowAnimated: true,
    flowMode: 'directional',
    pathStyle: 'smooth',
  },
});

const titleBlockData = (architectureName: string) => ({
  architectureName,
  author: 'Azure Architect',
  version: '1.0',
  date: new Date().toLocaleDateString(),
});

export const ARCHITECTURE_TEMPLATES: ArchitectureTemplate[] = [
  {
    id: 'secure-web-app',
    name: { en: 'Secure web application', ja: 'セキュアな Web アプリ' },
    description: {
      en: 'A production web path with global ingress, managed compute, private data, secrets, and monitoring.',
      ja: 'グローバル入口、マネージド コンピュート、プライベート データ、シークレット、監視を備えた運用向け Web 構成です。',
    },
    tags: [
      { en: 'Web', ja: 'Web' },
      { en: 'Private data', ja: 'プライベート データ' },
      { en: 'Monitoring', ja: '監視' },
    ],
    accent: '#0f6cbd',
    diagram: {
      nodes: [
        node('web-front-door', 'Azure Front Door', 'Azure Front Door', 'networking', 40, 150),
        node('web-app-service', 'App Service', 'App Service', 'compute', 330, 150),
        node('web-sql', 'Azure SQL Database', 'Azure SQL Database', 'databases', 650, 80),
        node('web-key-vault', 'Key Vault', 'Key Vault', 'security', 650, 235),
        node('web-monitor', 'Azure Monitor', 'Azure Monitor', 'management + governance', 940, 150),
      ],
      edges: [
        edge('web-e1', 'web-front-door', 'web-app-service', 'HTTPS'),
        edge('web-e2', 'web-app-service', 'web-sql', 'Private Link'),
        edge('web-e3', 'web-app-service', 'web-key-vault', 'Managed identity'),
        edge('web-e4', 'web-app-service', 'web-monitor', 'Telemetry'),
      ],
      titleBlockData: titleBlockData('Secure Web Application'),
      workflow: [
        {
          step: 1,
          title: 'Request',
          description: 'Users enter through Azure Front Door.',
          services: ['web-front-door'],
        },
        {
          step: 2,
          title: 'Process',
          description: 'App Service handles application requests.',
          services: ['web-app-service'],
        },
        {
          step: 3,
          title: 'Persist',
          description: 'Data and secrets stay on managed private services.',
          services: ['web-sql', 'web-key-vault'],
        },
      ],
    },
  },
  {
    id: 'event-driven',
    name: { en: 'Event-driven integration', ja: 'イベント駆動連携' },
    description: {
      en: 'Decoupled APIs and workers using managed messaging, serverless processing, and operational telemetry.',
      ja: 'マネージド メッセージング、サーバーレス処理、運用テレメトリを使用する疎結合 API とワーカーです。',
    },
    tags: [
      { en: 'Integration', ja: '連携' },
      { en: 'Serverless', ja: 'サーバーレス' },
      { en: 'Async', ja: '非同期' },
    ],
    accent: '#7c3aed',
    diagram: {
      nodes: [
        node('event-apim', 'API Management', 'API Management', 'integration', 40, 150),
        node('event-bus', 'Service Bus', 'Service Bus', 'integration', 330, 150),
        node('event-functions', 'Azure Functions', 'Azure Functions', 'compute', 620, 80),
        node('event-cosmos', 'Azure Cosmos DB', 'Azure Cosmos DB', 'databases', 910, 80),
        node('event-insights', 'Application Insights', 'Application Insights', 'management + governance', 620, 235),
      ],
      edges: [
        edge('event-e1', 'event-apim', 'event-bus', 'Event'),
        edge('event-e2', 'event-bus', 'event-functions', 'Event trigger'),
        edge('event-e3', 'event-functions', 'event-cosmos', 'Write'),
        edge('event-e4', 'event-functions', 'event-insights', 'Telemetry'),
      ],
      titleBlockData: titleBlockData('Event-Driven Integration'),
      workflow: [
        {
          step: 1,
          title: 'Accept',
          description: 'API Management validates incoming requests.',
          services: ['event-apim'],
        },
        {
          step: 2,
          title: 'Queue',
          description: 'Service Bus decouples producers and consumers.',
          services: ['event-bus'],
        },
        {
          step: 3,
          title: 'Process',
          description: 'Functions process events and persist outcomes.',
          services: ['event-functions', 'event-cosmos'],
        },
      ],
    },
  },
  {
    id: 'analytics-platform',
    name: { en: 'Modern analytics platform', ja: 'モダン分析プラットフォーム' },
    description: {
      en: 'A governed ingest-to-insight path for batch data engineering, lake storage, analytics, and reporting.',
      ja: 'バッチ データ エンジニアリング、レイク保存、分析、レポートのためのガバナンス対応取り込み経路です。',
    },
    tags: [
      { en: 'Data', ja: 'データ' },
      { en: 'Analytics', ja: '分析' },
      { en: 'Lakehouse', ja: 'レイクハウス' },
    ],
    accent: '#047857',
    diagram: {
      nodes: [
        node('data-factory', 'Data Factory', 'Data Factory', 'analytics', 40, 150),
        node('data-lake', 'Data Lake Storage', 'Storage Accounts', 'storage', 310, 150),
        node('data-databricks', 'Azure Databricks', 'Azure Databricks', 'analytics', 590, 70),
        node('data-synapse', 'Synapse Analytics', 'Azure Synapse Analytics', 'analytics', 590, 235),
        node('data-powerbi', 'Power BI', 'Power BI Embedded', 'analytics', 900, 150),
      ],
      edges: [
        edge('data-e1', 'data-factory', 'data-lake', 'Ingest'),
        edge('data-e2', 'data-lake', 'data-databricks', 'Transform'),
        edge('data-e3', 'data-lake', 'data-synapse', 'Query'),
        edge('data-e4', 'data-synapse', 'data-powerbi', 'Semantic model'),
      ],
      titleBlockData: titleBlockData('Modern Analytics Platform'),
      workflow: [
        {
          step: 1,
          title: 'Ingest',
          description: 'Data Factory orchestrates source ingestion.',
          services: ['data-factory', 'data-lake'],
        },
        {
          step: 2,
          title: 'Curate',
          description: 'Lake storage and compute prepare governed data.',
          services: ['data-lake', 'data-databricks'],
        },
        {
          step: 3,
          title: 'Serve',
          description: 'Synapse and Power BI deliver analytical insight.',
          services: ['data-synapse', 'data-powerbi'],
        },
      ],
    },
  },
  {
    id: 'aks-platform',
    name: { en: 'AKS application platform', ja: 'AKS アプリ基盤' },
    description: {
      en: 'A container platform with protected ingress, private images and secrets, and centralized observability.',
      ja: '保護された入口、プライベート イメージとシークレット、集中監視を備えたコンテナー基盤です。',
    },
    tags: [
      { en: 'Containers', ja: 'コンテナー' },
      { en: 'AKS', ja: 'AKS' },
      { en: 'Platform', ja: 'プラットフォーム' },
    ],
    accent: '#0891b2',
    diagram: {
      nodes: [
        node('aks-front-door', 'Azure Front Door', 'Azure Front Door', 'networking', 40, 150),
        node('aks-app-gateway', 'Application Gateway', 'Application Gateway', 'networking', 300, 150),
        node('aks-cluster', 'Azure Kubernetes Service', 'Azure Kubernetes Service', 'containers', 590, 150),
        node('aks-acr', 'Container Registry', 'Container Registries', 'containers', 880, 60),
        node('aks-key-vault', 'Key Vault', 'Key Vault', 'security', 880, 160),
        node('aks-monitor', 'Azure Monitor', 'Azure Monitor', 'management + governance', 880, 260),
      ],
      edges: [
        edge('aks-e1', 'aks-front-door', 'aks-app-gateway', 'HTTPS'),
        edge('aks-e2', 'aks-app-gateway', 'aks-cluster', 'Private ingress'),
        edge('aks-e3', 'aks-acr', 'aks-cluster', 'Pull image'),
        edge('aks-e4', 'aks-cluster', 'aks-key-vault', 'Managed identity'),
        edge('aks-e5', 'aks-cluster', 'aks-monitor', 'Telemetry'),
      ],
      titleBlockData: titleBlockData('AKS Application Platform'),
      workflow: [
        {
          step: 1,
          title: 'Route',
          description: 'Global and regional ingress protect workloads.',
          services: ['aks-front-door', 'aks-app-gateway'],
        },
        {
          step: 2,
          title: 'Run',
          description: 'AKS hosts containerized application services.',
          services: ['aks-cluster'],
        },
        {
          step: 3,
          title: 'Operate',
          description: 'Registry, secrets, and monitoring support the platform.',
          services: ['aks-acr', 'aks-key-vault', 'aks-monitor'],
        },
      ],
    },
  },
];
