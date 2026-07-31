// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export type IconPaletteCategoryId =
  | 'ai'
  | 'app-web'
  | 'compute'
  | 'containers'
  | 'data-analytics'
  | 'databases'
  | 'developer-devops'
  | 'hybrid-edge'
  | 'identity'
  | 'integration'
  | 'iot'
  | 'devices'
  | 'management'
  | 'monitoring'
  | 'networking'
  | 'security'
  | 'storage'
  | 'migration'
  | 'fabric'
  | 'general'
  | 'specialized';

export interface IconPaletteCategory {
  id: IconPaletteCategoryId;
  label: { en: string; ja: string };
  description: { en: string; ja: string };
  keywords: string[];
}

export const iconPaletteCategories: IconPaletteCategory[] = [
  {
    id: 'ai',
    label: { en: 'AI & Machine Learning', ja: 'AI・機械学習' },
    description: {
      en: 'Models, agents, cognitive APIs, search, and machine-learning services',
      ja: 'モデル、エージェント、Cognitive API、検索、機械学習サービス',
    },
    keywords: ['ai', 'ml', 'model', 'agent', 'foundry', 'openai', 'cognitive', 'vision', 'speech', 'language', '生成ai', '人工知能', '機械学習'],
  },
  {
    id: 'app-web',
    label: { en: 'Applications & Web', ja: 'アプリ・Web' },
    description: {
      en: 'Web apps, APIs, mobile backends, content delivery, and application hosting',
      ja: 'Webアプリ、API、モバイル バックエンド、コンテンツ配信、アプリ実行基盤',
    },
    keywords: ['app', 'web', 'api', 'mobile', 'cdn', 'frontend', 'backend', 'hosting', 'website', 'アプリ', 'ウェブ'],
  },
  {
    id: 'compute',
    label: { en: 'Compute', ja: 'コンピューティング' },
    description: {
      en: 'Virtual machines, desktops, batch, functions, and scalable compute',
      ja: '仮想マシン、デスクトップ、Batch、Functions、スケーラブルな計算基盤',
    },
    keywords: ['compute', 'vm', 'virtual machine', 'desktop', 'batch', 'function', 'hpc', '仮想マシン', '計算'],
  },
  {
    id: 'containers',
    label: { en: 'Containers & Kubernetes', ja: 'コンテナー・Kubernetes' },
    description: {
      en: 'AKS, container apps, registries, instances, and container operations',
      ja: 'AKS、Container Apps、Registry、Container Instances、コンテナー運用',
    },
    keywords: ['container', 'kubernetes', 'aks', 'acr', 'aci', 'docker', 'openshift', 'コンテナー', 'クーバネティス'],
  },
  {
    id: 'data-analytics',
    label: { en: 'Data & Analytics', ja: 'データ・分析' },
    description: {
      en: 'Ingestion, streaming, data engineering, analytics, and business intelligence',
      ja: 'データ取り込み、ストリーミング、データ エンジニアリング、分析、BI',
    },
    keywords: ['analytics', 'data', 'stream', 'event', 'synapse', 'factory', 'bi', 'etl', '分析', 'データ'],
  },
  {
    id: 'databases',
    label: { en: 'Databases & Cache', ja: 'データベース・キャッシュ' },
    description: {
      en: 'Relational, NoSQL, document, graph, and managed cache services',
      ja: 'リレーショナル、NoSQL、Document、Graph、マネージド キャッシュ',
    },
    keywords: ['database', 'sql', 'nosql', 'cosmos', 'documentdb', 'postgresql', 'mysql', 'redis', 'cache', 'db', 'データベース'],
  },
  {
    id: 'developer-devops',
    label: { en: 'Developer & DevOps', ja: '開発・DevOps' },
    description: {
      en: 'Developer environments, testing, source delivery, CI/CD, and API tooling',
      ja: '開発環境、テスト、ソース配布、CI/CD、API開発ツール',
    },
    keywords: ['developer', 'devops', 'test', 'lab', 'code', 'pipeline', 'deployment', 'ci', 'cd', '開発', 'テスト'],
  },
  {
    id: 'hybrid-edge',
    label: { en: 'Hybrid, Multicloud & Edge', ja: 'ハイブリッド・マルチクラウド・Edge' },
    description: {
      en: 'Azure Arc, Azure Local, edge hardware, operator, and disconnected services',
      ja: 'Azure Arc、Azure Local、Edge機器、通信事業者、切断環境向けサービス',
    },
    keywords: ['hybrid', 'multicloud', 'edge', 'arc', 'azure local', 'operator', 'disconnected', 'ハイブリッド', 'エッジ'],
  },
  {
    id: 'identity',
    label: { en: 'Identity & Access', ja: 'ID・アクセス' },
    description: {
      en: 'Microsoft Entra, users, applications, authentication, and authorization',
      ja: 'Microsoft Entra、ユーザー、アプリ、認証、認可',
    },
    keywords: ['identity', 'entra', 'active directory', 'user', 'group', 'authentication', 'authorization', 'rbac', '認証', 'id', 'アクセス'],
  },
  {
    id: 'integration',
    label: { en: 'Integration & Messaging', ja: '統合・メッセージング' },
    description: {
      en: 'API management, workflows, queues, topics, events, and service integration',
      ja: 'API Management、ワークフロー、Queue、Topic、Event、サービス統合',
    },
    keywords: ['integration', 'message', 'queue', 'topic', 'event grid', 'service bus', 'logic apps', 'api management', '統合', 'メッセージ'],
  },
  {
    id: 'iot',
    label: { en: 'IoT & Industrial', ja: 'IoT・産業' },
    description: {
      en: 'Devices, telemetry, digital twins, industrial systems, and real-time control',
      ja: 'デバイス、テレメトリ、Digital Twins、産業システム、リアルタイム制御',
    },
    keywords: ['iot', 'device', 'telemetry', 'digital twins', 'industrial', 'rtos', 'sensor', 'デバイス', '産業'],
  },
  {
    id: 'devices',
    label: { en: 'Devices & Endpoint Management', ja: 'デバイス・エンドポイント管理' },
    description: {
      en: 'Intune, enrollment, compliance, endpoint security, and device configuration',
      ja: 'Intune、登録、準拠性、エンドポイント セキュリティ、デバイス構成',
    },
    keywords: ['intune', 'endpoint', 'device management', 'enrollment', 'compliance', 'configuration', '端末', 'デバイス管理'],
  },
  {
    id: 'management',
    label: { en: 'Management & Governance', ja: '管理・ガバナンス' },
    description: {
      en: 'Cost, policy, resource organization, backup, compliance, and operations',
      ja: 'コスト、Policy、リソース整理、Backup、コンプライアンス、運用',
    },
    keywords: ['management', 'governance', 'policy', 'cost', 'billing', 'backup', 'quota', 'compliance', '運用', '管理', 'ガバナンス'],
  },
  {
    id: 'monitoring',
    label: { en: 'Monitoring & Observability', ja: '監視・可観測性' },
    description: {
      en: 'Metrics, logs, traces, dashboards, alerts, health, and troubleshooting',
      ja: 'メトリック、ログ、トレース、ダッシュボード、アラート、正常性、診断',
    },
    keywords: ['monitor', 'observability', 'log', 'metric', 'trace', 'alert', 'insights', 'grafana', 'prometheus', '監視', 'ログ'],
  },
  {
    id: 'networking',
    label: { en: 'Networking & Delivery', ja: 'ネットワーク・配信' },
    description: {
      en: 'Virtual networks, routing, gateways, load balancing, DNS, and private access',
      ja: 'Virtual Network、Routing、Gateway、Load Balancing、DNS、Private Access',
    },
    keywords: ['network', 'vnet', 'vpn', 'gateway', 'load balancer', 'dns', 'front door', 'private link', 'route', 'ネットワーク'],
  },
  {
    id: 'security',
    label: { en: 'Security', ja: 'セキュリティ' },
    description: {
      en: 'Threat protection, secrets, firewalls, posture, confidential computing, and SIEM',
      ja: '脅威保護、Secret、Firewall、セキュリティ態勢、Confidential Computing、SIEM',
    },
    keywords: ['security', 'defender', 'firewall', 'key vault', 'sentinel', 'hsm', 'confidential', 'waf', 'セキュリティ'],
  },
  {
    id: 'storage',
    label: { en: 'Storage', ja: 'ストレージ' },
    description: {
      en: 'Object, file, disk, queue, table, data lake, transfer, and storage management',
      ja: 'Object、File、Disk、Queue、Table、Data Lake、転送、ストレージ管理',
    },
    keywords: ['storage', 'blob', 'file', 'disk', 'queue', 'table', 'data lake', 'san', 'ストレージ'],
  },
  {
    id: 'migration',
    label: { en: 'Migration & Modernization', ja: '移行・モダナイズ' },
    description: {
      en: 'Assessment, migration, resource movement, modernization, and transfer services',
      ja: '評価、移行、リソース移動、モダナイズ、データ転送',
    },
    keywords: ['migration', 'migrate', 'modernize', 'resource mover', 'transfer', '移行', 'モダナイズ'],
  },
  {
    id: 'fabric',
    label: { en: 'Microsoft Fabric', ja: 'Microsoft Fabric' },
    description: {
      en: 'OneLake, lakehouse, warehouse, real-time intelligence, Power BI, and Fabric items',
      ja: 'OneLake、Lakehouse、Warehouse、Real-Time Intelligence、Power BI、Fabricアイテム',
    },
    keywords: ['fabric', 'onelake', 'lakehouse', 'warehouse', 'power bi', 'semantic model', 'eventhouse', 'ファブリック'],
  },
  {
    id: 'general',
    label: { en: 'General Azure Symbols', ja: 'Azure共通シンボル' },
    description: {
      en: 'Subscriptions, resource groups, regions, generic resources, and diagram symbols',
      ja: 'Subscription、Resource Group、Region、汎用リソース、作図用シンボル',
    },
    keywords: ['general', 'resource', 'subscription', 'region', 'template', 'symbol', 'generic', '共通', 'シンボル'],
  },
  {
    id: 'specialized',
    label: { en: 'Specialized & Industry', ja: '特殊・業界向け' },
    description: {
      en: 'Blockchain, mixed reality, healthcare, SAP, communications, and specialized services',
      ja: 'Blockchain、Mixed Reality、医療、SAP、通信、その他の専門サービス',
    },
    keywords: ['specialized', 'industry', 'blockchain', 'mixed reality', 'healthcare', 'sap', 'communication', '業界', '特殊'],
  },
];

const categoryById = new Map(iconPaletteCategories.map(category => [category.id, category]));

const directCategoryMap: Record<string, IconPaletteCategoryId> = {
  'ai + machine learning': 'ai',
  analytics: 'data-analytics',
  'app services': 'app-web',
  web: 'app-web',
  mobile: 'app-web',
  compute: 'compute',
  containers: 'containers',
  databases: 'databases',
  'developer tools': 'developer-devops',
  devops: 'developer-devops',
  'hybrid + multicloud': 'hybrid-edge',
  identity: 'identity',
  integration: 'integration',
  iot: 'iot',
  intune: 'devices',
  'management + governance': 'management',
  monitor: 'monitoring',
  networking: 'networking',
  security: 'security',
  storage: 'storage',
  migrate: 'migration',
  migration: 'migration',
  fabric: 'fabric',
  general: 'general',
  menu: 'general',
  blockchain: 'specialized',
  'mixed reality': 'specialized',
};

const semanticRules: Array<[IconPaletteCategoryId, RegExp]> = [
  ['ai', /\b(ai|foundry|machine learning|cognitive|openai|vision|speech|language|translator|agentic|agent service|planetary computer)\b/i],
  ['containers', /\b(aks|kubernetes|container|openshift|istio)\b/i],
  ['databases', /\b(sql|database|documentdb|postgresql|cassandra|redis|instance pool)\b/i],
  ['monitoring', /\b(monitor|log analytics|grafana|prometh|dashboard|health model|troubleshoot)\b/i],
  ['security', /\b(defender|security|hsm|attestation|enclave|confidential|compliance|resource guard|ssh key)\b/i],
  ['networking', /\b(network|vpn|vnet|gateway|load balancer|peering|route|private endpoint|ip prefix|mobile network|orbital)\b/i],
  ['storage', /\b(storage|disk|elastic san|data virtualization)\b/i],
  ['hybrid-edge', /\b(arc|azure local|edge|vmware|sap|disconnected|bare metal|modular data center)\b/i],
  ['iot', /\b(iot|sphere|rtos|industrial|vehicle|medtech)\b/i],
  ['identity', /\b(entra|identity|app registration|user subscription|multi tenancy)\b/i],
  ['integration', /\b(logic apps|pubsub|communication services|service group|data sharing|business process)\b/i],
  ['developer-devops', /\b(dev box|dev tunnel|deployment environment|app testing|test base|load testing|hpc workbench)\b/i],
  ['management', /\b(cost|quota|backup|capacity reservation|savings plan|resource mover|update management|sustainability|template spec|operation center)\b/i],
  ['compute', /\b(virtual desktop|virtual machine|vm app|cloud shell|compute|linux|scheduled action|stage map)\b/i],
  ['app-web', /\b(web app|web job|notification service)\b/i],
  ['migration', /\b(migrate|migration|mover|transfer)\b/i],
];

export function classifyIconPaletteCategory(
  sourceCategory: string,
  displayName: string,
  fileName: string,
): IconPaletteCategoryId {
  const direct = directCategoryMap[sourceCategory];
  if (direct) return direct;

  const searchable = `${displayName} ${fileName}`.replace(/[-_]+/g, ' ');
  for (const [category, pattern] of semanticRules) {
    if (pattern.test(searchable)) return category;
  }
  return 'specialized';
}

export function getIconPaletteCategory(id: IconPaletteCategoryId): IconPaletteCategory {
  return categoryById.get(id) ?? categoryById.get('specialized')!;
}
