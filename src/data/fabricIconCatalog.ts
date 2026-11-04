// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export type FabricIconKind =
  | 'platform'
  | 'capacity'
  | 'workload'
  | 'item'
  | 'state'
  | 'navigation'
  | 'sample'
  | 'legacy'
  | 'ui';

export interface FabricIconDefinition {
  serviceName: string;
  displayName: string;
  fileName: string;
  aliases: string[];
  group: string;
  kind: FabricIconKind;
  sourceAsset: string | null;
  sourceVersion: string;
  includeInServiceMap: boolean;
  consumesCapacity: boolean;
  hasPricingData: boolean;
  pricingServiceName?: string;
  isUsageBased?: boolean;
  costRange?: string;
}

type FabricIconOptions = Partial<
  Pick<
    FabricIconDefinition,
    | 'consumesCapacity'
    | 'hasPricingData'
    | 'pricingServiceName'
    | 'isUsageBased'
    | 'costRange'
    | 'sourceVersion'
    | 'includeInServiceMap'
  >
>;

export const FABRIC_ICON_PACKAGE_VERSION = '8.2.0';

function defineFabricIcon(
  serviceName: string,
  displayName: string,
  fileName: string,
  sourceAsset: string | null,
  group: string,
  kind: FabricIconKind,
  aliases: string[] = [],
  options: FabricIconOptions = {},
): FabricIconDefinition {
  const consumesCapacity = kind === 'workload' || kind === 'item' || kind === 'state';
  return {
    serviceName,
    displayName,
    fileName,
    aliases,
    group,
    kind,
    sourceAsset,
    sourceVersion: options.sourceVersion ?? (sourceAsset ? FABRIC_ICON_PACKAGE_VERSION : 'local'),
    includeInServiceMap: options.includeInServiceMap ?? !['navigation', 'sample', 'legacy', 'ui'].includes(kind),
    consumesCapacity: options.consumesCapacity ?? consumesCapacity,
    hasPricingData: options.hasPricingData ?? false,
    pricingServiceName: options.pricingServiceName,
    isUsageBased: options.isUsageBased,
    costRange: options.costRange ?? (consumesCapacity ? '$0 (consumes Fabric capacity)' : undefined),
  };
}

/**
 * Complete Microsoft Fabric architecture icon inventory.
 *
 * The 82 current architecture families use 8.2.0's 48px assets (Sample Workload
 * only provides 32px). The article's
 * historical designs use 6.1.0 / 1.1.37; its OneLake security UI glyphs are
 * hash-pinned separately. Legacy and UI symbols are drawing-only, never AI
 * services. The local Capacity symbol retains the existing billable F-SKU.
 */
export const FABRIC_ICON_CATALOG: FabricIconDefinition[] = [
  // Platform and capacity
  defineFabricIcon('Microsoft Fabric', 'Microsoft Fabric', 'microsoft-fabric', 'fabric_48_color.svg', 'Platform', 'platform', ['Fabric', 'MS Fabric', 'Fabric Platform'], { consumesCapacity: false, costRange: 'Platform (see Microsoft Fabric Capacity)' }),
  defineFabricIcon('Microsoft Fabric Capacity', 'Microsoft Fabric Capacity', 'fabric-capacity', null, 'Platform', 'capacity', ['Fabric Capacity', 'Fabric F SKU', 'Fabric F64', 'Fabric F2', 'Capacity Unit', 'F SKU'], { consumesCapacity: false, hasPricingData: true, pricingServiceName: 'Microsoft Fabric', isUsageBased: false, costRange: '$263-8,410/mo (F2-F64, PAYG)' }),
  defineFabricIcon('OneLake', 'OneLake', 'onelake', 'one_lake_48_color.svg', 'Platform', 'platform', ['One Lake', 'Fabric OneLake', 'OneLake Storage'], { consumesCapacity: false, hasPricingData: true, pricingServiceName: 'OneLake Storage', isUsageBased: true, costRange: 'Region-dependent storage pricing' }),

  // Workloads and experiences
  defineFabricIcon('Fabric Copilot', 'Copilot', 'fabric-workload-copilot', 'copilot_48_color.svg', 'Workloads', 'workload', ['Microsoft Fabric Copilot', 'Copilot in Fabric']),
  defineFabricIcon('Fabric Data Engineering', 'Data Engineering', 'fabric-workload-data-engineering', 'data_engineering_48_color.svg', 'Workloads', 'workload', ['Microsoft Fabric Data Engineering', 'Data Engineering (Fabric)']),
  defineFabricIcon('Fabric Data Factory Workload', 'Data Factory', 'fabric-workload-data-factory', 'data_factory_48_color.svg', 'Workloads', 'workload', ['Microsoft Fabric Data Factory', 'Data Factory Workload']),
  defineFabricIcon('Fabric Data Science', 'Data Science', 'fabric-workload-data-science', 'data_science_48_color.svg', 'Workloads', 'workload', ['Microsoft Fabric Data Science', 'Data Science (Fabric)']),
  defineFabricIcon('Fabric Data Warehouse Workload', 'Data Warehouse', 'fabric-workload-data-warehouse', 'data_warehouse_48_color.svg', 'Workloads', 'workload', ['Microsoft Fabric Data Warehouse', 'Data Warehouse Workload']),
  defineFabricIcon('Fabric Databases Workload', 'Databases', 'fabric-workload-databases', 'databases_48_color.svg', 'Workloads', 'workload', ['Microsoft Fabric Databases', 'Databases (Fabric)']),
  defineFabricIcon('Fabric Graph Intelligence', 'Graph Intelligence', 'fabric-workload-graph-intelligence', 'graph_intelligence_48_color.svg', 'Workloads', 'workload', ['Microsoft Fabric Graph Intelligence', 'Graph Intelligence (Fabric)']),
  defineFabricIcon('Fabric Industry Solutions', 'Industry Solutions', 'fabric-workload-industry-solutions', 'industry_solutions_48_color.svg', 'Workloads', 'workload', ['Microsoft Fabric Industry Solutions']),
  defineFabricIcon('Fabric Power BI Workload', 'Power BI', 'fabric-workload-power-bi', 'power_bi_48_color.svg', 'Workloads', 'workload', ['Power BI (Fabric)', 'Microsoft Fabric Power BI']),
  defineFabricIcon('Fabric Purview Workload', 'Purview', 'fabric-workload-purview', 'purview_48_color.svg', 'Workloads', 'workload', ['Purview (Fabric)', 'Microsoft Purview in Fabric']),
  defineFabricIcon('Fabric Real-Time Intelligence', 'Real-Time Intelligence', 'fabric-workload-real-time-intelligence', 'real_time_intelligence_48_color.svg', 'Workloads', 'workload', ['RTI', 'Real Time Intelligence', 'Microsoft Fabric Real-Time Intelligence']),
  defineFabricIcon('Fabric Sample Workload', 'Sample Workload', 'fabric-sample-workload', 'sample_workload_32_color.svg', 'Development and Samples', 'sample', ['Fabric Workload Sample', 'Sample Fabric Workload', 'サンプル ワークロード'], { consumesCapacity: false, includeInServiceMap: false, costRange: 'Development sample symbol' }),

  // Workspace, navigation, and developer sample symbols
  defineFabricIcon('Fabric Add Pipeline', 'Add Pipeline', 'fabric-navigation-add-pipeline', 'add_pipeline_48_non-item.svg', 'Workspace and Navigation', 'navigation', ['Create Pipeline', 'New Pipeline', 'パイプラインを追加'], { consumesCapacity: false, includeInServiceMap: false, costRange: 'Workspace and navigation symbol' }),
  defineFabricIcon('Fabric Eventhouse Navigation', 'Eventhouse (Navigation)', 'fabric-navigation-eventhouse', 'event_house_48_non-item.svg', 'Workspace and Navigation', 'navigation', ['Event House Navigation', 'Eventhouse Workspace', 'イベントハウス ナビゲーション'], { consumesCapacity: false, includeInServiceMap: false, costRange: 'Workspace and navigation symbol' }),
  defineFabricIcon('Fabric Folder', 'Folder', 'fabric-navigation-folder', 'folder_48_non-item.svg', 'Workspace and Navigation', 'navigation', ['Fabric Folder', 'Workspace Folder', 'フォルダー'], { consumesCapacity: false, includeInServiceMap: false, costRange: 'Workspace and navigation symbol' }),
  defineFabricIcon('Fabric Group Workspace', 'Group Workspace', 'fabric-navigation-group-workspace', 'group_workspace_48_non-item.svg', 'Workspace and Navigation', 'navigation', ['Workspace Group', 'Fabric Workspace Group', 'グループ ワークスペース'], { consumesCapacity: false, includeInServiceMap: false, costRange: 'Workspace and navigation symbol' }),
  defineFabricIcon('Fabric Import Notebook', 'Import Notebook', 'fabric-navigation-import-notebook', 'import_notebook_48_non-item.svg', 'Workspace and Navigation', 'navigation', ['Notebook Import', 'Import Fabric Notebook', 'ノートブックをインポート'], { consumesCapacity: false, includeInServiceMap: false, costRange: 'Workspace and navigation symbol' }),
  defineFabricIcon('Fabric My Workspace', 'My Workspace', 'fabric-navigation-my-workspace', 'my_workspace_48_non-item.svg', 'Workspace and Navigation', 'navigation', ['Personal Workspace', 'Fabric My Workspace', 'マイ ワークスペース'], { consumesCapacity: false, includeInServiceMap: false, costRange: 'Workspace and navigation symbol' }),
  defineFabricIcon('Fabric Sample Item', 'Sample Item', 'fabric-sample-item', 'sample_48_non-item.svg', 'Development and Samples', 'sample', ['Fabric Item Sample', 'Sample Fabric Item', 'サンプル項目'], { consumesCapacity: false, includeInServiceMap: false, costRange: 'Development sample symbol' }),

  // Existing item families
  defineFabricIcon('Fabric Data Agent', 'Fabric Data Agent', 'fabric-data-agent', 'data_agent_48_item.svg', 'AI and Data Science', 'item', ['Data Agent (Fabric)', 'Fabric AI Skill']),
  defineFabricIcon('Fabric Data Factory', 'Fabric Data Factory', 'fabric-data-factory', 'data_factory_48_item.svg', 'Data Factory', 'item', ['Data Factory (Fabric)', 'Fabric Data Factory Item']),
  defineFabricIcon('Fabric Data Pipeline', 'Fabric Data Pipeline', 'fabric-data-pipeline', 'pipeline_48_item.svg', 'Data Factory', 'item', ['Fabric Pipeline', 'Data Pipeline (Fabric)', 'Pipeline']),
  defineFabricIcon('Dataflow Gen2', 'Dataflow Gen2', 'fabric-dataflow-gen2', 'dataflow_gen2_48_item.svg', 'Data Factory', 'item', ['Fabric Dataflow Gen2', 'Dataflow Gen 2']),
  defineFabricIcon('Datamart', 'Datamart', 'fabric-datamart', 'datamart_48_item.svg', 'Data Warehouse', 'item', ['Fabric Datamart']),
  defineFabricIcon('Warehouse', 'Warehouse', 'fabric-warehouse', 'data_warehouse_48_item.svg', 'Data Warehouse', 'item', ['Fabric Warehouse', 'Data Warehouse (Fabric)', 'Warehouse (Gold)']),
  defineFabricIcon('Eventhouse', 'Eventhouse', 'fabric-eventhouse', 'event_house_48_item.svg', 'Real-Time Intelligence', 'item', ['Fabric Eventhouse', 'Event House']),
  defineFabricIcon('Eventstream', 'Eventstream', 'fabric-eventstream', 'eventstream_48_item.svg', 'Real-Time Intelligence', 'item', ['Fabric Eventstream', 'Event Stream']),
  defineFabricIcon('KQL Database', 'KQL Database', 'fabric-kql-database', 'kql_database_48_item.svg', 'Real-Time Intelligence', 'item', ['Fabric KQL Database', 'KQL DB', 'Kusto Database', 'Kusto Database (Fabric)', 'Eventhouse Database']),
  defineFabricIcon('Lakehouse', 'Lakehouse', 'fabric-lakehouse', 'lakehouse_48_item.svg', 'Data Engineering', 'item', ['Fabric Lakehouse', 'Lake House', 'Lakehouse (Bronze)', 'Lakehouse (Silver)', 'Lakehouse (Gold)']),
  defineFabricIcon('Mirrored Database', 'Mirrored Database', 'fabric-mirrored-database', 'mirrored_generic_database_48_item.svg', 'Databases', 'item', ['Fabric Mirroring', 'Database Mirroring (Fabric)', 'Mirrored Generic Database', 'Mirrored DB']),
  defineFabricIcon('Fabric Notebook', 'Fabric Notebook', 'fabric-notebook', 'notebook_48_item.svg', 'Data Engineering', 'item', ['Fabric Spark Notebook', 'Microsoft Fabric Notebook', 'Notebook']),
  defineFabricIcon('Power BI Report', 'Power BI Report', 'fabric-power-bi-report', 'report_48_item.svg', 'Power BI', 'item', ['Fabric Report', 'Report']),
  defineFabricIcon('Real-Time Dashboard', 'Real-Time Dashboard', 'fabric-real-time-dashboard', 'real_time_dashboard_48_item.svg', 'Real-Time Intelligence', 'item', ['Fabric Real-Time Dashboard', 'Real Time Dashboard']),
  defineFabricIcon('Semantic Model', 'Semantic Model', 'fabric-semantic-model', 'semantic_model_48_item.svg', 'Power BI', 'item', ['Power BI Semantic Model', 'Power BI Dataset', 'Direct Lake Semantic Model']),
  defineFabricIcon('Fabric Spark Job', 'Fabric Spark Job', 'fabric-spark-job', 'spark_job_direction_48_item.svg', 'Data Engineering', 'item', ['Fabric Spark', 'Spark Job Definition', 'Spark Job Direction']),
  defineFabricIcon('Fabric SQL Database', 'Fabric SQL Database', 'fabric-sql-database', 'sql_database_48_item.svg', 'Databases', 'item', ['SQL Database (Fabric)', 'Fabric SQL DB']),

  // Data Factory and integration items
  defineFabricIcon('Fabric Copy Job', 'Copy Job', 'fabric-item-copy-job', 'copy_job_48_item.svg', 'Data Factory', 'item', ['Copy Job (Fabric)']),
  defineFabricIcon('Fabric Custom Streaming Connector', 'Custom Streaming Connector', 'fabric-item-custom-streaming-connector', 'custom_streaming_connector_48_item.svg', 'Data Factory', 'item', ['Streaming Connector (Fabric)']),
  defineFabricIcon('Fabric Dataflow', 'Dataflow', 'fabric-item-dataflow', 'dataflow_48_item.svg', 'Data Factory', 'item', ['Dataflow (Fabric)', 'Power BI Dataflow']),
  defineFabricIcon('Fabric External Dataflow', 'External Dataflow', 'fabric-item-external-dataflow', 'external_dataflow_48_item.svg', 'Data Factory', 'item', ['External Dataflow (Fabric)']),
  defineFabricIcon('Fabric Function Set', 'Function Set', 'fabric-item-function-set', 'function_set_48_item.svg', 'Data Factory', 'item', ['Function Set (Fabric)']),
  defineFabricIcon('Fabric Streaming Dataflow', 'Streaming Dataflow', 'fabric-item-streaming-dataflow', 'streaming_dataflow_48_item.svg', 'Data Factory', 'item', ['Streaming Dataflow (Fabric)']),

  // Engineering, science, and data-management items
  defineFabricIcon('Fabric Environment', 'Environment', 'fabric-item-environment', 'environment_48_item.svg', 'Data Engineering', 'item', ['Environment (Fabric)']),
  defineFabricIcon('Fabric Experiments', 'Experiments', 'fabric-item-experiments', 'experiments_48_item.svg', 'Data Science', 'item', ['Machine Learning Experiments', 'Fabric Experiments']),
  defineFabricIcon('Fabric Exploration', 'Exploration', 'fabric-item-exploration', 'exploration_48_item.svg', 'Data Science', 'item', ['Exploration (Fabric)']),
  defineFabricIcon('Fabric External Datamart', 'External Datamart', 'fabric-item-external-datamart', 'external_datamart_48_item.svg', 'Data Warehouse', 'item', ['External Datamart (Fabric)']),
  defineFabricIcon('Fabric External Semantic Model', 'External Semantic Model', 'fabric-item-external-semantic-model', 'external_semantic_model_48_item.svg', 'Power BI', 'item', ['External Semantic Model (Fabric)']),
  defineFabricIcon('Fabric Generic Placeholder', 'Generic Placeholder', 'fabric-item-generic-placeholder', 'generic_placeholder_48_item.svg', 'Other Items', 'state', ['Fabric Placeholder']),
  defineFabricIcon('Fabric Mirrored Catalog', 'Mirrored Catalog', 'fabric-item-mirrored-catalog', 'mirrored_catalog_48_item.svg', 'Databases', 'item', ['Mirrored Catalog (Fabric)']),
  defineFabricIcon('Fabric Mirrored Generic Storage', 'Mirrored Generic Storage', 'fabric-item-mirrored-generic-storage', 'mirrored_generic_storage_48_item.svg', 'Databases', 'item', ['Mirrored Storage (Fabric)']),
  defineFabricIcon('Fabric Model', 'Model', 'fabric-item-model', 'model_48_item.svg', 'Data Science', 'item', ['Machine Learning Model (Fabric)']),
  defineFabricIcon('Fabric Runtime Lineage', 'Runtime Lineage', 'fabric-item-runtime-lineage', 'runtime_lineage_48_item.svg', 'Data Engineering', 'item', ['Runtime Lineage (Fabric)']),
  defineFabricIcon('Fabric Schema Model', 'Schema Model', 'fabric-item-schema-model', 'schema_model_48_item.svg', 'Data Engineering', 'item', ['Schema Model (Fabric)']),
  defineFabricIcon('Fabric User Data Function', 'User Data Function', 'fabric-item-user-data-function', 'user_data_function_48_item.svg', 'Data Engineering', 'item', ['User Data Functions', 'UDF (Fabric)']),
  defineFabricIcon('Fabric Variable Library', 'Variable Library', 'fabric-item-variable-library', 'variable_library_48_item.svg', 'Data Engineering', 'item', ['Variable Library (Fabric)']),

  // Real-Time Intelligence items
  defineFabricIcon('KQL Queryset', 'KQL Queryset', 'fabric-item-kql-queryset', 'kql_queryset_48_item.svg', 'Real-Time Intelligence', 'item', ['Kusto Queryset', 'KQL Query Set']),
  defineFabricIcon('KQL Script', 'KQL Script', 'fabric-item-kql-script', 'kql_script_48_item.svg', 'Real-Time Intelligence', 'item', ['Kusto Script']),
  defineFabricIcon('Fabric Operations Agent', 'Operations Agent', 'fabric-item-operations-agent', 'operations_agent_48_item.svg', 'Real-Time Intelligence', 'item', ['Operations Agent (Fabric)']),
  defineFabricIcon('Streaming Semantic Model', 'Streaming Semantic Model', 'fabric-item-streaming-semantic-model', 'streaming_semantic_model_48_item.svg', 'Real-Time Intelligence', 'item', ['Real-Time Semantic Model']),

  // Power BI and reporting items
  defineFabricIcon('Power BI App', 'Apps', 'fabric-item-apps', 'apps_48_item.svg', 'Power BI', 'item', ['Power BI Apps', 'Fabric Apps']),
  defineFabricIcon('Power BI Dashboard', 'Dashboard', 'fabric-item-dashboard', 'dashboard_48_item.svg', 'Power BI', 'item', ['Fabric Dashboard']),
  defineFabricIcon('Power BI Metric Sets', 'Metric Sets', 'fabric-item-metric-sets', 'metric_sets_48_item.svg', 'Power BI', 'item', ['Metric Set', 'Power BI Metrics']),
  defineFabricIcon('Power BI Mobile Report', 'Mobile Report', 'fabric-item-mobile-report', 'mobile_report_48_item.svg', 'Power BI', 'item', ['Mobile Report (Power BI)']),
  defineFabricIcon('No Access Semantic Model', 'No Access Semantic Model', 'fabric-item-no-access-semantic-model', 'no_access_semantic_model_48_item.svg', 'Power BI', 'state', ['Restricted Semantic Model']),
  defineFabricIcon('Power BI Paginated Report', 'Paginated Report', 'fabric-item-paginated-report', 'paginated_report_48_item.svg', 'Power BI', 'item', ['Fabric Paginated Report']),
  defineFabricIcon('Fabric Planning', 'Planning', 'fabric-item-planning', 'planning_48_item.svg', 'Power BI', 'item', ['Plan', 'Planning (Fabric)']),
  defineFabricIcon('Power BI RDL Report', 'RDL Report', 'fabric-item-rdl-report', 'rdl_report_48_item.svg', 'Power BI', 'item', ['Report Definition Language Report']),
  defineFabricIcon('Restricted Power BI Report', 'Restricted Report', 'fabric-item-restricted-report', 'restricted_report_48_item.svg', 'Power BI', 'state', ['Restricted Fabric Report']),
  defineFabricIcon('Restricted Scorecard', 'Restricted Scorecard', 'fabric-item-restricted-scorecard', 'restricted_scorecard_48_item.svg', 'Power BI', 'state', ['Restricted Power BI Scorecard']),
  defineFabricIcon('Power BI Scorecard', 'Scorecard', 'fabric-item-scorecard', 'scorecard_48_item.svg', 'Power BI', 'item', ['Fabric Scorecard']),
  defineFabricIcon('Shared Semantic Model', 'Shared Semantic Model', 'fabric-item-shared-semantic-model', 'shared_semantic_model_48_item.svg', 'Power BI', 'item', ['Shared Power BI Semantic Model']),

  // Industry, governance, and graph items
  defineFabricIcon('Fabric Cohort', 'Cohort', 'fabric-item-cohort', 'cohort_48_item.svg', 'Industry Solutions', 'item', ['Cohort (Fabric)']),
  defineFabricIcon('Fabric Data Quality', 'Data Quality', 'fabric-item-data-quality', 'data_quality_48_item.svg', 'Governance', 'item', ['Data Quality (Fabric)']),
  defineFabricIcon('Fabric Healthcare', 'Healthcare', 'fabric-item-healthcare', 'healthcare_48_item.svg', 'Industry Solutions', 'item', ['Healthcare (Fabric)']),
  defineFabricIcon('Fabric Links', 'Shortcuts', 'fabric-item-links', 'links_48_item.svg', 'OneLake', 'item', ['Links', 'Graph Links', 'Fabric Shortcuts', 'OneLake Shortcuts', 'OneLake Shortcut', 'ショートカット']),
  defineFabricIcon('Fabric Ontology', 'Ontology', 'fabric-item-ontology', 'ontology_48_item.svg', 'Graph Intelligence', 'item', ['Ontology (Fabric)']),
  defineFabricIcon('Fabric Retail', 'Retail', 'fabric-item-retail', 'retail_48_item.svg', 'Industry Solutions', 'item', ['Retail (Fabric)']),
  defineFabricIcon('Fabric Sustainability', 'Sustainability', 'fabric-item-sustainability', 'sustainability_48_item.svg', 'Industry Solutions', 'item', ['Sustainability (Fabric)']),
  defineFabricIcon('Fabric Graph Model', 'Graph Model', 'fabric-item-graph-model', 'graph_model_48.svg', 'Graph Intelligence', 'item', ['Graph Model (Fabric)', 'Graph Model Instance']),
  defineFabricIcon('Fabric Graph Queryset', 'Graph Queryset', 'fabric-item-graph-queryset', 'graph_queryset_48.svg', 'Graph Intelligence', 'item', ['Graph Query Set', 'Graph Model Instance Queryset']),

  // Drawing-only historical artwork, not current product services.
  defineFabricIcon('Fabric Legacy Graph Model Instance', 'Graph Model Instance (Legacy)', 'fabric-legacy-graph-model-instance', 'graph_model_instance_48_item.svg', 'Legacy / Deprecated', 'legacy', ['Historical Graph Model Instance', '旧 Graph Model Instance'], { sourceVersion: '6.1.0', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy Graph Model Instance Queryset', 'Graph Model Instance Queryset (Legacy)', 'fabric-legacy-graph-model-instance-queryset', 'graph_model_instance_queryset_48_item.svg', 'Legacy / Deprecated', 'legacy', ['Historical Graph Model Instance Queryset', '旧 Graph Model Instance Queryset'], { sourceVersion: '6.1.0', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy AI Skills', 'AI Skills (Deprecated)', 'fabric-legacy-ai-skills', 'ai_skills_64_item.svg', 'Legacy / Deprecated', 'legacy', ['Historical AI Skills', '旧 AI Skills'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy App Development', 'App Development (Deprecated)', 'fabric-legacy-app-development', 'app_development_48_color.svg', 'Legacy / Deprecated', 'legacy', ['Historical App Development'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy Function', 'Function (Deprecated)', 'fabric-legacy-function', 'function_64_item.svg', 'Legacy / Deprecated', 'legacy', ['Historical Fabric Function'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy OneRiver', 'OneRiver (Deprecated)', 'fabric-legacy-oneriver', 'oneriver_48_color.svg', 'Legacy / Deprecated', 'legacy', ['Historical OneRiver'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy Reflex Experience', 'Reflex Experience (Deprecated)', 'fabric-legacy-reflex-experience', 'reflex_48_color.svg', 'Legacy / Deprecated', 'legacy', ['Historical Reflex Experience'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy Reflex', 'Reflex Item (Deprecated)', 'fabric-legacy-reflex', 'reflex_64_item.svg', 'Legacy / Deprecated', 'legacy', ['Historical Reflex Item'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy Synapse', 'Synapse (Deprecated)', 'fabric-legacy-synapse', 'synapse_48_color .svg', 'Legacy / Deprecated', 'legacy', ['Historical Fabric Synapse'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy Synapse Data Engineering', 'Synapse Data Engineering (Deprecated)', 'fabric-legacy-synapse-data-engineering', 'synapse_data_engineering_28_color.svg', 'Legacy / Deprecated', 'legacy', ['Historical Synapse Data Engineering'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy Synapse Data Science', 'Synapse Data Science (Deprecated)', 'fabric-legacy-synapse-data-science', 'synapse_data_science_48_color.svg', 'Legacy / Deprecated', 'legacy', ['Historical Synapse Data Science'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy Synapse Data Warehouse', 'Synapse Data Warehouse (Deprecated)', 'fabric-legacy-synapse-data-warehouse', 'synapse_data_warehouse_48_color.svg', 'Legacy / Deprecated', 'legacy', ['Historical Synapse Data Warehouse'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy Synapse Real Time Intelligence', 'Synapse Real Time Intelligence (Deprecated)', 'fabric-legacy-synapse-real-time-intelligence', 'synapse_real_time_intelligence_48_color.svg', 'Legacy / Deprecated', 'legacy', ['Historical Synapse Real Time Intelligence'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric Legacy Variables', 'Variables (Deprecated)', 'fabric-legacy-variables', 'variables_64_item.svg', 'Legacy / Deprecated', 'legacy', ['Historical Fabric Variables'], { sourceVersion: '1.1.37', consumesCapacity: false, includeInServiceMap: false }),

  // Article-only UI glyphs: source URLs and immutable hashes live in the source lock.
  defineFabricIcon('Fabric UI Column Security Filled', 'Column Security (filled)', 'fabric-ui-column-security-filled', 'column_security_20_filled.svg', 'OneLake Security UI', 'ui', ['Column-level Security', 'CLS', '列レベル セキュリティ'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI Column Security Regular', 'Column Security (regular)', 'fabric-ui-column-security-regular', 'column_security_20_regular.svg', 'OneLake Security UI', 'ui', ['Column-level Security', 'CLS', '列レベル セキュリティ'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI Row Security Filled', 'Row Security (filled)', 'fabric-ui-row-security-filled', 'row_security_20_filled.svg', 'OneLake Security UI', 'ui', ['Row-level Security', 'RLS', '行レベル セキュリティ'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI Row Security Regular', 'Row Security (regular)', 'fabric-ui-row-security-regular', 'row_security_20_regular.svg', 'OneLake Security UI', 'ui', ['Row-level Security', 'RLS', '行レベル セキュリティ'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI OneLake Filled', 'OneLake (filled)', 'fabric-ui-onelake-filled', 'oneLake_20_filled.svg', 'OneLake Security UI', 'ui', ['OneLake UI'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI OneLake Regular', 'OneLake (regular)', 'fabric-ui-onelake-regular', 'oneLake_20_regular.svg', 'OneLake Security UI', 'ui', ['OneLake UI'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI OneLake Database', 'OneLake Database', 'fabric-ui-onelake-database', 'oneLake_database_20_regular.svg', 'OneLake Security UI', 'ui', ['OneLake Database UI'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI OneLake Schema Filled', 'OneLake Schema (filled)', 'fabric-ui-onelake-schema-filled', 'oneLake_schema_20_filled.svg', 'OneLake Security UI', 'ui', ['OneLake Schema UI'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI OneLake Schema Regular', 'OneLake Schema (regular)', 'fabric-ui-onelake-schema-regular', 'oneLake_schema_20_regular.svg', 'OneLake Security UI', 'ui', ['OneLake Schema UI'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI OneLake Table Filled', 'OneLake Table (filled)', 'fabric-ui-onelake-table-filled', 'oneLake_table_20_filled.svg', 'OneLake Security UI', 'ui', ['OneLake Table UI'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI OneLake Table Regular', 'OneLake Table (regular)', 'fabric-ui-onelake-table-regular', 'oneLake_table_20_regular.svg', 'OneLake Security UI', 'ui', ['OneLake Table UI'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI OneLake Table Link Filled', 'OneLake Table Link (filled)', 'fabric-ui-onelake-table-link-filled', 'oneLake_table_link_20_filled.svg', 'OneLake Security UI', 'ui', ['OneLake Table Link UI'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
  defineFabricIcon('Fabric UI OneLake Table Link Regular', 'OneLake Table Link (regular)', 'fabric-ui-onelake-table-link-regular', 'oneLake_table_link_20_regular.svg', 'OneLake Security UI', 'ui', ['OneLake Table Link UI'], { sourceVersion: 'article-2026', consumesCapacity: false, includeInServiceMap: false }),
];

const FABRIC_ICON_BY_FILE_NAME = new Map(
  FABRIC_ICON_CATALOG.map(definition => [definition.fileName, definition]),
);

export function getFabricIconByFileName(fileName: string): FabricIconDefinition | undefined {
  return FABRIC_ICON_BY_FILE_NAME.get(fileName);
}

export function isFabricUiIconPath(path: unknown): boolean {
  if (typeof path !== 'string') return false;
  const fileName = path.split('/').at(-1)?.replace(/\.svg(?:[?#].*)?$/, '');
  return fileName !== undefined && getFabricIconByFileName(fileName)?.kind === 'ui';
}
