// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Official Microsoft Power Platform, Dynamics 365, and Microsoft 365 icon
 * inventory.
 *
 * Source packages (pinned in `scripts/sync-microsoft-icons.mjs`):
 * - https://learn.microsoft.com/en-us/power-platform/guidance/icons
 * - https://learn.microsoft.com/en-us/dynamics365/get-started/icons
 * - https://learn.microsoft.com/en-us/previous-versions/microsoft-365/solutions/architecture-icons-templates
 *
 * Microsoft permits these icons in architectural diagrams, training materials,
 * and documentation. They must not be cropped, flipped, rotated, distorted, or
 * used to represent a non-Microsoft product.
 *
 * Power Platform and Dynamics 365 ship product logos, so each entry there names
 * one product. The Microsoft 365 package ships concept symbols instead — the
 * shapes Microsoft's own Microsoft 365 architecture diagrams are drawn from —
 * so those entries are `kind: 'symbol'` and stay out of the service map: they
 * describe a role in a drawing, not a billable service. Their aliases still
 * carry the workload names so searching "Teams" or "Purview" in the palette
 * surfaces the symbol Microsoft uses for it.
 */

export type MicrosoftProductIconFamily = 'power-platform' | 'dynamics-365' | 'microsoft-365';

export type MicrosoftProductIconCategory = 'power platform' | 'dynamics 365' | 'microsoft 365';

export type MicrosoftProductIconKind = 'platform' | 'product' | 'app' | 'symbol';

export interface MicrosoftProductIconDefinition {
  serviceName: string;
  displayName: string;
  /** Destination file name inside the icon library, without the extension. */
  fileName: string;
  /** Path of the asset inside the official package. */
  sourceAsset: string;
  family: MicrosoftProductIconFamily;
  /** Icon library folder, which is also the palette source category. */
  category: MicrosoftProductIconCategory;
  group: string;
  kind: MicrosoftProductIconKind;
  aliases: string[];
  includeInServiceMap: boolean;
  hasPricingData: boolean;
  isUsageBased?: boolean;
  costRange?: string;
}

/**
 * Revision of this catalog, mirrored into the manifest. Each package records
 * its own upstream release date under `packages[]`;
 * `scripts/sync-microsoft-icons.mjs` reads this value to detect drift between
 * the catalog and the generated manifest.
 */
export const MICROSOFT_PRODUCT_ICON_PACKAGE_VERSION = '2025-12';

const FAMILY_CATEGORY: Record<MicrosoftProductIconFamily, MicrosoftProductIconCategory> = {
  'power-platform': 'power platform',
  'dynamics-365': 'dynamics 365',
  'microsoft-365': 'microsoft 365',
};

const FAMILY_COST_RANGE: Record<MicrosoftProductIconFamily, string> = {
  'power-platform': 'Per-user or per-app subscription (not in Azure pricing)',
  'dynamics-365': 'Per-user subscription (not in Azure pricing)',
  'microsoft-365': 'Included with a Microsoft 365 subscription (not in Azure pricing)',
};

function defineMicrosoftProductIcon(
  serviceName: string,
  displayName: string,
  fileName: string,
  sourceAsset: string,
  family: MicrosoftProductIconFamily,
  group: string,
  kind: MicrosoftProductIconKind,
  aliases: string[] = [],
): MicrosoftProductIconDefinition {
  return {
    serviceName,
    displayName,
    fileName,
    sourceAsset,
    family,
    category: FAMILY_CATEGORY[family],
    group,
    kind,
    aliases,
    // A concept symbol names a role in a drawing, not a service, so it must not
    // become a service the AI can emit or the pricing engine can look up.
    includeInServiceMap: kind !== 'symbol',
    hasPricingData: false,
    isUsageBased: false,
    costRange: FAMILY_COST_RANGE[family],
  };
}

export const MICROSOFT_PRODUCT_ICON_CATALOG: MicrosoftProductIconDefinition[] = [
  // ========================================
  // Microsoft Power Platform
  // ========================================
  defineMicrosoftProductIcon('Microsoft Power Platform', 'Microsoft Power Platform', 'power-platform', 'Power Platform/PowerPlatform_scalable.svg', 'power-platform', 'Platform', 'platform', ['Power Platform', 'MS Power Platform', 'Microsoft Power Platform (Suite)', 'パワープラットフォーム']),
  defineMicrosoftProductIcon('Microsoft Power Apps', 'Power Apps', 'power-apps', 'Power Platform/PowerApps_scalable.svg', 'power-platform', 'Applications', 'product', ['Power Apps', 'PowerApps', 'Power App', 'Canvas App', 'Canvas Apps', 'Model-driven App', 'Model-driven Apps', 'パワーアップス']),
  defineMicrosoftProductIcon('Microsoft Power Automate', 'Power Automate', 'power-automate', 'Power Platform/PowerAutomate_scalable.svg', 'power-platform', 'Automation', 'product', ['Power Automate', 'PowerAutomate', 'Microsoft Flow', 'Power Automate Cloud Flow', 'Power Automate Desktop', 'Power Automate for desktop', 'パワーオートメイト']),
  defineMicrosoftProductIcon('Microsoft Power Pages', 'Power Pages', 'power-pages', 'Power Platform/PowerPages_scalable.svg', 'power-platform', 'Applications', 'product', ['Power Pages', 'PowerPages', 'Power Apps Portals', 'Power Apps Portal', 'パワーページズ']),
  defineMicrosoftProductIcon('Microsoft Dataverse', 'Dataverse', 'dataverse', 'Power Platform/Dataverse_scalable.svg', 'power-platform', 'Data', 'product', ['Dataverse', 'Microsoft Dataverse for Teams', 'Common Data Service', 'Dataverse Table', 'データバース']),
  defineMicrosoftProductIcon('AI Builder', 'AI Builder', 'ai-builder', 'Power Platform/AIBuilder_scalable.svg', 'power-platform', 'AI', 'product', ['AI Builder', 'AIBuilder', 'Power Platform AI Builder', 'AI ビルダー']),
  defineMicrosoftProductIcon('Microsoft Copilot Studio', 'Copilot Studio', 'copilot-studio', 'CopilotStudio_scalable.svg', 'power-platform', 'AI', 'product', ['Copilot Studio', 'CopilotStudio', 'Power Virtual Agents', 'Power Virtual Agent', 'Copilot Studio Agent', 'Declarative Agent', 'コパイロットスタジオ']),
  defineMicrosoftProductIcon('Microsoft Agent 365', 'Agent 365', 'agent-365', 'Agent365_scalable.svg', 'power-platform', 'AI', 'product', ['Agent 365', 'Agent365', 'Microsoft Agent365', 'エージェント365']),

  // ========================================
  // Microsoft Dynamics 365
  // ========================================
  defineMicrosoftProductIcon('Dynamics 365', 'Dynamics 365', 'dynamics-365', 'Dynamics 365 Product Family Icons/Dynamics365_scalable.svg', 'dynamics-365', 'Platform', 'platform', ['Microsoft Dynamics 365', 'D365', 'Dynamics', 'ダイナミクス365']),
  defineMicrosoftProductIcon('Dynamics 365 Sales', 'Dynamics 365 Sales', 'dynamics-365-sales', 'Dynamics 365 App Icons/Sales_scalable.svg', 'dynamics-365', 'Customer Engagement', 'app', ['D365 Sales', 'Dynamics Sales', 'Sales (Dynamics 365)', 'Dynamics CRM Sales']),
  defineMicrosoftProductIcon('Dynamics 365 Sales Insights', 'Dynamics 365 Sales Insights', 'dynamics-365-sales-insights', 'Dynamics 365 App Icons/SalesInsights_scalable.svg', 'dynamics-365', 'Customer Engagement', 'app', ['D365 Sales Insights', 'Sales Insights', 'Sales Insights (Dynamics 365)']),
  defineMicrosoftProductIcon('Dynamics 365 Customer Service', 'Dynamics 365 Customer Service', 'dynamics-365-customer-service', 'Dynamics 365 App Icons/CustomerServices_scalable.svg', 'dynamics-365', 'Customer Engagement', 'app', ['D365 Customer Service', 'Dynamics 365 Customer Services', 'Customer Service (Dynamics 365)']),
  defineMicrosoftProductIcon('Dynamics 365 Contact Center', 'Dynamics 365 Contact Center', 'dynamics-365-contact-center', 'Dynamics 365 App Icons/ContactCenter_scalable.svg', 'dynamics-365', 'Customer Engagement', 'app', ['D365 Contact Center', 'Contact Center (Dynamics 365)', 'Dynamics 365 Omnichannel', 'Omnichannel for Customer Service']),
  defineMicrosoftProductIcon('Dynamics 365 Customer Insights', 'Dynamics 365 Customer Insights', 'dynamics-365-customer-insights', 'Dynamics 365 App Icons/CustomerInsights_scalable.svg', 'dynamics-365', 'Customer Engagement', 'app', ['D365 Customer Insights', 'Customer Insights', 'Customer Insights - Data', 'Customer Insights - Journeys', 'Dynamics 365 Marketing']),
  defineMicrosoftProductIcon('Dynamics 365 Customer Voice', 'Dynamics 365 Customer Voice', 'dynamics-365-customer-voice', 'Dynamics 365 App Icons/CustomerVoice_scalable.svg', 'dynamics-365', 'Customer Engagement', 'app', ['D365 Customer Voice', 'Customer Voice', 'Microsoft Forms Pro']),
  defineMicrosoftProductIcon('Dynamics 365 Field Service', 'Dynamics 365 Field Service', 'dynamics-365-field-service', 'Dynamics 365 App Icons/FieldService_scalable.svg', 'dynamics-365', 'Service', 'app', ['D365 Field Service', 'Field Service', 'Field Service (Dynamics 365)']),
  defineMicrosoftProductIcon('Dynamics 365 Finance', 'Dynamics 365 Finance', 'dynamics-365-finance', 'Dynamics 365 App Icons/Finance_scalable.svg', 'dynamics-365', 'Finance and Operations', 'app', ['D365 Finance', 'Dynamics Finance', 'Finance (Dynamics 365)']),
  defineMicrosoftProductIcon('Dynamics 365 Finance and Operations', 'Dynamics 365 Finance and Operations', 'dynamics-365-finance-operations', 'Dynamics 365 App Icons/FinanceOperations_scalable.svg', 'dynamics-365', 'Finance and Operations', 'app', ['D365 Finance and Operations', 'Dynamics 365 Finance & Operations', 'Finance and Operations', 'Dynamics 365 F&O', 'Dynamics AX']),
  defineMicrosoftProductIcon('Dynamics 365 Supply Chain Management', 'Dynamics 365 Supply Chain Management', 'dynamics-365-supply-chain-management', 'Dynamics 365 App Icons/SupplyChainManagement_scalable.svg', 'dynamics-365', 'Finance and Operations', 'app', ['D365 Supply Chain Management', 'Supply Chain Management', 'Dynamics 365 Supply Chain']),
  defineMicrosoftProductIcon('Dynamics 365 Commerce', 'Dynamics 365 Commerce', 'dynamics-365-commerce', 'Dynamics 365 App Icons/Commerce_scalable.svg', 'dynamics-365', 'Finance and Operations', 'app', ['D365 Commerce', 'Dynamics 365 Retail', 'Commerce (Dynamics 365)']),
  defineMicrosoftProductIcon('Dynamics 365 Human Resources', 'Dynamics 365 Human Resources', 'dynamics-365-human-resources', 'Dynamics 365 App Icons/HumanResources_scalable.svg', 'dynamics-365', 'Finance and Operations', 'app', ['D365 Human Resources', 'Dynamics 365 HR', 'Human Resources (Dynamics 365)']),
  defineMicrosoftProductIcon('Dynamics 365 Intelligent Order Management', 'Dynamics 365 Intelligent Order Management', 'dynamics-365-intelligent-order-management', 'Dynamics 365 App Icons/IntelligentOrderManagement_scalable.svg', 'dynamics-365', 'Finance and Operations', 'app', ['D365 Intelligent Order Management', 'Intelligent Order Management']),
  defineMicrosoftProductIcon('Dynamics 365 Project Operations', 'Dynamics 365 Project Operations', 'dynamics-365-project-operations', 'Dynamics 365 App Icons/ProjectOperations_scalable.svg', 'dynamics-365', 'Finance and Operations', 'app', ['D365 Project Operations', 'Project Operations', 'Project Service Automation']),
  defineMicrosoftProductIcon('Dynamics 365 Business Central', 'Dynamics 365 Business Central', 'dynamics-365-business-central', 'Dynamics 365 App Icons/BusinessCentral_scalable.svg', 'dynamics-365', 'Small and Medium Business', 'app', ['D365 Business Central', 'Business Central', 'Dynamics NAV', 'Microsoft Dynamics NAV']),

  // ========================================
  // Microsoft 365 architecture symbols
  // ========================================
  defineMicrosoftProductIcon('Microsoft 365 Apps', 'Apps', 'm365-apps', 'Microsoft Blue/48x48 Light Blue Icon/Apps.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', ['Microsoft 365 Apps', 'Office Apps', 'アプリ']),
  defineMicrosoftProductIcon('Microsoft 365 Apps List Detail', 'Apps List Detail', 'm365-apps-list-detail', 'Microsoft Blue/48x48 Light Blue Icon/Apps List Detail.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Arrow Circle Left', 'Arrow Circle Left', 'm365-arrow-circle-left', 'Microsoft Blue/48x48 Light Blue Icon/Arrow Circle Left.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Arrow Circle Right', 'Arrow Circle Right', 'm365-arrow-circle-right', 'Microsoft Blue/48x48 Light Blue Icon/Arrow Circle Right.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Arrow Clockwise', 'Arrow Clockwise', 'm365-arrow-clockwise', 'Microsoft Blue/48x48 Light Blue Icon/Arrow Clockwise.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Arrow Down', 'Arrow Down', 'm365-arrow-down', 'Microsoft Blue/48x48 Light Blue Icon/Arrow Down.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Arrow Download', 'Arrow Download', 'm365-arrow-download', 'Microsoft Blue/48x48 Light Blue Icon/Arrow Download.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Arrow Sync', 'Arrow Sync', 'm365-arrow-sync', 'Microsoft Blue/48x48 Light Blue Icon/Arrow Sync.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft Entra Connect', 'Directory Sync', '同期']),
  defineMicrosoftProductIcon('Microsoft 365 Arrow Upload', 'Arrow Upload', 'm365-arrow-upload', 'Microsoft Blue/48x48 Light Blue Icon/Arrow Upload.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Attach', 'Attach', 'm365-attach', 'Microsoft Blue/48x48 Light Blue Icon/Attach.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Book Contacts', 'Book Contacts', 'm365-book-contacts', 'Microsoft Blue/48x48 Light Blue Icon/Book Contacts.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Exchange Online Address Book', 'Microsoft Entra Directory', 'アドレス帳']),
  defineMicrosoftProductIcon('Microsoft 365 Bot', 'Bot', 'm365-bot', 'Microsoft Blue/48x48 Light Blue Icon/Bot.svg', 'microsoft-365', 'General Symbols', 'symbol', ['Microsoft Copilot', 'Copilot', 'Copilot Studio Agent', 'Agent', 'コパイロット', 'エージェント']),
  defineMicrosoftProductIcon('Microsoft 365 Building', 'Building', 'm365-building', 'Microsoft Blue/48x48 Light Blue Icon/Building.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['On-premises', 'Corporate Headquarters', 'オンプレミス']),
  defineMicrosoftProductIcon('Microsoft 365 Building Cloud', 'Building Cloud', 'm365-building-cloud', 'Microsoft Blue/48x48 Light Blue Icon/Building Cloud.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft 365 Tenant', 'Cloud Organization', 'テナント']),
  defineMicrosoftProductIcon('Microsoft 365 Building Multiple', 'Building Multiple', 'm365-building-multiple', 'Microsoft Blue/48x48 Light Blue Icon/Building Multiple.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Multi-tenant Organization', 'Multi-geo', 'マルチテナント']),
  defineMicrosoftProductIcon('Microsoft 365 Building People', 'Building People', 'm365-building-people', 'Microsoft Blue/48x48 Light Blue Icon/Building People_Blue_Light.svg', 'microsoft-365', 'People and Organizations', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Building Retail', 'Building Retail', 'm365-building-retail', 'Microsoft Blue/48x48 Light Blue Icon/Building Retail.svg', 'microsoft-365', 'People and Organizations', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Calculator', 'Calculator', 'm365-calculator', 'Microsoft Blue/48x48 Light Blue Icon/Calculator.svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Calendar Month', 'Calendar Month', 'm365-calendar-month', 'Microsoft Blue/48x48 Light Blue Icon/Calendar Month.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Outlook Calendar', 'Microsoft Bookings', 'Bookings', '予定表']),
  defineMicrosoftProductIcon('Microsoft 365 Call', 'Call', 'm365-call', 'Microsoft Blue/48x48 Light Blue Icon/Call.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', ['Microsoft Teams Phone', 'Teams Calling', '通話']),
  defineMicrosoftProductIcon('Microsoft 365 Camera', 'Camera', 'm365-camera', 'Microsoft Blue/48x48 Light Blue Icon/Camera.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Cart', 'Cart', 'm365-cart', 'Microsoft Blue/48x48 Light Blue Icon/Cart.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Certificate', 'Certificate', 'm365-certificate', 'Microsoft Blue/48x48 Light Blue Icon/Certificate.svg', 'microsoft-365', 'Security and Compliance', 'symbol', ['Compliance Certificate', 'Microsoft Purview Compliance Manager', '認証']),
  defineMicrosoftProductIcon('Microsoft 365 Chart Multiple', 'Chart Multiple', 'm365-chart-multiple', 'Microsoft Blue/48x48 Light Blue Icon/Chart Multiple.svg', 'microsoft-365', 'Data and Analytics', 'symbol', ['Power BI', 'Viva Insights', 'Microsoft Excel', '分析']),
  defineMicrosoftProductIcon('Microsoft 365 Chat', 'Chat', 'm365-chat', 'Microsoft Blue/48x48 Light Blue Icon/Chat.svg', 'microsoft-365', 'Communication', 'symbol', ['Teams Chat', 'Instant Message', 'メッセージ']),
  defineMicrosoftProductIcon('Microsoft 365 Chat Multiple', 'Chat Multiple', 'm365-chat-multiple', 'Microsoft Blue/48x48 Light Blue Icon/Chat Multiple.svg', 'microsoft-365', 'Communication', 'symbol', ['Microsoft Teams Chat', 'Teams Chat', 'チャット']),
  defineMicrosoftProductIcon('Microsoft 365 Check', 'Check', 'm365-check', 'Microsoft Blue/48x48 Light Blue Icon/Check.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Approval', 'Validation', '検証']),
  defineMicrosoftProductIcon('Microsoft 365 Checkmark', 'Checkmark', 'm365-checkmark', 'Planner Green/48x48 SVG Icons/Checkmark_Dark.svg', 'microsoft-365', 'General Symbols', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Checkmark Circle', 'Checkmark Circle', 'm365-checkmark-circle', 'Microsoft Blue/48x48 Light Blue Icon/Checkmark Circle.svg', 'microsoft-365', 'General Symbols', 'symbol', ['Compliance Check', 'Approved', '承認']),
  defineMicrosoftProductIcon('Microsoft 365 Chevron Circle', 'Chevron Circle', 'm365-chevron-circle', 'Microsoft Blue/48x48 Light Blue Icon/Chevron Circle.svg', 'microsoft-365', 'General Symbols', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Clipboard', 'Clipboard', 'm365-clipboard', 'Planner Green/48x48 SVG Icons/Clipboard_Dark.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Clipboard Task List', 'Clipboard Task List', 'm365-clipboard-task-list', 'Microsoft Blue/48x48 Light Blue Icon/Clipboard Task List.svg', 'microsoft-365', 'Documents and Content', 'symbol', ['Microsoft Planner', 'Planner', 'To Do', 'Microsoft To Do', 'タスク']),
  defineMicrosoftProductIcon('Microsoft 365 Clock', 'Clock', 'm365-clock', 'Microsoft Blue/48x48 Light Blue Icon/Clock.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft Purview Retention', 'Retention', '保持']),
  defineMicrosoftProductIcon('Microsoft 365 Clock Alarm', 'Clock Alarm', 'm365-clock-alarm', 'Microsoft Blue/48x48 Light Blue Icon/Clock Alarm.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Retention Policy', 'Microsoft Purview Retention', '保持期限']),
  defineMicrosoftProductIcon('Microsoft 365 Cloud', 'Cloud', 'm365-cloud', 'Microsoft Blue/48x48 Light Blue Icon/Cloud.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Cloud Add', 'Cloud Add', 'm365-cloud-add', 'Microsoft Blue/48x48 Light Blue Icon/Cloud Add.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Cloud Arrow Down', 'Cloud Arrow Down', 'm365-cloud-arrow-down', 'Microsoft Blue/48x48 Light Blue Icon/Cloud Arrow Down.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Cloud Arrow Up', 'Cloud Arrow Up', 'm365-cloud-arrow-up', 'Microsoft Blue/48x48 Light Blue Icon/Cloud Arrow Up.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Cloud Beaker', 'Cloud Beaker', 'm365-cloud-beaker', 'Microsoft Blue/48x48 Light Blue Icon/Cloud Beaker.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Cloud Cube', 'Cloud Cube', 'm365-cloud-cube', 'Microsoft Blue/48x48 Light Blue Icon/Cloud Cube.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Cloud Database', 'Cloud Database', 'm365-cloud-database', 'Microsoft Blue/48x48 Light Blue Icon/Cloud Database.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', ['Microsoft Graph', 'Microsoft 365 Substrate', 'データストア']),
  defineMicrosoftProductIcon('Microsoft 365 Cloud Desktop', 'Cloud Desktop', 'm365-cloud-desktop', 'Microsoft Blue/48x48 Light Blue Icon/Cloud Desktop.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', ['Windows 365', 'Cloud PC', 'Azure Virtual Desktop', 'クラウド PC']),
  defineMicrosoftProductIcon('Microsoft 365 Code', 'Code', 'm365-code', 'Microsoft Blue/48x48 Light Blue Icon/Code.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', ['Microsoft Graph API', 'Developer', '開発']),
  defineMicrosoftProductIcon('Microsoft 365 Column Arrow Right', 'Column Arrow Right', 'm365-column-arrow-right', 'Microsoft Blue/48x48 Light Blue Icon/Column Arrow Right.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Comment', 'Comment', 'm365-comment', 'Microsoft Blue/48x48 Light Blue Icon/Comment.svg', 'microsoft-365', 'Communication', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Comment Multiple', 'Comment Multiple', 'm365-comment-multiple', 'Microsoft Blue/48x48 Light Blue Icon/Comment Multiple.svg', 'microsoft-365', 'Communication', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Data Area', 'Data Area', 'm365-data-area', 'Microsoft Blue/48x48 Light Blue Icon/Data Area.svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Data Bar Vertical', 'Data Bar Vertical', 'm365-data-bar-vertical', 'Microsoft Blue/48x48 Light Blue Icon/Data Bar Vertical.svg', 'microsoft-365', 'Data and Analytics', 'symbol', ['Microsoft Excel', 'Excel', 'Power BI', 'グラフ']),
  defineMicrosoftProductIcon('Microsoft 365 Data Pie', 'Data Pie', 'm365-data-pie', 'Microsoft Blue/48x48 Light Blue Icon/Data Pie.svg', 'microsoft-365', 'Data and Analytics', 'symbol', ['Power BI', 'Microsoft Excel', 'レポート']),
  defineMicrosoftProductIcon('Microsoft 365 Data Trending', 'Data Trending', 'm365-data-trending', 'Microsoft Blue/48x48 Light Blue Icon/Data Trending.svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Data Trending (1)', 'Data Trending (1)', 'm365-data-trending-1', 'Microsoft Blue/48x48 Light Blue Icon/Data Trending (1).svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Data Trending (2)', 'Data Trending (2)', 'm365-data-trending-2', 'Teams Purple/48x48 Light Purple Icon/Data Trending (2).svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Data Usage', 'Data Usage', 'm365-data-usage', 'Microsoft Blue/48x48 Light Blue Icon/Data Usage.svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Database', 'Database', 'm365-database', 'Microsoft Blue/48x48 Light Blue Icon/Database.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Deploy', 'Deploy', 'm365-deploy', 'Microsoft Blue/48x48 Light Blue Icon/Deploy.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', ['Microsoft 365 Deployment', 'Rollout', '展開']),
  defineMicrosoftProductIcon('Microsoft 365 Desktop Arrow Down', 'Desktop Arrow Down', 'm365-desktop-arrow-down', 'Microsoft Blue/48x48 Light Blue Icon/Desktop Arrow Down.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', ['Windows Autopilot', 'Microsoft Intune', 'プロビジョニング']),
  defineMicrosoftProductIcon('Microsoft 365 Desktop Tower', 'Desktop Tower', 'm365-desktop-tower', 'Microsoft Blue/48x48 Light Blue Icon/Desktop Tower.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Dismiss Circle', 'Dismiss Circle', 'm365-dismiss-circle', 'Microsoft Blue/48x48 Light Blue Icon/Dismiss Circle.svg', 'microsoft-365', 'Security and Compliance', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Document', 'Document', 'm365-document', 'Microsoft Blue/48x48 Light Blue Icon/Document.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Document Bullet List', 'Document Bullet List', 'm365-document-bullet-list', 'Microsoft Blue/48x48 Light Blue Icon/Document Bullet List.svg', 'microsoft-365', 'Documents and Content', 'symbol', ['Microsoft Lists', 'Lists', 'OneNote', 'Microsoft OneNote', 'リスト']),
  defineMicrosoftProductIcon('Microsoft 365 Document Bullet List Multiple', 'Document Bullet List Multiple', 'm365-document-bullet-list-multiple', 'Microsoft Blue/48x48 Light Blue Icon/Document Bullet List Multiple.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Document Key', 'Document Key', 'm365-document-key', 'Microsoft Blue/48x48 Light Blue Icon/Document Key.svg', 'microsoft-365', 'Security and Compliance', 'symbol', ['Microsoft Purview', 'Data Loss Prevention', 'DLP', 'データ損失防止']),
  defineMicrosoftProductIcon('Microsoft 365 Document Lock', 'Document Lock', 'm365-document-lock', 'Microsoft Blue/48x48 Light Blue Icon/Document Lock.svg', 'microsoft-365', 'Security and Compliance', 'symbol', ['Microsoft Purview Information Protection', 'Sensitivity Label', '機密ラベル']),
  defineMicrosoftProductIcon('Microsoft 365 Document Multiple', 'Document Multiple', 'm365-document-multiple', 'Microsoft Blue/48x48 Light Blue Icon/Document Multiple.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Document One Page', 'Document One Page', 'm365-document-one-page', 'Microsoft Blue/48x48 Light Blue Icon/Document One Page.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Document Text', 'Document Text', 'm365-document-text', 'Microsoft Blue/48x48 Light Blue Icon/Document Text.svg', 'microsoft-365', 'Documents and Content', 'symbol', ['Microsoft Word', 'Word', 'Loop', 'Microsoft Loop', '文書']),
  defineMicrosoftProductIcon('Microsoft 365 Edit', 'Edit', 'm365-edit', 'Microsoft Blue/48x48 Light Blue Icon/Edit.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Filmstrip Play', 'Filmstrip Play', 'm365-filmstrip-play', 'Microsoft Blue/48x48 Light Blue Icon/Filmstrip Play.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', ['Microsoft Stream', 'Stream', 'Clipchamp', 'Microsoft Clipchamp', '動画']),
  defineMicrosoftProductIcon('Microsoft 365 Fingerprint', 'Fingerprint', 'm365-fingerprint', 'Microsoft Blue/48x48 Light Blue Icon/Fingerprint.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft Entra ID', 'Multifactor Authentication', 'MFA', '多要素認証']),
  defineMicrosoftProductIcon('Microsoft 365 Flag', 'Flag', 'm365-flag', 'Microsoft Blue/48x48 Light Blue Icon/Flag.svg', 'microsoft-365', 'Security and Compliance', 'symbol', ['Microsoft Purview Alert', 'Compliance Alert', 'アラート']),
  defineMicrosoftProductIcon('Microsoft 365 Folder Multiple', 'Folder Multiple', 'm365-folder-multiple', 'Microsoft Blue/48x48 Light Blue Icon/Folder Multiple.svg', 'microsoft-365', 'Documents and Content', 'symbol', ['SharePoint Document Library', 'ドキュメント ライブラリ']),
  defineMicrosoftProductIcon('Microsoft 365 Folder Open', 'Folder Open', 'm365-folder-open', 'Microsoft Blue/48x48 Light Blue Icon/Folder Open.svg', 'microsoft-365', 'Documents and Content', 'symbol', ['OneDrive', 'OneDrive for Business', 'ファイル']),
  defineMicrosoftProductIcon('Microsoft 365 Folder Open Vertical', 'Folder Open Vertical', 'm365-folder-open-vertical', 'Microsoft Blue/48x48 Light Blue Icon/Folder Open Vertical.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Folder People', 'Folder People', 'm365-folder-people', 'Microsoft Blue/48x48 Light Blue Icon/Folder People.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['SharePoint', 'SharePoint Online', 'Team Site', 'SharePoint サイト']),
  defineMicrosoftProductIcon('Microsoft 365 Globe', 'Globe', 'm365-globe', 'Microsoft Blue/48x48 Light Blue Icon/Globe.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', ['Microsoft 365 Multi-Geo', 'Internet', 'インターネット']),
  defineMicrosoftProductIcon('Microsoft 365 Grid Circles', 'Grid Circles', 'm365-grid-circles', 'Microsoft Blue/48x48 Light Blue Icon/Grid Circles.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', ['Microsoft 365 App Launcher', 'Microsoft 365', 'アプリ ランチャー']),
  defineMicrosoftProductIcon('Microsoft 365 Handshake', 'Handshake', 'm365-handshake', 'Microsoft Blue/48x48 Light Blue Icon/Handshake.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft Teams External Access', 'B2B Collaboration', '外部連携']),
  defineMicrosoftProductIcon('Microsoft 365 Hat Graduation', 'Hat Graduation', 'm365-hat-graduation', 'Microsoft Blue/48x48 Light Blue Icon/Hat Graduation.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Viva Learning', 'Learning', '学習']),
  defineMicrosoftProductIcon('Microsoft 365 Headset', 'Headset', 'm365-headset', 'Microsoft Blue/48x48 Light Blue Icon/Headset.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft Teams Phone', 'Contact Center', 'コールセンター']),
  defineMicrosoftProductIcon('Microsoft 365 Heart', 'Heart', 'm365-heart', 'Microsoft Blue/48x48 Light Blue Icon/Heart.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Viva Insights Wellbeing', 'ウェルビーイング']),
  defineMicrosoftProductIcon('Microsoft 365 Heart Pulse', 'Heart Pulse', 'm365-heart-pulse', 'Microsoft Blue/48x48 Light Blue Icon/Heart Pulse.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Viva Insights', 'Service Health', '正常性']),
  defineMicrosoftProductIcon('Microsoft 365 Image', 'Image', 'm365-image', 'Microsoft Blue/48x48 Light Blue Icon/Image.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Key', 'Key', 'm365-key', 'Microsoft Blue/48x48 Light Blue Icon/Key.svg', 'microsoft-365', 'Security and Compliance', 'symbol', ['Microsoft Entra ID', 'Entra ID', 'Credential', '資格情報']),
  defineMicrosoftProductIcon('Microsoft 365 Laptop', 'Laptop', 'm365-laptop', 'Microsoft Blue/48x48 Light Blue Icon/Laptop.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Layer Diagonal', 'Layer Diagonal', 'm365-layer-diagonal', 'Microsoft Blue/48x48 Light Blue Icon/Layer Diagonal.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', ['Microsoft Graph API', 'Service Layer', 'レイヤー']),
  defineMicrosoftProductIcon('Microsoft 365 Lightbulb', 'Lightbulb', 'm365-lightbulb', 'Microsoft Blue/48x48 Light Blue Icon/Lightbulb.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Lightbulb Checkmark', 'Lightbulb Checkmark', 'm365-lightbulb-checkmark', 'Microsoft Blue/48x48 Light Blue Icon/Lightbulb Checkmark.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft Copilot', 'Copilot Insight', 'Recommendation', '提案']),
  defineMicrosoftProductIcon('Microsoft 365 Lock Closed', 'Lock Closed', 'm365-lock-closed', 'Microsoft Blue/48x48 Light Blue Icon/Lock Closed.svg', 'microsoft-365', 'Security and Compliance', 'symbol', ['Conditional Access', 'Microsoft Entra ID Protection', '条件付きアクセス']),
  defineMicrosoftProductIcon('Microsoft 365 Lock Shield', 'Lock Shield', 'm365-lock-shield', 'Microsoft Blue/48x48 Light Blue Icon/Lock Shield.svg', 'microsoft-365', 'Security and Compliance', 'symbol', ['Microsoft Purview', 'Microsoft Defender XDR', 'Defender', 'Information Protection', '情報保護']),
  defineMicrosoftProductIcon('Microsoft 365 Mail', 'Mail', 'm365-mail', 'Microsoft Blue/48x48 Light Blue Icon/Mail.svg', 'microsoft-365', 'Communication', 'symbol', ['Exchange Online', 'Outlook', 'Microsoft Outlook', 'メール']),
  defineMicrosoftProductIcon('Microsoft 365 Mail Alert', 'Mail Alert', 'm365-mail-alert', 'Microsoft Blue/48x48 Light Blue Icon/Mail Alert.svg', 'microsoft-365', 'Security and Compliance', 'symbol', ['Exchange Online Protection', 'Defender for Office 365', 'メール脅威']),
  defineMicrosoftProductIcon('Microsoft 365 Mail Attach', 'Mail Attach', 'm365-mail-attach', 'Microsoft Blue/48x48 Light Blue Icon/Mail Attach.svg', 'microsoft-365', 'Documents and Content', 'symbol', ['Safe Attachments', 'Defender for Office 365', '添付ファイル']),
  defineMicrosoftProductIcon('Microsoft 365 Mail Error', 'Mail Error', 'm365-mail-error', 'Microsoft Blue/48x48 Light Blue Icon/Mail Error.svg', 'microsoft-365', 'Security and Compliance', 'symbol', ['Quarantine', 'Defender for Office 365', '検疫']),
  defineMicrosoftProductIcon('Microsoft 365 Mail Read', 'Mail Read', 'm365-mail-read', 'Microsoft Blue/48x48 Light Blue Icon/Mail Read.svg', 'microsoft-365', 'Communication', 'symbol', ['Outlook', 'Exchange Online Mailbox', 'メールボックス']),
  defineMicrosoftProductIcon('Microsoft 365 Map', 'Map', 'm365-map', 'Microsoft Blue/48x48 Light Blue Icon/Map.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Megaphone Loud', 'Megaphone Loud', 'm365-megaphone-loud', 'Microsoft Blue/48x48 Light Blue Icon/Megaphone Loud.svg', 'microsoft-365', 'Communication', 'symbol', ['Viva Connections', 'Viva Amplify', 'Announcement', '社内発信']),
  defineMicrosoftProductIcon('Microsoft 365 Merge', 'Merge', 'm365-merge', 'Microsoft Blue/48x48 Light Blue Icon/Merge.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', ['Tenant Migration', 'Tenant-to-Tenant', 'テナント統合']),
  defineMicrosoftProductIcon('Microsoft 365 Notebook', 'Notebook', 'm365-notebook', 'Microsoft Blue/48x48 Light Blue Icon/Notebook.svg', 'microsoft-365', 'Documents and Content', 'symbol', ['Microsoft OneNote', 'OneNote', 'ノート']),
  defineMicrosoftProductIcon('Microsoft 365 Number Circle 1', 'Number Circle 1', 'm365-number-circle-1', 'Microsoft Blue/48x48 Light Blue Icon/Number Circle 1.svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Number Circle 2', 'Number Circle 2', 'm365-number-circle-2', 'Microsoft Blue/48x48 Light Blue Icon/Number Circle 2.svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Number Circle 3', 'Number Circle 3', 'm365-number-circle-3', 'Microsoft Blue/48x48 Light Blue Icon/Number Circle 3.svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Number Circle 4', 'Number Circle 4', 'm365-number-circle-4', 'Microsoft Blue/48x48 Light Blue Icon/Number Circle 4.svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Organization', 'Organization', 'm365-organization', 'Microsoft Blue/48x48 Light Blue Icon/Organization.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft Entra Administrative Unit', '組織単位']),
  defineMicrosoftProductIcon('Microsoft 365 Organization Horizontal', 'Organization Horizontal', 'm365-organization-horizontal', 'Microsoft Blue/48x48 Light Blue Icon/Organization Horizontal.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft Entra Organizational Structure', '組織図']),
  defineMicrosoftProductIcon('Microsoft 365 Panel Left Header', 'Panel Left Header', 'm365-panel-left-header', 'Microsoft Blue/48x48 Light Blue Icon/Panel Left Header.svg', 'microsoft-365', 'General Symbols', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Panel Left Header Key', 'Panel Left Header Key', 'm365-panel-left-header-key', 'Microsoft Blue/48x48 Light Blue Icon/Panel Left Header Key.svg', 'microsoft-365', 'Security and Compliance', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Payment', 'Payment', 'm365-payment', 'Microsoft Blue/48x48 Light Blue Icon/Payment.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft 365 Licensing', 'Subscription', 'サブスクリプション']),
  defineMicrosoftProductIcon('Microsoft 365 People', 'People', 'm365-people', 'Microsoft Blue/48x48 Light Blue Icon/People.svg', 'microsoft-365', 'People and Organizations', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 People Audience', 'People Audience', 'm365-people-audience', 'Microsoft Blue/48x48 Light Blue Icon/People Audience.svg', 'microsoft-365', 'People and Organizations', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 People Community', 'People Community', 'm365-people-community', 'Microsoft Blue/48x48 Light Blue Icon/People Community.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Viva Engage', 'Yammer', 'Community', 'コミュニティ']),
  defineMicrosoftProductIcon('Microsoft 365 People Settings', 'People Settings', 'm365-people-settings', 'Microsoft Blue/48x48 Light Blue Icon/People Settings.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft 365 Admin Center', 'Microsoft Entra Group', 'グループ管理']),
  defineMicrosoftProductIcon('Microsoft 365 People Team', 'People Team', 'm365-people-team', 'Microsoft Blue/48x48 Light Blue Icon/People Team.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft Teams', 'Teams', 'Team', 'チーム', 'Teams チャネル']),
  defineMicrosoftProductIcon('Microsoft 365 Person', 'Person', 'm365-person', 'Microsoft Blue/48x48 Light Blue Icon/Person.svg', 'microsoft-365', 'People and Organizations', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Person Accounts', 'Person Accounts', 'm365-person-accounts', 'Microsoft Blue/48x48 Light Blue Icon/Person Accounts.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft Entra ID', 'User Account', 'ユーザー アカウント']),
  defineMicrosoftProductIcon('Microsoft 365 Person Desktop', 'Person Desktop', 'm365-person-desktop', 'Microsoft Blue/48x48 Light Blue Icon/Person Desktop.svg', 'microsoft-365', 'People and Organizations', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Person Settings', 'Person Settings', 'm365-person-settings', 'Microsoft Blue/48x48 Light Blue Icon/Person Settings.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft Intune', 'Microsoft 365 Admin Center', 'Administrator', '管理者']),
  defineMicrosoftProductIcon('Microsoft 365 Person Square', 'Person Square', 'm365-person-square', 'Microsoft Blue/48x48 Light Blue Icon/Person Square.svg', 'microsoft-365', 'People and Organizations', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Person Wrench', 'Person Wrench', 'm365-person-wrench', 'Microsoft Blue/48x48 Light Blue Icon/Person Wrench.svg', 'microsoft-365', 'People and Organizations', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Phone', 'Phone', 'm365-phone', 'Microsoft Blue/48x48 Light Blue Icon/Phone.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', ['Microsoft Teams Phone', 'Mobile Device', 'モバイル']),
  defineMicrosoftProductIcon('Microsoft 365 Phone Desktop', 'Phone Desktop', 'm365-phone-desktop', 'Microsoft Blue/48x48 Light Blue Icon/Phone Desktop.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', ['Microsoft Intune', 'Intune', 'Managed Device', 'マネージド デバイス']),
  defineMicrosoftProductIcon('Microsoft 365 Phone Laptop', 'Phone Laptop', 'm365-phone-laptop', 'Microsoft Blue/48x48 Light Blue Icon/Phone Laptop.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Phone Tablet', 'Phone Tablet', 'm365-phone-tablet', 'Microsoft Blue/48x48 Light Blue Icon/Phone Tablet.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Phone Update', 'Phone Update', 'm365-phone-update', 'Microsoft Blue/48x48 Light Blue Icon/Phone Update.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', ['Microsoft Intune', 'Windows Autopatch', 'デバイス更新']),
  defineMicrosoftProductIcon('Microsoft 365 Presenter', 'Presenter', 'm365-presenter', 'Microsoft Blue/48x48 Light Blue Icon/Presenter.svg', 'microsoft-365', 'People and Organizations', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Question Circle', 'Question Circle', 'm365-question-circle', 'Microsoft Blue/48x48 Light Blue Icon/Question Circle.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft 365 Support', 'Help', 'ヘルプ']),
  defineMicrosoftProductIcon('Microsoft 365 Radar', 'Radar', 'm365-radar', 'Microsoft Blue/48x48 Light Blue Icon/Radar.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', ['Microsoft Defender XDR', 'Threat Detection', '検出']),
  defineMicrosoftProductIcon('Microsoft 365 Receipt', 'Receipt', 'm365-receipt', 'Microsoft Blue/48x48 Light Blue Icon/Receipt.svg', 'microsoft-365', 'Documents and Content', 'symbol', ['Microsoft 365 Licensing', 'Billing', '請求']),
  defineMicrosoftProductIcon('Microsoft 365 Ribbon', 'Ribbon', 'm365-ribbon', 'Microsoft Blue/48x48 Light Blue Icon/Ribbon.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Ribbon Star', 'Ribbon Star', 'm365-ribbon-star', 'Microsoft Blue/48x48 Light Blue Icon/Ribbon Star.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Viva Engage Praise', 'Recognition', '表彰']),
  defineMicrosoftProductIcon('Microsoft 365 Road Cone', 'Road Cone', 'm365-road-cone', 'Microsoft Blue/48x48 Light Blue Icon/Road Cone.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft 365 Roadmap', 'Preview', 'ロードマップ']),
  defineMicrosoftProductIcon('Microsoft 365 Script', 'Script', 'm365-script', 'Microsoft Blue/48x48 Light Blue Icon/Script.svg', 'microsoft-365', 'Documents and Content', 'symbol', ['Microsoft Graph PowerShell', 'PowerShell', 'スクリプト']),
  defineMicrosoftProductIcon('Microsoft 365 Search', 'Search', 'm365-search', 'Microsoft Blue/48x48 Light Blue Icon/Search.svg', 'microsoft-365', 'Data and Analytics', 'symbol', ['Microsoft Search', 'SharePoint Search', '検索']),
  defineMicrosoftProductIcon('Microsoft 365 Select All', 'Select All', 'm365-select-all', 'Microsoft Blue/48x48 Light Blue Icon/Select All.svg', 'microsoft-365', 'Data and Analytics', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Settings', 'Settings', 'm365-settings', 'Microsoft Blue/48x48 Light Blue Icon/Settings.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft 365 Admin Center', 'Policy', 'ポリシー']),
  defineMicrosoftProductIcon('Microsoft 365 Settings Cog Multiple', 'Settings Cog Multiple', 'm365-settings-cog-multiple', 'Microsoft Blue/48x48 Light Blue Icon/Settings Cog Multiple.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft 365 Admin Center', 'Tenant Settings', 'テナント設定']),
  defineMicrosoftProductIcon('Microsoft 365 Shapes Three', 'Shapes Three', 'm365-shapes-three', 'Microsoft Blue/48x48 Light Blue Icon/Shapes Three.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Share', 'Share', 'm365-share', 'Microsoft Blue/48x48 Light Blue Icon/Share.svg', 'microsoft-365', 'Communication', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Share Screen Person', 'Share Screen Person', 'm365-share-screen-person', 'Microsoft Blue/48x48 Light Blue Icon/Share Screen Person.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft Teams Meeting', 'Teams Meeting', '会議']),
  defineMicrosoftProductIcon('Microsoft 365 Share Screen Person Overlay', 'Share Screen Person Overlay', 'm365-share-screen-person-overlay', 'Microsoft Blue/48x48 Light Blue Icon/Share Screen Person Overlay.svg', 'microsoft-365', 'People and Organizations', 'symbol', ['Microsoft Teams Live Event', 'Teams Webinar', 'ウェビナー']),
  defineMicrosoftProductIcon('Microsoft 365 Shield Error', 'Shield Error', 'm365-shield-error', 'Microsoft Blue/48x48 Light Blue Icon/Shield Error.svg', 'microsoft-365', 'Security and Compliance', 'symbol', ['Microsoft Defender for Office 365', 'Microsoft Defender XDR', 'Threat', '脅威']),
  defineMicrosoftProductIcon('Microsoft 365 Shifts Activity', 'Shifts Activity', 'm365-shifts-activity', 'Microsoft Blue/48x48 Light Blue Icon/Shifts Activity.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft Teams Shifts', 'Shifts', 'シフト']),
  defineMicrosoftProductIcon('Microsoft 365 Shopping Bag', 'Shopping Bag', 'm365-shopping-bag', 'Microsoft Blue/48x48 Light Blue Icon/Shopping Bag.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Signature', 'Signature', 'm365-signature', 'Microsoft Blue/48x48 Light Blue Icon/Signature.svg', 'microsoft-365', 'People and Organizations', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Slide Layout', 'Slide Layout', 'm365-slide-layout', 'Microsoft Blue/48x48 Light Blue Icon/Slide Layout.svg', 'microsoft-365', 'Documents and Content', 'symbol', ['Microsoft PowerPoint', 'PowerPoint', 'Sway', 'プレゼンテーション']),
  defineMicrosoftProductIcon('Microsoft 365 Tablet', 'Tablet', 'm365-tablet', 'Microsoft Blue/48x48 Light Blue Icon/Tablet.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Tablet Laptop', 'Tablet Laptop', 'm365-tablet-laptop', 'Microsoft Blue/48x48 Light Blue Icon/Tablet Laptop.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Tap', 'Tap', 'm365-tap', 'Microsoft Blue/48x48 Light Blue Icon/Tap.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Task List LTR', 'Task List LTR', 'm365-task-list-ltr', 'Planner Green/48x48 SVG Icons/Task_List_LTR_Dark.svg', 'microsoft-365', 'Process and Tasks', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Task List Square', 'Task List Square', 'm365-task-list-square', 'Microsoft Blue/48x48 Light Blue Icon/Task List Square.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft Planner', 'Planner', 'Microsoft Lists', 'Lists', 'タスク一覧']),
  defineMicrosoftProductIcon('Microsoft 365 Text Bullet List', 'Text Bullet List', 'm365-text-bullet-list', 'Project Green/48x48 SVG Icons/Text_Bullet_List_Dark.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Text Bullet List Square', 'Text Bullet List Square', 'm365-text-bullet-list-square', 'Project Green/48x48 SVG Icons/Text_Bullet_List_Square_Dark.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Text First Line', 'Text First Line', 'm365-text-first-line', 'Microsoft Blue/48x48 Light Blue Icon/Text First Line.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Text Number List', 'Text Number List', 'm365-text-number-list', 'Microsoft Blue/48x48 Light Blue Icon/Text Number List.svg', 'microsoft-365', 'Documents and Content', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Toolbox', 'Toolbox', 'm365-toolbox', 'Microsoft Blue/48x48 Light Blue Icon/Toolbox.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft 365 Admin Center', 'Admin Tools', '管理ツール']),
  defineMicrosoftProductIcon('Microsoft 365 Top Speed', 'Top Speed', 'm365-top-speed', 'Microsoft Blue/48x48 Light Blue Icon/Top Speed.svg', 'microsoft-365', 'Cloud and Infrastructure', 'symbol', ['Microsoft 365 Performance', 'Network Optimization', 'パフォーマンス']),
  defineMicrosoftProductIcon('Microsoft 365 Video', 'Video', 'm365-video', 'Microsoft Blue/48x48 Light Blue Icon/Video.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', ['Microsoft Stream', 'Teams Meeting Recording', '録画']),
  defineMicrosoftProductIcon('Microsoft 365 Wallet Credit Card', 'Wallet Credit Card', 'm365-wallet-credit-card', 'Microsoft Blue/48x48 Light Blue Icon/Wallet Credit Card.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft 365 Licensing', 'Subscription', 'ライセンス']),
  defineMicrosoftProductIcon('Microsoft 365 Window', 'Window', 'm365-window', 'Microsoft Blue/48x48 Light Blue Icon/Window.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Window Dev Edit', 'Window Dev Edit', 'm365-window-dev-edit', 'Microsoft Blue/48x48 Light Blue Icon/Window Dev Edit.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Window Edit', 'Window Edit', 'm365-window-edit', 'Microsoft Blue/48x48 Light Blue Icon/Window Edit.svg', 'microsoft-365', 'Devices and Endpoints', 'symbol', []),
  defineMicrosoftProductIcon('Microsoft 365 Wrench', 'Wrench', 'm365-wrench', 'Microsoft Blue/48x48 Light Blue Icon/Wrench.svg', 'microsoft-365', 'Process and Tasks', 'symbol', ['Microsoft 365 Admin Center', 'Configuration', '構成']),

  // ========================================
  // Microsoft 365 workloads
  //
  // The same official glyphs in the brand treatment the package ships for each
  // workload — Teams purple, SharePoint teal, Planner green, Project green, and
  // the dark blue treatment for the rest. These carry the real product names so
  // a generated diagram that says "Exchange Online" gets Microsoft's own symbol
  // for it instead of an empty box.
  // ========================================
  defineMicrosoftProductIcon('Microsoft 365', 'Microsoft 365', 'm365-app-microsoft-365', 'Microsoft Blue/48x48 Dark Blue Icon/Grid Circles.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'platform', ['M365', 'Office 365', 'O365', 'Microsoft 365 Suite', 'マイクロソフト365']),
  defineMicrosoftProductIcon('Microsoft 365 Tenant', 'Microsoft 365 Tenant', 'm365-app-tenant', 'Microsoft Blue/48x48 Dark Blue Icon/Building Cloud.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'platform', ['M365 Tenant', 'Office 365 Tenant', 'Microsoft 365 Organization', 'テナント']),
  defineMicrosoftProductIcon('Microsoft 365 Admin Center', 'Microsoft 365 Admin Center', 'm365-app-admin-center', 'Microsoft Blue/48x48 Dark Blue Icon/Toolbox.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['M365 Admin Center', 'Microsoft 365 Admin Portal', 'Office 365 Admin Center', '管理センター']),
  defineMicrosoftProductIcon('Microsoft 365 Copilot', 'Microsoft 365 Copilot', 'm365-app-copilot', 'Microsoft Blue/48x48 Dark Blue Icon/Bot.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Microsoft Copilot', 'Copilot for Microsoft 365', 'M365 Copilot', 'Copilot Chat', 'Microsoft 365 コパイロット']),
  defineMicrosoftProductIcon('Microsoft Teams', 'Microsoft Teams', 'm365-app-teams', 'Teams Purple/Building People_Teams_Light.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Teams', 'MS Teams', 'Teams Workspace', 'チームズ']),
  defineMicrosoftProductIcon('Microsoft Teams Channel', 'Microsoft Teams Channel', 'm365-app-teams-channel', 'Teams Purple/48x48 Light Purple Icon/People Team.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Teams Channel', 'Teams Team', 'Teams チャネル']),
  defineMicrosoftProductIcon('Microsoft Teams Chat', 'Microsoft Teams Chat', 'm365-app-teams-chat', 'Teams Purple/48x48 Light Purple Icon/Chat Multiple.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Teams Chat', 'Teams Messaging', 'Teams チャット']),
  defineMicrosoftProductIcon('Microsoft Teams Meeting', 'Microsoft Teams Meeting', 'm365-app-teams-meeting', 'Teams Purple/48x48 Light Purple Icon/Share Screen Person.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Teams Meeting', 'Teams Webinar', 'Teams Live Event', 'Teams 会議']),
  defineMicrosoftProductIcon('Microsoft Teams Phone', 'Microsoft Teams Phone', 'm365-app-teams-phone', 'Teams Purple/48x48 Light Purple Icon/Call.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Teams Phone', 'Teams Calling', 'Teams Voice', 'Teams 電話']),
  defineMicrosoftProductIcon('Microsoft Teams Rooms', 'Microsoft Teams Rooms', 'm365-app-teams-rooms', 'Teams Purple/48x48 Light Purple Icon/Presenter.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Teams Room', 'Teams Rooms', 'MTR', 'Teams ルーム']),
  defineMicrosoftProductIcon('Microsoft Teams Shifts', 'Microsoft Teams Shifts', 'm365-app-teams-shifts', 'Teams Purple/48x48 Light Purple Icon/Shifts Activity.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Shifts', 'Teams Shifts', 'シフト']),
  defineMicrosoftProductIcon('SharePoint Online', 'SharePoint Online', 'm365-app-sharepoint', 'SharePoint Teal/48x48 SVG Icon/Organization_Light.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['SharePoint', 'Microsoft SharePoint', 'SharePoint Site', 'SharePoint Embedded', 'シェアポイント']),
  defineMicrosoftProductIcon('OneDrive for Business', 'OneDrive for Business', 'm365-app-onedrive', 'Microsoft Blue/48x48 Dark Blue Icon/Cloud Arrow Up.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['OneDrive', 'Microsoft OneDrive', 'ワンドライブ']),
  defineMicrosoftProductIcon('Exchange Online', 'Exchange Online', 'm365-app-exchange-online', 'Microsoft Blue/48x48 Dark Blue Icon/Mail.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Exchange', 'Microsoft Exchange Online', 'Exchange Online Mailbox', 'エクスチェンジ']),
  defineMicrosoftProductIcon('Microsoft Outlook', 'Microsoft Outlook', 'm365-app-outlook', 'Microsoft Blue/48x48 Dark Blue Icon/Mail Read.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Outlook', 'Outlook on the web', 'アウトルック']),
  defineMicrosoftProductIcon('Microsoft Purview Information Protection', 'Microsoft Purview Information Protection', 'm365-app-purview-information-protection', 'Microsoft Blue/48x48 Dark Blue Icon/Lock Shield.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Purview Information Protection', 'Microsoft Information Protection', 'Sensitivity Label', 'MIP', '情報保護']),
  defineMicrosoftProductIcon('Microsoft Purview Data Loss Prevention', 'Microsoft Purview Data Loss Prevention', 'm365-app-purview-dlp', 'Microsoft Blue/48x48 Dark Blue Icon/Document Key.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Purview DLP', 'Data Loss Prevention', 'DLP', 'データ損失防止']),
  defineMicrosoftProductIcon('Microsoft Purview Compliance Manager', 'Microsoft Purview Compliance Manager', 'm365-app-purview-compliance-manager', 'Microsoft Blue/48x48 Dark Blue Icon/Certificate.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Compliance Manager', 'Purview Compliance Manager', 'コンプライアンス マネージャー']),
  defineMicrosoftProductIcon('Microsoft Defender for Office 365', 'Microsoft Defender for Office 365', 'm365-app-defender-office-365', 'Microsoft Blue/48x48 Dark Blue Icon/Shield Error.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Defender for Office 365', 'MDO', 'Office 365 ATP', 'Exchange Online Protection']),
  defineMicrosoftProductIcon('Microsoft Defender XDR', 'Microsoft Defender XDR', 'm365-app-defender-xdr', 'Microsoft Blue/48x48 Dark Blue Icon/Radar.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Defender XDR', 'Microsoft 365 Defender', 'M365 Defender', 'XDR']),
  defineMicrosoftProductIcon('Microsoft Intune', 'Microsoft Intune', 'm365-app-intune', 'Microsoft Blue/48x48 Dark Blue Icon/Phone Desktop.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Intune', 'Microsoft Endpoint Manager', 'Intune Managed Device', 'インチューン']),
  defineMicrosoftProductIcon('Windows 365', 'Windows 365', 'm365-app-windows-365', 'Microsoft Blue/48x48 Dark Blue Icon/Cloud Desktop.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Cloud PC', 'Windows 365 Cloud PC', 'ウィンドウズ365']),
  defineMicrosoftProductIcon('Microsoft Viva Connections', 'Microsoft Viva Connections', 'm365-app-viva-connections', 'Microsoft Blue/48x48 Dark Blue Icon/Megaphone Loud.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Viva Connections', 'Viva Amplify', 'ビバ コネクション']),
  defineMicrosoftProductIcon('Microsoft Viva Engage', 'Microsoft Viva Engage', 'm365-app-viva-engage', 'Microsoft Blue/48x48 Dark Blue Icon/People Community.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Viva Engage', 'Yammer', 'Microsoft Yammer', 'ビバ エンゲージ']),
  defineMicrosoftProductIcon('Microsoft Viva Learning', 'Microsoft Viva Learning', 'm365-app-viva-learning', 'Microsoft Blue/48x48 Dark Blue Icon/Hat Graduation.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Viva Learning', 'ビバ ラーニング']),
  defineMicrosoftProductIcon('Microsoft Viva Insights', 'Microsoft Viva Insights', 'm365-app-viva-insights', 'Microsoft Blue/48x48 Dark Blue Icon/Heart Pulse.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Viva Insights', 'Workplace Analytics', 'ビバ インサイト']),
  defineMicrosoftProductIcon('Microsoft Planner', 'Microsoft Planner', 'm365-app-planner', 'Planner Green/48x48 SVG Icons/Ribbon_Planner_Light.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Planner', 'Microsoft Planner Plan', 'プランナー']),
  defineMicrosoftProductIcon('Microsoft To Do', 'Microsoft To Do', 'm365-app-to-do', 'Planner Green/48x48 SVG Icons/Task_List_Square_Light.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['To Do', 'Microsoft ToDo', 'To-Do', 'トゥドゥ']),
  defineMicrosoftProductIcon('Microsoft Project', 'Microsoft Project', 'm365-app-project', 'Project Green/48x48 SVG Icons/Text_Bullet_List_Square_Light.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Project', 'Project for the web', 'Project Online', 'プロジェクト']),
  defineMicrosoftProductIcon('Microsoft OneNote', 'Microsoft OneNote', 'm365-app-onenote', 'Microsoft Blue/48x48 Dark Blue Icon/Notebook.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['OneNote', 'ワンノート']),
  defineMicrosoftProductIcon('Microsoft Word', 'Microsoft Word', 'm365-app-word', 'Microsoft Blue/48x48 Dark Blue Icon/Document Text.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Word', 'ワード']),
  defineMicrosoftProductIcon('Microsoft Excel', 'Microsoft Excel', 'm365-app-excel', 'Microsoft Blue/48x48 Dark Blue Icon/Data Bar Vertical.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Excel', 'エクセル']),
  defineMicrosoftProductIcon('Microsoft PowerPoint', 'Microsoft PowerPoint', 'm365-app-powerpoint', 'Microsoft Blue/48x48 Dark Blue Icon/Slide Layout.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['PowerPoint', 'パワーポイント']),
  defineMicrosoftProductIcon('Microsoft Loop', 'Microsoft Loop', 'm365-app-loop', 'Microsoft Blue/48x48 Dark Blue Icon/Shapes Three.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Loop', 'Loop Component', 'Loop Workspace', 'ループ']),
  defineMicrosoftProductIcon('Microsoft Lists', 'Microsoft Lists', 'm365-app-lists', 'Microsoft Blue/48x48 Dark Blue Icon/Document Bullet List.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Lists', 'SharePoint List', 'リスト']),
  defineMicrosoftProductIcon('Microsoft Forms', 'Microsoft Forms', 'm365-app-forms', 'Microsoft Blue/48x48 Dark Blue Icon/Clipboard Task List.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Forms', 'フォーム']),
  defineMicrosoftProductIcon('Microsoft Bookings', 'Microsoft Bookings', 'm365-app-bookings', 'Microsoft Blue/48x48 Dark Blue Icon/Calendar Month.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Bookings', 'Outlook Calendar', '予約']),
  defineMicrosoftProductIcon('Microsoft Stream', 'Microsoft Stream', 'm365-app-stream', 'Microsoft Blue/48x48 Dark Blue Icon/Filmstrip Play.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Stream', 'Microsoft Stream on SharePoint', 'ストリーム']),
  defineMicrosoftProductIcon('Microsoft Clipchamp', 'Microsoft Clipchamp', 'm365-app-clipchamp', 'Microsoft Blue/48x48 Dark Blue Icon/Video.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Clipchamp', 'クリップチャンプ']),
  defineMicrosoftProductIcon('Microsoft Whiteboard', 'Microsoft Whiteboard', 'm365-app-whiteboard', 'Microsoft Blue/48x48 Dark Blue Icon/Edit.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Whiteboard', 'ホワイトボード']),
  defineMicrosoftProductIcon('Microsoft Sway', 'Microsoft Sway', 'm365-app-sway', 'Microsoft Blue/48x48 Dark Blue Icon/Presenter.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Sway', 'スウェイ']),
  defineMicrosoftProductIcon('Microsoft Search', 'Microsoft Search', 'm365-app-search', 'Microsoft Blue/48x48 Dark Blue Icon/Search.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Microsoft Search in Bing', 'SharePoint Search', 'M365 検索']),
  defineMicrosoftProductIcon('Microsoft Graph', 'Microsoft Graph', 'm365-app-graph', 'Microsoft Blue/48x48 Dark Blue Icon/Layer Diagonal.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Microsoft Graph API', 'Graph API', 'Microsoft Graph Connector', 'グラフ API']),
  defineMicrosoftProductIcon('Microsoft Graph PowerShell', 'Microsoft Graph PowerShell', 'm365-app-graph-powershell', 'Microsoft Blue/48x48 Dark Blue Icon/Script.svg', 'microsoft-365', 'Microsoft 365 Workloads', 'app', ['Graph PowerShell', 'Microsoft 365 PowerShell', 'PowerShell スクリプト']),
];

const MICROSOFT_PRODUCT_ICON_BY_FILE_NAME = new Map(
  MICROSOFT_PRODUCT_ICON_CATALOG.map(definition => [definition.fileName, definition]),
);

export function getMicrosoftProductIconByFileName(
  fileName: string,
): MicrosoftProductIconDefinition | undefined {
  return MICROSOFT_PRODUCT_ICON_BY_FILE_NAME.get(fileName);
}

export const MICROSOFT_PRODUCT_ICON_CATEGORIES: MicrosoftProductIconCategory[] = [
  'power platform',
  'dynamics 365',
  'microsoft 365',
];

export function isMicrosoftProductIconCategory(
  category: string,
): category is MicrosoftProductIconCategory {
  return (MICROSOFT_PRODUCT_ICON_CATEGORIES as string[]).includes(category);
}
