// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Official Microsoft Power Platform and Dynamics 365 icon inventory.
 *
 * Source packages (pinned in `scripts/sync-microsoft-icons.mjs`):
 * - https://learn.microsoft.com/en-us/power-platform/guidance/icons
 * - https://learn.microsoft.com/en-us/dynamics365/get-started/icons
 *
 * Microsoft permits these icons in architectural diagrams, training materials,
 * and documentation. They must not be cropped, flipped, rotated, distorted, or
 * used to represent a non-Microsoft product.
 */

export type MicrosoftProductIconFamily = 'power-platform' | 'dynamics-365';

export type MicrosoftProductIconCategory = 'power platform' | 'dynamics 365';

export type MicrosoftProductIconKind = 'platform' | 'product' | 'app';

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
 * Release date of the official packages, as published in the "Icon updates"
 * table on the Microsoft Learn source pages. `scripts/sync-microsoft-icons.mjs`
 * reads this value to detect drift.
 */
export const MICROSOFT_PRODUCT_ICON_PACKAGE_VERSION = '2025-12';

const FAMILY_CATEGORY: Record<MicrosoftProductIconFamily, MicrosoftProductIconCategory> = {
  'power-platform': 'power platform',
  'dynamics-365': 'dynamics 365',
};

const FAMILY_COST_RANGE: Record<MicrosoftProductIconFamily, string> = {
  'power-platform': 'Per-user or per-app subscription (not in Azure pricing)',
  'dynamics-365': 'Per-user subscription (not in Azure pricing)',
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
    includeInServiceMap: true,
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
];

export function isMicrosoftProductIconCategory(
  category: string,
): category is MicrosoftProductIconCategory {
  return (MICROSOFT_PRODUCT_ICON_CATEGORIES as string[]).includes(category);
}
