// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useCallback, useRef, useEffect, useLayoutEffect, useMemo } from 'react';
import ReactFlow, {
  MiniMap,
  Controls,
  Background,
  useNodesState,
  useEdgesState,
  addEdge,
  Connection,
  Edge,
  Node,
  BackgroundVariant,
  MarkerType,
} from 'reactflow';
import 'reactflow/dist/style.css';
import { captureDiagramAsPng, captureDiagramAsSvg } from './utils/captureCanvas';
import { animateEdgeFlow } from './utils/animateEdges';
import { sequenceWorkflowSvg } from './utils/sequenceWorkflow';
import { buildWorkflowMarkdown } from './services/workflowNarrativeExporter';
import { Download, Save, Upload, DollarSign, Shield, ShieldCheck, FileText, FileCode, ChevronDown, ChevronRight, Clock, Camera, Loader, GitCompare, RefreshCw, PanelLeftClose, Minimize2, Maximize2, Presentation, MessageSquare, MessagesSquare, HelpCircle, Hand, ZoomIn, Frame, X, PanelTopClose, PanelTopOpen, DownloadCloud, Sun, Moon, Play, Pause, Eye, EyeOff, Cloud, Copy, Trash2, Ungroup, Boxes } from 'lucide-react';
import IconPalette from './components/IconPalette';
import AzureNode from './components/AzureNode';
import GroupNode from './components/GroupNode';
import AIArchitectureGenerator from './components/AIArchitectureGenerator';
import ArchitectureChatPanel from './components/ArchitectureChatPanel';
import HelpLearnPanel from './components/GuidedHelpPanel';
import { exportReferenceArchitectureAsPng } from './utils/exportReferencePng';
import type { ReferenceArchitecture } from './services/referenceArchitectureAI';
import { exportBlueprintArchitectureAsPng } from './utils/exportBlueprintPng';
import type { BlueprintArchitecture } from './services/blueprintArchitectureAI';
import ReferenceImageViewer from './components/ReferenceImageViewer';
import TitleBlock from './components/TitleBlock';
import ModelBadge from './components/ModelBadge';
import Legend from './components/Legend';
import EditableEdge from './components/EditableEdge';
import AlignmentToolbar from './components/AlignmentToolbar';
import WorkflowPanel from './components/WorkflowPanel';
import RegionSelector from './components/RegionSelector';
import ValidationModal from './components/ValidationModal';
import DeploymentGuideModal from './components/DeploymentGuideModal';
import IaCRoundTripModal from './components/IaCRoundTripModal';
import VersionHistoryModal from './components/VersionHistoryModal';
import SaveSnapshotModal from './components/SaveSnapshotModal';
import CloudWorkspaceModal from './components/CloudWorkspaceModal';
import PricingScenarioModal from './components/PricingScenarioModal';
import AzureImportModal from './components/AzureImportModal';
import ModelSettingsPopover from './components/ModelSettingsPopover';
import CompareModelsModal from './components/CompareModelsModal';
import CompareValidationModal from './components/CompareValidationModal';
import { loadIconsFromCategory, type AzureIcon } from './utils/iconLoader';
import { getServiceIconMapping, isCapacityConsumed } from './data/serviceIconMapping';
import { layoutArchitecture } from './utils/layoutEngine';
import { layoutArchitecture as elkLayoutArchitecture } from './utils/elkLayoutEngine';
import { initializeNodePricing, updateNodePricing, calculateCostBreakdown, exportCostBreakdownCSV, exportCostBreakdownJSON, getCostSummaryMarkdown, refreshAllNodePricing, type PricingMode } from './services/costEstimationService';
import { prefetchCommonServices } from './services/azurePricingService';
import { preloadCommonServices, getActiveRegion, AzureRegion, AVAILABLE_REGIONS, RegionInfo } from './services/regionalPricingService';
import JSZip from 'jszip';
import { formatMonthlyCost, getPricingFreshness } from './utils/pricingHelpers';
import { hasPricingData, PRICING_DATA_AS_OF } from './data/azurePricing';
import { costReportToHtml } from './utils/costReportHtml';
import { validateArchitecture, ArchitectureValidation } from './services/architectureValidator';
import { bandLabel } from './services/wafMaturity';
import { generateDeploymentGuide, DeploymentGuide } from './services/deploymentGuideGenerator';
import { generateArchitectureWithAI } from './services/azureOpenAI';
import { MODEL_CONFIG, DEPLOYMENT_NAMES, type ModelType } from './stores/modelSettingsStore';
import { usePricingDisplayPrefs } from './stores/pricingDisplayStore';
import {
  useNodePricingEditor,
  closeNodePricingEditor,
  getNodePricingEditorStateVersion,
  openNodePricingEditor,
} from './stores/nodePricingEditorStore';
import NodePricingEditor from './components/NodePricingEditor';
import type { NodePricingConfig, PricingScenario } from './types/pricing';
import {
  loadPricingScenarios,
  normalizePricingScenarios,
  savePricingScenarios,
} from './services/pricingScenarioService';
import { createSnapshot, DiagramVersion, getVersion } from './services/versionStorageService';
import { useCloudDiagramSync } from './hooks/useCloudDiagramSync';
import { exportAndDownloadDrawio } from './services/drawioExporter';
import { buildVsdxBlob } from './services/visioVsdxExporter';
import { exportDiagramAsPptx, exportArchitectureDeck, type DeckService } from './services/pptxExporter';
import { extractArchitectureFromArm, summarizeCoverage } from './services/armExtractor';
import {
  buildIaCBaseline,
  buildStarterTemplate,
  compareDiagramToBaseline,
  parseDeploymentPlan,
  restoreIaCBaseline,
  type DriftPlanSummary,
  type IaCBaseline,
  type StarterTemplateFormat,
} from './services/iacRoundTrip';
import { buildArchitectureFromResources } from './services/resourceGraphAdapter';
import { getResources as getAzureResources } from './services/azureImportProvider';
import { isDelegatedAuthConfigured, getSignedInName, consumeReopenFlag } from './services/msalAuth';
import { exportDiagramAsHtml } from './services/htmlDiagramExporter';
import {
  applyLayoutPreset,
  type LayoutPreset,
  type LayoutSpacing,
  type LayoutEdgeStyle,
  type LayoutEngineType,
} from './utils/layoutPresets';
import { generateModelFilename, setSourceModel, clearSourceModel } from './utils/modelNaming';
import {
  collectNodeAndDescendantIds,
  deleteNodesPreservingGroupChildren,
  detachChildrenFromGroups,
  fitGroupToContent,
  fitAllGroupsToContent,
} from './utils/groupUtils';
import {
  buildAbsolutePositionMap,
  preserveManualLayout,
  selectHorizontalConnectionHandles,
} from './utils/preserveManualLayout';
import { trackArchitectureGeneration, trackValidation, trackValidationHandoff, trackDeploymentGuide, trackExport, trackTemplateImport, trackModelComparison, trackRecommendationsApplied, trackVersionOperation, trackStartFresh, trackValidationFindings } from './services/telemetryService';
import { classifyValidationTopics } from './services/validationConsensus';
import type { IaCFormat } from './services/azureOpenAI';
import FeedbackModal from './components/FeedbackModal';
import FeedbackToast from './components/FeedbackToast';
import AccessManagementModal from './components/AccessManagementModal';
import LanguageSwitch from './components/LanguageSwitch';
import ValidationHandoffToast from './components/ValidationHandoffToast';
import { FEEDBACK_DONE_KEY } from './services/feedbackService';
import { getAccessIdentity, type AccessIdentity } from './services/accessControlService';
import './App.css';
import { useLanguage } from './i18n/LanguageContext';
import { localize } from './i18n/localization';
import { decodeUtf8Base64 } from './utils/base64Utf8';
import { csvTextCell } from './utils/csv';
import { toFileNameSegment } from './utils/fileName';
import { readBooleanPreference, readLocalStorage, writeLocalStorage } from './utils/safeStorage';

const nodeTypes = {
  azureNode: AzureNode,
  groupNode: GroupNode,
};

const edgeTypes = {
  editableEdge: EditableEdge,
};

type ExportHistoryKind = 'png' | 'svg' | 'animated-svg' | 'workflow-animation' | 'workflow-md' | 'costs' | 'json' | 'drawio' | 'vsdx' | 'pptx' | 'html';

type ExportHistoryItem = {
  id: string;
  kind: ExportHistoryKind;
  fileName: string;
  createdAt: number;
};

const EXPORT_HISTORY_STORAGE_KEY = 'azure-diagram-builder.exportHistory.v1';
const EDGE_STYLE_STORAGE_KEY = 'azure-diagram-builder.edgeStyle.v1';
const EDGE_ANIMATION_STORAGE_KEY = 'azure-diagram-builder.edgeAnimation.v1';
const CANVAS_HINT_STORAGE_KEY = 'azure-diagram-builder.canvasHintDismissed.v1';
const HEADER_COLLAPSED_STORAGE_KEY = 'azure-diagram-builder.headerCollapsed.v1';
const TOOLBAR_SECTIONS_STORAGE_KEY = 'azure-diagram-builder.toolbarSections.v1';
const RIBBON_TAB_STORAGE_KEY = 'azure-diagram-builder.ribbonTab.v1';
const EDGE_CONTEXT_MENU_WIDTH = 220;
const EDGE_CONTEXT_MENU_HEIGHT = 280;
const EDGE_CONTEXT_MENU_MARGIN = 8;
const NODE_CONTEXT_MENU_WIDTH = 280;
const NODE_CONTEXT_MENU_HEIGHT = 360;

type NodeContextMenuState = {
  x: number;
  y: number;
  nodeId: string;
};

function clampContextMenuPosition(
  clientX: number,
  clientY: number,
  width: number,
  height: number,
): { x: number; y: number } {
  const maxX = Math.max(
    EDGE_CONTEXT_MENU_MARGIN,
    window.innerWidth - width - EDGE_CONTEXT_MENU_MARGIN,
  );
  const maxY = Math.max(
    EDGE_CONTEXT_MENU_MARGIN,
    window.innerHeight - height - EDGE_CONTEXT_MENU_MARGIN,
  );
  return {
    x: Math.min(Math.max(EDGE_CONTEXT_MENU_MARGIN, clientX), maxX),
    y: Math.min(Math.max(EDGE_CONTEXT_MENU_MARGIN, clientY), maxY),
  };
}

function pricingFingerprint(pricing: NodePricingConfig | undefined): string {
  if (!pricing) return 'none';
  return JSON.stringify([
    pricing.estimatedCost,
    pricing.customPrice ?? null,
    pricing.tier,
    pricing.tierId ?? null,
    pricing.skuName,
    pricing.quantity,
    pricing.region,
    pricing.unit,
    pricing.lastUpdated,
    pricing.isCustom,
    pricing.isUsageBased ?? false,
    pricing.reserved1yrCost ?? null,
    pricing.reservedIsSavingsPlan ?? false,
    pricing.usageEstimate?.type ?? null,
    pricing.usageEstimate?.description ?? null,
  ]);
}

function nodePricingFingerprint(node: Node): string {
  return JSON.stringify([
    node.type,
    node.data?.label ?? null,
    node.data?.serviceName ?? null,
    pricingFingerprint(node.data?.pricing as NodePricingConfig | undefined),
  ]);
}

function createCustomPricingDraft(region: AzureRegion): NodePricingConfig {
  return {
    estimatedCost: 0,
    tier: 'Custom',
    skuName: 'Custom',
    quantity: 1,
    region,
    unit: 'per month',
    lastUpdated: new Date().toISOString(),
    isCustom: true,
    customPrice: 0,
  };
}

type RegionalCostResult = {
  info: RegionInfo;
  total: number;
  annual: number;
  breakdown: ReturnType<typeof calculateCostBreakdown>;
};

type RegionalCostFailure = {
  info: RegionInfo;
  reason: string;
};

async function calculateRegionalCostComparison(
  nodes: Node[],
  pricingMode: PricingMode,
): Promise<{ results: RegionalCostResult[]; failures: RegionalCostFailure[] }> {
  const results: RegionalCostResult[] = [];
  const failures: RegionalCostFailure[] = [];

  for (const info of AVAILABLE_REGIONS) {
    try {
      const repricedNodes = await refreshAllNodePricing(nodes, info.id);
      const breakdown = calculateCostBreakdown(repricedNodes, info.id, pricingMode);
      results.push({
        info,
        total: breakdown.totalMonthlyCost,
        annual: breakdown.totalMonthlyCost * 12,
        breakdown,
      });
    } catch (error) {
      const reason = error instanceof Error
        ? error.message.replace(/\s+/g, ' ').slice(0, 240)
        : 'Pricing data is unavailable for at least one selected SKU.';
      failures.push({ info, reason });
    }
  }

  results.sort((left, right) => left.total - right.total);
  return { results, failures };
}

const TOOLBAR_SECTION_IDS = [
  'context',
  'create',
  'import',
  'file',
  'workspace',
  'history',
  'arrange',
  'review',
] as const;
type ToolbarSectionId = typeof TOOLBAR_SECTION_IDS[number];
const TOOLBAR_SECTION_ID_SET = new Set<string>(TOOLBAR_SECTION_IDS);
const RIBBON_TAB_IDS = ['home', 'create', 'design', 'review'] as const;
type RibbonTabId = typeof RIBBON_TAB_IDS[number];
const RIBBON_TAB_ID_SET = new Set<string>(RIBBON_TAB_IDS);

const normalizeLayoutEdgeStyle = (value: unknown): LayoutEdgeStyle =>
  value === 'straight' || value === 'smooth' || value === 'orthogonal'
    ? value
    : 'orthogonal';

type RestoredWorkflowStep = Record<string, unknown> & {
  step: number;
  description: string;
  services: string[];
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function validateRestoredNodes(restoredNodes: unknown[]): Node[] {
  const nodeIds = new Set<string>();
  const nodes = restoredNodes.map((value, index) => {
    if (!isRecord(value)) {
      throw new Error(`Invalid node at index ${index}`);
    }
    if (typeof value.id !== 'string' || value.id.trim() === '') {
      throw new Error(`Node at index ${index} must have an id`);
    }
    if (nodeIds.has(value.id)) {
      throw new Error(`Duplicate node id: ${value.id}`);
    }
    if (
      !isRecord(value.position)
      || !Number.isFinite(value.position.x)
      || !Number.isFinite(value.position.y)
    ) {
      throw new Error(`Node ${value.id} must have a finite position`);
    }
    if (!isRecord(value.data)) {
      throw new Error(`Node ${value.id} must have a data object`);
    }
    if (value.type !== undefined && typeof value.type !== 'string') {
      throw new Error(`Node ${value.id} has an invalid type`);
    }
    if (value.parentNode !== undefined && typeof value.parentNode !== 'string') {
      throw new Error(`Node ${value.id} has an invalid parent`);
    }

    const data = value.data;
    const stringDataFields = [
      'label',
      'serviceName',
      'category',
      'iconPath',
      'description',
      'stylePreset',
      'groupId',
      'groupLabel',
    ];
    for (const field of stringDataFields) {
      if (data[field] !== undefined && typeof data[field] !== 'string') {
        throw new Error(`Node ${value.id} has an invalid ${field}`);
      }
    }
    if (data.pricing !== undefined && data.pricing !== null) {
      if (!isRecord(data.pricing)) {
        throw new Error(`Node ${value.id} has invalid pricing data`);
      }
      const pricing = data.pricing;
      for (const field of ['estimatedCost', 'customPrice', 'reserved1yrCost']) {
        const amount = pricing[field];
        if (
          amount !== undefined
          && amount !== null
          && (typeof amount !== 'number' || !Number.isFinite(amount) || amount < 0)
        ) {
          throw new Error(`Node ${value.id} has invalid pricing ${field}`);
        }
      }
      const estimatedCost = pricing.estimatedCost;
      if (
        typeof estimatedCost !== 'number'
        || !Number.isFinite(estimatedCost)
        || estimatedCost < 0
      ) {
        throw new Error(`Node ${value.id} has invalid pricing estimatedCost`);
      }
      const importedQuantity = pricing.quantity;
      let quantity = 1;
      if (importedQuantity === undefined || importedQuantity === null) {
        pricing.quantity = quantity;
      } else if (
        typeof importedQuantity !== 'number'
        || !Number.isInteger(importedQuantity)
        || importedQuantity < 1
        || importedQuantity > 100_000
      ) {
        throw new Error(`Node ${value.id} has invalid pricing quantity`);
      } else {
        quantity = importedQuantity;
      }
      if (!Number.isFinite(estimatedCost * quantity)) {
        throw new Error(`Node ${value.id} has an invalid total price`);
      }
      for (const field of ['isCustom', 'isUsageBased', 'reservedIsSavingsPlan']) {
        const flag = pricing[field];
        if (flag !== undefined && flag !== null && typeof flag !== 'boolean') {
          throw new Error(`Node ${value.id} has invalid pricing ${field}`);
        }
      }
      for (const field of ['tier', 'tierId', 'skuName', 'unit', 'lastUpdated']) {
        const text = pricing[field];
        if (text !== undefined && text !== null && typeof text !== 'string') {
          throw new Error(`Node ${value.id} has invalid pricing ${field}`);
        }
      }
      if (
        pricing.region === undefined
        || pricing.region === null
        || (typeof pricing.region === 'string' && pricing.region.trim() === '')
      ) {
        pricing.region = 'Unknown';
      } else {
        if (typeof pricing.region !== 'string') {
          throw new Error(`Node ${value.id} has invalid pricing region`);
        }
        const region = pricing.region.trim();
        if (!/^[A-Za-z0-9][A-Za-z0-9 ._()-]{0,127}$/.test(region)) {
          throw new Error(`Node ${value.id} has invalid pricing region`);
        }
        pricing.region = region;
      }
    }
    if (data.customColor !== undefined && data.customColor !== null) {
      if (
        !isRecord(data.customColor)
        || typeof data.customColor.bg !== 'string'
        || typeof data.customColor.border !== 'string'
        || typeof data.customColor.header !== 'string'
      ) {
        throw new Error(`Node ${value.id} has invalid custom colors`);
      }
    }

    nodeIds.add(value.id);
    return value as unknown as Node;
  });

  for (const node of nodes) {
    if (node.parentNode !== undefined && (!nodeIds.has(node.parentNode) || node.parentNode === node.id)) {
      throw new Error(`Node ${node.id} references an invalid parent`);
    }
  }
  return nodes;
}

function validateRestoredEdges(restoredEdges: unknown[], nodeIds: Set<string>): unknown[] {
  const edgeIds = new Set<string>();
  return restoredEdges.map((value, index) => {
    if (!isRecord(value)) {
      throw new Error(`Invalid edge at index ${index}`);
    }
    if (typeof value.id !== 'string' || value.id.trim() === '') {
      throw new Error(`Edge at index ${index} must have an id`);
    }
    if (edgeIds.has(value.id)) {
      throw new Error(`Duplicate edge id: ${value.id}`);
    }
    if (
      typeof value.source !== 'string'
      || typeof value.target !== 'string'
      || !nodeIds.has(value.source)
      || !nodeIds.has(value.target)
    ) {
      throw new Error(`Edge ${value.id} references an unknown node`);
    }
    if (value.data !== undefined && value.data !== null && !isRecord(value.data)) {
      throw new Error(`Edge ${value.id} has invalid data`);
    }
    if (value.type !== undefined && typeof value.type !== 'string') {
      throw new Error(`Edge ${value.id} has an invalid type`);
    }
    for (const field of ['sourceHandle', 'targetHandle']) {
      if (value[field] !== undefined && value[field] !== null && typeof value[field] !== 'string') {
        throw new Error(`Edge ${value.id} has an invalid ${field}`);
      }
    }
    edgeIds.add(value.id);
    return value;
  });
}

function validateRestoredWorkflow(value: unknown): RestoredWorkflowStep[] {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value) || value.length > 2000) {
    throw new Error('Invalid workflow in diagram payload');
  }
  return value.map((step, index) => {
    if (
      !isRecord(step)
      || !Number.isFinite(step.step)
      || typeof step.description !== 'string'
      || !Array.isArray(step.services)
      || step.services.some((service: unknown) => typeof service !== 'string')
    ) {
      throw new Error(`Invalid workflow step at index ${index}`);
    }
    return step as RestoredWorkflowStep;
  });
}

function fitToolbarMenuToViewport(menu: HTMLElement) {
  const edgeGap = 12;
  const triggerGap = 10;

  menu.style.position = '';
  menu.style.top = '';
  menu.style.bottom = '';
  menu.style.left = '';
  menu.style.right = '';
  menu.style.width = '';
  menu.style.minWidth = '';
  menu.style.maxWidth = '';
  menu.style.maxHeight = '';
  menu.style.overflowX = '';
  menu.style.overflowY = '';
  menu.style.transform = '';

  const visualViewport = window.visualViewport;
  const viewportLeft = visualViewport?.offsetLeft ?? 0;
  const viewportTop = visualViewport?.offsetTop ?? 0;
  const viewportWidth = visualViewport?.width ?? window.innerWidth;
  const viewportHeight = visualViewport?.height ?? window.innerHeight;
  const viewportRight = viewportLeft + viewportWidth;
  const viewportBottom = viewportTop + viewportHeight;
  const menuMaxWidth = Math.max(0, viewportWidth - edgeGap * 2);
  menu.style.maxWidth = `${menuMaxWidth}px`;
  const configuredMinWidth = Number.parseFloat(window.getComputedStyle(menu).minWidth);
  if (Number.isFinite(configuredMinWidth) && configuredMinWidth > menuMaxWidth) {
    menu.style.minWidth = `${menuMaxWidth}px`;
  }
  menu.style.overflowX = 'auto';
  const trigger = menu.parentElement?.getBoundingClientRect();

  if (window.matchMedia('(max-width: 640px), (max-width: 1180px) and (max-height: 600px)').matches) {
    const ribbonTabsBottom = document.querySelector('.ribbon-tabs')?.getBoundingClientRect().bottom;
    const minimumTop = Math.max(viewportTop + edgeGap, (ribbonTabsBottom ?? 78) + 4);
    let top = Math.max(minimumTop, (trigger?.bottom ?? minimumTop) + triggerGap);
    if (viewportBottom - top - edgeGap < 80) {
      top = minimumTop;
    }
    menu.style.position = 'fixed';
    menu.style.top = `${top}px`;
    menu.style.left = `${viewportLeft + edgeGap}px`;
    menu.style.right = 'auto';
    menu.style.width = `${menuMaxWidth}px`;
    menu.style.maxHeight = `${Math.max(0, viewportBottom - top - edgeGap)}px`;
    menu.style.overflowY = 'auto';
    return;
  }

  if (!trigger) return;

  const availableBelow = viewportBottom - trigger.bottom - triggerGap - edgeGap;
  const availableAbove = trigger.top - viewportTop - triggerGap - edgeGap;
  const naturalHeight = menu.scrollHeight;
  const placeAbove = naturalHeight > availableBelow && availableAbove > availableBelow;
  const availableHeight = placeAbove ? availableAbove : availableBelow;
  const triggerOutsideViewport = trigger.bottom <= viewportTop + edgeGap
    || trigger.top >= viewportBottom - edgeGap;

  if (triggerOutsideViewport || availableHeight < 80) {
    const menuWidth = menu.getBoundingClientRect().width;
    const left = Math.min(
      Math.max(trigger.left, viewportLeft + edgeGap),
      Math.max(viewportLeft + edgeGap, viewportRight - menuWidth - edgeGap),
    );
    menu.style.position = 'fixed';
    menu.style.top = `${viewportTop + edgeGap}px`;
    menu.style.left = `${left}px`;
    menu.style.right = 'auto';
    menu.style.maxHeight = `${Math.max(0, viewportHeight - edgeGap * 2)}px`;
  } else {
    menu.style.top = placeAbove ? 'auto' : `calc(100% + ${triggerGap}px)`;
    menu.style.bottom = placeAbove ? `calc(100% + ${triggerGap}px)` : 'auto';
    menu.style.maxHeight = `${availableHeight}px`;
  }
  menu.style.overflowY = 'auto';

  const rect = menu.getBoundingClientRect();
  const horizontalShift = rect.left < viewportLeft + edgeGap
    ? viewportLeft + edgeGap - rect.left
    : rect.right > viewportRight - edgeGap
      ? viewportRight - edgeGap - rect.right
      : 0;
  if (horizontalShift !== 0) {
    menu.style.transform = `translateX(${horizontalShift}px)`;
  }
}

// Derive a short, human-friendly architecture title from a free-form prompt
// (used as a fallback when no manifest title is available). Strips common
// prefixes like "MODIFY EXISTING ARCHITECTURE: ...", "Build a", "Design a",
// then takes the first ~8 words and title-cases them.
function deriveTitleFromPrompt(prompt: string | undefined | null): string | undefined {
  if (!prompt) return undefined;
  let p = String(prompt).trim();
  // Strip our own context-prompt prefix if present.
  const modMatch = p.match(/CHANGE REQUESTED:\s*(.+)$/is);
  if (modMatch) p = modMatch[1].trim();
  // Drop leading verbs / fillers.
  p = p.replace(/^(please\s+)?(build|design|create|generate|make|draw|architect|show|produce)\s+(me\s+)?(a|an|the)?\s*/i, '');
  // Take first sentence / line.
  p = p.split(/[\n.!?]/)[0].trim();
  if (!p) return undefined;
  const words = p.split(/\s+/).slice(0, 8);
  if (words.length === 0) return undefined;
  const titled = words
    .map((w) => {
      if (w.length <= 3 && w === w.toUpperCase()) return w; // keep acronyms (IoT, AKS)
      return w.charAt(0).toUpperCase() + w.slice(1);
    })
    .join(' ')
    .replace(/[^\w\s\-/&]/g, '')
    .trim();
  return titled.length > 0 ? titled : undefined;
}

function App() {
  const { t, translate, language } = useLanguage();
  const [nodes, setNodes, onNodesChangeBase] = useNodesState([]);
  const latestNodesRef = useRef<Node[]>(nodes);
  latestNodesRef.current = nodes;
  const [edges, setEdges, onEdgesChange] = useEdgesState([]);
  const [architecturePrompt, setArchitecturePrompt] = useState<string>('');
  // The FIRST prompt of the current diagram lineage. Unlike architecturePrompt
  // (which each chat refinement overwrites), this is captured once when the
  // canvas is empty so the customer deck's "brief" reflects the original ask.
  const [originalPrompt, setOriginalPrompt] = useState<string>('');
  const [promptBannerPosition, setPromptBannerPosition] = useState({ x: 0, y: 0 });
  const [isDraggingBanner, setIsDraggingBanner] = useState(false);
  const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 });

  const [isImportingTemplate, setIsImportingTemplate] = useState(false);
  const [isAzureImportOpen, setIsAzureImportOpen] = useState(false);
  // After a delegated sign-in redirect returns, re-open the "Import from Azure"
  // modal so the user lands back where they left off (now signed in).
  useEffect(() => {
    if (!isDelegatedAuthConfigured()) return;
    getSignedInName()
      .then(() => {
        if (consumeReopenFlag()) setIsAzureImportOpen(true);
      })
      .catch((error) => {
        console.error('Failed to complete delegated Azure sign-in:', error);
        if (consumeReopenFlag()) setIsAzureImportOpen(true);
      });
  }, []);
  const [importFormatLabel, setImportFormatLabel] = useState('Template');
  const [isApplyingRecommendations, setIsApplyingRecommendations] = useState(false);
  const [reactFlowInstance, setReactFlowInstance] = useState<any>(null);
  
  const onNodesChange = useCallback((changes: any[]) => {
    onNodesChangeBase(changes);
  }, [onNodesChangeBase]);

  
  const [workflow, setWorkflow] = useState<any[]>([]);
  // const [showWorkflow, setShowWorkflow] = useState(false);
  const [highlightedServices, setHighlightedServices] = useState<string[]>([]);
  const [nodeContextMenu, setNodeContextMenu] = useState<NodeContextMenuState | null>(null);
  const nodeContextMenuRef = useRef<HTMLDivElement>(null);
  const nodeContextMenuReturnFocusRef = useRef<HTMLElement | null>(null);
  const [edgeContextMenu, setEdgeContextMenu] = useState<{ x: number; y: number; edgeId: string } | null>(null);
  const [totalMonthlyCost, setTotalMonthlyCost] = useState(0);
  const [pricingMode, setPricingMode] = useState<PricingMode>('payg');
  const [pricingScenarios, setPricingScenarios] = useState<PricingScenario[]>(
    () => loadPricingScenarios(),
  );
  useEffect(() => {
    savePricingScenarios(pricingScenarios);
  }, [pricingScenarios]);
  const hasCostReportData = nodes.some(
    (node) => node.type === 'azureNode' && Boolean(node.data?.pricing),
  );
  const hasCostDisplayData = nodes.some((node) => {
    if (node.type !== 'azureNode') return false;
    const serviceName = String(node.data?.serviceName || node.data?.label || '');
    return Boolean(node.data?.pricing) || isCapacityConsumed(serviceName);
  });
  // Whether cost estimates are shown at all (persisted, independent of stylePreset).
  const [pricingPrefs, setPricingPrefs] = usePricingDisplayPrefs();
  // Node whose per-node cost editor is open (opened from its cost badge).
  const pricingEditorNodeId = useNodePricingEditor();
  const [pricingEditorDraft, setPricingEditorDraft] = useState<{
    nodeId: string;
    pricing: NodePricingConfig;
  } | null>(null);
  const pricingEditorOpenRunRef = useRef(0);
  const cancelPendingPricingEditorOpen = useCallback(() => {
    pricingEditorOpenRunRef.current += 1;
  }, []);

  const deleteCanvasNodes = useCallback((
    nodeIds: Iterable<string>,
    preserveGroupChildren = true,
  ) => {
    const deletedIds = new Set(nodeIds);
    if (deletedIds.size === 0) return;
    setNodes((currentNodes) => (
      preserveGroupChildren
        ? deleteNodesPreservingGroupChildren(currentNodes, deletedIds)
        : currentNodes.filter(node => !deletedIds.has(node.id))
    ));
    setEdges((currentEdges) => currentEdges.filter(edge => (
      !deletedIds.has(edge.source) && !deletedIds.has(edge.target)
    )));
  }, [setEdges, setNodes]);

  const [titleBlockData, setTitleBlockData] = useState({
    architectureName: 'Untitled Architecture',
    author: 'Azure Architect',
    version: '1.0',
    date: new Date().toLocaleDateString(),
  });

  useEffect(() => {
    setTitleBlockData((current) => {
      const architectureName = current.architectureName === 'Untitled Architecture'
        || current.architectureName === '無題のアーキテクチャ'
        ? translate('Untitled Architecture')
        : current.architectureName;
      const author = current.author === 'Azure Architect'
        || current.author === 'Azure アーキテクト'
        ? translate('Azure Architect')
        : current.author;
      if (architectureName === current.architectureName && author === current.author) return current;
      return { ...current, architectureName, author };
    });
  }, [language, translate]);

  useEffect(() => {
    const labels = [
      ['.react-flow__controls-zoomin', translate('Zoom in')],
      ['.react-flow__controls-zoomout', translate('Zoom out')],
      ['.react-flow__controls-fitview', translate('Fit diagram to view')],
      ['.react-flow__controls-interactive', translate('Toggle interactivity')],
    ] as const;

    const updateControlLabels = () => {
      for (const [selector, label] of labels) {
        document.querySelectorAll<HTMLButtonElement>(selector).forEach(control => {
          control.title = label;
          control.setAttribute('aria-label', label);
        });
      }
    };

    updateControlLabels();
    const observer = new MutationObserver(updateControlLabels);
    observer.observe(document.body, { childList: true, subtree: true });
    return () => observer.disconnect();
  }, [language, translate]);
  
  // Theme State
  const [isDarkMode, setIsDarkMode] = useState<boolean>(() => readBooleanPreference('darkMode', false));
  
  // Premium Features State
  const [validationResult, setValidationResult] = useState<ArchitectureValidation | null>(null);
  const [isValidationModalOpen, setIsValidationModalOpen] = useState(false);
  const [isValidating, setIsValidating] = useState(false);
  const [deploymentGuide, setDeploymentGuide] = useState<DeploymentGuide | null>(null);
  const [isDeploymentGuideModalOpen, setIsDeploymentGuideModalOpen] = useState(false);
  const [isGeneratingGuide, setIsGeneratingGuide] = useState(false);
  const [generatedWithModel, setGeneratedWithModel] = useState<{ name: string; timeMs?: number } | null>(null);
  const [iacBaseline, setIaCBaseline] = useState<IaCBaseline | null>(null);
  const [isIaCRoundTripModalOpen, setIsIaCRoundTripModalOpen] = useState(false);
  const [driftPlanSummary, setDriftPlanSummary] = useState<DriftPlanSummary | null>(null);

  // Version History State
  const [isVersionHistoryModalOpen, setIsVersionHistoryModalOpen] = useState(false);
  const [isSaveSnapshotModalOpen, setIsSaveSnapshotModalOpen] = useState(false);
  const [isCloudWorkspaceOpen, setIsCloudWorkspaceOpen] = useState(false);
  const [isPricingScenarioModalOpen, setIsPricingScenarioModalOpen] = useState(false);
  const [isCompareModelsOpen, setIsCompareModelsOpen] = useState(false);
  const [isChatOpen, setIsChatOpen] = useState(false);
  const [isHelpOpen, setIsHelpOpen] = useState(false);
  // First-run nudge: pulse the Help button until it has been opened once.
  const [helpSeen, setHelpSeen] = useState<boolean>(() => readLocalStorage('help.seen') === '1');
  // Canvas navigation hint: teaches scroll-to-zoom / drag-to-pan / fit-view.
  // Dismissed permanently once the user closes it (persisted in localStorage).
  const [showCanvasHint, setShowCanvasHint] = useState<boolean>(() => readLocalStorage(CANVAS_HINT_STORAGE_KEY) !== '1');
  // Collapses the top toolbar rows to maximize canvas height. Independent of
  // the "Focus" button (which collapses the side panels). Persisted so the
  // user's preference sticks across sessions.
  const [isHeaderCollapsed, setIsHeaderCollapsed] = useState<boolean>(() => {
    const stored = readLocalStorage(HEADER_COLLAPSED_STORAGE_KEY);
    if (stored !== null) return stored === '1';
    return window.matchMedia('(max-width: 1440px)').matches;
  });
  const [activeRibbonTab, setActiveRibbonTab] = useState<RibbonTabId>(() => {
    const stored = readLocalStorage(RIBBON_TAB_STORAGE_KEY);
    return stored && RIBBON_TAB_ID_SET.has(stored)
      ? stored as RibbonTabId
      : 'home';
  });
  const [collapsedToolbarSections, setCollapsedToolbarSections] = useState<Set<ToolbarSectionId>>(() => {
    const stored = readLocalStorage(TOOLBAR_SECTIONS_STORAGE_KEY);
    if (!stored) return new Set();
    try {
      const parsed = JSON.parse(stored);
      if (!Array.isArray(parsed)) return new Set();
      return new Set(
        parsed.filter(
          (sectionId): sectionId is ToolbarSectionId =>
            typeof sectionId === 'string' && TOOLBAR_SECTION_ID_SET.has(sectionId)
        )
      );
    } catch {
      return new Set();
    }
  });
  const [isCompareValidationOpen, setIsCompareValidationOpen] = useState(false);
  const [isFeedbackModalOpen, setIsFeedbackModalOpen] = useState(false);
  const [isFeedbackToastOpen, setIsFeedbackToastOpen] = useState(false);
  const [validationHandoff, setValidationHandoff] = useState<{
    source: 'generation' | 'modification';
    serviceCount: number;
  } | null>(null);
  const validationHandoffShownRef = useRef<typeof validationHandoff>(null);
  const [feedbackPreselectedRating, setFeedbackPreselectedRating] = useState<number | undefined>(undefined);
  const [feedbackFabPulse, setFeedbackFabPulse] = useState(false);
  const [accessIdentity, setAccessIdentity] = useState<AccessIdentity | null>(null);
  const [isAccessManagementOpen, setIsAccessManagementOpen] = useState(false);
  // Counts successful AI generations this session so we can ask for feedback
  // after a "success moment" (the 2nd diagram) rather than nagging up front.
  const generationCountRef = useRef(0);
  const aiPricingRunRef = useRef(0);
  const handleAIGenerateRef = useRef<(
    (
      architecture: any,
      prompt: string,
      autoSnapshot?: boolean,
      preserveExistingLayout?: boolean,
    ) => Promise<void>
  ) | null>(null);
  const feedbackAfterValidationRef = useRef(false);
  const [referenceImageUrl, setReferenceImageUrl] = useState<string | null>(null);
  const [lastReferenceArchitecture, setLastReferenceArchitecture] = useState<ReferenceArchitecture | null>(null);
  const [lastBlueprintArchitecture, setLastBlueprintArchitecture] = useState<BlueprintArchitecture | null>(null);
  const [panelsCollapsedSignal, setPanelsCollapsedSignal] = useState(0);

  useEffect(() => {
    if (nodes.length === 0) {
      setValidationHandoff(null);
      feedbackAfterValidationRef.current = false;
    }
  }, [nodes.length]);
  // Focus mode: hides canvas chrome (side panels via the signal above, plus the
  // "Generated from" prompt banner and the "Generated with" model badge) so only
  // the diagram itself remains. Toggled by the Focus button.
  const [focusMode, setFocusMode] = useState(false);

  useEffect(() => {
    if (!validationHandoff || focusMode || validationHandoffShownRef.current === validationHandoff) return;
    validationHandoffShownRef.current = validationHandoff;
    trackValidationHandoff({ action: 'shown', ...validationHandoff });
  }, [focusMode, validationHandoff]);

  const [isExportMenuOpen, setIsExportMenuOpen] = useState(false);
  const exportMenuRef = useRef<HTMLDivElement | null>(null);

  const [isLayoutMenuOpen, setIsLayoutMenuOpen] = useState(false);
  const layoutMenuRef = useRef<HTMLDivElement | null>(null);

  const [layoutPreset, setLayoutPreset] = useState<LayoutPreset>('flow-lr');
  const [layoutSpacing, setLayoutSpacing] = useState<LayoutSpacing>('comfortable');
  // Default to orthogonal (right-angle) routing — matches the blueprint PNG and
  // draw.io exports, and is the convention for architecture diagrams. The
  // user's choice is remembered across sessions (see persistence effect below).
  const [layoutEdgeStyle, setLayoutEdgeStyle] = useState<LayoutEdgeStyle>(() => {
    try {
      const saved = localStorage.getItem(EDGE_STYLE_STORAGE_KEY);
      if (saved === 'straight' || saved === 'smooth' || saved === 'orthogonal') {
        return saved;
      }
    } catch {
      /* localStorage unavailable — fall through to default */
    }
    return 'orthogonal';
  });
  const [animateConnections, setAnimateConnections] = useState(() =>
    readBooleanPreference(EDGE_ANIMATION_STORAGE_KEY, true)
  );
  const [layoutEmphasizePrimaryPath, setLayoutEmphasizePrimaryPath] = useState(false);
  const [layoutEngine, setLayoutEngine] = useState<LayoutEngineType>('dagre');
  
  const [isBulkSelectMenuOpen, setIsBulkSelectMenuOpen] = useState(false);
  const bulkSelectMenuRef = useRef<HTMLDivElement | null>(null);
  
  const [isStylePresetMenuOpen, setIsStylePresetMenuOpen] = useState(false);
  const stylePresetMenuRef = useRef<HTMLDivElement | null>(null);

  const [isModelSettingsOpen, setIsModelSettingsOpen] = useState(false);
  const modelSettingsRef = useRef<HTMLDivElement | null>(null);
  const [stylePreset, setStylePreset] = useState<'detailed' | 'presentation'>('detailed');

  // Collapse / expand all groups
  const [allGroupsCollapsed, setAllGroupsCollapsed] = useState(false);
  const preCollapseGroupSizes = useRef<Map<string, { width: number; height: number }>>(new Map());



  const [exportHistory, setExportHistory] = useState<ExportHistoryItem[]>(() => {
    try {
      const raw = localStorage.getItem(EXPORT_HISTORY_STORAGE_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed
        .filter((v) => v && typeof v === 'object')
        .slice(0, 25) as ExportHistoryItem[];
    } catch {
      return [];
    }
  });
  
  const reactFlowWrapper = useRef<HTMLDivElement>(null);

  useEffect(() => {
    try {
      localStorage.setItem(EXPORT_HISTORY_STORAGE_KEY, JSON.stringify(exportHistory.slice(0, 25)));
    } catch {
      // ignore
    }
  }, [exportHistory]);

  // Remember the chosen edge style across sessions so a user who prefers smooth
  // (or straight) keeps it instead of reverting to the orthogonal default.
  useEffect(() => {
    try {
      localStorage.setItem(EDGE_STYLE_STORAGE_KEY, layoutEdgeStyle);
    } catch {
      // ignore
    }
  }, [layoutEdgeStyle]);

  useEffect(() => {
    writeLocalStorage(EDGE_ANIMATION_STORAGE_KEY, animateConnections ? '1' : '0');
  }, [animateConnections]);

  useEffect(() => {
    writeLocalStorage(
      TOOLBAR_SECTIONS_STORAGE_KEY,
      JSON.stringify([...collapsedToolbarSections])
    );
  }, [collapsedToolbarSections]);

  // One-time gentle pulse on the feedback button ~15s after load so it earns a
  // glance without looping/nagging. Suppressed once feedback has been given.
  useEffect(() => {
    let alreadyDone = false;
    try {
      alreadyDone = sessionStorage.getItem(FEEDBACK_DONE_KEY) === '1';
    } catch {
      /* sessionStorage unavailable — ignore */
    }
    if (alreadyDone) return;
    const startT = window.setTimeout(() => setFeedbackFabPulse(true), 15000);
    const stopT = window.setTimeout(() => setFeedbackFabPulse(false), 19500);
    return () => {
      window.clearTimeout(startT);
      window.clearTimeout(stopT);
    };
  }, []);

  useEffect(() => {
    let active = true;
    getAccessIdentity()
      .then((identity) => {
        if (active) setAccessIdentity(identity);
      })
      .catch((error) => {
        console.error('[access] failed to read current identity:', error);
      });
    return () => {
      active = false;
    };
  }, []);

  const recordExport = useCallback((kind: ExportHistoryKind, fileName: string) => {
    const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    const item: ExportHistoryItem = { id, kind, fileName, createdAt: Date.now() };
    setExportHistory((prev) => [item, ...prev].slice(0, 25));
  }, []);

  const formatTimeAgo = useCallback((ts: number) => {
    const diffMs = Date.now() - ts;
    const diffSec = Math.max(0, Math.floor(diffMs / 1000));
    if (diffSec < 10) return localize(language, { en: 'just now', ja: 'たった今' });
    if (diffSec < 60) return localize(language, { en: `${diffSec}s ago`, ja: `${diffSec}秒前` });
    const diffMin = Math.floor(diffSec / 60);
    if (diffMin < 60) return localize(language, { en: `${diffMin}m ago`, ja: `${diffMin}分前` });
    const diffHr = Math.floor(diffMin / 60);
    if (diffHr < 24) return localize(language, { en: `${diffHr}h ago`, ja: `${diffHr}時間前` });
    const diffDay = Math.floor(diffHr / 24);
    return localize(language, { en: `${diffDay}d ago`, ja: `${diffDay}日前` });
  }, [language]);

  useEffect(() => {
    const handlePointerDown = (e: PointerEvent) => {
      if (!isExportMenuOpen && !isLayoutMenuOpen && !isBulkSelectMenuOpen && !isStylePresetMenuOpen && !isModelSettingsOpen) return;
      const target = e.target as unknown as globalThis.Node | null;
      if (!target) return;

      if (isExportMenuOpen && exportMenuRef.current && !exportMenuRef.current.contains(target)) {
        setIsExportMenuOpen(false);
      }

      if (isLayoutMenuOpen && layoutMenuRef.current && !layoutMenuRef.current.contains(target)) {
        setIsLayoutMenuOpen(false);
      }

      if (isBulkSelectMenuOpen && bulkSelectMenuRef.current && !bulkSelectMenuRef.current.contains(target)) {
        setIsBulkSelectMenuOpen(false);
      }

      if (isStylePresetMenuOpen && stylePresetMenuRef.current && !stylePresetMenuRef.current.contains(target)) {
        setIsStylePresetMenuOpen(false);
      }

      if (isModelSettingsOpen && modelSettingsRef.current && !modelSettingsRef.current.contains(target)) {
        setIsModelSettingsOpen(false);
      }
    };

    const handleKeyDown = (e: KeyboardEvent) => {
      if (!isExportMenuOpen && !isLayoutMenuOpen && !isBulkSelectMenuOpen && !isStylePresetMenuOpen && !isModelSettingsOpen) return;
      if (e.key === 'Escape') {
        setIsExportMenuOpen(false);
        setIsLayoutMenuOpen(false);
        setIsBulkSelectMenuOpen(false);
        setIsStylePresetMenuOpen(false);
        setIsModelSettingsOpen(false);
      }
    };

    document.addEventListener('pointerdown', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('pointerdown', handlePointerDown);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [isExportMenuOpen, isLayoutMenuOpen, isBulkSelectMenuOpen, isStylePresetMenuOpen, isModelSettingsOpen]);

  useLayoutEffect(() => {
    if (!isExportMenuOpen && !isLayoutMenuOpen && !isBulkSelectMenuOpen && !isStylePresetMenuOpen && !isModelSettingsOpen) {
      return;
    }

    let frame = 0;
    let settleTimer = 0;
    const fitOpenMenus = () => {
      document
        .querySelectorAll<HTMLElement>('.app-header .toolbar-dropdown-menu')
        .forEach(fitToolbarMenuToViewport);
    };
    const scheduleFitOpenMenus = () => {
      window.cancelAnimationFrame(frame);
      window.clearTimeout(settleTimer);
      frame = window.requestAnimationFrame(fitOpenMenus);
      settleTimer = window.setTimeout(fitOpenMenus, 250);
    };

    fitOpenMenus();
    window.addEventListener('resize', scheduleFitOpenMenus);
    window.visualViewport?.addEventListener('resize', scheduleFitOpenMenus);
    window.visualViewport?.addEventListener('scroll', scheduleFitOpenMenus);
    const header = document.querySelector('.app-header');
    const resizeObserver = header && typeof ResizeObserver !== 'undefined'
      ? new ResizeObserver(scheduleFitOpenMenus)
      : null;
    if (header && resizeObserver) resizeObserver.observe(header);
    return () => {
      window.cancelAnimationFrame(frame);
      window.clearTimeout(settleTimer);
      window.removeEventListener('resize', scheduleFitOpenMenus);
      window.visualViewport?.removeEventListener('resize', scheduleFitOpenMenus);
      window.visualViewport?.removeEventListener('scroll', scheduleFitOpenMenus);
      resizeObserver?.disconnect();
    };
  }, [isExportMenuOpen, isLayoutMenuOpen, isBulkSelectMenuOpen, isStylePresetMenuOpen, isModelSettingsOpen]);

  // Keyboard shortcuts: Delete and Ctrl+D (duplicate)
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') cancelPendingPricingEditorOpen();

      const target = e.target as HTMLElement;
      if (
        pricingEditorNodeId
        || target.tagName === 'INPUT'
        || target.tagName === 'TEXTAREA'
        || target.tagName === 'SELECT'
        || target.tagName === 'BUTTON'
        || target.isContentEditable
        || target.closest('[role="dialog"], [role="menu"]')
      ) {
        return;
      }

      // Delete key - remove selected nodes and edges
      if (e.key === 'Delete' || e.key === 'Backspace') {
        const selectedNodes = nodes.filter(n => n.selected);
        const selectedEdges = edges.filter(e => e.selected);
        
        if (selectedNodes.length > 0 || selectedEdges.length > 0) {
          e.preventDefault();
          
          // Remove selected nodes
          if (selectedNodes.length > 0) {
            const nodeIdsToRemove = selectedNodes.map(n => n.id);
            deleteCanvasNodes(nodeIdsToRemove);
          }
          
          // Remove selected edges
          if (selectedEdges.length > 0) {
            const edgeIdsToRemove = selectedEdges.map(e => e.id);
            setEdges(eds => eds.filter(e => !edgeIdsToRemove.includes(e.id)));
          }
        }
      }

      // Ctrl+D or Cmd+D - duplicate selected nodes
      if ((e.ctrlKey || e.metaKey) && e.key === 'd') {
        const selectedNodes = nodes.filter(n => n.selected);
        
        if (selectedNodes.length > 0) {
          e.preventDefault();
          
          const duplicatedNodes = selectedNodes.map(node => {
            const newId = `${Date.now()}-${Math.random()}`;
            return {
              ...node,
              id: newId,
              position: {
                x: node.position.x + 50, // Offset by 50px
                y: node.position.y + 50,
              },
              selected: true, // Select the new nodes
            };
          });
          
          // Deselect original nodes
          setNodes(nds => [
            ...nds.map(n => ({ ...n, selected: false })),
            ...duplicatedNodes
          ]);
        }
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [
    nodes,
    edges,
    setNodes,
    setEdges,
    deleteCanvasNodes,
    pricingEditorNodeId,
    cancelPendingPricingEditorOpen,
  ]);

  // Keep edge rendering style in sync even without re-layout.
  useEffect(() => {
    setEdges((eds) =>
      eds.map((e) => ({
        ...e,
        data: { ...(e.data ?? {}), pathStyle: layoutEdgeStyle },
      }))
    );
  }, [layoutEdgeStyle, setEdges]);

  const addGroupBox = useCallback(() => {
    const newNode: Node = {
      id: `group-${Date.now()}`,
      type: 'groupNode',
      position: { x: 250, y: 150 },
      data: { 
        label: 'Group Label',
      },
      style: {
        width: 400,
        height: 300,
      },
    };
    setNodes((nds) => nds.concat(newNode));
  }, [setNodes]);

  // Apply dark mode class to body and persist preference
  useEffect(() => {
    if (isDarkMode) {
      document.body.classList.add('dark-mode');
    } else {
      document.body.classList.remove('dark-mode');
    }
    writeLocalStorage('darkMode', JSON.stringify(isDarkMode));
  }, [isDarkMode]);

  // Preload pricing data on mount
  useEffect(() => {
    preloadCommonServices().catch(err => 
      console.warn('Failed to preload regional pricing:', err)
    );
    prefetchCommonServices(getActiveRegion()).catch(err =>
      console.warn('Failed to prefetch API pricing:', err)
    );
  }, []);

  // Recalculate total cost whenever nodes change
  useEffect(() => {
    const breakdown = calculateCostBreakdown(nodes, undefined, pricingMode);
    setTotalMonthlyCost(breakdown.totalMonthlyCost);
  }, [nodes, pricingMode]);

  // Handle region change
  //
  // Pricing lookups are async, so the node list can change while they are in
  // flight (drags, additions, deletions, or another region change). Replacing
  // the whole array with a pre-await snapshot would silently discard those
  // edits, so the results are merged into the current nodes by id and stale
  // runs are dropped.
  const regionPricingRunRef = useRef(0);
  const handleRegionChange = useCallback(async (region: AzureRegion) => {
    cancelPendingPricingEditorOpen();
    console.log(`🌍 Region changed to ${region}, updating all node pricing...`);
    const runId = ++regionPricingRunRef.current;

    type RegionPricingResult = {
      nodeId: string;
      expectedFingerprint: string;
      pricing: NodePricingConfig | null;
      failed: boolean;
    };

    const refreshNodePricing = async (node: Node): Promise<RegionPricingResult | null> => {
      if (node.type !== 'azureNode') return null;

      const currentPricing = node.data.pricing as NodePricingConfig | undefined;
      const serviceType = String(node.data.serviceName || node.data.label || '');
      if (!serviceType) return null;
      const expectedFingerprint = nodePricingFingerprint(node);
      try {
        const newPricing = currentPricing
          ? currentPricing.isCustom
            ? {
                ...currentPricing,
                region,
                lastUpdated: new Date().toISOString(),
              }
            : await updateNodePricing(
                serviceType,
                currentPricing,
                currentPricing.tierId || currentPricing.skuName || currentPricing.tier,
                currentPricing.quantity,
                region,
              )
          : await initializeNodePricing(serviceType, region);

        return {
          nodeId: node.id,
          expectedFingerprint,
          pricing: newPricing,
          failed: false,
        };
      } catch (error) {
        console.error(`Failed to refresh pricing for ${serviceType}:`, error);
        return {
          nodeId: node.id,
          expectedFingerprint,
          pricing: null,
          failed: true,
        };
      }
    };

    const initialResults = (await Promise.all(nodes.map(refreshNodePricing)))
      .filter((result): result is RegionPricingResult => result !== null);

    if (runId !== regionPricingRunRef.current) return;

    // If a user edited a node while regional prices were loading, reprice the
    // latest configuration once instead of overwriting it with the stale
    // result calculated from the original snapshot.
    const resultsByNodeId = new Map(initialResults.map(result => [result.nodeId, result]));
    const retryNodes = latestNodesRef.current.filter((node) => {
      const result = resultsByNodeId.get(node.id);
      return result && nodePricingFingerprint(node) !== result.expectedFingerprint;
    });
    if (retryNodes.length > 0) {
      const retryResults = (await Promise.all(retryNodes.map(refreshNodePricing)))
        .filter((result): result is RegionPricingResult => result !== null);
      retryResults.forEach(result => resultsByNodeId.set(result.nodeId, result));
    }

    if (runId !== regionPricingRunRef.current) return;

    const results = [...resultsByNodeId.values()];
    const currentNodesById = new Map(latestNodesRef.current.map(node => [node.id, node]));
    const skippedForConcurrentEdits = results.filter((result) => {
      const currentNode = currentNodesById.get(result.nodeId);
      return currentNode
        && nodePricingFingerprint(currentNode) !== result.expectedFingerprint;
    }).length;

    setNodes((currentNodes) => currentNodes.map((node) => {
      const result = resultsByNodeId.get(node.id);
      if (
        !result?.pricing
        || nodePricingFingerprint(node) !== result.expectedFingerprint
      ) {
        return node;
      }
      return { ...node, data: { ...node.data, pricing: result.pricing } };
    }));

    const failedCount = results.filter(result => result.failed).length;
    if (failedCount > 0 || skippedForConcurrentEdits > 0) {
      const failedMessage = failedCount > 0
        ? localize(language, {
            en: `${failedCount} service price${failedCount === 1 ? '' : 's'} could not be matched in the selected region. Existing estimates keep their original region.`,
            ja: `${failedCount} 件のサービス価格を選択したリージョンで特定できなかったため、既存の見積もりと元のリージョンを保持しました。`,
          })
        : '';
      const skippedMessage = skippedForConcurrentEdits > 0
        ? localize(language, {
            en: `${skippedForConcurrentEdits} concurrent pricing edit${skippedForConcurrentEdits === 1 ? ' was' : 's were'} preserved.`,
            ja: `更新中に行われた ${skippedForConcurrentEdits} 件の価格編集を保持しました。`,
          })
        : '';
      alert([failedMessage, skippedMessage].filter(Boolean).join('\n'));
    }
  }, [language, nodes, setNodes]);

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (isDraggingBanner) {
        setPromptBannerPosition({
          x: e.clientX - dragOffset.x,
          y: e.clientY - dragOffset.y,
        });
      }
    };

    const handleMouseUp = () => {
      setIsDraggingBanner(false);
    };

    if (isDraggingBanner) {
      document.addEventListener('mousemove', handleMouseMove);
      document.addEventListener('mouseup', handleMouseUp);
    }

    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isDraggingBanner, dragOffset]);

  const handleEdgeLabelChange = useCallback((edgeId: string, newLabel: string) => {
    setEdges((eds) =>
      eds.map((edge) => {
        if (edge.id === edgeId) {
          return {
            ...edge,
            label: newLabel,
            data: {
              ...edge.data,
              label: newLabel,
              onLabelChange: handleEdgeLabelChange,
            },
          };
        }
        return edge;
      })
    );
  }, [setEdges]);

  const handleEdgeLabelOffsetChange = useCallback((edgeId: string, offsetX: number, offsetY: number) => {
    setEdges((eds) =>
      eds.map((edge) => {
        if (edge.id === edgeId) {
          return {
            ...edge,
            data: { 
              ...edge.data, 
              labelOffsetX: offsetX, 
              labelOffsetY: offsetY,
              onLabelChange: handleEdgeLabelChange,
              onLabelOffsetChange: handleEdgeLabelOffsetChange,
            },
          };
        }
        return edge;
      })
    );
  }, [setEdges, handleEdgeLabelChange]);

  const normalizeRestoredEdges = useCallback((restoredEdges: unknown[]): Edge[] => {
    const SRC_FIX: Record<string, string> = { top: 'top-source', left: 'left-source' };
    const TGT_FIX: Record<string, string> = { bottom: 'bottom-target', right: 'right-target' };
    return restoredEdges.map((edge) => {
      if (!edge || typeof edge !== 'object' || Array.isArray(edge)) {
        throw new Error('Invalid edge in diagram payload');
      }
      const next = { ...(edge as Edge) };
      if (typeof next.sourceHandle === 'string' && SRC_FIX[next.sourceHandle]) {
        next.sourceHandle = SRC_FIX[next.sourceHandle];
      }
      if (typeof next.targetHandle === 'string' && TGT_FIX[next.targetHandle]) {
        next.targetHandle = TGT_FIX[next.targetHandle];
      }
      const baseFlowAnimated = Boolean(next.data?.baseFlowAnimated ?? next.data?.flowAnimated ?? true);
      const edgeAnimationPreference = typeof next.data?.flowAnimated === 'boolean'
        ? next.data.flowAnimated
        : baseFlowAnimated;
      next.data = {
        ...next.data,
        baseFlowAnimated,
        flowAnimated: animateConnections && edgeAnimationPreference,
        pathStyle: normalizeLayoutEdgeStyle(next.data?.pathStyle ?? layoutEdgeStyle),
        onLabelChange: handleEdgeLabelChange,
        onLabelOffsetChange: handleEdgeLabelOffsetChange,
      };
      return next;
    });
  }, [handleEdgeLabelChange, handleEdgeLabelOffsetChange, animateConnections, layoutEdgeStyle]);

  const onConnect = useCallback(
    (params: Connection | Edge) => setEdges((eds) => addEdge({ 
      ...params, 
      animated: false,
      type: 'editableEdge',
      label: '',
      markerEnd: { type: MarkerType.ArrowClosed, color: '#0078d4' },
      labelStyle: { fontSize: 14, fill: '#333', fontWeight: 'bold' },
      labelBgStyle: { fill: 'white', fillOpacity: 0.9, stroke: '#000', strokeWidth: 1.5 },
      data: {
        onLabelChange: handleEdgeLabelChange,
        onLabelOffsetChange: handleEdgeLabelOffsetChange,
        connectionType: 'sync',
        direction: 'forward',
        baseFlowAnimated: true,
        flowAnimated: animateConnections,
        flowMode: 'directional',
        pathStyle: layoutEdgeStyle,
        labelOffsetX: 0,
        labelOffsetY: 0,
      },
    }, eds)),
    [setEdges, handleEdgeLabelChange, handleEdgeLabelOffsetChange, animateConnections, layoutEdgeStyle]
  );

  // Bulk select operations
  const selectAllNodesOfType = useCallback((serviceType: string) => {
    setNodes((nds) => 
      nds.map(node => ({
        ...node,
        selected: node.type === 'azureNode' && node.data.label === serviceType
      }))
    );
    setIsBulkSelectMenuOpen(false);
  }, [setNodes]);

  const selectAllNodes = useCallback(() => {
    setNodes((nds) => nds.map(node => ({ ...node, selected: true })));
    setIsBulkSelectMenuOpen(false);
  }, [setNodes]);

  const deselectAll = useCallback(() => {
    setNodes((nds) => nds.map(node => ({ ...node, selected: false })));
    setEdges((eds) => eds.map(edge => ({ ...edge, selected: false })));
    setIsBulkSelectMenuOpen(false);
  }, [setNodes, setEdges]);

  // Toggle collapse / expand all group nodes
  const toggleCollapseAllGroups = useCallback(() => {
    if (!allGroupsCollapsed) {
      // Save current sizes before collapsing
      const sizeMap = new Map<string, { width: number; height: number }>();
      nodes.forEach(n => {
        if (n.type === 'groupNode') {
          sizeMap.set(n.id, {
            width: (n.style?.width as number) || (n.width as number) || 400,
            height: (n.style?.height as number) || (n.height as number) || 300,
          });
        }
      });
      preCollapseGroupSizes.current = sizeMap;

      // Collapse all groups to fit content
      const collapsed = fitAllGroupsToContent(nodes);
      setNodes(collapsed);
      setAllGroupsCollapsed(true);

      // Zoom out to show the full picture
      setTimeout(() => {
        reactFlowInstance?.fitView?.({ padding: 0.3, duration: 300 });
      }, 50);
    } else {
      // Restore saved sizes
      const sizeMap = preCollapseGroupSizes.current;
      setNodes(nds =>
        nds.map(n => {
          if (n.type === 'groupNode' && sizeMap.has(n.id)) {
            const { width, height } = sizeMap.get(n.id)!;
            return { ...n, style: { ...n.style, width, height } };
          }
          return n;
        })
      );
      setAllGroupsCollapsed(false);
      preCollapseGroupSizes.current = new Map();

      setTimeout(() => {
        reactFlowInstance?.fitView?.({ padding: 0.2, duration: 300 });
      }, 50);
    }
  }, [allGroupsCollapsed, nodes, setNodes, reactFlowInstance]);

  // Get unique service types from current diagram
  const getServiceTypes = useCallback(() => {
    const types = new Set<string>();
    nodes.forEach((node) => {
      if (node.type === 'azureNode' && node.data.label) {
        types.add(node.data.label);
      }
    });
    return Array.from(types).sort();
  }, [nodes]);

  // Style preset functions
  const applyStylePreset = useCallback((preset: 'detailed' | 'presentation') => {
    setStylePreset(preset);

    // Update nodes with style data
    setNodes((nds) =>
      nds.map((node) => ({
        ...node,
        data: {
          ...node.data,
          stylePreset: preset,
        },
      }))
    );

    // Update edges based on preset (non-destructive: don't wipe labels)
    setEdges((eds) =>
      eds.map((edge) => {
        const nextStyle: any = { ...(edge.style ?? {}) };
        const nextLabelStyle: any = { ...(edge.labelStyle ?? {}) };
        const nextLabelBgStyle: any = { ...(edge.labelBgStyle ?? {}) };

        switch (preset) {
          case 'presentation':
            nextStyle.strokeWidth = 2.5;
            delete nextStyle.strokeDasharray;
            nextLabelStyle.opacity = 1;
            nextLabelStyle.fontSize = 15;
            nextLabelStyle.fontWeight = '600';
            nextLabelBgStyle.fillOpacity = 0.95;
            nextLabelBgStyle.strokeWidth = 2;
            nextLabelBgStyle.rx = 6;
            break;
          case 'detailed':
          default:
            delete nextStyle.strokeWidth;
            delete nextStyle.strokeDasharray;
            nextLabelStyle.opacity = 1;
            nextLabelBgStyle.fillOpacity = 0.9;
            nextLabelBgStyle.strokeWidth = 1.5;
            break;
        }

        return {
          ...edge,
          style: nextStyle,
          labelStyle: nextLabelStyle,
          labelBgStyle: nextLabelBgStyle,
        };
      })
    );

    setIsStylePresetMenuOpen(false);
  }, [setEdges, setNodes]);

  const applyLayout = useCallback(async () => {
    const selectedAzureNodeId = nodes.find((n) => n.type === 'azureNode' && (n as any).selected)?.id;
    const shouldEmphasize =
      layoutEmphasizePrimaryPath && (layoutPreset === 'flow-lr' || layoutPreset === 'flow-tb');

    const result = await applyLayoutPreset(nodes as any, edges as any, {
      preset: layoutPreset,
      spacing: layoutSpacing,
      edgeStyle: layoutEdgeStyle,
      emphasizePrimaryPath: shouldEmphasize,
      selectedNodeId: selectedAzureNodeId,
      layoutEngine,
    });

    setNodes(result.nodes as any);
    setEdges(result.edges as any);

    requestAnimationFrame(() => {
      reactFlowInstance?.fitView?.({ padding: 0.2, duration: 250 });
    });
  }, [
    nodes,
    edges,
    layoutPreset,
    layoutSpacing,
    layoutEdgeStyle,
    layoutEmphasizePrimaryPath,
    layoutEngine,
    reactFlowInstance,
    setNodes,
    setEdges,
  ]);

  const layoutPresetLabel: Record<LayoutPreset, string> = {
    'flow-lr': 'Flow (L→R)',
    'flow-tb': 'Flow (Top→Bottom)',
    swimlanes: 'Swimlanes by Group',
    radial: 'Radial',
  };

  const onReconnect = useCallback(
    (oldEdge: Edge, newConnection: Connection) => {
      setEdges((eds) => {
        // Remove the old edge and add the new connection
        const filtered = eds.filter(e => e.id !== oldEdge.id);
        return addEdge({
          ...newConnection,
          animated: false,
          type: oldEdge.type,
          label: oldEdge.label,
          markerEnd: oldEdge.markerEnd,
          markerStart: (oldEdge as any).markerStart,
          labelStyle: oldEdge.labelStyle,
          labelBgStyle: oldEdge.labelBgStyle,
          data: oldEdge.data,
        }, filtered);
      });
    },
    [setEdges]
  );

  const closeNodeContextMenu = useCallback(() => {
    setNodeContextMenu(null);
    const returnFocus = nodeContextMenuReturnFocusRef.current;
    nodeContextMenuReturnFocusRef.current = null;
    window.requestAnimationFrame(() => {
      if (returnFocus?.isConnected) returnFocus.focus();
    });
  }, []);

  const dismissNodeContextMenu = useCallback(() => {
    nodeContextMenuReturnFocusRef.current = null;
    setNodeContextMenu(null);
  }, []);

  const closeEdgeContextMenu = useCallback(() => {
    setEdgeContextMenu(null);
  }, []);

  const closeCanvasContextMenus = useCallback(() => {
    cancelPendingPricingEditorOpen();
    closeNodeContextMenu();
    closeEdgeContextMenu();
  }, [cancelPendingPricingEditorOpen, closeEdgeContextMenu, closeNodeContextMenu]);

  const dismissCanvasContextMenus = useCallback(() => {
    cancelPendingPricingEditorOpen();
    dismissNodeContextMenu();
    closeEdgeContextMenu();
  }, [cancelPendingPricingEditorOpen, closeEdgeContextMenu, dismissNodeContextMenu]);

  const onNodeContextMenu = useCallback((event: React.MouseEvent, node: Node) => {
    cancelPendingPricingEditorOpen();
    event.preventDefault();
    event.stopPropagation();
    const position = clampContextMenuPosition(
      event.clientX,
      event.clientY,
      NODE_CONTEXT_MENU_WIDTH,
      NODE_CONTEXT_MENU_HEIGHT,
    );
    if (!node.selected) {
      setNodes((currentNodes) => currentNodes.map(currentNode => ({
        ...currentNode,
        selected: currentNode.id === node.id,
      })));
      setEdges((currentEdges) => currentEdges.map(edge => ({ ...edge, selected: false })));
    }
    const eventTarget = event.target as HTMLElement;
    nodeContextMenuReturnFocusRef.current = eventTarget.closest<HTMLElement>('.react-flow__node')
      || (document.activeElement instanceof HTMLElement ? document.activeElement : null);
    setEdgeContextMenu(null);
    setNodeContextMenu({ ...position, nodeId: node.id });
  }, [cancelPendingPricingEditorOpen, setEdges, setNodes]);

  const onEdgeContextMenu = useCallback((event: React.MouseEvent, edge: Edge) => {
    cancelPendingPricingEditorOpen();
    event.preventDefault();
    const position = clampContextMenuPosition(
      event.clientX,
      event.clientY,
      EDGE_CONTEXT_MENU_WIDTH,
      EDGE_CONTEXT_MENU_HEIGHT,
    );
    dismissNodeContextMenu();
    setEdgeContextMenu({ ...position, edgeId: edge.id });
  }, [cancelPendingPricingEditorOpen, dismissNodeContextMenu]);

  const duplicateServiceNode = useCallback((nodeId: string) => {
    setNodes((currentNodes) => {
      const source = currentNodes.find(node => node.id === nodeId);
      if (!source || source.type !== 'azureNode') return currentNodes;
      const duplicateId = `${source.id}-copy-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      const duplicate: Node = {
        ...source,
        id: duplicateId,
        position: {
          x: source.position.x + 40,
          y: source.position.y + 40,
        },
        positionAbsolute: undefined,
        selected: true,
        dragging: false,
        data: { ...source.data },
      };
      return [
        ...currentNodes.map(node => ({ ...node, selected: false })),
        duplicate,
      ];
    });
    setEdges((currentEdges) => currentEdges.map(edge => ({ ...edge, selected: false })));
    closeNodeContextMenu();
  }, [closeNodeContextMenu, setEdges, setNodes]);

  const editContextNodePricing = useCallback(async (nodeId: string) => {
    const runId = ++pricingEditorOpenRunRef.current;
    const editorStateVersion = getNodePricingEditorStateVersion();
    const node = latestNodesRef.current.find(candidate => candidate.id === nodeId);
    dismissNodeContextMenu();
    if (!node || node.type !== 'azureNode') return;

    const storedPricing = node.data?.pricing as NodePricingConfig | undefined;
    if (storedPricing) {
      setPricingEditorDraft(null);
      openNodePricingEditor(nodeId);
      return;
    }

    const serviceType = String(node.data?.serviceName || node.data?.label || '');
    const region = getActiveRegion();
    const initializedPricing = serviceType && hasPricingData(serviceType)
      ? await initializeNodePricing(serviceType, region)
      : null;
    if (
      runId !== pricingEditorOpenRunRef.current
      || editorStateVersion !== getNodePricingEditorStateVersion()
      || getActiveRegion() !== region
    ) return;

    const latestNode = latestNodesRef.current.find(candidate => candidate.id === nodeId);
    const latestServiceType = String(
      latestNode?.data?.serviceName || latestNode?.data?.label || '',
    );
    if (!latestNode || latestNode.type !== 'azureNode' || latestServiceType !== serviceType) return;

    const latestPricing = latestNode.data?.pricing as NodePricingConfig | undefined;
    setPricingEditorDraft(latestPricing
      ? null
      : {
          nodeId,
          pricing: initializedPricing ?? createCustomPricingDraft(region),
        });
    openNodePricingEditor(nodeId);
  }, [dismissNodeContextMenu]);

  const closePricingEditor = useCallback(() => {
    cancelPendingPricingEditorOpen();
    setPricingEditorDraft(null);
    closeNodePricingEditor();
  }, [cancelPendingPricingEditorOpen]);

  const fitContextGroupToContent = useCallback((groupId: string) => {
    setNodes((currentNodes) => fitGroupToContent(currentNodes, groupId) ?? currentNodes);
    closeNodeContextMenu();
  }, [closeNodeContextMenu, setNodes]);

  const deleteContextNode = useCallback((nodeId: string) => {
    deleteCanvasNodes([nodeId]);
    closeNodeContextMenu();
  }, [closeNodeContextMenu, deleteCanvasNodes]);

  const deleteContextGroupWithContents = useCallback((groupId: string) => {
    const deletedIds = collectNodeAndDescendantIds(nodes, [groupId]);
    const containedCount = Math.max(0, deletedIds.size - 1);
    const confirmed = window.confirm(localize(language, {
      en: `Delete this layer and ${containedCount} contained item${containedCount === 1 ? '' : 's'}? This cannot be undone.`,
      ja: `このレイヤーと、その中の${containedCount}件の項目を削除しますか？この操作は元に戻せません。`,
    }));
    if (!confirmed) return;
    deleteCanvasNodes(deletedIds, false);
    closeNodeContextMenu();
  }, [closeNodeContextMenu, deleteCanvasNodes, language, nodes]);

  useLayoutEffect(() => {
    if (!nodeContextMenu) return;
    nodeContextMenuRef.current
      ?.querySelector<HTMLButtonElement>('[role="menuitem"]:not(:disabled)')
      ?.focus();
  }, [nodeContextMenu]);

  const handleNodeContextMenuKeyDown = useCallback((
    event: React.KeyboardEvent<HTMLDivElement>,
  ) => {
    const items = Array.from(
      event.currentTarget.querySelectorAll<HTMLButtonElement>('[role="menuitem"]:not(:disabled)'),
    );
    if (items.length === 0) return;
    const currentIndex = items.indexOf(document.activeElement as HTMLButtonElement);

    let nextIndex: number | null = null;
    if (event.key === 'ArrowDown' || (event.key === 'Tab' && !event.shiftKey)) {
      nextIndex = currentIndex < 0 ? 0 : (currentIndex + 1) % items.length;
    } else if (event.key === 'ArrowUp' || (event.key === 'Tab' && event.shiftKey)) {
      nextIndex = currentIndex <= 0 ? items.length - 1 : currentIndex - 1;
    } else if (event.key === 'Home') {
      nextIndex = 0;
    } else if (event.key === 'End') {
      nextIndex = items.length - 1;
    } else if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      closeNodeContextMenu();
      return;
    }

    if (nextIndex !== null) {
      event.preventDefault();
      event.stopPropagation();
      items[nextIndex].focus();
    }
  }, [closeNodeContextMenu]);

  useEffect(() => {
    if (!nodeContextMenu && !edgeContextMenu) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') closeCanvasContextMenus();
    };
    const handleResize = () => closeCanvasContextMenus();
    document.addEventListener('keydown', handleKeyDown);
    window.addEventListener('resize', handleResize);
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('resize', handleResize);
    };
  }, [closeCanvasContextMenus, edgeContextMenu, nodeContextMenu]);

  const setConnectionAnimations = useCallback((enabled: boolean) => {
    setAnimateConnections(enabled);
    setEdges((currentEdges) => currentEdges.map((edge) => {
      const baseFlowAnimated = Boolean(edge.data?.baseFlowAnimated ?? edge.data?.flowAnimated ?? true);
      return {
        ...edge,
        animated: false,
        data: {
          ...edge.data,
          baseFlowAnimated,
          flowAnimated: enabled && baseFlowAnimated,
          onLabelChange: handleEdgeLabelChange,
          onLabelOffsetChange: handleEdgeLabelOffsetChange,
        },
      };
    }));
  }, [handleEdgeLabelChange, handleEdgeLabelOffsetChange, setEdges]);

  const toggleEdgeAnimation = useCallback((edgeId: string) => {
    setEdges((currentEdges) => currentEdges.map((edge) => {
      if (edge.id !== edgeId) return edge;
      const flowAnimated = !Boolean(edge.data?.flowAnimated);
      return {
        ...edge,
        animated: false,
        data: {
          ...edge.data,
          baseFlowAnimated: flowAnimated
            ? true
            : Boolean(edge.data?.baseFlowAnimated ?? true),
          flowAnimated,
          onLabelChange: handleEdgeLabelChange,
          onLabelOffsetChange: handleEdgeLabelOffsetChange,
        },
      };
    }));
    closeEdgeContextMenu();
  }, [closeEdgeContextMenu, handleEdgeLabelChange, handleEdgeLabelOffsetChange, setEdges]);

  const setEdgeDirection = useCallback((edgeId: string, direction: 'forward' | 'reverse' | 'bidirectional') => {
    setEdges((eds) => eds.map((edge) => {
      if (edge.id === edgeId) {
        let markerEnd: any = undefined;
        let markerStart: any = undefined;
        const baseFlowAnimated = Boolean(edge.data?.baseFlowAnimated ?? edge.data?.flowAnimated ?? true);
        const flowAnimated = animateConnections && baseFlowAnimated;
        const flowMode = direction === 'bidirectional' ? 'pulse' : 'directional';
        
        switch (direction) {
          case 'forward':
            markerEnd = { type: MarkerType.ArrowClosed, color: '#0078d4' };
            break;
          case 'reverse':
            markerStart = { type: MarkerType.ArrowClosed, color: '#0078d4' };
            break;
          case 'bidirectional':
            markerEnd = { type: MarkerType.ArrowClosed, color: '#0078d4' };
            markerStart = { type: MarkerType.ArrowClosed, color: '#0078d4' };
            break;
        }
        
        return {
          ...edge,
          markerEnd,
          markerStart,
          animated: false,
          data: { 
            ...edge.data, 
            direction, 
            baseFlowAnimated, 
            flowAnimated, 
            flowMode,
            onLabelChange: handleEdgeLabelChange,
            onLabelOffsetChange: handleEdgeLabelOffsetChange,
          }
        };
      }
      return edge;
    }));
    closeEdgeContextMenu();
  }, [animateConnections, setEdges, closeEdgeContextMenu]);

  const onDragOver = useCallback((event: React.DragEvent) => {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
  }, []);

  const addServiceNodeAtPosition = useCallback((
    service: { iconPath: string; iconName: string; serviceName: string; category?: string },
    position: Node['position'],
  ) => {
    if (!reactFlowInstance) return;

    let parentGroup: Node | undefined;
    for (const node of reactFlowInstance.getNodes()) {
      if (node.type !== 'groupNode') continue;
      const groupWidth = (node.style?.width as number) || node.width || 400;
      const groupHeight = (node.style?.height as number) || node.height || 300;
      if (
        position.x >= node.position.x
        && position.x <= node.position.x + groupWidth
        && position.y >= node.position.y
        && position.y <= node.position.y + groupHeight
      ) {
        parentGroup = node;
        break;
      }
    }

    const newNode: Node = {
      id: `service-${globalThis.crypto.randomUUID()}`,
      type: 'azureNode',
      position: parentGroup
        ? {
            x: position.x - parentGroup.position.x,
            y: position.y - parentGroup.position.y,
          }
        : position,
      data: {
        label: service.iconName,
        serviceName: service.serviceName,
        category: service.category,
        iconPath: service.iconPath,
      },
      parentNode: parentGroup?.id,
      extent: parentGroup ? 'parent' : undefined,
    };

    setNodes((current) => current.concat(newNode));

    const currentRegion = getActiveRegion();
    void initializeNodePricing(service.serviceName, currentRegion)
      .then((pricing) => {
        if (!pricing) return;
        setNodes((current) => current.map((node) => (
          node.id === newNode.id && !node.data.pricing
            ? { ...node, data: { ...node.data, pricing } }
            : node
        )));
      })
      .catch((error) => console.warn('Failed to initialize pricing:', error));
  }, [reactFlowInstance, setNodes]);

  const handleAddService = useCallback((icon: AzureIcon) => {
    if (!reactFlowInstance || !reactFlowWrapper.current) return;

    const bounds = reactFlowWrapper.current.getBoundingClientRect();
    const serviceCount = reactFlowInstance.getNodes().filter((node: Node) => node.type === 'azureNode').length;
    const slot = serviceCount % 9;
    const column = (slot % 3) - 1;
    const row = Math.floor(slot / 3) - 1;
    const spacingX = Math.min(180, bounds.width / 4);
    const spacingY = Math.min(140, bounds.height / 4);
    const position = reactFlowInstance.screenToFlowPosition({
      x: bounds.left + (bounds.width / 2) + (column * spacingX),
      y: bounds.top + (bounds.height / 2) + (row * spacingY),
    });

    addServiceNodeAtPosition({
      iconPath: icon.path,
      iconName: icon.name,
      serviceName: icon.serviceName,
      category: icon.category,
    }, position);
  }, [addServiceNodeAtPosition, reactFlowInstance]);

  // Handle node deletion - convert child nodes to absolute positions when parent group is deleted
  const onNodesDelete = useCallback((deleted: any[]) => {
    const deletedGroupIds = deleted.filter(n => n.type === 'groupNode').map(n => n.id);
    
    if (deletedGroupIds.length > 0) {
      setNodes((nds) => detachChildrenFromGroups(nds, deletedGroupIds));
    }
  }, [setNodes]);

  const onDrop = useCallback(
    (event: React.DragEvent) => {
      event.preventDefault();

      if (!reactFlowInstance) return;

      const type = event.dataTransfer.getData('application/reactflow');
      const iconPath = event.dataTransfer.getData('iconPath');
      const iconName = event.dataTransfer.getData('iconName');
      const serviceName = event.dataTransfer.getData('iconServiceName') || iconName;
      const category = event.dataTransfer.getData('iconCategory') || undefined;

      if (type !== 'azureNode' || !iconPath || !iconName) return;

      const position = reactFlowInstance.screenToFlowPosition({
        x: event.clientX,
        y: event.clientY,
      });

      addServiceNodeAtPosition({ iconPath, iconName, serviceName, category }, position);
    },
    [addServiceNodeAtPosition, reactFlowInstance]
  );

  const exportDiagram = useCallback(async () => {
    if (!reactFlowWrapper.current || !reactFlowInstance) {
      return;
    }

    // Fit all nodes into view with no animation for immediate rendering
    reactFlowInstance.fitView({ padding: 0.2, duration: 0 });

    // Wait for fitView to settle, then capture
    setTimeout(async () => {
      try {
        const dataUrl = await captureDiagramAsPng(reactFlowWrapper.current as HTMLElement, {
          backgroundColor: '#ffffff',
        });

        const res = await fetch(dataUrl);
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        const fileName = generateModelFilename('azure-diagram', 'png');
        link.download = fileName;
        link.href = url;
        link.click();
        URL.revokeObjectURL(url);
        recordExport('png', fileName);
        trackExport('png', nodes.filter(n => n.type === 'azureNode').length);
      } catch (err) {
        console.error('Error exporting diagram:', err);
        alert(t("Failed to export diagram. Please try again."));
      }
    }, 800);
  }, [reactFlowInstance, recordExport, nodes, t]);

  const exportAsSvg = useCallback(async () => {
    if (!reactFlowWrapper.current || !reactFlowInstance) {
      return;
    }

    // Fit all nodes into view with no animation for immediate rendering
    reactFlowInstance.fitView({ padding: 0.2, duration: 0 });

    // Wait for fitView to settle, then capture
    setTimeout(async () => {
      try {
        // captureDiagramAsSvg serialises the DOM natively — SVG edge paths
        // (curves, dashes, orthogonal bends) are preserved as vector data.
        const svgText = await captureDiagramAsSvg(reactFlowWrapper.current as HTMLElement, {
          backgroundColor: '#f8fafc',
          excludePanels: true,
        });

        const blob = new Blob([svgText], { type: 'image/svg+xml;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        const fileName = generateModelFilename('azure-diagram', 'svg');
        link.download = fileName;
        link.click();
        URL.revokeObjectURL(url);
        recordExport('svg', fileName);
        trackExport('svg', nodes.filter(n => n.type === 'azureNode').length);
      } catch (err) {
        console.error('Error exporting SVG:', err);
        alert(t("Failed to export SVG. Please try again."));
      }
    }, 800);
  }, [reactFlowInstance, recordExport, nodes, t]);

  // Export the workflow narrative (title, prompt, services, step-by-step flow,
  // connections, optional validation/cost) as a Markdown document.
  const exportWorkflowMarkdown = useCallback(() => {
    if (nodes.filter(n => n.type === 'azureNode').length === 0) {
      alert(t("Add or generate an architecture first, then export its workflow narrative."));
      return;
    }
    try {
      const pricingBreakdown = calculateCostBreakdown(nodes, undefined, pricingMode);
      const md = buildWorkflowMarkdown({
        title: titleBlockData,
        prompt: architecturePrompt,
        model: generatedWithModel,
        nodes,
        edges,
        workflow,
        validationScore: validationResult ? validationResult.overallScore : null,
        totalMonthlyCost: pricingBreakdown.totalMonthlyCost,
        pricingMode,
        region: pricingBreakdown.region,
      });
      const blob = new Blob([md], { type: 'text/markdown;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      const fileName = generateModelFilename('azure-diagram-workflow', 'md');
      link.download = fileName;
      link.click();
      URL.revokeObjectURL(url);
      recordExport('workflow-md', fileName);
      trackExport('workflow-md', nodes.filter(n => n.type === 'azureNode').length);
    } catch (err) {
      console.error('Error exporting workflow markdown:', err);
      alert(t("Failed to export workflow narrative. Please try again."));
    }
  }, [nodes, edges, workflow, titleBlockData, architecturePrompt, generatedWithModel, validationResult, pricingMode, recordExport, t]);

  // Export as an Animated SVG: same vector capture as exportAsSvg, but with
  // flowing data-flow circles injected onto each edge. Pure client-side — the
  // motion is carried by the SVG (open in a browser to view). For README/Teams
  // surfaces that strip SVG animation, export a GIF/WebP instead.
  const exportAsAnimatedSvg = useCallback(async () => {
    if (!reactFlowWrapper.current || !reactFlowInstance) {
      return;
    }

    reactFlowInstance.fitView({ padding: 0.2, duration: 0 });

    setTimeout(async () => {
      try {
        const svgText = await captureDiagramAsSvg(reactFlowWrapper.current as HTMLElement, {
          backgroundColor: '#f8fafc',
          excludePanels: true,
        });
        const animatedSvg = animateEdgeFlow(svgText);

        const blob = new Blob([animatedSvg], { type: 'image/svg+xml;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        const fileName = generateModelFilename('azure-diagram-animated', 'svg');
        link.download = fileName;
        link.click();
        URL.revokeObjectURL(url);
        recordExport('animated-svg', fileName);
        trackExport('animated-svg', nodes.filter(n => n.type === 'azureNode').length);
      } catch (err) {
        console.error('Error exporting animated SVG:', err);
        alert(t("Failed to export animated SVG. Please try again."));
      }
    }, 800);
  }, [reactFlowInstance, recordExport, nodes, t]);

  // Export a SEQUENCED "workflow animation" SVG: plays the diagram's workflow
  // steps chronologically (one edge flows at a time) with a caption per step and
  // pulsing highlights on the involved nodes. Client-side; the motion is carried
  // by the SVG (open in a browser). Requires a workflow on the diagram.
  const exportWorkflowAnimation = useCallback(async () => {
    if (!reactFlowWrapper.current || !reactFlowInstance) {
      return;
    }
    if (!workflow || workflow.length === 0) {
      alert(t("This diagram has no workflow steps to animate. Generate a diagram with a workflow first."));
      return;
    }

    reactFlowInstance.fitView({ padding: 0.2, duration: 0 });

    setTimeout(async () => {
      try {
        const svgText = await captureDiagramAsSvg(reactFlowWrapper.current as HTMLElement, {
          backgroundColor: '#f8fafc',
          excludePanels: true,
        });
        const sequenced = sequenceWorkflowSvg(svgText, { nodes, edges, workflow, stepDurSec: 3 });

        const blob = new Blob([sequenced], { type: 'image/svg+xml;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        const fileName = generateModelFilename('azure-diagram-workflow', 'svg');
        link.download = fileName;
        link.click();
        URL.revokeObjectURL(url);
        recordExport('workflow-animation', fileName);
        trackExport('workflow-animation', nodes.filter(n => n.type === 'azureNode').length);
      } catch (err) {
        console.error('Error exporting workflow animation:', err);
        alert(t("Failed to export workflow animation. Please try again."));
      }
    }, 800);
  }, [reactFlowInstance, recordExport, nodes, edges, workflow, t]);

  const exportAsDrawio = useCallback(async () => {
    try {
      const diagramName = titleBlockData.architectureName || 'Azure Architecture';
      const fileName = await exportAndDownloadDrawio(nodes, edges, diagramName);
      recordExport('drawio', fileName);
      trackExport('drawio', nodes.filter(n => n.type === 'azureNode').length);
    } catch (err) {
      console.error('Error exporting Draw.io:', err);
      alert(t("Failed to export Draw.io file. Please try again."));
    }
  }, [nodes, edges, titleBlockData.architectureName, recordExport, t]);

  const exportAsVsdx = useCallback(async () => {
    if (nodes.filter(n => n.type === 'azureNode').length === 0) {
      alert(t("Add or generate an architecture first, then export to Visio."));
      return;
    }
    try {
      const diagramName = titleBlockData.architectureName || 'Azure Architecture';
      const blob = await buildVsdxBlob(nodes, edges, diagramName);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      const fileName = generateModelFilename('azure-diagram', 'vsdx');
      link.download = fileName;
      link.click();
      URL.revokeObjectURL(url);
      recordExport('vsdx', fileName);
      trackExport('vsdx', nodes.filter(n => n.type === 'azureNode').length);
    } catch (err) {
      console.error('Error exporting Visio VSDX:', err);
      alert(t("Failed to export Visio file. Please try again."));
    }
  }, [nodes, edges, titleBlockData.architectureName, recordExport, t]);

  const exportAsHtml = useCallback(() => {
    try {
      const diagramName = titleBlockData.architectureName || 'Azure Architecture';
      exportDiagramAsHtml(nodes, edges, diagramName);
      const fileName = `${diagramName.replace(/[^a-zA-Z0-9-_ ]/g, '').replace(/\s+/g, '-').toLowerCase()}.html`;
      recordExport('html', fileName);
      trackExport('html', nodes.filter(n => n.type === 'azureNode').length);
    } catch (err) {
      console.error('Error exporting HTML diagram:', err);
      alert(t("Failed to export HTML diagram. Please try again."));
    }
  }, [nodes, edges, titleBlockData.architectureName, recordExport, t]);

  const exportAsPptx = useCallback(async () => {
    if (!reactFlowWrapper.current || !reactFlowInstance) return;

    reactFlowInstance.fitView({ padding: 0.2, duration: 0 });

    setTimeout(async () => {
      try {
        const imageDataUrl = await captureDiagramAsPng(reactFlowWrapper.current as HTMLElement, {
          backgroundColor: isDarkMode ? '#1e293b' : '#f8fafc',
          excludePanels: true,
        });

        const fileName = await exportDiagramAsPptx(imageDataUrl, {
          diagramName: titleBlockData.architectureName || 'Azure Architecture',
          author: titleBlockData.author || 'Azure Architect',
          date: titleBlockData.date || new Date().toLocaleDateString(),
          isDarkMode,
        });

        recordExport('pptx', fileName);
        trackExport('pptx', nodes.filter(n => n.type === 'azureNode').length);
      } catch (err) {
        console.error('Error exporting PPTX:', err);
        alert(t("Failed to export PowerPoint slide. Please try again."));
      }
    }, 800);
  }, [reactFlowInstance, recordExport, nodes, isDarkMode, titleBlockData, t]);

  const exportCustomerDeck = useCallback(async () => {
    if (!reactFlowWrapper.current || !reactFlowInstance) return;

    const azureNodes = nodes.filter(n => n.type === 'azureNode');
    if (azureNodes.length === 0) {
      alert(t('Add or generate an architecture first, then export a customer deck.'));
      return;
    }

    reactFlowInstance.fitView({ padding: 0.2, duration: 0 });

    setTimeout(async () => {
      try {
        const imageDataUrl = await captureDiagramAsPng(reactFlowWrapper.current as HTMLElement, {
          backgroundColor: isDarkMode ? '#1e293b' : '#f8fafc',
          excludePanels: true,
        });

        // Service inventory from the diagram nodes. Group membership is via
        // React Flow's parent link (parentNode/parentId) → the group node's
        // label; category is derived from the icon path (/Icons/<category>/…).
        const groupLabelById = new Map<string, string>();
        nodes.filter(n => n.type === 'groupNode').forEach(g => {
          const label = (g.data?.label as string) || '';
          if (label) groupLabelById.set(g.id, label);
        });
        const categoryFromIcon = (iconPath?: string): string | undefined => {
          const m = iconPath?.match(/\/Icons\/([^/]+)\//i);
          if (!m) return undefined;
          return m[1].replace(/[-_]+/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
        };
        const services: DeckService[] = azureNodes.map(n => {
          const parentId = (n as Node & { parentNode?: string; parentId?: string }).parentNode
            ?? (n as Node & { parentNode?: string; parentId?: string }).parentId;
          return {
            name: (n.data?.label as string) || 'Unnamed service',
            category: categoryFromIcon(n.data?.iconPath as string),
            group: (parentId ? groupLabelById.get(parentId) : undefined) || undefined,
          };
        });

        // Optional WAF review
        const validation = validationResult ? {
          overallScore: validationResult.overallScore,
          overallLabel: bandLabel(validationResult.overallScore),
          summary: validationResult.summary,
          pillars: validationResult.pillars.map(p => ({ pillar: p.pillar, score: p.score, maturity: bandLabel(p.score) })),
          findings: (validationResult.quickWins.length > 0
            ? validationResult.quickWins
            : validationResult.pillars.flatMap(p => p.findings)
          )
            .slice()
            .sort((a, b) => {
              const rank = { critical: 0, high: 1, medium: 2, low: 3 } as Record<string, number>;
              return (rank[a.severity] ?? 9) - (rank[b.severity] ?? 9);
            })
            .slice(0, 6)
            .map(f => ({ severity: f.severity, category: f.category, issue: f.issue, recommendation: f.recommendation })),
          modelUsed: validationResult.modelUsed,
        } : null;

        // Optional cost estimate — enriched to mirror the Cost Intelligence
        // report (annual projection, fixed vs usage split, top drivers %, and
        // an optional multi-region comparison).
        const breakdown = calculateCostBreakdown(nodes, undefined, pricingMode);
        const hasCost = breakdown.byService.length > 0;

        // Multi-region comparison (best-effort; reprices over local pricing data)
        let regions: Array<{ name: string; flag?: string; monthly: number; annual: number; isCurrent: boolean; isCheapest: boolean }> | undefined;
        let unavailableRegions: string[] | undefined;
        if (hasCost) {
          const comparison = await calculateRegionalCostComparison(nodes, pricingMode);
          const comparisonComplete = comparison.failures.length === 0;
          unavailableRegions = comparison.failures.map(
            failure => `${failure.info.displayName}: ${failure.reason}`,
          );
          if (comparison.results.length > 1) {
            const min = comparison.results[0].total;
            const activeRegion = AVAILABLE_REGIONS.some(r => r.id === breakdown.region)
              ? breakdown.region
              : undefined;
            regions = comparison.results.map(r => ({
              name: r.info.displayName,
              flag: r.info.flag,
              monthly: r.total,
              annual: r.annual,
              isCurrent: activeRegion !== undefined && r.info.id === activeRegion,
              isCheapest: comparisonComplete && r.total === min,
            }));
          }
        }

        const azureServiceNodes = nodes.filter(n => n.type === 'azureNode');
        const fixedCost = hasCost
          ? breakdown.byService
              .filter(svc => { const node = azureServiceNodes.find(n => n.id === svc.nodeId); return !(node?.data?.pricing as any)?.isUsageBased; })
              .reduce((sum, svc) => sum + svc.cost, 0)
          : 0;
        const cost = hasCost ? {
          totalMonthly: breakdown.totalMonthlyCost,
          annual: breakdown.totalMonthlyCost * 12,
          currency: breakdown.currency || 'USD',
          term: breakdown.pricingTerm,
          region: breakdown.region,
          pricesAsOf: breakdown.pricesAsOf,
          fixedCost,
          usageCost: breakdown.totalMonthlyCost - fixedCost,
          byCategory: breakdown.byCategory
            .slice()
            .sort((a, b) => b.cost - a.cost)
            .map(c => ({ category: c.category, cost: c.cost, percentage: c.percentage })),
          topServices: breakdown.byService
            .slice()
            .sort((a, b) => b.cost - a.cost)
            .slice(0, 10)
            .map(s => ({
              serviceName: s.serviceName,
              cost: s.cost,
              tier: s.tier,
              percentage: breakdown.totalMonthlyCost > 0 ? (s.cost / breakdown.totalMonthlyCost) * 100 : 0,
            })),
          regions,
          regionComparisonIncomplete: Boolean(unavailableRegions?.length),
          unavailableRegions,
        } : null;

        const fileName = await exportArchitectureDeck(imageDataUrl, {
          diagramName: titleBlockData.architectureName || 'Azure Architecture',
          author: titleBlockData.author || 'Azure Architect',
          date: titleBlockData.date || new Date().toLocaleDateString(),
          isDarkMode,
          prompt: (originalPrompt || architecturePrompt) || undefined,
          model: generatedWithModel?.name,
          services,
          validation,
          cost,
        });

        recordExport('pptx', fileName);
        trackExport('pptx-deck', azureNodes.length);
      } catch (err) {
        console.error('Error exporting customer deck:', err);
        alert(t('Failed to export the customer deck. Please try again.'));
      }
    }, 800);
  }, [reactFlowInstance, recordExport, nodes, isDarkMode, titleBlockData, validationResult, pricingMode, architecturePrompt, originalPrompt, generatedWithModel, t]);

  // ── az prototype export removed (feature unused) ───────────────────────

  const saveDiagram = useCallback(() => {
    const flow = reactFlowInstance?.toObject();
    const diagramData = {
      ...flow,
      metadata: {
        ...titleBlockData,
        savedAt: new Date().toISOString(),
      },
      workflow: workflow.length > 0 ? workflow : undefined,
      pricingScenarios,
      architecturePrompt: architecturePrompt || undefined,
      originalPrompt: originalPrompt || architecturePrompt || undefined,
      iacBaseline,
    };
    const dataStr = JSON.stringify(diagramData, null, 2);
    const dataUri = 'data:application/json;charset=utf-8,' + encodeURIComponent(dataStr);
    
    const link = document.createElement('a');
    link.setAttribute('href', dataUri);
    const fileName = generateModelFilename('azure-diagram', 'json');
    link.setAttribute('download', fileName);
    link.click();
    recordExport('json', fileName);
    trackExport('json', nodes.filter(n => n.type === 'azureNode').length);
  }, [reactFlowInstance, recordExport, titleBlockData, workflow, pricingScenarios, architecturePrompt, originalPrompt, nodes, iacBaseline]);

  const exportCostBreakdown = useCallback(() => {
    // Calculate the cost breakdown
    const breakdown = calculateCostBreakdown(nodes, undefined, pricingMode);
    
    // Check if there's any cost data
    if (breakdown.byService.length === 0) {
      alert(t("No costing information available. Please ensure your diagram contains Azure services with pricing data."));
      return;
    }

    // Export as CSV
    const csvData = exportCostBreakdownCSV(breakdown, nodes);
    const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const baseName = generateModelFilename('azure-cost-breakdown', 'csv');
    // Insert region before the extension for cost exports
    const fileName = baseName.replace(
      '.csv',
      `-${toFileNameSegment(breakdown.region)}.csv`,
    );
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url);
    recordExport('costs', fileName);
    trackExport('csv', nodes.filter(n => n.type === 'azureNode').length);
  }, [nodes, recordExport, pricingMode, t]);

  const exportCostBreakdownZip = useCallback(async () => {
    const breakdown = calculateCostBreakdown(nodes, undefined, pricingMode);
    if (breakdown.byService.length === 0) {
      alert(t("No costing information available. Please ensure your diagram contains Azure services with pricing data."));
      return;
    }

    // ── Multi-region comparison ──────────────────────────────────────────────
    const comparison = await calculateRegionalCostComparison(nodes, pricingMode);
    const regionResults = comparison.results;
    const regionFailures = comparison.failures;
    const regionalComparisonComplete = regionFailures.length === 0;
    const lowestComparable = regionResults[0];
    const highestComparable = regionResults[regionResults.length - 1];
    const cheapest = regionalComparisonComplete ? lowestComparable : undefined;
    const mostExpensive = regionalComparisonComplete ? highestComparable : undefined;

    // Build intelligent analysis text
    const reportRegion = breakdown.region;
    const regionInfo = AVAILABLE_REGIONS.find(r => r.id === reportRegion);
    const currentRegionId = regionInfo?.id;
    const reportRegionLabel = regionInfo
      ? `${regionInfo.displayName} (${regionInfo.id})`
      : reportRegion;
    const annual = breakdown.totalMonthlyCost * 12;
    const sortedServices = [...breakdown.byService].sort((a, b) => b.cost - a.cost);
    const topDrivers = sortedServices.slice(0, 5);
    const topService = sortedServices[0];
    const topServicePct = breakdown.totalMonthlyCost > 0 ? (topService.cost / breakdown.totalMonthlyCost) * 100 : 0;
    const azureServiceNodes = nodes.filter(n => n.type === 'azureNode');
    const usageBasedCount = breakdown.byService.filter(svc => {
      const node = azureServiceNodes.find(n => n.id === svc.nodeId);
      return (node?.data?.pricing as any)?.isUsageBased;
    }).length;
    const fixedCost = breakdown.byService
      .filter(svc => { const node = azureServiceNodes.find(n => n.id === svc.nodeId); return !(node?.data?.pricing as any)?.isUsageBased; })
      .reduce((sum, svc) => sum + svc.cost, 0);
    const usageCost = breakdown.totalMonthlyCost - fixedCost;

    // ── Build intelligent analysis as Markdown ───────────────────────────────
    const mdBar = (pct: number) => '█'.repeat(Math.max(0, Math.round(pct / 5))).padEnd(20, '░');
    const fixedPct = breakdown.totalMonthlyCost > 0 ? ((fixedCost / breakdown.totalMonthlyCost) * 100).toFixed(1) : '0.0';
    const usagePct = breakdown.totalMonthlyCost > 0 ? ((usageCost / breakdown.totalMonthlyCost) * 100).toFixed(1) : '0.0';

    const analysisLines: string[] = [
      '# Azure Architecture — Cost Intelligence Report',
      '',
      `_Generated ${new Date().toLocaleString()} · Region: \`${reportRegionLabel}\` · ${azureServiceNodes.length} service(s) on diagram_`,
      '',
    ];

    // TL;DR callout
    analysisLines.push('> **TL;DR**');
    analysisLines.push(`> Estimated **$${breakdown.totalMonthlyCost.toFixed(2)}/mo** (**$${annual.toFixed(2)}/yr**).`);
    if (cheapest) {
      analysisLines.push(`> Cheapest region: **${cheapest.info.flag} ${cheapest.info.displayName}** at $${cheapest.total.toFixed(2)}/mo.`);
    } else if (regionFailures.length > 0) {
      analysisLines.push(`> Regional comparison is partial: ${regionResults.length} of ${AVAILABLE_REGIONS.length} regions are comparable. No cheapest-region recommendation is made.`);
    }
    {
      const currentResult = currentRegionId
        ? regionResults.find(r => r.info.id === currentRegionId)
        : undefined;
      if (currentResult && cheapest && currentResult.info.id !== cheapest.info.id) {
        const savingsMonthly = currentResult.total - cheapest.total;
        analysisLines.push(`> Potential saving by switching region: ~$${savingsMonthly.toFixed(2)}/mo ($${(savingsMonthly * 12).toFixed(2)}/yr).`);
      } else if (currentResult && cheapest && currentResult.info.id === cheapest.info.id) {
        analysisLines.push('> Already on the cheapest available region. ✅');
      }
    }
    analysisLines.push('');

    // Cost summary
    analysisLines.push('## Cost summary');
    analysisLines.push('');
    analysisLines.push('| Metric | Value |');
    analysisLines.push('| --- | ---: |');
    analysisLines.push(`| Monthly estimate | **$${breakdown.totalMonthlyCost.toFixed(2)}** |`);
    analysisLines.push(`| Annual projection | $${annual.toFixed(2)} |`);
    analysisLines.push(`| Fixed costs | $${fixedCost.toFixed(2)}/mo (${fixedPct}%) |`);
    analysisLines.push(`| Usage-based costs | $${usageCost.toFixed(2)}/mo (${usagePct}%) — actual may vary |`);
    analysisLines.push('');

    // Top cost drivers
    analysisLines.push('## Top cost drivers');
    analysisLines.push('');
    analysisLines.push('| # | Service | Monthly cost | Share | |');
    analysisLines.push('| ---: | --- | ---: | ---: | --- |');
    topDrivers.forEach((svc, i) => {
      const pct = breakdown.totalMonthlyCost > 0 ? (svc.cost / breakdown.totalMonthlyCost) * 100 : 0;
      analysisLines.push(`| ${i + 1} | ${svc.serviceName.replace(/\|/g, '\\|')} | $${svc.cost.toFixed(2)} | ${pct.toFixed(1)}% | \`${mdBar(pct)}\` |`);
    });
    analysisLines.push('');
    // Pie chart of the top cost drivers (per-service): top 8 services + an
    // aggregated "Other services" slice gives a meaningful distribution every
    // time, unlike the category breakdown which often collapses to "Other".
    {
      const nonZero = sortedServices.filter(s => s.cost > 0);
      const pieRows: { label: string; value: number }[] = nonZero
        .slice(0, 8)
        .map(s => ({ label: s.serviceName, value: s.cost }));
      const rest = nonZero.slice(8).reduce((sum, s) => sum + s.cost, 0);
      if (rest > 0) pieRows.push({ label: 'Other services', value: rest });
      if (pieRows.length > 0) {
        analysisLines.push('```mermaid');
        analysisLines.push('pie showData title Top cost drivers (monthly USD)');
        pieRows.forEach(r => {
          const label = r.label.replace(/"/g, "'");
          analysisLines.push(`    "${label}" : ${r.value.toFixed(2)}`);
        });
        analysisLines.push('```');
        analysisLines.push('');
      }
    }

    // Flags & recommendations
    analysisLines.push('## Flags & recommendations');
    analysisLines.push('');
    let hasFlags = false;
    if (topServicePct > 50) {
      analysisLines.push(`- ⚠️ **Cost concentration:** "${topService.serviceName}" is ${topServicePct.toFixed(0)}% of total. Consider reviewing tier/quantity or splitting the workload.`);
      hasFlags = true;
    }
    if (usageBasedCount > 0) {
      analysisLines.push(`- ℹ️ **${usageBasedCount} usage-based service(s)** detected (e.g. Functions, OpenAI). Actual monthly spend may differ significantly from estimates.`);
      hasFlags = true;
    }
    if (annual > 100000) {
      analysisLines.push('- 💡 **Annual spend >$100k** — consider Azure Reserved Instances / Savings Plans (typically 30–40% savings on compute with a 1- or 3-year commitment).');
      hasFlags = true;
    } else if (annual > 12000) {
      analysisLines.push('- 💡 Azure Reserved Instances may offer savings on compute-heavy services.');
      hasFlags = true;
    }
    if (!hasFlags) {
      analysisLines.push('- ✅ No cost concentration or optimization flags raised for this architecture.');
    }
    analysisLines.push('');

    // Multi-region comparison
    analysisLines.push('## Multi-region cost comparison');
    analysisLines.push('');
    if (regionResults.length === 0) {
      analysisLines.push('_No region could preserve every selected SKU, so a like-for-like comparison is unavailable._');
      if (regionFailures.length > 0) {
        analysisLines.push('');
        analysisLines.push('**Unavailable regions**');
        regionFailures.forEach(failure => {
          analysisLines.push(`- ${failure.info.flag} ${failure.info.displayName} (\`${failure.info.id}\`): ${failure.reason}`);
        });
      }
      analysisLines.push('');
    } else {
      analysisLines.push(`| Rank | Region | Monthly | Annual | ${regionalComparisonComplete ? 'vs Cheapest' : 'vs Lowest shown'} | ${currentRegionId ? 'vs Current' : 'vs Diagram estimate'} |`);
      analysisLines.push('| ---: | --- | ---: | ---: | ---: | ---: |');
      const currentTotal = breakdown.totalMonthlyCost;
      regionResults.forEach((r, idx) => {
        const isCurrent = currentRegionId !== undefined && r.info.id === currentRegionId;
        const isLowestShown = idx === 0;
        const vsLowest = isLowestShown
          ? 'baseline'
          : lowestComparable && lowestComparable.total > 0
            ? `+${(((r.total - lowestComparable.total) / lowestComparable.total) * 100).toFixed(1)}%`
            : '—';
        const vsCurrent = isCurrent
          ? 'current'
          : r.total < currentTotal
            ? `−${(((currentTotal - r.total) / currentTotal) * 100).toFixed(1)}% 💰`
            : `+${(((r.total - currentTotal) / currentTotal) * 100).toFixed(1)}%`;
        const marker = regionalComparisonComplete && isLowestShown ? ' ★' : isCurrent ? ' ◀' : '';
        analysisLines.push(`| ${idx + 1}${marker} | ${r.info.flag} ${r.info.displayName} (\`${r.info.id}\`) | $${r.total.toFixed(2)} | $${r.annual.toFixed(2)} | ${vsLowest} | ${vsCurrent} |`);
      });
      analysisLines.push('');
      if (!regionalComparisonComplete) {
        analysisLines.push(`- ⚠️ **Partial comparison:** ${regionResults.length} of ${AVAILABLE_REGIONS.length} regions preserve every selected SKU. No global cheapest or savings recommendation is shown.`);
        if (lowestComparable) {
          analysisLines.push(`- **Lowest shown:** ${lowestComparable.info.flag} ${lowestComparable.info.displayName} — $${lowestComparable.total.toFixed(2)}/mo. This is not a global cheapest-region claim.`);
        }
        analysisLines.push('- **Unavailable regions:**');
        regionFailures.forEach(failure => {
          analysisLines.push(`  - ${failure.info.flag} ${failure.info.displayName} (\`${failure.info.id}\`): ${failure.reason}`);
        });
      } else if (cheapest) {
        analysisLines.push(`- ★ **Cheapest:** ${cheapest.info.flag} ${cheapest.info.displayName} — $${cheapest.total.toFixed(2)}/mo ($${cheapest.annual.toFixed(2)}/yr)`);
        if (mostExpensive) {
          const premiumPct = cheapest.total > 0 ? (((mostExpensive.total - cheapest.total) / cheapest.total) * 100).toFixed(1) : '0.0';
          analysisLines.push(`- 🔥 **Priciest:** ${mostExpensive.info.flag} ${mostExpensive.info.displayName} — $${mostExpensive.total.toFixed(2)}/mo (+${premiumPct}% above cheapest)`);
        }
        const currentResult = currentRegionId
          ? regionResults.find(r => r.info.id === currentRegionId)
          : undefined;
        if (currentResult && currentResult.info.id !== cheapest.info.id) {
          const savingsMonthly = currentResult.total - cheapest.total;
          analysisLines.push(`- 💡 **Potential savings:** switching ${currentResult.info.displayName} → ${cheapest.info.displayName} saves ~$${savingsMonthly.toFixed(2)}/mo ($${(savingsMonthly * 12).toFixed(2)}/yr) — verify service availability before migrating.`);
        } else if (currentResult && currentResult.info.id === cheapest.info.id) {
          analysisLines.push('- ✅ You are already on the cheapest available region for this architecture.');
        }
      }
      analysisLines.push('');

      // Per-service regional variance — top 3 services with biggest price spread
      if (regionResults.length >= 2) {
        const serviceVariance: { name: string; min: number; max: number; spread: number }[] = [];
        breakdown.byService.forEach(svc => {
          const prices = regionResults
            .map(r => r.breakdown.byService.find(s => s.nodeId === svc.nodeId)?.cost ?? 0)
            .filter(p => p > 0);
          if (prices.length > 1) {
            const minP = Math.min(...prices);
            const maxP = Math.max(...prices);
            serviceVariance.push({ name: svc.serviceName, min: minP, max: maxP, spread: maxP - minP });
          }
        });
        serviceVariance.sort((a, b) => b.spread - a.spread);
        const top3 = serviceVariance.slice(0, 3);
        if (top3.length > 0) {
          analysisLines.push('### Top services by regional price variance');
          analysisLines.push('');
          analysisLines.push('| Service | Min | Max | Spread |');
          analysisLines.push('| --- | ---: | ---: | ---: |');
          top3.forEach(sv => {
            const spreadPct = sv.min > 0 ? (((sv.max - sv.min) / sv.min) * 100).toFixed(1) : '0.0';
            analysisLines.push(`| ${sv.name.replace(/\|/g, '\\|')} | $${sv.min.toFixed(2)} | $${sv.max.toFixed(2)} | ${spreadPct}% |`);
          });
          analysisLines.push('');
        }
      }
    }

    // Cost by group
    analysisLines.push('## Cost by group');
    analysisLines.push('');
    if (breakdown.byGroup.length === 0) {
      analysisLines.push('_No groups defined in this diagram._');
    } else {
      analysisLines.push('| Group | Monthly cost | Services | Share |');
      analysisLines.push('| --- | ---: | ---: | ---: |');
      breakdown.byGroup.forEach(grp => {
        const pct = breakdown.totalMonthlyCost > 0 ? ((grp.cost / breakdown.totalMonthlyCost) * 100).toFixed(1) : '0.0';
        analysisLines.push(`| ${grp.groupLabel.replace(/\|/g, '\\|')} | $${grp.cost.toFixed(2)} | ${grp.serviceCount} | ${pct}% |`);
      });
    }
    analysisLines.push('');
    analysisLines.push('---');
    analysisLines.push('');
    analysisLines.push('_Generated by Azure Architecture Diagram Builder. Estimates are indicative and exclude taxes, bandwidth egress, and support plans unless modeled explicitly._');


    // Build ZIP
    const zip = new JSZip();
    const baseName = generateModelFilename('azure-cost', 'zip').replace('.zip', '');
    const fileBase = `${baseName}-${toFileNameSegment(reportRegion)}`;

    const summaryMd = getCostSummaryMarkdown(breakdown);
    const analysisMd = analysisLines.join('\n');
    // Combined "one file to read": summary first, then the full analysis.
    const combinedMd = [
      summaryMd,
      '',
      '<br />',
      '',
      analysisMd,
    ].join('\n');

    zip.file(`${fileBase}.csv`, exportCostBreakdownCSV(breakdown, nodes));
    zip.file(`${fileBase}.json`, exportCostBreakdownJSON(breakdown));
    zip.file(`${fileBase}-summary.md`, summaryMd);
    zip.file(`${fileBase}-analysis.md`, analysisMd);
    zip.file(`${fileBase}-report.md`, combinedMd);
    // HTML render of the combined report for non-Markdown viewers.
    zip.file(`${fileBase}-report.html`, costReportToHtml('Azure Architecture Cost Report', combinedMd));
    // Tiny manifest explaining each file in the bundle.
    zip.file('README.md', [
      '# Azure Architecture Cost Export',
      '',
      `Generated ${new Date().toLocaleString()} for region \`${reportRegionLabel}\` by Azure Architecture Diagram Builder.`,
      '',
      ...(regionFailures.length > 0
        ? [`> Regional comparison is partial: ${regionResults.length} of ${AVAILABLE_REGIONS.length} regions preserve every selected SKU. Unavailable regions are listed in the analysis and comparison CSV.`, '']
        : []),
      'This bundle contains the same cost estimate in several formats — open whichever suits your tooling:',
      '',
      '| File | Format | Best for |',
      '| --- | --- | --- |',
      `| \`${fileBase}-report.md\` | Markdown | **Start here** — combined summary + full analysis in one file |`,
      `| \`${fileBase}-report.html\` | HTML | Same combined report, viewable in any browser (Mermaid pie chart included) |`,
      `| \`${fileBase}-summary.md\` | Markdown | Quick cost summary tables (by service, group, category) |`,
      `| \`${fileBase}-analysis.md\` | Markdown | Detailed intelligence report (drivers, flags, multi-region comparison) |`,
      `| \`${fileBase}.csv\` | CSV | Per-service breakdown for Excel / spreadsheets |`,
      `| \`${fileBase}.json\` | JSON | Structured data for programmatic use / automation |`,
      `| \`${fileBase}-multiregion-comparison.csv\` | CSV | Per-service pricing and SKU availability across all regions |`,
      '',
      '> Estimates are indicative and exclude taxes, bandwidth egress, and support plans unless modeled explicitly.',
      '> Usage-based services (e.g. Functions, OpenAI) may vary with actual consumption.',
    ].join('\n'));

    // Multi-region comparison CSV
    if (regionResults.length > 0 || regionFailures.length > 0) {
      const mrLines: string[] = [
        `Region,Region ID,Geography,Flag,Type,Monthly Cost (USD),Annual Cost (USD),${regionalComparisonComplete ? 'vs Cheapest (%)' : 'vs Lowest Shown (%)'},${currentRegionId ? 'vs Current Region (%)' : 'vs Diagram Estimate (%)'},Status,Error`,
      ];
      const currentTotal = breakdown.totalMonthlyCost;
      regionResults.forEach(r => {
        const vsLowest = lowestComparable && lowestComparable.total > 0
          ? r.info.id === lowestComparable.info.id ? '0.00' : (((r.total - lowestComparable.total) / lowestComparable.total) * 100).toFixed(2)
          : '';
        const vsCurrent = currentTotal > 0
          ? currentRegionId && r.info.id === currentRegionId
            ? '0.00'
            : (((r.total - currentTotal) / currentTotal) * 100).toFixed(2)
          : '';
        mrLines.push(`${csvTextCell(r.info.displayName, true)},${csvTextCell(r.info.id)},${csvTextCell(r.info.geography, true)},${csvTextCell(r.info.flag)},${csvTextCell(r.info.regionType)},${r.total.toFixed(2)},${r.annual.toFixed(2)},${vsLowest},${vsCurrent},Comparable,`);
      });
      regionFailures.forEach(failure => {
        mrLines.push(`${csvTextCell(failure.info.displayName, true)},${csvTextCell(failure.info.id)},${csvTextCell(failure.info.geography, true)},${csvTextCell(failure.info.flag)},${csvTextCell(failure.info.regionType)},,,,,Unavailable,${csvTextCell(failure.reason, true)}`);
      });
      // Per-service per-region detail sheet
      mrLines.push('');
      mrLines.push('Service,Node ID,' + AVAILABLE_REGIONS.map(r => r.displayName).join(','));
      const resultsByRegionId = new Map(regionResults.map(result => [result.info.id, result]));
      breakdown.byService.forEach(svc => {
        const prices = AVAILABLE_REGIONS.map(region => {
          const result = resultsByRegionId.get(region.id);
          const match = result?.breakdown.byService.find(s => s.nodeId === svc.nodeId);
          return match ? match.cost.toFixed(2) : '';
        });
        mrLines.push(`${csvTextCell(svc.serviceName, true)},${csvTextCell(svc.nodeId)},${prices.join(',')}`);
      });
      zip.file(`${fileBase}-multiregion-comparison.csv`, mrLines.join('\n'));
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(zipBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${fileBase}-all-formats.zip`;
    link.click();
    URL.revokeObjectURL(url);
    recordExport('costs', `${fileBase}-all-formats.zip`);
    trackExport('csv', azureServiceNodes.length);
  }, [nodes, recordExport, pricingMode, t]);

  const applyFlowObject = useCallback(
    (flow: unknown) => {
      if (!isRecord(flow)) {
        throw new Error('Invalid diagram payload');
      }
      if (!Array.isArray(flow.nodes) || !Array.isArray(flow.edges)) {
        throw new Error('Diagram payload must contain node and edge arrays');
      }
      if (flow.nodes.length > 2000 || flow.edges.length > 5000) {
        throw new Error('Diagram payload is too large');
      }
      const restoredNodes = validateRestoredNodes(flow.nodes);
      const nodeIds = new Set(restoredNodes.map((node) => node.id));
      const restoredEdges = validateRestoredEdges(flow.edges, nodeIds);
      const restoredWorkflow = validateRestoredWorkflow(flow.workflow);
      let restoredViewport: { x: number; y: number; zoom: number } | undefined;
      if (flow.viewport !== undefined && flow.viewport !== null) {
        const viewport = flow.viewport;
        if (
          !isRecord(viewport)
          || typeof viewport.x !== 'number'
          || !Number.isFinite(viewport.x)
          || typeof viewport.y !== 'number'
          || !Number.isFinite(viewport.y)
          || typeof viewport.zoom !== 'number'
          || !Number.isFinite(viewport.zoom)
          || viewport.zoom <= 0
        ) {
          throw new Error('Invalid viewport in diagram payload');
        }
        restoredViewport = {
          x: viewport.x,
          y: viewport.y,
          zoom: viewport.zoom,
        };
      }

      // Normalize edge handle ids. Some scenes (e.g. exported by the MCP server
      // or hand-authored) use bare position names for handles, but AzureNode's
      // handles are asymmetric: valid sources are top-source/left-source/right/
      // bottom; valid targets are top/left/right-target/bottom-target. A bare
      // sourceHandle "top"/"left" or targetHandle "bottom"/"right" points at a
      // non-existent handle, so the edge silently fails to render. Remap the
      // invalid bare names to the correct handle id (valid ids pass through).
      const fixedEdges = normalizeRestoredEdges(restoredEdges);
      setNodes(restoredNodes);
      setEdges(fixedEdges);

      if (restoredViewport && reactFlowInstance?.setViewport) {
        reactFlowInstance.setViewport(restoredViewport);
      }

      // Restore metadata if present
      const restoredTitle = isRecord(flow.titleBlockData)
        ? flow.titleBlockData
        : flow.metadata;
      if (isRecord(restoredTitle)) {
        setTitleBlockData({
          architectureName: typeof restoredTitle.architectureName === 'string'
            ? restoredTitle.architectureName
            : 'Untitled Architecture',
          author: typeof restoredTitle.author === 'string' ? restoredTitle.author : 'Azure Architect',
          version: typeof restoredTitle.version === 'string' ? restoredTitle.version : '1.0',
          date: typeof restoredTitle.date === 'string'
            ? restoredTitle.date
            : new Date().toLocaleDateString(),
        });
      } else {
        setTitleBlockData({
          architectureName: 'Untitled Architecture',
          author: 'Azure Architect',
          version: '1.0',
          date: new Date().toLocaleDateString(),
        });
      }

      // Restore workflow if present
      setWorkflow(restoredWorkflow);
      if (Array.isArray(flow.pricingScenarios)) {
        setPricingScenarios(normalizePricingScenarios(flow.pricingScenarios));
      }
      setIaCBaseline(restoreIaCBaseline(flow.iacBaseline));
      setDriftPlanSummary(null);

      // Restore architecture prompt if present
      const restoredPrompt = typeof flow.architecturePrompt === 'string' ? flow.architecturePrompt : '';
      setArchitecturePrompt(restoredPrompt);
      setOriginalPrompt(typeof flow.originalPrompt === 'string' ? flow.originalPrompt : restoredPrompt);
      setReferenceImageUrl(null);
      setLastReferenceArchitecture(null);
      setLastBlueprintArchitecture(null);
    },
    [setNodes, setEdges, reactFlowInstance, normalizeRestoredEdges]
  );

  const cloudDiagramPayload = useMemo(() => ({
    nodes,
    edges,
    architecturePrompt,
    originalPrompt: originalPrompt || architecturePrompt || undefined,
    validationScore: validationResult?.overallScore,
    titleBlockData,
    workflow,
    pricingScenarios,
    iacBaseline,
  }), [
    nodes,
    edges,
    architecturePrompt,
    originalPrompt,
    validationResult?.overallScore,
    titleBlockData,
    workflow,
    pricingScenarios,
    iacBaseline,
  ]);

  const cloudSync = useCloudDiagramSync({
    diagramName: titleBlockData.architectureName,
    payload: cloudDiagramPayload,
    enabled: nodes.length > 0,
    onLoad: applyFlowObject,
  });

  const startFreshDiagram = useCallback(async (): Promise<boolean> => {
    const preserveAsCopy = cloudSync.context?.role === 'viewer';
    const confirmed = window.confirm(localize(language, {
      en: preserveAsCopy
        ? 'Start a new diagram? Your current view will first be saved as a personal copy, then the canvas will be cleared.'
        : 'Start a new diagram? The current cloud diagram will remain saved, and the canvas will be cleared.',
      ja: preserveAsCopy
        ? '新しい図面を作成しますか？現在の内容を先に個人用コピーとして保存してから、キャンバスをクリアします。'
        : '新しい図面を作成しますか？現在のクラウド図面は保存されたまま残り、キャンバスがクリアされます。',
    }));
    if (!confirmed) return false;

    try {
      if (preserveAsCopy) {
        await cloudSync.saveAsCopy();
      } else {
        await cloudSync.saveNow();
      }
    } catch {
      const discardUnsavedChanges = window.confirm(localize(language, {
        en: 'The current diagram could not be saved to the cloud. Start a new diagram anyway and discard these unsaved changes?',
        ja: '現在の図面をクラウドに保存できませんでした。未保存の変更を破棄して、新しい図面を開始しますか？',
      }));
      if (!discardUnsavedChanges) return false;
    }
    trackStartFresh();
    cloudSync.reset();
    setNodes([]);
    setEdges([]);
    setArchitecturePrompt('');
    setOriginalPrompt('');
    setWorkflow([]);
    setHighlightedServices([]);
    setNodeContextMenu(null);
    setEdgeContextMenu(null);
    setGeneratedWithModel(null);
    setValidationResult(null);
    setDeploymentGuide(null);
    setIaCBaseline(null);
    setDriftPlanSummary(null);
    setReferenceImageUrl(null);
    setLastReferenceArchitecture(null);
    setLastBlueprintArchitecture(null);
    setPricingScenarios([]);
    setTitleBlockData({
      architectureName: translate('Untitled Architecture'),
      author: translate('Azure Architect'),
      date: new Date().toISOString().split('T')[0],
      version: '1.0',
    });
    return true;
  }, [
    cloudSync,
    language,
    setEdges,
    setNodes,
    translate,
  ]);

  const iacComparison = useMemo(
    () => compareDiagramToBaseline(nodes, iacBaseline),
    [nodes, iacBaseline],
  );
  const bicepStarterTemplate = useMemo(
    () => buildStarterTemplate(nodes, 'bicep'),
    [nodes],
  );
  const terraformStarterTemplate = useMemo(
    () => buildStarterTemplate(nodes, 'terraform'),
    [nodes],
  );

  const downloadStarterTemplate = useCallback((format: StarterTemplateFormat) => {
    const template = format === 'bicep' ? bicepStarterTemplate : terraformStarterTemplate;
    const blob = new Blob([template.content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = template.fileName;
    link.click();
    URL.revokeObjectURL(url);
  }, [bicepStarterTemplate, terraformStarterTemplate]);

  const importDriftPlan = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;
    if (files.length > 1) {
      alert(localize(language, {
        en: 'Import a single Azure what-if or Terraform plan JSON file at a time.',
        ja: 'Azure what-if または Terraform plan の JSON は一度に1ファイルだけ取り込んでください。',
      }));
      event.target.value = '';
      return;
    }

    const file = files[0];
    if (!file.name.toLowerCase().endsWith('.json')) {
      alert(localize(language, {
        en: 'Only JSON deployment-plan files are supported.',
        ja: '対応しているのは JSON のデプロイ プラン ファイルのみです。',
      }));
      event.target.value = '';
      return;
    }
    if (file.size > 5 * 1024 * 1024) {
      alert(localize(language, {
        en: 'Deployment-plan files must be 5 MB or smaller.',
        ja: 'デプロイ プラン ファイルは 5 MB 以下にしてください。',
      }));
      event.target.value = '';
      return;
    }

    try {
      const summary = parseDeploymentPlan(file.name, await file.text(), new Date().toISOString());
      setDriftPlanSummary(summary);
    } catch (error: any) {
      console.error('Deployment plan import failed:', error);
      alert(localize(language, {
        en: `Failed to import deployment plan: ${error.message}`,
        ja: `デプロイ プランの取り込みに失敗しました: ${error.message}`,
      }));
    } finally {
      event.target.value = '';
    }
  }, [language]);

  // Load version from URL hash (for "Open in New Tab" feature)
  useEffect(() => {
    const hash = window.location.hash;
    const clearVersionHash = () => {
      window.history.replaceState(null, '', window.location.pathname + window.location.search);
    };
    if (hash.startsWith('#version-id-')) {
      try {
        const versionId = decodeURIComponent(hash.substring(12));
        if (!versionId) throw new Error('Version id is missing');
        getVersion(versionId)
          .then((version) => {
            if (!version) throw new Error('Version not found');
            applyFlowObject({
              nodes: version.nodes,
              edges: version.edges,
              metadata: version.metadata,
              workflow: version.workflow,
              architecturePrompt: version.architecturePrompt,
              titleBlockData: version.titleBlockData,
              pricingScenarios: version.pricingScenarios,
              iacBaseline: version.iacBaseline,
            });
            clearVersionHash();
          })
          .catch((error) => {
            console.error('Failed to load version from storage:', error);
            clearVersionHash();
          });
      } catch (error) {
        console.error('Failed to read version id from URL:', error);
        clearVersionHash();
      }
    } else if (hash.startsWith('#version-')) {
      try {
        const encodedData = hash.substring(9); // Remove '#version-'
        const decodedData = decodeUtf8Base64(encodedData);
        const diagramData = JSON.parse(decodedData);
        
        // Apply the diagram data
        applyFlowObject(diagramData);
        
        // Clear the hash
        clearVersionHash();
      } catch (error) {
        console.error('Failed to load version from URL:', error);
        clearVersionHash();
      }
    }
  }, [applyFlowObject]);



  const loadDiagram = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    const input = event.currentTarget;
    const file = event.target.files?.[0];
    if (!file) return;
    if (file.size > 10 * 1024 * 1024) {
      alert(localize(language, {
        en: 'The diagram file is too large. The maximum size is 10 MB.',
        ja: '図のファイルが大きすぎます。最大サイズは10 MBです。',
      }));
      input.value = '';
      return;
    }

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const flow = JSON.parse(e.target?.result as string);
        if (
          flow?.format === 'azurediagarm-ai-architecture'
          && Array.isArray(flow.services)
          && Array.isArray(flow.connections)
        ) {
          if (!handleAIGenerateRef.current) {
            throw new Error('The architecture renderer is not ready');
          }
          cloudSync.reset();
          await handleAIGenerateRef.current(
            flow,
            typeof flow.metadata?.prompt === 'string' ? flow.metadata.prompt : file.name,
            false,
          );
        } else {
          cloudSync.reset();
          applyFlowObject(flow);
        }
      } catch (error) {
        console.error('Error loading diagram:', error);
        alert(t("Error loading diagram file"));
      } finally {
        input.value = '';
      }
    };
    reader.onerror = () => {
      console.error('Error reading diagram file:', reader.error);
      alert(t("Error loading diagram file"));
      input.value = '';
    };
    reader.readAsText(file);
  }, [applyFlowObject, cloudSync, t, language]);

  // Restore a version from history
  const restoreVersion = useCallback((version: DiagramVersion) => {
    try {
      applyFlowObject({
        nodes: version.nodes,
        edges: version.edges,
        titleBlockData: version.titleBlockData || version.metadata,
        workflow: version.workflow || [],
        pricingScenarios: version.pricingScenarios,
        architecturePrompt: version.architecturePrompt || '',
        originalPrompt: version.originalPrompt || version.architecturePrompt || '',
        iacBaseline: version.iacBaseline,
      });
      
      console.log('✅ Version restored successfully');
      trackVersionOperation('restore');
    } catch (error) {
      console.error('Failed to restore version:', error);
      alert(t("Failed to restore version"));
    }
  }, [applyFlowObject, t]);

  // Manual snapshot save handler
  const handleSaveSnapshot = useCallback(async (notes: string) => {
    try {
      await createSnapshot(
        nodes,
        edges,
        titleBlockData.architectureName,
        {
          architecturePrompt,
          originalPrompt: originalPrompt || architecturePrompt || undefined,
          validationScore: validationResult?.overallScore,
          notes: notes || 'Manual snapshot',
          titleBlockData,
          workflow,
          pricingScenarios,
          iacBaseline,
        }
      );
      try {
        await cloudSync.saveSnapshot(notes || 'Manual snapshot');
      } catch (cloudError) {
        console.warn('Cloud snapshot was unavailable; the local snapshot was preserved:', cloudError);
      }
      console.log('✅ Manual snapshot saved successfully');
      trackVersionOperation('save');
    } catch (error) {
      console.error('Failed to save manual snapshot:', error);
      throw error;
    }
  }, [nodes, edges, titleBlockData, architecturePrompt, originalPrompt, validationResult, workflow, pricingScenarios, cloudSync, iacBaseline]);

  const handleAIGenerate = useCallback(async (
    architecture: any,
    prompt: string,
    autoSnapshot: boolean = true,
    preserveExistingLayout: boolean = false,
  ) => {
    try {
      console.log('Generating architecture from:', architecture);
      const isRefinement = preserveExistingLayout && nodes.length > 0;
      const { services, connections, workflow: workflowSteps } = architecture;
      let { groups } = architecture;
      
      if (!Array.isArray(services) || services.length === 0) {
        throw new Error(t("No services were identified in your description. Please try a more detailed description."));
      }
      if (!Array.isArray(connections)) {
        throw new Error('Generated architecture is missing its connections array.');
      }
      if (groups != null && !Array.isArray(groups)) {
        throw new Error('Generated architecture contains an invalid groups value.');
      }

      // ── Guard: remove empty groups & reassign orphaned services ──
      if (groups && Array.isArray(groups) && services && Array.isArray(services)) {
        const groupIds = new Set((groups as any[]).map((g: any) => g.id));
        const populated = new Set<string>();
        for (const s of services) {
          if (s.groupId && groupIds.has(s.groupId)) populated.add(s.groupId);
        }
        const emptyGroupIds = (groups as any[]).filter((g: any) => !populated.has(g.id)).map((g: any) => g.id);
        if (emptyGroupIds.length > 0) {
          console.warn(`⚠️ Removing ${emptyGroupIds.length} empty group(s): ${emptyGroupIds.join(', ')}`);
          groups = (groups as any[]).filter((g: any) => populated.has(g.id));
          // Clear any service groupId that points to a removed group
          for (const s of services) {
            if (s.groupId && !populated.has(s.groupId)) {
              console.warn(`  → Clearing orphaned groupId "${s.groupId}" on service "${s.id}"`);
              s.groupId = null;
            }
          }
        }
      }

      console.log(`Processing ${services.length} services, ${connections?.length || 0} connections, ${groups?.length || 0} groups`);

      // Auto-save snapshot before regenerating (if enabled and there are existing nodes)
      if (autoSnapshot && nodes.length > 0) {
        console.log('📸 Auto-saving snapshot before regeneration...');
        console.log(`Current state: ${nodes.length} nodes, ${edges.length} edges, name: "${titleBlockData.architectureName}"`);
        try {
          await createSnapshot(
            nodes,
            edges,
            titleBlockData.architectureName,
            {
              architecturePrompt: architecturePrompt || 'Previous version',
              originalPrompt: originalPrompt || architecturePrompt || undefined,
              validationScore: validationResult?.overallScore,
              notes: 'Auto-saved before AI regeneration',
              titleBlockData,
              workflow,
              pricingScenarios,
              iacBaseline,
            }
          );
          try {
            await cloudSync.saveSnapshot('Auto-saved before AI regeneration');
          } catch (cloudError) {
            console.warn('Cloud snapshot was unavailable; the local snapshot was preserved:', cloudError);
          }
          console.log('✅ Snapshot saved successfully!');
        } catch (err) {
          console.error('❌ Failed to save snapshot:', err);
        }
      } else {
        console.log('ℹ️ No existing nodes to snapshot');
      }

      // Pick up an architecture name from the AI payload (manifest.title in
      // Both mode) or derive a short title from the prompt so the banner
      // doesn't read "Untitled Architecture" after every generation.
      const incomingName: string | undefined = (architecture?.architectureName && String(architecture.architectureName).trim())
        || deriveTitleFromPrompt(prompt);

      const newNodes: Node[] = [];
      const serviceMap = new Map();

    // Load all required icons first
    const iconCache = new Map();
    
    // Category correction map - AI categorizes services differently than icon folders
    const correctCategory = (serviceType: string, aiCategory: string): string => {
      // Check SERVICE_ICON_MAP first for authoritative category
      const mapping = getServiceIconMapping(serviceType);
      if (mapping) return mapping.category;
      
      const corrections: Record<string, string> = {
        'Azure Functions': 'compute',
        'Logic Apps': 'integration',
        'API Management': 'integration',
      };
      return corrections[serviceType] || aiCategory;
    };
    
    // Load icons in parallel with timeout protection
    console.log(`⏳ Loading icons for ${services.length} services...`);
    const iconLoadingPromises = services.map(async (service: any) => {
      try {
        // Prefer the canonical service type; display names can be ambiguous
        // across Azure and Microsoft Fabric.
        let mapping = getServiceIconMapping(service.type);
        if (!mapping) mapping = getServiceIconMapping(service.name);
        if (mapping) {
          console.log(`  🎯 Found mapping for "${service.name}" (type: ${service.type}): ${mapping.iconFile}`);
          const iconPath = `/Azure_Public_Service_Icons/Icons/${mapping.category}/${mapping.iconFile}.svg`;
          return { 
            serviceId: service.id, 
            icon: {
              name: mapping.displayName,
              path: iconPath,
              category: mapping.category
            }
          };
        }
        
        // SECOND: Fall back to category search if no mapping found
        const correctedCategory = correctCategory(service.type, service.category);
        const icons = await Promise.race([
          loadIconsFromCategory(correctedCategory),
          new Promise<any[]>((_, reject) => 
            setTimeout(() => reject(new Error('Icon loading timeout')), 5000)
          )
        ]);
        
        console.log(`🎨 Loaded ${icons.length} icons for: ${service.name} (${correctedCategory})`);
        
        if (icons.length > 0) {
          // Try to find the best matching icon
          let icon = null;
          
          console.log(`  🔍 Searching for: "${service.type}"`);
          
          // First: Try exact match (case-insensitive)
          icon = icons.find(i => 
            i.name.toLowerCase() === service.type.toLowerCase()
          );
          if (icon) console.log(`  ✅ Exact match: ${icon.name}`);
          
          // Third: Try to match all significant words (skip common words like "Azure", "Service")
          if (!icon) {
            const serviceWords = service.type.toLowerCase()
              .split(/[\s-]+/)
              .filter((w: string) => !['azure', 'service', 'microsoft'].includes(w));
            
            icon = icons.find(i => {
              const iconWords = i.name.toLowerCase().split(/[\s-]+/);
              return serviceWords.every((word: string) => 
                iconWords.some((iw: string) => iw.includes(word) || word.includes(iw))
              );
            });
            if (icon) console.log(`  ✅ Multi-word match: ${icon.name}`);
          }
          
          // Fourth: Try matching just the primary word (first meaningful word)
          if (!icon) {
            const primaryWord = service.type.toLowerCase()
              .split(/[\s-]+/)
              .find((w: string) => !['azure', 'microsoft', 'service'].includes(w));
            
            if (primaryWord) {
              icon = icons.find(i => 
                i.name.toLowerCase().includes(primaryWord)
              );
              if (icon) console.log(`  ✅ Primary word match: ${icon.name}`);
            }
          }
          
          // Fifth: Fallback to first icon in category
          if (!icon) {
            icon = icons[0];
            console.log(`  ⚠️ Using fallback: ${icon.name}`);
          }
          
          return { serviceId: service.id, icon };
        } else {
          console.warn(`  ❌ No icons found for: ${service.name}`);
          return { serviceId: service.id, icon: null };
        }
      } catch (error) {
        console.error(`  ❌ Error loading icon for ${service.name}:`, error);
        return { serviceId: service.id, icon: null };
      }
    });
    
    // Wait for all icon loading with overall timeout
    const iconResults = await Promise.race([
      Promise.all(iconLoadingPromises),
      new Promise<any[]>((_, reject) => 
        setTimeout(() => reject(new Error('Overall icon loading timeout')), 15000)
      )
    ]).catch(error => {
      console.error('Icon loading failed:', error);
      return services.map((s: any) => ({ serviceId: s.id, icon: null }));
    });
    
    // Build icon cache from results
    iconResults.forEach((result: any) => {
      if (result.icon) {
        iconCache.set(result.serviceId, result.icon);
      }
    });
    
    console.log(`✅ Icon loading complete. Loaded ${iconCache.size}/${services.length} icons`);

    // ============================================================================
    // LAYOUT ENGINE: Calculate optimal positions using selected algorithm
    // ============================================================================
    const engineLabel = layoutEngine === 'elk' ? 'ELK' : 'Dagre';
    console.log(`📐 Calculating layout with ${engineLabel} algorithm...`);
    console.log('📦 Groups before layout:', groups);

    let positionedServices: any[];
    let positionedGroups: any[];

    if (layoutEngine === 'elk') {
      const result = await elkLayoutArchitecture(
        services,
        connections,
        groups || [],
        { direction: 'LR' }
      );
      positionedServices = result.services;
      positionedGroups = result.groups;
    } else {
      const result = layoutArchitecture(
        services,
        connections,
        groups || [],
        { direction: 'LR' }
      );
      positionedServices = result.services;
      positionedGroups = result.groups;
    }
    console.log('📦 Positioned groups after layout:', positionedGroups);

    // Create group nodes with calculated positions and sizes
    if (positionedGroups && positionedGroups.length > 0) {
      positionedGroups.forEach((group: any) => {
        const groupNode: Node = {
          id: group.id,
          type: 'groupNode',
          position: group.position,
          data: {
            label: group.label || group.id || 'Unnamed Group',
          },
          style: {
            width: group.width,
            height: group.height,
          },
        };
        newNodes.push(groupNode);
      });
    }

    // Create service nodes with calculated positions
    positionedServices.forEach((service: any) => {
      const icon = iconCache.get(service.id);
      
      const node: Node = {
        id: service.id,
        type: 'azureNode',
        position: service.position,  // ✅ Use position from layout engine
        data: {
          label: service.name,
          serviceName: service.type || service.name,
          category: icon?.category || service.category,
          iconPath: icon?.path || '',
        },
        parentNode: service.groupId || undefined,  // Link to group if exists
        extent: service.groupId ? 'parent' : undefined,  // Keep within parent bounds
      };

      newNodes.push(node);
      serviceMap.set(service.id, node);
    });

    // Existing services/groups retain their manually edited geometry during a
    // refinement. New elements use the generated layout positions.
    const finalNodes = isRefinement
      ? preserveManualLayout(nodes, newNodes)
      : newNodes;

    // Build absolute position map for smart edge routing
    // after manual positions have been restored.
    const absolutePositions = buildAbsolutePositionMap(finalNodes);

    // Smart handle selection based on relative node positions
    // Picks handles that create the shortest, least-crossing edge paths
    const getConnectionPositions = (sourceId: string, targetId: string, _conn: any) => {
      return selectHorizontalConnectionHandles(absolutePositions, sourceId, targetId);
    };

    // Function to determine arrow direction based on edge label
    const determineEdgeDirection = (label: string): { direction: 'forward' | 'reverse' | 'bidirectional', markerEnd?: any, markerStart?: any, flowMode: 'directional' | 'pulse' } => {
      const lowerLabel = label.toLowerCase();
      
      // Keywords that indicate reverse flow
      const reverseKeywords = ['response', 'callback', 'return', 'acknowledge', 'ack', 'reply'];
      
      // Keywords that indicate bidirectional flow
      const bidirectionalKeywords = ['sync', 'bidirectional', 'two-way', 'exchange', 'communicate'];
      
      // Check for bidirectional
      if (bidirectionalKeywords.some(keyword => lowerLabel.includes(keyword))) {
        return {
          direction: 'bidirectional',
          markerEnd: { type: MarkerType.ArrowClosed, color: '#0078d4' },
          markerStart: { type: MarkerType.ArrowClosed, color: '#0078d4' },
          flowMode: 'pulse',
        };
      }
      
      // Check for reverse
      if (reverseKeywords.some(keyword => lowerLabel.includes(keyword))) {
        return {
          direction: 'reverse',
          markerStart: { type: MarkerType.ArrowClosed, color: '#0078d4' },
          markerEnd: undefined,
          flowMode: 'directional',
        };
      }
      
      // Default to forward
      return {
        direction: 'forward',
        markerEnd: { type: MarkerType.ArrowClosed, color: '#0078d4' },
        markerStart: undefined,
        flowMode: 'directional',
      };
    };

    // Create edges from connections
    const newEdges: Edge[] = connections.map((conn: any, index: number) => {
      const positions = getConnectionPositions(conn.from, conn.to, conn);
      
      // Determine edge direction based on label
      const edgeDirection = determineEdgeDirection(conn.label || '');
      
      // Determine edge style based on connection type
      const connectionType = conn.type || 'sync';
      let edgeStyle = {};
      let baseFlowAnimated = true;
      
      switch (connectionType) {
        case 'async':
          // Dashed line for asynchronous
          edgeStyle = { strokeDasharray: '5, 5' };
          baseFlowAnimated = true;
          break;
        case 'optional':
          // Dotted line for optional
          edgeStyle = { strokeDasharray: '2, 4', opacity: 0.6 };
          baseFlowAnimated = false;
          break;
        case 'sync':
        default:
          // Solid line for synchronous (default)
          edgeStyle = {};
          baseFlowAnimated = true;
          break;
      }

      const flowAnimated = animateConnections && baseFlowAnimated;
      
      return {
        id: `edge-${index}`,
        source: conn.from,
        target: conn.to,
        sourceHandle: positions.sourceHandle,
        targetHandle: positions.targetHandle,
        animated: false,
        type: 'editableEdge',
        label: conn.label || '',
        markerEnd: edgeDirection.markerEnd,
        markerStart: edgeDirection.markerStart,
        labelStyle: { fontSize: 14, fill: '#333', fontWeight: 'bold' },
        labelBgStyle: { fill: 'white', fillOpacity: 0.9, stroke: '#000', strokeWidth: 1.5 },
        style: edgeStyle,
        data: {
          connectionType,
          direction: edgeDirection.direction,
          baseFlowAnimated,
          flowAnimated,
          flowMode: edgeDirection.flowMode,
          pathStyle: layoutEdgeStyle,
          onLabelChange: handleEdgeLabelChange,
          onLabelOffsetChange: handleEdgeLabelOffsetChange,
          labelOffsetX: 0,
          labelOffsetY: 0,
        },
      };
    });

    // Add the new nodes and edges
    console.log(`Setting ${finalNodes.length} nodes and ${newEdges.length} edges`);
    const pricingRunId = ++aiPricingRunRef.current;
    setValidationResult(null);
    setValidationHandoff(null);
    feedbackAfterValidationRef.current = false;
    setLastReferenceArchitecture(architecture?.__referenceArchitecture ?? null);
    setNodes(finalNodes);
    setEdges(newEdges);
    setArchitecturePrompt(prompt);
    if (!isRefinement) setOriginalPrompt(prompt);
    setWorkflow(Array.isArray(workflowSteps) ? workflowSteps : []);
    if (incomingName && incomingName !== 'Untitled Architecture') {
      setTitleBlockData((prev) => ({ ...prev, architectureName: incomingName }));
    }

    // Set the model badge from metrics
    if (architecture.metrics) {
      const modelKey = Object.keys(MODEL_CONFIG).find(
        k => DEPLOYMENT_NAMES[k as ModelType] === architecture.metrics!.model
      );
      const displayName = modelKey
        ? MODEL_CONFIG[modelKey as keyof typeof MODEL_CONFIG].displayName
        : architecture.metrics.model || 'AI';
      setGeneratedWithModel({ name: displayName, timeMs: architecture.metrics.elapsedTimeMs });
    } else {
      setGeneratedWithModel(null);
    }

    // Initialize only nodes that do not already carry editor-owned pricing.
    const currentRegion = getActiveRegion();
    const finalNodesById = new Map(finalNodes.map(node => [node.id, node]));
    const pricingTargets = services.filter(
      (service: any) => !finalNodesById.get(service.id)?.data?.pricing,
    );
    console.log(`💰 Initializing pricing for ${pricingTargets.length} services in region: ${currentRegion}`);

    const pricingPromises = pricingTargets.map(async (service: any) => {
      const serviceType = String(service.type || service.name);
      console.log(`  → Fetching pricing for: ${service.name} (type: ${service.type}, ID: ${service.id})`);
      const pricing = await initializeNodePricing(serviceType, currentRegion);
      console.log(`  ${pricing ? '✅' : '❌'} Pricing result for ${service.name}:`, pricing ? 'Found' : 'Not found');
      return { id: service.id, serviceType, pricing };
    });
    
    Promise.all(pricingPromises)
      .then(pricingResults => {
        if (pricingRunId !== aiPricingRunRef.current) return;
        console.log(`📊 Pricing results ready, updating ${pricingResults.length} nodes`);
        const resultsWithPricing = pricingResults.filter(r => r.pricing);
        console.log(`  → ${resultsWithPricing.length}/${pricingResults.length} nodes have pricing data`);
        
        setNodes((nds) => 
          nds.map(node => {
            const result = pricingResults.find(r => r.id === node.id);
            const currentServiceType = String(node.data.serviceName || node.data.label || '');
            if (
              result?.pricing
              && !node.data.pricing
              && currentServiceType === result.serviceType
            ) {
              console.log(`  💵 Adding pricing to node ${node.id}:`, result.pricing.estimatedCost);
              return { ...node, data: { ...node.data, pricing: result.pricing } };
            }
            return node;
          })
        );
        console.log(`✅ Pricing initialization complete`);
      })
      .catch(err => console.error('❌ Failed to initialize pricing for AI nodes:', err));

    // Collapse all panels to maximize diagram view
    setPanelsCollapsedSignal(prev => prev + 1);

    // Track architecture generation telemetry
    const aiMetrics = (architecture as any)?.metrics || {};
    const aiIntegrity = (architecture as any)?.integrity || {};
    trackArchitectureGeneration({
      model: aiMetrics.model,
      reasoningEffort: aiMetrics.reasoningEffort,
      promptLength: prompt?.length,
      serviceCount: services?.length,
      connectionCount: connections?.length,
      groupCount: groups?.length,
      workflowStepCount: workflowSteps?.length,
      elapsedTimeMs: aiMetrics.elapsedTimeMs,
      totalTokens: aiMetrics.totalTokens,
      isModification: isRefinement,
      orphanCount: aiIntegrity.orphanCount,
      repairedEdges: aiIntegrity.repairedEdges,
      droppedEdges: aiIntegrity.droppedEdges,
    });

    const handoffContext = {
      source: isRefinement ? 'modification' as const : 'generation' as const,
      serviceCount: services.length,
    };
    setValidationHandoff(handoffContext);

    // ── Success-moment feedback ask ──────────────────────────────────────
    // After the 2nd successful generation this session, surface the one-click
    // toast — the user now has a real opinion. Fires once and only if they
    // haven't already given feedback.
    generationCountRef.current += 1;
    let feedbackAlreadyDone = false;
    try {
      feedbackAlreadyDone = sessionStorage.getItem(FEEDBACK_DONE_KEY) === '1';
    } catch {
      /* sessionStorage unavailable — ignore */
    }
    if (!feedbackAlreadyDone && generationCountRef.current === 2 && !isFeedbackModalOpen) {
      feedbackAfterValidationRef.current = true;
    }

    // A refinement keeps the user's pan/zoom. Only frame a newly generated
    // diagram, where no prior editorial viewport exists.
    if (!isRefinement) {
      setTimeout(() => {
        reactFlowInstance?.fitView({ padding: 0.2 });
      }, 100);
    }
    } catch (error) {
      console.error('Error in handleAIGenerate:', error);
      alert(t("Failed to generate diagram. Check console for details."));
      throw error;
    }
  }, [setNodes, setEdges, reactFlowInstance, nodes, edges, titleBlockData, architecturePrompt, originalPrompt, validationResult, workflow, pricingScenarios, isFeedbackModalOpen, animateConnections, layoutEdgeStyle, t, cloudSync, iacBaseline]);
  handleAIGenerateRef.current = handleAIGenerate;

  // ── az prototype import ──────────────────────────────────────────────
  // (Import az prototype UI removed — feature unused.)

  /** Detect IaC format from file extension and content */
  const detectIaCFormat = useCallback((filename: string, text: string): { format: IaCFormat; label: string } | null => {
    const ext = filename.split('.').pop()?.toLowerCase();

    if (ext === 'bicep') {
      if (/\b(resource|module)\b/.test(text)) {
        return { format: 'bicep', label: 'Bicep' };
      }
      alert(localize(language, {
        en: `Invalid Bicep file: "${filename}" does not contain resource or module declarations.`,
        ja: `無効なBicepファイルです: "${filename}" にresourceまたはmodule宣言がありません。`,
      }));
      return null;
    }

    if (ext === 'tf') {
      if (/\b(resource|provider|module|data)\b/.test(text)) {
        return { format: 'terraform-hcl', label: 'Terraform' };
      }
      alert(localize(language, {
        en: `Invalid Terraform file: "${filename}" does not contain resource or provider blocks.`,
        ja: `無効なTerraformファイルです: "${filename}" にresourceまたはprovider blockがありません。`,
      }));
      return null;
    }

    if (ext === 'json') {
      try {
        const json = JSON.parse(text);
        // Terraform state file
        if (json.version !== undefined && json.resources && json.terraform_version) {
          return { format: 'terraform-state', label: 'Terraform State' };
        }
        // ARM template
        if (json.$schema && json.resources) {
          return { format: 'arm', label: 'ARM' };
        }
        alert(localize(language, {
          en: `Unrecognized JSON file: "${filename}". Expected an ARM template ($schema + resources) or Terraform state file.`,
          ja: `認識できないJSONファイルです: "${filename}"。ARM template（$schema + resources）またはTerraform state fileが必要です。`,
        }));
        return null;
      } catch {
        alert(localize(language, {
          en: `Invalid JSON in "${filename}".`,
          ja: `"${filename}" のJSONが無効です。`,
        }));
        return null;
      }
    }

    alert(localize(language, {
      en: `Unsupported file type: .${ext}. Supported formats: .bicep, .tf, .json`,
      ja: `未対応のファイル形式です: .${ext}。対応形式: .bicep、.tf、.json`,
    }));
    return null;
  }, [language]);

  const uploadTemplate = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;
    const selectedFiles = Array.from(files);
    const totalBytes = selectedFiles.reduce((sum, file) => sum + file.size, 0);
    if (selectedFiles.length > 20 || totalBytes > 10 * 1024 * 1024) {
      alert(localize(language, {
        en: 'Select no more than 20 template files with a combined size of 10 MB or less.',
        ja: 'templateファイルは20個以下、合計10 MB以下で選択してください。',
      }));
      event.target.value = '';
      return;
    }

    setIsImportingTemplate(true);

    try {
      // Read all selected files
      const fileContents: { name: string; text: string }[] = [];
      for (const file of selectedFiles) {
        const text = await file.text();
        fileContents.push({ name: file.name, text });
      }

      // Detect format from the first file
      const detection = detectIaCFormat(fileContents[0].name, fileContents[0].text);
      if (!detection) {
        setIsImportingTemplate(false);
        event.target.value = '';
        return;
      }

      setImportFormatLabel(detection.label);
      const filenames = fileContents.map(f => f.name);
      const extraCount = filenames.length > 1 ? ` (+${filenames.length - 1} files)` : '';
      const importedAt = new Date().toISOString();
      const baseline = buildIaCBaseline({
        format: detection.format,
        files: fileContents.map((file, index) => ({
          name: file.name,
          text: file.text,
          size: selectedFiles[index]?.size,
        })),
        importedAt,
      });

      // ── ARM: deterministic extraction (faithful mirror of the template) ──
      // Parse resources + real dependsOn/resourceId edges directly instead of
      // asking the LLM to interpret. Falls back to the LLM path only when the
      // template contains no recognizable resources.
      if (detection.format === 'arm') {
        const template = JSON.parse(fileContents[0].text);
        const { architecture, coverage } = extractArchitectureFromArm(template);
        if (architecture.services.length > 0) {
          clearSourceModel();
          const promptLabel = `ARM Template: ${filenames[0]}${extraCount} — ${summarizeCoverage(coverage)}`;
          trackTemplateImport('arm', filenames[0], filenames.length);
          await handleAIGenerate(architecture, promptLabel);
          setIaCBaseline(baseline);
          setDriftPlanSummary(null);
          return;
        }
        console.warn('Deterministic ARM extraction found no mappable resources; falling back to LLM.');
      }

      let content: string | object;
      if (detection.format === 'arm') {
        content = JSON.parse(fileContents[0].text);
      } else if (detection.format === 'terraform-state') {
        content = JSON.parse(fileContents[0].text);
      } else {
        // Bicep or Terraform HCL — concatenate multiple files with headers
        content = fileContents
          .map(f => `// === ${f.name} ===\n${f.text}`)
          .join('\n\n');
      }

      const { generateArchitectureFromIaC } = await import('./services/azureOpenAI');
      const result = await generateArchitectureFromIaC({
        format: detection.format,
        content,
        filenames,
      }, language);

      clearSourceModel();

      // Build descriptive prompt label
      const promptLabel = localize(language, {
        en: `${detection.label} Template: ${filenames[0]}${extraCount}`,
        ja: `${detection.label}テンプレート: ${filenames[0]}${extraCount}`,
      });

      trackTemplateImport(detection.format, filenames[0], filenames.length);
      await handleAIGenerate(result, promptLabel);
      setIaCBaseline(baseline);
      setDriftPlanSummary(null);
    } catch (error: any) {
      console.error('Template import error:', error);
      alert(localize(language, {
        en: `Failed to import template: ${error.message}`,
        ja: `テンプレートのインポートに失敗しました: ${error.message}`,
      }));
    } finally {
      setIsImportingTemplate(false);
      setImportFormatLabel('Template');
      event.target.value = '';
    }
  }, [handleAIGenerate, detectIaCFormat, language]);

  // Reverse-engineer a live Azure resource group into a diagram via Azure
  // Resource Graph (Reader-sufficient, returns only real top-level resources).
  // Edges are inferred from resource IDs embedded in properties. The same
  // deterministic mapping is used as the file-based ARM import.
  const importFromAzure = useCallback(async (subscriptionId: string, resourceGroup: string) => {
    const resources = await getAzureResources(subscriptionId, resourceGroup);
    const { architecture, coverage } = buildArchitectureFromResources(resources);
    if (architecture.services.length === 0) {
      throw new Error(localize(language, {
        en: 'No mappable Azure resources were found in this resource group.',
        ja: 'このResource Groupには図に変換できるAzureリソースが見つかりませんでした。',
      }));
    }
    clearSourceModel();
    const promptLabel = localize(language, {
      en: `Azure Resource Group: ${resourceGroup} — ${summarizeCoverage(coverage)}`,
      ja: `Azure Resource Group: ${resourceGroup} — ${summarizeCoverage(coverage)}`,
    });
    trackTemplateImport('arm', `rg:${resourceGroup}`, 1);
    await handleAIGenerate(architecture, promptLabel);
    setIaCBaseline(null);
    setDriftPlanSummary(null);
  }, [handleAIGenerate, language]);

  const handleAlign = useCallback((type: string) => {
    const selectedNodes = nodes.filter(n => n.selected);
    if (selectedNodes.length < 2) return;

    const updatedNodes = [...nodes];
    
    switch (type) {
      case 'left': {
        const minX = Math.min(...selectedNodes.map(n => n.position.x));
        selectedNodes.forEach(node => {
          const idx = updatedNodes.findIndex(n => n.id === node.id);
          updatedNodes[idx] = { ...updatedNodes[idx], position: { ...updatedNodes[idx].position, x: minX } };
        });
        break;
      }
      case 'right': {
        const maxX = Math.max(...selectedNodes.map(n => n.position.x + (n.width || 150)));
        selectedNodes.forEach(node => {
          const idx = updatedNodes.findIndex(n => n.id === node.id);
          const nodeWidth = node.width || 150;
          updatedNodes[idx] = { ...updatedNodes[idx], position: { ...updatedNodes[idx].position, x: maxX - nodeWidth } };
        });
        break;
      }
      case 'center-h': {
        const minX = Math.min(...selectedNodes.map(n => n.position.x));
        const maxX = Math.max(...selectedNodes.map(n => n.position.x + (n.width || 150)));
        const centerX = (minX + maxX) / 2;
        selectedNodes.forEach(node => {
          const idx = updatedNodes.findIndex(n => n.id === node.id);
          const nodeWidth = node.width || 150;
          updatedNodes[idx] = { ...updatedNodes[idx], position: { ...updatedNodes[idx].position, x: centerX - nodeWidth / 2 } };
        });
        break;
      }
      case 'top': {
        const minY = Math.min(...selectedNodes.map(n => n.position.y));
        selectedNodes.forEach(node => {
          const idx = updatedNodes.findIndex(n => n.id === node.id);
          updatedNodes[idx] = { ...updatedNodes[idx], position: { ...updatedNodes[idx].position, y: minY } };
        });
        break;
      }
      case 'bottom': {
        const maxY = Math.max(...selectedNodes.map(n => n.position.y + (n.height || 100)));
        selectedNodes.forEach(node => {
          const idx = updatedNodes.findIndex(n => n.id === node.id);
          const nodeHeight = node.height || 100;
          updatedNodes[idx] = { ...updatedNodes[idx], position: { ...updatedNodes[idx].position, y: maxY - nodeHeight } };
        });
        break;
      }
      case 'center-v': {
        const minY = Math.min(...selectedNodes.map(n => n.position.y));
        const maxY = Math.max(...selectedNodes.map(n => n.position.y + (n.height || 100)));
        const centerY = (minY + maxY) / 2;
        selectedNodes.forEach(node => {
          const idx = updatedNodes.findIndex(n => n.id === node.id);
          const nodeHeight = node.height || 100;
          updatedNodes[idx] = { ...updatedNodes[idx], position: { ...updatedNodes[idx].position, y: centerY - nodeHeight / 2 } };
        });
        break;
      }
      case 'distribute-h': {
        const sorted = [...selectedNodes].sort((a, b) => a.position.x - b.position.x);
        const minX = sorted[0].position.x;
        const maxX = sorted[sorted.length - 1].position.x;
        const spacing = (maxX - minX) / (sorted.length - 1);
        sorted.forEach((node, i) => {
          if (i === 0 || i === sorted.length - 1) return;
          const idx = updatedNodes.findIndex(n => n.id === node.id);
          updatedNodes[idx] = { ...updatedNodes[idx], position: { ...updatedNodes[idx].position, x: minX + spacing * i } };
        });
        break;
      }
      case 'distribute-v': {
        const sorted = [...selectedNodes].sort((a, b) => a.position.y - b.position.y);
        const minY = sorted[0].position.y;
        const maxY = sorted[sorted.length - 1].position.y;
        const spacing = (maxY - minY) / (sorted.length - 1);
        sorted.forEach((node, i) => {
          if (i === 0 || i === sorted.length - 1) return;
          const idx = updatedNodes.findIndex(n => n.id === node.id);
          updatedNodes[idx] = { ...updatedNodes[idx], position: { ...updatedNodes[idx].position, y: minY + spacing * i } };
        });
        break;
      }
    }

    setNodes(updatedNodes);
  }, [cancelPendingPricingEditorOpen, nodes, setNodes]);

  // Premium Feature Handlers
  const handleValidateArchitecture = useCallback(async () => {
    // Group boxes are nodes too, but the validator only reasons about Azure
    // services — running it on a canvas that has none wastes an AI call.
    if (nodes.filter(n => n.type === 'azureNode').length === 0) {
      alert(t("Please create an architecture diagram first."));
      return;
    }

    setValidationHandoff(null);

    // Capture diagram snapshot BEFORE opening the modal overlay
    let diagramImageDataUrl: string | undefined;
    if (reactFlowWrapper.current && reactFlowInstance) {
      const previousViewport = reactFlowInstance.getViewport();
      try {
        reactFlowInstance.fitView({ padding: 0.2, duration: 0 });
        // Brief delay for fitView to settle before capture
        await new Promise(resolve => setTimeout(resolve, 400));
        const isDark = document.body.classList.contains('dark-mode');
        diagramImageDataUrl = await captureDiagramAsPng(reactFlowWrapper.current, {
          backgroundColor: isDark ? '#1a1a2e' : '#f8fafc',
        });
        console.log('\uD83D\uDCF8 Diagram snapshot captured for validation report');
      } catch (err) {
        console.warn('Could not capture diagram snapshot:', err);
      } finally {
        // fitView only exists to frame the capture, so put the canvas back
        // where the user left it instead of silently zooming their view.
        reactFlowInstance.setViewport(previousViewport, { duration: 0 });
      }
    }

    // Now show the modal and start validation
    setIsValidating(true);
    setIsValidationModalOpen(true);

    try {
      // Extract services data
      const services = nodes
        .filter(n => n.type === 'azureNode')
        .map(n => ({
          name: n.data.label || n.data.serviceName || 'Unknown Service',
          type: n.data.serviceName || n.data.label || 'Unknown',
          category: n.data.category || 'General',
        }));

      // Extract connections
      const connections = edges.map(e => ({
        from: nodes.find(n => n.id === e.source)?.data?.label || e.source,
        to: nodes.find(n => n.id === e.target)?.data?.label || e.target,
        label: String(e.label || ''),
      }));

      // Extract groups
      const groups = nodes
        .filter(n => n.type === 'groupNode')
        .map(n => ({
          name: n.data.label || 'Group',
          services: nodes
            .filter(child => child.parentNode === n.id)
            .map(child => child.data.label || child.data.serviceName || 'Unknown'),
        }));

      const result = await validateArchitecture(
        services,
        connections,
        groups,
        architecturePrompt || titleBlockData.architectureName,
        undefined,
        language,
      );

      // Attach diagram snapshot to results
      if (diagramImageDataUrl) {
        result.diagramImageDataUrl = diagramImageDataUrl;
      }
      setValidationResult(result);
      trackValidation({
        model: result.metrics?.model,
        overallScore: result.overallScore,
        serviceCount: services.length,
        findingCount: result.pillars?.reduce((sum: number, p: any) => sum + (p.findings?.length || 0), 0),
        elapsedTimeMs: result.metrics?.elapsedTimeMs,
      });
      trackValidationFindings({
        source: 'single',
        model: result.metrics?.model,
        overallScore: result.overallScore,
        serviceCount: services.length,
        topics: classifyValidationTopics(result).map(t => ({ id: t.id, label: t.label, pillar: t.pillar, severity: t.severity })),
      });
      if (feedbackAfterValidationRef.current && !isFeedbackModalOpen) {
        feedbackAfterValidationRef.current = false;
        setIsFeedbackToastOpen(true);
      }
      // Collapse panels to maximize diagram view
      setPanelsCollapsedSignal(prev => prev + 1);
    } catch (error: any) {
      console.error('Validation error:', error);
      alert(localize(language, {
        en: `Failed to validate architecture: ${error.message}`,
        ja: `アーキテクチャの検証に失敗しました: ${error.message}`,
      }));
      setIsValidationModalOpen(false);
    } finally {
      setIsValidating(false);
    }
  }, [nodes, edges, architecturePrompt, titleBlockData.architectureName, reactFlowInstance, isFeedbackModalOpen, t, language]);

  const handleValidationHandoffStart = useCallback(() => {
    if (!validationHandoff) return;
    trackValidationHandoff({ action: 'started', ...validationHandoff });
    setValidationHandoff(null);
    setIsFeedbackToastOpen(false);
    void handleValidateArchitecture();
  }, [handleValidateArchitecture, validationHandoff]);

  const handleValidationHandoffDismiss = useCallback(() => {
    if (!validationHandoff) return;
    trackValidationHandoff({ action: 'dismissed', ...validationHandoff });
    setValidationHandoff(null);
    if (feedbackAfterValidationRef.current && !isFeedbackModalOpen) {
      feedbackAfterValidationRef.current = false;
      setIsFeedbackToastOpen(true);
    }
  }, [isFeedbackModalOpen, validationHandoff]);

  const handleGenerateDeploymentGuide = useCallback(async () => {
    // Same as validation: a canvas of group boxes alone has nothing to deploy.
    if (nodes.filter(n => n.type === 'azureNode').length === 0) {
      alert(t("Please create an architecture diagram first."));
      return;
    }

    setIsGeneratingGuide(true);
    setIsDeploymentGuideModalOpen(true);

    try {
      // Extract services data
      const services = nodes
        .filter(n => n.type === 'azureNode')
        .map(n => ({
          name: n.data.label || n.data.serviceName || 'Unknown Service',
          type: n.data.serviceName || n.data.label || 'Unknown',
          category: n.data.category || 'General',
        }));

      // Extract connections
      const connections = edges.map(e => ({
        from: nodes.find(n => n.id === e.source)?.data?.label || e.source,
        to: nodes.find(n => n.id === e.target)?.data?.label || e.target,
        label: String(e.label || ''),
      }));

      // Extract groups
      const groups = nodes
        .filter(n => n.type === 'groupNode')
        .map(n => ({
          name: n.data.label || 'Group',
          services: nodes
            .filter(child => child.parentNode === n.id)
            .map(child => child.data.label || child.data.serviceName || 'Unknown'),
        }));

      const guide = await generateDeploymentGuide(
        services,
        connections,
        groups,
        architecturePrompt || titleBlockData.architectureName,
        totalMonthlyCost,
        language,
      );

      setDeploymentGuide(guide);
      trackDeploymentGuide({
        model: guide.metrics?.model,
        serviceCount: services.length,
        bicepFileCount: guide.bicepTemplates?.length,
        elapsedTimeMs: guide.metrics?.elapsedTimeMs,
      });
    } catch (error: any) {
      console.error('Guide generation error:', error);
      alert(t('error.deploymentGuide', { message: error.message }));
      setIsDeploymentGuideModalOpen(false);
    } finally {
      setIsGeneratingGuide(false);
    }
  }, [nodes, edges, architecturePrompt, titleBlockData.architectureName, totalMonthlyCost, t, language]);

  const toggleToolbarSection = useCallback((sectionId: ToolbarSectionId) => {
    if (sectionId === 'create') setIsModelSettingsOpen(false);
    if (sectionId === 'file') setIsExportMenuOpen(false);
    if (sectionId === 'arrange') {
      setIsLayoutMenuOpen(false);
      setIsBulkSelectMenuOpen(false);
      setIsStylePresetMenuOpen(false);
    }

    setCollapsedToolbarSections((current) => {
      const next = new Set(current);
      if (next.has(sectionId)) next.delete(sectionId);
      else next.add(sectionId);
      return next;
    });
  }, []);

  const activateRibbonTab = useCallback((tabId: RibbonTabId) => {
    setActiveRibbonTab(tabId);
    writeLocalStorage(RIBBON_TAB_STORAGE_KEY, tabId);
    setIsExportMenuOpen(false);
    setIsLayoutMenuOpen(false);
    setIsBulkSelectMenuOpen(false);
    setIsStylePresetMenuOpen(false);
    setIsModelSettingsOpen(false);
  }, []);

  const toolbarSectionHeading = (sectionId: ToolbarSectionId, label: string) => {
    const isCollapsed = collapsedToolbarSections.has(sectionId);
    return (
      <button
        type="button"
        className="toolbar-group-label"
        aria-expanded={!isCollapsed}
        onClick={() => toggleToolbarSection(sectionId)}
        title={localize(language, {
          en: isCollapsed ? 'Expand section' : 'Collapse section',
          ja: isCollapsed ? 'セクションを展開' : 'セクションを折りたたむ',
        })}
      >
        {isCollapsed ? <ChevronRight size={12} aria-hidden="true" /> : <ChevronDown size={12} aria-hidden="true" />}
        <span>{label}</span>
      </button>
    );
  };

  const contextMenuEdge = edgeContextMenu
    ? edges.find((edge) => edge.id === edgeContextMenu.edgeId)
    : undefined;
  const contextMenuNode = nodeContextMenu
    ? nodes.find((node) => node.id === nodeContextMenu.nodeId)
    : undefined;
  const contextMenuNodeIsGroup = contextMenuNode?.type === 'groupNode';
  const contextMenuNodeLabel = String(
    contextMenuNode?.data?.serviceName
      || contextMenuNode?.data?.label
      || localize(language, { en: 'Selected item', ja: '選択した項目' }),
  );
  const contextMenuContainedCount = contextMenuNodeIsGroup && contextMenuNode
    ? Math.max(
        0,
        collectNodeAndDescendantIds(nodes, [contextMenuNode.id]).size - 1,
      )
    : 0;
  const ribbonTabs: Array<{ id: RibbonTabId; label: string }> = [
    { id: 'home', label: localize(language, { en: 'Home', ja: 'ホーム' }) },
    { id: 'create', label: localize(language, { en: 'Create', ja: '作成' }) },
    { id: 'design', label: localize(language, { en: 'Design', ja: 'デザイン' }) },
    { id: 'review', label: localize(language, { en: 'Review', ja: 'レビュー' }) },
  ];
  const handleRibbonTabKeyDown = (
    event: React.KeyboardEvent<HTMLButtonElement>,
    tabIndex: number,
  ) => {
    let nextIndex: number | null = null;
    if (event.key === 'ArrowRight') nextIndex = (tabIndex + 1) % ribbonTabs.length;
    if (event.key === 'ArrowLeft') nextIndex = (tabIndex - 1 + ribbonTabs.length) % ribbonTabs.length;
    if (event.key === 'Home') nextIndex = 0;
    if (event.key === 'End') nextIndex = ribbonTabs.length - 1;
    if (nextIndex === null) return;

    event.preventDefault();
    const nextTab = ribbonTabs[nextIndex];
    activateRibbonTab(nextTab.id);
    window.requestAnimationFrame(() => {
      const nextTabElement = document.getElementById(`ribbon-tab-${nextTab.id}`);
      nextTabElement?.focus();
      nextTabElement?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    });
  };
  const cloudSyncLabel = (() => {
    switch (cloudSync.status) {
      case 'saving':
        return localize(language, { en: 'Saving...', ja: '保存中...' });
      case 'saved':
        return localize(language, { en: 'Cloud saved', ja: 'クラウド保存済み' });
      case 'readonly':
        return localize(language, { en: 'Cloud read-only', ja: 'クラウド閲覧のみ' });
      case 'conflict':
        return localize(language, { en: 'Sync conflict', ja: '同期競合' });
      case 'offline':
      case 'unavailable':
      case 'error':
        return localize(language, { en: 'Local only', ja: 'ローカルのみ' });
      default:
        return localize(language, { en: 'Cloud', ja: 'クラウド' });
    }
  })();

  return (
    <div className="app">
      <header className={`app-header${isHeaderCollapsed ? ' header-collapsed' : ''}`}>
        <div className="header-content">
          <div className="header-brand">
            <div className="microsoft-logo" role="img" aria-label={t("Microsoft")}>
              <span className="microsoft-symbol" aria-hidden="true">
                <span className="microsoft-square microsoft-square-red" />
                <span className="microsoft-square microsoft-square-green" />
                <span className="microsoft-square microsoft-square-blue" />
                <span className="microsoft-square microsoft-square-yellow" />
              </span>
              <span className="microsoft-wordmark" aria-hidden="true">Microsoft</span>
            </div>
            <h1>{t("Azure Architecture Diagram Builder")}</h1>
          </div>
          <div
            className="header-actions-wrapper"
            id="application-toolbar"
            role="region"
            aria-label={t('toolbar.label')}
          >
            <div
              className="ribbon-tabs"
              role="tablist"
              aria-label={localize(language, { en: 'Ribbon tabs', ja: 'リボンタブ' })}
            >
              {ribbonTabs.map((tab, tabIndex) => (
                <button
                  key={tab.id}
                  id={`ribbon-tab-${tab.id}`}
                  type="button"
                  role="tab"
                  className={`ribbon-tab${activeRibbonTab === tab.id ? ' active' : ''}`}
                  aria-selected={activeRibbonTab === tab.id}
                  aria-controls="ribbon-command-strip"
                  tabIndex={activeRibbonTab === tab.id ? 0 : -1}
                  onClick={() => activateRibbonTab(tab.id)}
                  onKeyDown={(event) => handleRibbonTabKeyDown(event, tabIndex)}
                >
                  {tab.label}
                </button>
              ))}
            </div>
            <div
              id="ribbon-command-strip"
              className="ribbon-command-strip"
              role="tabpanel"
              aria-labelledby={`ribbon-tab-${activeRibbonTab}`}
            >
            {/* Row 1: Context, creation, import, file, and workspace actions */}
            <div className="header-actions">
              <div
                hidden={activeRibbonTab !== 'home'}
                className={`toolbar-group toolbar-group--labeled${collapsedToolbarSections.has('context') ? ' toolbar-group-collapsed' : ''}`}
                data-label={localize(language, { en: 'Pricing estimate region', ja: '料金見積リージョン' })}
                aria-label={localize(language, { en: 'Pricing estimate settings', ja: '料金見積の設定' })}
              >
                {toolbarSectionHeading('context', t('pricing.regionLabel'))}
                <RegionSelector
                  isActive={
                    activeRibbonTab === 'home'
                    && !isHeaderCollapsed
                    && !collapsedToolbarSections.has('context')
                  }
                  onRegionChange={handleRegionChange}
                />
                {hasCostDisplayData && (
                  <button
                    className={`cost-visibility-toggle${pricingPrefs.showCostBadges ? '' : ' is-off'}`}
                    onClick={() => setPricingPrefs({ showCostBadges: !pricingPrefs.showCostBadges })}
                    aria-pressed={pricingPrefs.showCostBadges}
                    title={localize(language, {
                      en: pricingPrefs.showCostBadges
                        ? 'Hide indicative cost estimates without changing the diagram style'
                        : 'Show indicative cost estimates',
                      ja: pricingPrefs.showCostBadges
                        ? '図のスタイルを変えずに参考コストを非表示'
                        : '参考コストを表示',
                    })}
                  >
                    {pricingPrefs.showCostBadges ? <Eye size={14} /> : <EyeOff size={14} />}
                    {localize(language, {
                      en: pricingPrefs.showCostBadges ? 'Cost' : 'Cost hidden',
                      ja: pricingPrefs.showCostBadges ? 'コスト' : 'コスト非表示',
                    })}
                  </button>
                )}
                {hasCostReportData && pricingPrefs.showCostBadges && (
                  <>
                    <div
                      className="cost-indicator"
                      title={t('pricing.totalMonthly', {
                        term: pricingMode === 'reserved1yr' ? t('1-year savings plan') : t('pay-as-you-go'),
                      })}
                    >
                      {' '}{t("💰")}{' '}
                      {totalMonthlyCost === 0 ? '$0.00/mo' : formatMonthlyCost(totalMonthlyCost)}
                    </div>
                    <div className="pricing-mode-toggle" role="group" aria-label={t("Pricing term")}>
                      <button
                        className={`pricing-mode-btn${pricingMode === 'payg' ? ' active' : ''}`}
                        onClick={() => setPricingMode('payg')}
                        title={t("Pay-as-you-go list pricing")}
                      >
                        {' '}{t("PAYG")}{' '}</button>
                      <button
                        className={`pricing-mode-btn${pricingMode === 'reserved1yr' ? ' active' : ''}`}
                        onClick={() => setPricingMode('reserved1yr')}
                        title={t("1-year Savings Plan pricing. Uses each meter's real 1-year savings-plan rate where available, otherwise a representative discount on reservation-eligible services. Usage-based services stay at PAYG.")}
                      >
                        {' '}{t("Savings 1yr")}{' '}</button>
                    </div>
                    <button
                      className="pricing-scenario-launch"
                      onClick={() => setIsPricingScenarioModalOpen(true)}
                      title={localize(language, {
                        en: 'Compare development, production, and custom pricing scenarios',
                        ja: '開発、運用、カスタムの料金シナリオを比較',
                      })}
                    >
                      <GitCompare size={14} />
                      <span>{localize(language, { en: 'Scenarios', ja: 'シナリオ' })}</span>
                    </button>
                    {(() => {
                      const f = getPricingFreshness(PRICING_DATA_AS_OF, new Date(), language);
                      return (
                        <div
                          className={`pricing-freshness pricing-freshness--${f.level}`}
                          title={
                            f.isStale
                              ? t('pricing.stale', { age: f.ageLabel, date: f.dateLabel })
                              : t('pricing.current', { age: f.ageLabel, date: f.dateLabel })
                          }
                          role="status"
                        >
                          {f.isStale && <span aria-hidden="true">{t("⚠️")}{' '}</span>}
                          <span className="pricing-freshness-label">{t("as of")}{' '}{f.dateLabel}</span>
                          {f.isStale && (
                            <span className="pricing-freshness-age"> {' '}{t("·")}{' '}{f.ageLabel}</span>
                          )}
                        </div>
                      );
                    })()}
                  </>
                )}
              </div>

              <div
                hidden={activeRibbonTab !== 'create'}
                className={`toolbar-group toolbar-group--labeled${collapsedToolbarSections.has('create') ? ' toolbar-group-collapsed' : ''}`}
                data-label={localize(language, { en: 'Create & AI', ja: '作成・AI' })}
                aria-label={localize(language, { en: 'Create and AI tools', ja: '作成とAIツール' })}
              >
                {toolbarSectionHeading('create', localize(language, { en: 'Create & AI', ja: '作成・AI' }))}
                <button onClick={addGroupBox} className="btn btn-secondary" title={t("Add grouping box")}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <rect x="3" y="3" width="18" height="18" rx="2" strokeDasharray="4 4" />
                  </svg>
                  {' '}{t("Add Group")}{' '}</button>
                <AIArchitectureGenerator 
                  onGenerate={async (arch, prompt, autoSnap, refImageUrl) => {
                    await handleAIGenerate(arch, prompt, autoSnap, nodes.length > 0);
                    clearSourceModel();
                    setReferenceImageUrl(refImageUrl ?? null);
                    setLastBlueprintArchitecture(null);
                  }}
                  onReferenceArchitecture={(ref) => {
                    // Reference mode does not push a topology onto the canvas;
                    // just remember the ref so the toolbar can re-export the PNG.
                    setLastReferenceArchitecture(ref ?? null);
                  }}
                  onBlueprintArchitecture={(bp) => {
                    // Blueprint mode is also PNG-only; stash for re-export.
                    setLastBlueprintArchitecture(bp ?? null);
                  }}
                  currentArchitecture={{
                    nodes,
                    edges,
                    architectureName: titleBlockData.architectureName
                  }}
                />
                <ModelSettingsPopover
                  ref={modelSettingsRef}
                  isOpen={isModelSettingsOpen}
                  onToggle={() => setIsModelSettingsOpen(v => !v)}
                />
                <button
                  className={`btn btn-secondary${isChatOpen ? ' btn-active' : ''}`}
                  onClick={() => setIsChatOpen((v) => !v)}
                  aria-pressed={isChatOpen}
                  title={isChatOpen
                    ? t("Close Architecture Chat")
                    : t("Open Architecture Chat — start or refine your diagram in plain English")}
                >
                  <MessagesSquare size={18} />
                  {' '}{t("Chat")}{' '}</button>
                <button
                  className="btn btn-secondary"
                  onClick={() => setIsCompareModelsOpen(true)}
                  title={t("Compare architecture output across multiple AI models")}
                >
                  <GitCompare size={18} />
                  {' '}{t("Compare Models")}{' '}</button>
              </div>

              <div
                hidden={activeRibbonTab !== 'create'}
                className={`toolbar-group toolbar-group--labeled${collapsedToolbarSections.has('import') ? ' toolbar-group-collapsed' : ''}`}
                data-label={localize(language, { en: 'Import', ja: 'インポート' })}
                aria-label={localize(language, { en: 'Import architecture', ja: 'アーキテクチャのインポート' })}
              >
                {toolbarSectionHeading('import', localize(language, { en: 'Import', ja: 'インポート' }))}
                <label className={`btn btn-secondary${isImportingTemplate ? ' btn-parsing' : ''}`} title={t("Import Bicep, Terraform, or ARM template to generate diagram")}>
                  {isImportingTemplate ? <Loader size={18} className="spin-icon" /> : <FileCode size={18} />}
                  {isImportingTemplate ? t("Parsing...") : t("Import Template")}
                  <input
                    type="file"
                    accept=".json,.bicep,.tf"
                    multiple
                    onChange={uploadTemplate}
                    style={{ display: 'none' }}
                    disabled={isImportingTemplate}
                  />
                </label>
                <button
                  className="btn btn-secondary"
                  onClick={() => setIsAzureImportOpen(true)}
                  title={localize(language, {
                    en: 'Import an accessible Azure Resource Group using your own Azure permissions',
                    ja: '自分のAzure権限でアクセス可能なResource Groupをインポート',
                  })}
                >
                  <DownloadCloud size={18} />
                  {localize(language, {
                    en: 'Import from Azure',
                    ja: 'Azureからインポート',
                  })}
                </button>
              </div>

              <div
                hidden={activeRibbonTab !== 'home'}
                className={`toolbar-group toolbar-group--labeled${collapsedToolbarSections.has('file') ? ' toolbar-group-collapsed' : ''}`}
                data-label={localize(language, { en: 'File & export', ja: 'ファイル・出力' })}
                aria-label={localize(language, { en: 'File and export actions', ja: 'ファイルと出力操作' })}
              >
                {toolbarSectionHeading('file', localize(language, { en: 'File & export', ja: 'ファイル・出力' }))}
                <button onClick={saveDiagram} className="btn btn-secondary" title={t("Save diagram")}>
                  <Save size={18} />
                  {' '}{t("Save")}{' '}</button>

                <label className="btn btn-secondary" title={t("Load diagram")}>
                  <Upload size={18} />
                  {' '}{t("Load")}{' '}<input
                    type="file"
                    accept=".json"
                    onChange={loadDiagram}
                    style={{ display: 'none' }}
                  />
                </label>
                <div className="toolbar-dropdown" ref={exportMenuRef}>
                  <button
                    onClick={() => setIsExportMenuOpen((v) => !v)}
                    className="btn btn-secondary"
                    title={t("Export")}
                    aria-haspopup="menu"
                    aria-expanded={isExportMenuOpen}
                  >
                    <Download size={18} />
                    {' '}{t("Export")}{' '}<ChevronDown size={16} style={{ marginLeft: 2 }} />
                  </button>

                  {isExportMenuOpen && (
                    <div className="toolbar-dropdown-menu toolbar-dropdown-menu--export" role="menu" aria-label={t("Export options")}>
                      <div className="toolbar-dropdown-heading">
                        {localize(language, { en: 'Images & animation', ja: '画像・アニメーション' })}
                      </div>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportDiagram();
                        }}
                        title={t("Export as PNG")}
                      >
                        <Download size={18} />
                        {' '}{t("Export PNG")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        disabled={!lastReferenceArchitecture}
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          if (!lastReferenceArchitecture) return;
                          exportReferenceArchitectureAsPng(lastReferenceArchitecture).catch((err) => {
                            console.error('Editorial PNG export failed:', err);
                            alert(t("Editorial PNG export failed. See console for details."));
                          });
                        }}
                        title={
                          lastReferenceArchitecture
                            ? t("Re-download the publication-style reference-architecture PNG")
                            : t("Generate a diagram in Reference Architecture mode to enable this export")
                        }
                      >
                        <Download size={18} />
                        {' '}{t("Export Editorial PNG")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        disabled={!lastBlueprintArchitecture}
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          if (!lastBlueprintArchitecture) return;
                          const savedLegend = readLocalStorage('aiGenerator.blueprintLegendPosition');
                          const legendPosition =
                            savedLegend === 'bottom' || savedLegend === 'right' || savedLegend === 'auto'
                              ? (savedLegend as 'bottom' | 'right' | 'auto')
                              : 'auto';
                          exportBlueprintArchitectureAsPng(lastBlueprintArchitecture, { legendPosition }).catch((err) => {
                            console.error('Blueprint PNG export failed:', err);
                            alert(t("Blueprint PNG export failed. See console for details."));
                          });
                        }}
                        title={
                          lastBlueprintArchitecture
                            ? t("Re-download the hand-drawn whiteboard-style blueprint PNG")
                            : t("Generate a diagram in Blueprint mode to enable this export")
                        }
                      >
                        <Download size={18} />
                        {' '}{t("Export Blueprint PNG")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportAsSvg();
                        }}
                        title={t("Export as SVG (vector format)")}
                      >
                        <Download size={18} />
                        {' '}{t("Export SVG")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportAsAnimatedSvg();
                        }}
                        title={t("Export as Animated SVG — flowing data-flow arrows (open in a browser to see motion)")}
                      >
                        <Download size={18} />
                        {' '}{t("Export Animated SVG")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        disabled={workflow.length === 0}
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportWorkflowAnimation();
                        }}
                        title={
                          workflow.length > 0
                            ? t("Export an animated SVG that plays the workflow step-by-step with captions")
                            : t("No workflow steps in this diagram to animate")
                        }
                      >
                        <Download size={18} />
                        {' '}{t("Export Workflow Animation")}{' '}</button>
                      <div className="toolbar-dropdown-separator" role="separator" />
                      <div className="toolbar-dropdown-heading">
                        {localize(language, { en: 'Documents & editable formats', ja: 'ドキュメント・編集形式' })}
                      </div>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        disabled={nodes.filter(n => n.type === 'azureNode').length === 0}
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportWorkflowMarkdown();
                        }}
                        title={t("Export the workflow narrative (services, step-by-step flow, connections) as a Markdown file")}
                      >
                        <FileText size={18} />
                        {' '}{t("Export Workflow (Markdown)")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportAsPptx();
                        }}
                        title={t("Export current diagram as a PowerPoint slide (.pptx)")}
                      >
                        <Presentation size={18} />
                        {' '}{t("Export PPTX Slide")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        disabled={nodes.filter(n => n.type === 'azureNode').length === 0}
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportCustomerDeck();
                        }}
                        title={t("Export a customer-ready PowerPoint deck: title, diagram, services, plus WAF review and cost estimate when available")}
                      >
                        <Presentation size={18} />
                        {t("Export Customer Deck (PPTX)")}
                      </button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportAsDrawio();
                        }}
                        title={t("Export for Draw.io / diagrams.net (editable diagram format)")}
                      >
                        <Download size={18} />
                        {' '}{t("Export Draw.io")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        disabled={nodes.filter(n => n.type === 'azureNode').length === 0}
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportAsVsdx();
                        }}
                        title={t("Export a native Visio drawing (.vsdx). Opens in desktop Visio and Visio for the web; also importable into diagrams.net. Generic editable shapes + connectors.")}
                      >
                        <Download size={18} />
                        {' '}{t("Export Visio (VSDX)")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        disabled={nodes.filter(n => n.type === 'azureNode').length === 0}
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportAsHtml();
                        }}
                        title={t("Export as interactive HTML with pan, zoom, and tooltips")}
                      >
                        <FileCode size={18} />
                        {' '}{t("Export Interactive HTML")}{' '}</button>
                      <div className="toolbar-dropdown-separator" role="separator" />
                      <div className="toolbar-dropdown-heading">
                        {localize(language, { en: 'Cost reports', ja: 'コストレポート' })}
                      </div>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        disabled={!hasCostReportData}
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportCostBreakdown();
                        }}
                        title={!hasCostReportData ? t("Add services to estimate costs first") : t("Export cost breakdown as CSV")}
                      >
                        <DollarSign size={18} />
                        {' '}{t("Export Costs (CSV)")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        disabled={!hasCostReportData}
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          exportCostBreakdownZip();
                        }}
                        title={!hasCostReportData ? t("Add services to estimate costs first") : t("Export CSV, JSON, summary and intelligent analysis as a ZIP")}
                      >
                        <DollarSign size={18} />
                        {' '}{t("Export Costs (All Formats)")}{' '}</button>

                      <div className="toolbar-dropdown-separator" role="separator" />

                      <div className="toolbar-dropdown-heading">{t("Recent exports")}</div>
                      {exportHistory.length === 0 ? (
                        <div className="toolbar-dropdown-hint toolbar-dropdown-hint--muted">{t("No exports yet")}</div>
                      ) : (
                        <div className="toolbar-dropdown-history">
                          {exportHistory.slice(0, 6).map((item) => (
                            <div key={item.id} className="toolbar-dropdown-history-item">
                              <div className="toolbar-dropdown-history-file">{item.fileName}</div>
                              <div className="toolbar-dropdown-history-meta">
                                {item.kind.toUpperCase()} {' '}{t("•")}{' '}{formatTimeAgo(item.createdAt)}
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>

              <div
                hidden={activeRibbonTab !== 'home'}
                className={`toolbar-group toolbar-group--labeled${collapsedToolbarSections.has('workspace') ? ' toolbar-group-collapsed' : ''}`}
                data-label={localize(language, { en: 'Workspace', ja: '表示・操作' })}
                aria-label={localize(language, { en: 'Workspace and help actions', ja: '表示・操作とヘルプ' })}
              >
                {toolbarSectionHeading('workspace', localize(language, { en: 'Workspace', ja: '表示・操作' }))}
                <button
                  className="btn btn-secondary"
                  onClick={() => {
                    setIsHelpOpen(true);
                    if (!helpSeen) {
                      setHelpSeen(true);
                      writeLocalStorage('help.seen', '1');
                    }
                  }}
                  title={t("How to use the tool & learn the features")}
                >
                  <HelpCircle size={18} />
                  {' '}{t("Help")}{' '}</button>
                <button 
                  onClick={() => setIsDarkMode(!isDarkMode)} 
                  className="btn btn-secondary" 
                  title={isDarkMode ? t("Switch to Light Mode") : t("Switch to Dark Mode")}
                  aria-label={isDarkMode ? t("Switch to Light Mode") : t("Switch to Dark Mode")}
                >
                  {isDarkMode ? <Sun size={18} /> : <Moon size={18} />}
                  {localize(language, {
                    en: isDarkMode ? 'Light' : 'Dark',
                    ja: isDarkMode ? 'ライト' : 'ダーク',
                  })}
                </button>
                <button
                  onClick={() => void startFreshDiagram()}
                  className="btn btn-secondary"
                  title={t("Clear diagram and start fresh")}
                >
                  <RefreshCw size={18} />
                  {localize(language, { en: 'New', ja: '新規' })}
                </button>
              </div>
            </div>

            {/* Row 2: History, arrangement, and review actions */}
            <div className="header-actions">
              <div
                hidden={activeRibbonTab !== 'review'}
                className={`toolbar-group toolbar-group--labeled${collapsedToolbarSections.has('history') ? ' toolbar-group-collapsed' : ''}`}
                data-label={localize(language, { en: 'History', ja: '履歴' })}
                aria-label={localize(language, { en: 'History and snapshots', ja: '履歴とスナップショット' })}
              >
                {toolbarSectionHeading('history', localize(language, { en: 'History', ja: '履歴' }))}
                <button
                  onClick={() => setIsCloudWorkspaceOpen(true)}
                  className={`btn btn-secondary${cloudSync.document ? ' btn-active' : ''}`}
                  title={cloudSync.errorMessage || localize(language, {
                    en: 'Open cloud autosave, sharing, comments, and snapshots',
                    ja: 'クラウド自動保存、共有、コメント、スナップショットを開く',
                  })}
                  aria-label={localize(language, {
                    en: `Cloud workspace: ${cloudSyncLabel}`,
                    ja: `クラウド ワークスペース: ${cloudSyncLabel}`,
                  })}
                >
                  {cloudSync.status === 'saving'
                    ? <Loader size={18} className="ribbon-cloud-saving" />
                    : <Cloud size={18} />}
                  {cloudSyncLabel}
                </button>
                <button 
                  onClick={() => setIsVersionHistoryModalOpen(true)} 
                  className="btn btn-secondary" 
                  title={t("View version history")}
                >
                  <Clock size={18} />
                  {' '}{t("History")}{' '}</button>
                <button 
                  onClick={() => setIsSaveSnapshotModalOpen(true)} 
                  className="btn btn-secondary" 
                  title={t("Save current diagram as snapshot")}
                  disabled={nodes.length === 0}
                >
                  <Camera size={18} />
                  {' '}{t("Snapshot")}{' '}</button>
              </div>

              <div
                hidden={activeRibbonTab !== 'design'}
                className={`toolbar-group toolbar-group--labeled${collapsedToolbarSections.has('arrange') ? ' toolbar-group-collapsed' : ''}`}
                data-label={localize(language, { en: 'Arrange', ja: '配置・選択' })}
                aria-label={localize(language, { en: 'Arrange, select, and style', ja: '配置、選択、スタイル' })}
              >
                {toolbarSectionHeading('arrange', localize(language, { en: 'Arrange', ja: '配置・選択' }))}
                <div className="toolbar-dropdown" ref={layoutMenuRef}>
                  <button
                    onClick={() => setIsLayoutMenuOpen((v) => !v)}
                    className="btn btn-secondary"
                    title={t("Layout presets")}
                    aria-haspopup="menu"
                    aria-expanded={isLayoutMenuOpen}
                  >
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M4 4h7v7H4z" />
                      <path d="M13 4h7v7h-7z" />
                      <path d="M4 13h7v7H4z" />
                      <path d="M13 13h7v7h-7z" />
                    </svg>
                    {' '}{t("Layout")}{' '}<ChevronDown size={16} style={{ marginLeft: 2 }} />
                  </button>

                  {isLayoutMenuOpen && (
                    <div className="toolbar-dropdown-menu toolbar-dropdown-menu--layout" role="menu" aria-label={t("Layout options")}>
                      <div className="toolbar-dropdown-heading">{t("Preset")}</div>
                      <select
                        className="toolbar-dropdown-select"
                        value={layoutPreset}
                        onChange={(e) => setLayoutPreset(e.target.value as LayoutPreset)}
                        aria-label={t("Layout preset")}
                      >
                        <option value="flow-lr">{t("Flow (L→R)")}</option>
                        <option value="flow-tb">{t("Flow (Top→Bottom)")}</option>
                        <option value="swimlanes">{t("Swimlanes by Group")}</option>
                        <option value="radial">{t("Radial")}</option>
                      </select>

                      <div className="toolbar-dropdown-separator" role="separator" />

                      <div className="toolbar-dropdown-row">
                        <label className="toolbar-dropdown-label" htmlFor="layoutEngine">
                          {' '}{t("Engine")}{' '}</label>
                        <select
                          id="layoutEngine"
                          className="toolbar-dropdown-select"
                          value={layoutEngine}
                          onChange={(e) => setLayoutEngine(e.target.value as LayoutEngineType)}
                        >
                          <option value="dagre">{t("Dagre")}</option>
                          <option value="elk">{t("ELK")}</option>
                        </select>
                      </div>

                      <div className="toolbar-dropdown-separator" role="separator" />

                      <div className="toolbar-dropdown-row">
                        <label className="toolbar-dropdown-label" htmlFor="layoutSpacing">
                          {' '}{t("Spacing")}{' '}</label>
                        <select
                          id="layoutSpacing"
                          className="toolbar-dropdown-select"
                          value={layoutSpacing}
                          onChange={(e) => setLayoutSpacing(e.target.value as LayoutSpacing)}
                        >
                          <option value="compact">{t("Compact")}</option>
                          <option value="comfortable">{t("Comfortable")}</option>
                        </select>
                      </div>

                      <div className="toolbar-dropdown-row">
                        <label className="toolbar-dropdown-label" htmlFor="edgeStyle">
                          {' '}{t("Edge style")}{' '}</label>
                        <select
                          id="edgeStyle"
                          className="toolbar-dropdown-select"
                          value={layoutEdgeStyle}
                          onChange={(e) => setLayoutEdgeStyle(e.target.value as LayoutEdgeStyle)}
                        >
                          <option value="straight">{t("Straight")}</option>
                          <option value="smooth">{t("Smooth")}</option>
                          <option value="orthogonal">{t("Orthogonal")}</option>
                        </select>
                      </div>

                      <label className="toolbar-dropdown-checkbox">
                        <input
                          type="checkbox"
                          checked={layoutEmphasizePrimaryPath}
                          onChange={(e) => setLayoutEmphasizePrimaryPath(e.target.checked)}
                          disabled={!(layoutPreset === 'flow-lr' || layoutPreset === 'flow-tb')}
                        />
                        {' '}{t("Emphasize primary path")}{' '}</label>

                      <div className="toolbar-dropdown-hint">
                        {' '}{t("Current:")}{' '}{translate(layoutPresetLabel[layoutPreset])} {' '}{t("• Engine:")}{' '}{layoutEngine === 'elk' ? 'ELK' : 'Dagre'}
                        {layoutPreset === 'radial' ? t(" (centers on selected node when possible)") : ''}
                      </div>

                      <div className="toolbar-dropdown-separator" role="separator" />

                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        disabled={nodes.length === 0}
                        onClick={() => {
                          setIsLayoutMenuOpen(false);
                          applyLayout();
                        }}
                        title={nodes.length === 0 ? t("Add services first") : t("Apply selected layout preset")}
                      >
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 12a9 9 0 1 1-9-9" />
                          <path d="M21 3v9h-9" />
                        </svg>
                        {' '}{t("Apply Layout")}{' '}</button>
                    </div>
                  )}
                </div>

                <button
                  type="button"
                  className={`btn btn-secondary${animateConnections ? ' btn-active' : ''}`}
                  onClick={() => setConnectionAnimations(!animateConnections)}
                  aria-pressed={animateConnections}
                  title={localize(language, {
                    en: animateConnections
                      ? 'Pause animated connection flow'
                      : 'Animate supported connection flow',
                    ja: animateConnections
                      ? '接続線のアニメーションを停止'
                      : '対応する接続線をアニメーション表示',
                  })}
                >
                  {animateConnections ? <Pause size={18} /> : <Play size={18} />}
                  {localize(language, { en: 'Flow motion', ja: '線アニメ' })}
                </button>

                <div className="toolbar-dropdown" ref={bulkSelectMenuRef}>
                  <button
                    onClick={() => setIsBulkSelectMenuOpen((v) => !v)}
                    className="btn btn-secondary"
                    title={t("Bulk select operations")}
                    aria-haspopup="menu"
                    aria-expanded={isBulkSelectMenuOpen}
                    disabled={nodes.length === 0}
                  >
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="3" y="3" width="7" height="7" rx="1" />
                      <rect x="14" y="3" width="7" height="7" rx="1" />
                      <rect x="3" y="14" width="7" height="7" rx="1" />
                      <rect x="14" y="14" width="7" height="7" rx="1" />
                    </svg>
                    {' '}{t("Select")}{' '}<ChevronDown size={16} style={{ marginLeft: 2 }} />
                  </button>

                  {isBulkSelectMenuOpen && (
                    <div className="toolbar-dropdown-menu" role="menu" aria-label={t("Bulk select options")}>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        onClick={selectAllNodes}
                      >
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M5 13l4 4L19 7" />
                        </svg>
                        {' '}{t("Select All Nodes")}{' '}</button>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        onClick={deselectAll}
                      >
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <circle cx="12" cy="12" r="10" />
                        </svg>
                        {' '}{t("Deselect All")}{' '}</button>
                      
                      {getServiceTypes().length > 0 && (
                        <>
                          <div className="toolbar-dropdown-separator" role="separator" />
                          <div className="toolbar-dropdown-heading">{t("Select by Service Type")}</div>
                          {getServiceTypes().map(serviceType => (
                            <button
                              key={serviceType}
                              className="toolbar-dropdown-item"
                              role="menuitem"
                              onClick={() => selectAllNodesOfType(serviceType)}
                            >
                              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                <rect x="3" y="3" width="18" height="18" rx="2" />
                              </svg>
                              {serviceType}
                            </button>
                          ))}
                        </>
                      )}
                    </div>
                  )}
                </div>

                <div className="toolbar-dropdown" ref={stylePresetMenuRef}>
                  <button
                    onClick={() => setIsStylePresetMenuOpen((v) => !v)}
                    className="btn btn-secondary"
                    title={t("Change diagram style")}
                    aria-haspopup="menu"
                    aria-expanded={isStylePresetMenuOpen}
                  >
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M12 2.69l5.66 5.66a8 8 0 1 1-11.31 0z" />
                    </svg>
                    {' '}{t("Style")}{' '}<ChevronDown size={16} style={{ marginLeft: 2 }} />
                  </button>

                  {isStylePresetMenuOpen && (
                    <div className="toolbar-dropdown-menu" role="menu" aria-label={t("Style preset options")}>
                      <div className="toolbar-dropdown-heading">{t("Visual Style")}</div>
                      <button
                        className={`toolbar-dropdown-item ${stylePreset === 'detailed' ? 'active' : ''}`}
                        role="menuitem"
                        onClick={() => applyStylePreset('detailed')}
                      >
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <rect x="3" y="3" width="18" height="18" rx="2" />
                          <path d="M3 9h18M3 15h18M9 3v18" />
                        </svg>
                        {' '}{t("Detailed (Default)")}{' '}</button>
                      <button
                        className={`toolbar-dropdown-item ${stylePreset === 'presentation' ? 'active' : ''}`}
                        role="menuitem"
                        onClick={() => applyStylePreset('presentation')}
                      >
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <rect x="2" y="3" width="20" height="14" rx="2" />
                          <path d="M8 21h8M12 17v4" />
                        </svg>
                        {' '}{t("Presentation (Professional)")}{' '}</button>
                      <div className="toolbar-dropdown-separator" role="separator" />
                      <div className="toolbar-dropdown-hint">
                        {stylePreset === 'detailed' && t("Shows all labels, pricing, and details")}
                        {stylePreset === 'presentation' && t("Professional look with bold connections")}
                      </div>
                    </div>
                  )}
                </div>

                <button
                  onClick={() => {
                    setFocusMode(prev => {
                      const next = !prev;
                      // Entering focus also collapses the side panels & legend
                      // (their existing one-way signal behavior).
                      if (next) setPanelsCollapsedSignal(p => p + 1);
                      return next;
                    });
                  }}
                  className={`btn btn-secondary${focusMode ? ' btn-active' : ''}`}
                  title={focusMode ? t("Show panels and diagram info") : t("Hide panels and diagram info for maximum diagram space")}
                  aria-pressed={focusMode}
                >
                  <PanelLeftClose size={18} />
                  {focusMode ? t("Exit Focus") : t("Focus")}
                </button>

                <button
                  onClick={toggleCollapseAllGroups}
                  className={`btn btn-secondary${allGroupsCollapsed ? ' btn-active' : ''}`}
                  title={allGroupsCollapsed ? t("Expand all groups to original size") : t("Collapse all groups to fit their content")}
                  disabled={nodes.filter(n => n.type === 'groupNode').length === 0}
                >
                  {allGroupsCollapsed ? <Maximize2 size={18} /> : <Minimize2 size={18} />}
                  {allGroupsCollapsed ? t("Expand Groups") : t("Collapse Groups")}
                </button>
              </div>

              <div
                hidden={activeRibbonTab !== 'review'}
                className={`toolbar-group toolbar-group--labeled${collapsedToolbarSections.has('review') ? ' toolbar-group-collapsed' : ''}`}
                data-label={localize(language, { en: 'Review', ja: 'レビュー・ガイド' })}
                aria-label={localize(language, { en: 'Architecture review and guides', ja: 'アーキテクチャのレビューとガイド' })}
              >
                {toolbarSectionHeading('review', localize(language, { en: 'Review', ja: 'レビュー・ガイド' }))}
                <button
                  onClick={handleValidateArchitecture}
                  className="btn btn-secondary"
                  title={t("Validate architecture against Azure Well-Architected Framework")}
                  disabled={nodes.length === 0}
                >
                  <Shield size={18} />
                  {' '}{t("Validate Architecture")}{' '}</button>
                <button
                  className="btn btn-secondary"
                  onClick={() => setIsCompareValidationOpen(true)}
                  title={t("Compare WAF validation results across multiple AI models")}
                  disabled={nodes.length === 0}
                >
                  <GitCompare size={18} />
                  {' '}{t("Compare Validation")}{' '}</button>
                {validationResult && (
                  <button
                    onClick={() => setIsValidationModalOpen(true)}
                    className="btn btn-secondary"
                    title={t("Open last validation results")}
                  >
                    <Shield size={18} />
                    {' '}{t("Validation:")}{' '}{translate(bandLabel(validationResult.overallScore))}
                  </button>
                )}
                <button
                  onClick={handleGenerateDeploymentGuide}
                  className="btn btn-secondary"
                  title={t("Generate comprehensive deployment guide")}
                  disabled={nodes.length === 0}
                >
                  <FileText size={18} />
                  {' '}{t("Deployment Guide")}{' '}</button>
                <button
                  onClick={() => setIsIaCRoundTripModalOpen(true)}
                  className={`btn btn-secondary${isIaCRoundTripModalOpen ? ' btn-active' : ''}`}
                  title={localize(language, {
                    en: 'Open deterministic IaC round-trip comparison, drift import, and starter export',
                    ja: '決定論的な IaC ラウンドトリップ比較、ドリフト取り込み、スターター出力を開く',
                  })}
                  disabled={!iacBaseline && nodes.filter(n => n.type === 'azureNode').length === 0}
                >
                  <GitCompare size={18} />
                  {localize(language, {
                    en: 'IaC Round-trip',
                    ja: 'IaC ラウンドトリップ',
                  })}
                </button>
                {deploymentGuide && (
                  <button
                    onClick={() => setIsDeploymentGuideModalOpen(true)}
                    className="btn btn-secondary"
                    title={t("Open last deployment guide")}
                  >
                    <FileText size={18} />
                    {' '}{t("View Guide")}{' '}</button>
                )}
              </div>
            </div>
            </div>
          </div>
          <div className="header-identity-actions">
            {accessIdentity?.enabled && accessIdentity.isAdmin && (
              <button
                type="button"
                className="access-admin-button"
                onClick={() => setIsAccessManagementOpen(true)}
                title={localize(language, {
                  en: `Manage application access (${accessIdentity.email})`,
                  ja: `アプリのアクセスを管理 (${accessIdentity.email})`,
                })}
              >
                <ShieldCheck size={17} />
                <span>{localize(language, { en: 'Access', ja: 'アクセス管理' })}</span>
              </button>
            )}
            <LanguageSwitch />
          </div>
          <button
            className="header-collapse-toggle"
            onClick={() => {
              setIsHeaderCollapsed(v => {
                const next = !v;
                writeLocalStorage(HEADER_COLLAPSED_STORAGE_KEY, next ? '1' : '0');
                return next;
              });
            }}
            title={isHeaderCollapsed ? t('header.showToolbar') : t('header.hideToolbarTitle')}
            aria-label={isHeaderCollapsed ? t('header.showToolbar') : t('header.hideToolbar')}
            aria-controls="application-toolbar"
            aria-expanded={!isHeaderCollapsed}
          >
            {isHeaderCollapsed ? <PanelTopOpen size={18} /> : <PanelTopClose size={18} />}
            <span>{isHeaderCollapsed ? t('header.showToolbar') : t('header.hideToolbar')}</span>
          </button>
        </div>
      </header>
      
      <div className="workspace">
        <IconPalette
          forceCollapsed={panelsCollapsedSignal > 0 ? panelsCollapsedSignal : undefined}
          onAddIcon={handleAddService}
        />
        
        <div className="canvas-container" ref={reactFlowWrapper}>
          <ReactFlow
            nodes={nodes}
            edges={edges}
            onNodesChange={onNodesChange}
            onEdgesChange={onEdgesChange}
            onNodesDelete={onNodesDelete}
            onConnect={onConnect}
            onReconnect={onReconnect}
            onNodeContextMenu={onNodeContextMenu}
            onEdgeContextMenu={onEdgeContextMenu}
            onPaneClick={dismissCanvasContextMenus}
            onInit={setReactFlowInstance}
            onDrop={onDrop}
            onDragOver={onDragOver}
            nodeTypes={nodeTypes}
            edgeTypes={edgeTypes}
            deleteKeyCode={null}
            fitView
            snapToGrid={true}
            snapGrid={[20, 20]}
            selectionOnDrag={true}
            panOnDrag={[1, 2]}
            elevateNodesOnSelect={false}
            reconnectRadius={20}
            attributionPosition="bottom-left"
          >
            <Controls />
            {nodes.length > 0 && (
              <>
                <div className="nav-minimap-caption">{t('canvas.miniMapCaption')}</div>
                <MiniMap
                  pannable
                  zoomable
                  position="bottom-right"
                  className="nav-minimap"
                  style={{ bottom: 84 }}
                  ariaLabel={t('canvas.miniMap')}
                  nodeColor="#60a5fa"
                  nodeStrokeColor="#3b82f6"
                  maskColor="rgba(30, 41, 59, 0.45)"
                />
              </>
            )}
            <Background 
              variant={BackgroundVariant.Dots} 
              gap={20} 
              size={2.5} 
              color="#60a5fa"
              style={{ backgroundColor: '#f8fafc' }}
            />
            {/* Canvas navigation hint — teaches pan/zoom so large diagrams
                aren't perceived as "stuck" or too big to view. Shown only when
                a diagram exists and until the user dismisses it. */}
            {showCanvasHint && nodes.length > 0 && (
              <div className="canvas-nav-hint" role="note" aria-label={t("Canvas navigation tips")}>
                <div className="canvas-nav-hint-tips">
                  <span className="canvas-nav-hint-tip canvas-nav-hint-desktop"><ZoomIn size={15} /> {' '}{t("Scroll to zoom in / out")}</span>
                  <span className="canvas-nav-hint-sep canvas-nav-hint-desktop" aria-hidden="true">{t("·")}</span>
                  <span className="canvas-nav-hint-tip canvas-nav-hint-desktop"><Hand size={15} /> {' '}{t("Right-click + drag to pan")}</span>
                  <span className="canvas-nav-hint-sep canvas-nav-hint-desktop" aria-hidden="true">{t("·")}</span>
                  <span className="canvas-nav-hint-tip canvas-nav-hint-mobile"><Hand size={15} /> {' '}{t('canvas.touchNavigation')}</span>
                  <button
                    type="button"
                    className="canvas-nav-hint-fit"
                    onClick={() => reactFlowInstance?.fitView?.({ padding: 0.2, duration: 400 })}
                    title={t("Zoom to fit the whole diagram in view")}
                  >
                    <Frame size={15} /> {' '}{t("Fit to view")}{' '}</button>
                </div>
                <button
                  type="button"
                  className="canvas-nav-hint-close"
                  onClick={() => {
                    setShowCanvasHint(false);
                    writeLocalStorage(CANVAS_HINT_STORAGE_KEY, '1');
                  }}
                  title={t("Dismiss (won't show again)")}
                  aria-label={t("Dismiss navigation tips")}
                >
                  <X size={14} />
                </button>
              </div>
            )}
            {/* Empty-canvas call-to-action — turns the blank grid into an
                obvious starting point that opens the Architecture Chat. Hidden
                once a diagram exists or while the chat panel is already open.
                pointer-events are disabled on the wrapper so drag-and-drop of
                services onto the canvas still works; only the button is
                clickable. */}
            {nodes.length === 0 && !isChatOpen && (
              <div className="canvas-empty-cta" role="note" aria-label={t("Get started")}>
                <div className="canvas-empty-cta-inner">
                  <MessagesSquare size={34} className="canvas-empty-cta-icon" />
                  <h2 className="canvas-empty-cta-title">{t("Start with a conversation")}</h2>
                  <p className="canvas-empty-cta-desc">
                    {' '}{t("Describe what you want to build in plain English — I’ll draw the first version, then you refine it step by step.")}{' '}</p>
                  <button
                    type="button"
                    className="canvas-empty-cta-btn"
                    onClick={() => setIsChatOpen(true)}
                  >
                    <MessagesSquare size={18} /> {' '}{t("Start with a conversation")}{' '}</button>
                  <span className="canvas-empty-cta-alt">
                    {' '}{t("or use")}{' '}<strong>{t("Generate with AI")}</strong> {' '}{t("· or add services from the left panel")}{' '}</span>
                </div>
              </div>
            )}
            <style>
              {highlightedServices.map(id => 
                `.react-flow__node[data-id="${id}"] {
                  filter: drop-shadow(0 0 12px rgba(96, 165, 250, 1)) drop-shadow(0 0 24px rgba(96, 165, 250, 0.9)) drop-shadow(0 0 36px rgba(96, 165, 250, 0.6)) !important;
                  z-index: 1000 !important;
                  animation: pulse-glow 1.5s ease-in-out infinite;
                }
                @keyframes pulse-glow {
                  0%, 100% { filter: drop-shadow(0 0 12px rgba(96, 165, 250, 1)) drop-shadow(0 0 24px rgba(96, 165, 250, 0.9)) drop-shadow(0 0 36px rgba(96, 165, 250, 0.6)); }
                  50% { filter: drop-shadow(0 0 18px rgba(96, 165, 250, 1)) drop-shadow(0 0 32px rgba(96, 165, 250, 1)) drop-shadow(0 0 48px rgba(96, 165, 250, 0.8)); }
                }
                
                body:not(.dark-mode) .react-flow__node[data-id="${id}"] {
                  filter: drop-shadow(0 0 8px rgba(0, 120, 212, 1)) drop-shadow(0 0 16px rgba(0, 120, 212, 0.8)) !important;
                }
                body:not(.dark-mode) @keyframes pulse-glow {
                  0%, 100% { filter: drop-shadow(0 0 8px rgba(0, 120, 212, 1)) drop-shadow(0 0 16px rgba(0, 120, 212, 0.8)); }
                  50% { filter: drop-shadow(0 0 12px rgba(0, 120, 212, 1)) drop-shadow(0 0 24px rgba(0, 120, 212, 0.9)); }
                }`
              ).join('\n')}
            </style>
            {/* Loading banner for applying recommendations */}
            {isApplyingRecommendations && (
              <div
                className="prompt-banner loading-banner"
                style={{
                  position: 'absolute',
                  left: '50%',
                  top: '10px',
                  transform: 'translateX(-50%)',
                  zIndex: 1001,
                  background: 'linear-gradient(135deg, #2563eb 0%, #1d4ed8 50%, #1e40af 100%)',
                  border: '2px solid #60a5fa',
                  padding: '1rem 2.5rem',
                  borderRadius: '12px',
                  boxShadow: '0 0 20px rgba(59, 130, 246, 0.5), 0 0 60px rgba(59, 130, 246, 0.2)',
                  maxWidth: '700px',
                  animation: 'pulse 2s ease-in-out infinite',
                }}
              >
                <div className="prompt-text" style={{ fontSize: '1.1rem', fontWeight: 600, letterSpacing: '0.3px' }}>
                  <strong style={{ fontSize: '1.2rem' }}>{t("⏳ Applying recommendations...")}</strong> {' '}{t("Regenerating architecture with improvements")}{' '}</div>
              </div>
            )}

            {/* Loading banner for template parsing */}
            {isImportingTemplate && (
              <div
                className="prompt-banner loading-banner"
                style={{
                  position: 'absolute',
                  left: '50%',
                  top: '10px',
                  transform: 'translateX(-50%)',
                  zIndex: 1001,
                  background: 'linear-gradient(135deg, #7c3aed 0%, #6d28d9 50%, #5b21b6 100%)',
                  border: '2px solid #a78bfa',
                  padding: '1rem 2.5rem',
                  borderRadius: '12px',
                  boxShadow: '0 0 20px rgba(139, 92, 246, 0.5), 0 0 60px rgba(139, 92, 246, 0.2)',
                  maxWidth: '700px',
                  animation: 'pulse 2s ease-in-out infinite',
                }}
              >
                <div className="prompt-text" style={{ fontSize: '1.1rem', fontWeight: 600, letterSpacing: '0.3px' }}>
                  <strong style={{ fontSize: '1.2rem' }}>{t("📄 Parsing")}{' '}{importFormatLabel} {' '}{t("Template...")}</strong> {' '}{t("Analyzing resources and generating architecture diagram")}{' '}</div>
              </div>
            )}

            {/* Architecture generation prompt banner */}
            {architecturePrompt && !focusMode && (
              <div
                className="prompt-banner draggable"
                style={{
                  position: 'absolute',
                  left: promptBannerPosition.x === 0 ? '50%' : `${promptBannerPosition.x}px`,
                  top: promptBannerPosition.y === 0 ? '10px' : `${promptBannerPosition.y}px`,
                  transform: promptBannerPosition.x === 0 ? 'translateX(-50%)' : 'none',
                  cursor: isDraggingBanner ? 'grabbing' : 'grab',
                  zIndex: 1000,
                }}
                onMouseDown={(e) => {
                  e.preventDefault();
                  const rect = e.currentTarget.getBoundingClientRect();
                  // Calculate offset from mouse position to element's top-left corner
                  setDragOffset({
                    x: e.clientX - rect.left,
                    y: e.clientY - rect.top,
                  });
                  // Store the current absolute position
                  if (promptBannerPosition.x === 0) {
                    // First time dragging - calculate initial center position
                    const initialX = window.innerWidth / 2 - rect.width / 2;
                    setPromptBannerPosition({ x: initialX, y: 10 });
                  }
                  setIsDraggingBanner(true);
                }}
              >
                <div className="prompt-text">
                  <strong>{t("Generated from:")}</strong> {architecturePrompt}
                </div>
              </div>
            )}

            {nodes.length > 0 && (
              <TitleBlock
                architectureName={titleBlockData.architectureName}
                author={titleBlockData.author}
                version={titleBlockData.version}
                date={titleBlockData.date}
                onUpdate={(data) => setTitleBlockData({ ...titleBlockData, ...data })}
              />
            )}
            {generatedWithModel && !focusMode && (
              <ModelBadge
                modelName={generatedWithModel.name}
                elapsedTimeMs={generatedWithModel.timeMs}
              />
            )}
            <Legend forceCollapsed={panelsCollapsedSignal > 0 ? panelsCollapsedSignal : undefined} />
            {referenceImageUrl && (
              <ReferenceImageViewer
                imageUrl={referenceImageUrl}
                onDismiss={() => setReferenceImageUrl(null)}
              />
            )}
          </ReactFlow>
          <AlignmentToolbar 
            selectedNodes={nodes.filter(n => n.selected)}
            onAlign={handleAlign}
          />
        </div>
        {workflow.length > 0 && (
          <WorkflowPanel 
            workflow={workflow}
            onServiceHover={(refs) => {
              // Workflow steps reference services by node id (app-generated) OR
              // by label (e.g. scenes exported by the MCP server). The glow CSS
              // targets nodes by id, so resolve each ref to matching node ids by
              // id or case-insensitive label before highlighting.
              const ids = (refs || []).flatMap((ref) => {
                const refLc = String(ref).toLowerCase();
                return nodes
                  .filter((n) => n.type === 'azureNode'
                    && (n.id === ref || String(n.data?.label ?? '').toLowerCase() === refLc))
                  .map((n) => n.id);
              });
              setHighlightedServices(ids.length > 0 ? ids : (refs || []));
            }}
            onServiceLeave={() => setHighlightedServices([])}
            forceCollapsed={panelsCollapsedSignal > 0 ? panelsCollapsedSignal : undefined}
          />
        )}
        
        {/* Service / Layer Context Menu */}
        {nodeContextMenu && contextMenuNode && (
          <>
            <div
              className="edge-context-menu-overlay"
              onClick={dismissNodeContextMenu}
              onContextMenu={(event) => {
                event.preventDefault();
                dismissNodeContextMenu();
              }}
            />
            <div
              ref={nodeContextMenuRef}
              className="node-context-menu"
              role="menu"
              onKeyDown={handleNodeContextMenuKeyDown}
              aria-label={localize(language, {
                en: `${contextMenuNodeIsGroup ? 'Layer' : 'Service'} actions`,
                ja: `${contextMenuNodeIsGroup ? 'レイヤー' : 'サービス'}の操作`,
              })}
              style={{
                position: 'fixed',
                top: nodeContextMenu.y,
                left: nodeContextMenu.x,
                zIndex: 10000,
              }}
            >
              <div className="context-menu-header node-context-menu-header">
                <span>
                  {localize(language, {
                    en: contextMenuNodeIsGroup ? 'Layer' : 'Service',
                    ja: contextMenuNodeIsGroup ? 'レイヤー' : 'サービス',
                  })}
                </span>
                <strong title={contextMenuNodeLabel}>{contextMenuNodeLabel}</strong>
              </div>

              {contextMenuNodeIsGroup ? (
                <>
                  <button
                    className="context-menu-item"
                    role="menuitem"
                    disabled={contextMenuContainedCount === 0}
                    onClick={() => fitContextGroupToContent(contextMenuNode.id)}
                  >
                    <span className="menu-icon"><Minimize2 size={16} /></span>
                    <span className="context-menu-item-copy">
                      <span>{localize(language, { en: 'Fit layer to contents', ja: '内容に合わせてレイヤーを調整' })}</span>
                      <small>{localize(language, { en: 'Resize around contained items', ja: '内部の項目に合わせてサイズを変更' })}</small>
                    </span>
                  </button>
                  <div className="context-menu-separator" role="separator" />
                  <button
                    className="context-menu-item danger"
                    role="menuitem"
                    onClick={() => deleteContextNode(contextMenuNode.id)}
                  >
                    <span className="menu-icon"><Ungroup size={16} /></span>
                    <span className="context-menu-item-copy">
                      <span>
                        {localize(language, {
                          en: contextMenuContainedCount > 0 ? 'Delete layer only' : 'Delete layer',
                          ja: contextMenuContainedCount > 0 ? 'レイヤーだけを削除' : 'レイヤーを削除',
                        })}
                      </span>
                      {contextMenuContainedCount > 0 && (
                        <small>{localize(language, { en: 'Keep all contained items', ja: '内部の項目はすべて保持' })}</small>
                      )}
                    </span>
                  </button>
                  {contextMenuContainedCount > 0 && (
                    <button
                      className="context-menu-item danger"
                      role="menuitem"
                      onClick={() => deleteContextGroupWithContents(contextMenuNode.id)}
                    >
                      <span className="menu-icon"><Boxes size={16} /></span>
                      <span className="context-menu-item-copy">
                        <span>
                          {localize(language, {
                            en: `Delete layer and ${contextMenuContainedCount} item${contextMenuContainedCount === 1 ? '' : 's'}`,
                            ja: `レイヤーと内部の${contextMenuContainedCount}件を削除`,
                          })}
                        </span>
                        <small>{localize(language, { en: 'Requires confirmation', ja: '確認後に削除' })}</small>
                      </span>
                    </button>
                  )}
                </>
              ) : (
                <>
                  <button
                    className="context-menu-item"
                    role="menuitem"
                    onClick={() => duplicateServiceNode(contextMenuNode.id)}
                  >
                    <span className="menu-icon"><Copy size={16} /></span>
                    <span>{localize(language, { en: 'Duplicate service', ja: 'サービスを複製' })}</span>
                  </button>
                  <button
                    className="context-menu-item"
                    role="menuitem"
                    onClick={() => editContextNodePricing(contextMenuNode.id)}
                  >
                    <span className="menu-icon"><DollarSign size={16} /></span>
                    <span>
                      {contextMenuNode.data?.pricing
                        ? localize(language, { en: 'Edit cost estimate', ja: 'コスト見積を編集' })
                        : localize(language, { en: 'Set cost estimate', ja: 'コスト見積を設定' })}
                    </span>
                  </button>
                  <div className="context-menu-separator" role="separator" />
                  <button
                    className="context-menu-item danger"
                    role="menuitem"
                    onClick={() => deleteContextNode(contextMenuNode.id)}
                  >
                    <span className="menu-icon"><Trash2 size={16} /></span>
                    <span>{localize(language, { en: 'Delete service', ja: 'サービスを削除' })}</span>
                  </button>
                </>
              )}
            </div>
          </>
        )}

        {/* Edge Context Menu */}
        {edgeContextMenu && (
          <>
            <div
              className="edge-context-menu-overlay"
              onClick={closeEdgeContextMenu}
              onContextMenu={(event) => {
                event.preventDefault();
                closeEdgeContextMenu();
              }}
            />
            <div 
              className="edge-context-menu"
              style={{
                position: 'fixed',
                top: edgeContextMenu.y,
                left: edgeContextMenu.x,
                zIndex: 10000,
              }}
            >
              <div className="context-menu-header">{t("Edge Direction")}</div>
              <button
                className="context-menu-item"
                onClick={() => setEdgeDirection(edgeContextMenu.edgeId, 'forward')}
              >
                <span className="menu-icon">{t("→")}</span>
                <span>{t("One-way (Forward)")}</span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => setEdgeDirection(edgeContextMenu.edgeId, 'reverse')}
              >
                <span className="menu-icon">{t("←")}</span>
                <span>{t("One-way (Reverse)")}</span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => setEdgeDirection(edgeContextMenu.edgeId, 'bidirectional')}
              >
                <span className="menu-icon">{t("↔")}</span>
                <span>{t("Bidirectional")}</span>
              </button>
              <div className="context-menu-separator" role="separator" />
              <div className="context-menu-header">
                {localize(language, { en: 'Flow animation', ja: 'フローアニメーション' })}
              </div>
              <button
                className="context-menu-item"
                onClick={() => toggleEdgeAnimation(edgeContextMenu.edgeId)}
                aria-pressed={Boolean(contextMenuEdge?.data?.flowAnimated)}
              >
                <span className="menu-icon">
                  {contextMenuEdge?.data?.flowAnimated ? <Pause size={16} /> : <Play size={16} />}
                </span>
                <span>
                  {localize(language, {
                    en: contextMenuEdge?.data?.flowAnimated ? 'Pause this connection' : 'Animate this connection',
                    ja: contextMenuEdge?.data?.flowAnimated ? 'この接続線を停止' : 'この接続線をアニメーション',
                  })}
                </span>
              </button>
            </div>
          </>
        )}
      </div>

      {/* Premium Feature Modals */}
      <ValidationModal
        validation={validationResult}
        isOpen={isValidationModalOpen}
        onClose={() => setIsValidationModalOpen(false)}
        isLoading={isValidating}
        onRevalidate={handleValidateArchitecture}
        onApplyRecommendations={async (selectedFindings) => {
          console.log('📝 User selected recommendations to apply:', selectedFindings);
          
          // Close validation modal and show loading state
          setIsValidationModalOpen(false);
          setIsApplyingRecommendations(true);
          
          // Get current architecture state
          const currentServices = nodes
            .filter(n => n.type === 'azureNode')
            .map(n => ({
              id: n.id,
              name: n.data.label,
              type: n.data.serviceName || n.data.service || n.data.label,
              category: n.data.category || 'General',
              description: n.data.description || '',
            }));
          
          const currentConnections = edges.map(e => ({
            from: e.source,
            to: e.target,
            label: e.label || '',
            type: e.data?.type || 'sync'
          }));
          
          const currentGroups = nodes
            .filter(n => n.type === 'groupNode')
            .map(n => ({
              id: n.id,
              label: n.data.label,
            }));
          
          // Format recommendations for the prompt
          const recommendationsText = selectedFindings
            .map((f, i) => [
              `${i + 1}. [${f.severity.toUpperCase()}] ${f.category}`,
              `   Issue: ${f.issue}`,
              `   Recommendation: ${f.recommendation}`,
              ...(f.evidence || []).map((item) => `   Evidence: ${item}`),
              ...(f.remediation || []).map((step, stepIndex) => `   Step ${stepIndex + 1}: ${step}`),
              f.applyAction
                ? `   Intended diagram action: ${f.applyAction.label}${f.applyAction.serviceType ? ` (${f.applyAction.serviceType})` : ''}`
                : '',
            ].filter(Boolean).join('\n'))
            .join('\n');
          
          // Build regeneration prompt
          const regenerationPrompt = `You are regenerating an existing Azure architecture with improvements based on Well-Architected Framework recommendations.

CURRENT ARCHITECTURE:
Services: ${currentServices.map(s => `${s.name} (${s.type})`).join(', ')}
Connections: ${currentConnections.length} connections
Groups: ${currentGroups.map(g => g.label).join(', ')}

SELECTED RECOMMENDATIONS TO APPLY:
${recommendationsText}

CRITICAL INSTRUCTIONS:
1. Keep all existing services that are working well and not affected by recommendations
2. Add new services recommended (e.g., Azure DevOps, Azure Monitor, Application Insights, Azure Front Door, Redis Cache, etc.)
3. **IMPORTANT**: Place ALL new services into appropriate logical groups:
   - DevOps/CI/CD services → "DevOps & Deployment" or "CI/CD Pipeline" group
   - Monitoring services → Add to existing monitoring group or create "Monitoring & Observability" group
   - Security services → Add to existing security group or create "Security & Compliance" group
   - Caching services → Add to "Data & Cache" or similar group
   - Never leave services ungrouped unless they are truly standalone
4. Update connections to reflect security improvements, monitoring, and best practices
5. Maintain and extend the logical grouping structure - create new groups as needed for new service categories
6. Ensure the architecture implements the selected recommendations

MULTI-REGION RESILIENCY — CRITICAL:
When a recommendation involves multi-region, failover, geo-redundancy, or disaster recovery:
- **ADD DUPLICATE SERVICES** in a secondary region. For example, if the primary has "API Management", add a second node "API Management (Secondary)" in a separate region group.
- **CREATE REGION GROUPS**: Replace a single "Application & AI" group with "Primary Region (East US)" and "Secondary Region (West US)" groups, each containing their own instances of the affected services.
- **SHOW FAILOVER ROUTING**: Add Azure Front Door or Traffic Manager as the entry point that routes to both regional deployments via origin groups or priority routing.
- **DATABASE REPLICATION**: For Cosmos DB, add a second node "Azure Cosmos DB (Replica)" in the secondary region group with a replication connection between the two.
- **DO NOT just change edge labels** to mention "primary" or "failover" — that is NOT implementing multi-region. You must ADD actual service nodes in a secondary region.
- Target 14-18 services total for multi-region architectures (roughly double the region-scoped services).

SERVICE MAPPING — CRITICAL:
- When adding private endpoints or Private Link, add "Azure Private Link" as an explicit service node AND "Azure DNS" for private DNS resolution. Connect source services → Azure Private Link → target PaaS services.
- When adding WAF capabilities, add "Web Application Firewall" as a service node if not already covered by Application Gateway or Azure Front Door.
- When adding SIEM/security monitoring, add "Microsoft Sentinel" as a service node.
- Always use exact service names from the known services list in the system prompt.

LAYOUT RULES:
7. Limit total connections to 12-18. Only show primary data/control flow, not obvious implicit relationships. Show only 1 representative Key Vault edge, not one per service.
8. For monitoring: connect ONLY the primary compute service to Azure Monitor, then a SINGLE edge to Log Analytics. Maximum 2-3 monitoring edges total.
9. Arrange groups in directional flow: Ingress → Application → Data (left-to-right). Security at bottom-left, Monitoring at bottom-right.
10. Minimize cross-group edges. Place tightly-coupled services in the SAME group. Aim for 1-2 outgoing edges per group.
11. Total service count: 8-18 depending on complexity. Multi-region architectures will have more services. Only add security/identity services when the recommendations explicitly require them.

Return the IMPROVED architecture in the same JSON format as before with proper group assignments.`;

          console.log('🔄 Regenerating architecture with recommendations...');
          console.log('📋 Prompt:', regenerationPrompt);
          
          // Call Azure OpenAI to regenerate
          try {
            const improvedArchitecture = await generateArchitectureWithAI(
              regenerationPrompt,
              undefined,
              undefined,
              language,
            );
            
            if (improvedArchitecture) {
              // Detect newly added services
              const existingServiceNames = new Set(currentServices.map(s => s.name.toLowerCase()));
              const newServices = improvedArchitecture.services
                .filter((s: any) => !existingServiceNames.has(s.name.toLowerCase()))
                .map((s: any) => s.name);
              
              // Build descriptive banner text
              let bannerText = localize(language, {
                en: `Original architecture improved with ${selectedFindings.length} WAF recommendation${selectedFindings.length > 1 ? 's' : ''}`,
                ja: `元のアーキテクチャへ${selectedFindings.length}件のWAF推奨事項を適用`,
              });
              if (newServices.length > 0) {
                bannerText += localize(language, {
                  en: `. Added: ${newServices.join(', ')}`,
                  ja: `。追加: ${newServices.join('、')}`,
                });
              }
              
              // Apply the improved architecture
              await handleAIGenerate(improvedArchitecture, bannerText, true, true);
              trackRecommendationsApplied(selectedFindings.length);
              
              setIsApplyingRecommendations(false);
              alert(localize(language, {
                en: `✅ Architecture regenerated successfully!\n\nApplied ${selectedFindings.length} recommendations.\n${newServices.length > 0 ? `\nAdded ${newServices.length} new services: ${newServices.join(', ')}` : ''}`,
                ja: `✅ アーキテクチャを再生成しました。\n\n${selectedFindings.length}件の推奨事項を適用しました。\n${newServices.length > 0 ? `\n${newServices.length}件の新しいサービスを追加: ${newServices.join('、')}` : ''}`,
              }));
            }
          } catch (error) {
            console.error('❌ Failed to regenerate architecture:', error);
            setIsApplyingRecommendations(false);
            alert(t("Failed to regenerate architecture. Please try again."));
          }
        }}
      />
      <DeploymentGuideModal
        guide={deploymentGuide}
        isOpen={isDeploymentGuideModalOpen}
        onClose={() => setIsDeploymentGuideModalOpen(false)}
        isLoading={isGeneratingGuide}
      />
      <IaCRoundTripModal
        isOpen={isIaCRoundTripModalOpen}
        onClose={() => setIsIaCRoundTripModalOpen(false)}
        baseline={iacBaseline}
        comparison={iacComparison}
        driftPlan={driftPlanSummary}
        onImportDriftPlan={importDriftPlan}
        onClearDriftPlan={() => setDriftPlanSummary(null)}
        onDownloadStarter={downloadStarterTemplate}
        bicepStarter={bicepStarterTemplate}
        terraformStarter={terraformStarterTemplate}
        diagramServiceCount={nodes.filter(n => n.type === 'azureNode').length}
      />
      <VersionHistoryModal
        isOpen={isVersionHistoryModalOpen}
        onClose={() => setIsVersionHistoryModalOpen(false)}
        onRestoreVersion={restoreVersion}
        currentDiagramName={titleBlockData.architectureName}
      />
      <CloudWorkspaceModal
        isOpen={isCloudWorkspaceOpen}
        onClose={() => setIsCloudWorkspaceOpen(false)}
        currentDocument={cloudSync.document}
        currentContext={cloudSync.context}
        syncStatus={cloudSync.status}
        syncError={cloudSync.errorMessage}
        lastSavedAt={cloudSync.lastSavedAt}
        onOpenDocument={cloudSync.openDocument}
        onRestoreVersion={cloudSync.restoreVersion}
        onDocumentUpdated={cloudSync.replaceCurrentDocument}
        onResetCurrent={cloudSync.reset}
        onReloadRemote={cloudSync.reloadRemote}
        onSaveAsCopy={cloudSync.saveAsCopy}
        onCreateNew={startFreshDiagram}
      />
      <PricingScenarioModal
        isOpen={isPricingScenarioModalOpen}
        onClose={() => setIsPricingScenarioModalOpen(false)}
        nodes={nodes}
        scenarios={pricingScenarios}
        onChange={setPricingScenarios}
      />
      <SaveSnapshotModal
        isOpen={isSaveSnapshotModalOpen}
        onClose={() => setIsSaveSnapshotModalOpen(false)}
        onSave={handleSaveSnapshot}
        diagramName={titleBlockData.architectureName}
        serviceCount={nodes.filter(n => n.type === 'azureNode').length}
      />
      <AzureImportModal
        isOpen={isAzureImportOpen}
        onClose={() => setIsAzureImportOpen(false)}
        onImport={importFromAzure}
      />
      <CompareModelsModal
        isOpen={isCompareModelsOpen}
        onClose={() => setIsCompareModelsOpen(false)}
        onApply={async (architecture, prompt, sourceModel, sourceReasoningEffort) => {
          trackModelComparison({ selectedModel: sourceModel });
          if (sourceModel && sourceReasoningEffort) {
            setSourceModel(sourceModel, sourceReasoningEffort);
          }
          await handleAIGenerate(architecture, prompt, true, false);
        }}
        onCaptureBatch={async (items) => {
          // Render each architecture on the main canvas in turn, capture as PNG,
          // and trigger a download. Filenames are supplied by the modal so the
          // PNG file always pairs 1:1 with the JSON saved via "Save All Diagrams".
          for (let i = 0; i < items.length; i++) {
            const item = items[i];
            try {
              // Apply this architecture to the canvas (no auto-snapshot to
              // avoid spamming the snapshot history with N intermediate states).
              await handleAIGenerate(item.architecture, item.prompt, false, false);
              // Give icons, layout, and the post-generate fitView a moment to settle.
              await new Promise(res => setTimeout(res, 1500));
              if (reactFlowInstance) {
                reactFlowInstance.fitView({ padding: 0.2, duration: 0 });
                await new Promise(res => setTimeout(res, 400));
              }
              if (!reactFlowWrapper.current) continue;
              const dataUrl = await captureDiagramAsPng(reactFlowWrapper.current, {
                backgroundColor: '#ffffff',
              });
              const a = document.createElement('a');
              a.href = dataUrl;
              a.download = item.filename;
              a.click();
              // Small gap so the browser doesn't throttle / merge downloads.
              await new Promise(res => setTimeout(res, 350));
            } catch (err) {
              console.error(`Failed to capture PNG for ${item.filename}:`, err);
            }
          }
        }}
      />
      <CompareValidationModal
        isOpen={isCompareValidationOpen}
        onClose={() => setIsCompareValidationOpen(false)}
        onApply={(validation) => {
          setValidationResult(validation);
          setIsValidationModalOpen(true);
          setPanelsCollapsedSignal(prev => prev + 1);
        }}
        services={nodes
          .filter(n => n.type === 'azureNode')
          .map(n => ({
            name: n.data.label || n.data.serviceName || 'Unknown Service',
            type: n.data.serviceName || n.data.label || 'Unknown',
            category: n.data.category || 'General',
          }))}
        connections={edges.map(e => ({
          from: nodes.find(n => n.id === e.source)?.data?.label || e.source,
          to: nodes.find(n => n.id === e.target)?.data?.label || e.target,
          label: String(e.label || ''),
        }))}
        groups={nodes
          .filter(n => n.type === 'groupNode')
          .map(n => ({
            name: n.data.label || 'Group',
            services: nodes
              .filter(child => child.parentNode === n.id)
              .map(child => child.data.label || child.data.serviceName || 'Unknown'),
          }))}
        architectureDescription={architecturePrompt || titleBlockData.architectureName}
      />

      <ValidationHandoffToast
        isOpen={validationHandoff !== null && !focusMode}
        isModification={validationHandoff?.source === 'modification'}
        isChatOpen={isChatOpen}
        onValidate={handleValidationHandoffStart}
        onDismiss={handleValidationHandoffDismiss}
      />

      {!isChatOpen && (
        <button
          className={`feedback-fab${feedbackFabPulse ? ' pulse-once' : ''}`}
          onClick={() => setIsFeedbackModalOpen(true)}
          title={t("Share feedback")}
        >
          <MessageSquare size={18} />
          {' '}{t("Feedback")}{' '}
        </button>
      )}
      <ArchitectureChatPanel
        isOpen={isChatOpen}
        onClose={() => setIsChatOpen(false)}
        currentArchitecture={{
          nodes,
          edges,
          architectureName: titleBlockData.architectureName,
        }}
        onApply={(architecture, prompt, autoSnapshot) => (
          handleAIGenerate(architecture, prompt, autoSnapshot, nodes.length > 0)
        )}
      />
      <HelpLearnPanel
        isOpen={isHelpOpen}
        onClose={() => setIsHelpOpen(false)}
      />
      <FeedbackToast
        isOpen={isFeedbackToastOpen}
        onClose={() => setIsFeedbackToastOpen(false)}
        onAddComment={(rating) => {
          setIsFeedbackToastOpen(false);
          setFeedbackPreselectedRating(rating);
          setIsFeedbackModalOpen(true);
        }}
        context={{
          diagramName: titleBlockData.architectureName,
          serviceCount: nodes.filter(n => n.type === 'azureNode').length,
          model: generatedWithModel?.name,
        }}
      />
      <FeedbackModal
        isOpen={isFeedbackModalOpen}
        onClose={() => {
          setIsFeedbackModalOpen(false);
          setFeedbackPreselectedRating(undefined);
        }}
        preselectedRating={feedbackPreselectedRating}
        context={{
          diagramName: titleBlockData.architectureName,
          serviceCount: nodes.filter(n => n.type === 'azureNode').length,
          model: generatedWithModel?.name,
        }}
      />
      {accessIdentity?.enabled && accessIdentity.isAdmin && (
        <AccessManagementModal
          isOpen={isAccessManagementOpen}
          identity={accessIdentity}
          onClose={() => setIsAccessManagementOpen(false)}
        />
      )}
      {(() => {
        if (!pricingEditorNodeId) return null;
        const node = nodes.find(n => n.id === pricingEditorNodeId);
        const storedPricing = node?.data?.pricing as NodePricingConfig | undefined;
        const nodePricing = pricingEditorDraft?.nodeId === pricingEditorNodeId
          ? pricingEditorDraft.pricing
          : storedPricing;
        if (!node || !nodePricing) return null;
        return (
          <NodePricingEditor
            serviceType={String(node.data.serviceName || node.data.label || 'Unknown')}
            pricing={nodePricing}
            onClose={closePricingEditor}
            onApply={(updated) => {
              // Total cost recalculates from `nodes` via the existing effect.
              setNodes(nds =>
                nds.map(n =>
                  n.id === pricingEditorNodeId
                    ? { ...n, data: { ...n.data, pricing: updated } }
                    : n,
                ),
              );
            }}
          />
        );
      })()}
    </div>
  );
}

export default App;
