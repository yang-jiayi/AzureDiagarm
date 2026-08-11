// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { lazy, Suspense, useState, useCallback, useRef, useEffect, useLayoutEffect, useMemo } from 'react';
import ReactFlow, {
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
  getNodesBounds,
  type NodeChange,
  type ReactFlowInstance,
} from 'reactflow';
import 'reactflow/dist/style.css';
import type { CaptureOptions } from './utils/captureCanvas';
import { type ExportBackground } from './utils/captureCanvas';
import { animateEdgeFlow } from './utils/animateEdges';
import { sequenceWorkflowSvg } from './utils/sequenceWorkflow';
import { buildWorkflowMarkdown } from './services/workflowNarrativeExporter';
import { Download, Save, Upload, DollarSign, Shield, ShieldCheck, FileText, FileCode, ChevronDown, ChevronRight, Clock, Camera, Loader, GitCompare, RefreshCw, PanelLeftClose, Minimize2, Maximize2, Presentation, MessagesSquare, HelpCircle, Info, Frame, PanelTopClose, PanelTopOpen, DownloadCloud, Sun, Moon, Play, Pause, Eye, EyeOff, Cloud, Copy, Trash2, Ungroup, Boxes, CheckSquare, Sparkles, Undo2, Redo2, LayoutTemplate, ScanSearch } from 'lucide-react';
import AboutDialog from './components/AboutDialog';
import IconPalette from './components/IconPalette';
import AzureNode from './components/AzureNode';
import GroupNode from './components/GroupNode';
import AIArchitectureGenerator from './components/AIArchitectureGenerator';
import ArchitectureChatPanel from './components/ArchitectureChatPanel';
import BYOAISettingsDialog from './components/BYOAISettingsDialog';
import CanvasActivityOverlay from './components/CanvasActivityOverlay';
import CanvasChrome from './components/CanvasChrome';
import CommandPalette, { type CommandPaletteAction } from './components/CommandPalette';
import { DeliverChooser } from './components/GuidedJourney';
import HelpLearnPanel from './components/GuidedHelpPanel';
import MobileCommandBar from './components/MobileCommandBar';
import MobileNodeInspector from './components/MobileNodeInspector';
import PrivacyPreflightDialog from './components/PrivacyPreflightDialog';
import ResponsiveRibbonSurface from './components/ResponsiveRibbonSurface';
import DocumentStatus from './components/DocumentStatus';
import TemplateGallery from './components/TemplateGallery';
import ThreatModelOverlay from './components/ThreatModelOverlay';
import WorkflowStepper from './components/WorkflowStepper';
import type { ReferenceArchitecture } from './services/referenceArchitectureAI';
import type { BlueprintArchitecture } from './services/blueprintArchitectureAI';
import ReferenceImageViewer from './components/ReferenceImageViewer';
import EditableEdge from './components/EditableEdge';
import AlignmentToolbar, {
  type BulkEditRequest,
  type BulkEditResult,
} from './components/AlignmentToolbar';
import WorkflowPanel from './components/WorkflowPanel';
import RegionSelector from './components/RegionSelector';
import ModelSettingsPopover from './components/ModelSettingsPopover';
import { loadIconsFromCategory, type AzureIcon } from './utils/iconLoader';
import { resolveServiceIconLoose } from './utils/serviceIconFuzzy';
import { getServiceIconMapping, isCapacityConsumed } from './data/serviceIconMapping';
import { initializeNodePricing, updateNodePricing, setCustomPricing, calculateCostBreakdown, exportCostBreakdownCSV, exportCostBreakdownJSON, getCostSummaryMarkdown, refreshAllNodePricing, type PricingMode } from './services/costEstimationService';
import { prefetchCommonServices } from './services/azurePricingService';
import { preloadCommonServices, getActiveRegion, AzureRegion, AVAILABLE_REGIONS, RegionInfo } from './services/regionalPricingService';
import { formatMonthlyCost, getPricingFreshness } from './utils/pricingHelpers';
import { hasPricingData, PRICING_DATA_AS_OF } from './data/azurePricing';
import { costReportToHtml } from './utils/costReportHtml';
import { validateArchitecture, ArchitectureValidation } from './services/architectureValidator';
import { bandLabel } from './services/wafMaturity';
import type { DeploymentGuide } from './services/deploymentGuideGenerator';
import { generateArchitectureWithAI } from './services/azureOpenAI';
import { MODEL_CONFIG, getDeploymentNames, type ModelType } from './stores/modelSettingsStore';
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
  DEFAULT_PRICING_SCENARIOS,
  loadPricingScenarios,
  normalizePricingScenarios,
  savePricingScenarios,
} from './services/pricingScenarioService';
import { createSnapshot, DiagramVersion, getVersion } from './services/versionStorageService';
import {
  CloudDiagramOperationCancelledError,
  useCloudDiagramSync,
} from './hooks/useCloudDiagramSync';
import {
  type CloudCommentAnchor,
  getCloudDiagram,
  type CloudDiagramSummary,
  CloudDiagramDocument,
  CloudDiagramVersion,
} from './services/cloudDiagramService';
import {
  getRecentWorkSessionId,
  saveRecentWork,
  type RecentWorkRecord,
  type RecentWorkSyncState,
} from './services/recentWorkService';
import { useDiagramHistory } from './hooks/useDiagramHistory';
import type { DeckService } from './services/pptxExporter';
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
  detachNodeFromGroup,
  duplicateSelectedSubgraph,
  fitGroupToContent,
  fitAllGroupsToContent,
  captureGroupLayout,
  restoreGroupLayout,
  type GroupLayoutSnapshot,
} from './utils/groupUtils';
import {
  buildAbsolutePositionMap,
  preserveManualLayout,
  selectHorizontalConnectionHandles,
} from './utils/preserveManualLayout';
import { mergeLayoutEdges, mergeLayoutNodes } from './utils/layoutResultMerge';
import { OperationGeneration } from './utils/operationGeneration';
import {
  getCurrentValidationScore,
  resolveValidationFreshness,
} from './utils/validationFreshness';
import { trackArchitectureGeneration, trackValidation, trackValidationHandoff, trackDeploymentGuide, trackExport, trackTemplateImport, trackModelComparison, trackRecommendationsApplied, trackVersionOperation, trackStartFresh, trackValidationFindings, trackGuidedJourney } from './services/telemetryService';
import { classifyValidationTopics } from './services/validationConsensus';
import type { IaCFormat } from './services/azureOpenAI';
import FeedbackToast from './components/FeedbackToast';
import HeaderUtilityMenu from './components/HeaderUtilityMenu';
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
import type { ArchitectureTemplate } from './data/architectureTemplates';
import {
  anonymizeDiagramPayload,
  detectSensitiveData,
  detectSensitiveDataInValue,
  type SensitiveFinding,
} from './utils/privacyPreflight';
import { analyzeThreatModel } from './utils/threatModel';
import { findAvailableServicePosition } from './utils/serviceNodePlacement';
import { MEDIA_QUERIES } from './styles/breakpoints';
import { useMediaQuery } from './hooks/useMediaQuery';
import {
  applyAutomaticEdgeLabelOffsets,
  shouldRecalculateAutomaticEdgeLabels,
} from './utils/edgeLabelLayout';
import { classifyEdgeDirection } from './utils/edgeDirection';
import { mapWorkflowStepsToEdges } from './utils/workflowStepMapping';
import { applySelectedVersionChanges } from './utils/versionDiff';
import {
  alignSelectedNodes,
  applyBulkNodeEdits,
  BULK_GROUP_COLORS,
  type BulkAlignmentType,
} from './utils/bulkNodeEditing';
import {
  analyzeDiagramQuality,
  applyDiagramQualityFixes,
  type DiagramQualityFinding,
} from './utils/diagramQuality';
import {
  getConnectionPresentation,
  normalizeConnectionType,
  type DiagramConnectionType,
} from './utils/edgePresentation';
import {
  expandDiagramContentBounds,
  screenRectToDiagramBounds,
  type DiagramContentBounds,
} from './utils/exportComposition';
import { LiveAnnouncer } from './components/LiveAnnouncer';
import { announce } from './a11y/liveAnnouncer';
import {
  KEYBOARD_CONNECT_EVENT,
  type KeyboardConnectDetail,
} from './hooks/useKeyboardConnection';

type LazyFeatureBoundaryProps = {
  active: boolean;
  children: React.ReactNode;
  onClose?: () => void;
};

class LazyFeatureBoundary extends React.Component<
  LazyFeatureBoundaryProps,
  { error: Error | null }
> {
  state = { error: null };

  static getDerivedStateFromError(error: Error) {
    return { error };
  }

  componentDidCatch(error: Error) {
    console.error('Failed to load an optional feature chunk:', error);
  }

  render() {
    if (this.state.error) {
      if (!this.props.active) return null;
      return (
        <div className="modal-overlay" role="alert">
          <div className="modal-content" style={{ maxWidth: 480, padding: 24 }}>
            <h2>Feature unavailable / 機能を読み込めません</h2>
            <p>
              Refresh the page and try again. Your current diagram remains open until you choose to reload.
              {' / '}
              ページを再読み込みして、もう一度お試しください。
            </p>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
              {this.props.onClose && (
                <button className="btn btn-secondary" onClick={this.props.onClose}>
                  Close / 閉じる
                </button>
              )}
              <button className="btn btn-primary" onClick={() => window.location.reload()}>
                Reload / 再読み込み
              </button>
            </div>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

function renderLazyFeature<T extends React.ComponentType<any>>(
  LazyComponent: React.LazyExoticComponent<T>,
  props: React.ComponentProps<T>,
) {
  const modalProps = props as { isOpen?: boolean; onClose?: () => void };
  return (
    <LazyFeatureBoundary active={Boolean(modalProps.isOpen)} onClose={modalProps.onClose}>
      <Suspense fallback={null}>
        <LazyComponent {...props} />
      </Suspense>
    </LazyFeatureBoundary>
  );
}

function lazyWhenOpen<T extends React.ComponentType<any>>(
  loader: () => Promise<{ default: T }>,
) {
  const LazyComponent = lazy(loader);
  return (props: React.ComponentProps<T>) => {
    if (!(props as { isOpen?: boolean }).isOpen) return null;
    return renderLazyFeature(LazyComponent, props);
  };
}

function lazyOnceOpened<T extends React.ComponentType<any>>(
  loader: () => Promise<{ default: T }>,
) {
  const LazyComponent = lazy(loader);
  return (props: React.ComponentProps<T>) => {
    const isOpen = Boolean((props as { isOpen?: boolean }).isOpen);
    const [wasOpened, setWasOpened] = useState(isOpen);
    useEffect(() => {
      if (isOpen) setWasOpened(true);
    }, [isOpen]);
    if (!wasOpened && !isOpen) return null;
    return renderLazyFeature(LazyComponent, props);
  };
}

const ValidationModal = lazyWhenOpen(() => import('./components/ValidationModal'));
const DeploymentGuideModal = lazyWhenOpen(() => import('./components/DeploymentGuideModal'));
const IaCRoundTripModal = lazyWhenOpen(() => import('./components/IaCRoundTripModal'));
const VersionHistoryModal = lazyWhenOpen(() => import('./components/VersionHistoryModal'));
const SaveSnapshotModal = lazyWhenOpen(() => import('./components/SaveSnapshotModal'));
const CloudWorkspaceModal = lazyWhenOpen(() => import('./components/CloudWorkspaceModal'));
const RecentWorkModal = lazyWhenOpen(() => import('./components/RecentWorkModal'));
const DiagramQualityDialog = lazyWhenOpen(() => import('./components/DiagramQualityDialog'));
const PricingScenarioModal = lazyWhenOpen(() => import('./components/PricingScenarioModal'));
const AzureImportModal = lazyWhenOpen(() => import('./components/AzureImportModal'));
const CompareModelsModal = lazyOnceOpened(() => import('./components/CompareModelsModal'));
const CompareValidationModal = lazyOnceOpened(() => import('./components/CompareValidationModal'));
const FeedbackModal = lazyWhenOpen(() => import('./components/FeedbackModal'));
const AccessManagementModal = lazyWhenOpen(() => import('./components/AccessManagementModal'));

async function captureDiagramAsPng(
  element: HTMLElement,
  options: CaptureOptions,
): Promise<string> {
  const capture = await import('./utils/captureCanvas');
  return capture.captureDiagramAsPng(element, options);
}

async function captureDiagramAsSvg(
  element: HTMLElement,
  options: CaptureOptions,
): Promise<string> {
  const capture = await import('./utils/captureCanvas');
  return capture.captureDiagramAsSvg(element, options);
}

async function exportReferenceArchitectureAsPng(
  architecture: ReferenceArchitecture,
): Promise<void> {
  const exporter = await import('./utils/exportReferencePng');
  await exporter.exportReferenceArchitectureAsPng(architecture);
}

async function exportBlueprintArchitectureAsPng(
  architecture: BlueprintArchitecture,
  options: { legendPosition?: 'bottom' | 'right' | 'auto' } = {},
): Promise<void> {
  const exporter = await import('./utils/exportBlueprintPng');
  await exporter.exportBlueprintArchitectureAsPng(architecture, options);
}

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
const EXPORT_BACKGROUND_STORAGE_KEY = 'azure-diagram-builder.exportBackground.v1';
const EDGE_STYLE_STORAGE_KEY = 'azure-diagram-builder.edgeStyle.v1';
const EDGE_ANIMATION_STORAGE_KEY = 'azure-diagram-builder.edgeAnimation.v1';
const CANVAS_HINT_STORAGE_KEY = 'azure-diagram-builder.canvasHintDismissed.v1';
const HEADER_COLLAPSED_STORAGE_KEY = 'azure-diagram-builder.headerCollapsed.v1';
const FOCUS_MODE_STORAGE_KEY = 'azure-diagram-builder.focusMode.v1';
const TOOLBAR_SECTIONS_STORAGE_KEY = 'azure-diagram-builder.toolbarSections.v1';
const RIBBON_TAB_STORAGE_KEY = 'azure-diagram-builder.ribbonTab.v1';
const THREAT_OVERLAY_STORAGE_KEY = 'azure-diagram-builder.threatOverlay.v1';
// Stable empty arrays so gated memos can skip work without creating a new
// reference each render (which would defeat downstream memoization).
const EMPTY_THREAT_MARKERS: ReturnType<typeof analyzeThreatModel> = [];
const EMPTY_SENSITIVE_FINDINGS: SensitiveFinding[] = [];
const DEFAULT_EDGE_COLOR = getConnectionPresentation('sync').stroke;
const EDGE_CONTEXT_MENU_WIDTH = 220;
const EDGE_CONTEXT_MENU_HEIGHT = 360;
const EDGE_CONTEXT_MENU_MARGIN = 8;
const NODE_CONTEXT_MENU_WIDTH = 280;
const NODE_CONTEXT_MENU_HEIGHT = 360;
const PANE_CONTEXT_MENU_WIDTH = 220;
const PANE_CONTEXT_MENU_HEIGHT = 190;

function createLocalDiagramLineageId(prefix = 'local'): string {
  const id = typeof crypto.randomUUID === 'function'
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
  return `${prefix}:${id}`;
}

type NodeContextMenuState = {
  x: number;
  y: number;
  nodeId: string;
};

type EdgeContextMenuState = {
  x: number;
  y: number;
  edgeId: string;
};

type PaneContextMenuState = {
  x: number;
  y: number;
  flowPosition: Node['position'];
};

type PrivacyRequest = {
  purpose: 'export' | 'share' | 'review';
  findings: SensitiveFinding[];
  canAnonymize: boolean;
  onProceed: () => void;
  onCancel: () => void;
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

function getContextMenuAnchor(
  event: React.MouseEvent,
  target: HTMLElement | null,
): { x: number; y: number } {
  if (event.clientX !== 0 || event.clientY !== 0) {
    return { x: event.clientX, y: event.clientY };
  }
  const bounds = target?.getBoundingClientRect();
  return bounds
    ? { x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2 }
    : { x: window.innerWidth / 2, y: window.innerHeight / 2 };
}

function handleContextMenuNavigation(
  event: React.KeyboardEvent<HTMLDivElement>,
  closeMenu: () => void,
) {
  const items = Array.from(
    event.currentTarget.querySelectorAll<HTMLButtonElement>('[role^="menuitem"]:not(:disabled)'),
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
    closeMenu();
    return;
  }

  if (nextIndex !== null) {
    event.preventDefault();
    event.stopPropagation();
    items[nextIndex].focus();
  }
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
    if (
      data.tags !== undefined
      && (
        !Array.isArray(data.tags)
        || data.tags.length > 12
        || data.tags.some(tag => typeof tag !== 'string' || tag.length > 40)
      )
    ) {
      throw new Error(`Node ${value.id} has invalid tags`);
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

  if (window.matchMedia(MEDIA_QUERIES.compactOrShortWorkspace).matches) {
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

const EDGE_LABEL_CAPTURE_PADDING = 16;

function getRenderedEdgeLabelBounds(
  wrapper: HTMLElement | null,
  instance: ReactFlowInstance | null,
): DiagramContentBounds[] {
  if (!wrapper || !instance) return [];

  const flowOrigin = instance.flowToScreenPosition({ x: 0, y: 0 });
  const { zoom } = instance.getViewport();

  return [...wrapper.querySelectorAll<HTMLElement>('.editable-edge-label-shell')]
    .filter((shell) => (
      !shell.querySelector('.editable-edge-label.is-empty')
      && shell.getClientRects().length > 0
    ))
    .flatMap((shell) => {
      const rect = shell.getBoundingClientRect();
      const bounds = screenRectToDiagramBounds({
        left: rect.left,
        top: rect.top,
        width: rect.width,
        height: rect.height,
      }, flowOrigin, zoom);
      return bounds ? [bounds] : [];
    });
}

interface DiagramHistorySnapshot {
  nodes: Node[];
  edges: Edge[];
  architecturePrompt: string;
  originalPrompt: string;
  validationScore?: number;
  titleBlockData: {
    architectureName: string;
    author: string;
    version: string;
    date: string;
  };
  workflow: any[];
  pricingScenarios: PricingScenario[];
  iacBaseline: IaCBaseline | null;
}

function stripTransientNodeState(node: Node): Node {
  const snapshot = { ...node } as Node & Record<string, unknown>;
  delete snapshot.selected;
  delete snapshot.dragging;
  delete snapshot.width;
  delete snapshot.height;
  delete snapshot.positionAbsolute;
  delete snapshot.measured;
  delete snapshot.resizing;
  return snapshot;
}

function stripTransientEdgeState(edge: Edge): Edge {
  const snapshot = { ...edge } as Edge & Record<string, unknown>;
  delete snapshot.selected;
  return snapshot;
}

function App() {
  const { t, translate, language } = useLanguage();
  // Locale tag for Date#toLocale* formatting so Japanese users get ja-JP dates.
  const localeTag = language === 'ja' ? 'ja-JP' : 'en-US';
  const [nodes, setNodes, onNodesChangeBase] = useNodesState([]);
  const latestNodesRef = useRef<Node[]>(nodes);
  latestNodesRef.current = nodes;
  const [edges, setEdges, onEdgesChange] = useEdgesState([]);
  const latestEdgesRef = useRef<Edge[]>(edges);
  latestEdgesRef.current = edges;
  const layoutGenerationRef = useRef(new OperationGeneration());
  const diagramRevisionGenerationRef = useRef(new OperationGeneration());
  const aiGenerationRef = useRef(new OperationGeneration());
  const validationGenerationRef = useRef(new OperationGeneration());
  const deploymentGuideGenerationRef = useRef(new OperationGeneration());
  const intentionalLineageTransitionRef = useRef<string | null>(null);
  const [localDiagramLineageId, setLocalDiagramLineageId] = useState(
    () => createLocalDiagramLineageId(),
  );
  const activeDiagramLineageIdRef = useRef(localDiagramLineageId);
  const [architecturePrompt, setArchitecturePrompt] = useState<string>('');
  // The FIRST prompt of the current diagram lineage. Unlike architecturePrompt
  // (which each chat refinement overwrites), this is captured once when the
  // canvas is empty so the customer deck's "brief" reflects the original ask.
  const [originalPrompt, setOriginalPrompt] = useState<string>('');

  const [isImportingTemplate, setIsImportingTemplate] = useState(false);
  const templateInputRef = useRef<HTMLInputElement>(null);
  const [isTemplateGalleryOpen, setIsTemplateGalleryOpen] = useState(false);
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
  const [reactFlowInstance, setReactFlowInstance] = useState<ReactFlowInstance | null>(null);
  const shouldRefreshAutomaticEdgeLabelsRef = useRef(false);
  
  const onNodesChange = useCallback((changes: NodeChange[]) => {
    if (shouldRecalculateAutomaticEdgeLabels(changes)) {
      shouldRefreshAutomaticEdgeLabelsRef.current = true;
    }
    onNodesChangeBase(changes);
  }, [onNodesChangeBase]);

  useEffect(() => {
    if (!shouldRefreshAutomaticEdgeLabelsRef.current) return;
    shouldRefreshAutomaticEdgeLabelsRef.current = false;
    setEdges(currentEdges => applyAutomaticEdgeLabelOffsets(nodes, currentEdges));
  }, [nodes, setEdges]);

  
  const [workflow, setWorkflow] = useState<any[]>([]);
  // const [showWorkflow, setShowWorkflow] = useState(false);
  const [highlightedServices, setHighlightedServices] = useState<string[]>([]);
  const [nodeContextMenu, setNodeContextMenu] = useState<NodeContextMenuState | null>(null);
  const nodeContextMenuRef = useRef<HTMLDivElement>(null);
  const nodeContextMenuReturnFocusRef = useRef<HTMLElement | null>(null);
  const [edgeContextMenu, setEdgeContextMenu] = useState<EdgeContextMenuState | null>(null);
  const edgeContextMenuRef = useRef<HTMLDivElement>(null);
  const edgeContextMenuReturnFocusRef = useRef<HTMLElement | null>(null);
  const [paneContextMenu, setPaneContextMenu] = useState<PaneContextMenuState | null>(null);
  const paneContextMenuRef = useRef<HTMLDivElement>(null);
  const paneContextMenuReturnFocusRef = useRef<HTMLElement | null>(null);
  const [totalMonthlyCost, setTotalMonthlyCost] = useState(0);
  // Services with no published price are excluded from the total; surface the
  // count so a partial estimate is never presented as complete.
  const [unpricedCount, setUnpricedCount] = useState(0);
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
  const pricingEditorReturnFocusRef = useRef<{
    runId: number;
    nodeId: string;
    element: HTMLElement | null;
  } | null>(null);
  const pricingEditorOpenRunRef = useRef(0);
  const cancelPendingPricingEditorOpen = useCallback(() => {
    pricingEditorOpenRunRef.current += 1;
    pricingEditorReturnFocusRef.current = null;
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
    date: new Date().toLocaleDateString(localeTag),
  });
  const latestTitleBlockDataRef = useRef(titleBlockData);
  latestTitleBlockDataRef.current = titleBlockData;

  useEffect(() => {
    diagramRevisionGenerationRef.current.advance();
  }, [architecturePrompt, edges, nodes, titleBlockData.architectureName]);

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
  const exportCanvasBackground = isDarkMode ? '#1a1a1a' : '#f8fafc';

  // Premium Features State
  const [validationResult, setValidationResult] = useState<ArchitectureValidation | null>(null);
  const [persistedValidationScore, setPersistedValidationScore] = useState<number | undefined>();
  const [validationNeedsRefresh, setValidationNeedsRefresh] = useState(false);
  const currentValidationScore = getCurrentValidationScore(
    validationResult?.overallScore ?? persistedValidationScore,
    validationNeedsRefresh,
  );
  const currentValidationResult = validationNeedsRefresh ? null : validationResult;
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
  const [isRecentWorkOpen, setIsRecentWorkOpen] = useState(false);
  const [isQualityDoctorOpen, setIsQualityDoctorOpen] = useState(false);
  const recentWorkSessionId = useMemo(() => getRecentWorkSessionId(), []);
  const [privacyRequest, setPrivacyRequest] = useState<PrivacyRequest | null>(null);
  const [threatOverlayEnabled, setThreatOverlayEnabled] = useState<boolean>(
    () => readBooleanPreference(THREAT_OVERLAY_STORAGE_KEY, false),
  );
  const [isPricingScenarioModalOpen, setIsPricingScenarioModalOpen] = useState(false);
  const [isCompareModelsOpen, setIsCompareModelsOpen] = useState(false);
  const [isChatOpen, setIsChatOpen] = useState(false);
  const [isDeliverChooserOpen, setIsDeliverChooserOpen] = useState(false);
  const [generatorOpenSignal, setGeneratorOpenSignal] = useState(0);
  const generatorOpenSourceRef = useRef<'first-start' | 'journey-strip' | 'toolbar'>('toolbar');
  const [isHelpOpen, setIsHelpOpen] = useState(false);
  const [isAboutOpen, setIsAboutOpen] = useState(false);
  const [isBYOAISettingsOpen, setIsBYOAISettingsOpen] = useState(false);
  const [isCommandPaletteOpen, setIsCommandPaletteOpen] = useState(false);
  const [isMobileRibbonOpen, setIsMobileRibbonOpen] = useState(false);
  const [paletteOpenSignal, setPaletteOpenSignal] = useState(0);
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
    return false;
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
      beforeApply?: () => string | void,
      reportErrors?: boolean,
      preserveValidationForRecheck?: boolean,
    ) => Promise<void>
  ) | null>(null);
  const feedbackAfterValidationRef = useRef(false);
  const [referenceImageUrl, setReferenceImageUrl] = useState<string | null>(null);
  const [lastReferenceArchitecture, setLastReferenceArchitecture] = useState<ReferenceArchitecture | null>(null);
  const [lastBlueprintArchitecture, setLastBlueprintArchitecture] = useState<BlueprintArchitecture | null>(null);
  const [panelsCollapsedSignal, setPanelsCollapsedSignal] = useState(0);
  const isCompactViewport = useMediaQuery(MEDIA_QUERIES.compact);
  const mobileCanvasFirstActive = isCompactViewport && nodes.length > 0;
  const [isMobileJourneyExpanded, setIsMobileJourneyExpanded] = useState(false);
  const wasMobileCanvasFirstRef = useRef(false);
  useEffect(() => {
    if (isChatOpen) setPanelsCollapsedSignal((previous) => previous + 1);
  }, [isChatOpen]);

  useEffect(() => {
    if (mobileCanvasFirstActive && !wasMobileCanvasFirstRef.current) {
      setIsMobileJourneyExpanded(false);
      setIsMobileRibbonOpen(false);
      setPanelsCollapsedSignal((previous) => previous + 1);
    }
    if (!mobileCanvasFirstActive) setIsMobileJourneyExpanded(false);
    wasMobileCanvasFirstRef.current = mobileCanvasFirstActive;
  }, [mobileCanvasFirstActive]);

  useEffect(() => {
    if (nodes.length === 0) {
      setValidationHandoff(null);
      feedbackAfterValidationRef.current = false;
    }
  }, [nodes.length]);
  const [focusMode, setFocusMode] = useState<boolean>(() => (
    readBooleanPreference(FOCUS_MODE_STORAGE_KEY, false)
  ));
  const focusModeReturnTargetRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    writeLocalStorage(FOCUS_MODE_STORAGE_KEY, focusMode ? '1' : '0');
    if (!focusMode) return;
    setIsMobileRibbonOpen(false);
    setIsCommandPaletteOpen(false);
    setIsChatOpen(false);
    setIsHelpOpen(false);
    setIsCloudWorkspaceOpen(false);
    setIsRecentWorkOpen(false);
    setIsQualityDoctorOpen(false);
    setPanelsCollapsedSignal((previous) => previous + 1);
  }, [focusMode]);

  const openCommandPalette = useCallback(() => {
    setIsMobileRibbonOpen(false);
    setIsChatOpen(false);
    setIsHelpOpen(false);
    setIsCloudWorkspaceOpen(false);
    setIsRecentWorkOpen(false);
    setIsQualityDoctorOpen(false);
    setPanelsCollapsedSignal((previous) => previous + 1);
    setIsCommandPaletteOpen(true);
  }, []);

  const openServicesPanel = useCallback(() => {
    setFocusMode(false);
    setIsMobileRibbonOpen(false);
    setIsCommandPaletteOpen(false);
    setIsChatOpen(false);
    setIsHelpOpen(false);
    setIsCloudWorkspaceOpen(false);
    setIsRecentWorkOpen(false);
    setIsQualityDoctorOpen(false);
    setPanelsCollapsedSignal((previous) => previous + 1);
    setPaletteOpenSignal((previous) => previous + 1);
  }, []);

  const toggleChatPanel = useCallback(() => {
    setFocusMode(false);
    setIsMobileRibbonOpen(false);
    setIsCommandPaletteOpen(false);
    setIsHelpOpen(false);
    setIsCloudWorkspaceOpen(false);
    setIsRecentWorkOpen(false);
    setIsQualityDoctorOpen(false);
    setIsChatOpen((current) => !current);
  }, []);

  const openCloudWorkspace = useCallback(() => {
    setFocusMode(false);
    setIsMobileRibbonOpen(false);
    setIsCommandPaletteOpen(false);
    setIsChatOpen(false);
    setIsHelpOpen(false);
    setIsRecentWorkOpen(false);
    setIsQualityDoctorOpen(false);
    setPanelsCollapsedSignal((previous) => previous + 1);
    setIsCloudWorkspaceOpen(true);
  }, []);

  const openRecentWork = useCallback(() => {
    setFocusMode(false);
    setIsMobileRibbonOpen(false);
    setIsCommandPaletteOpen(false);
    setIsChatOpen(false);
    setIsHelpOpen(false);
    setIsCloudWorkspaceOpen(false);
    setIsQualityDoctorOpen(false);
    setPanelsCollapsedSignal((previous) => previous + 1);
    setIsRecentWorkOpen(true);
  }, []);

  const locateReviewAnchor = useCallback((anchor: CloudCommentAnchor) => {
    const currentNodes = latestNodesRef.current;
    const currentEdges = latestEdgesRef.current;
    if (anchor.type === 'canvas' || !anchor.targetId) {
      setNodes(nodes => nodes.map(node => ({ ...node, selected: false })));
      setEdges(edges => edges.map(edge => ({ ...edge, selected: false })));
      window.requestAnimationFrame(() => {
        void reactFlowInstance?.fitView({ padding: 0.2, duration: 350, maxZoom: 1.2 });
      });
      return;
    }

    const selectedNodeIds = new Set<string>();
    const selectedEdgeIds = new Set<string>();
    if (anchor.type === 'node') {
      selectedNodeIds.add(anchor.targetId);
    } else {
      selectedEdgeIds.add(anchor.targetId);
      const targetEdge = currentEdges.find(edge => edge.id === anchor.targetId);
      if (targetEdge) {
        selectedNodeIds.add(targetEdge.source);
        selectedNodeIds.add(targetEdge.target);
      }
    }
    setNodes(nodes => nodes.map(node => ({
      ...node,
      selected: selectedNodeIds.has(node.id),
    })));
    setEdges(edges => edges.map(edge => ({
      ...edge,
      selected: selectedEdgeIds.has(edge.id),
    })));
    const targetNodes = currentNodes.filter(node => selectedNodeIds.has(node.id));
    window.requestAnimationFrame(() => {
      void reactFlowInstance?.fitView({
        nodes: targetNodes,
        padding: 0.4,
        duration: 350,
        maxZoom: 1.35,
      });
      reactFlowWrapper.current?.focus();
    });
  }, [reactFlowInstance, setEdges, setNodes]);

  const openQualityDoctor = useCallback(() => {
    setFocusMode(false);
    setIsMobileRibbonOpen(false);
    setIsCommandPaletteOpen(false);
    setIsChatOpen(false);
    setIsHelpOpen(false);
    setIsCloudWorkspaceOpen(false);
    setIsRecentWorkOpen(false);
    setPanelsCollapsedSignal((previous) => previous + 1);
    setIsQualityDoctorOpen(true);
  }, []);

  const enterFocusMode = useCallback(() => {
    const activeElement = document.activeElement;
    focusModeReturnTargetRef.current = activeElement instanceof HTMLElement
      && activeElement !== document.body
      ? activeElement
      : null;
    setFocusMode(true);
  }, []);

  const exitFocusMode = useCallback(() => {
    setFocusMode(false);
    window.requestAnimationFrame(() => {
      const returnTarget = focusModeReturnTargetRef.current;
      focusModeReturnTargetRef.current = null;
      if (
        returnTarget
        && returnTarget.isConnected
        && returnTarget.getClientRects().length > 0
        && !returnTarget.closest('[inert], [aria-hidden="true"]')
      ) {
        returnTarget.focus();
        if (document.activeElement === returnTarget) return;
      }
      reactFlowWrapper.current?.focus();
    });
  }, []);

  const toggleFocusMode = useCallback(() => {
    if (focusMode) exitFocusMode();
    else enterFocusMode();
  }, [enterFocusMode, exitFocusMode, focusMode]);

  useEffect(() => {
    if (!validationHandoff || focusMode || validationHandoffShownRef.current === validationHandoff) return;
    validationHandoffShownRef.current = validationHandoff;
    trackValidationHandoff({ action: 'shown', ...validationHandoff });
  }, [focusMode, validationHandoff]);

  const [isExportMenuOpen, setIsExportMenuOpen] = useState(false);
  const [exportBackground, setExportBackground] = useState<ExportBackground>(() => {
    const saved = localStorage.getItem(EXPORT_BACKGROUND_STORAGE_KEY);
    return saved === 'dots' || saved === 'grid' ? saved : 'plain';
  });
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
  const [layoutEngine, setLayoutEngine] = useState<LayoutEngineType>('elk');
  
  const [isBulkSelectMenuOpen, setIsBulkSelectMenuOpen] = useState(false);
  const bulkSelectMenuRef = useRef<HTMLDivElement | null>(null);
  
  const [isStylePresetMenuOpen, setIsStylePresetMenuOpen] = useState(false);
  const stylePresetMenuRef = useRef<HTMLDivElement | null>(null);

  const [isModelSettingsOpen, setIsModelSettingsOpen] = useState(false);
  const modelSettingsRef = useRef<HTMLDivElement | null>(null);
  const [stylePreset, setStylePreset] = useState<'detailed' | 'presentation'>('detailed');

  // Collapse / expand all groups
  const [allGroupsCollapsed, setAllGroupsCollapsed] = useState(false);
  const preCollapseGroupLayout = useRef<GroupLayoutSnapshot>(new Map());



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
    // Exports finish asynchronously and never move focus, so this is the only
    // completion signal a screen-reader user gets (WCAG 4.1.3).
    announce(localize(language, {
      en: `Export complete. ${fileName} downloaded.`,
      ja: `エクスポートが完了しました。${fileName} をダウンロードしました。`,
    }));
  }, [language]);

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

  // Keep edge rendering style in sync even without re-layout.
  useEffect(() => {
    setEdges((eds) =>
      eds.map((e) => ({
        ...e,
        data: { ...(e.data ?? {}), pathStyle: layoutEdgeStyle },
      }))
    );
  }, [layoutEdgeStyle, setEdges]);

  const addGroupBoxAtPosition = useCallback((position: Node['position']) => {
    const newNode: Node = {
      id: `group-${Date.now()}`,
      type: 'groupNode',
      position,
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

  const addGroupBox = useCallback(() => {
    addGroupBoxAtPosition({ x: 250, y: 150 });
  }, [addGroupBoxAtPosition]);

  // Apply dark mode class to body and persist preference
  useEffect(() => {
    if (isDarkMode) {
      document.body.classList.add('dark-mode');
    } else {
      document.body.classList.remove('dark-mode');
    }
    writeLocalStorage('darkMode', JSON.stringify(isDarkMode));
  }, [isDarkMode]);

  // Preload pricing data once the browser is idle so it never competes with
  // first paint. Falls back to a short timeout where requestIdleCallback is
  // unavailable.
  useEffect(() => {
    let idleHandle: number | undefined;
    let timeoutHandle: ReturnType<typeof setTimeout> | undefined;
    const run = () => {
      preloadCommonServices().catch(err =>
        console.warn('Failed to preload regional pricing:', err)
      );
      prefetchCommonServices(getActiveRegion()).catch(err =>
        console.warn('Failed to prefetch API pricing:', err)
      );
    };
    if (typeof requestIdleCallback === 'function') {
      idleHandle = requestIdleCallback(run, { timeout: 4000 });
    } else {
      timeoutHandle = setTimeout(run, 2000);
    }
    return () => {
      if (idleHandle !== undefined && typeof cancelIdleCallback === 'function') {
        cancelIdleCallback(idleHandle);
      }
      if (timeoutHandle !== undefined) {
        clearTimeout(timeoutHandle);
      }
    };
  }, []);

  // A structural signature of every node's pricing-relevant fields (ids +
  // pricing config), deliberately excluding positions so drag frames do not
  // change it and therefore skip the cost recompute below.
  const costSignature = useMemo(
    () => nodes.map(node => `${node.id}:${nodePricingFingerprint(node)}`).join('|'),
    [nodes],
  );

  // Recalculate total cost when the pricing signature changes. Debounced so a
  // burst of edits collapses into a single recompute + render pass, and reads
  // the freshest nodes from the ref when it fires.
  useEffect(() => {
    const handle = setTimeout(() => {
      const breakdown = calculateCostBreakdown(latestNodesRef.current, undefined, pricingMode);
      setTotalMonthlyCost(breakdown.totalMonthlyCost);
      setUnpricedCount(breakdown.unpricedServices?.length ?? 0);
    }, 150);
    return () => clearTimeout(handle);
    // costSignature intentionally keys this effect so position-only node
    // changes are skipped; nodes are read fresh from the ref in the timer.
  }, [costSignature, pricingMode]);

  // AlignmentToolbar only reads the selection's ids/types and each group's
  // id/label, never positions. Key both memos on a position-free signature so
  // drag frames reuse the same array references instead of reallocating (which
  // would defeat the toolbar's memoization).
  const alignmentSelectionSignature = useMemo(
    () => nodes.filter(node => node.selected).map(node => `${node.id}:${node.type ?? ''}`).join('|'),
    [nodes],
  );
  const alignmentSelectedNodes = useMemo(
    () => nodes.filter(node => node.selected),
    // Keyed on the selection signature so drag frames keep the same reference.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [alignmentSelectionSignature],
  );
  const alignmentGroupSignature = useMemo(
    () => nodes
      .filter(node => node.type === 'groupNode')
      .map(node => `${node.id}:${String(node.data?.label || 'Group')}`)
      .join('|'),
    [nodes],
  );
  const alignmentGroups = useMemo(
    () => nodes
      .filter(node => node.type === 'groupNode')
      .map(node => ({ id: node.id, label: String(node.data?.label || 'Group') })),
    // Keyed on the group signature so drag frames keep the same reference.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [alignmentGroupSignature],
  );

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
  }, [cancelPendingPricingEditorOpen, language, nodes, setNodes]);

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
              labelOffsetAuto: false,
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
      const presentation = getConnectionPresentation(next.data?.connectionType);
      const semanticStyle = {
        ...next.style,
        stroke: presentation.stroke,
      };
      if (presentation.strokeDasharray) {
        semanticStyle.strokeDasharray = presentation.strokeDasharray;
      } else {
        delete semanticStyle.strokeDasharray;
      }
      if (presentation.opacity !== undefined) {
        semanticStyle.opacity = presentation.opacity;
      } else {
        delete semanticStyle.opacity;
      }
      next.style = semanticStyle;
      if (next.markerEnd && typeof next.markerEnd === 'object') {
        next.markerEnd = { ...next.markerEnd, color: presentation.stroke };
      }
      if (next.markerStart && typeof next.markerStart === 'object') {
        next.markerStart = { ...next.markerStart, color: presentation.stroke };
      }
      const baseFlowAnimated = Boolean(
        next.data?.baseFlowAnimated
        ?? next.data?.flowAnimated
        ?? presentation.baseFlowAnimated,
      );
      const edgeAnimationPreference = typeof next.data?.flowAnimated === 'boolean'
        ? next.data.flowAnimated
        : baseFlowAnimated;
      const labelOffsetX = Number(next.data?.labelOffsetX) || 0;
      const labelOffsetY = Number(next.data?.labelOffsetY) || 0;
      next.data = {
        ...next.data,
        connectionType: presentation.type,
        baseFlowAnimated,
        flowAnimated: animateConnections && edgeAnimationPreference,
        pathStyle: normalizeLayoutEdgeStyle(next.data?.pathStyle ?? layoutEdgeStyle),
        labelOffsetAuto: typeof next.data?.labelOffsetAuto === 'boolean'
          ? next.data.labelOffsetAuto
          : labelOffsetX === 0 && labelOffsetY === 0,
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
      markerEnd: { type: MarkerType.ArrowClosed, color: DEFAULT_EDGE_COLOR },
      labelStyle: { fontSize: 13, fill: '#334155', fontWeight: 600 },
      labelBgStyle: { fill: 'white', fillOpacity: 0.94, stroke: '#cbd5e1', strokeWidth: 1 },
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
        labelOffsetAuto: true,
      },
    }, eds)),
    [setEdges, handleEdgeLabelChange, handleEdgeLabelOffsetChange, animateConnections, layoutEdgeStyle]
  );

  // Keyboard-created edges (press C on the source node, then C on the target)
  // are funnelled through the same onConnect as pointer drags so both routes
  // produce identical edges. Handles match React Flow's default left-to-right
  // flow; the edge renderer re-routes them anyway.
  //
  // The pending source is held in a module store with no lifecycle tie to the
  // node it names, so it can outlive a delete, an AI regeneration or a diagram
  // load. React Flow's addEdge does not check that either endpoint still
  // exists and its renderer then drops the edge without a warning, leaving a
  // dangling edge in saved JSON and exports — hence the explicit re-check here,
  // against live instance state rather than a captured closure.
  useEffect(() => {
    const handleKeyboardConnect = (event: Event) => {
      const detail = (event as CustomEvent<KeyboardConnectDetail>).detail;
      if (!detail?.source || !detail?.target || detail.source === detail.target) return;

      const failed = localize(language, {
        en: 'Connection failed. One of the nodes is no longer on the canvas.',
        ja: '接続できませんでした。対象のノードがキャンバス上にありません。',
      });
      const sourceNode = reactFlowInstance?.getNode(detail.source);
      const targetNode = reactFlowInstance?.getNode(detail.target);
      if (!sourceNode || !targetNode) {
        announce(failed, 'assertive');
        return;
      }

      const sourceHandle = 'right';
      const targetHandle = 'left';
      // addEdge silently drops an identical connection, which would otherwise
      // be announced as a second successful connection.
      const alreadyConnected = (reactFlowInstance?.getEdges() ?? []).some((edge) => (
        edge.source === detail.source
        && edge.target === detail.target
        && (edge.sourceHandle ?? null) === sourceHandle
        && (edge.targetHandle ?? null) === targetHandle
      ));

      const sourceLabel = String(sourceNode.data?.label ?? detail.source);
      const targetLabel = String(targetNode.data?.label ?? detail.target);

      if (alreadyConnected) {
        announce(localize(language, {
          en: `${sourceLabel} is already connected to ${targetLabel}.`,
          ja: `${sourceLabel} は既に ${targetLabel} に接続されています。`,
        }));
        return;
      }

      onConnect({
        source: detail.source,
        target: detail.target,
        sourceHandle,
        targetHandle,
      });
      announce(localize(language, {
        en: `Connected ${sourceLabel} to ${targetLabel}.`,
        ja: `${sourceLabel} から ${targetLabel} へ接続しました。`,
      }));
    };
    window.addEventListener(KEYBOARD_CONNECT_EVENT, handleKeyboardConnect);
    return () => window.removeEventListener(KEYBOARD_CONNECT_EVENT, handleKeyboardConnect);
  }, [language, onConnect, reactFlowInstance]);

  const diagramHistoryState = useMemo<DiagramHistorySnapshot>(() => ({
    nodes: nodes.map(stripTransientNodeState),
    edges: edges.map(stripTransientEdgeState),
    architecturePrompt,
    originalPrompt,
    validationScore: currentValidationScore,
    titleBlockData,
    workflow,
    pricingScenarios,
    iacBaseline,
  }), [
    architecturePrompt,
    currentValidationScore,
    edges,
    iacBaseline,
    nodes,
    originalPrompt,
    pricingScenarios,
    titleBlockData,
    workflow,
  ]);
  const diagramHistoryStateRef = useRef(diagramHistoryState);
  diagramHistoryStateRef.current = diagramHistoryState;

  const restoreDiagramHistory = useCallback((snapshot: DiagramHistorySnapshot) => {
    setNodes(snapshot.nodes);
    setEdges(normalizeRestoredEdges(snapshot.edges));
    setArchitecturePrompt(snapshot.architecturePrompt);
    setOriginalPrompt(snapshot.originalPrompt);
    setTitleBlockData(snapshot.titleBlockData);
    setWorkflow(snapshot.workflow);
    setPricingScenarios(snapshot.pricingScenarios);
    setIaCBaseline(snapshot.iacBaseline);
    setDriftPlanSummary(null);
    setValidationResult(null);
    setPersistedValidationScore(snapshot.validationScore);
    setValidationNeedsRefresh(false);
    setValidationHandoff(null);
    feedbackAfterValidationRef.current = false;
    setDeploymentGuide(null);
    setReferenceImageUrl(null);
    setLastReferenceArchitecture(null);
    setLastBlueprintArchitecture(null);
    window.requestAnimationFrame(() => {
      reactFlowWrapper.current?.focus();
    });
  }, [normalizeRestoredEdges, setEdges, setNodes]);

  const {
    canUndo: canUndoDiagram,
    canRedo: canRedoDiagram,
    undo: undoDiagram,
    redo: redoDiagram,
    reset: resetDiagramHistory,
  } = useDiagramHistory(diagramHistoryState, restoreDiagramHistory, {
    delayMs: 250,
    limit: 50,
  });

  // Keyboard shortcuts: undo/redo, delete, and duplicate.
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') cancelPendingPricingEditorOpen();

      const target = e.target;
      if (
        !(target instanceof Element)
        || e.defaultPrevented
        || e.isComposing
        || e.keyCode === 229
        || pricingEditorNodeId
        || target.tagName === 'INPUT'
        || target.tagName === 'TEXTAREA'
        || target.tagName === 'SELECT'
        || (target.tagName === 'BUTTON' && !(focusMode && e.key === 'Escape'))
        || (target instanceof HTMLElement && target.isContentEditable)
        || target.closest('[role="dialog"], [role="menu"]')
        || document.querySelector('[role="dialog"], [role="menu"]')
      ) {
        return;
      }

      if (e.key === 'Escape' && focusMode) {
        e.preventDefault();
        exitFocusMode();
        return;
      }

      if (
        (e.ctrlKey || e.metaKey)
        && !e.altKey
        && e.key.toLowerCase() === 'k'
      ) {
        e.preventDefault();
        openCommandPalette();
        return;
      }

      if ((e.ctrlKey || e.metaKey) && !e.altKey && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) redoDiagram();
        else undoDiagram();
        return;
      }

      if (
        (e.ctrlKey || e.metaKey)
        && !e.altKey
        && e.key.toLowerCase() === 'y'
      ) {
        e.preventDefault();
        redoDiagram();
        return;
      }

      if (e.key === 'Delete' || e.key === 'Backspace') {
        const selectedNodes = nodes.filter(n => n.selected);
        const selectedEdges = edges.filter(e => e.selected);

        if (selectedNodes.length > 0 || selectedEdges.length > 0) {
          e.preventDefault();
          if (selectedNodes.length > 0) {
            deleteCanvasNodes(selectedNodes.map(n => n.id));
          }
          if (selectedEdges.length > 0) {
            const edgeIdsToRemove = selectedEdges.map(edge => edge.id);
            setEdges(current => current.filter(edge => !edgeIdsToRemove.includes(edge.id)));
          }
          window.requestAnimationFrame(() => {
            reactFlowWrapper.current?.focus();
          });
        }
        return;
      }

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'd') {
        const selectedNodes = nodes.filter(n => n.selected);
        if (selectedNodes.length === 0) return;

        e.preventDefault();
        const duplicateRunId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
        let duplicateSequence = 0;
        const result = duplicateSelectedSubgraph(
          nodes,
          edges,
          selectedNodes.map(node => node.id),
          (kind, sourceId) => (
            `${sourceId}-copy-${kind}-${duplicateRunId}-${duplicateSequence++}`
          ),
        );
        setNodes(result.nodes);
        setEdges(result.edges);
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [
    cancelPendingPricingEditorOpen,
    deleteCanvasNodes,
    edges,
    exitFocusMode,
    focusMode,
    nodes,
    openCommandPalette,
    pricingEditorNodeId,
    redoDiagram,
    setEdges,
    setNodes,
    undoDiagram,
  ]);

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
      preCollapseGroupLayout.current = captureGroupLayout(nodes);

      // Collapse all groups to fit content
      const collapsed = fitAllGroupsToContent(nodes);
      setNodes(collapsed);
      setAllGroupsCollapsed(true);

      // Zoom out to show the full picture
      setTimeout(() => {
        reactFlowInstance?.fitView?.({ padding: 0.3, duration: 300, maxZoom: 1.2 });
      }, 50);
    } else {
      const snapshot = preCollapseGroupLayout.current;
      setNodes(nds => restoreGroupLayout(nds, snapshot));
      setAllGroupsCollapsed(false);
      preCollapseGroupLayout.current = new Map();

      setTimeout(() => {
        reactFlowInstance?.fitView?.({ padding: 0.2, duration: 300, maxZoom: 1.2 });
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
    const sourceNodes = latestNodesRef.current;
    const sourceEdges = latestEdgesRef.current;
    const generation = layoutGenerationRef.current.advance();
    const selectedAzureNodeId = sourceNodes.find(
      (node) => node.type === 'azureNode' && node.selected,
    )?.id;
    const shouldEmphasize =
      layoutEmphasizePrimaryPath && (layoutPreset === 'flow-lr' || layoutPreset === 'flow-tb');

    try {
      const result = await applyLayoutPreset(sourceNodes, sourceEdges, {
        preset: layoutPreset,
        spacing: layoutSpacing,
        edgeStyle: layoutEdgeStyle,
        emphasizePrimaryPath: shouldEmphasize,
        selectedNodeId: selectedAzureNodeId,
        layoutEngine,
      });
      if (!layoutGenerationRef.current.isCurrent(generation)) return;

      const arrangedEdges = applyAutomaticEdgeLabelOffsets(result.nodes, result.edges);
      setNodes(currentNodes => mergeLayoutNodes(currentNodes, sourceNodes, result.nodes));
      setEdges(currentEdges => mergeLayoutEdges(currentEdges, sourceEdges, arrangedEdges));

      requestAnimationFrame(() => {
        reactFlowInstance?.fitView?.({ padding: 0.2, duration: 250, maxZoom: 1.2 });
      });
    } catch (error) {
      if (!layoutGenerationRef.current.isCurrent(generation)) return;
      console.error('Failed to apply diagram layout:', error);
      alert(localize(language, {
        en: 'Failed to arrange the diagram. Please try again.',
        ja: '図を再配置できませんでした。もう一度お試しください。',
      }));
    }
  }, [
    layoutPreset,
    layoutSpacing,
    layoutEdgeStyle,
    layoutEmphasizePrimaryPath,
    layoutEngine,
    reactFlowInstance,
    setNodes,
    setEdges,
    language,
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
      const focusTarget = returnFocus?.isConnected ? returnFocus : reactFlowWrapper.current;
      focusTarget?.focus();
    });
  }, []);

  const dismissNodeContextMenu = useCallback(() => {
    nodeContextMenuReturnFocusRef.current = null;
    setNodeContextMenu(null);
  }, []);

  const closeEdgeContextMenu = useCallback(() => {
    setEdgeContextMenu(null);
    const returnFocus = edgeContextMenuReturnFocusRef.current;
    edgeContextMenuReturnFocusRef.current = null;
    window.requestAnimationFrame(() => {
      const focusTarget = returnFocus?.isConnected ? returnFocus : reactFlowWrapper.current;
      focusTarget?.focus();
    });
  }, []);

  const dismissEdgeContextMenu = useCallback(() => {
    edgeContextMenuReturnFocusRef.current = null;
    setEdgeContextMenu(null);
  }, []);

  const closePaneContextMenu = useCallback(() => {
    setPaneContextMenu(null);
    const returnFocus = paneContextMenuReturnFocusRef.current;
    paneContextMenuReturnFocusRef.current = null;
    window.requestAnimationFrame(() => {
      const focusTarget = returnFocus?.isConnected ? returnFocus : reactFlowWrapper.current;
      focusTarget?.focus();
    });
  }, []);

  const dismissPaneContextMenu = useCallback(() => {
    paneContextMenuReturnFocusRef.current = null;
    setPaneContextMenu(null);
  }, []);

  const closeCanvasContextMenus = useCallback(() => {
    cancelPendingPricingEditorOpen();
    if (nodeContextMenu) {
      closeNodeContextMenu();
    } else if (edgeContextMenu) {
      closeEdgeContextMenu();
    } else if (paneContextMenu) {
      closePaneContextMenu();
    }
  }, [
    cancelPendingPricingEditorOpen,
    closeEdgeContextMenu,
    closeNodeContextMenu,
    closePaneContextMenu,
    edgeContextMenu,
    nodeContextMenu,
    paneContextMenu,
  ]);

  const dismissCanvasContextMenus = useCallback(() => {
    cancelPendingPricingEditorOpen();
    dismissNodeContextMenu();
    dismissEdgeContextMenu();
    dismissPaneContextMenu();
  }, [
    cancelPendingPricingEditorOpen,
    dismissEdgeContextMenu,
    dismissNodeContextMenu,
    dismissPaneContextMenu,
  ]);

  const onNodeContextMenu = useCallback((event: React.MouseEvent, node: Node) => {
    cancelPendingPricingEditorOpen();
    event.preventDefault();
    event.stopPropagation();
    const eventTarget = event.target as HTMLElement;
    const nodeElement = eventTarget.closest<HTMLElement>('.react-flow__node');
    const focusTarget = eventTarget.closest<HTMLElement>(
      '[data-node-keyboard-target], button, input, textarea, select',
    ) || nodeElement?.querySelector<HTMLElement>('[data-node-keyboard-target]') || null;
    const anchor = getContextMenuAnchor(event, focusTarget);
    const position = clampContextMenuPosition(
      anchor.x,
      anchor.y,
      NODE_CONTEXT_MENU_WIDTH,
      NODE_CONTEXT_MENU_HEIGHT,
    );
    // Context-menu commands always target the clicked item, never a hidden
    // multi-selection that the menu does not describe.
    setNodes((currentNodes) => currentNodes.map(currentNode => ({
      ...currentNode,
      selected: currentNode.id === node.id,
    })));
    setEdges((currentEdges) => currentEdges.map(edge => ({ ...edge, selected: false })));
    nodeContextMenuReturnFocusRef.current = focusTarget
      || (document.activeElement instanceof HTMLElement ? document.activeElement : null);
    dismissEdgeContextMenu();
    dismissPaneContextMenu();
    setNodeContextMenu({ ...position, nodeId: node.id });
  }, [
    cancelPendingPricingEditorOpen,
    dismissEdgeContextMenu,
    dismissPaneContextMenu,
    setEdges,
    setNodes,
  ]);

  const onEdgeContextMenu = useCallback((event: React.MouseEvent, edge: Edge) => {
    cancelPendingPricingEditorOpen();
    event.preventDefault();
    event.stopPropagation();
    const eventTarget = event.target as HTMLElement;
    const focusTarget = eventTarget.closest<HTMLElement>(
      '.react-flow__edge, [data-edge-label-id]',
    );
    const anchor = getContextMenuAnchor(event, focusTarget);
    const position = clampContextMenuPosition(
      anchor.x,
      anchor.y,
      EDGE_CONTEXT_MENU_WIDTH,
      EDGE_CONTEXT_MENU_HEIGHT,
    );
    dismissNodeContextMenu();
    dismissPaneContextMenu();
    setNodes((currentNodes) => currentNodes.map(node => ({ ...node, selected: false })));
    setEdges((currentEdges) => currentEdges.map(currentEdge => ({
      ...currentEdge,
      selected: currentEdge.id === edge.id,
    })));
    edgeContextMenuReturnFocusRef.current = focusTarget
      || (document.activeElement instanceof HTMLElement ? document.activeElement : null);
    setEdgeContextMenu({ ...position, edgeId: edge.id });
  }, [
    cancelPendingPricingEditorOpen,
    dismissNodeContextMenu,
    dismissPaneContextMenu,
    setEdges,
    setNodes,
  ]);

  const handleCanvasContextMenuCapture = useCallback((
    event: React.MouseEvent<HTMLDivElement>,
  ) => {
    const target = event.target;
    if (!(target instanceof Element)) return;
    const edgeLabel = target.closest<HTMLElement>('[data-edge-label-id]');
    const edgeId = edgeLabel?.dataset.edgeLabelId;
    const edge = edgeId ? edges.find(candidate => candidate.id === edgeId) : undefined;
    if (edge) onEdgeContextMenu(event, edge);
  }, [edges, onEdgeContextMenu]);

  const openPaneContextMenu = useCallback((
    clientX: number,
    clientY: number,
    returnFocus: HTMLElement | null,
  ) => {
    cancelPendingPricingEditorOpen();
    const position = clampContextMenuPosition(
      clientX,
      clientY,
      PANE_CONTEXT_MENU_WIDTH,
      PANE_CONTEXT_MENU_HEIGHT,
    );
    const flowPosition = reactFlowInstance?.screenToFlowPosition?.({
      x: clientX,
      y: clientY,
    }) ?? { x: 250, y: 150 };
    dismissNodeContextMenu();
    dismissEdgeContextMenu();
    setNodes((currentNodes) => currentNodes.map(node => ({ ...node, selected: false })));
    setEdges((currentEdges) => currentEdges.map(edge => ({ ...edge, selected: false })));
    paneContextMenuReturnFocusRef.current = returnFocus;
    setPaneContextMenu({ ...position, flowPosition });
  }, [
    cancelPendingPricingEditorOpen,
    dismissEdgeContextMenu,
    dismissNodeContextMenu,
    reactFlowInstance,
    setEdges,
    setNodes,
  ]);

  const onPaneContextMenu = useCallback((event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    const anchor = getContextMenuAnchor(event, reactFlowWrapper.current);
    openPaneContextMenu(anchor.x, anchor.y, reactFlowWrapper.current);
  }, [openPaneContextMenu]);

  const handleCanvasKeyDown = useCallback((event: React.KeyboardEvent<HTMLDivElement>) => {
    if (!((event.shiftKey && event.key === 'F10') || event.key === 'ContextMenu')) return;
    const target = event.target as HTMLElement;
    if (target.closest(
      '.react-flow__node, .react-flow__edge, [data-edge-label-id], button, input, textarea, select, [contenteditable="true"]',
    )) return;

    event.preventDefault();
    event.stopPropagation();
    const bounds = reactFlowWrapper.current?.getBoundingClientRect();
    openPaneContextMenu(
      bounds ? bounds.left + bounds.width / 2 : window.innerWidth / 2,
      bounds ? bounds.top + bounds.height / 2 : window.innerHeight / 2,
      reactFlowWrapper.current,
    );
  }, [openPaneContextMenu]);

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
    if (!node || node.type !== 'azureNode') {
      closeNodeContextMenu();
      return;
    }
    pricingEditorReturnFocusRef.current = {
      runId,
      nodeId,
      element: nodeContextMenuReturnFocusRef.current,
    };
    closeNodeContextMenu();

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
    ) {
      if (pricingEditorReturnFocusRef.current?.runId === runId) {
        pricingEditorReturnFocusRef.current = null;
      }
      return;
    }

    const latestNode = latestNodesRef.current.find(candidate => candidate.id === nodeId);
    const latestServiceType = String(
      latestNode?.data?.serviceName || latestNode?.data?.label || '',
    );
    if (!latestNode || latestNode.type !== 'azureNode' || latestServiceType !== serviceType) {
      if (pricingEditorReturnFocusRef.current?.runId === runId) {
        pricingEditorReturnFocusRef.current = null;
      }
      return;
    }

    const latestPricing = latestNode.data?.pricing as NodePricingConfig | undefined;
    setPricingEditorDraft(latestPricing
      ? null
      : {
          nodeId,
          pricing: initializedPricing ?? createCustomPricingDraft(region),
        });
    openNodePricingEditor(nodeId);
  }, [closeNodeContextMenu]);

  const closePricingEditor = useCallback(() => {
    cancelPendingPricingEditorOpen();
    setPricingEditorDraft(null);
    closeNodePricingEditor();
    pricingEditorReturnFocusRef.current = null;
  }, [cancelPendingPricingEditorOpen]);

  const fitContextGroupToContent = useCallback((groupId: string) => {
    setNodes((currentNodes) => fitGroupToContent(currentNodes, groupId) ?? currentNodes);
    closeNodeContextMenu();
  }, [closeNodeContextMenu, setNodes]);

  const detachContextNodeFromGroup = useCallback((nodeId: string) => {
    setNodes((currentNodes) => detachNodeFromGroup(currentNodes, nodeId));
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

  const deleteContextEdge = useCallback((edgeId: string) => {
    setEdges((currentEdges) => currentEdges.filter(edge => edge.id !== edgeId));
    closeEdgeContextMenu();
  }, [closeEdgeContextMenu, setEdges]);

  const addContextGroup = useCallback(() => {
    if (!paneContextMenu) return;
    addGroupBoxAtPosition(paneContextMenu.flowPosition);
    closePaneContextMenu();
  }, [addGroupBoxAtPosition, closePaneContextMenu, paneContextMenu]);

  const fitContextDiagram = useCallback(() => {
    reactFlowInstance?.fitView?.({ padding: 0.2, duration: 300, maxZoom: 1.2 });
    closePaneContextMenu();
  }, [closePaneContextMenu, reactFlowInstance]);

  const selectAllContextItems = useCallback(() => {
    setNodes((currentNodes) => currentNodes.map(node => ({ ...node, selected: true })));
    setEdges((currentEdges) => currentEdges.map(edge => ({ ...edge, selected: true })));
    closePaneContextMenu();
  }, [closePaneContextMenu, setEdges, setNodes]);

  useLayoutEffect(() => {
    const menu = nodeContextMenu
      ? nodeContextMenuRef.current
      : edgeContextMenu
        ? edgeContextMenuRef.current
        : paneContextMenu
          ? paneContextMenuRef.current
          : null;
    menu
      ?.querySelector<HTMLButtonElement>('[role^="menuitem"]:not(:disabled)')
      ?.focus();
  }, [edgeContextMenu, nodeContextMenu, paneContextMenu]);

  useEffect(() => {
    if (!nodeContextMenu && !edgeContextMenu && !paneContextMenu) return;
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
  }, [closeCanvasContextMenus, edgeContextMenu, nodeContextMenu, paneContextMenu]);

  const setConnectionAnimations = useCallback((enabled: boolean) => {
    setAnimateConnections(enabled);
    setEdges((currentEdges) => currentEdges.map((edge) => {
      const presentation = getConnectionPresentation(edge.data?.connectionType);
      const baseFlowAnimated = Boolean(
        edge.data?.baseFlowAnimated
        ?? edge.data?.flowAnimated
        ?? presentation.baseFlowAnimated,
      );
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
        const markerColor = getConnectionPresentation(edge.data?.connectionType).stroke;
        const baseFlowAnimated = Boolean(edge.data?.baseFlowAnimated ?? edge.data?.flowAnimated ?? true);
        const flowAnimated = animateConnections && baseFlowAnimated;
        const flowMode = direction === 'bidirectional' ? 'pulse' : 'directional';
        
        switch (direction) {
          case 'forward':
            markerEnd = { type: MarkerType.ArrowClosed, color: markerColor };
            break;
          case 'reverse':
            markerStart = { type: MarkerType.ArrowClosed, color: markerColor };
            break;
          case 'bidirectional':
            markerEnd = { type: MarkerType.ArrowClosed, color: markerColor };
            markerStart = { type: MarkerType.ArrowClosed, color: markerColor };
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
  }, [
    animateConnections,
    closeEdgeContextMenu,
    handleEdgeLabelChange,
    handleEdgeLabelOffsetChange,
    setEdges,
  ]);

  const onDragOver = useCallback((event: React.DragEvent) => {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
  }, []);

  const addServiceNodeAtPosition = useCallback((
    service: { iconPath: string; iconName: string; serviceName: string; category?: string },
    position: Node['position'],
    options: { avoidCollisions?: boolean } = {},
  ) => {
    if (!reactFlowInstance) return;

    const nodeId = `service-${globalThis.crypto.randomUUID()}`;
    const newNode: Node = {
      id: nodeId,
      type: 'azureNode',
      position,
      data: {
        label: service.iconName,
        serviceName: service.serviceName,
        category: service.category,
        iconPath: service.iconPath,
      },
    };

    setNodes((current) => {
      const parentGroup = current.find((node) => {
        if (node.type !== 'groupNode') return false;
        const groupWidth = (node.style?.width as number) || node.width || 400;
        const groupHeight = (node.style?.height as number) || node.height || 300;
        return position.x >= node.position.x
          && position.x <= node.position.x + groupWidth
          && position.y >= node.position.y
          && position.y <= node.position.y + groupHeight;
      });
      const relativePosition = parentGroup
        ? {
            x: position.x - parentGroup.position.x,
            y: position.y - parentGroup.position.y,
          }
        : position;
      const siblings = current.filter((node) => (
        node.type === 'azureNode'
        && node.parentNode === parentGroup?.id
      ));
      const resolvedPosition = options.avoidCollisions
        ? findAvailableServicePosition(relativePosition, siblings)
        : relativePosition;
      return current.concat({
        ...newNode,
        position: resolvedPosition,
        parentNode: parentGroup?.id,
        extent: parentGroup ? 'parent' : undefined,
      });
    });

    const currentRegion = getActiveRegion();
    void initializeNodePricing(service.serviceName, currentRegion)
      .then((pricing) => {
        if (!pricing) return;
        setNodes((current) => current.map((node) => (
          node.id === nodeId && !node.data.pricing
            ? { ...node, data: { ...node.data, pricing } }
            : node
        )));
      })
      .catch((error) => console.warn('Failed to initialize pricing:', error));
  }, [reactFlowInstance, setNodes]);

  const handleAddService = useCallback((icon: AzureIcon) => {
    if (!reactFlowInstance || !reactFlowWrapper.current) return;

    const bounds = reactFlowWrapper.current.getBoundingClientRect();
    const position = reactFlowInstance.screenToFlowPosition({
      x: bounds.left + (bounds.width / 2),
      y: bounds.top + (bounds.height / 2),
    });

    addServiceNodeAtPosition({
      iconPath: icon.path,
      iconName: icon.name,
      serviceName: icon.serviceName,
      category: icon.category,
    }, position, { avoidCollisions: true });
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

  const createDiagramCaptureOptions = useCallback((excludePanels = true): CaptureOptions => {
    const visibleNodes = latestNodesRef.current.filter(node => !node.hidden);
    const bounds = expandDiagramContentBounds(
      getNodesBounds(visibleNodes),
      getRenderedEdgeLabelBounds(reactFlowWrapper.current, reactFlowInstance),
      EDGE_LABEL_CAPTURE_PADDING,
    );
    const latestTitle = latestTitleBlockDataRef.current;
    const sync = getConnectionPresentation('sync');
    const asyncConnection = getConnectionPresentation('async');
    const optional = getConnectionPresentation('optional');
    const security = getConnectionPresentation('security');
    const telemetry = getConnectionPresentation('telemetry');
    const usedConnectionTypes = new Set<DiagramConnectionType>(
      latestEdgesRef.current
        .filter(edge => !edge.hidden)
        .map(edge => normalizeConnectionType(edge.data?.connectionType)),
    );
    const legendItems = [
      {
        type: 'sync' as const,
        label: t('Synchronous'),
        description: t('Real-time, request-response (HTTP, SQL)'),
        color: sync.stroke,
        lineStyle: 'solid' as const,
      },
      {
        type: 'async' as const,
        label: t('Asynchronous'),
        description: t('Message-based, event-driven (queues, events)'),
        color: asyncConnection.stroke,
        lineStyle: 'dashed' as const,
      },
      {
        type: 'optional' as const,
        label: t('Optional'),
        description: t('Conditional, fallback paths'),
        color: optional.stroke,
        lineStyle: 'dotted' as const,
      },
      {
        type: 'security' as const,
        label: t('Security'),
        description: t('Identity, trust, and policy enforcement'),
        color: security.stroke,
        lineStyle: 'dotted' as const,
      },
      {
        type: 'telemetry' as const,
        label: t('Telemetry'),
        description: t('Metrics, logs, traces, and diagnostics'),
        color: telemetry.stroke,
        lineStyle: 'dashed' as const,
      },
    ]
      .filter(item => usedConnectionTypes.has(item.type))
      .map(({ type: _type, ...item }) => item);

    return {
      backgroundColor: exportCanvasBackground,
      excludePanels,
      exportBackground,
      composition: {
        bounds,
        title: latestTitle.architectureName || 'Azure Architecture',
        subtitle: [
          latestTitle.author,
          latestTitle.date,
          latestTitle.version ? `v${latestTitle.version}` : '',
        ].filter(Boolean).join(' · '),
        legendTitle: localize(language, {
          en: 'Connection legend',
          ja: '接続凡例',
        }),
        legendItems,
      },
    };
  }, [
    exportBackground,
    exportCanvasBackground,
    language,
    reactFlowInstance,
    t,
  ]);

  const exportDiagram = useCallback(async () => {
    if (!reactFlowWrapper.current || !reactFlowInstance) {
      return;
    }

    try {
      const dataUrl = await captureDiagramAsPng(
        reactFlowWrapper.current,
        createDiagramCaptureOptions(false),
      );
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
      trackExport('png', nodes.filter(n => n.type === 'azureNode').length, exportBackground);
    } catch (err) {
      console.error('Error exporting diagram:', err);
      alert(t("Failed to export diagram. Please try again."));
    }
  }, [
    createDiagramCaptureOptions,
    exportBackground,
    nodes,
    reactFlowInstance,
    recordExport,
    t,
  ]);

  const exportAsSvg = useCallback(async () => {
    if (!reactFlowWrapper.current || !reactFlowInstance) {
      return;
    }

    try {
      const svgText = await captureDiagramAsSvg(
        reactFlowWrapper.current,
        createDiagramCaptureOptions(),
      );
      const blob = new Blob([svgText], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      const fileName = generateModelFilename('azure-diagram', 'svg');
      link.download = fileName;
      link.click();
      URL.revokeObjectURL(url);
      recordExport('svg', fileName);
      trackExport('svg', nodes.filter(n => n.type === 'azureNode').length, exportBackground);
    } catch (err) {
      console.error('Error exporting SVG:', err);
      alert(t("Failed to export SVG. Please try again."));
    }
  }, [
    createDiagramCaptureOptions,
    exportBackground,
    nodes,
    reactFlowInstance,
    recordExport,
    t,
  ]);

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
        validationScore: currentValidationScore ?? null,
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
  }, [nodes, edges, workflow, titleBlockData, architecturePrompt, generatedWithModel, currentValidationScore, pricingMode, recordExport, t]);

  // Export as an Animated SVG: same vector capture as exportAsSvg, but with
  // flowing data-flow circles injected onto each edge. Pure client-side — the
  // motion is carried by the SVG (open in a browser to view). For README/Teams
  // surfaces that strip SVG animation, export a GIF/WebP instead.
  const exportAsAnimatedSvg = useCallback(async () => {
    if (!reactFlowWrapper.current || !reactFlowInstance) {
      return;
    }

    try {
      const svgText = await captureDiagramAsSvg(
        reactFlowWrapper.current,
        createDiagramCaptureOptions(),
      );
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
      trackExport('animated-svg', nodes.filter(n => n.type === 'azureNode').length, exportBackground);
    } catch (err) {
      console.error('Error exporting animated SVG:', err);
      alert(t("Failed to export animated SVG. Please try again."));
    }
  }, [
    createDiagramCaptureOptions,
    exportBackground,
    nodes,
    reactFlowInstance,
    recordExport,
    t,
  ]);

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

    try {
      const svgText = await captureDiagramAsSvg(
        reactFlowWrapper.current,
        createDiagramCaptureOptions(),
      );
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
      trackExport('workflow-animation', nodes.filter(n => n.type === 'azureNode').length, exportBackground);
    } catch (err) {
      console.error('Error exporting workflow animation:', err);
      alert(t("Failed to export workflow animation. Please try again."));
    }
  }, [
    createDiagramCaptureOptions,
    edges,
    exportBackground,
    nodes,
    reactFlowInstance,
    recordExport,
    t,
    workflow,
  ]);

  const exportAsDrawio = useCallback(async () => {
    try {
      const { exportAndDownloadDrawio } = await import('./services/drawioExporter');
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
      const { buildVsdxBlob } = await import('./services/visioVsdxExporter');
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

  const exportAsHtml = useCallback(async () => {
    try {
      const { exportDiagramAsHtml } = await import('./services/htmlDiagramExporter');
      const diagramName = titleBlockData.architectureName || 'Azure Architecture';
      await exportDiagramAsHtml(nodes, edges, diagramName);
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

    try {
      const { exportDiagramAsPptx } = await import('./services/pptxExporter');
      const imageDataUrl = await captureDiagramAsPng(
        reactFlowWrapper.current,
        createDiagramCaptureOptions(),
      );
      const fileName = await exportDiagramAsPptx(imageDataUrl, {
        diagramName: titleBlockData.architectureName || 'Azure Architecture',
        author: titleBlockData.author || 'Azure Architect',
        date: titleBlockData.date || new Date().toLocaleDateString(localeTag),
        isDarkMode,
        // Supplying the canvas turns the slide into native, editable shapes
        // instead of a flat screenshot.
        diagram: { nodes, edges },
      });
      recordExport('pptx', fileName);
      trackExport('pptx', nodes.filter(n => n.type === 'azureNode').length, exportBackground);
    } catch (err) {
      console.error('Error exporting PPTX:', err);
      alert(t("Failed to export PowerPoint slide. Please try again."));
    }
  }, [
    createDiagramCaptureOptions,
    edges,
    exportBackground,
    isDarkMode,
    localeTag,
    nodes,
    reactFlowInstance,
    recordExport,
    t,
    titleBlockData,
  ]);

  const exportCustomerDeck = useCallback(async () => {
    if (!reactFlowWrapper.current || !reactFlowInstance) return;

    const azureNodes = nodes.filter(n => n.type === 'azureNode');
    if (azureNodes.length === 0) {
      alert(t('Add or generate an architecture first, then export a customer deck.'));
      return;
    }

    try {
        const { exportArchitectureDeck } = await import('./services/pptxExporter');
        const imageDataUrl = await captureDiagramAsPng(
          reactFlowWrapper.current,
          createDiagramCaptureOptions(),
        );

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
        const validation = currentValidationResult ? {
          overallScore: currentValidationResult.overallScore,
          overallLabel: bandLabel(currentValidationResult.overallScore),
          summary: currentValidationResult.summary,
          pillars: currentValidationResult.pillars.map(p => ({ pillar: p.pillar, score: p.score, maturity: bandLabel(p.score) })),
          findings: (currentValidationResult.quickWins.length > 0
            ? currentValidationResult.quickWins
            : currentValidationResult.pillars.flatMap(p => p.findings)
          )
            .slice()
            .sort((a, b) => {
              const rank = { critical: 0, high: 1, medium: 2, low: 3 } as Record<string, number>;
              return (rank[a.severity] ?? 9) - (rank[b.severity] ?? 9);
            })
            .slice(0, 6)
            .map(f => ({ severity: f.severity, category: f.category, issue: f.issue, recommendation: f.recommendation })),
          modelUsed: currentValidationResult.modelUsed,
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
          date: titleBlockData.date || new Date().toLocaleDateString(localeTag),
          isDarkMode,
          prompt: (originalPrompt || architecturePrompt) || undefined,
          model: generatedWithModel?.name,
          services,
          workflow: Array.isArray(workflow) && workflow.length > 0 ? workflow : null,
          validation,
          cost,
          diagram: { nodes, edges },
        });

        recordExport('pptx', fileName);
        trackExport('pptx-deck', azureNodes.length, exportBackground);
    } catch (err) {
      console.error('Error exporting customer deck:', err);
      alert(t('Failed to export the customer deck. Please try again.'));
    }
  }, [
    architecturePrompt,
    createDiagramCaptureOptions,
    currentValidationResult,
    edges,
    exportBackground,
    generatedWithModel,
    isDarkMode,
    localeTag,
    nodes,
    originalPrompt,
    pricingMode,
    reactFlowInstance,
    recordExport,
    t,
    titleBlockData,
    workflow,
  ]);

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
      validationScore: currentValidationScore,
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
  }, [reactFlowInstance, recordExport, titleBlockData, workflow, pricingScenarios, architecturePrompt, originalPrompt, nodes, iacBaseline, currentValidationScore]);

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
      `_Generated ${new Date().toLocaleString(localeTag)} · Region: \`${reportRegionLabel}\` · ${azureServiceNodes.length} service(s) on diagram_`,
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
    analysisLines.push('_Generated by Microsoft Product Architecture Diagram Builder. Estimates are indicative and exclude taxes, bandwidth egress, and support plans unless modeled explicitly._');


    // Build ZIP
    const { default: JSZip } = await import('jszip');
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
      `Generated ${new Date().toLocaleString(localeTag)} for region \`${reportRegionLabel}\` by Microsoft Product Architecture Diagram Builder.`,
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
  }, [localeTag, nodes, recordExport, pricingMode, t]);

  const prepareFlowObject = useCallback(
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

      // Restore metadata if present
      const restoredTitle = isRecord(flow.titleBlockData)
        ? flow.titleBlockData
        : flow.metadata;
      const titleBlock = isRecord(restoredTitle)
        ? {
          architectureName: typeof restoredTitle.architectureName === 'string'
            ? restoredTitle.architectureName
            : 'Untitled Architecture',
          author: typeof restoredTitle.author === 'string' ? restoredTitle.author : 'Azure Architect',
          version: typeof restoredTitle.version === 'string' ? restoredTitle.version : '1.0',
          date: typeof restoredTitle.date === 'string'
            ? restoredTitle.date
            : new Date().toLocaleDateString(localeTag),
        }
        : {
          architectureName: 'Untitled Architecture',
          author: 'Azure Architect',
          version: '1.0',
          date: new Date().toLocaleDateString(localeTag),
        };
      const pricing = normalizePricingScenarios(flow.pricingScenarios);
      const restoredIaCBaseline = restoreIaCBaseline(flow.iacBaseline);
      const validationScore = (
        typeof flow.validationScore === 'number' && Number.isFinite(flow.validationScore)
      )
        ? flow.validationScore
        : undefined;
      const restoredPrompt = typeof flow.architecturePrompt === 'string' ? flow.architecturePrompt : '';
      const restoredOriginalPrompt = typeof flow.originalPrompt === 'string'
        ? flow.originalPrompt
        : restoredPrompt;

      return {
        nodes: restoredNodes,
        edges: fixedEdges,
        viewport: restoredViewport,
        titleBlockData: titleBlock,
        workflow: restoredWorkflow,
        pricingScenarios: pricing,
        iacBaseline: restoredIaCBaseline,
        validationScore,
        architecturePrompt: restoredPrompt,
        originalPrompt: restoredOriginalPrompt,
      };
    },
    [localeTag, normalizeRestoredEdges],
  );

  const applyPreparedFlowObject = useCallback(
    (restored: ReturnType<typeof prepareFlowObject>) => {
      setNodes(restored.nodes);
      setEdges(restored.edges);

      if (restored.viewport && reactFlowInstance?.setViewport) {
        reactFlowInstance.setViewport(restored.viewport);
      }

      setTitleBlockData(restored.titleBlockData);

      // Restore workflow if present
      setWorkflow(restored.workflow);
      setPricingScenarios(restored.pricingScenarios);
      setIaCBaseline(restored.iacBaseline);
      setDriftPlanSummary(null);
      setValidationResult(null);
      setPersistedValidationScore(restored.validationScore);
      setValidationNeedsRefresh(false);
      setValidationHandoff(null);
      feedbackAfterValidationRef.current = false;
      setDeploymentGuide(null);

      // Restore architecture prompt if present
      setArchitecturePrompt(restored.architecturePrompt);
      setOriginalPrompt(restored.originalPrompt);
      setReferenceImageUrl(null);
      setLastReferenceArchitecture(null);
      setLastBlueprintArchitecture(null);
    },
    [reactFlowInstance, setEdges, setNodes],
  );

  const applyFlowObject = useCallback(
    (flow: unknown) => {
      applyPreparedFlowObject(prepareFlowObject(flow));
    },
    [applyPreparedFlowObject, prepareFlowObject],
  );

  const cloudValidationScore = currentValidationScore;
  const hasCustomizedPricingScenarios = useMemo(
    () => JSON.stringify(pricingScenarios) !== JSON.stringify(DEFAULT_PRICING_SCENARIOS),
    [pricingScenarios],
  );

  const cloudDiagramPayload = useMemo(() => ({
    nodes,
    edges,
    architecturePrompt,
    originalPrompt: originalPrompt || architecturePrompt || undefined,
    validationScore: cloudValidationScore,
    titleBlockData,
    workflow,
    pricingScenarios,
    iacBaseline,
  }), [
    nodes,
    edges,
    architecturePrompt,
    originalPrompt,
    cloudValidationScore,
    titleBlockData,
    workflow,
    pricingScenarios,
    iacBaseline,
  ]);
  const cloudDiagramPayloadRef = useRef(cloudDiagramPayload);
  cloudDiagramPayloadRef.current = cloudDiagramPayload;
  // Only scan for sensitive data while a privacy or cloud dialog is open. The
  // scan walks the whole payload, so keeping it out of the hot path avoids
  // recomputing on every drag frame when nothing is looking at the result.
  const privacyScanActive = privacyRequest !== null || isCloudWorkspaceOpen;
  const privacyFindings = useMemo(
    () => (privacyScanActive ? detectSensitiveData(cloudDiagramPayload) : EMPTY_SENSITIVE_FINDINGS),
    [privacyScanActive, cloudDiagramPayload],
  );
  const threatMarkers = useMemo(
    () => (threatOverlayEnabled ? analyzeThreatModel(nodes, edges) : EMPTY_THREAT_MARKERS),
    [threatOverlayEnabled, edges, nodes],
  );
  const updateThreatOverlay = useCallback((enabled: boolean) => {
    setThreatOverlayEnabled(enabled);
    writeLocalStorage(THREAT_OVERLAY_STORAGE_KEY, enabled ? '1' : '0');
  }, []);

  const cloudDraftHasContent = useMemo(() => {
    const architectureName = titleBlockData.architectureName.trim();
    const author = titleBlockData.author.trim();
    const version = titleBlockData.version.trim();
    const date = titleBlockData.date.trim();
    const defaultDates = new Set([
      new Date().toLocaleDateString(localeTag),
      new Date().toISOString().split('T')[0],
    ]);
    const hasCustomizedTitle = (
      (
        architectureName.length > 0
        && !['Untitled Architecture', '無題のアーキテクチャ'].includes(architectureName)
      )
      || !['Azure Architect', 'Azure アーキテクト'].includes(author)
      || version !== '1.0'
      || !defaultDates.has(date)
    );

    return (
      nodes.length > 0
      || edges.length > 0
      || architecturePrompt.trim().length > 0
      || originalPrompt.trim().length > 0
      || workflow.length > 0
      || cloudValidationScore !== undefined
      || hasCustomizedPricingScenarios
      || iacBaseline !== null
      || hasCustomizedTitle
    );
  }, [
    architecturePrompt,
    edges.length,
    iacBaseline,
    localeTag,
    nodes.length,
    originalPrompt,
    titleBlockData,
    cloudValidationScore,
    hasCustomizedPricingScenarios,
    workflow.length,
  ]);

  const cloudSync = useCloudDiagramSync({
    diagramName: titleBlockData.architectureName,
    payload: cloudDiagramPayload,
    enabled: cloudDraftHasContent,
    onLoad: applyFlowObject,
  });
  const activeDiagramLineageId = cloudSync.context?.documentId
    ? `cloud:${cloudSync.context.documentId}`
    : localDiagramLineageId;
  activeDiagramLineageIdRef.current = activeDiagramLineageId;

  const recentWorkSyncState: RecentWorkSyncState = (() => {
    if (!cloudSync.context) return 'local';
    switch (cloudSync.status) {
      case 'saving':
      case 'readonly':
      case 'offline':
      case 'unavailable':
      case 'conflict':
      case 'error':
        return cloudSync.status;
      case 'saved':
        return 'synced';
      default:
        return 'synced';
    }
  })();

  const saveCurrentRecovery = useCallback(async (): Promise<RecentWorkRecord | null> => {
    if (!cloudDraftHasContent) return null;
    const recoveryPayload = JSON.parse(JSON.stringify({
      ...cloudDiagramPayload,
      viewport: reactFlowInstance?.getViewport(),
    })) as RecentWorkRecord['payload'];
    const recoveryRecord: RecentWorkRecord = {
      id: activeDiagramLineageId.slice(0, 320),
      lineageId: activeDiagramLineageId.slice(0, 320),
      sessionId: recentWorkSessionId,
      diagramName: (titleBlockData.architectureName.trim() || 'Untitled Architecture').slice(0, 200),
      updatedAt: Date.now(),
      payload: recoveryPayload,
      syncState: recentWorkSyncState,
      cloudDocumentId: cloudSync.context?.documentId,
      cloudRevision: cloudSync.document?.revision,
    };
    await saveRecentWork(recoveryRecord);
    return recoveryRecord;
  }, [
    activeDiagramLineageId,
    cloudDiagramPayload,
    cloudDraftHasContent,
    cloudSync.context?.documentId,
    cloudSync.document?.revision,
    reactFlowInstance,
    recentWorkSessionId,
    recentWorkSyncState,
    titleBlockData.architectureName,
  ]);

  useEffect(() => {
    if (!cloudDraftHasContent) return;
    const timeout = window.setTimeout(() => {
      void saveCurrentRecovery().catch((error) => {
        console.error('Failed to preserve recent work locally:', error);
      });
    }, 650);
    return () => window.clearTimeout(timeout);
  }, [cloudDraftHasContent, saveCurrentRecovery]);

  useEffect(() => {
    const preserveWhenHidden = () => {
      if (document.visibilityState !== 'hidden') return;
      void saveCurrentRecovery().catch((error) => {
        console.error('Failed to preserve recent work while leaving the page:', error);
      });
    };
    const preserveOnPageHide = () => {
      void saveCurrentRecovery().catch((error) => {
        console.error('Failed to preserve recent work during page hide:', error);
      });
    };
    document.addEventListener('visibilitychange', preserveWhenHidden);
    window.addEventListener('pagehide', preserveOnPageHide);
    return () => {
      document.removeEventListener('visibilitychange', preserveWhenHidden);
      window.removeEventListener('pagehide', preserveOnPageHide);
    };
  }, [saveCurrentRecovery]);

  useEffect(() => {
    if (intentionalLineageTransitionRef.current === activeDiagramLineageId) {
      intentionalLineageTransitionRef.current = null;
    } else {
      aiGenerationRef.current.advance();
    }
    validationGenerationRef.current.advance();
    deploymentGuideGenerationRef.current.advance();
    setIsValidating(false);
    setIsGeneratingGuide(false);
    setIsValidationModalOpen(false);
    setIsDeploymentGuideModalOpen(false);
    resetDiagramHistory(diagramHistoryStateRef.current);
  }, [activeDiagramLineageId, resetDiagramHistory]);

  const requestPrivacyAction = useCallback((
    purpose: PrivacyRequest['purpose'],
    onProceed: () => void,
    onCancel: () => void = () => undefined,
    options?: {
      findings?: SensitiveFinding[];
      canAnonymize?: boolean;
    },
  ) => {
    const findings = options?.findings ?? detectSensitiveData(cloudDiagramPayloadRef.current);
    if (purpose !== 'review' && findings.length === 0) {
      onProceed();
      return;
    }
    setPrivacyRequest({
      purpose,
      findings,
      canAnonymize: options?.canAnonymize ?? true,
      onProceed,
      onCancel,
    });
  }, []);

  const guardSensitiveExport = useCallback((action: () => void | Promise<void>) => {
    requestPrivacyAction('export', () => {
      void action();
    });
  }, [requestPrivacyAction]);

  const confirmBeforeShare = useCallback((
    document: CloudDiagramDocument,
    versions: CloudDiagramVersion[],
  ) => new Promise<boolean>((resolve) => {
    requestPrivacyAction(
      'share',
      () => resolve(true),
      () => resolve(false),
      {
        findings: detectSensitiveDataInValue({
          current: {
            diagramName: document.diagramName,
            payload: document.payload,
            comments: document.comments,
          },
          versions,
        }, 'share'),
        canAnonymize: document.id === cloudSync.document?.id,
      },
    );
  }), [cloudSync.document?.id, requestPrivacyAction]);

  const finishPrivacyRequest = useCallback((proceed: boolean) => {
    const request = privacyRequest;
    setPrivacyRequest(null);
    if (!request) return;
    if (proceed) request.onProceed();
    else request.onCancel();
  }, [privacyRequest]);

  const anonymizeCurrentDiagram = useCallback(() => {
    if (privacyRequest && !privacyRequest.canAnonymize) return;
    const sanitized = anonymizeDiagramPayload(cloudDiagramPayload);
    if (cloudSync.context) {
      cloudSync.reset();
      setLocalDiagramLineageId(createLocalDiagramLineageId('anonymized'));
    }
    applyFlowObject(sanitized);
    if (privacyRequest?.purpose === 'share') {
      setIsCloudWorkspaceOpen(false);
    }
    finishPrivacyRequest(false);
    alert(localize(language, {
      en: 'An anonymized working copy is ready. Review it, then repeat the export or share action.',
      ja: '匿名化した作業コピーを用意しました。内容を確認してから、エクスポートまたは共有操作をもう一度実行してください。',
    }));
  }, [
    applyFlowObject,
    cloudDiagramPayload,
    cloudSync,
    finishPrivacyRequest,
    language,
    privacyRequest,
  ]);

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
      await saveCurrentRecovery();
    } catch (error) {
      console.error('Failed to create the local recovery copy before starting fresh:', error);
    }
    try {
      const savedDocument = preserveAsCopy
        ? await cloudSync.saveAsCopy()
        : await cloudSync.saveNow({ force: true });
      if (!savedDocument && cloudDraftHasContent) {
        throw new Error(localize(language, {
          en: 'The current diagram has not been saved to the cloud.',
          ja: '現在の図面はクラウドに保存されていません。',
        }));
      }
    } catch (error) {
      if (error instanceof CloudDiagramOperationCancelledError) return false;
      const discardUnsavedChanges = window.confirm(localize(language, {
        en: 'The current diagram could not be saved to the cloud. Start a new diagram anyway and discard these unsaved changes?',
        ja: '現在の図面をクラウドに保存できませんでした。未保存の変更を破棄して、新しい図面を開始しますか？',
      }));
      if (!discardUnsavedChanges) return false;
    }
    trackStartFresh();
    cloudSync.reset();
    setLocalDiagramLineageId(createLocalDiagramLineageId());
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
    setPersistedValidationScore(undefined);
    setValidationNeedsRefresh(false);
    setDeploymentGuide(null);
    setIaCBaseline(null);
    setDriftPlanSummary(null);
    setReferenceImageUrl(null);
    setLastReferenceArchitecture(null);
    setLastBlueprintArchitecture(null);
    setPricingScenarios(DEFAULT_PRICING_SCENARIOS.map((scenario) => ({ ...scenario })));
    setTitleBlockData({
      architectureName: translate('Untitled Architecture'),
      author: translate('Azure Architect'),
      date: new Date().toISOString().split('T')[0],
      version: '1.0',
    });
    return true;
  }, [
    cloudSync,
    cloudDraftHasContent,
    language,
    saveCurrentRecovery,
    setEdges,
    setNodes,
    translate,
  ]);

  const confirmRecentWorkReplacement = useCallback(async (targetName: string) => {
    if (!cloudDraftHasContent) return true;
    try {
      await saveCurrentRecovery();
    } catch (error) {
      console.error('Failed to preserve the current diagram before switching work:', error);
      const continueWithoutRecovery = window.confirm(localize(language, {
        en: 'The current diagram could not be saved to local recovery storage. Continue switching work and discard it?',
        ja: '現在の図面をローカル復旧ストレージへ保存できませんでした。破棄して作業を切り替えますか？',
      }));
      if (!continueWithoutRecovery) return false;
    }
    return window.confirm(localize(language, {
      en: `Open "${targetName}"? The current canvas will be replaced, and its latest local recovery copy will remain in Recent work.`,
      ja: `「${targetName}」を開きますか？ 現在のキャンバスは置き換えられ、最新のローカル復旧コピーは「最近の作業」に残ります。`,
    }));
  }, [cloudDraftHasContent, language, saveCurrentRecovery]);

  const resumeRecentLocalWork = useCallback(async (record: RecentWorkRecord) => {
    if (record.lineageId === activeDiagramLineageId) return true;
    if (!await confirmRecentWorkReplacement(record.diagramName)) return false;
    try {
      cloudSync.reset();
      setLocalDiagramLineageId(createLocalDiagramLineageId(
        `recovered-${record.lineageId}`,
      ));
      applyFlowObject(record.payload);
      window.setTimeout(() => {
        void reactFlowInstance?.fitView({
          padding: 0.2,
          duration: 350,
          maxZoom: 1.2,
        });
      }, 100);
      return true;
    } catch (error) {
      console.error('Failed to resume local recent work:', error);
      alert(localize(language, {
        en: 'This local recovery copy could not be opened because its diagram data is invalid.',
        ja: '図面データが無効なため、このローカル復旧コピーを開けませんでした。',
      }));
      return false;
    }
  }, [
    activeDiagramLineageId,
    applyFlowObject,
    cloudSync,
    confirmRecentWorkReplacement,
    language,
    reactFlowInstance,
  ]);

  const openRecentCloudWork = useCallback(async (summary: CloudDiagramSummary) => {
    if (cloudSync.document?.id === summary.id) return true;
    if (!await confirmRecentWorkReplacement(summary.diagramName)) return false;
    try {
      const document = await getCloudDiagram(summary.id);
      cloudSync.openDocument(document, {
        documentId: document.id,
        access: 'owner',
        role: 'owner',
      });
      return true;
    } catch (error) {
      console.error('Failed to open a recent cloud diagram:', error);
      alert(error instanceof Error
        ? error.message
        : localize(language, {
          en: 'The cloud diagram could not be opened.',
          ja: 'クラウド図面を開けませんでした。',
        }));
      return false;
    }
  }, [
    cloudSync,
    confirmRecentWorkReplacement,
    language,
  ]);

  const applyArchitectureTemplate = useCallback(async (template: ArchitectureTemplate) => {
    if (nodes.length > 0 || cloudSync.context) {
      const cleared = await startFreshDiagram();
      if (!cleared) return;
    } else {
      cloudSync.reset();
      setLocalDiagramLineageId(createLocalDiagramLineageId(`template-${template.id}`));
    }

    applyFlowObject(template.diagram);
    setIsTemplateGalleryOpen(false);
    trackGuidedJourney({
      action: 'path-selected',
      step: 'create',
      path: 'template-import',
      source: 'first-start',
      hasDiagram: false,
    });
    window.setTimeout(() => {
      reactFlowInstance?.fitView({ padding: 0.2, duration: 350, maxZoom: 1.2 });
    }, 100);
  }, [
    applyFlowObject,
    cloudSync,
    nodes.length,
    reactFlowInstance,
    startFreshDiagram,
  ]);

  const iacComparison = useMemo(
    () => compareDiagramToBaseline(nodes, iacBaseline),
    [nodes, iacBaseline],
  );
  const bicepStarterTemplate = useMemo(
    () => buildStarterTemplate(nodes, 'bicep', edges),
    [nodes, edges],
  );
  const terraformStarterTemplate = useMemo(
    () => buildStarterTemplate(nodes, 'terraform', edges),
    [nodes, edges],
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
            setLocalDiagramLineageId(createLocalDiagramLineageId(
              `restored-${version.lineageId || version.versionId}`,
            ));
            applyFlowObject({
              nodes: version.nodes,
              edges: version.edges,
              metadata: version.metadata,
              workflow: version.workflow,
              architecturePrompt: version.architecturePrompt,
              originalPrompt: version.originalPrompt,
              validationScore: version.validationScore,
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
        setLocalDiagramLineageId(createLocalDiagramLineageId('restored-legacy'));
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
          await handleAIGenerateRef.current(
            flow,
            typeof flow.metadata?.prompt === 'string' ? flow.metadata.prompt : file.name,
            false,
            false,
            () => {
              const lineageId = createLocalDiagramLineageId('loaded');
              cloudSync.reset();
              setLocalDiagramLineageId(lineageId);
              return lineageId;
            },
            false,
          );
        } else {
          const preparedFlow = prepareFlowObject(flow);
          cloudSync.reset();
          setLocalDiagramLineageId(createLocalDiagramLineageId('loaded'));
          applyPreparedFlowObject(preparedFlow);
        }
      } catch (error) {
        if (error instanceof CloudDiagramOperationCancelledError) return;
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
  }, [applyPreparedFlowObject, cloudSync, language, prepareFlowObject, t]);

  // Restore a version from history
  const restoreVersion = useCallback((version: DiagramVersion, restoreAsCopy: boolean) => {
    try {
      if (restoreAsCopy) {
        cloudSync.reset();
        setLocalDiagramLineageId(createLocalDiagramLineageId(
          `restored-${version.lineageId || version.versionId}`,
        ));
      }
      applyFlowObject({
        nodes: version.nodes,
        edges: version.edges,
        titleBlockData: version.titleBlockData || version.metadata,
        workflow: version.workflow || [],
        pricingScenarios: version.pricingScenarios,
        architecturePrompt: version.architecturePrompt || '',
        originalPrompt: version.originalPrompt || version.architecturePrompt || '',
        validationScore: version.validationScore,
        iacBaseline: version.iacBaseline,
      });
      
      console.log('✅ Version restored successfully');
      trackVersionOperation('restore');
    } catch (error) {
      console.error('Failed to restore version:', error);
      alert(t("Failed to restore version"));
    }
  }, [applyFlowObject, cloudSync, t]);

  const restoreSelectedVersionChanges = useCallback(async (
    version: DiagramVersion,
    selectedKeys: string[],
  ): Promise<boolean> => {
    if (selectedKeys.length === 0) return false;
    try {
      await createSnapshot(
        nodes,
        edges,
        titleBlockData.architectureName,
        {
          lineageId: activeDiagramLineageId,
          architecturePrompt,
          originalPrompt: originalPrompt || architecturePrompt || undefined,
          validationScore: cloudValidationScore,
          notes: localize(language, {
            en: 'Automatic backup before selective version restore',
            ja: 'バージョンの選択復元前の自動バックアップ',
          }),
          titleBlockData,
          workflow,
          pricingScenarios,
          iacBaseline,
        },
      );
      try {
        await cloudSync.saveSnapshot(localize(language, {
          en: 'Automatic backup before selective version restore',
          ja: 'バージョンの選択復元前の自動バックアップ',
        }));
      } catch (cloudError) {
        console.warn('Cloud backup was unavailable before selective restore:', cloudError);
      }

      const restored = applySelectedVersionChanges(
        nodes,
        edges,
        version.nodes,
        version.edges,
        selectedKeys,
      );
      if (restored.appliedKeys.length === 0) {
        alert(localize(language, {
          en: 'None of the selected changes could be applied safely.',
          ja: '選択した変更を安全に適用できませんでした。',
        }));
        return false;
      }

      setNodes(restored.nodes);
      setEdges(applyAutomaticEdgeLabelOffsets(restored.nodes, restored.edges));
      if (validationResult || persistedValidationScore !== undefined) {
        setValidationNeedsRefresh(true);
      }
      setDeploymentGuide(null);
      setDriftPlanSummary(null);
      trackVersionOperation('restore');
      alert(localize(language, {
        en: [
          `Applied ${restored.appliedKeys.length} selected change${restored.appliedKeys.length === 1 ? '' : 's'}.`,
          restored.autoAddedNodeIds.length > 0
            ? `${restored.autoAddedNodeIds.length} required parent group${restored.autoAddedNodeIds.length === 1 ? ' was' : 's were'} included automatically.`
            : '',
          restored.autoRemovedEdgeIds.length > 0
            ? `${restored.autoRemovedEdgeIds.length} dangling connection${restored.autoRemovedEdgeIds.length === 1 ? ' was' : 's were'} removed safely.`
            : '',
          restored.skippedKeys.length > 0
            ? `${restored.skippedKeys.length} incompatible change${restored.skippedKeys.length === 1 ? ' was' : 's were'} skipped.`
            : '',
        ].filter(Boolean).join('\n'),
        ja: [
          `選択した${restored.appliedKeys.length}件の変更を適用しました。`,
          restored.autoAddedNodeIds.length > 0
            ? `必要な親グループ${restored.autoAddedNodeIds.length}件を自動的に追加しました。`
            : '',
          restored.autoRemovedEdgeIds.length > 0
            ? `参照先のない接続${restored.autoRemovedEdgeIds.length}件を安全に削除しました。`
            : '',
          restored.skippedKeys.length > 0
            ? `互換性のない変更${restored.skippedKeys.length}件をスキップしました。`
            : '',
        ].filter(Boolean).join('\n'),
      }));
      return true;
    } catch (error) {
      console.error('Failed to restore selected historical changes:', error);
      alert(localize(language, {
        en: 'The backup or selective restore could not be completed.',
        ja: 'バックアップまたは選択復元を完了できませんでした。',
      }));
      return false;
    }
  }, [
    activeDiagramLineageId,
    architecturePrompt,
    cloudSync,
    cloudValidationScore,
    edges,
    iacBaseline,
    language,
    nodes,
    originalPrompt,
    persistedValidationScore,
    pricingScenarios,
    setEdges,
    setNodes,
    titleBlockData,
    validationResult,
    workflow,
  ]);

  const locateQualityFinding = useCallback((finding: DiagramQualityFinding) => {
    const nodeIds = new Set(finding.nodeIds);
    const edgeIds = new Set(finding.edgeIds);
    for (const edge of edges) {
      if (!edgeIds.has(edge.id)) continue;
      nodeIds.add(edge.source);
      nodeIds.add(edge.target);
    }
    setNodes(currentNodes => currentNodes.map(node => ({
      ...node,
      selected: nodeIds.has(node.id),
    })));
    setEdges(currentEdges => currentEdges.map(edge => ({
      ...edge,
      selected: edgeIds.has(edge.id),
    })));
    setIsQualityDoctorOpen(false);

    const locatedNodes = nodes.filter(node => nodeIds.has(node.id));
    window.requestAnimationFrame(() => {
      void reactFlowInstance?.fitView({
        nodes: locatedNodes,
        padding: 0.35,
        duration: 350,
        maxZoom: 1.35,
      });
      reactFlowWrapper.current?.focus();
    });
  }, [edges, nodes, reactFlowInstance, setEdges, setNodes]);

  const applyQualityDoctorFixes = useCallback(async (findingIds: string[]) => {
    if (findingIds.length === 0) return;
    try {
      const report = analyzeDiagramQuality(nodes, edges);
      const backupNotes = localize(language, {
        en: 'Automatic backup before Diagram Quality Doctor fixes',
        ja: 'ダイアグラム品質診断の修正前の自動バックアップ',
      });

      await createSnapshot(
        nodes,
        edges,
        titleBlockData.architectureName,
        {
          lineageId: activeDiagramLineageId,
          architecturePrompt,
          originalPrompt: originalPrompt || architecturePrompt || undefined,
          validationScore: cloudValidationScore,
          notes: backupNotes,
          titleBlockData,
          workflow,
          pricingScenarios,
          iacBaseline,
        },
      );
      try {
        await cloudSync.saveSnapshot(backupNotes);
      } catch (cloudError) {
        console.warn('Cloud backup was unavailable before quality fixes:', cloudError);
      }

      const fixed = applyDiagramQualityFixes(
        nodes,
        edges,
        report.findings,
        findingIds,
      );
      if (fixed.fixedFindingIds.length === 0) {
        alert(localize(language, {
          en: 'None of the selected improvements could be applied safely.',
          ja: '選択した改善項目を安全に適用できませんでした。',
        }));
        return;
      }

      let nextNodes = fixed.nodes;
      let nextEdges = fixed.edges;
      if (fixed.requiresLayout) {
        const layoutResult = await applyLayoutPreset(nextNodes, nextEdges, {
          preset: layoutPreset,
          spacing: layoutSpacing,
          edgeStyle: layoutEdgeStyle,
          emphasizePrimaryPath: layoutEmphasizePrimaryPath
            && (layoutPreset === 'flow-lr' || layoutPreset === 'flow-tb'),
          selectedNodeId: nextNodes.find(
            node => node.type === 'azureNode' && node.selected,
          )?.id,
          layoutEngine,
        });
        nextNodes = layoutResult.nodes;
        nextEdges = layoutResult.edges;
      }
      nextEdges = applyAutomaticEdgeLabelOffsets(nextNodes, nextEdges);

      setNodes(nextNodes);
      setEdges(nextEdges);
      if (validationResult || persistedValidationScore !== undefined) {
        setValidationNeedsRefresh(true);
      }
      setDeploymentGuide(null);
      setDriftPlanSummary(null);
      setIsQualityDoctorOpen(false);
      window.requestAnimationFrame(() => {
        void reactFlowInstance?.fitView({
          padding: 0.2,
          duration: 350,
          maxZoom: 1.2,
        });
      });
    } catch (error) {
      console.error('Failed to apply Diagram Quality Doctor fixes:', error);
      alert(localize(language, {
        en: 'The backup or selected quality fixes could not be completed.',
        ja: 'バックアップまたは選択した品質修正を完了できませんでした。',
      }));
    }
  }, [
    activeDiagramLineageId,
    architecturePrompt,
    cloudSync,
    cloudValidationScore,
    edges,
    iacBaseline,
    language,
    layoutEdgeStyle,
    layoutEmphasizePrimaryPath,
    layoutEngine,
    layoutPreset,
    layoutSpacing,
    nodes,
    originalPrompt,
    persistedValidationScore,
    pricingScenarios,
    reactFlowInstance,
    setEdges,
    setNodes,
    titleBlockData,
    validationResult,
    workflow,
  ]);

  // Manual snapshot save handler
  const handleSaveSnapshot = useCallback(async (notes: string) => {
    try {
      const snapshotNotes = notes || 'Manual snapshot';
      try {
        await cloudSync.saveSnapshot(snapshotNotes);
      } catch (cloudError) {
        console.warn('Cloud snapshot was unavailable; the local snapshot will still be preserved:', cloudError);
      }
      await createSnapshot(
        nodes,
        edges,
        titleBlockData.architectureName,
        {
          lineageId: activeDiagramLineageId,
          architecturePrompt,
          originalPrompt: originalPrompt || architecturePrompt || undefined,
          validationScore: cloudValidationScore,
          notes: snapshotNotes,
          titleBlockData,
          workflow,
          pricingScenarios,
          iacBaseline,
        }
      );
      console.log('✅ Manual snapshot saved successfully');
      trackVersionOperation('save');
    } catch (error) {
      console.error('Failed to save manual snapshot:', error);
      throw error;
    }
  }, [
    activeDiagramLineageId,
    nodes,
    edges,
    titleBlockData,
    architecturePrompt,
    originalPrompt,
    cloudValidationScore,
    workflow,
    pricingScenarios,
    cloudSync,
    iacBaseline,
  ]);

  const handleAIGenerate = useCallback(async (
    architecture: any,
    prompt: string,
    autoSnapshot: boolean = true,
    preserveExistingLayout: boolean = false,
    beforeApply?: () => string | void,
    reportErrors: boolean = true,
    preserveValidationForRecheck: boolean = false,
  ) => {
    const operationGeneration = aiGenerationRef.current.advance();
    const sourceLineageId = activeDiagramLineageIdRef.current;
    const sourceDiagramRevision = diagramRevisionGenerationRef.current.current();
    const assertIntentCurrent = () => {
      if (
        !aiGenerationRef.current.isCurrent(operationGeneration)
        || activeDiagramLineageIdRef.current !== sourceLineageId
      ) {
        throw new CloudDiagramOperationCancelledError();
      }
    };
    const assertSourceCurrent = () => {
      assertIntentCurrent();
      if (!diagramRevisionGenerationRef.current.isCurrent(sourceDiagramRevision)) {
        throw new CloudDiagramOperationCancelledError();
      }
    };
    try {
      assertSourceCurrent();
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
              lineageId: activeDiagramLineageId,
              architecturePrompt: architecturePrompt || 'Previous version',
              originalPrompt: originalPrompt || architecturePrompt || undefined,
              validationScore: cloudValidationScore,
              notes: 'Auto-saved before AI regeneration',
              titleBlockData,
              workflow,
              pricingScenarios,
              iacBaseline,
            }
          );
          assertSourceCurrent();
          try {
            await cloudSync.saveSnapshot('Auto-saved before AI regeneration');
            assertSourceCurrent();
          } catch (cloudError) {
            if (cloudError instanceof CloudDiagramOperationCancelledError) throw cloudError;
            console.warn('Cloud snapshot was unavailable; the local snapshot was preserved:', cloudError);
          }
          console.log('✅ Snapshot saved successfully!');
        } catch (err) {
          if (err instanceof CloudDiagramOperationCancelledError) throw err;
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
        // Last-resort fuzzy match so AI-invented or misspelled names still
        // resolve to a real icon instead of an endless placeholder spinner.
        if (!mapping) mapping = resolveServiceIconLoose(service.type) || resolveServiceIconLoose(service.name);
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
    assertSourceCurrent();
    
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
      const { layoutArchitecture: elkLayoutArchitecture } = await import('./utils/elkLayoutEngine');
      const result = await elkLayoutArchitecture(
        services,
        connections,
        groups || [],
        { direction: 'LR' }
      );
      assertSourceCurrent();
      positionedServices = result.services;
      positionedGroups = result.groups;
    } else {
      const { layoutArchitecture } = await import('./utils/layoutEngine');
      const result = layoutArchitecture(
        services,
        connections,
        groups || [],
        { direction: 'LR' }
      );
      positionedServices = result.services;
      positionedGroups = result.groups;
    }
    assertSourceCurrent();
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
    const determineEdgeDirection = (
      label: string,
      markerColor: string,
    ): { direction: 'forward' | 'reverse' | 'bidirectional', markerEnd?: any, markerStart?: any, flowMode: 'directional' | 'pulse' } => {
      const direction = classifyEdgeDirection(label);

      if (direction === 'bidirectional') {
        return {
          direction: 'bidirectional',
          markerEnd: { type: MarkerType.ArrowClosed, color: markerColor },
          markerStart: { type: MarkerType.ArrowClosed, color: markerColor },
          flowMode: 'pulse',
        };
      }

      if (direction === 'reverse') {
        return {
          direction: 'reverse',
          markerStart: { type: MarkerType.ArrowClosed, color: markerColor },
          markerEnd: undefined,
          flowMode: 'directional',
        };
      }

      return {
        direction: 'forward',
        markerEnd: { type: MarkerType.ArrowClosed, color: markerColor },
        markerStart: undefined,
        flowMode: 'directional',
      };
    };

    // Create edges from connections
    // Numbering follows the Azure Architecture Center convention: the arrow
    // carries the same number as the workflow step that narrates it, so the
    // diagram and the step list in every export refer to the same thing.
    const workflowStepByEdgeId = mapWorkflowStepsToEdges(
      connections.map((conn: any, index: number) => ({
        id: `edge-${index}`,
        source: conn.from,
        target: conn.to,
      })),
      Array.isArray(workflowSteps) ? workflowSteps : [],
    );
    // The prose travels with the arrow so every exporter can rebuild the
    // numbered workflow list from the same data that drew the badges.
    const workflowProseByStep = new Map<number, string>(
      (Array.isArray(workflowSteps) ? workflowSteps : [])
        .filter((s: any) => Number.isInteger(s?.step) && typeof s?.description === 'string')
        .map((s: any) => [s.step as number, String(s.description).trim()]),
    );
    const generatedEdges: Edge[] = connections.map((conn: any, index: number) => {
      const positions = getConnectionPositions(conn.from, conn.to, conn);
      const presentation = getConnectionPresentation(conn.type);
      const edgeDirection = determineEdgeDirection(conn.label || '', presentation.stroke);
      const edgeStyle = {
        stroke: presentation.stroke,
        ...(presentation.strokeDasharray
          ? { strokeDasharray: presentation.strokeDasharray }
          : {}),
        ...(presentation.opacity !== undefined ? { opacity: presentation.opacity } : {}),
      };
      const baseFlowAnimated = presentation.baseFlowAnimated;

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
        labelStyle: { fontSize: 14, fill: '#334155', fontWeight: 650 },
        labelBgStyle: { fill: 'white', fillOpacity: 0.94, stroke: '#cbd5e1', strokeWidth: 1 },
        style: edgeStyle,
        data: {
          connectionType: presentation.type,
          direction: edgeDirection.direction,
          ...(workflowStepByEdgeId.has(`edge-${index}`)
            ? {
                stepNumber: workflowStepByEdgeId.get(`edge-${index}`),
                stepDescription: workflowProseByStep.get(
                  workflowStepByEdgeId.get(`edge-${index}`) as number,
                ),
              }
            : {}),
          baseFlowAnimated,
          flowAnimated,
          flowMode: edgeDirection.flowMode,
          pathStyle: layoutEdgeStyle,
          onLabelChange: handleEdgeLabelChange,
          onLabelOffsetChange: handleEdgeLabelOffsetChange,
          labelOffsetX: 0,
          labelOffsetY: 0,
          labelOffsetAuto: true,
        },
      };
    });
    const newEdges = applyAutomaticEdgeLabelOffsets(finalNodes, generatedEdges);

    // Add the new nodes and edges
    console.log(`Setting ${finalNodes.length} nodes and ${newEdges.length} edges`);
    assertSourceCurrent();
    const transitionedLineageId = beforeApply?.();
    const appliedLineageId = transitionedLineageId || sourceLineageId;
    if (transitionedLineageId && transitionedLineageId !== sourceLineageId) {
      intentionalLineageTransitionRef.current = transitionedLineageId;
      activeDiagramLineageIdRef.current = transitionedLineageId;
    }
    const pricingRunId = ++aiPricingRunRef.current;
    const validationTransition = resolveValidationFreshness(
      validationResult !== null,
      preserveValidationForRecheck,
    );
    if (!validationTransition.keepResult) {
      setValidationResult(null);
      setPersistedValidationScore(undefined);
    }
    setValidationNeedsRefresh(validationTransition.needsRefresh);
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
      const deploymentNames = getDeploymentNames();
      const modelKey = Object.keys(MODEL_CONFIG).find(
        k => deploymentNames[k as ModelType] === architecture.metrics!.model
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
        if (
          pricingRunId !== aiPricingRunRef.current
          || !aiGenerationRef.current.isCurrent(operationGeneration)
          || activeDiagramLineageIdRef.current !== appliedLineageId
        ) return;
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

    // AI generation runs for tens of seconds and swaps the canvas without moving
    // focus, so the completion is otherwise silent for assistive technology.
    announce(localize(language, {
      en: isRefinement
        ? `Diagram updated. ${services.length} services and ${connections.length} connections.`
        : `Diagram generated. ${services.length} services and ${connections.length} connections.`,
      ja: isRefinement
        ? `図を更新しました。サービス ${services.length} 件、接続 ${connections.length} 件です。`
        : `図を生成しました。サービス ${services.length} 件、接続 ${connections.length} 件です。`,
    }));

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
      setIsFeedbackToastOpen(true);
    }

    // A refinement keeps the user's pan/zoom. Only frame a newly generated
    // diagram, where no prior editorial viewport exists.
    if (!isRefinement) {
      setTimeout(() => {
        if (
          aiGenerationRef.current.isCurrent(operationGeneration)
          && activeDiagramLineageIdRef.current === appliedLineageId
        ) {
          reactFlowInstance?.fitView({ padding: 0.2, maxZoom: 1.2 });
        }
      }, 100);
    }
    } catch (error) {
      if (error instanceof CloudDiagramOperationCancelledError) throw error;
      console.error('Error in handleAIGenerate:', error);
      if (reportErrors) {
        announce(localize(language, {
          en: 'Diagram generation failed.',
          ja: '図の生成に失敗しました。',
        }), 'assertive');
        alert(t("Failed to generate diagram. Check console for details."));
      }
      throw error;
    }
  }, [
    activeDiagramLineageId,
    animateConnections,
    architecturePrompt,
    cloudSync,
    edges,
    handleEdgeLabelChange,
    handleEdgeLabelOffsetChange,
    iacBaseline,
    isFeedbackModalOpen,
    language,
    layoutEdgeStyle,
    layoutEngine,
    nodes,
    originalPrompt,
    pricingScenarios,
    reactFlowInstance,
    setEdges,
    setNodes,
    t,
    titleBlockData,
    cloudValidationScore,
    validationResult,
    workflow,
  ]);
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

  const handleAlign = useCallback((type: BulkAlignmentType) => {
    setNodes(currentNodes => alignSelectedNodes(currentNodes, type));
  }, [setNodes]);

  const handleBulkEdit = useCallback(async (
    request: BulkEditRequest,
  ): Promise<BulkEditResult> => {
    const sourceNodes = latestNodesRef.current;
    const selectedNodes = sourceNodes.filter(node => node.selected);
    const selectedNodeIds = new Set(selectedNodes.map(node => node.id));
    const selectedServices = selectedNodes.filter(node => node.type === 'azureNode');
    if (selectedNodeIds.size < 2) {
      throw new Error(localize(language, {
        en: 'Select at least two items before applying bulk edits.',
        ja: '一括編集を適用する前に2件以上を選択してください。',
      }));
    }
    if (
      request.targetGroupId
      && !sourceNodes.some(node => node.id === request.targetGroupId && node.type === 'groupNode')
    ) {
      throw new Error(localize(language, {
        en: 'The selected target group no longer exists.',
        ja: '選択した移動先グループは存在しません。',
      }));
    }
    if (
      request.quantity !== undefined
      && (
        !Number.isInteger(request.quantity)
        || request.quantity < 1
        || request.quantity > 100_000
      )
    ) {
      throw new Error(localize(language, {
        en: 'Quantity must be a whole number from 1 to 100,000.',
        ja: '数量は1から100,000までの整数で指定してください。',
      }));
    }
    if (
      request.customPrice !== undefined
      && (
        !Number.isFinite(request.customPrice)
        || request.customPrice < 0
        || request.customPrice > 1_000_000_000
      )
    ) {
      throw new Error(localize(language, {
        en: 'Custom monthly price must be between 0 and 1,000,000,000.',
        ja: '独自の月額単価は0から1,000,000,000の範囲で指定してください。',
      }));
    }

    const pricingByNodeId = new Map<string, NodePricingConfig>();
    let pricingFailureCount = 0;
    const shouldEditPricing = request.region !== undefined
      || request.quantity !== undefined
      || request.customPrice !== undefined;
    if (shouldEditPricing) {
      await Promise.all(selectedServices.map(async (node) => {
        const currentPricing = node.data?.pricing as NodePricingConfig | undefined;
        const serviceName = String(node.data?.serviceName || node.data?.label || '');
        const targetRegion = request.region || currentPricing?.region || getActiveRegion();
        try {
          let nextPricing: NodePricingConfig | null | undefined;
          if (currentPricing) {
            if (request.region && !currentPricing.isCustom && request.customPrice === undefined) {
              nextPricing = await updateNodePricing(
                serviceName,
                currentPricing,
                undefined,
                request.quantity,
                request.region,
              );
            } else {
              nextPricing = {
                ...currentPricing,
                region: targetRegion,
                quantity: request.quantity ?? currentPricing.quantity,
                lastUpdated: new Date().toISOString(),
              };
            }
          } else if (request.customPrice !== undefined) {
            nextPricing = createCustomPricingDraft(targetRegion as AzureRegion);
            nextPricing.quantity = request.quantity ?? 1;
          } else if (request.region) {
            nextPricing = await initializeNodePricing(serviceName, request.region);
            if (nextPricing && request.quantity !== undefined) {
              nextPricing = { ...nextPricing, quantity: request.quantity };
            }
          }

          if (nextPricing && request.customPrice !== undefined) {
            nextPricing = setCustomPricing(nextPricing, request.customPrice);
          }
          if (nextPricing) pricingByNodeId.set(node.id, nextPricing);
        } catch (error) {
          pricingFailureCount += 1;
          console.error(`Failed to update pricing for ${serviceName}:`, error);
          if (currentPricing && request.quantity !== undefined) {
            pricingByNodeId.set(node.id, {
              ...currentPricing,
              quantity: request.quantity,
              lastUpdated: new Date().toISOString(),
            });
          }
        }
      }));
    }

    const groupColor = request.groupColorName === undefined
      ? undefined
      : request.groupColorName === null
        ? null
        : BULK_GROUP_COLORS.find(color => color.name === request.groupColorName);
    if (request.groupColorName && !groupColor) {
      throw new Error(localize(language, {
        en: 'The selected group color is invalid.',
        ja: '選択したグループ色は無効です。',
      }));
    }

    setNodes(currentNodes => applyBulkNodeEdits(
      currentNodes,
      selectedNodeIds,
      {
        targetGroupId: request.targetGroupId,
        stylePreset: request.stylePreset,
        groupColor,
        tags: request.tags,
        pricingByNodeId,
      },
    ));
    if (validationResult || persistedValidationScore !== undefined) {
      setValidationNeedsRefresh(true);
    }
    setDeploymentGuide(null);
    setDriftPlanSummary(null);
    return {
      updatedCount: selectedNodeIds.size,
      pricingFailureCount,
    };
  }, [
    language,
    persistedValidationScore,
    setNodes,
    validationResult,
  ]);

  // Premium Feature Handlers
  const handleValidateArchitecture = useCallback(async () => {
    // Group boxes are nodes too, but the validator only reasons about Azure
    // services — running it on a canvas that has none wastes an AI call.
    if (nodes.filter(n => n.type === 'azureNode').length === 0) {
      alert(t("Please create an architecture diagram first."));
      return;
    }

    const requestGeneration = validationGenerationRef.current.advance();
    const diagramGeneration = diagramRevisionGenerationRef.current.current();
    const lineageId = activeDiagramLineageIdRef.current;
    const isCurrentValidation = () => (
      validationGenerationRef.current.isCurrent(requestGeneration)
      && diagramRevisionGenerationRef.current.isCurrent(diagramGeneration)
      && activeDiagramLineageIdRef.current === lineageId
    );
    setValidationHandoff(null);

    // Capture diagram snapshot BEFORE opening the modal overlay
    let diagramImageDataUrl: string | undefined;
    if (reactFlowWrapper.current) {
      try {
        if (!isCurrentValidation()) return;
        diagramImageDataUrl = await captureDiagramAsPng(
          reactFlowWrapper.current,
          createDiagramCaptureOptions(),
        );
        if (!isCurrentValidation()) return;
        console.log('\uD83D\uDCF8 Diagram snapshot captured for validation report');
      } catch (err) {
        console.warn('Could not capture diagram snapshot:', err);
      }
    }
    if (!isCurrentValidation()) return;

    // Now show the modal and start validation
    setValidationResult(null);
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
      if (!isCurrentValidation()) return;

      // Attach diagram snapshot to results
      if (diagramImageDataUrl) {
        result.diagramImageDataUrl = diagramImageDataUrl;
      }
      setValidationResult(result);
      setPersistedValidationScore(result.overallScore);
      setValidationNeedsRefresh(false);
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
      // Collapse panels to maximize diagram view
      setPanelsCollapsedSignal(prev => prev + 1);
    } catch (error: any) {
      if (!isCurrentValidation()) return;
      console.error('Validation error:', error);
      alert(localize(language, {
        en: `Failed to validate architecture: ${error.message}`,
        ja: `アーキテクチャの検証に失敗しました: ${error.message}`,
      }));
      setIsValidationModalOpen(false);
    } finally {
      if (validationGenerationRef.current.isCurrent(requestGeneration)) {
        setIsValidating(false);
        if (!isCurrentValidation()) setIsValidationModalOpen(false);
      }
    }
  }, [nodes, edges, architecturePrompt, titleBlockData.architectureName, createDiagramCaptureOptions, t, language]);

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

    const requestGeneration = deploymentGuideGenerationRef.current.advance();
    const diagramGeneration = diagramRevisionGenerationRef.current.current();
    const lineageId = activeDiagramLineageIdRef.current;
    const isCurrentGuide = () => (
      deploymentGuideGenerationRef.current.isCurrent(requestGeneration)
      && diagramRevisionGenerationRef.current.isCurrent(diagramGeneration)
      && activeDiagramLineageIdRef.current === lineageId
    );
    setDeploymentGuide(null);
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

      const { generateDeploymentGuide } = await import('./services/deploymentGuideGenerator');
      const guide = await generateDeploymentGuide(
        services,
        connections,
        groups,
        architecturePrompt || titleBlockData.architectureName,
        totalMonthlyCost,
        language,
      );
      if (!isCurrentGuide()) return;

      setDeploymentGuide(guide);
      trackDeploymentGuide({
        model: guide.metrics?.model,
        serviceCount: services.length,
        bicepFileCount: guide.bicepTemplates?.length,
        elapsedTimeMs: guide.metrics?.elapsedTimeMs,
      });
    } catch (error: any) {
      if (!isCurrentGuide()) return;
      console.error('Guide generation error:', error);
      alert(t('error.deploymentGuide', { message: error.message }));
      setIsDeploymentGuideModalOpen(false);
    } finally {
      if (deploymentGuideGenerationRef.current.isCurrent(requestGeneration)) {
        setIsGeneratingGuide(false);
        if (!isCurrentGuide()) setIsDeploymentGuideModalOpen(false);
      }
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

  const openWorkflowGenerator = useCallback(() => {
    trackGuidedJourney({
      action: 'step-selected',
      step: 'create',
      source: 'journey-strip',
      hasDiagram: nodes.some(node => node.type === 'azureNode'),
    });
    if (!isChatOpen) toggleChatPanel();
  }, [isChatOpen, nodes, toggleChatPanel]);

  const openWorkflowValidation = useCallback(() => {
    trackGuidedJourney({
      action: 'step-selected',
      step: 'validate',
      source: 'journey-strip',
      hasDiagram: nodes.some(node => node.type === 'azureNode'),
    });
    if (validationResult) {
      setIsValidationModalOpen(true);
      return;
    }
    void handleValidateArchitecture();
  }, [handleValidateArchitecture, nodes, validationResult]);

  const openWorkflowCostReview = useCallback(() => {
    activateRibbonTab('home');
    setPricingPrefs({ showCostBadges: true });
    setIsPricingScenarioModalOpen(true);
  }, [activateRibbonTab, setPricingPrefs]);

  const openWorkflowDeployment = useCallback(() => {
    const hasDiagram = nodes.some(node => node.type === 'azureNode');
    trackGuidedJourney({
      action: 'step-selected',
      step: 'deliver',
      path: deploymentGuide ? 'deployment-guide' : 'export',
      source: 'journey-strip',
      hasDiagram,
    });
    if (!hasDiagram) {
      alert(localize(language, {
        en: 'Create or import a diagram before sharing or building artifacts.',
        ja: '共有またはビルドの前に、図を作成またはインポートしてください。',
      }));
      return;
    }
    if (deploymentGuide) {
      setIsDeploymentGuideModalOpen(true);
      return;
    }
    setIsDeliverChooserOpen(true);
  }, [deploymentGuide, language, nodes]);

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
  const commandPaletteCommands: CommandPaletteAction[] = [
    {
      id: 'open-services',
      label: localize(language, { en: 'Open Azure services', ja: 'Azure サービスを開く' }),
      description: localize(language, {
        en: 'Browse the complete service catalog',
        ja: 'サービス カタログ全体を参照します',
      }),
      keywords: ['icons', 'catalog', 'palette', 'services'],
      group: localize(language, { en: 'Navigate', ja: '移動' }),
      icon: <Boxes size={17} aria-hidden="true" />,
      run: openServicesPanel,
    },
    {
      id: 'open-chat',
      label: localize(language, { en: 'Open Architecture Chat', ja: 'Architecture Chat を開く' }),
      description: localize(language, {
        en: 'Create or refine the diagram conversationally',
        ja: '会話形式で図を作成または改善します',
      }),
      keywords: ['ai', 'assistant', 'prompt', 'conversation'],
      group: localize(language, { en: 'Create', ja: '作成' }),
      icon: <MessagesSquare size={17} aria-hidden="true" />,
      run: toggleChatPanel,
    },
    {
      id: 'configure-byo-ai',
      label: localize(language, {
        en: 'Configure custom AI endpoint',
        ja: 'カスタム AI エンドポイントを設定',
      }),
      description: localize(language, {
        en: 'Use your Azure OpenAI or official OpenAI endpoint and model',
        ja: '独自の Azure OpenAI または公式 OpenAI のエンドポイントとモデルを使用します',
      }),
      keywords: ['ai', 'byo', 'custom', 'endpoint', 'model', 'openai', 'azure'],
      group: localize(language, { en: 'Create', ja: '作成' }),
      icon: <Info size={17} aria-hidden="true" />,
      run: () => setIsBYOAISettingsOpen(true),
    },
    {
      id: 'save-diagram',
      label: localize(language, { en: 'Save diagram file', ja: '図ファイルを保存' }),
      description: localize(language, {
        en: 'Download the editable JSON diagram',
        ja: '編集可能な JSON 図をダウンロードします',
      }),
      keywords: ['download', 'json', 'file'],
      group: localize(language, { en: 'File', ja: 'ファイル' }),
      icon: <Save size={17} aria-hidden="true" />,
      run: () => guardSensitiveExport(saveDiagram),
    },
    {
      id: 'export-png',
      label: localize(language, { en: 'Export PNG', ja: 'PNG を出力' }),
      description: localize(language, {
        en: 'Capture the current architecture as an image',
        ja: '現在のアーキテクチャを画像として保存します',
      }),
      keywords: ['image', 'download', 'capture'],
      group: localize(language, { en: 'Export', ja: '出力' }),
      icon: <Download size={17} aria-hidden="true" />,
      disabled: nodes.length === 0,
      run: () => guardSensitiveExport(exportDiagram),
    },
    {
      id: 'privacy-preflight',
      label: localize(language, { en: 'Review privacy before sharing', ja: '共有前にプライバシーを確認' }),
      description: localize(language, {
        en: 'Detect credentials, internal identifiers, addresses, and resource IDs',
        ja: '資格情報、内部識別子、アドレス、リソース ID を検出します',
      }),
      keywords: ['privacy', 'security', 'sensitive', 'redact', 'anonymize'],
      group: localize(language, { en: 'Review', ja: 'レビュー' }),
      icon: <ShieldCheck size={17} aria-hidden="true" />,
      disabled: nodes.length === 0,
      run: () => requestPrivacyAction('review', () => undefined),
    },
    {
      id: 'toggle-threat-overlay',
      label: localize(language, {
        en: threatOverlayEnabled ? 'Hide threat-model overlay' : 'Show threat-model overlay',
        ja: threatOverlayEnabled ? '脅威モデル オーバーレイを非表示' : '脅威モデル オーバーレイを表示',
      }),
      description: localize(language, {
        en: 'Highlight exposure, data, identity, secrets, and detection controls',
        ja: '公開、データ、ID、シークレット、検出制御を強調します',
      }),
      keywords: ['threat', 'security', 'overlay', 'trust boundary'],
      group: localize(language, { en: 'Review', ja: 'レビュー' }),
      icon: <Shield size={17} aria-hidden="true" />,
      disabled: nodes.length === 0,
      run: () => updateThreatOverlay(!threatOverlayEnabled),
    },
    {
      id: 'fit-view',
      label: localize(language, { en: 'Fit diagram to view', ja: '図全体を表示' }),
      description: localize(language, {
        en: 'Center and scale all diagram elements',
        ja: '図の全要素を中央に収めます',
      }),
      keywords: ['zoom', 'center', 'canvas'],
      group: localize(language, { en: 'Canvas', ja: 'キャンバス' }),
      icon: <Frame size={17} aria-hidden="true" />,
      disabled: nodes.length === 0,
      run: () => {
        void reactFlowInstance?.fitView({
          padding: 0.2,
          duration: 400,
          maxZoom: 1.2,
        });
      },
    },
    {
      id: 'validate-architecture',
      label: localize(language, { en: 'Validate architecture', ja: 'アーキテクチャを検証' }),
      description: localize(language, {
        en: 'Run the Azure Well-Architected review',
        ja: 'Azure Well-Architected レビューを実行します',
      }),
      keywords: ['waf', 'review', 'reliability', 'security'],
      group: localize(language, { en: 'Review', ja: 'レビュー' }),
      icon: <Shield size={17} aria-hidden="true" />,
      disabled: nodes.length === 0 || isValidating,
      run: handleValidateArchitecture,
    },
    {
      id: 'diagram-quality-doctor',
      label: localize(language, {
        en: 'Run Diagram Quality Doctor',
        ja: 'ダイアグラム品質診断を実行',
      }),
      description: localize(language, {
        en: 'Find visual clutter, disconnected services, spacing, label, and contrast issues',
        ja: '視覚的な混雑、未接続サービス、余白、ラベル、コントラストの問題を検出します',
      }),
      keywords: ['quality', 'doctor', 'layout', 'overlap', 'crossing', 'contrast', 'readability'],
      group: localize(language, { en: 'Review', ja: 'レビュー' }),
      icon: <ScanSearch size={17} aria-hidden="true" />,
      disabled: nodes.length === 0,
      run: openQualityDoctor,
    },
    {
      id: 'recent-work',
      label: localize(language, { en: 'Resume recent work', ja: '最近の作業を再開' }),
      description: localize(language, {
        en: 'Continue local drafts, recovered sessions, unsynced work, or cloud diagrams',
        ja: 'ローカル下書き、復旧セッション、未同期作業、クラウド図面を続行します',
      }),
      keywords: ['recent', 'resume', 'recover', 'draft', 'unsynced', 'cloud'],
      group: localize(language, { en: 'Workspace', ja: 'ワークスペース' }),
      icon: <Clock size={17} aria-hidden="true" />,
      run: openRecentWork,
    },
    {
      id: 'cloud-workspace',
      label: localize(language, { en: 'Open cloud workspace', ja: 'クラウド ワークスペースを開く' }),
      description: localize(language, {
        en: 'Manage saved diagrams, versions, and sharing',
        ja: '保存済み図、バージョン、共有を管理します',
      }),
      keywords: ['sync', 'versions', 'share', 'collaboration'],
      group: localize(language, { en: 'Workspace', ja: 'ワークスペース' }),
      icon: <Cloud size={17} aria-hidden="true" />,
      run: openCloudWorkspace,
    },
    {
      id: 'toggle-focus',
      label: localize(language, {
        en: focusMode ? 'Exit focus mode' : 'Enter focus mode',
        ja: focusMode ? '集中モードを終了' : '集中モードを開始',
      }),
      description: localize(language, {
        en: 'Show only the architecture canvas',
        ja: 'アーキテクチャ キャンバスだけを表示します',
      }),
      keywords: ['presentation', 'fullscreen', 'canvas'],
      group: localize(language, { en: 'View', ja: '表示' }),
      icon: <PanelLeftClose size={17} aria-hidden="true" />,
      run: toggleFocusMode,
    },
    {
      id: 'toggle-theme',
      label: localize(language, {
        en: isDarkMode ? 'Use light theme' : 'Use dark theme',
        ja: isDarkMode ? 'ライト テーマに切り替え' : 'ダーク テーマに切り替え',
      }),
      description: localize(language, {
        en: 'Switch the application color theme',
        ja: 'アプリケーションの配色を切り替えます',
      }),
      keywords: ['dark', 'light', 'appearance'],
      group: localize(language, { en: 'View', ja: '表示' }),
      icon: isDarkMode
        ? <Sun size={17} aria-hidden="true" />
        : <Moon size={17} aria-hidden="true" />,
      run: () => setIsDarkMode((current) => !current),
    },
    {
      id: 'start-fresh',
      label: localize(language, { en: 'Start a new diagram', ja: '新しい図を開始' }),
      description: localize(language, {
        en: 'Clear the workspace after preserving cloud changes',
        ja: 'クラウド変更を保持してワークスペースをクリアします',
      }),
      keywords: ['new', 'clear', 'reset'],
      group: localize(language, { en: 'File', ja: 'ファイル' }),
      icon: <Trash2 size={17} aria-hidden="true" />,
      run: async () => {
        await startFreshDiagram();
      },
    },
    {
      id: 'open-help',
      label: localize(language, { en: 'Open help and learning', ja: 'ヘルプと学習を開く' }),
      description: localize(language, {
        en: 'Learn canvas controls and architecture workflows',
        ja: 'キャンバス操作とアーキテクチャ手順を確認します',
      }),
      keywords: ['guide', 'learn', 'shortcuts'],
      group: localize(language, { en: 'Help', ja: 'ヘルプ' }),
      icon: <HelpCircle size={17} aria-hidden="true" />,
      run: () => {
        setFocusMode(false);
        setIsHelpOpen(true);
      },
    },
    {
      id: 'open-about',
      label: localize(language, { en: 'About this application', ja: 'このアプリについて' }),
      description: localize(language, {
        en: 'View version, attribution, license, and repository details',
        ja: 'バージョン、作成者、ライセンス、リポジトリ情報を表示します',
      }),
      keywords: ['about', 'version', 'credits', 'license', 'repository'],
      group: localize(language, { en: 'Help', ja: 'ヘルプ' }),
      icon: <Info size={17} aria-hidden="true" />,
      run: () => setIsAboutOpen(true),
    },
  ];

  const openGeneratorFrom = (source: 'first-start' | 'journey-strip' | 'toolbar') => {
    generatorOpenSourceRef.current = source;
    setGeneratorOpenSignal(value => value + 1);
  };

  const openGuidedChat = (source: 'first-start' | 'journey-strip' | 'toolbar') => {
    trackGuidedJourney({
      action: source === 'journey-strip' ? 'step-selected' : 'path-selected',
      step: nodes.length > 0 ? 'refine' : 'create',
      path: 'guided-chat',
      source,
      hasDiagram: nodes.length > 0,
    });
    setIsChatOpen(true);
  };

  const aiGeneratorHost = (
    <AIArchitectureGenerator
      showTrigger={false}
      openSignal={generatorOpenSignal}
      onOpen={() => {
        const source = generatorOpenSourceRef.current;
        trackGuidedJourney({
          action: 'path-selected',
          step: nodes.length > 0 ? 'refine' : 'create',
          path: 'brief-image',
          source,
          hasDiagram: nodes.length > 0,
        });
        generatorOpenSourceRef.current = 'toolbar';
      }}
      onGenerate={async (arch, prompt, autoSnap, refImageUrl) => {
        await handleAIGenerate(arch, prompt, autoSnap, nodes.length > 0);
        clearSourceModel();
        setReferenceImageUrl(refImageUrl ?? null);
        setLastBlueprintArchitecture(null);
      }}
      onReferenceArchitecture={(ref) => {
        setLastReferenceArchitecture(ref ?? null);
      }}
      onBlueprintArchitecture={(bp) => {
        setLastBlueprintArchitecture(bp ?? null);
      }}
      currentArchitecture={{
        nodes,
        edges,
        architectureName: titleBlockData.architectureName,
      }}
      onContinueInChat={() => {
        trackGuidedJourney({ action: 'post-generation-action', step: 'refine', path: 'guided-chat', source: 'generator-success', hasDiagram: true });
        setIsChatOpen(true);
      }}
      onReview={() => {
        trackGuidedJourney({ action: 'post-generation-action', step: 'refine', path: 'canvas', source: 'generator-success', hasDiagram: true });
        window.setTimeout(() => reactFlowInstance?.fitView({ padding: 0.2, duration: 300 }), 100);
      }}
      onValidate={() => {
        trackGuidedJourney({ action: 'post-generation-action', step: 'validate', source: 'generator-success', hasDiagram: true });
        void handleValidateArchitecture();
      }}
    />
  );


  return (
    <div className={`app${isChatOpen ? ' chat-open' : ''}${focusMode ? ' focus-mode' : ''}${mobileCanvasFirstActive ? ' mobile-canvas-first' : ''}`}>
      <a className="skip-link" href="#main-canvas">
        {localize(language, { en: 'Skip to canvas', ja: 'キャンバスへスキップ' })}
      </a>
      <LiveAnnouncer />
      {/* Referenced by every node's aria-describedby so the keyboard contract is
          discoverable without hunting through the help dialog. */}
      <span id="azd-node-keyboard-help" className="azd-visually-hidden">
        {localize(language, {
          en: 'Arrow keys move the node, Shift with an arrow key moves it further, F2 renames it, C starts or completes a connection to another node, Escape cancels.',
          ja: '矢印キーでノードを移動、Shift+矢印キーで大きく移動、F2 で名前を変更、C で他のノードへの接続を開始または確定、Escape で取り消します。',
        })}
      </span>
      <header className={`app-header${isHeaderCollapsed ? ' header-collapsed' : ''}${isMobileRibbonOpen ? ' mobile-ribbon-open' : ''}`}>
        <div className="header-content">
          <div className="header-brand">
            <div
              className="community-brand"
              role="img"
              aria-label={localize(language, {
                en: 'Independent community project',
                ja: '独立したコミュニティ プロジェクト',
              })}
            >
              <span className="community-brand-symbol" aria-hidden="true">
                <Boxes size={18} />
              </span>
              <span className="community-brand-wordmark" aria-hidden="true">
                {localize(language, { en: 'Community project', ja: 'コミュニティ版' })}
              </span>
            </div>
            <h1>{t("Microsoft Product Architecture Diagram Builder")}</h1>
          </div>
          <ResponsiveRibbonSurface
            isOpen={isMobileRibbonOpen}
            onClose={() => setIsMobileRibbonOpen(false)}
          >
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
                role="group"
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
                      {unpricedCount > 0 && (
                        <span
                          className="cost-indicator-partial"
                          title={localize(language, {
                            en: `${unpricedCount} service(s) have no published price and are excluded from this total.`,
                            ja: `${unpricedCount} 件のサービスは公開価格がないため、この合計に含まれていません。`,
                          })}
                        >
                          {' '}+{unpricedCount}
                        </span>
                      )}
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
                role="group"
                aria-label={localize(language, { en: 'Create and AI tools', ja: '作成とAIツール' })}
              >
                {toolbarSectionHeading('create', localize(language, { en: 'Create & AI', ja: '作成・AI' }))}
                <button onClick={addGroupBox} className="btn btn-secondary" title={t("Add grouping box")}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <rect x="3" y="3" width="18" height="18" rx="2" strokeDasharray="4 4" />
                  </svg>
                  {' '}{t("Add Group")}{' '}</button>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setIsTemplateGalleryOpen(true)}
                  title={localize(language, {
                    en: 'Browse previewable starter architectures',
                    ja: 'プレビュー可能なスターター アーキテクチャを参照',
                  })}
                >
                  <LayoutTemplate size={18} />
                  {localize(language, { en: 'Templates', ja: 'テンプレート' })}
                </button>
                <button
                  type="button"
                  className="btn btn-ai btn-generate-ai"
                  onClick={() => openGeneratorFrom('toolbar')}
                  title={localize(language, {
                    en: 'Generate a diagram from detailed requirements or an uploaded image',
                    ja: '詳細な要件またはアップロードした画像から図を生成',
                  })}
                >
                  <Sparkles size={18} />
                  {localize(language, { en: 'Generate Diagram', ja: '図を生成' })}
                </button>
                <ModelSettingsPopover
                  ref={modelSettingsRef}
                  isOpen={isModelSettingsOpen}
                  onToggle={() => setIsModelSettingsOpen(v => !v)}
                  onOpenBYOSettings={() => setIsBYOAISettingsOpen(true)}
                />
                <button
                  className={`btn btn-secondary${isChatOpen ? ' btn-active' : ''}`}
                  onClick={() => {
                    if (!isChatOpen) {
                      trackGuidedJourney({
                        action: 'path-selected',
                        step: nodes.length > 0 ? 'refine' : 'create',
                        path: 'guided-chat',
                        source: 'toolbar',
                        hasDiagram: nodes.length > 0,
                      });
                    }
                    toggleChatPanel();
                  }}
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
                role="group"
                aria-label={localize(language, { en: 'Import architecture', ja: 'アーキテクチャのインポート' })}
              >
                {toolbarSectionHeading('import', localize(language, { en: 'Import', ja: 'インポート' }))}
                <label className={`btn btn-secondary${isImportingTemplate ? ' btn-parsing' : ''}`} title={t("Import Bicep, Terraform, or ARM template to generate diagram")}>
                  {isImportingTemplate ? <Loader size={18} className="spin-icon" /> : <FileCode size={18} />}
                  {isImportingTemplate ? t("Parsing...") : t("Import Template")}
                  <input
                    ref={templateInputRef}
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
                role="group"
                aria-label={localize(language, { en: 'File and export actions', ja: 'ファイルと出力操作' })}
              >
                {toolbarSectionHeading('file', localize(language, { en: 'File & export', ja: 'ファイル・出力' }))}
                <button onClick={() => guardSensitiveExport(saveDiagram)} className="btn btn-secondary" title={t("Save diagram")}>
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
                        {localize(language, { en: 'Presentation', ja: 'プレゼンテーション' })}
                      </div>
                      <div className="toolbar-dropdown-row">
                        <label className="toolbar-dropdown-label" htmlFor="export-background">
                          {localize(language, { en: 'Export background', ja: '出力背景' })}
                        </label>
                        <select
                          id="export-background"
                          className="toolbar-dropdown-select"
                          value={exportBackground}
                          onChange={(event) => {
                            const next = event.target.value as ExportBackground;
                            setExportBackground(next);
                            try { localStorage.setItem(EXPORT_BACKGROUND_STORAGE_KEY, next); } catch { /* ignore */ }
                          }}
                          aria-label={localize(language, { en: 'Export background', ja: '出力背景' })}
                        >
                          <option value="plain">{localize(language, { en: 'Plain (recommended)', ja: 'プレーン（推奨）' })}</option>
                          <option value="dots">{localize(language, { en: 'Dots', ja: 'ドット' })}</option>
                          <option value="grid">{localize(language, { en: 'Grid', ja: 'グリッド' })}</option>
                        </select>
                      </div>
                      <div className="toolbar-dropdown-hint toolbar-dropdown-hint--muted">
                        {localize(language, {
                          en: 'Affects PNG, SVG, animated SVG, and PowerPoint captures. The editing canvas stays dotted.',
                          ja: 'PNG、SVG、アニメーションSVG、PowerPointの出力に反映されます。編集キャンバスはドット表示のままです。',
                        })}
                      </div>
                      <div className="toolbar-dropdown-separator" />
                      <div className="toolbar-dropdown-heading">
                        {localize(language, { en: 'Images & animation', ja: '画像・アニメーション' })}
                      </div>
                      <button
                        className="toolbar-dropdown-item"
                        role="menuitem"
                        onClick={() => {
                          setIsExportMenuOpen(false);
                          guardSensitiveExport(exportDiagram);
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
                          guardSensitiveExport(async () => {
                            await exportReferenceArchitectureAsPng(lastReferenceArchitecture).catch((err) => {
                              console.error('Editorial PNG export failed:', err);
                              alert(t("Editorial PNG export failed. See console for details."));
                            });
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
                          guardSensitiveExport(async () => {
                            await exportBlueprintArchitectureAsPng(lastBlueprintArchitecture, { legendPosition }).catch((err) => {
                              console.error('Blueprint PNG export failed:', err);
                              alert(t("Blueprint PNG export failed. See console for details."));
                            });
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
                          guardSensitiveExport(exportAsSvg);
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
                          guardSensitiveExport(exportAsAnimatedSvg);
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
                          guardSensitiveExport(exportWorkflowAnimation);
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
                          guardSensitiveExport(exportWorkflowMarkdown);
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
                          guardSensitiveExport(exportAsPptx);
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
                          guardSensitiveExport(exportCustomerDeck);
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
                          guardSensitiveExport(exportAsDrawio);
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
                          guardSensitiveExport(exportAsVsdx);
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
                          guardSensitiveExport(exportAsHtml);
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
                          guardSensitiveExport(exportCostBreakdown);
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
                          guardSensitiveExport(exportCostBreakdownZip);
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
                role="group"
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
                role="group"
                aria-label={localize(language, { en: 'History and snapshots', ja: '履歴とスナップショット' })}
              >
                {toolbarSectionHeading('history', localize(language, { en: 'History', ja: '履歴' }))}
                <button
                  onClick={openCloudWorkspace}
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
                role="group"
                aria-label={localize(language, { en: 'Arrange, select, and style', ja: '配置、選択、スタイル' })}
              >
                {toolbarSectionHeading('arrange', localize(language, { en: 'Arrange', ja: '配置・選択' }))}
                <div
                  className="diagram-history-controls"
                  role="group"
                  aria-label={localize(language, { en: 'Diagram history', ja: '図の操作履歴' })}
                >
                  <button
                    type="button"
                    className="btn btn-secondary"
                    onClick={undoDiagram}
                    disabled={!canUndoDiagram}
                    title={localize(language, { en: 'Undo (Ctrl+Z)', ja: '元に戻す (Ctrl+Z)' })}
                    aria-label={localize(language, { en: 'Undo (Ctrl+Z)', ja: '元に戻す (Ctrl+Z)' })}
                  >
                    <Undo2 size={18} />
                    {localize(language, { en: 'Undo', ja: '元に戻す' })}
                  </button>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    onClick={redoDiagram}
                    disabled={!canRedoDiagram}
                    title={localize(language, { en: 'Redo (Ctrl+Y)', ja: 'やり直す (Ctrl+Y)' })}
                    aria-label={localize(language, { en: 'Redo (Ctrl+Y)', ja: 'やり直す (Ctrl+Y)' })}
                  >
                    <Redo2 size={18} />
                    {localize(language, { en: 'Redo', ja: 'やり直す' })}
                  </button>
                </div>
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
                  onClick={toggleFocusMode}
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
                role="group"
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
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => requestPrivacyAction('review', () => undefined)}
                  disabled={nodes.length === 0}
                  title={localize(language, {
                    en: 'Detect sensitive values before export or sharing',
                    ja: 'エクスポートまたは共有前に機密情報を検出',
                  })}
                >
                  <ShieldCheck size={18} />
                  {localize(language, { en: 'Privacy', ja: 'プライバシー' })}
                  {privacyFindings.length > 0 && (
                    <span className="toolbar-count-badge">{privacyFindings.length}</span>
                  )}
                </button>
                <button
                  type="button"
                  className={`btn btn-secondary${threatOverlayEnabled ? ' btn-active' : ''}`}
                  onClick={() => updateThreatOverlay(!threatOverlayEnabled)}
                  disabled={nodes.length === 0}
                  aria-pressed={threatOverlayEnabled}
                  title={localize(language, {
                    en: 'Highlight internet exposure, sensitive data, identity, secrets, and detection controls',
                    ja: 'インターネット公開、機密データ、ID、シークレット、検出制御を強調',
                  })}
                >
                  <Eye size={18} />
                  {localize(language, { en: 'Threats', ja: '脅威表示' })}
                </button>
                {validationResult && (
                  <button
                    onClick={() => setIsValidationModalOpen(true)}
                    className="btn btn-secondary"
                    title={validationNeedsRefresh
                      ? localize(language, {
                          en: 'Architecture changed after recommendations. Open the previous results and revalidate.',
                          ja: '推奨事項の適用後にアーキテクチャが変更されました。以前の結果を開いて再検証してください。',
                        })
                      : t("Open last validation results")}
                  >
                    {validationNeedsRefresh ? <RefreshCw size={18} /> : <Shield size={18} />}
                    {validationNeedsRefresh
                      ? localize(language, { en: 'Revalidate Needed', ja: '再検証が必要' })
                      : (<>{' '}{t("Validation:")}{' '}{translate(bandLabel(validationResult.overallScore))}</>)}
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
          </ResponsiveRibbonSurface>
          <div className="header-identity-actions">
            <DocumentStatus
              status={cloudSync.status}
              lastSavedAt={cloudSync.lastSavedAt}
              hasCloudDocument={Boolean(cloudSync.context)}
              onOpen={() => setIsCloudWorkspaceOpen(true)}
            />
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
            <HeaderUtilityMenu
              onOpenRecentWork={openRecentWork}
              onOpenQualityDoctor={openQualityDoctor}
              onOpenAbout={() => setIsAboutOpen(true)}
            />
          </div>
          <MobileCommandBar
            activeSection={ribbonTabs.find((tab) => tab.id === activeRibbonTab)?.label || ''}
            commandSheetOpen={isMobileRibbonOpen}
            focusMode={focusMode}
            chatOpen={isChatOpen}
            onOpenCommands={() => setIsMobileRibbonOpen((current) => !current)}
            onOpenCommandPalette={openCommandPalette}
            onOpenServices={openServicesPanel}
            onToggleChat={toggleChatPanel}
            onToggleFocus={toggleFocusMode}
          />
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

      {aiGeneratorHost}

      {!focusMode && (
        <WorkflowStepper
          serviceCount={nodes.filter((node) => node.type === 'azureNode').length}
          validationScore={currentValidationScore ?? null}
          hasCostData={hasCostReportData}
          monthlyCostLabel={
            totalMonthlyCost === 0 ? '$0.00/mo' : formatMonthlyCost(totalMonthlyCost)
          }
          hasDeploymentGuide={deploymentGuide !== null}
          isValidating={isValidating}
          isGeneratingGuide={isGeneratingGuide}
          collapsed={mobileCanvasFirstActive && !isMobileJourneyExpanded}
          onToggleCollapsed={mobileCanvasFirstActive
            ? () => setIsMobileJourneyExpanded(current => !current)
            : undefined}
          onGenerate={openWorkflowGenerator}
          onValidate={openWorkflowValidation}
          onReviewCost={openWorkflowCostReview}
          onDeploy={openWorkflowDeployment}
        />
      )}
      
      <main id="main-canvas" className="workspace" tabIndex={-1}>
        <IconPalette
          forceCollapsed={panelsCollapsedSignal > 0 ? panelsCollapsedSignal : undefined}
          openSignal={paletteOpenSignal > 0 ? paletteOpenSignal : undefined}
          onAddIcon={handleAddService}
        />
        
        <div
          className="canvas-container"
          ref={reactFlowWrapper}
          role="region"
          aria-label={localize(language, {
            en: 'Architecture canvas',
            ja: 'アーキテクチャ キャンバス',
          })}
          data-modal-focus-fallback
          tabIndex={0}
          onKeyDown={handleCanvasKeyDown}
          onContextMenuCapture={handleCanvasContextMenuCapture}
        >
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
            onPaneContextMenu={onPaneContextMenu}
            onPaneClick={dismissCanvasContextMenus}
            onInit={setReactFlowInstance}
            onDrop={onDrop}
            onDragOver={onDragOver}
            nodeTypes={nodeTypes}
            edgeTypes={edgeTypes}
            nodesFocusable={false}
            deleteKeyCode={null}
            fitView
            fitViewOptions={{ padding: 0.2, maxZoom: 1.2 }}
            snapToGrid={true}
            snapGrid={[20, 20]}
            selectionOnDrag={true}
            panOnDrag={[1, 2]}
            elevateNodesOnSelect={false}
            reconnectRadius={20}
            attributionPosition="bottom-left"
          >
            {nodes.length > 0 && (
              <Controls fitViewOptions={{ padding: 0.2, maxZoom: 1.2 }} />
            )}
            <Background 
              variant={BackgroundVariant.Dots} 
              gap={20} 
              size={1.5}
              color={isDarkMode ? '#334155' : '#cbd5e1'}
              style={{ backgroundColor: isDarkMode ? '#0f172a' : '#f8fafc' }}
            />
            <CanvasChrome
              hasNodes={nodes.length > 0}
              focusMode={focusMode}
              showNavigationHint={showCanvasHint}
              showEmptyState={!isChatOpen}
              showFeedback={!isChatOpen}
              titleBlockData={titleBlockData}
              generatedWithModel={focusMode ? null : generatedWithModel}
              forceCollapsed={panelsCollapsedSignal > 0 ? panelsCollapsedSignal : undefined}
              feedbackPulse={feedbackFabPulse}
              onDismissNavigationHint={() => {
                setShowCanvasHint(false);
                writeLocalStorage(CANVAS_HINT_STORAGE_KEY, '1');
              }}
              onFitView={() => reactFlowInstance?.fitView?.({
                padding: 0.2,
                duration: 400,
                maxZoom: 1.2,
              })}
              onGuidedChat={() => openGuidedChat('first-start')}
              onGenerateDiagram={() => openGeneratorFrom('first-start')}
              onBrowseTemplates={() => setIsTemplateGalleryOpen(true)}
              onImportTemplate={() => {
                trackGuidedJourney({ action: 'path-selected', step: 'create', path: 'template-import', source: 'first-start', hasDiagram: false });
                templateInputRef.current?.click();
              }}
              onImportAzure={() => {
                trackGuidedJourney({ action: 'path-selected', step: 'create', path: 'azure-import', source: 'first-start', hasDiagram: false });
                setIsAzureImportOpen(true);
              }}
              onTitleBlockUpdate={(data) => {
                setTitleBlockData((current) => ({ ...current, ...data }));
              }}
              onFeedback={() => setIsFeedbackModalOpen(true)}
              onExitFocus={exitFocusMode}
            />
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
            <CanvasActivityOverlay
              isApplyingRecommendations={isApplyingRecommendations}
              isImportingTemplate={isImportingTemplate}
              importFormatLabel={importFormatLabel}
              architecturePrompt={architecturePrompt}
              showArchitecturePrompt={!focusMode}
            />

            {referenceImageUrl && (
              <ReferenceImageViewer
                imageUrl={referenceImageUrl}
                onDismiss={() => setReferenceImageUrl(null)}
              />
            )}
          </ReactFlow>
          <ThreatModelOverlay
            enabled={threatOverlayEnabled}
            markers={threatMarkers}
            onClose={() => updateThreatOverlay(false)}
          />
          <AlignmentToolbar 
            selectedNodes={alignmentSelectedNodes}
            groups={alignmentGroups}
            onAlign={handleAlign}
            onApplyBulkEdit={handleBulkEdit}
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
              onKeyDown={(event) => handleContextMenuNavigation(event, closeNodeContextMenu)}
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
                    className={`context-menu-item${contextMenuContainedCount === 0 ? ' danger' : ''}`}
                    role="menuitem"
                    onClick={() => deleteContextNode(contextMenuNode.id)}
                  >
                    <span className="menu-icon"><Ungroup size={16} /></span>
                    <span className="context-menu-item-copy">
                      <span>
                        {localize(language, {
                          en: contextMenuContainedCount > 0 ? 'Ungroup layer' : 'Delete empty layer',
                          ja: contextMenuContainedCount > 0 ? 'レイヤーをグループ解除' : '空のレイヤーを削除',
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
                  {contextMenuNode.parentNode && (
                    <button
                      className="context-menu-item"
                      role="menuitem"
                      onClick={() => detachContextNodeFromGroup(contextMenuNode.id)}
                    >
                      <span className="menu-icon"><Ungroup size={16} /></span>
                      <span>{localize(language, { en: 'Remove from layer', ja: 'レイヤーから外す' })}</span>
                    </button>
                  )}
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
              onClick={dismissEdgeContextMenu}
              onContextMenu={(event) => {
                event.preventDefault();
                dismissEdgeContextMenu();
              }}
            />
            <div
              ref={edgeContextMenuRef}
              className="edge-context-menu"
              role="menu"
              aria-label={localize(language, {
                en: 'Connection actions',
                ja: '接続線の操作',
              })}
              onKeyDown={(event) => handleContextMenuNavigation(event, closeEdgeContextMenu)}
              style={{
                position: 'fixed',
                top: edgeContextMenu.y,
                left: edgeContextMenu.x,
                zIndex: 10000,
              }}
            >
              <div className="context-menu-header">{t("Edge Direction")}</div>
              <button
                type="button"
                className="context-menu-item"
                role="menuitem"
                onClick={() => setEdgeDirection(edgeContextMenu.edgeId, 'forward')}
              >
                <span className="menu-icon">{t("→")}</span>
                <span>{t("One-way (Forward)")}</span>
              </button>
              <button
                type="button"
                className="context-menu-item"
                role="menuitem"
                onClick={() => setEdgeDirection(edgeContextMenu.edgeId, 'reverse')}
              >
                <span className="menu-icon">{t("←")}</span>
                <span>{t("One-way (Reverse)")}</span>
              </button>
              <button
                type="button"
                className="context-menu-item"
                role="menuitem"
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
                type="button"
                className="context-menu-item"
                role="menuitemcheckbox"
                onClick={() => toggleEdgeAnimation(edgeContextMenu.edgeId)}
                aria-checked={Boolean(contextMenuEdge?.data?.flowAnimated)}
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
              <div className="context-menu-separator" role="separator" />
              <button
                type="button"
                className="context-menu-item danger"
                role="menuitem"
                onClick={() => deleteContextEdge(edgeContextMenu.edgeId)}
              >
                <span className="menu-icon"><Trash2 size={16} /></span>
                <span>
                  {localize(language, {
                    en: 'Delete connection',
                    ja: '接続線を削除',
                  })}
                </span>
              </button>
            </div>
          </>
        )}

        {/* Canvas Context Menu */}
        {paneContextMenu && (
          <>
            <div
              className="edge-context-menu-overlay"
              onClick={dismissPaneContextMenu}
              onContextMenu={(event) => {
                event.preventDefault();
                dismissPaneContextMenu();
              }}
            />
            <div
              ref={paneContextMenuRef}
              className="edge-context-menu pane-context-menu"
              role="menu"
              aria-label={localize(language, {
                en: 'Canvas actions',
                ja: 'キャンバスの操作',
              })}
              onKeyDown={(event) => handleContextMenuNavigation(event, closePaneContextMenu)}
              style={{
                position: 'fixed',
                top: paneContextMenu.y,
                left: paneContextMenu.x,
                zIndex: 10000,
              }}
            >
              <div className="context-menu-header">
                {localize(language, { en: 'Canvas', ja: 'キャンバス' })}
              </div>
              <button
                type="button"
                className="context-menu-item"
                role="menuitem"
                onClick={addContextGroup}
              >
                <span className="menu-icon"><Boxes size={16} /></span>
                <span>{localize(language, { en: 'Add layer here', ja: 'ここにレイヤーを追加' })}</span>
              </button>
              <button
                type="button"
                className="context-menu-item"
                role="menuitem"
                disabled={nodes.length === 0}
                onClick={fitContextDiagram}
              >
                <span className="menu-icon"><Frame size={16} /></span>
                <span>{localize(language, { en: 'Fit diagram to view', ja: '図全体を表示' })}</span>
              </button>
              <button
                type="button"
                className="context-menu-item"
                role="menuitem"
                disabled={nodes.length === 0 && edges.length === 0}
                onClick={selectAllContextItems}
              >
                <span className="menu-icon"><CheckSquare size={16} /></span>
                <span>{localize(language, { en: 'Select all items', ja: 'すべての項目を選択' })}</span>
              </button>
            </div>
          </>
        )}
      </main>

      <MobileNodeInspector
        node={nodes.find(node => node.selected) ?? null}
        onUpdate={(nodeId, patch) => {
          setNodes(current => current.map(node => (
            node.id === nodeId
              ? { ...node, data: { ...node.data, ...patch } }
              : node
          )));
        }}
        onDelete={(nodeId) => deleteCanvasNodes([nodeId])}
        onOpenPricing={(nodeId) => {
          void editContextNodePricing(nodeId);
        }}
      />

      {/* Premium Feature Modals */}
      <TemplateGallery
        isOpen={isTemplateGalleryOpen}
        onClose={() => setIsTemplateGalleryOpen(false)}
        onApply={(template) => {
          void applyArchitectureTemplate(template);
        }}
      />
      <ValidationModal
        validation={validationResult}
        isOpen={isValidationModalOpen}
        onClose={() => setIsValidationModalOpen(false)}
        isLoading={isValidating}
        isStale={validationNeedsRefresh}
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
              await handleAIGenerate(
                improvedArchitecture,
                bannerText,
                true,
                true,
                undefined,
                true,
                true,
              );
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
        onRestoreSelection={restoreSelectedVersionChanges}
        currentLineageId={activeDiagramLineageId}
        currentNodes={nodes}
        currentEdges={edges}
      />
      <CloudWorkspaceModal
        isOpen={isCloudWorkspaceOpen}
        onClose={() => setIsCloudWorkspaceOpen(false)}
        isCloseBlocked={privacyRequest !== null}
        currentDocument={cloudSync.document}
        currentContext={cloudSync.context}
        syncStatus={cloudSync.status}
        syncError={cloudSync.errorMessage}
        lastSavedAt={cloudSync.lastSavedAt}
        hasLocalDraft={cloudDraftHasContent}
        onOpenDocument={cloudSync.openDocument}
        onRestoreVersion={cloudSync.restoreVersion}
        onDocumentUpdated={cloudSync.replaceCurrentDocument}
        onResetCurrent={cloudSync.reset}
        onSaveCurrent={cloudSync.saveNow}
        onReloadRemote={cloudSync.reloadRemote}
        onSaveAsCopy={cloudSync.saveAsCopy}
        onSaveAsDetachedCopy={cloudSync.saveAsDetachedCopy}
        onCloudConflict={cloudSync.reportConflict}
        onDiscardPendingSave={cloudSync.discardPendingSave}
        onCreateNew={startFreshDiagram}
        onBeforeShare={confirmBeforeShare}
        currentUserEmail={accessIdentity?.email || undefined}
        onLocateReviewAnchor={locateReviewAnchor}
      />
      <RecentWorkModal
        isOpen={isRecentWorkOpen}
        currentSessionId={recentWorkSessionId}
        currentLineageId={activeDiagramLineageId}
        onClose={() => setIsRecentWorkOpen(false)}
        onResumeLocal={resumeRecentLocalWork}
        onOpenCloud={openRecentCloudWork}
      />
      <DiagramQualityDialog
        isOpen={isQualityDoctorOpen}
        nodes={nodes}
        edges={edges}
        language={language}
        onClose={() => setIsQualityDoctorOpen(false)}
        onLocateFinding={locateQualityFinding}
        onApplyFixes={applyQualityDoctorFixes}
      />
      <PrivacyPreflightDialog
        isOpen={privacyRequest !== null}
        purpose={privacyRequest?.purpose ?? 'review'}
        findings={privacyRequest?.findings ?? privacyFindings}
        canAnonymize={privacyRequest?.canAnonymize ?? true}
        threatOverlayEnabled={threatOverlayEnabled}
        onThreatOverlayChange={updateThreatOverlay}
        onCancel={() => finishPrivacyRequest(false)}
        onProceed={() => finishPrivacyRequest(true)}
        onAnonymize={anonymizeCurrentDiagram}
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
              // Give icons and layout a moment to settle before cloning the viewport.
              await new Promise(res => setTimeout(res, 1500));
              if (!reactFlowWrapper.current) continue;
              const dataUrl = await captureDiagramAsPng(
                reactFlowWrapper.current,
                createDiagramCaptureOptions(),
              );
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
          setPersistedValidationScore(validation.overallScore);
          setValidationNeedsRefresh(false);
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

      <CommandPalette
        isOpen={isCommandPaletteOpen}
        onClose={() => setIsCommandPaletteOpen(false)}
        commands={commandPaletteCommands}
        onAddService={handleAddService}
      />

      <ArchitectureChatPanel
        isOpen={isChatOpen}
        onClose={() => setIsChatOpen(false)}
        diagramKey={activeDiagramLineageId}
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
      <AboutDialog
        isOpen={isAboutOpen}
        onClose={() => setIsAboutOpen(false)}
      />
      <BYOAISettingsDialog
        isOpen={isBYOAISettingsOpen}
        onClose={() => setIsBYOAISettingsOpen(false)}
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
      {isDeliverChooserOpen && (
        <DeliverChooser
          isBuilding={isGeneratingGuide}
          onClose={() => setIsDeliverChooserOpen(false)}
          onShare={() => {
            trackGuidedJourney({ action: 'path-selected', step: 'deliver', path: 'export', source: 'journey-strip', hasDiagram: true });
            setIsDeliverChooserOpen(false);
            activateRibbonTab('home');
            setIsHeaderCollapsed(false);
            writeLocalStorage(HEADER_COLLAPSED_STORAGE_KEY, '0');
            setCollapsedToolbarSections((current) => {
              if (!current.has('file')) return current;
              const next = new Set(current);
              next.delete('file');
              return next;
            });
            if (window.matchMedia(MEDIA_QUERIES.compactOrShortWorkspace).matches) {
              setIsMobileRibbonOpen(true);
            }
            setIsExportMenuOpen(true);
            window.requestAnimationFrame(() => {
              exportMenuRef.current?.querySelector<HTMLButtonElement>('button')?.focus();
            });
          }}
          onBuild={() => {
            trackGuidedJourney({ action: 'path-selected', step: 'deliver', path: 'deployment-guide', source: 'journey-strip', hasDiagram: true });
            setIsDeliverChooserOpen(false);
            void handleGenerateDeploymentGuide();
          }}
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
            returnFocusTarget={
              pricingEditorReturnFocusRef.current?.nodeId === pricingEditorNodeId
                ? pricingEditorReturnFocusRef.current.element
                : null
            }
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
