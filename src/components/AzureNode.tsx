// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { memo, useEffect, useRef, useState } from 'react';
import { Handle, Position, NodeProps, useReactFlow, useStore } from 'reactflow';
import { Zap, Unlink, Layers } from 'lucide-react';
import { loadIcon, loadIconsFromCategory } from '../utils/iconLoader';
import { NodePricingConfig } from '../types/pricing';
import { formatMonthlyCost, getCostColor } from '../utils/pricingHelpers';
import { isCapacityConsumed } from '../data/serviceIconMapping';
import { usePricingDisplayPrefs } from '../stores/pricingDisplayStore';
import { openNodePricingEditor } from '../stores/nodePricingEditorStore';
import './AzureNode.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { useNodeKeyboardInteraction } from '../hooks/useNodeKeyboardInteraction';
import { usePendingConnection } from '../hooks/useKeyboardConnection';
import { detachNodeFromGroup } from '../utils/groupUtils';
import { shallowArrayEqual, shallowEqual } from '../utils/shallowEqual';
import { categoryAccent } from '../utils/canvasPalette';

// Map categories to colors
const getCategoryColor = (category: string): string => categoryAccent(category);

const AzureNode: React.FC<NodeProps> = memo(({ data, selected, id }) => {
  const { t, language } = useLanguage();
  const [iconUrl, setIconUrl] = useState<string>('');
  // When no icon can be resolved at all (neither the node's own iconPath nor a
  // generic icon for its category), we render a static initial glyph instead of
  // an endless loading spinner.
  const [iconFallbackFailed, setIconFallbackFailed] = useState(false);
  const [isEditingLabel, setIsEditingLabel] = useState(false);
  const [label, setLabel] = useState(data.label || 'Azure Service');
  const cancelLabelEditRef = useRef(false);
  const labelRef = useRef<HTMLDivElement>(null);
  const { setNodes } = useReactFlow();
  // Group membership lives on the node itself, not in `data`, and getNode() is
  // a non-reactive snapshot read. Subscribing to the store is what makes the
  // ungroup button appear and disappear as the node joins or leaves a group;
  // reading it through getNode() left the button on screen after an ungroup
  // until some unrelated change happened to re-render the node.
  const parentNode = useStore((state) => state.nodeInternals.get(id)?.parentNode);
  const pendingConnection = usePendingConnection();
  const isConnectSource = pendingConnection?.nodeId === id;

  // Extract pricing data
  const pricing = data.pricing as NodePricingConfig | undefined;
  const hasPricing = !!pricing
    && Number.isFinite(pricing.estimatedCost)
    && pricing.estimatedCost >= 0;
  const totalCost = pricing ? pricing.estimatedCost * pricing.quantity : 0;

  // Fabric workload items consume Capacity Units from the shared Fabric
  // Capacity rather than billing separately — show an "incl. capacity" badge.
  const serviceKey = (data.serviceName as string) || (data.label as string) || '';
  const capacityConsumed = !hasPricing && isCapacityConsumed(serviceKey);
  
  // Extract style preset
  const stylePreset = (data as any).stylePreset || 'detailed';
  const showLabels = true; // Always show labels
  const labelMaxWidth = (
    typeof data.labelMaxWidth === 'number'
    && Number.isFinite(data.labelMaxWidth)
    && data.labelMaxWidth >= 120
  )
    ? Math.min(260, data.labelMaxWidth)
    : undefined;
  const rawTags: unknown[] = Array.isArray(data.tags) ? data.tags : [];
  const tags = rawTags
    .filter((tag: unknown): tag is string => typeof tag === 'string' && tag.trim() !== '')
    .slice(0, 12);
  // Cost badges are suppressed by presentation styling OR by the standalone
  // cost-visibility preference, so a user can drop the figures without
  // restyling the whole diagram.
  const [pricingPrefs] = usePricingDisplayPrefs();
  const showPricing = stylePreset === 'detailed' && pricingPrefs.showCostBadges;

  useEffect(() => {
    let cancelled = false;
    const category = typeof data.category === 'string' ? data.category.trim() : '';

    const resolve = async () => {
      // 1. Preferred: the icon the node was generated/mapped with.
      if (data.iconPath) {
        const url = await loadIcon(data.iconPath as string);
        if (cancelled) return;
        if (url) {
          setIconUrl(url);
          setIconFallbackFailed(false);
          return;
        }
      }

      // 2. Last resort: any generic icon from the node's category folder so an
      //    unresolved/misspelled service name still shows a meaningful glyph
      //    rather than a spinner that never stops.
      if (category) {
        const icons = await loadIconsFromCategory(category);
        if (cancelled) return;
        const generic = icons[0];
        if (generic) {
          const url = await loadIcon(generic.path);
          if (cancelled) return;
          if (url) {
            setIconUrl(url);
            setIconFallbackFailed(false);
            return;
          }
        }
      }

      // 3. Nothing resolved — fall back to a static initial glyph.
      if (!cancelled) {
        setIconUrl('');
        setIconFallbackFailed(true);
      }
    };

    setIconUrl('');
    setIconFallbackFailed(false);
    resolve();

    return () => {
      cancelled = true;
    };
  }, [data.iconPath, data.category]);

  useEffect(() => {
    if (!isEditingLabel) {
      setLabel(data.label || 'Azure Service');
    }
  }, [data.label, isEditingLabel]);

  const handleLabelDoubleClick = () => {
    cancelLabelEditRef.current = false;
    setIsEditingLabel(true);
  };
  const {
    handleFocus: handleNodeFocus,
    handleKeyDown: handleNodeKeyDown,
  } = useNodeKeyboardInteraction(id, handleLabelDoubleClick, data.label);

  const handleLabelChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setLabel(e.target.value);
  };

  const restoreLabelFocus = () => {
    window.requestAnimationFrame(() => labelRef.current?.focus());
  };

  const commitLabel = (restoreFocus = false) => {
    if (cancelLabelEditRef.current) {
      cancelLabelEditRef.current = false;
      return;
    }
    setNodes(nodes => nodes.map(node => (
      node.id === id
        ? { ...node, data: { ...node.data, label } }
        : node
    )));
    setIsEditingLabel(false);
    if (restoreFocus) restoreLabelFocus();
  };

  const handleLabelKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      commitLabel(true);
    } else if (e.key === 'Escape') {
      cancelLabelEditRef.current = true;
      setLabel(data.label || 'Azure Service');
      setIsEditingLabel(false);
      restoreLabelFocus();
    }
  };

  const handleUngroup = (e: React.MouseEvent) => {
    e.stopPropagation();
    setNodes((nodes) => detachNodeFromGroup(nodes, id));
  };

  const categoryColor = getCategoryColor(data.category);
  // First letter of the service/label, used for the static fallback glyph when
  // no icon (primary or category-generic) can be resolved.
  const nodeInitial = (
    (typeof label === 'string' && label.trim())
      || serviceKey
      || 'Azure Service'
  ).trim().charAt(0).toUpperCase() || '?';
  const borderStyle = {
    borderLeft: `4px solid ${categoryColor}`,
    borderTop: '1px solid #d8e1ea',
    borderRight: '1px solid #d8e1ea',
    borderBottom: '1px solid #d8e1ea',
  };

  return (
    <div
      className={`azure-node ${selected ? 'selected' : ''} style-${stylePreset}${isConnectSource ? ' connect-source' : ''}`}
      style={borderStyle}
      onFocus={handleNodeFocus}
      data-connect-source={isConnectSource ? 'true' : undefined}
    >
      {parentNode && selected && (
        <button
          type="button"
          className="ungroup-button"
          onClick={handleUngroup}
          title={t("Remove from group (you can then drag into another group)")}
          aria-label={t("Remove from group (you can then drag into another group)")}
        >
          <Unlink size={14} aria-hidden="true" />
        </button>
      )}
      <Handle
        type="target"
        position={Position.Top}
        id="top"
        className="node-handle"
        isConnectable={true}
      />
      <Handle
        type="source"
        position={Position.Top}
        id="top-source"
        className="node-handle"
        isConnectable={true}
      />
      <Handle
        type="target"
        position={Position.Left}
        id="left"
        className="node-handle"
        isConnectable={true}
      />
      <Handle
        type="source"
        position={Position.Left}
        id="left-source"
        className="node-handle"
        isConnectable={true}
      />
      
      <div className="node-content">
        {hasPricing && showPricing && (
          <button
            type="button"
            className="cost-badge cost-badge--editable"
            onClick={(e) => { e.stopPropagation(); openNodePricingEditor(id); }}
            title={
              pricing.isUsageBased
                ? localize(language, {
                    en: `Usage-based pricing estimate\n~${formatMonthlyCost(totalCost)}\nBased on typical usage patterns\nActual cost varies with consumption\n\nTier: ${pricing.tier}\nRegion: ${pricing.region}\n\nClick to change tier, quantity, or set your own price`,
                    ja: `従量課金の参考見積もり\n約${formatMonthlyCost(totalCost)}\n一般的な使用量に基づく参考値です\n実際の料金は使用量により変動します\n\nTier: ${pricing.tier}\nRegion: ${pricing.region}\n\nクリックしてTier、数量、独自価格を変更`,
                  })
                : localize(language, {
                    en: `Estimated monthly cost\nTier: ${pricing.tier}\nQuantity: ${pricing.quantity}\nRegion: ${pricing.region}\n${pricing.isCustom ? 'Custom pricing' : 'Auto-calculated'}\n\nClick to change tier, quantity, or set your own price`,
                    ja: `月額参考見積もり\nTier: ${pricing.tier}\n数量: ${pricing.quantity}\nRegion: ${pricing.region}\n${pricing.isCustom ? '独自価格' : '自動計算'}\n\nクリックしてTier、数量、独自価格を変更`,
                  })
            }
            style={{
              '--cost-accent': pricing.isUsageBased ? '#2563eb' : getCostColor(totalCost),
            } as React.CSSProperties}
          >
            {pricing.isUsageBased && <Zap size={12} style={{ marginRight: '2px', display: 'inline-block', verticalAlign: 'middle' }} />}
            {pricing.isUsageBased && '~'}{formatMonthlyCost(totalCost)}
          </button>
        )}
        {capacityConsumed && showPricing && (
          <div
            className="cost-badge cost-badge--capacity"
            title={t("Cost included in Fabric Capacity\nThis item consumes Capacity Units (CUs) from the workspace's Fabric Capacity\nrather than billing separately. See the Microsoft Fabric Capacity node for the cost.")}
          >
            <Layers size={12} style={{ marginRight: '2px', display: 'inline-block', verticalAlign: 'middle' }} />
            {' '}{t("incl. capacity")}{' '}</div>
        )}
        {iconUrl ? (
          <img src={iconUrl} alt={label} className={`node-icon ${stylePreset === 'presentation' ? 'node-icon--presentation' : ''}`} />
        ) : iconFallbackFailed ? (
          <div
            className="node-icon-placeholder"
            style={{ borderLeft: `3px solid ${categoryColor}` }}
            role="img"
            aria-label={label}
            title={label}
          >
            <span
              aria-hidden="true"
              style={{
                fontSize: '20px',
                fontWeight: 700,
                lineHeight: 1,
                color: categoryColor,
              }}
            >
              {nodeInitial}
            </span>
          </div>
        ) : (
          <div className="node-icon-placeholder">
            <div className="loading-spinner"></div>
          </div>
        )}
        
        {showLabels && (
          <>
            {isEditingLabel ? (
              <input
                type="text"
                value={label}
                onChange={handleLabelChange}
                onBlur={() => commitLabel()}
                onKeyDown={handleLabelKeyDown}
                autoFocus
                className="node-label-input"
                style={labelMaxWidth ? { maxWidth: `${labelMaxWidth}px` } : undefined}
              />
            ) : (
              <div
                ref={labelRef}
                data-node-keyboard-target
                className="node-label"
                style={labelMaxWidth ? { maxWidth: `${labelMaxWidth}px` } : undefined}
                onDoubleClick={handleLabelDoubleClick}
                onKeyDown={handleNodeKeyDown}
                role="button"
                tabIndex={0}
                aria-keyshortcuts="F2 C Escape ArrowUp ArrowDown ArrowLeft ArrowRight"
                aria-describedby="azd-node-keyboard-help"
                title={t("Double-click to edit")}
              >
                {label}
              </div>
            )}
          </>
        )}
        {tags.length > 0 && (
          <div className="node-tags" role="group" aria-label={tags.join(', ')}>
            {tags.slice(0, 2).map(tag => (
              <span key={tag} title={tag}>{tag}</span>
            ))}
            {tags.length > 2 && <span title={tags.slice(2).join(', ')}>+{tags.length - 2}</span>}
          </div>
        )}
      </div>
      
      <Handle
        type="source"
        position={Position.Right}
        id="right"
        className="node-handle"
        isConnectable={true}
      />
      <Handle
        type="target"
        position={Position.Right}
        id="right-target"
        className="node-handle"
        isConnectable={true}
      />
      <Handle
        type="source"
        position={Position.Bottom}
        id="bottom"
        className="node-handle"
        isConnectable={true}
      />
      <Handle
        type="target"
        position={Position.Bottom}
        id="bottom-target"
        className="node-handle"
        isConnectable={true}
      />
    </div>
  );
}, (prevProps, nextProps) => {
  // Custom comparison for memo - re-render if these props change.
  //
  // This runs for every node on every React Flow store update (drag, pan,
  // selection, zoom), so it has to stay allocation-free — JSON.stringify here
  // meant two serialisations per node per frame on diagrams with hundreds of
  // nodes. shallowEqual is also strictly more correct: it is insensitive to key
  // order, which stringify is not, and it distinguishes an explicitly
  // undefined field from a missing one, which stringify collapses.
  //
  // Group membership is deliberately absent: it lives on the node rather than
  // in `data`, so it is read reactively from the store instead. A
  // `data.parentNode` term here compared undefined to undefined forever and
  // only looked like it was doing something.
  return (
    prevProps.id === nextProps.id &&
    prevProps.selected === nextProps.selected &&
    prevProps.data.label === nextProps.data.label &&
    // Drives the "incl. capacity" badge independently of the label.
    prevProps.data.serviceName === nextProps.data.serviceName &&
    prevProps.data.category === nextProps.data.category &&
    prevProps.data.labelMaxWidth === nextProps.data.labelMaxWidth &&
    shallowArrayEqual(prevProps.data.tags, nextProps.data.tags) &&
    prevProps.data.iconPath === nextProps.data.iconPath &&
    prevProps.data.stylePreset === nextProps.data.stylePreset &&
    shallowEqual(prevProps.data.pricing, nextProps.data.pricing)
  );
});

AzureNode.displayName = 'AzureNode';

export default AzureNode;