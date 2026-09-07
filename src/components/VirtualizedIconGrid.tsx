// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { useMediaQuery } from '../hooks/useMediaQuery';
import { MEDIA_QUERIES } from '../styles/breakpoints';
import type { AzureIcon } from '../utils/iconLoader';

interface VirtualizedIconGridProps {
  icons: AzureIcon[];
  renderIcon: (icon: AzureIcon) => React.ReactNode;
  onVisibleIconsChange?: (icons: AzureIcon[]) => void;
  ariaLabel: string;
  maxHeight?: number;
  fillAvailableHeight?: boolean;
  layout?: 'grid' | 'list';
}

const GRID_ROW_HEIGHT = 128;
const TOUCH_GRID_ROW_HEIGHT = 164;
const LIST_ROW_HEIGHT = 72;
const ROW_GAP = 8;
const GRID_PADDING = 4;
const OVERSCAN_ROWS = 2;

const VirtualizedIconGrid: React.FC<VirtualizedIconGridProps> = ({
  icons,
  renderIcon,
  onVisibleIconsChange,
  ariaLabel,
  maxHeight = 420,
  fillAvailableHeight = false,
  layout = 'grid',
}) => {
  const viewportRef = useRef<HTMLDivElement>(null);
  const hasCoarsePointer = useMediaQuery(MEDIA_QUERIES.coarsePointer);
  const [width, setWidth] = useState(260);
  const [availableHeight, setAvailableHeight] = useState(maxHeight);
  const [scrollTop, setScrollTop] = useState(0);

  useEffect(() => {
    const element = viewportRef.current;
    if (!element) return;
    const parent = element.parentElement;
    const updateSize = () => {
      setWidth(element.clientWidth || 260);
      if (fillAvailableHeight && parent) {
        const top = element.getBoundingClientRect().top
          - parent.getBoundingClientRect().top + parent.scrollTop;
        const bottomPadding = Number.parseFloat(getComputedStyle(parent).paddingBottom) || 0;
        setAvailableHeight(Math.max(0, parent.clientHeight - top - bottomPadding));
      }
    };
    updateSize();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(updateSize);
    observer.observe(element);
    if (fillAvailableHeight && parent) observer.observe(parent);
    return () => observer.disconnect();
  }, [fillAvailableHeight]);

  useEffect(() => {
    setScrollTop(0);
    if (viewportRef.current) viewportRef.current.scrollTop = 0;
  }, [icons, layout, hasCoarsePointer]);

  const rowHeight = layout === 'list'
    ? LIST_ROW_HEIGHT
    : hasCoarsePointer ? TOUCH_GRID_ROW_HEIGHT : GRID_ROW_HEIGHT;
  const minimumColumnWidth = hasCoarsePointer ? 104 : 72;
  const columns = layout === 'list'
    ? 1
    : Math.max(1, Math.min(3, Math.floor(
        (width - GRID_PADDING * 2 + ROW_GAP) / (minimumColumnWidth + ROW_GAP),
      )));
  const rowCount = Math.ceil(icons.length / columns);
  const totalHeight = rowCount * rowHeight;
  const heightLimit = fillAvailableHeight ? Math.max(rowHeight, availableHeight) : maxHeight;
  const viewportHeight = Math.min(Math.max(rowHeight, totalHeight), heightLimit);
  const startRow = Math.max(0, Math.floor(scrollTop / rowHeight) - OVERSCAN_ROWS);
  const visibleRows = Math.ceil(viewportHeight / rowHeight) + (OVERSCAN_ROWS * 2);
  const endRow = Math.min(rowCount, startRow + visibleRows);
  const visibleIcons = useMemo(
    () => icons.slice(startRow * columns, endRow * columns),
    [columns, endRow, icons, startRow],
  );

  useEffect(() => {
    onVisibleIconsChange?.(visibleIcons);
  }, [onVisibleIconsChange, visibleIcons]);

  if (icons.length === 0) return null;

  return (
    <div
      ref={viewportRef}
      className={`virtualized-icons-viewport virtualized-icons-viewport--${layout}`}
      style={{
        height: viewportHeight,
        '--palette-row-height': `${rowHeight}px`,
        '--palette-row-gap': `${ROW_GAP}px`,
        '--palette-grid-padding': `${GRID_PADDING}px`,
        '--palette-action-size': hasCoarsePointer ? '44px' : '24px',
      } as React.CSSProperties}
      onScroll={(event) => setScrollTop(event.currentTarget.scrollTop)}
      role="region"
      aria-label={ariaLabel}
      tabIndex={0}
    >
      <div className="virtualized-icons-spacer" style={{ height: totalHeight }}>
        <div
          className="virtualized-icons-window"
          style={{
            top: startRow * rowHeight,
            gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))`,
          }}
        >
          {visibleIcons.map(renderIcon)}
        </div>
      </div>
    </div>
  );
};

export default React.memo(VirtualizedIconGrid);
