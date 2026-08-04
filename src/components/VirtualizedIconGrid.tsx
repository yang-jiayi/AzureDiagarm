// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import type { AzureIcon } from '../utils/iconLoader';

interface VirtualizedIconGridProps {
  icons: AzureIcon[];
  renderIcon: (icon: AzureIcon) => React.ReactNode;
  onVisibleIconsChange?: (icons: AzureIcon[]) => void;
  ariaLabel: string;
  maxHeight?: number;
  layout?: 'grid' | 'list';
}

const GRID_ROW_HEIGHT = 112;
const LIST_ROW_HEIGHT = 72;
const OVERSCAN_ROWS = 2;

const VirtualizedIconGrid: React.FC<VirtualizedIconGridProps> = ({
  icons,
  renderIcon,
  onVisibleIconsChange,
  ariaLabel,
  maxHeight = 420,
  layout = 'grid',
}) => {
  const viewportRef = useRef<HTMLDivElement>(null);
  const [width, setWidth] = useState(260);
  const [scrollTop, setScrollTop] = useState(0);

  useEffect(() => {
    const element = viewportRef.current;
    if (!element) return;
    const updateWidth = () => setWidth(element.clientWidth || 260);
    updateWidth();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(updateWidth);
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    setScrollTop(0);
    if (viewportRef.current) viewportRef.current.scrollTop = 0;
  }, [icons, layout]);

  const rowHeight = layout === 'list' ? LIST_ROW_HEIGHT : GRID_ROW_HEIGHT;
  const columns = layout === 'list' ? 1 : width >= 230 ? 3 : 2;
  const rowCount = Math.ceil(icons.length / columns);
  const totalHeight = rowCount * rowHeight;
  const viewportHeight = Math.min(Math.max(rowHeight, totalHeight), maxHeight);
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
      style={{ height: viewportHeight }}
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
