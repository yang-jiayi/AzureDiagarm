// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useRef } from 'react';
import {
  EdgeProps,
  getBezierPath,
  getStraightPath,
  getSmoothStepPath,
  EdgeLabelRenderer,
  BaseEdge,
} from 'reactflow';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './EditableEdge.css';

const EditableEdge: React.FC<EdgeProps> = ({
  id,
  sourceX,
  sourceY,
  targetX,
  targetY,
  sourcePosition,
  targetPosition,
  style = {},
  markerEnd,
  markerStart,
  data,
  label,
  selected,
}) => {
  const { t, language } = useLanguage();
  const [isEditing, setIsEditing] = useState(false);
  const [editLabel, setEditLabel] = useState(label?.toString() || '');
  const [isDragging, setIsDragging] = useState(false);
  const dragStartRef = useRef<{ x: number; y: number; offsetX: number; offsetY: number } | null>(null);

  React.useEffect(() => {
    if (!isEditing) {
      setEditLabel(label?.toString() || '');
    }
  }, [label, isEditing]);

  const pathStyle = (data as any)?.pathStyle as 'straight' | 'smooth' | 'orthogonal' | undefined;

  const pathFn =
    pathStyle === 'straight'
      ? getStraightPath
      : pathStyle === 'orthogonal'
        ? getSmoothStepPath
        : getBezierPath;

  const [edgePath, labelX, labelY] = pathFn({
    sourceX,
    sourceY,
    sourcePosition,
    targetX,
    targetY,
    targetPosition,
  } as any);

  // Get stored offset from edge data
  const offsetX = (data as any)?.labelOffsetX ?? 0;
  const offsetY = (data as any)?.labelOffsetY ?? 0;

  const handleLabelDoubleClick = (e: React.MouseEvent) => {
    e.stopPropagation();
    if (!isDragging) {
      setIsEditing(true);
    }
  };

  const handleLabelChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setEditLabel(e.target.value);
  };

  const handleLabelBlur = () => {
    setIsEditing(false);
    // Update the edge data
    if (data?.onLabelChange) {
      data.onLabelChange(id, editLabel);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleLabelBlur();
    } else if (e.key === 'Escape') {
      setEditLabel(label?.toString() || '');
      setIsEditing(false);
    }
  };

  const handleMouseDown = (e: React.MouseEvent) => {
    if (isEditing) return;
    e.stopPropagation();
    setIsDragging(true);
    dragStartRef.current = {
      x: e.clientX,
      y: e.clientY,
      offsetX,
      offsetY,
    };
  };

  // Add/remove global mouse event listeners
  React.useEffect(() => {
    if (!isDragging) return;
    const handleMouseMove = (event: MouseEvent) => {
      if (!dragStartRef.current) return;
      const dx = event.clientX - dragStartRef.current.x;
      const dy = event.clientY - dragStartRef.current.y;
      data?.onLabelOffsetChange?.(
        id,
        dragStartRef.current.offsetX + dx,
        dragStartRef.current.offsetY + dy,
      );
    };
    const handleMouseUp = () => {
      setIsDragging(false);
      dragStartRef.current = null;
    };
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, [data, id, isDragging]);

  const direction = (data?.direction ?? 'forward') as 'forward' | 'reverse' | 'bidirectional';
  const flowMode = (data?.flowMode ?? (direction === 'bidirectional' ? 'pulse' : 'directional')) as
    | 'directional'
    | 'pulse';

  const flowAnimated = Boolean(data?.flowAnimated);
  const shouldDirectionalFlow = flowAnimated && flowMode === 'directional' && (direction === 'forward' || direction === 'reverse');
  const shouldPulseFlow = flowAnimated && flowMode === 'pulse' && direction === 'bidirectional';
  const edgeStroke = selected ? '#0f6cbd' : ((style as any)?.stroke ?? '#64748b');
  const edgeStrokeWidth = Number((style as any)?.strokeWidth) || 1.75;

  return (
    <>
      <BaseEdge
        path={edgePath}
        markerEnd={markerEnd}
        markerStart={markerStart}
        style={{
          ...style,
          stroke: edgeStroke,
          strokeWidth: selected ? Math.max(2.5, edgeStrokeWidth) : edgeStrokeWidth,
          animation: undefined,
        }}
      />
      {flowAnimated && (
        <path
          d={edgePath}
          fill="none"
          stroke={edgeStroke}
          strokeWidth={Math.max(2, edgeStrokeWidth)}
          strokeLinecap="round"
          pointerEvents="none"
          className={`editable-edge-flow-path${selected ? ' is-selected' : ''}`}
          style={{
            strokeDasharray: shouldPulseFlow ? '2 10' : '3 9',
            animation: shouldPulseFlow
              ? 'edge-pulse 1.4s ease-in-out infinite'
              : shouldDirectionalFlow
                ? direction === 'reverse'
                  ? 'edge-dash-reverse 0.9s linear infinite'
                  : 'edge-dash-forward 0.9s linear infinite'
                : undefined,
          }}
          aria-hidden="true"
        />
      )}
      <EdgeLabelRenderer>
        <div
          style={{
            position: 'absolute',
            transform: `translate(-50%, -50%) translate(${labelX + offsetX}px,${labelY + offsetY}px)`,
            pointerEvents: 'all',
          }}
          data-label-offset-x={offsetX}
          data-label-offset-y={offsetY}
          data-label-offset-auto={(data as any)?.labelOffsetAuto !== false}
          className={`nodrag nopan editable-edge-label-shell${selected ? ' is-selected' : ''}`}
        >
          {isEditing ? (
            <input
              type="text"
              value={editLabel}
              onChange={handleLabelChange}
              onBlur={handleLabelBlur}
              onKeyDown={handleKeyDown}
              autoFocus
              className="editable-edge-label-input"
            />
          ) : (
            <div
              data-edge-label-id={id}
              onMouseDown={handleMouseDown}
              onDoubleClick={handleLabelDoubleClick}
              onKeyDown={(event) => {
                if (event.key === 'Enter' || event.key === ' ' || event.key === 'F2') {
                  event.preventDefault();
                  setIsEditing(true);
                }
              }}
              role="button"
              tabIndex={0}
              className={`editable-edge-label${isDragging ? ' is-dragging' : ''}${editLabel ? '' : ' is-empty'}`}
              title={`${editLabel || t("Double-click to edit label")}\n${localize(language, {
                en: '(Drag to reposition)',
                ja: '（ドラッグして位置を変更）',
              })}`}
            >
              {editLabel || localize(language, {
                en: '(click to add label)',
                ja: '（クリックしてラベルを追加）',
              })}
            </div>
          )}
        </div>
      </EdgeLabelRenderer>
    </>
  );
};

export default React.memo(EditableEdge);
