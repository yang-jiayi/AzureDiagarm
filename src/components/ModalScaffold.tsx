// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect } from 'react';
import { createPortal } from 'react-dom';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';

let pageLockCount = 0;
let previousRootInert = false;
let previousRootAriaHidden: string | null = null;
let previousBodyOverflow = '';

function lockBackground(): () => void {
  const root = document.getElementById('root');
  if (pageLockCount === 0) {
    previousBodyOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    if (root) {
      previousRootInert = root.inert;
      previousRootAriaHidden = root.getAttribute('aria-hidden');
      root.inert = true;
      root.setAttribute('aria-hidden', 'true');
    }
  }
  pageLockCount += 1;

  return () => {
    pageLockCount = Math.max(0, pageLockCount - 1);
    if (pageLockCount !== 0) return;

    document.body.style.overflow = previousBodyOverflow;
    if (!root) return;
    root.inert = previousRootInert;
    if (previousRootAriaHidden === null) root.removeAttribute('aria-hidden');
    else root.setAttribute('aria-hidden', previousRootAriaHidden);
  };
}

interface ModalScaffoldProps extends Omit<
  React.HTMLAttributes<HTMLDivElement>,
  'children' | 'className' | 'role' | 'tabIndex'
> {
  isOpen: boolean;
  onClose: () => void;
  children: React.ReactNode;
  className?: string;
  overlayClassName?: string;
  ariaLabel?: string;
  ariaLabelledBy?: string;
  closeOnBackdrop?: boolean;
  closeOnEscape?: boolean;
  returnFocusTarget?: HTMLElement | null;
}

export default function ModalScaffold({
  isOpen,
  onClose,
  children,
  className = '',
  overlayClassName = '',
  ariaLabel,
  ariaLabelledBy,
  closeOnBackdrop = true,
  closeOnEscape = true,
  returnFocusTarget = null,
  ...dialogProps
}: ModalScaffoldProps) {
  const dialogRef = useModalFocus<HTMLDivElement>(isOpen, returnFocusTarget);
  useEscapeKey(isOpen && closeOnEscape, onClose);

  useEffect(() => {
    if (!isOpen) return;
    return lockBackground();
  }, [isOpen]);

  if (!isOpen) return null;

  return createPortal(
    <div
      className={`modal-overlay ${overlayClassName}`.trim()}
      onMouseDown={(event) => {
        if (closeOnBackdrop && event.target === event.currentTarget) onClose();
      }}
      data-modal-scaffold-overlay
    >
      <div
        {...dialogProps}
        ref={dialogRef}
        className={`modal-content ${className}`.trim()}
        role="dialog"
        aria-modal="true"
        aria-label={ariaLabel}
        aria-labelledby={ariaLabelledBy}
        tabIndex={-1}
        data-modal-scaffold
      >
        {children}
      </div>
    </div>,
    document.body,
  );
}
