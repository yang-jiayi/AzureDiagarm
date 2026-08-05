// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect } from 'react';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';
import './ResponsiveDrawer.css';

type DrawerPlacement = 'docked' | 'left' | 'right' | 'bottom' | 'center';

interface ResponsiveDrawerProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'role'> {
  isOpen: boolean;
  modal: boolean;
  placement: DrawerPlacement;
  ariaLabel: string;
  onClose: () => void;
  role?: React.AriaRole;
  backdropClassName?: string;
  backgroundSelectors?: string[];
}

export default function ResponsiveDrawer({
  isOpen,
  modal,
  placement,
  ariaLabel,
  onClose,
  role = 'region',
  backdropClassName = '',
  backgroundSelectors = [],
  className = '',
  tabIndex,
  children,
  ...panelProps
}: ResponsiveDrawerProps) {
  const panelRef = useModalFocus<HTMLDivElement>(isOpen && modal);
  useEscapeKey(isOpen && modal, onClose);
  const backgroundSelectorKey = backgroundSelectors.join('\u0000');

  useEffect(() => {
    if (!isOpen || !modal) return;
    const panel = panelRef.current;
    if (!panel) return;
    const selectors = backgroundSelectorKey
      ? backgroundSelectorKey.split('\u0000')
      : [];

    const backgroundElements = [...new Set(
      selectors.flatMap(selector => (
        Array.from(document.querySelectorAll<HTMLElement>(selector))
      )),
    )].filter(element => element !== panel && !element.contains(panel));
    const previousState = backgroundElements.map(element => ({
      element,
      inert: element.inert,
      ariaHidden: element.getAttribute('aria-hidden'),
    }));

    for (const element of backgroundElements) {
      element.inert = true;
      element.setAttribute('aria-hidden', 'true');
    }

    return () => {
      for (const { element, inert, ariaHidden } of previousState) {
        element.inert = inert;
        if (ariaHidden === null) element.removeAttribute('aria-hidden');
        else element.setAttribute('aria-hidden', ariaHidden);
      }
    };
  }, [backgroundSelectorKey, isOpen, modal, panelRef]);

  if (!isOpen) return null;

  return (
    <>
      {modal && (
        <button
          type="button"
          className={`responsive-drawer-backdrop ${backdropClassName}`.trim()}
          onClick={onClose}
          tabIndex={-1}
          aria-hidden="true"
        />
      )}
      <div
        {...panelProps}
        ref={panelRef}
        className={`responsive-drawer-panel responsive-drawer-panel--${placement} ${className}`.trim()}
        role={modal ? 'dialog' : role}
        aria-modal={modal ? 'true' : undefined}
        aria-label={ariaLabel}
        tabIndex={modal ? -1 : tabIndex}
        data-responsive-drawer
        data-modal={modal ? 'true' : 'false'}
        data-placement={placement}
      >
        {children}
      </div>
    </>
  );
}
