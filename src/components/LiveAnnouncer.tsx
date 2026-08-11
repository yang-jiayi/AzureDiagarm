import { useSyncExternalStore } from 'react';

import {
  getAnnouncement,
  subscribeToAnnouncements,
  type LiveMessage,
} from '../a11y/liveAnnouncer';

const politeSnapshot = () => getAnnouncement('polite');
const assertiveSnapshot = () => getAnnouncement('assertive');

/**
 * The two ARIA live regions every `announce()` call writes into.
 *
 * Mounted once, near the top of the app so the regions exist in the accessibility
 * tree before any message arrives — a live region added to the DOM at the same
 * time as its content is usually not announced.
 *
 * `key={message.id}` forces a fresh text node for repeated identical messages,
 * which assistive technology otherwise treats as "no change" and stays silent on.
 */
export function LiveAnnouncer() {
  const polite: LiveMessage = useSyncExternalStore(
    subscribeToAnnouncements,
    politeSnapshot,
    politeSnapshot,
  );
  const assertive: LiveMessage = useSyncExternalStore(
    subscribeToAnnouncements,
    assertiveSnapshot,
    assertiveSnapshot,
  );

  return (
    <>
      <div
        className="azd-visually-hidden"
        role="status"
        aria-live="polite"
        aria-atomic="true"
        data-testid="live-region-polite"
      >
        <span key={polite.id}>{polite.text}</span>
      </div>
      <div
        className="azd-visually-hidden"
        role="alert"
        aria-live="assertive"
        aria-atomic="true"
        data-testid="live-region-assertive"
      >
        <span key={assertive.id}>{assertive.text}</span>
      </div>
    </>
  );
}

export default LiveAnnouncer;
