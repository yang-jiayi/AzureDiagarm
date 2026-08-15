// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Reading assets the exporters embed, under the policy the site actually ships.
 *
 * Every export that embeds an icon has, at some point, been written as
 * `fetch(url)` -- and every one of them broke in production while passing every
 * test. The reason is worth stating once, here, so the next exporter does not
 * repeat it:
 *
 * Vite inlines any asset under `build.assetsInlineLimit` (4 KB by default) as a
 * `data:` URL. 1,012 of this app's 1,114 service icons are under that limit, so
 * a real build emits 86 `.svg` files and inlines the other 1,017. `loadIcon`
 * therefore hands back a `data:` URL for roughly 91 percent of icons.
 *
 * Fetching a `data:` URL is a *connection*, governed by CSP `connect-src`, and
 * this app's policy (nginx.conf) lists only `'self'` and `blob:`. So the fetch
 * is refused, the icon is dropped, and -- because the failure is usually inside
 * a `try/catch` that falls back to "no icon" -- the export completes and looks
 * fine until someone opens it.
 *
 * The canvas never showed a symptom, because an `<img>` is governed by
 * `img-src`, which does allow `data:`. And the dev server inlines nothing, so
 * no test that runs against it can reproduce any of this.
 *
 * A data URL already carries its own payload. Decode it here rather than asking
 * the network for something we are holding. Widening `connect-src` to `data:`
 * would also make the symptom go away, at the cost of the directive that makes
 * exfiltration hard; it is not the trade to make.
 */

export interface DecodedDataUrl {
  mediaType: string;
  bytes: Uint8Array;
}

/** Split a `data:` URL into its media type and raw bytes. */
export function decodeDataUrl(url: string): DecodedDataUrl | null {
  if (!url.startsWith('data:')) return null;
  const comma = url.indexOf(',');
  if (comma < 0) return null;
  const meta = url.slice(5, comma);
  const payload = url.slice(comma + 1);
  try {
    if (/;base64/i.test(meta)) {
      const binary = atob(payload);
      return {
        mediaType: meta.replace(/;base64.*$/i, '') || 'text/plain',
        bytes: Uint8Array.from(binary, (char) => char.charCodeAt(0)),
      };
    }
    return {
      mediaType: meta.split(';')[0] || 'text/plain',
      bytes: new TextEncoder().encode(decodeURIComponent(payload)),
    };
  } catch {
    return null;
  }
}

/**
 * Turn a `data:` URL into a Blob without going through the network.
 *
 * `URL.createObjectURL` on the result gives a `blob:` URL, which the policy
 * does allow, so downloads keep working the way they always have.
 */
export function dataUrlToBlob(url: string): Blob | null {
  const decoded = decodeDataUrl(url);
  if (!decoded) return null;
  return new Blob([decoded.bytes as unknown as BlobPart], { type: decoded.mediaType });
}

/**
 * Read a text asset, whether the build inlined it or emitted it as a file.
 *
 * Returns null rather than throwing, because every caller's fallback is to
 * carry on without the asset.
 */
export async function readTextAsset(url: string): Promise<string | null> {
  const decoded = decodeDataUrl(url);
  if (decoded) return new TextDecoder('utf-8').decode(decoded.bytes);
  if (url.startsWith('data:')) return null;
  try {
    const response = await fetch(url);
    if (!response.ok) return null;
    return await response.text();
  } catch {
    return null;
  }
}

/**
 * A `data:` URL for an SVG source, in the one form every consumer accepts.
 *
 * Without the `;base64` marker a data URL is percent-decoded rather than
 * base64-decoded, so a base64 payload is read as literal text and the image
 * silently fails to render. The Draw.io export shipped that way.
 */
export function svgToDataUrl(svg: string): string {
  const bytes = new TextEncoder().encode(svg);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return `data:image/svg+xml;base64,${btoa(binary)}`;
}
