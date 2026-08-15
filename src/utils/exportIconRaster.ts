// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Icon rasterisation for native-shape exports (PowerPoint / Visio).
 *
 * Office formats cannot embed the app's SVG icons directly (Visio needs a
 * bitmap ForeignData part, PowerPoint's image support for SVG is inconsistent
 * across versions), so both exporters embed a PNG rendered from the same SVG
 * the canvas uses. Results are cached per path + size because a diagram often
 * repeats the same service icon.
 *
 * Browser-only: returns null in non-DOM environments or on any failure so the
 * caller can fall back to a plain shape.
 */

export interface RasterizedIcon {
  bytes: Uint8Array;
  dataUrl: string;
  sizePx: number;
}

const cache = new Map<string, Promise<RasterizedIcon | null>>();

function canRasterize(): boolean {
  return typeof document !== 'undefined' && typeof Image !== 'undefined';
}

/**
 * Read an icon's SVG source from whatever URL the build handed back.
 *
 * `fetch` is not usable here and never was. Vite inlines any asset under
 * `build.assetsInlineLimit` (4 KB by default) as a `data:` URL, and 1,012 of
 * this app's 1,114 service icons are under that limit -- so in a production
 * build `loadIcon` returns a data URL for almost every icon. Fetching a `data:`
 * URL is a *connection*, governed by CSP `connect-src`, and this app's policy
 * lists only `'self'` and `blob:`. The browser blocked the request, rasterizing
 * returned null, and the PowerPoint and Visio exports came out with every icon
 * missing -- while the canvas still drew them all, because an `<img>` is
 * governed by `img-src`, which does allow `data:`.
 *
 * That is why it never reproduced in dev or under Playwright: the dev server
 * inlines nothing, so every icon came back as a real URL and fetch was fine.
 *
 * A data URL already contains the source, so decode it here instead.
 */
async function readSvg(url: string): Promise<string | null> {
  if (url.startsWith('data:')) {
    const comma = url.indexOf(',');
    if (comma < 0) return null;
    const meta = url.slice(5, comma);
    const payload = url.slice(comma + 1);
    try {
      if (/;base64/i.test(meta)) {
        const binary = atob(payload);
        const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
        return new TextDecoder('utf-8').decode(bytes);
      }
      return decodeURIComponent(payload);
    } catch {
      return null;
    }
  }
  const response = await fetch(url);
  if (!response.ok) return null;
  return response.text();
}

async function rasterize(iconPath: string, sizePx: number): Promise<RasterizedIcon | null> {
  if (!canRasterize()) return null;
  try {
    // Loaded lazily: iconLoader uses `import.meta.glob`, which only exists in
    // the Vite runtime, and this module is also imported by Node-side tests.
    const { loadIcon } = await import('./iconLoader');
    const url = await loadIcon(iconPath);
    if (!url) return null;

    let svg = await readSvg(url);
    if (!svg) return null;
    // Many Azure icon SVGs declare only a viewBox. Without explicit pixel
    // dimensions some browsers rasterize a zero-sized (blank) image.
    if (!/<svg[^>]*\bwidth\s*=/.test(svg)) {
      svg = svg.replace(/<svg\b/, `<svg width="${sizePx}" height="${sizePx}"`);
    }

    const svgUrl = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    const image = new Image();
    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();
      image.onerror = () => reject(new Error('icon image load failed'));
      image.src = svgUrl;
    });

    const canvas = document.createElement('canvas');
    canvas.width = sizePx;
    canvas.height = sizePx;
    const context = canvas.getContext('2d');
    if (!context) return null;
    context.drawImage(image, 0, 0, sizePx, sizePx);

    const dataUrl = canvas.toDataURL('image/png');
    const blob: Blob | null = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (!blob) return null;
    return { bytes: new Uint8Array(await blob.arrayBuffer()), dataUrl, sizePx };
  } catch {
    return null;
  }
}

/** Rasterize an icon to PNG bytes + data URL, memoised per path and size. */
export function rasterizeIconToPng(
  iconPath: string,
  sizePx = 128,
): Promise<RasterizedIcon | null> {
  if (!iconPath) return Promise.resolve(null);
  const key = `${iconPath}@${sizePx}`;
  const cached = cache.get(key);
  if (cached) return cached;
  const pending = rasterize(iconPath, sizePx).catch(() => null);
  cache.set(key, pending);
  return pending;
}

/** Rasterize many icons in parallel; failures resolve to null. */
export async function rasterizeIcons(
  iconPaths: Array<string | undefined>,
  sizePx = 128,
): Promise<Map<string, RasterizedIcon>> {
  const unique = Array.from(new Set(iconPaths.filter((path): path is string => !!path)));
  const results = await Promise.all(unique.map((path) => rasterizeIconToPng(path, sizePx)));
  const map = new Map<string, RasterizedIcon>();
  unique.forEach((path, index) => {
    const result = results[index];
    if (result) map.set(path, result);
  });
  return map;
}
