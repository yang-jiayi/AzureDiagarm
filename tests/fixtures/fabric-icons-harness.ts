import { loadIcon, loadIconsFromCategory, loadIconsFromPaletteCategory } from '../../src/utils/iconLoader';
import { rasterizeIconToPng } from '../../src/utils/exportIconRaster';
import { readTextAsset } from '../../src/utils/assetSource';

export async function inspectFabricIcons() {
  const icons = await loadIconsFromCategory('fabric');
  let inlined = 0;
  let rasterizedUi = 0;
  for (const icon of icons) {
    const url = await loadIcon(icon.path);
    if (!url) throw new Error(`Missing icon URL: ${icon.name}`);
    if (url.startsWith('data:')) inlined++;
    const image = new Image();
    image.src = url;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = canvas.height = 64;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (!context) throw new Error('Canvas context unavailable');
    context.drawImage(image, 0, 0, 64, 64);
    const pixels = context.getImageData(0, 0, 64, 64).data;
    if (!pixels.some((value, index) => index % 4 === 3 && value > 0)) {
      throw new Error(`Blank Fabric icon: ${icon.name}`);
    }
    if (icon.fabricDrawingKind === 'ui') {
      const raster = await rasterizeIconToPng(icon.path, 64);
      if (!raster || raster.svg !== await readTextAsset(url)) {
        throw new Error(`Native export lost original UI artwork: ${icon.name}`);
      }
      const bitmap = new Image();
      bitmap.src = raster.dataUrl;
      await bitmap.decode();
      context.clearRect(0, 0, 64, 64);
      context.drawImage(bitmap, 0, 0);
      const bitmapPixels = context.getImageData(0, 0, 64, 64).data;
      if (!bitmapPixels.some((value, index) => index % 4 === 3 && value > 0)) {
        throw new Error(`Blank native export bitmap: ${icon.name}`);
      }
      rasterizedUi++;
    }
  }
  return {
    sourceCount: icons.length,
    categoryCount: (await loadIconsFromPaletteCategory('fabric')).length,
    inlined, rasterizedUi,
    uiNames: icons.filter(icon => icon.fabricDrawingKind === 'ui').map(icon => icon.name),
  };
}
