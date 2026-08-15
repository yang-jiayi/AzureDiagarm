import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { nativizePackage } from '../../src/services/pptxNativeShapes';

/**
 * The four services the user actually exported, drawn as empty boxes in Visio.
 *
 * The existing icon spec hands the exporter an icon path it wrote itself, which
 * proves rasterisation works but cannot prove the pipeline ever produces one.
 * These four came from the palette, and every one of them lost its icon on the
 * sheet, so the failure is somewhere between "the library has this icon" and
 * "the .vsdx embeds it" -- exactly the span a hard-coded path skips over.
 *
 * So this resolves them the way the palette does, builds the nodes the way a
 * drag onto the canvas does, and then asks the package what it contains.
 */
const SERVICES = ['Applens', 'Azure Blockchain Service', 'Azure Chaos Studio', 'Azure Access Point'];

test('services dragged from the palette keep their icons in every native export', async ({ page }) => {
  const consoleErrors: string[] = [];
  page.on('console', (msg) => {
    if (msg.type() === 'error') consoleErrors.push(msg.text());
  });

  await page.goto('/');
  await page.waitForLoadState('domcontentloaded');

  const result = await page.evaluate(async (wanted: string[]) => {
    const loader = await import('/src/utils/iconLoader.ts');
    const vsdx = await import('/src/services/visioVsdxExporter.ts');
    const pptxMod = await import('/src/services/pptxExporter.ts');

    const categories = ['azure ecosystem', 'blockchain', 'other', 'new icons', 'general', 'compute'];
    const library: Array<{ name: string; serviceName: string; path: string; category: string }> = [];
    for (const category of categories) {
      library.push(...(await loader.loadIconsFromCategory(category)));
    }

    const picked = wanted.map((want) => {
      const key = want.toLowerCase().replace(/[^a-z0-9]/g, '');
      return library.find((icon) => icon.name.toLowerCase().replace(/[^a-z0-9]/g, '') === key)
        ?? library.find((icon) => icon.path.toLowerCase().replace(/[^a-z0-9]/g, '').includes(key));
    });

    // Exactly what `addServiceNodeAtPosition` stores, plus the size React Flow
    // measures onto the node once it is on the canvas.
    const nodes = picked.map((icon, index) => ({
      id: `service-${index}`,
      type: 'azureNode',
      position: { x: 80 + index * 220, y: 120 },
      width: 150,
      height: 75,
      data: {
        label: icon?.name ?? wanted[index],
        serviceName: icon?.serviceName ?? wanted[index],
        category: icon?.category,
        iconPath: icon?.path ?? '',
      },
    }));
    const edges = [
      { id: 'e1', source: 'service-0', target: 'service-1', label: 'diagnoses' },
      { id: 'e2', source: 'service-1', target: 'service-3', label: 'publishes to' },
    ];

    const pkg = await vsdx.buildVsdxPackage(nodes as never, edges as never, 'Palette probe');
    const page1 = pkg.parts.find((p: { path: string }) => /page1\.xml$/i.test(p.path));
    const pageXml = typeof page1?.data === 'string' ? page1.data : '';

    const pptx = await pptxMod.buildDiagramSlidePptx(
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
      { diagramName: 'Palette probe', author: 'A', date: 'd', isDarkMode: false, diagram: { nodes, edges } } as never,
    );
    // The builder's own output is handed back for Node to finish and inspect:
    // the browser has no bare-specifier resolver, and the second half of the
    // export is a pure function that runs identically on either side.
    const built = (await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
    let binary = '';
    for (const byte of new Uint8Array(built)) binary += String.fromCharCode(byte);

    return {
      resolved: picked.map((icon, index) => ({ want: wanted[index], path: icon?.path ?? null })),
      vsdxMedia: pkg.parts.filter((p: { path: string }) => /\/media\//i.test(p.path)).length,
      vsdxForeignData: (pageXml.match(/<ForeignData /g) ?? []).length,
      pptxBase64: btoa(binary),
    };
  }, SERVICES);

  const authored = Buffer.from(result.pptxBase64, 'base64');
  const delivered = await (await nativizePackage(await JSZip.loadAsync(authored)))
    .generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  const deliveredZip = await JSZip.loadAsync(delivered);
  const slide = await deliveredZip.file('ppt/slides/slide1.xml')!.async('string');
  const pptxMedia = Object.keys(deliveredZip.files).filter((name) => /^ppt\/media\//.test(name)).length;
  const pptxPics = (slide.match(/<p:pic>/g) ?? []).length;
  const illegalConnectors = (slide.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g) ?? [])
    .filter((xml) => xml.includes('<a:custGeom>')).length;

  const unresolved = result.resolved.filter((entry) => !entry.path).map((entry) => entry.want);
  expect(unresolved, 'every palette service must resolve to an icon file').toEqual([]);
  expect(result.vsdxMedia, 'each service must embed an icon in the .vsdx').toBe(SERVICES.length);
  expect(result.vsdxForeignData, 'each .vsdx icon must be referenced by a ForeignData shape').toBe(SERVICES.length);
  expect(pptxMedia, 'the delivered deck must carry the tile icons as media').toBeGreaterThanOrEqual(SERVICES.length);
  expect(pptxPics, 'each tile must draw its icon on the slide').toBeGreaterThanOrEqual(SERVICES.length);
  expect(illegalConnectors, 'a <p:cxnSp> with custom geometry makes PowerPoint refuse the file').toBe(0);
  expect(consoleErrors.filter((e) => /icon|raster|export/i.test(e))).toEqual([]);
});
