import { test, expect } from '@playwright/test';

/**
 * Icon coverage for the native exports.
 *
 * `rasterizeIcons` needs a DOM, so under Node it silently returns an empty map
 * and every Node-side assertion about embedded icons is a false positive. This
 * spec runs the real export pipeline inside the browser, which is the only
 * place the icon path can actually be proven end to end.
 */
test('native PPTX and VSDX exports embed service icons', async ({ page }) => {
  const consoleErrors: string[] = [];
  page.on('console', (msg) => {
    if (msg.type() === 'error') consoleErrors.push(msg.text());
  });

  await page.goto('/');
  await page.waitForLoadState('domcontentloaded');

  const result = await page.evaluate(async () => {
    const raster = await import('/src/utils/exportIconRaster.ts');
    const vsdx = await import('/src/services/visioVsdxExporter.ts');
    const pptxMod = await import('/src/services/pptxExporter.ts');
    const mapping = await import('/src/data/serviceIconMapping.ts');

    const AKS = '/Azure_Public_Service_Icons/Icons/containers/10023-icon-service-Kubernetes-Services.svg';
    const FUNCTIONS = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
    const iconPath = '/Azure_Public_Service_Icons/Icons/compute/10021-icon-service-Virtual-Machine.svg';
    const svc = (id: string, label: string, x: number, y: number, parent?: string, icon = iconPath) => ({
      id,
      type: 'azureNode',
      position: { x, y },
      width: 150,
      height: 75,
      ...(parent ? { parentNode: parent } : {}),
      data: { label, serviceName: label, iconPath: icon },
    });
    const nodes = [
      { id: 'z', type: 'groupNode', position: { x: 0, y: 0 }, style: { width: 520, height: 320 }, data: { label: 'Application zone' } },
      svc('a', 'Azure Kubernetes Service', 60, 80, 'z', AKS),
      svc('b', 'Azure Functions', 320, 80, 'z', FUNCTIONS),
      // A third tile repeating an icon already on the sheet. Icons are embedded
      // once per DISTINCT icon and shared, so this must add a ForeignData and
      // no media part -- the drawing used to carry a byte-identical copy per
      // tile, which is how a 40-tile sheet reached 800KB of duplicate PNG.
      svc('c', 'Azure Functions (EU)', 60, 200, 'z', FUNCTIONS),
    ];
    const edges = [{ id: 'e', source: 'a', target: 'b', label: 'Managed identity authentication' }];

    const single = await raster.rasterizeIconToPng(iconPath, 128);
    const pkg = await vsdx.buildVsdxPackage(nodes as never, edges as never, 'Icon probe');
    const media = pkg.parts.filter((p: { path: string }) => /\/media\//i.test(p.path));
    const page1 = pkg.parts.find((p: { path: string }) => /page1\.xml$/i.test(p.path));
    const pageXml = typeof page1?.data === 'string' ? page1.data : '';

    const pptx = await pptxMod.buildDiagramSlidePptx(
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
      { diagramName: 'Icon probe', author: 'A', date: 'd', isDarkMode: false, diagram: { nodes, edges } } as never,
    );
    const blob = (await pptx.write({ outputType: 'blob' })) as Blob;

    // A no-icon control run: the difference proves the bytes really are icons.
    const bare = nodes.map((node) => (node.type === 'groupNode'
      ? node
      : { ...node, data: { label: node.data.label, serviceName: node.data.label } }));
    const bareBlob = (await (await pptxMod.buildDiagramSlidePptx(
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
      { diagramName: 'Icon probe', author: 'A', date: 'd', isDarkMode: false, diagram: { nodes: bare, edges } } as never,
    )).write({ outputType: 'blob' })) as Blob;

    return {
      singleIconBytes: single ? single.bytes.length : 0,
      vsdxMediaParts: media.length,
      vsdxForeignData: (pageXml.match(/<ForeignData /g) ?? []).length,
      pptxBytes: blob.size,
      pptxBytesWithoutIcons: bareBlob.size,
      mappedServices: Object.keys(mapping.SERVICE_ICON_MAP ?? {}).length,
    };
  });

  expect(result.singleIconBytes, 'a service SVG must rasterise to PNG bytes').toBeGreaterThan(500);
  expect(result.vsdxMediaParts, 'the two distinct icons must be embedded once each').toBe(2);
  expect(result.vsdxForeignData, 'each of the three tiles must reference an icon').toBe(3);
  expect(result.pptxBytes, 'the PPTX must be larger with icons than without').toBeGreaterThan(result.pptxBytesWithoutIcons);
  expect(result.mappedServices, 'the service icon map must be populated').toBeGreaterThan(100);
  expect(consoleErrors.filter((e) => /icon|raster|export/i.test(e))).toEqual([]);
});
