import { expect, test } from '@playwright/test';
import type { Edge, Node } from 'reactflow';
import { buildInteractiveDiagramHtml } from '../../src/services/htmlDiagramExporter';

test.use({ viewport: { width: 1600, height: 1000 } });

const service = (id: string, x = 0, y = 0): Node => ({
  id, type: 'azureNode', position: { x, y }, width: 150, height: 75,
  data: { label: id, serviceName: 'App Service' },
});
const english = 'Retry request after a transient processing failure';
const japanese = '一時的な処理エラーが発生した場合は要求を再試行してください';
const cases: Array<{ name: string; nodes: Node[]; edges: Edge[] }> = [
  ...[english, japanese].flatMap(label => [1, 4].map(count => ({
    name: `${label === english ? 'English' : 'Japanese'} ${count} self-loop labels`,
    nodes: [service('worker')],
    edges: Array.from({ length: count }, (_, index) => ({
      id: `retry-${index}`, source: 'worker', target: 'worker',
      data: { label: count === 1 ? label : `${label} ${index + 1}`, stepNumber: index + 1 },
    })),
  }))),
  ...[{ x: 300, y: 0 }, { x: 0, y: 300 }].flatMap(position =>
    ['forward', 'reverse', 'bidirectional'].map(direction => ({
      name: `${position.x ? 'horizontal' : 'vertical'} ${direction} READ label`,
      nodes: [service('api'), service('db', position.x, position.y)],
      edges: [{ id: 'read', source: 'api', target: 'db', data: { label: 'READ', stepNumber: 1, direction } }],
    }))),
  {
    name: 'manual loop text offset keeps its badge on the route',
    nodes: [service('worker')],
    edges: [{
      id: 'manual', source: 'worker', target: 'worker',
      data: { label: english, stepNumber: 1, labelOffsetAuto: false, labelOffsetX: 400, labelOffsetY: 100 },
    }],
  },
];

// Each Playwright test receives a fresh document; exported scripts declare top-level consts.
for (const scenario of cases) {
  test(scenario.name, async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.route('**/*', route => route.abort());
    const html = await buildInteractiveDiagramHtml(scenario.nodes, scenario.edges, scenario.name);
    expect(html).not.toBeNull();
    await page.setContent(html!, { waitUntil: 'domcontentloaded' });
    await page.evaluate(() => document.fonts.ready);
    const result = await page.evaluate(() => {
      const svg = document.querySelector<SVGSVGElement>('.edges-layer')!;
      const width = Number(svg.getAttribute('width'));
      const height = Number(svg.getAttribute('height'));
      const boxes = Array.from(document.querySelectorAll<HTMLElement>('.node'), node => ({
        x: parseFloat(node.style.left), y: parseFloat(node.style.top),
        width: parseFloat(node.style.width), height: parseFloat(node.style.height),
      }));
      const labels = Array.from(svg.querySelectorAll<SVGTextElement>('.edge-label'), label => {
        const box = label.getBBox();
        return { text: label.textContent, x: box.x, y: box.y, width: box.width, height: box.height };
      });
      const overlaps = (a: typeof boxes[number], b: typeof boxes[number]) =>
        a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;
      const paths = Array.from(svg.querySelectorAll<SVGPathElement>('.edge-path'));
      const steps = Array.from(svg.querySelectorAll<SVGTextElement>('.edge-step'));
      return {
        labels: labels.map(label => label.text),
        nodeOverlaps: labels.flatMap(label => boxes.filter(box => overlaps(label, box))).length,
        labelOverlaps: labels.flatMap((label, index) => labels.slice(index + 1).filter(other => overlaps(label, other))).length,
        clipped: labels.filter(label => label.x - 1.5 < 0 || label.y - 1.5 < 0
          || label.x + label.width + 1.5 > width || label.y + label.height + 1.5 > height).length,
        stepDistances: steps.map((step, index) => {
          const x = Number(step.getAttribute('x'));
          const y = Number(step.getAttribute('y')) - 4;
          const path = paths[index];
          const length = path.getTotalLength();
          let distance = Infinity;
          for (let sample = 0; sample <= 512; sample++) {
            const point = path.getPointAtLength(length * sample / 512);
            distance = Math.min(distance, Math.hypot(point.x - x, point.y - y));
          }
          return distance;
        }),
      };
    });
    expect(errors).toEqual([]);
    expect(result.labels).toEqual(scenario.edges.map(edge => edge.data.label));
    expect(result.clipped, 'actual SVG text ink and its stroke fit the viewport').toBe(0);
    expect(result.nodeOverlaps, 'no sentence is hidden by a subsequently painted node').toBe(0);
    expect(result.labelOverlaps, 'loop sentences remain individually readable').toBe(0);
    for (const distance of result.stepDistances) expect(distance, 'badge stays on its own route').toBeLessThan(1);
  });
}

for (const paint of [
  { name: 'faded custom dash', dash: '10 2 3 2', opacity: 0.45 },
  { name: 'invisible solid', dash: 'none', opacity: 0 },
]) {
  test(`standalone HTML preserves ${paint.name} on the path and legend`, async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.route('**/*', route => route.abort());
    const html = await buildInteractiveDiagramHtml([service('a'), service('b', 400)], [{
      id: 'authored', source: 'a', target: 'b', data: { connectionType: 'security' },
      style: { stroke: '#006D77', strokeDasharray: paint.dash, opacity: paint.opacity },
    }], paint.name);
    expect(html).not.toBeNull();
    await page.setContent(html!, { waitUntil: 'domcontentloaded' });
    const actual = await page.locator('.edge-path, .legend-line path').evaluateAll(elements =>
      elements.map(element => {
        const style = getComputedStyle(element);
        return { color: style.stroke, opacity: Number(style.opacity),
          dash: style.strokeDasharray === 'none' ? [] : style.strokeDasharray.split(/[,\s]+/).map(parseFloat) };
      }));
    const expected = {
      color: 'rgb(0, 109, 119)', opacity: paint.opacity,
      dash: paint.dash === 'none' ? [] : [10, 2, 3, 2],
    };
    expect(actual).toEqual([expected, expected]);
    expect(errors).toEqual([]);
  });
}

test('standalone HTML does not invent a swatch for mixed connector styles', async ({ page }) => {
  await page.route('**/*', route => route.abort());
  const html = await buildInteractiveDiagramHtml([service('a'), service('b', 400)], [
    { id: 'default', source: 'a', target: 'b', data: { connectionType: 'security' } },
    {
      id: 'authored', source: 'a', target: 'b', data: { connectionType: 'security' },
      style: { stroke: '#006d77', strokeDasharray: 'none', opacity: 0.45 },
    },
  ], 'Varied');
  expect(html).not.toBeNull();
  await page.setContent(html!, { waitUntil: 'domcontentloaded' });
  await expect(page.locator('#legend')).toContainText('Security (varied)');
  await expect(page.locator('.legend-line')).toHaveCount(0);
});
