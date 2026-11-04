# Microsoft Fabric icons

These graphics and trademarks are the property and copyright of the Microsoft
group of companies, **not** MIT-licensed application code.

- [Microsoft Fabric icon terms](https://learn.microsoft.com/en-us/fabric/fundamentals/icons)
- [Microsoft Trademark and Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks)
- [Sam Debruyn's 2026 diagramming collection](https://debruyn.dev/2026/all-microsoft-fabric-icons-for-diagramming-2026-edition/)

Microsoft permits copying, distribution, and display for architectural diagrams,
training materials, and documentation. Label the product, experience, or item near
its icon. Do not crop, flip, rotate, distort, or change the icon shape, and do not
use Microsoft icons to represent your own product or service.

## Inventory and provenance

The palette includes **110** Fabric icons:

| Set | Count | Source |
| --- | ---: | --- |
| Current architecture families | 82 | Official `@fabric-msft/svg-icons` **8.2.0** |
| Historical graph designs | 2 | Official package **6.1.0** |
| Deprecated / historical designs | 12 | Official package **1.1.37** |
| OneLake security UI glyphs | 13 | De Bruyn's article, individually SHA-256 pinned |
| Fabric Capacity | 1 | Existing application-specific capacity symbol |

All **101 SVGs** listed in the article are represented byte-for-byte, including
the historical graph designs. The 74 overlapping current icons use the exact
article artwork from the official 8.2.0 package. Current Graph Model and Graph
Queryset retain their newer product designs; the article's older designs are
available separately with Legacy labels. Eight other current official families
and the local Capacity symbol are retained.

Current icons use the official 48px variant where available. Sample Workload has
no 48px variant in 8.2.0 and keeps its native 32px asset. Historical assets retain
their original native sizes (including 28px and 64px). No SVG bytes are edited.

The 13 monochrome UI glyphs are not branded product/service icons. They appear in
the **OneLake Security UI** drawing group; the 14 historical designs appear in
**Legacy / Deprecated**. Neither group is supplied to AI service generation, MCP
service discovery, or pricing. A white image background keeps their `currentColor`
(black in an external SVG image) visible in either theme without altering SVGs.

Existing service keys and icon paths are unchanged. The former Links icon is
displayed as **Shortcuts**, in the OneLake group, retaining Links / Graph Links
aliases. OneLake Shortcut(s) now resolve to this artwork, not the OneLake storage
logo; saved nodes explicitly using the OneLake path still keep their original icon.

## Reproducible synchronization

```bash
npm run icons:sync:fabric
node mcp-server/scripts/sync-icon-map.mjs
npm run test:icons
node scripts/test-fabric-icons.mjs
```

`scripts/fabric-icon-sources.json` locks the three official npm archives by npm
SHA-512 integrity and SHA-256, and every article SVG by URL, title, section,
SHA-256, and stable catalog destination. `fabric-manifest.json` records each
shipped file's source type, version or URL, and SHA-256 (including Capacity).
Article-only SVGs are vendored; synchronization reuses verified local copies and
downloads only missing or hash-mismatched ones. Mismatched downloaded artwork is
rejected rather than silently accepting a new design.

The synchronizer validates package identity, complete current-family coverage,
article hashes, and inert self-contained SVG content before writing files.
Newer npm `latest` versions are **reported, not adopted or treated as an error**.
To upgrade, review catalog coverage, changed artwork, terms, and the source lock
explicitly. Do not regenerate the source lock automatically from `latest`.

The refresh is an additive overlay: it does not delete unrelated upstream assets,
rename stable destinations, or change the deploy/upstream-sync workflow.
