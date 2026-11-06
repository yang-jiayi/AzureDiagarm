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

The palette includes **117** Fabric icons:

| Set | Count | Source |
| --- | ---: | --- |
| Current architecture families | 88 | Official `@fabric-msft/svg-icons` **12.0.0**, published 2026-11-05 |
| Historical and compatibility designs | 15 | Official packages **8.2.0**, **6.1.0**, and **1.1.37** |
| OneLake security UI glyphs | 13 | De Bruyn's article, individually SHA-256 pinned |
| Fabric Capacity | 1 | Existing application-specific capacity symbol |

The package comparison is complete and recorded in `fabric-manifest.json`.
Compared with 8.2.0's 82 architecture families, 12.0.0 has 7 new assets,
renames/replaces `user_data_function_48_item.svg` with
`user_data_functions_48_item.svg`, removes the Function Set asset, changes the
same-name Copilot and Mirrored Catalog artwork, and leaves 78 same-name assets
unchanged. The renamed user-data path and the retired Function Set artwork remain
available for existing diagrams; Function Set keeps its former service-name and
capacity-resolution semantics but is excluded from new AI/MCP recommendations.

All **101** article SVGs remain hash-verified. Ninety-nine are also byte-identical
to the assets at their active catalog paths. The original Copilot and User Data
Function artwork, which differs from the current npm release, is retained at
`Article_Fabric_Icons/2026/` and listed separately in the manifest. The active
catalog continues to use the current 12.0.0 designs.

The `Action`, `Database Hub`, `Feature Store`, `Managed Identity`, `PostgreSQL
Database`, and `Rayfin` assets are labeled **Concept / drawing only**. They are
not emitted as services, passed to MCP service discovery, or treated as
capacity-consuming. Rayfin's product identity and purpose are unverified, so it
is shown as an unclassified symbol; the icons alone do not establish service or
billing semantics for these concepts. No cost range is assigned to them. Fabric
Apps is a separate current item and is classified as capacity consuming.
Function Set remains a Legacy item rather than being treated as retired based
solely on its removal from the new archive.

Current icons use the official 48px variant where available. Sample Workload has
no 48px variant and keeps its native 32px asset. Historical assets retain their
original native sizes. No SVG artwork is edited.

The 13 monochrome UI glyphs are not branded product/service icons. They appear in
the **OneLake Security UI** drawing group; the 15 historical and compatibility
designs appear in **Legacy / Deprecated**. Neither group is supplied to AI
service generation, MCP service discovery, or pricing. A white image background
keeps `currentColor` visible in either theme without altering SVGs.

Existing service keys and stable icon paths are retained. The former Links icon
is displayed as **Shortcuts**, in the OneLake group, retaining Links / Graph
Links aliases. OneLake Shortcut(s) resolve to this artwork, not the OneLake
storage logo; saved nodes explicitly using the OneLake path still keep their
original icon.

## Reproducible synchronization

```bash
npm run icons:sync:fabric
npm run icons:check:fabric
node mcp-server/scripts/sync-icon-map.mjs
npm run test:unit -- tests/fabricIcons.test.ts tests/fabricIconSync.test.ts
node scripts/test-fabric-icons.mjs
```

`scripts/fabric-icon-sources.json` locks the official npm archives by npm
SHA-512 integrity and SHA-256, records the reviewed 8.2.0-to-12.0.0 filename
replacement and publication timestamp, and pins every article SVG by URL,
title, section, SHA-256, and stable catalog destination. `fabric-manifest.json`
records each shipped file's source, version or URL, and SHA-256, as well as the
complete package comparison and article-asset inventory. The synchronizer
validates package identity, all 88 current families, the version comparison,
article hashes, and inert self-contained SVG content before writing files.

Newer npm `latest` versions are **reported, not adopted**. To upgrade, review
catalog coverage, changed artwork, terms, and the source lock explicitly. Do not
regenerate the source lock automatically from `latest`.

The refresh is an additive overlay: it does not delete unrelated upstream
assets, rename stable destinations, or change the deploy/upstream-sync workflow.
