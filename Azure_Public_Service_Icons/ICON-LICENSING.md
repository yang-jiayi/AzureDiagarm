# Microsoft icon sources, terms, and attribution

The SVGs are Microsoft graphics and trademarks, not MIT-licensed application
code. Each upstream source grants only the use described by its own terms; the
presence of an icon in this repository does not grant broader trademark or
product rights. Use icons only for the permitted purposes, label products near
their icons, and do not crop, flip, rotate, distort, or change their shape.

| Set | Official source and permitted use | Pinned inventory |
| --- | --- | --- |
| Azure architecture and service icons | [Azure icon guidance](https://learn.microsoft.com/en-us/azure/architecture/icons/): architectural diagrams, training materials, and documentation. Do not use Microsoft product icons to represent another product or service. | V25, 721 SVGs; `manifest.json`; source lock `../scripts/azure-icon-source.json` |
| Microsoft Fabric | [Fabric icon guidance](https://learn.microsoft.com/en-us/fabric/fundamentals/icons): architectural diagrams, training materials, slide decks, and documentation. | `@fabric-msft/svg-icons` 12.0.0; `fabric-manifest.json`; source lock `../scripts/fabric-icon-sources.json` |
| Microsoft Entra | [Entra architecture icon guidance](https://learn.microsoft.com/en-us/entra/architecture/architecture-icons): architectural diagrams, training materials, and documentation. Marketing communications are not permitted. | October 2023, 14 SVGs; `entra-manifest.json`; source map `../scripts/entra-icon-sources.json` |
| Microsoft 365 | [Archived Microsoft 365 architecture icon guidance](https://learn.microsoft.com/en-us/previous-versions/microsoft-365/solutions/architecture-icons-templates): architectural diagrams, training materials, and documentation. | April 2024 archive; `microsoft-product-manifest.json`; package pin in `../scripts/sync-microsoft-icons.mjs` |
| Power Platform and Dynamics 365 | [Power Platform icon guidance](https://learn.microsoft.com/en-us/power-platform/guidance/icons) and [Dynamics 365 icon guidance](https://learn.microsoft.com/en-us/dynamics365/get-started/icons). | December 2025 package set; `microsoft-product-manifest.json`; package pins in `../scripts/sync-microsoft-icons.mjs` |

The Microsoft 365 download is an **architecture-symbol set**, not an app-logo
collection. Its 963 SVG files are organized into six color/theme folders; the
205 selected palette assets are generic shapes such as chat, calendar, people,
document, or cloud symbols. Theme folders such as “Teams Purple” and “Planner
Green” do not make those generic shapes product logos. The package does not
provide official product-specific SVG logos for the Microsoft 365 applications
currently searched by those symbols. The palette therefore labels them as
symbols, not logos. Existing Microsoft 365 service names and aliases remain
available to generation, but the selected artwork is still a concept symbol and
is not assigned Azure pricing data.

Power Platform and Dynamics 365 are different: their official archives include
product-specific logos, so the catalog identifies those entries by product.
No unverified or third-party Microsoft 365 logos are added.

The Entra package maps all 14 official source SVGs individually. Existing Azure
identity art is reused where it already covers the same icon purpose. Only
Microsoft Entra Product Family and Microsoft Entra Workload ID are added as new
palette art; monochrome variants remain represented by their color counterpart
instead of being duplicated.

The Fabric inventory also retains the 101 SVGs from Sam Debruyn's
[diagramming collection](https://debruyn.dev/2026/all-microsoft-fabric-icons-for-diagramming-2026-edition/).
The compilation is attributed to the author; the Microsoft artwork remains
subject to Microsoft's icon terms. The two older designs that differ from the
current npm package are retained separately in `Article_Fabric_Icons/2026/`.

All downloaded packages are pinned by SHA-256 (and npm SHA-512 integrity where
available). Azure per-file manifests keep both the exact package-byte SHA-256
and a canonical LF-normalized SVG SHA-256 so integrity checks remain stable
across Git line-ending conversions. Other per-file hashes and the exact source
URLs are recorded in the generated manifests. The repeatable commands are:

```bash
npm run icons:sync
npm run icons:check
```

Synchronization does not self-update a pin from a moving “latest” release. A
source upgrade requires an explicit review of the official guidance, the
complete package inventory, terms, classifications, and generated hash
comparison before changing the source lock.
