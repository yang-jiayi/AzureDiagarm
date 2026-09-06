# Microsoft Product Architecture Diagram Builder

<div align="center">

![Azure](https://img.shields.io/badge/Azure-0078D4?style=for-the-badge&logo=microsoft-azure&logoColor=white)
![React](https://img.shields.io/badge/React-61DAFB?style=for-the-badge&logo=react&logoColor=black)
![TypeScript](https://img.shields.io/badge/TypeScript-3178C6?style=for-the-badge&logo=typescript&logoColor=white)
![GPT-5.x](https://img.shields.io/badge/GPT--5.x-412991?style=for-the-badge&logo=openai&logoColor=white)
![DeepSeek](https://img.shields.io/badge/DeepSeek-4D6BFF?style=for-the-badge&logoColor=white)
![Grok](https://img.shields.io/badge/Grok-000000?style=for-the-badge&logoColor=white)
![Mistral](https://img.shields.io/badge/Mistral-FF7000?style=for-the-badge&logoColor=white)
![Kimi](https://img.shields.io/badge/Kimi-1A1A1A?style=for-the-badge&logoColor=white)

**A professional AI-powered tool for designing, validating, and deploying Microsoft cloud architectures across Azure, Microsoft Fabric, Power Platform, and Dynamics 365**

[Live Demo](https://azurediagarm.mssql.biz) • [Upstream Project](https://github.com/Arturo-Quiroga-MSFT/azure-architecture-diagram-builder) • [Documentation](DOCS/ARCHITECTURE.md) • [Report Bug](../../issues)

</div>

---

> [!NOTE]
> This repository is a production-hardened deployment fork of the
> [Azure Architecture Diagram Builder](https://github.com/Arturo-Quiroga-MSFT/azure-architecture-diagram-builder).
> It retains the upstream project and license while adding secure Azure hosting,
> cloud collaboration safeguards, deployment automation, and responsive workspace improvements.

## 👤 Maintainers and attribution

This fork is maintained at [`yang-jiayi/AzureDiagarm`](https://github.com/yang-jiayi/AzureDiagarm).
The original project was created by **Arturo Quiroga**.
The latest customization and production hardening were created by **Swarm Data SE, Jiayi Yang**.

---

## 📖 Overview

Microsoft Product Architecture Diagram Builder is an enterprise-grade web application that empowers cloud architects to design, visualize, validate, and deploy Azure solutions. Supporting **16 AI models** across multiple providers — **GPT-6 Astra, GPT-5.1, GPT-5.2, GPT-5.4, GPT-5.4 Mini, GPT-5.6 Sol, GPT-5.6 Terra, GPT-5.6 Luna, Claude Opus 5, DeepSeek V3.2 Speciale, DeepSeek V4 Pro, Grok 4.1 Fast, Grok 4.3, Mistral Large 3, Kimi K2.5, and Kimi K2.7 Code** (via configured Azure OpenAI and Microsoft Foundry model deployments) — it transforms natural language descriptions into professional architecture diagrams while providing real-time cost estimates, Well-Architected Framework validation, multi-model comparison, and Infrastructure as Code generation.

Beyond editable **topology** diagrams, the app can also produce polished, whiteboard-style **Blueprint** diagrams (BETA) as shareable PNGs — ideal for presentations and design reviews.

### Why This Tool?

- **Speed**: Go from idea to deployable architecture in minutes, not hours
- **Accuracy**: Official Azure icons, real-time pricing from Azure Retail Prices API
- **Best Practices**: Built-in WAF validation ensures your architecture follows Microsoft recommendations
- **Actionable Output**: Generate deployment guides with Bicep/ARM templates ready for production

### 🎯 Scope & Intended Use

This tool is designed for **greenfield Azure** architecture design — sketching, validating, and
costing **new** solutions, and as an accelerator in architecture-design sessions and technical
workshops. It supports **Azure services only**.

The WAF validation produces a **diagram-only, design-time signal** intended to guide new designs.
It is **not** an audit of a deployed environment, and it is **not** intended for direct deployment
into existing, complex customer environments without further validation and review by a qualified
architect.

---

## ✨ Key Features

### 🤖 AI-Powered Architecture Generation
Describe your architecture in plain language to create a complete diagram with logical service groupings. **GPT-6 Astra** is the default for generation, validation, deployment guidance, and blueprints when its deployment is configured. Other configured models remain available for explicit selection and comparison.

**13 curated example prompts** included — from simple web apps to complex enterprise scenarios:
- Zero Trust enterprise networks with security segmentation
- Healthcare HIPAA-compliant platforms with FHIR APIs
- Black Friday e-commerce handling 50K orders/hour
- Industrial IoT with 5,000+ sensors and predictive maintenance
- Global multiplayer gaming backends for 500K+ concurrent players
- AI-powered chatbots, document processing, content moderation
- And more...

### 🖼️ Architecture Image Import
Upload an existing architecture diagram image (screenshot, whiteboard photo, or exported PNG) and let AI analyze it to recreate the architecture as an editable, interactive diagram with proper Azure service mapping.

### 💬 Architecture Chat (Conversational Refinement)
Refine your diagram through a natural back-and-forth conversation instead of one-shot prompts. Click the **Chat** button in the toolbar to open a docked side panel where you can iterate in plain English:

- Type changes like *"add Azure Front Door with WAF"* → *"now make it zone-redundant"* → *"add a Redis cache between the API and the database"*
- Each turn reads the **live canvas** as the source of truth, so follow-up requests naturally build on previous ones
- The assistant replies with a concise summary of proposed service and connection changes
- Review and select changes before applying; accepted edits can be undone. When automatic snapshots are enabled, the pre-apply snapshot must save successfully before changes are applied
- Suggestion chips help you get started, and the panel shows which model is active

### ✏️ Blueprint Diagrams (BETA)
Generate a hand-drawn, **whiteboard-style blueprint** of your architecture — nested zones (Azure / VNet / On-prem) with numbered, labeled arrows that trace the end-to-end flow, just like an architect explaining a system at a whiteboard. Three generation modes are available in the AI Generator modal:

- **Topology** — the classic deployable, editable diagram on the canvas
- **Blueprint** *(BETA)* — a polished whiteboard-style PNG (the PNG is the deliverable; re-download any time via **Export > Export Blueprint PNG**)
- **Both** *(BETA)* — a deployable topology **and** a Blueprint PNG from the same prompt, optionally generated in parallel

> Blueprint and Both modes require a compatible general-purpose OpenAI model, such as GPT-6 Astra or GPT-5.x. The app auto-switches if a third-party model is selected. A configurable legend position keeps the output presentation-ready.

Explicitly identified provider/proxy rate limits can retry automatically with the same prompt, model, reasoning and output limit, up to three HTTP attempts and two minutes of total cooldown. Budget exhaustion, unclassified 429 responses, request timeouts and incomplete output are surfaced instead of silently lowering generation quality. Countdown waits and active requests remain cancellable.

Both mode admits work against the server's concurrency budget. Retrying an incomplete run with the same brief regenerates only the missing deliverable, preserving accepted output and reusing its existing component manifest. Editing the brief starts a new run.

### 📋 IaC Import, Export & Drift Review
Import ARM JSON, Bicep, Terraform HCL, or Terraform state and turn the declared resources into an editable diagram. The round-trip workspace preserves the source baseline, compares it with the current canvas, exports Bicep or Terraform starter templates, and summarizes Azure what-if or Terraform plan JSON without ever running an apply operation.

Bicep baseline inspection recognizes conditional and nested declarations without executing expressions. Unexpanded modules, unresolved loop instance counts, and parsing limits are explicitly marked as incomplete. The workspace labels their counts and comparison results as provisional instead of implying that the full source has no differences.

### 🎯 Well-Architected Framework Validation
Validate your architecture against all five WAF pillars:
- **Security** — Identity, encryption, network isolation
- **Reliability** — High availability, disaster recovery
- **Performance** — Scaling, caching, optimization
- **Cost Optimization** — Right-sizing, reserved instances
- **Operational Excellence** — Monitoring, automation

Select specific recommendations to generate an improved proposal, then review the changes before applying them. Reviews show available finding sources, link findings to diagram resources, and retain a bounded review history. Edited diagrams and older comparison results are marked as stale; a finding not detected in a later review is not treated as proof of remediation. During analysis, a dismiss hint lets you close the panel and return later via the **Validation Score** button in the toolbar.

### 🔀 Multi-Model Comparison
Compare AI output side-by-side across all 15 models:

- **Architecture Comparison** — Run the same prompt through multiple models and compare service counts, connection counts, groups, workflow steps, token usage, and latency
- **Validation Comparison** — Run WAF validation across models and compare overall scores, pillar-level scores, severity breakdowns, finding counts, and quick wins. An inline WAF info box explains the five pillars being assessed
- **Save All Diagrams** — Download each model's architecture as a separate JSON file
- **Save Comparison Report** — Download a combined JSON report for offline analysis
- **Present Critique** — Click "Present" to have a talking avatar narrate the AI ranking with live word-by-word closed captions (requires `VITE_SPEECH_REGION`)
- **Apply Winner** — Pick a result and review its proposed changes before applying them to the canvas

### 🎙️ Avatar Presenter
After completing a model comparison, use **Present Critique** to have a photorealistic talking avatar narrate the AI ranking results aloud — or click **Narrate** in the Workflow Panel to have the avatar walk through every architecture step:
- A 3D avatar appears in a **draggable, resizable** floating panel — grab the header to reposition anywhere on screen, drag the bottom-right corner to resize

### 🖼️ Draggable Reference Image Viewer
When a sketch or image is uploaded for AI generation, the reference image stays visible as a floating panel:
- **Drag** by the header bar to reposition anywhere on the canvas
- **Resize** by dragging the purple corner handle — scales from 160 × 110 px up to 700 × 700 px
- **Expand** to full-screen overlay for detail
- **Collapse** to a small pill to stay out of the way
- Live **word-by-word closed captions** highlight each spoken word in real time, synchronized via the Speech SDK `wordBoundary` event
- **Keyless authentication** — no API keys stored; a lightweight Express.js token server runs co-located with nginx inside the container, acquiring an AAD token via `DefaultAzureCredential` (Azure Managed Identity) and returning it as `aad#{resourceId}#{aadToken}` on each request
- The "Present" / "Narrate" buttons are only visible when `VITE_SPEECH_REGION` is configured at image build time; no UI impact when not set

### 🗂️ Collapse All Groups
Toggle button to collapse or expand all groups at once for a bird's-eye view of the architecture. Restores original group sizes on expand.

### 🔄 Workflow Animation & Data Flow
Visualize how data flows through your architecture step-by-step:
- Interactive step-by-step walkthrough of the architecture
- Service highlighting — each step highlights the involved services on the canvas
- Animated connections use a static base line plus a moving flow overlay so direction stays readable
- Pause or resume supported connection motion from the toolbar, or control one connection from its right-click menu
- Respects the operating system's reduced-motion preference
- AI-generated descriptions for each workflow step
- **Narrate** button (when Speech is configured) — avatar speaks all steps aloud with live closed captions in a draggable, resizable panel

### 📄 Deployment Guide Generation with Bicep
Generate comprehensive deployment documentation including:
- Prerequisites and Azure resource requirements
- Step-by-step deployment instructions
- **Bicep templates** for each service (Infrastructure as Code)
- Post-deployment verification steps
- Security configuration recommendations
- **Grounded in Microsoft Learn** — before generating, the app searches official Microsoft Learn documentation for your services (via a server-side proxy to the Microsoft Learn MCP endpoint) and feeds the results into the model so commands, API versions, and Bicep schemas reflect current docs. A **“Grounded with Microsoft Learn”** references section lists the cited pages, which are also included in the exported Markdown. Grounding is best-effort: if docs are unavailable the guide still generates.

### 💰 Real-Time Multi-Region Cost Estimation
Get instant cost estimates across **8 Azure regions**:
- 🇺🇸 East US 2 · 🇦🇺 Australia East · 🇨🇦 Canada Central · 🇧🇷 Brazil South · 🇲🇽 Mexico Central · 🇳🇱 West Europe · 🇸🇪 Sweden Central · 🇸🇬 Southeast Asia

Features include:
- **Show/hide estimates** — keep detailed diagram styling while suppressing indicative cost figures
- **Per-node cost editor** — click a service cost badge to change its Tier/SKU, quantity, or enter a custom monthly unit price
- **PAYG ↔ Savings Plan (1-year) toggle** — flip the entire estimate between pay-as-you-go and 1-year commitment pricing. Each meter's **real 1-year Savings Plan rate** (from the Azure Retail Prices API) is used per-SKU when available; services without a savings-plan meter fall back to a representative discount, and Microsoft Fabric Capacity is **exact**. Usage-based services stay at PAYG.
- **“Prices as of” stamp** — every cost export records the pricing-data refresh date and the selected billing term.
- **Development vs production scenarios** — compare editable capacity, usage, commitment, negotiated-discount, support, currency, and planning-FX assumptions without changing the diagram.
- **True per-region meters** — pricing is pre-fetched per region from the Azure Retail Prices API (refresh anytime with `npm run pricing:refresh`), including per-region **Microsoft Fabric** capacity (CU) and OneLake storage rates.
- Color-coded legend (green/yellow/red based on cost thresholds)
- SKU and tier information for each service
- **Export Costs (CSV)** — per-service cost breakdown spreadsheet for the active region
- **Export Costs (All Formats)** — downloads a ZIP containing:
  - `README.md` — manifest explaining every file in the bundle
  - `-report.md` — **start here**: combined summary + full analysis in one Markdown file
  - `-report.html` — the same combined report as a self-contained HTML page (with the Mermaid pie chart rendered) for non-Markdown viewers
  - `-summary.md` — Markdown summary with tables for by-service, by-group, and by-category costs
  - `-analysis.md` — intelligent Markdown report: TL;DR callout, top cost drivers, a **Mermaid pie chart** of cost by category, fixed vs usage-based split, Reserved Instance flags, and a **ranked multi-region comparison table** showing cheapest/most expensive region and potential savings
  - `.csv` — spreadsheet for Excel
  - `.json` — structured breakdown for programmatic use
  - `-multiregion-comparison.csv` — per-service pricing across all 8 regions for side-by-side comparison

### 🟦 Microsoft Fabric Support
Design **Microsoft Fabric** data platforms alongside core Azure services:
- **83 Fabric icons** — all 82 architecture-oriented families from the official `@fabric-msft/svg-icons` 8.2.0 package, plus the app's Fabric Capacity symbol. Includes workloads, items, workspace/navigation symbols, and developer samples.
- **Capacity-aware costing** — Fabric Capacity (F-SKU) carries the cost; compute items show an **“incl. capacity”** badge instead of double-counting, and OneLake is billed as usage-based storage. The full F2→F2048 ladder (PAYG + 1-yr reserved) is built in.
- **Fabric example prompts** — medallion lakehouse, real-time intelligence, and Direct Lake Power BI scenarios

### 🟪 Power Platform, Copilot Studio & Dynamics 365 Support
Design business-application architectures next to Azure and Fabric, using Microsoft's current first-party logos:
- **8 Power Platform icons** — Power Platform, Power Apps, Power Automate, Power Pages, Dataverse, AI Builder, **Microsoft Copilot Studio**, and **Microsoft Agent 365**, from the official December 2025 Power Platform icon package.
- **16 Dynamics 365 icons** — the Dynamics 365 product family plus Sales, Sales Insights, Customer Service, Contact Center, Customer Insights, Customer Voice, Field Service, Finance, Finance and Operations, Supply Chain Management, Commerce, Human Resources, Intelligent Order Management, Project Operations, and Business Central.
- **Two dedicated palette categories** with bilingual (EN/JA) labels, descriptions, and search keywords.
- **AI-aware naming** — legacy and shorthand names resolve to the right icon (`Power Virtual Agents` → Copilot Studio, `Dynamics NAV` → Business Central, `D365 Sales` → Dynamics 365 Sales), and the generation prompt knows to group these products around **Microsoft Dataverse**.
- **Honest costing** — these products are licensed per user, so they never carry a fabricated Azure meter cost.

### ❓ Help & Learn Panel
An in-app **Help** button opens a centered guide so new users can get productive fast — Quick Start, a feature tour, example prompts, tips & FAQ, and resource links. (Opening it fires a `Help_Opened` telemetry event.)

### 💬 User Feedback
A built-in feedback widget captures a rating, category, and optional free-text comment. The token server delivers submissions through **Azure Communication Services Email** and can archive them in Azure Table Storage or Cosmos DB using managed identity. Diagnostic metadata is **opt-in**, with a preview of the submitted payload; diagram names, URL paths/queries, and browser details are excluded. Comments, prompts, and optional follow-up contact are never used as telemetry fallbacks. Archived feedback has a configurable retention period, and the submission receipt supports authorized deletion. Sent email copies and storage backups have separate retention policies; deleting an archive entry does not delete those copies.

### 🧠 Smart Layout Engine
- **Dagre-based hierarchical layout** with compound node support
- **12 AI layout rules** for clean, readable diagrams (directional flow, hub-and-spoke monitoring, connection caps)
- **Automatic group overlap resolution** — post-processing that detects and separates overlapping groups
- **Resizable group nodes** — drag handles to adjust group boundaries

### 📸 Auto-Snapshot & Version History
- When automatic snapshots are enabled, the pre-apply snapshot must save before accepted AI changes are applied
- Save named snapshots with descriptions
- Browse and restore previous versions
- Track architecture evolution over time
- Open stored snapshots in another tab; download JSON for sharing
- Entra-authenticated cloud autosave with immutable snapshots, comments, optimistic concurrency, and revocable viewer/editor links

### Reversible Editing and Local Drafts
- **Undo / Redo** covers diagram edits, including service and group labels, colors, connections, and layout changes. Use the toolbar or `Ctrl/Cmd+Z` and `Ctrl/Cmd+Shift+Z`.
- **Local autosave** keeps the active draft in IndexedDB and reports when a transaction has committed. A recovery prompt lets you restore or download the draft after reopening the page.
- Local drafts are **browser-local, not cloud backups**; authenticated cloud autosave is a separate feature. Clearing browser data removes local drafts, so use cloud storage or a JSON download for a portable copy. A concurrent tab cannot silently overwrite a newer draft revision.
- **AI change review** presents additions, removals, and modifications before they reach the canvas. Review selected changes, cancel a proposal, or retry generation without replacing current edits.
- **Create / Review / Export** tabs keep task-specific controls together. Service settings edit pricing inputs; WAF findings can locate their affected shapes on the canvas.

Workspace regression commands:

```sh
npm run test:workspace
npm run test:workspace:browser
npm run test:ai-ui
npm run test:inspector-ui
npm run test:modal-focus
```

The browser check uses the existing Playwright installation. Set `WORKSPACE_BROWSER_CHANNEL=msedge`
to use installed Microsoft Edge, and `WORKSPACE_ARTIFACT_DIR` to retain screenshots.

Public deployments fail closed without explicit authentication, ingress, access-list, deployment allowlist, and shared-budget configuration. Per-user daily token reservations and concurrency limits are shared across replicas; local development uses a clearly separate mode. See [runtime controls and privacy](server/SECURITY.md) for required settings, budget accounting, retention, and deletion behavior.

### 🎨 Professional Diagramming
- **Complete official Azure V24 icon package** — all 714 SVGs from Microsoft's July 2026 package, hash-verified from a committed manifest, plus the Microsoft Fabric, Power Platform, and Dynamics 365 icon sets
- **Purpose-based icon catalog** — 23 bilingual categories organize services by meaning while preserving every official source folder for compatibility
- **Semantic icon search** — search names, acronyms, aliases, categories, purposes, and Japanese keywords
- **Personal icon workspace** — favorites, recently used services, and custom collections persist locally
- **Virtualized catalog rendering** — only visible rows are mounted, keeping the complete 830-icon catalog responsive
- **185+ AI-mapped services** — with pricing, categories, and icon resolution (including Microsoft Fabric items, Power Platform, and Dynamics 365 applications)
- **Smart Grouping** — Logical organization (Frontend, Backend, Data, Security)
- **Editable Connections** — Labels, direction, per-edge animation, and custom styling
- **Alignment Tools** — Professional layout assistance
- **Title Block & Legend** — Document-ready diagrams
- **Canvas navigation hint** — a dismissable pill teaches scroll-to-zoom, right-click-drag to pan, and one-click **Fit to view** (so large diagrams are never "stuck")
- **Maximize the canvas** — collapse any of the eight toolbar sections independently, hide the complete toolbar, or use **Focus** mode to hide side panels and canvas chrome; preferences persist across sessions

### 📤 Export Options
| Format | Use Case |
|--------|----------|
| **PNG** | Documentation, presentations |
| **Editorial PNG** | Publication-style reference-architecture PNG |
| **Blueprint PNG** | Hand-drawn, whiteboard-style blueprint PNG (BETA) |
| **SVG** | Scalable vector graphics (true vector — edges preserved as paths) |
| **PPTX Slide** | Editable native PowerPoint shapes/text and embedded icons, with aspect-fitted layout and dark/light theme matching the canvas |
| **Interactive HTML** | Self-contained HTML with pan, zoom, and tooltips |
| **Visio (VSDX)** | Native service groups with embedded icons and editable text, glued connectors preserving arrows/styles, wrapped labels, nested-position support, and dark/light themes |
| **Draw.io** | Edit in diagrams.net — orthogonal (right-angle) connectors with wrapped, auto-sized edge-label boxes |
| **Workflow (Markdown)** | The workflow narrative as a `.md` doc — title block, prompt, grouped services, ordered step-by-step flow (service names resolved), connections table, optional WAF score + cost |
| **JSON** | Backup, version control |
| **CSV** | Cost analysis in Excel (single region) |
| **ZIP (All Formats)** | CSV + JSON + TXT summary + intelligent analysis + multi-region comparison |

### 📊 Application Insights Telemetry

- **Automatic tracking** — page views, session duration, unique users, geography
- **Feature usage events** — every key action is tracked as a custom event:
  | Event | Properties |
  |-------|------------|
  | `Architecture_Generated` | model, reasoning effort, prompt length, service/connection/group counts, elapsed time, tokens |
  | `Architecture_Validated` | model, overall WAF score, finding count, elapsed time |
  | `DeploymentGuide_Generated` | model, service count, bicep file count, elapsed time |
  | `Diagram_Exported` | format (png/svg/vsdx/drawio/pptx/html/workflow-md/json/csv), service count |
  | `ARM_Template_Imported` | filename, resource count |
  | `Image_Imported` | — |
  | `Models_Compared` | selected model |
  | `Recommendations_Applied` | recommendation count |
  | `Version_Operation` | save / restore |
  | `Region_Changed` | region ID |
  | `Start_Fresh` | — |
  | `Avatar_Presentation_Started` | model count, critique length |
- **Zero-impact when disabled** — if `VITE_APPINSIGHTS_CONNECTION_STRING` is not set, all tracking calls are no-ops
- **Privacy-friendly** — no PII collected; anonymous user IDs via cookies

---

## 🏗️ Architecture

### Application Flow

```mermaid
flowchart TD
    subgraph User["👤 User Interface"]
        A[Natural Language Input] --> B[AI Generator]
        C[ARM Template Upload] --> D[Template Parser]
        E[Drag & Drop Icons] --> F[Manual Design]
        G[Image Upload] --> H[Vision Analyzer]
    end

    subgraph AI["🤖 AI Services (14 Models)"]
        B --> I[Architecture Generation]
        D --> I
        H --> I
        I --> J[Diagram Specification]
    end

    subgraph Core["⚙️ Core Engine"]
        J --> K[React Flow Canvas]
        F --> K
        K --> L[Node Manager]
        K --> M[Connection Manager]
        K --> N[Group Manager]
        K --> O[Dagre Layout + Overlap Resolution]
    end

    subgraph Services["🔧 Services"]
        L --> P[Cost Estimation]
        L --> Q[WAF Validation]
        L --> R[Deployment Guide]
        P --> S[Azure Pricing API]
        Q --> T[AI Validator]
        R --> U[Bicep Generator]
    end

    subgraph Export["📤 Export"]
        K --> V[PNG/SVG]
        K --> W[Draw.io XML]
        K --> X[JSON Backup]
        P --> Y[CSV Cost Report]
        R --> Z[Deployment Docs]
        U --> AA[Bicep Templates]
    end

    style AI fill:#412991,color:#fff
    style Core fill:#0078D4,color:#fff
    style Services fill:#50E6FF,color:#000
    style Export fill:#00A36C,color:#fff
```

### Data Flow

```mermaid
sequenceDiagram
    participant U as User
    participant UI as React UI
    participant MS as Model Settings Store
    participant TS as Token Server (/api)
    participant AI as Azure OpenAI
    participant L as Microsoft Learn MCP
    participant P as Azure Retail Prices API
    participant FD as Feedback delivery / archive

    U->>UI: Describe architecture
    UI->>MS: Get model selection
    MS-->>UI: Selected model + settings
    UI->>TS: POST /api/openai (no key in browser)
    TS->>AI: Generate request (managed identity / key fallback)
    AI-->>TS: Diagram specification (JSON)
    TS-->>UI: Diagram specification (JSON)
    UI->>UI: Render nodes & auto-layout (Dagre + overlap resolution)
    UI->>P: Fetch regional pricing (pre-fetched per region)
    P-->>UI: Cost data (8 regions, PAYG / Reserved)

    U->>UI: Refine via Architecture Chat
    UI->>TS: POST /api/openai (modification prompt, live canvas)
    TS->>AI: Apply change
    AI-->>TS: Updated specification
    TS-->>UI: Updated diagram + change summary

    U->>UI: Validate architecture
    UI->>TS: POST /api/openai (WAF validation)
    TS->>AI: Validation request
    AI-->>TS: Recommendations by pillar
    TS-->>UI: Findings + score

    U->>UI: Generate deployment guide
    UI->>TS: POST /api/docs-search (grounding)
    TS->>L: Search Microsoft Learn
    L-->>TS: Cited doc snippets
    TS-->>UI: Sources
    UI->>TS: POST /api/openai (guide + Bicep, grounded)
    TS->>AI: Documentation request
    AI-->>TS: Guide + Bicep templates
    TS-->>UI: Guide + references

    U->>UI: Submit feedback
    UI->>TS: POST /api/feedback
    TS->>FD: Deliver or archive (managed identity)
```

### Component Architecture

```mermaid
graph TB
    subgraph Frontend["Frontend (React + TypeScript)"]
        App[App.tsx]
        App --> Canvas[React Flow Canvas]
        App --> AIGen[AI Generator Modal]
        App --> Chat[Architecture Chat Panel]
        App --> Help[Help & Learn Panel]
        App --> Validation[Validation Modal]
        App --> Deploy[Deployment Guide Modal]
        App --> Feedback[Feedback Widget]
        Canvas --> AzureNode[Azure Node + real icons]
        Canvas --> Layout[Layout Engine + overlap resolution]
    end

    subgraph Services["Services Layer"]
        azureOpenAI[azureOpenAI.ts]
        apiHelper[apiHelper.ts]
        modificationPrompt[modificationPrompt.ts]
        costService[costEstimationService.ts]
        pricing[regionalPricingService.ts]
        validator[architectureValidator.ts]
        deployGen[deploymentGuideGenerator.ts]
        docsGrounding[docsGroundingService.ts]
        feedbackService[feedbackService.ts]
        telemetry[telemetryService.ts]
    end

    subgraph Server["Server (co-located with nginx)"]
        TokenServer["token-server.js<br/>/api/openai · /api/docs-search<br/>/api/feedback · /api/speech-token"]
    end

    subgraph External["External APIs"]
        OpenAI[Azure OpenAI + Microsoft Foundry<br/>15 models]
        LearnMCP[Microsoft Learn MCP]
        PricingAPI[Azure Retail Prices API]
        Cosmos[(Azure Cosmos DB)]
        AppInsights[Application Insights]
        SpeechAPI[Azure Speech]
    end

    AIGen --> azureOpenAI
    Chat --> modificationPrompt --> azureOpenAI
    Validation --> validator
    Deploy --> deployGen --> docsGrounding
    Feedback --> feedbackService
    costService --> pricing

    azureOpenAI --> apiHelper --> TokenServer
    validator --> TokenServer
    deployGen --> TokenServer
    docsGrounding --> TokenServer
    feedbackService --> TokenServer
    pricing --> PricingAPI
    telemetry --> AppInsights

    TokenServer --> OpenAI
    TokenServer --> LearnMCP
    TokenServer --> Cosmos
    TokenServer --> SpeechAPI

    style Frontend fill:#61DAFB,color:#000
    style Services fill:#3178C6,color:#fff
    style Server fill:#412991,color:#fff
    style External fill:#0078D4,color:#fff
```

---

## 🔌 MCP Server & Microsoft Scout Integration

The Diagram Builder ships a **Model Context Protocol (MCP) server** (`mcp-server/`) that exposes its core capabilities as **12 tools, 3 resources, and 3 prompts**, so any MCP-compatible client — including **[Microsoft Scout](https://learn.microsoft.com/en-us/microsoft-scout/get-started)** — can design, validate, cost, and render Azure architectures conversationally.

### Tools

| Tool | Purpose |
|------|---------|
| `list_services` | Browse the Azure service catalog (categories, aliases, pricing, cost ranges) |
| `validate_architecture` | Score a design against Well-Architected Framework rules (deterministic, no LLM) |
| `harden_architecture` | **NEW** — deterministically clear pattern-level WAF anti-patterns (identity, WAF, API gateway, DB replica, cache, Key Vault, backup, monitoring, multi-region) and re-validate; collapses the manual add-service → re-validate loop into one call |
| `estimate_costs` | **Numeric** monthly costs (low/expected/high) from a distilled Azure Retail Prices snapshot — region- and term-aware (PAYG / 1-year reserved), with by-category totals. Instance-priced services use a representative SKU; Microsoft Fabric uses F-SKU capacity; usage-based services report curated catalog ranges |
| `generate_bicep` | Emit deployable Bicep with Well-Architected secure defaults pre-set (HTTPS-only + TLS 1.2, managed identity, Key Vault soft-delete/purge, health check, autoscale, staging slots, Storage/Cosmos/Redis hardening) + a structured map of which WAF finding each setting resolves. Design-time only |
| `generate_terraform` | **NEW** — deployable Terraform (azurerm) with the same Well-Architected secure defaults as `generate_bicep` |
| `generate_deployment_guide` | **NEW** — step-by-step Markdown deploy runbook (Bicep or Terraform): prereqs, deploy commands, a post-deploy hardening checklist, smoke tests, and teardown |
| `generate_manifest` | Emit an `az prototype` interchange manifest |
| `get_waf_rules` | Query WAF rules by pillar or service type |
| `render_diagram` | Render a diagram as SVG/HTML — with **real Azure icons**, smooth edges, and tiered layout |
| `export_reactflow_scene` | Produce a React Flow scene for the web app |
| `import_architecture` | **NEW** — inverse of the export tools: parse a manifest / React Flow scene back to the canonical `{services, connections, groups}` shape |

> **Structured outputs:** `validate_architecture`, `estimate_costs`, and `get_waf_rules` return typed `structuredContent` (validated against a declared `outputSchema`) alongside a concise human summary, and carry read-only/idempotent tool annotations — so agents consume the data machine-readably instead of parsing prose.

> **Resources & prompts:** beyond tools, the server publishes read-only **resources** (`azure://catalog/services`, `azure://waf/rules`, `azure://pricing/meta`) and starter **prompts** (`design-secure-web-app`, `design-event-driven-platform`, `harden-and-cost`) so any MCP client gets browsable reference data and guided entry points. Full reference: [`mcp-server/TOOLS.md`](mcp-server/TOOLS.md).

### Transport & auth
- **Dual transport** — stdio (local clients) and **Streamable-HTTP** (remote clients). Launch HTTP with `npm run start:http` (or `MCP_TRANSPORT=http`).
- **Bearer-token auth** — set `MCP_AUTH_TOKEN`; the server enforces `Authorization: Bearer <token>` with a constant-time comparison. A `/healthz` probe and a pre-auth liveness response on `/mcp` keep connector wizards happy.
- **Ops-ready** — stateless HTTP mode for multi-replica deployments, bounded stateful sessions for single-replica/local use, CORS preflight, and graceful shutdown.

### Use it from Scout
Register the deployed MCP endpoint (`https://<your-mcp-host>/mcp`) as a **custom remote MCP server** in Scout's Extensions panel with your Bearer token (stored encrypted). See [`SCOUT/README.md`](SCOUT/README.md) for the walkthrough, and deploy an isolated MCP instance with [`scripts/deploy-mcp-instance.sh`](scripts/deploy-mcp-instance.sh).

### Use it from VS Code (GitHub Copilot)
The MCP server also works in **GitHub Copilot agent mode** in VS Code — no code changes, just a config entry. Create a `.vscode/mcp.json` pointing at the deployed server, with the bearer token supplied via an input prompt so no secret is committed:

```jsonc
{
  "servers": {
    "azure-diagram-builder": {
      "type": "http",
      "url": "https://<your-mcp-host>/mcp",
      "headers": { "Authorization": "Bearer ${input:aadb-token}" }
    }
  },
  "inputs": [
    { "id": "aadb-token", "type": "promptString", "description": "AADB MCP bearer token", "password": true }
  ]
}
```

Reload the MCP servers (**MCP: List Servers**), paste your token when prompted (the value in `.env.mcp`), and the 12 tools appear in Copilot Chat. Attach resources via **Add Context > MCP Resources**, and invoke prompts with `/azure-diagram-builder.design-secure-web-app`. Prefer local development? The bundled config also defines a `stdio` server that runs `mcp-server/dist/index.js` (run `npm run build` in `mcp-server/` first).

---

## 🔐 Production deployment

The generic Azure Developer CLI (`azd`) path is intentionally retired in this
secured fork. It did not configure the required Azure Front Door route, Easy
Auth enterprise-application assignment, Conditional Access boundaries, or
direct-origin isolation. The retained `azure.yaml` exits before provisioning,
packaging, or deployment so it cannot accidentally create a weaker public
environment.

Production updates run only through
[`AzureDiagarm sync and deploy`](.github/workflows/azurediagarm-sync-deploy.yml).
Merge a reviewed pull request into `main`; its push starts the release
automatically. **Do not use manual workflow dispatch for ordinary releases.**
Dispatch on `main` is reserved for guarded upstream synchronization.
The workflow validates the application and servers, builds and pushes an image to ACR, creates a
Container Apps revision with health probes, preserves authentication and origin
controls, purges Front Door, and verifies the deployed security boundary.
It verifies the validated checkout against current `main` before Azure changes
and again immediately before the revision update. See the
[production runbook](deployment/azurediagarm/README.md#safe-retries-and-rollbacks)
before retrying a failed run: older workflow versions do **not** inherit this guard.

Configure `ACCESS_ADMIN_EMAIL` as a GitHub Actions **secret**, not a repository
variable, so the administrator address is masked in public workflow logs.
Existing installations using a variable must copy the same value to the
like-named secret before deploying; the application administrator does not
change.

OpenAI proxy quotas use an atomic Azure Table Storage counter when
`AZURE_TABLES_ENDPOINT` is configured; the production workflow requires this
shared backend so all Container Apps replicas enforce one per-client limit.
Storage failures fail closed with a short retry interval. For an externally
managed Storage account, grant the runtime identity `Storage Table Data
Contributor`; the application ensures the configured `feedback` and
`ratelimit` tables exist. The workflow also rejects broad `Owner`,
`Contributor`, or `User Access Administrator` assignments on the runtime
identity. Production MCP HTTP requests use stateless mode so they can move
safely across replicas. Stateful mode remains bounded to 100 sessions by
default, expires sessions after 30 minutes idle or two hours absolute, and
closes them during graceful shutdown.

---

## 🚀 Getting Started

### Prerequisites

- **Node.js 22**
- **npm** or **yarn**
- **Azure OpenAI** resource with a model deployment for managed models, or a
  user-owned Azure OpenAI / official OpenAI endpoint when BYO AI is enabled

### Installation

1. **Clone the repository**
```bash
git clone https://github.com/yang-jiayi/AzureDiagarm.git
cd AzureDiagarm
```

2. **Install dependencies**
```bash
npm ci
```

3. **Configure environment variables**

Create a `.env` file in the project root:

```bash
# Azure OpenAI Configuration (Required)
#
# SECURITY: Azure OpenAI calls are proxied server-side by the co-located token
# server (server/token-server.js) via the /api/openai endpoint. The API key is
# NEVER shipped to the browser. Keyless auth (managed identity / `az login`) is
# preferred; a key is only used as a fallback when AZURE_OPENAI_API_KEY is set.
#
# VITE_AZURE_OPENAI_ENDPOINT is a non-secret build-time flag that signals the
# UI that AI is configured. In dev, scripts/start-token-server.sh bridges the
# VITE_ values to the server-side names (AZURE_OPENAI_ENDPOINT / _API_KEY).
VITE_AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
AZURE_OPENAI_API_KEY=your-api-key-here   # optional server-side fallback; prefer managed identity
VITE_AZURE_OPENAI_DEPLOYMENT=your-default-deployment

# Optional user-owned endpoints. Disabled by default for self-hosted installs.
ALLOW_BYO_AI_ENDPOINTS=false

# Multi-model deployments (16 supported models; configure only real deployments)
# GPT-6 Astra is the preferred application model
VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA=your-gpt6-astra-deployment
# Optional OpenAI GPT-5.x deployments
VITE_AZURE_OPENAI_DEPLOYMENT_GPT51=your-gpt51-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_GPT52=your-gpt52-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_GPT54=your-gpt54-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_GPT54MINI=your-gpt54-mini-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_GPT56SOL=your-gpt56-sol-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_GPT56TERRA=your-gpt56-terra-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_GPT56LUNA=your-gpt56-luna-deployment
# Anthropic Messages API in Microsoft Foundry
VITE_AZURE_FOUNDRY_ENDPOINT=https://your-resource.services.ai.azure.com/
VITE_AZURE_FOUNDRY_DEPLOYMENT_CLAUDE_OPUS5=your-claude-opus-5-deployment
AZURE_FOUNDRY_ENDPOINT=https://your-resource.services.ai.azure.com/
AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS=your-claude-opus-5-deployment
# Partner models (Chat Completions API)
VITE_AZURE_OPENAI_DEPLOYMENT_DEEPSEEK=your-deepseek-v32-speciale-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_DEEPSEEK_V4_PRO=your-deepseek-v4-pro-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_GROK4FAST=your-grok-41-fast-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_GROK43=your-grok-43-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_MISTRALLARGE3=your-mistral-large-3-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_KIMIK25=your-kimi-k2-5-deployment
VITE_AZURE_OPENAI_DEPLOYMENT_KIMIK27CODE=your-kimi-k2-7-code-deployment

# Reasoning model configuration
VITE_REASONING_EFFORT=medium  # none | low | medium | high

# Optional: Cloud storage for sharing
AZURE_COSMOS_ENDPOINT=https://your-cosmos.documents.azure.com:443/
COSMOS_DATABASE_ID=diagrams
COSMOS_CONTAINER_ID=diagrams

# Optional: Application Insights telemetry
# Create an App Insights resource in Azure Portal and paste the connection string
VITE_APPINSIGHTS_CONNECTION_STRING=InstrumentationKey=...;IngestionEndpoint=...

# Optional: Avatar Presenter (enables "Present Critique" button in Compare Models)
# Requires an Azure Speech resource with Custom Subdomain enabled and
# the ACA managed identity assigned the "Cognitive Services Speech User" RBAC role
VITE_SPEECH_REGION=westus2                 # Build-time: controls visibility of the "Present" button
AZURE_SPEECH_REGION=westus2               # Runtime: read by the co-located token server
AZURE_SPEECH_RESOURCE_ID=/subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.CognitiveServices/accounts/<speech-account-name>
```

`VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA` must name a real GPT-6 Astra deployment,
not a renamed GPT-5.6 deployment. Include the same name in the API's managed-model
allowlist. The deployment-only template `infra/gpt6-astra.bicep` codifies the
verified model/version and usage-based SKU under an existing OpenAI account;
it does not recreate the account or modify networking, roles, or legacy models.
Confirm model availability and quota before applying it in another environment.
On first use with Astra configured, saved GPT-5.6 selections and
feature overrides migrate to Astra, preserving supported reasoning levels.
The migration is persisted and does not repeatedly overwrite later explicit
model choices. Separate bring-your-own endpoint settings are not migrated.
Installations without Astra retain their existing configured-model portfolio.

4. **Start the development server**
```bash
# Frontend only — Vite on http://localhost:3000.
# NOTE: /api/openai is NOT served, so AI generation/chat will fail with a
# 500/503. Use this only for pure UI work that doesn't call the AI backend.
npm run dev

# Recommended for local testing: starts the token server (:3001, serves
# /api/openai) AND Vite (:3000) together, with pre-flight checks and cleanup.
# Reads .env, bridges VITE_AZURE_OPENAI_* → server-side AZURE_OPENAI_* so the
# /api/openai proxy works, checks `az login` (warns if not on the expected
# subscription), confirms ports 3000/3001 are free, installs deps if missing,
# then runs Vite in the foreground. Single Ctrl-C cleans up all children.
# Logs land in `.dev-logs/`. The AZURE_SPEECH_* vars are OPTIONAL — without
# them the avatar "Present" button is disabled but everything else works.
npm run dev:full                    # alias for ./scripts/dev-all.sh
./scripts/dev-all.sh                # token server + Vite (same thing)
./scripts/dev-all.sh --with-mcp     # also build & start the MCP server
./scripts/dev-all.sh --skip-az-check  # skip the Azure CLI verification

# With avatar presenter on an alternate port (Vite :3002 + token server):
npm run dev:avatar
```

### Official Icon Library Maintenance

```bash
npm run icons:sync             # overlay every official Microsoft icon package
npm run icons:sync:azure       # Azure architecture icons only
npm run icons:sync:fabric      # Microsoft Fabric icons only
npm run icons:sync:microsoft   # Power Platform + Dynamics 365 icons only
npm run icons:check:microsoft  # fail if the pinned packages drifted from the catalog
npm run test:icons             # verify all manifest files and SHA-256 hashes
```

The sync is additive: it preserves legacy paths used by saved diagrams and service mappings.
Official Microsoft icon assets remain subject to Microsoft's own terms, not this repository's
MIT License:

| Icon set | Terms |
| --- | --- |
| Azure architecture icons | [Azure icon terms](https://learn.microsoft.com/azure/architecture/icons/) |
| Microsoft Fabric icons | [Fabric icon terms](https://learn.microsoft.com/fabric/fundamentals/icons) |
| Power Platform & Copilot Studio icons | [Power Platform icon terms](https://learn.microsoft.com/power-platform/guidance/icons) |
| Dynamics 365 icons | [Dynamics 365 icon terms](https://learn.microsoft.com/dynamics365/get-started/icons) |

Microsoft permits these icons in architectural diagrams, training materials, and documentation.
They must not be cropped, flipped, rotated, or distorted, and must not be used to represent a
non-Microsoft product.

#### Avatar narrator troubleshooting

If the avatar panel opens but the video stays blank (audio may also fail), open DevTools and look for `[avatar] ICE state: failed`. That means the WebRTC peer connection cannot reach `relay.communication.microsoft.com:3478` (UDP) — common on corporate networks, VPNs, and some home ISPs.

The app already mitigates this: it offers both the UDP candidate and a TCP/443 fallback (`turn:relay.communication.microsoft.com:443?transport=tcp`) and forces `iceTransportPolicy: 'relay'`. If you still see ICE failures, your network is also blocking outbound 443 to that host — escalate to your network team or test from a different network.

To experiment with the legacy UDP-only path, run this in the browser console **before** clicking Narrate:

```js
window.__AVATAR_FORCE_TCP__ = false;
```

Microsoft Edge with **Strict** Tracking Prevention may log warnings such as `Tracking Prevention blocked access to storage for …tts.speech.microsoft.com…`. These are harmless — the Speech SDK does not need site storage for the WebRTC flow.

5. **Open your browser**
Navigate to `http://localhost:3000`

### Docker Deployment (Local Only)

The image defaults to authenticated **public** mode. A workstation run must
explicitly select local mode and bind the published port to loopback. Never
expose this local configuration to other users or copy it into a public deployment.

```bash
# Use your real Azure OpenAI endpoint and genuine GPT-6 Astra deployment name.
# These two values are public configuration, not API credentials.
OPENAI_ENDPOINT="https://your-resource.openai.azure.com/"
ASTRA_DEPLOYMENT="gpt-6-astra"

# Build-time VITE_* values select the model; credentials stay server-side.
docker build -t azure-diagram-builder \
  --build-arg VITE_AZURE_OPENAI_ENDPOINT="$OPENAI_ENDPOINT" \
  --build-arg VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA="$ASTRA_DEPLOYMENT" .

# Forward an optional API key from your shell without putting its value in
# the command. Obtain it through your approved local secret-management process.
docker run --rm -p 127.0.0.1:8080:80 \
  -e APP_DEPLOYMENT_MODE=local \
  -e ACCESS_CONTROL_ENABLED=false \
  -e AI_BUDGET_STORE=memory \
  -e AZURE_OPENAI_ENDPOINT="$OPENAI_ENDPOINT" \
  -e AZURE_OPENAI_ALLOWED_DEPLOYMENTS="$ASTRA_DEPLOYMENT" \
  -e AZURE_OPENAI_API_KEY \
  azure-diagram-builder
```

Open `http://127.0.0.1:8080`. The editor can start without a model credential;
real AI calls require a server-side credential authorized for the configured
deployment. A host-side `az login` is not automatically available inside Docker.
In Azure, production uses the configured managed identity instead of an API key.
Never pass a provider API key as a `VITE_*` value or Docker build argument.

Additional model deployments need matching public build arguments and
server-side provider allowlists; see [runtime controls](server/SECURITY.md).

### Azure Container Apps Deployment

For the existing AzureDiagarm production app, follow the
[production runbook](deployment/azurediagarm/README.md): reviewed changes merged
into `main` release automatically. Manual workflow dispatch is only for upstream
synchronization, not an alternative release button or a historical rollback.
Do not use `azd` or the generic `scripts/deploy_aca.sh` compatibility helper for
ordinary AzureDiagarm releases; they are not substitutes for this release workflow.

### Public Deployment Authentication (Required)

Entra Easy Auth, the application access list, protected Front Door ingress, and
durable shared budgets are mandatory in public mode. Setting runtime flags alone
does not provision or prove these controls. Keep the existing tenant, app
registration, managed identity, secrets, and access assignments; a UI release
does not require replacing them.

See [access and authentication](DOCS/APP-ACCESS-AND-AUTH.md) and the
[public runtime requirements](server/SECURITY.md#local-versus-public).
Anonymous app/API requests must remain protected. Crawler blocking and missing
anonymous social previews are intentional for this private application.

---

## 📚 Usage Guide

### Creating Diagrams

#### Method 1: AI Generation (Recommended)
1. Click **"Generate with AI"** in the toolbar
2. Describe your architecture in natural language, or pick from **13 curated example prompts**
3. Choose a **diagram mode** — Topology, Blueprint (BETA), or Both (BETA)
4. Select your AI model (any of the 12 options) and reasoning level
5. Click **Generate** — the architecture is created with auto-layout and workflow animation

#### Method 2: Image Import
1. Click **"AI Generate"** and expand the image upload section
2. Upload a screenshot or photo of an existing architecture diagram
3. AI analyzes the image and generates an editable description
4. Click **Generate** to recreate it as an interactive diagram

#### Method 3: ARM Template Import
1. Click **"Import ARM"** in the toolbar
2. Paste your ARM template JSON
3. AI parses and visualizes your existing infrastructure

#### Method 4: Manual Design
1. Browse the icon palette (left sidebar)
2. Drag services onto the canvas
3. Connect services by clicking and dragging between them
4. Double-click labels to edit

### Validating Architecture

1. Design or generate your architecture
2. Click **"Validate Architecture"** in the toolbar
3. Review recommendations by WAF pillar
4. Check the improvements you want to implement
5. Click **"Regenerate with Selected"** to apply

### Comparing Models

#### Architecture Comparison
1. Click **"Compare Models"** in the toolbar
2. Select which models to include and set reasoning effort
3. Enter a prompt (or pick from sample prompts)
4. Click **Compare** — all models run in parallel
5. Review side-by-side results (service count, tokens, latency)
6. Click **"Use This Architecture"** on the best result

#### Validation Comparison
1. Generate an architecture first
2. Click **"Compare Validation"** in the toolbar
3. Select models and click **Compare**
4. Compare WAF scores, pillar breakdowns, severity counts
5. Click **"Use This Validation"** on the preferred result

### Generating Deployment Guide

1. Complete your architecture design
2. Click **"Deployment Guide"** in the toolbar
3. Review the generated documentation:
   - Prerequisites
   - Deployment steps
   - Bicep templates (expandable)
   - Security recommendations
4. Download individual Bicep files or all as ZIP

### Working with Costs

- Costs update automatically as you add services
- Use the **Region Selector** to compare pricing
- Legend shows color-coded cost ranges
- Export to CSV for detailed analysis

---

## 🛠️ Technology Stack

| Category | Technologies |
|----------|-------------|
| **Frontend** | React 18, TypeScript, React Flow, Vite |
| **AI** | Azure OpenAI + Microsoft Foundry: GPT-5.x, Claude Opus 5, and partner models; Responses, Chat Completions, and Anthropic Messages APIs |
| **Styling** | CSS3, html-to-image |
| **Serving** | nginx:alpine (Docker), Vite dev server (local) |
| **APIs** | Azure Retail Prices API |
| **Export** | JSZip, Draw.io XML format, PptxGenJS (client-side PPTX) |
| **Avatar** | Azure Cognitive Services Speech SDK (TTS Avatar), `DefaultAzureCredential` (keyless), Express.js token server |
| **MCP** | Model Context Protocol server (`@modelcontextprotocol/sdk`), stdio + Streamable-HTTP, Bearer auth — consumable by Microsoft Scout |
| **Persistence** | Azure Blob Storage for authenticated diagrams; ACS Email with optional Table Storage or Cosmos DB archives for feedback |
| **Docs grounding** | Microsoft Learn MCP endpoint (via server-side `/api/docs-search` proxy) |
| **Deployment** | Docker, Azure Container Apps |

---

## 📁 Project Structure

```
azure-diagrams/
├── src/
│   ├── components/           # React components
│   │   ├── AIArchitectureGenerator.tsx  # AI generation modal
│   │   ├── ImageUploader.tsx  # Diagram image import
│   │   ├── WorkflowPanel.tsx  # Workflow animation
│   │   ├── ValidationModal.tsx  # WAF validation
│   │   ├── CompareModelsModal.tsx  # Multi-model architecture comparison
│   │   ├── CompareValidationModal.tsx  # Multi-model validation comparison
│   │   ├── DeploymentGuideModal.tsx  # Deployment guides
│   │   ├── ModelSettingsPopover.tsx  # Model selector
│   │   ├── IconPalette.tsx
│   │   ├── AzureNode.tsx / GroupNode.tsx
│   │   ├── Legend.tsx / TitleBlock.tsx
│   │   └── ...
│   ├── services/             # Business logic
│   │   ├── azureOpenAI.ts    # AI integration (Responses + Chat Completions API), via /api/openai proxy
│   │   ├── architectureValidator.ts  # WAF validation with ModelOverride support
│   │   ├── deploymentGuideGenerator.ts  # Guides & Bicep generation
│   │   ├── docsGroundingService.ts  # Microsoft Learn grounding for deployment guides
│   │   ├── modificationPrompt.ts  # Architecture Chat: live-canvas modification prompts
│   │   ├── feedbackService.ts  # User feedback delivery and metadata-only telemetry
│   │   ├── costEstimationService.ts  # Pricing engine (PAYG / Reserved)
│   │   ├── drawioExporter.ts  # Draw.io export
│   │   ├── pptxExporter.ts   # PowerPoint slide export (PptxGenJS, dark/light theme)
│   │   ├── regionalPricingService.ts  # Multi-region pricing
│   │   ├── apiHelper.ts      # Dual API format builder (Responses/Chat Completions)
│   │   ├── versionStorageService.ts  # Version history
│   │   ├── wafPatternDetector.ts  # Rule-based WAF pattern checks
│   │   ├── avatarPresenter.ts   # Talking avatar: Speech SDK, ICE relay, word-boundary captions
│   │   └── telemetryService.ts  # Application Insights telemetry
│   ├── stores/               # State management
│   │   └── modelSettingsStore.ts  # Multi-model settings (15 models)
│   ├── hooks/                # Shared React hooks
│   │   └── useDraggableResizable.ts  # Pointer-capture drag-to-move + drag-to-resize hook
│   ├── data/                 # Static data
│   │   ├── pricing/          # Regional pricing data (568 files: 71 services × 8 regions)
│   │   ├── azurePricing.ts   # Service mappings
│   │   └── serviceIconMapping.ts  # Icon mappings
│   ├── utils/                # Utilities
│   │   ├── iconLoader.ts     # Icon matching
│   │   ├── layoutEngine.ts   # Dagre layout + overlap resolution
│   │   ├── layoutPresets.ts  # Reference architectures
│   │   ├── groupUtils.ts     # Shared group collapse/fit utilities
│   │   ├── captureCanvas.ts  # html-to-image capture with SVG edge pre-inlining
│   │   └── modelNaming.ts    # Model display names
│   └── App.tsx               # Main application
├── server/                   # Token server (co-located with nginx in the container)
│   └── token-server.js       # Express.js: /api/speech-token + /api/ice-token + /api/openai + /api/docs-search + /api/feedback (Managed Identity, keyless)
├── scripts/                  # Deployment & data scripts
│   ├── deploy_aca.sh         # Configurable ACA deployment (reads from .env)
│   ├── update_aca.sh         # Author's ACA deployment (hardcoded resources)
│   ├── deploy-mcp-instance.sh  # Deploy the isolated MCP server ACA instance
│   └── fetch-multi-region-pricing.sh  # Refresh per-region pricing (npm run pricing:refresh)
├── Azure_Public_Service_Icons/  # 714 Azure + 83 Fabric + 24 Power Platform/Dynamics 365 icons
├── mcp-server/               # MCP server (12 tools + 3 resources + 3 prompts, stdio + HTTP, Bearer auth)
│   └── src/                  # serviceCatalog, wafDetector, layoutEngine, svgRenderer, htmlRenderer
├── SCOUT/                    # Microsoft Scout integration notes & sample session
├── DOCS/                     # Documentation
└── Dockerfile               # Container configuration
```

---

## 📖 Documentation

- **[System Architecture](DOCS/ARCHITECTURE.md)** - Technical deep-dive
- **[Layout Engines: Dagre vs ELK](DOCS/LAYOUT_ENGINES_COMPARISON.md)** - Comparison, pros/cons, and best fit analysis
- **[Regional Pricing](DOCS/REGIONAL_PRICING_IMPLEMENTATION.md)** - Cost estimation details
- **[Services Pricing](DOCS/services_pricing.md)** - Supported services and tiers
- **[Icon Mapping](DOCS/ICON_MAPPING.md)** - Service to icon reference

---

## 🌟 What's New

### July 2026 — New Frontier Models & Expanded MCP Toolset

#### 🤖 14 AI models
Added three new GPT-5.6 reasoning variants — **GPT-5.6 Sol**, **GPT-5.6 Terra**, and **GPT-5.6 Luna** — plus **Kimi K2.7 Code**, and retired the GPT-5.2 / GPT-5.3 Codex deployments. The lineup is now **14 models**, each assignable per feature (generation, validation, deployment guide, blueprint).

#### 🔌 MCP server grew to 12 tools + resources + prompts
Four new tools: **`harden_architecture`** (deterministically clears topology WAF anti-patterns in one call), **`generate_terraform`** (azurerm IaC with the same secure defaults as Bicep), **`generate_deployment_guide`** (Markdown deploy runbook), and **`import_architecture`** (round-trips a manifest / React Flow scene back in). The server now also exposes **3 MCP resources** (catalog, WAF rules, pricing) and **3 starter prompts**. See [`mcp-server/TOOLS.md`](mcp-server/TOOLS.md).

#### 🎨 Diagram rendering polish
Two-line wrapped edge labels with collision-avoided placement, opaque label chips (no more strike-through), a redesigned footer band (wrapped legend + cost total), distinct per-group header colors, and cleaner type-badge abbreviations.

### June 2026 — MCP Server, Microsoft Scout, Fabric & Pricing Upgrades

#### 🔌 MCP server + Microsoft Scout
The Diagram Builder is now an **MCP server** (8 tools: list / validate / estimate / **generate_bicep** / render / export / manifest / WAF) with stdio + Streamable-HTTP transports and Bearer auth, registerable as a remote extension in **Microsoft Scout**. `estimate_costs` returns numeric live-derived pricing, `generate_bicep` emits WAF-hardened IaC, and three tools now return typed `structuredContent`. SVG rendering gained **real Azure icons** (embedded glyphs, emoji fallback), smooth bezier edges, tighter layout, and far fewer edge crossings.

#### 🟦 Microsoft Fabric support
83 Fabric icons with complete official 8.2.0 architecture-family coverage, capacity-aware costing (F2→F2048 ladder, “incl. capacity” badges), per-region Fabric/OneLake meters, and Fabric example prompts.

#### 💰 Pricing upgrades
PAYG ↔ Reserved (1-year) toggle, a “Prices as of” stamp on exports, true per-region meters refreshable with `npm run pricing:refresh`, and corrected OneLake/Fabric rates.

#### 🔒 Security & resilience
All Azure OpenAI traffic is now **proxied server-side** (`/api/openai`) — the key never reaches the browser. Deployment guides are **grounded in Microsoft Learn** (`/api/docs-search`). A new **Help & Learn** panel and **User Feedback** round out the release; current feedback privacy and retention controls are described above.

---

### June 2026 — Blueprint Diagrams (BETA) & 12-Model Lineup

#### ✏️ Blueprint Diagrams (BETA)
The AI Generator modal now offers three **diagram modes**:
- **Topology** — the classic deployable, editable canvas diagram
- **Blueprint** *(BETA)* — a hand-drawn, whiteboard-style PNG with nested zones (Azure / VNet / On-prem) and numbered, labeled arrows tracing the end-to-end flow
- **Both** *(BETA)* — generate a deployable topology **and** a Blueprint PNG from the same prompt (optionally in parallel)

Blueprint output is delivered as a polished PNG (the PNG is the deliverable, not a canvas render) and can be re-downloaded anytime via **Export › Export Blueprint PNG**. A configurable legend position keeps results presentation-ready. Blueprint/Both modes require a general-purpose OpenAI model (GPT-5.x); the app auto-switches if a partner model is selected.

#### 🤖 Expanded to 12 AI Models
The model lineup grew from 7 to **12**, adding **GPT-5.4 Mini**, **DeepSeek V4 Pro**, **Grok 4.3**, **Mistral Large 3**, and **Kimi K2.5** alongside the existing GPT-5.1, GPT-5.2, GPT-5.2 Codex, GPT-5.3 Codex, GPT-5.4, DeepSeek V3.2 Speciale, and Grok 4.1 Fast. Every feature (generation, validation, comparison) can be assigned its own model.

#### 📤 New Export Formats
- **Interactive HTML** — self-contained page with pan, zoom, and tooltips
- **Blueprint PNG** — re-export the whiteboard-style blueprint

---

### March 14, 2026 — Workflow Avatar Narrator & Draggable/Resizable Panels

#### 🎙️ Narrate Workflow (new)
The Workflow Panel (right side of canvas) now has a **Narrate** button in its header. Click it to have a talking avatar speak every architecture step aloud:
- Narration text is built from the existing workflow steps — `"Step 1: … Step 2: … "` — no extra AI call
- Same avatar session, closed-caption, and token-server infrastructure as the Compare Models presenter
- Button is only rendered when `VITE_SPEECH_REGION` is set

#### 🖱️ Draggable & Resizable Avatar Panels
Both avatar panels (Workflow Narrator and Compare Models Presenter) are now fully interactive:
- **Drag** the panel header to reposition anywhere on the viewport
- **Resize** by dragging the diagonal-stripe handle in the bottom-right corner
- Position and size are clamped to the viewport so the panel can never be dragged off-screen
- Panel resets to its default position/size when dismissed

#### 🔧 Infrastructure
- `src/hooks/useDraggableResizable.ts` — new shared hook using pointer capture (`el.setPointerCapture`) for smooth, lag-free drag and resize; React `currentTarget` captured into locals before closures to avoid the synthetic-event nullification bug
- `pointercancel` listener added to both drag and resize handlers for clean-up on focus-loss or touch cancel

---

### March 13, 2026 — Talking Avatar Presenter

#### 🎙️ Present Critique (new)
Compare AI model critiques, then click **"Present"** to have a photorealistic **talking avatar** narrate the ranked results aloud, right in the browser:

- **Floating avatar panel** — 3D avatar appears at bottom-right inside the Compare Models modal while speaking
- **Word-by-word closed captions** — each word highlights in real time as the avatar speaks, driven by the Speech SDK `wordBoundary` event
- **Keyless authentication** — no API keys stored: `server/token-server.js` (Express.js, port 3001) runs co-located with nginx. On each `/api/speech-token` request it acquires an AAD token via `DefaultAzureCredential` and returns `aad#{resourceId}#{aadToken}` directly to the Speech SDK
- **ICE relay** — `/api/ice-token` endpoint fetches WebRTC relay credentials from Azure so avatar video works through corporate firewalls
- **Build-time feature flag** — the "Present" button is only rendered when `VITE_SPEECH_REGION` is set at image build time

#### 🔒 Server-side AI proxy
The same token server also brokers Azure OpenAI and Microsoft Foundry so credentials never reach the browser:

- **`/api/openai`** — proxies architecture generation, chat refinement, validation, and deployment-guide calls. Azure OpenAI uses `AZURE_OPENAI_ENDPOINT`; Claude Opus 5 uses the Microsoft Foundry Anthropic Messages endpoint configured by `AZURE_FOUNDRY_ENDPOINT` and the fail-closed `AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS` allowlist. Both prefer managed identity (`DefaultAzureCredential`) with optional server-side API-key fallback. The client only sends the request body, deployment name, and API format — keys are never bundled.
- **Bring your own AI endpoint (optional)** — when the explicit server kill
  switch `ALLOW_BYO_AI_ENDPOINTS=true` is enabled, users can connect a deployment
  on a trusted Azure OpenAI or Microsoft Foundry resource host, or a model on the
  official `api.openai.com` API. Azure requests use the implicit-versioned
  `/openai/v1/responses` or `/openai/v1/chat/completions` routes. The key is kept
  only in browser-tab memory and passes through `/api/openai` for the request; it
  is not written to local storage, diagrams, URLs, telemetry, or server logs.
  Reloading the tab keeps the selected public configuration but blocks AI calls
  until the user re-enters and verifies the key. The client reads
  `/api/runtime-config` so an administrator-disabled server switch is visible
  before users attempt a connection. Arbitrary OpenAI-compatible hosts are
  intentionally rejected to prevent SSRF and open-proxy abuse.
- **`/api/docs-search`** — grounds deployment guides in official Microsoft Learn documentation by calling the Microsoft Learn MCP endpoint server-side and returning citable `{title, url, excerpt}` results. Best-effort (soft-fails to empty).

#### 🔧 Infrastructure
- `server/token-server.js` — new Express.js token server started by `start.sh` before nginx
- `src/services/avatarPresenter.ts` — Speech SDK avatar session, ICE relay, word-boundary callback
- `Dockerfile` — extended build stage with `ARG/ENV VITE_SPEECH_REGION`; production stage installs token server deps
- `scripts/update_aca.sh` — adds `VITE_SPEECH_REGION` build arg and `AZURE_SPEECH_REGION` / `AZURE_SPEECH_RESOURCE_ID` runtime env vars
- ACA managed identity assigned `Cognitive Services Speech User` role on the Speech resource (no stored credentials)

For GitHub Actions deployments, set the repository variable
`ALLOW_BYO_AI_ENDPOINTS=true` only after approving BYO access for that
environment. Missing variables and all self-hosted deployment defaults remain
fail-closed (`false`).

---

### March 12, 2026 — PPTX Export & SVG Edge Rendering Fix

#### 🖼️ Export Diagram as PowerPoint Slide (new)
- **"Export PPTX Slide"** added to the Export dropdown menu
- Generates a single widescreen 16:9 `.pptx` file via **PptxGenJS** — entirely client-side, no backend required
- Slide theme automatically mirrors the current canvas mode:
  - **Dark mode** → slate-900 background, white title, Azure-blue accent bars
  - **Light mode** → slate-50 background, dark title, same accent
- Slide includes: diagram name, author, date (from the Architecture Diagram title block), the diagram image (aspect-ratio preserved), and a footer
- Export is recorded in the Recent Exports history like all other formats

#### 🔧 SVG Edge Rendering Fix — all exports (PNG, SVG, PPTX, validation snapshots)

**The problem:** ReactFlow edges (smooth, bezier, orthogonal, dashed) were invisible in all exported images.

**Root cause:** ReactFlow draws edges as SVG `<path>` elements whose `stroke` colour comes solely from the `reactflow/dist/style.css` stylesheet via the `.react-flow__edge-path` CSS class. The previous `html2canvas` library dropped SVG content almost entirely. After switching to `html-to-image`, the DOM is serialised correctly — but inside the resulting SVG `<foreignObject>`, the page's external stylesheets are no longer in scope, so every path renders with no stroke (invisible).

**Fix — `src/utils/captureCanvas.ts`:**
1. `html2canvas` replaced with `html-to-image` across all four capture call sites (PNG export, SVG export, PPTX export, validation snapshot)
2. A new `prepareEdgesForCapture()` helper runs synchronously before every capture. It iterates every `svg path/line/polyline/circle` inside the ReactFlow wrapper, reads each element's **computed CSS** via `window.getComputedStyle()`, and writes the results back as **SVG presentation attributes** (`stroke`, `stroke-width`, `stroke-dasharray`, `fill`, `opacity`, `marker-end`, etc.) directly on the element. Presentation attributes survive serialisation regardless of whether stylesheets are present
3. After capture completes (or throws), all attributes are restored to their original values so the live canvas is unaffected
4. Transparent fills are normalised to `none` (SVG convention) to avoid invisible filled areas

All edge types now render correctly: solid sync edges, dashed async edges, dotted optional edges, animated directional-flow edges, and bidirectional pulse edges.

### February 28, 2026 — Multi-Model Expansion & Comparison
- **7-Model Support** — Added GPT-5.1, GPT-5.3 Codex, GPT-5.4, DeepSeek V3.2 Speciale, and Grok 4.1 Fast alongside existing GPT-5.2 and GPT-5.2 Codex
- **Chat Completions API Adapter** — Dual API support: Responses API for GPT models, Chat Completions API for third-party models (DeepSeek, Grok)
- **Multi-Model Validation Comparison** — Compare WAF validation results across all 7 models with score, pillar, severity, and finding breakdowns
- **Collapse All Groups** — Toggle button to collapse/expand all groups for bird's-eye view, with size persistence
- **ARM Parsing Banner** — Glowing purple gradient banner during ARM template parsing
- **Seven-Model Comparison Report** — Formal analysis document ranking all 7 models across 4 prompts
- **Save All Diagrams** — Download each model's architecture comparison result as individual JSON files
- **Save Comparison Report** — Download combined comparison results as a single JSON for offline analysis
- **Shared Group Utilities** — Extracted `fitGroupToContent` into reusable `groupUtils.ts`
- **Bug Fixes** — Fixed Grok 404/string-groups crash, DeepSeek circular parent crash, `kbStats.serviceCount` typo

### February 14, 2026 — UI Polish, Auth & Deployment
- **Entra ID Authentication** — ACA built-in auth with per-user assignment (no code changes needed)
- **Configurable Deploy Script** — New `scripts/deploy_aca.sh` reads all config from `.env` — clone, configure, deploy
- **GPT-5.2 Codex Deployment Support** — Added to Dockerfile and deploy pipeline
- **Compare Models Button Styling** — Amber gradient with pulse animation, dark mode compatible
- **Remove Share Feature** — Removed broken Share button, Express server, and Cosmos DB backend
- **Categorized AI Prompts** — 6 color-coded categories (Web, Security, IoT, AI, E-commerce, Healthcare) replacing flat list
- **Dark Mode Improvements** — Full dark mode support for AI modal, image uploader, form elements
- **Auto-Collapse Panels** — Icon palette, workflow panel, and legend collapse after AI generation
- **Start Fresh Button** — One-click reset with confirmation to clear entire diagram state
- **Compare Models Verbose Prompts** — 8 sample prompts (4 concise + 4 detailed enterprise scenarios)
- **Form UX Improvements** — Textarea above image upload, purple/blue fill colors, improved labels
- **Dockerfile Optimized** — Switched from Node.js server to nginx:alpine static serving
- **Sidebar Search Fix** — Icon search now works across all categories
- **Azure Backup Icon Fix** — Corrected category mapping
- **Stream Analytics Alias Fix** — Corrected icon resolution
- **Power BI Embedded Pricing** — Added missing pricing data
- **Azure Functions & Stream Analytics Pricing** — Added regional pricing data
- **Dashboard Services** — Added Power BI, Grafana, Azure Dashboard to icon mapping

### February 2026 — Core Features
- **Architecture Image Import** — Upload diagram images for AI-powered recreation
- **Workflow Animation Panel** — Step-by-step data flow visualization with service highlighting
- **Multi-Model Support** — GPT-5.1, GPT-5.2, GPT-5.2 Codex, GPT-5.3 Codex, DeepSeek V3.2 Speciale, Grok 4.1 Fast with per-feature overrides
- **Dual API Support** — Responses API for GPT models, Chat Completions API for third-party models
- **Model Selector UI** — Toolbar dropdown with reasoning effort configuration
- **Model Comparison** — Side-by-side architecture and validation comparison across all models
- **Bicep Templates** — IaC generation in deployment guides
- **Reasoning Effort** — Configurable AI thinking depth (GPT-5.x: none/low/medium/high)
- **Smart Layout Engine** — Dagre-based auto-layout with group overlap resolution
- **ELK.js Layout Engine** — Alternative layout with toggle
- **Microsoft Logo** — Added to header banner
- **12 AI Layout Rules** — Directional flow, hub-and-spoke, connection caps, cross-group edge minimization
- **Auto-Snapshot** — Automatic version save before AI regeneration
- **13 Curated Example Prompts** — Security, healthcare, gaming, e-commerce, IoT, AI services
- **68 Mapped Azure Services** — Full icon resolution, categorization, and pricing
- **Resizable Group Nodes** — Drag handles to adjust group boundaries
- **Iterative Regeneration** — Regenerate with selected WAF improvements applied
- **Security-Focused Prompts** — Zero Trust, SOC, and enterprise security scenarios
- **ModelBadge** — Shows which AI model generated the current diagram
- **Chat Completions Fallback** — Automatic fallback for models not supporting Responses API
- **Two-Row Toolbar** — Split toolbar for better fit on normal-width windows

### March 2026
- **8 Azure Regions** — Expanded from 5 to 8 regions: added Australia East (HERO), Southeast Asia (Singapore), Mexico Central (Querétaro) with full pricing data (568 JSON files)
- **Export Costs (All Formats)** — One-click ZIP: CSV + JSON + Markdown summary + intelligent Markdown analysis report (with Mermaid pie chart)
- **Multi-Region Cost Comparison** — Ranked table across all 8 regions in the analysis report, with cheapest/priciest callouts, potential savings estimate, and per-service regional variance
- **Draggable & Resizable Reference Image** — Reference sketch panel can now be dragged anywhere and resized via a corner handle
- **Legacy azd template retired in secured fork** — production deployment now uses the protected Front Door and Easy Auth workflow only

### January 2026
- **WAF Validation** — Well-Architected Framework checks across all 5 pillars
- **Iterative Improvement** — Select and apply WAF recommendations
- **Version History** — Named snapshots with time travel
- **Draw.io Export** — Edit in diagrams.net

---

## 🤝 Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately. Simply follow the instructions provided by the bot.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

See [CONTRIBUTING.md](CONTRIBUTING.md) for detailed guidelines.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).

This project uses the official Microsoft Azure icon library. Please refer to [Microsoft's usage guidelines](https://docs.microsoft.com/en-us/azure/architecture/icons/) for the icons.

---

## ™️ Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft
trademarks or logos is subject to and must follow
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.

---

<div align="center">

**Built with ❤️ for the Azure community**

*Empowering cloud architects to design better solutions faster*

</div>
