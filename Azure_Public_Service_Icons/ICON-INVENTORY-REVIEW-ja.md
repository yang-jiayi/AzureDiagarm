# Microsoft アイコン網羅更新レビュー

## 変更概要

公式アーカイブを完全走査し、Azure V25、Fabric 12.0.0、Microsoft Entra
October 2023 の現行資産を統合した。出典URLとパッケージSHA-256を固定し、
全SVGのハッシュを各manifestへ記録した。同期・checkは固定ピンに対して実行し、
ネットワーク上の「latest」を自動採用しない。

| 配布元 | 旧 → 新 | 比較結果 |
| --- | --- | --- |
| Azure Architecture Icons | V24: 714 → V25: 721 | 新規サービス7件、Microsoft Discoveryのカテゴリ移動1件（内容も更新）、Translator Textの内容更新1件、削除0件。V25 ZIP SHA-256: `f293d99252b3638557dd08ccbff9f2529083473d39b373fb2af29a7200351339` |
| Fabric architecture package | 8.2.0: 82 → 12.0.0: 88 | 新規7件、User Data Functionのファイル名置換1件、Function Set asset削除1件、既存名のSVG更新2件、同名・同hash 78件。12.0.0公開日時: 2026-11-05 03:02:36.505 UTC |
| Microsoft Entra architecture icons | October 2023: 14件 | 追加2件、既存サービスアイコンに対応付け9件、白黒variant 3件は色付きcounterpartと重複するため追加せず。ZIP SHA-256: `4e07536706a2d092e6524e5417e2c861333fdbdb41c36f78d44dfa07ccc5eedc` |
| Microsoft 365 architecture symbols | April 2024 archive | 963 SVG、6つの配色・テーマ別フォルダーを照合。アプリ専用ロゴSVGとは確認できず、製品ロゴとして追加したものは0件。現行205項目は概念シンボルとして保持。 |

Azure V25で追加した7サービスは Azure Confidential Clean Rooms、ACA
Sandboxes、Azure HorizonDB、Azure Encrypted Transport、AI Manager、App
Service Builder、Azure Device Registry。Microsoft Discoveryは製品新設ではなく
`other/` から `new icons/` への移動である。既存の保存図で使われるV24の旧
Microsoft Discoveryパスも互換用に残した。

Fabric 12.0.0の7新規ファイルは `action_48_item.svg`、
`database_hub_48_color.svg`、`fabric_apps_48_item.svg`、
`feature_store_48_item.svg`、`managed_identity_48_item.svg`、
`postgre_sql_database_48_item.svg`、`rayfin_48_color.svg`。名称置換は
`user_data_function_48_item.svg` → `user_data_functions_48_item.svg`。
`function_set_48_item.svg` は12.0.0から削除されたが、8.2.0由来のLegacy画像、
旧サービス名解決、容量消費意味を維持した。V12で内容が更新された同名SVGは
CopilotとMirrored Catalog。Sample Workloadは公式packageに48px assetがないため
32pxのままとした。

Sam Debruynの2026年コレクションに含まれる101 SVGはすべてSHA-256を照合した。
99件はactive catalog path上でも元記事と同一である。現行12.0.0の絵柄に更新した
CopilotとUser Data Functionの元画像は
`Azure_Public_Service_Icons/Article_Fabric_Icons/2026/` に別保存し、記事の
コンパイル著者とMicrosoftの図版権利を区別して帰属を記載した。

Entraの新規画像は Entra Product Family と Entra Workload ID の2件のみ。
Identity Governance、Internet Access、Private Access、Verified IDなどは既存
Azureアイコンとの用途重複として再追加せず、白黒variantは対応する色付き画像を
使う。全14件のsourcePath、分類、source hash、追加先または重複先は
`entra-manifest.json` に記録した。

## 製品名付きMicrosoft 365検索とロゴの限界

公式Microsoft 365配布元はアーカイブ済みのarchitecture icon/symbol setであり、
製品ロゴ集ではない。配色フォルダー（例: “Teams Purple”, “Planner Green”）に
ある同名ファイルも汎用の人物・会話・予定表・作業アイコンで、TeamsやPlannerの
製品ロゴではない。許諾範囲の明確な公式M365アプリロゴ配布元をこの調査では確認
できなかったため、非公式画像や創作ロゴは追加していない。

したがって次の製品名は、現在も専用ロゴではなく、Microsoft 365カテゴリ内の
汎用「概念シンボル」への検索alias／既存サービスアイコンターゲットとして残る。
paletteに「概念シンボル (製品ロゴではありません)」と明示する。従来の
AI/MCPサービス名と別名は維持し、pricing dataは付与しない。

- Teams / Teams Chat・Phone・Meeting・Webinar・Shifts、SharePoint、OneDrive、
  Exchange Online、Outlook、Bookings
- Microsoft 365 Apps、Word、Excel、PowerPoint、OneNote、Loop、Forms、Stream、
  Planner、To Do、Project、Lists、Clipchamp、Whiteboard、Sway
- Copilot、Copilot Studio、Viva Insights・Learning・Connections・Engage・Amplify
- Microsoft Purview、Compliance Manager、Retention、Information Protection、
  Microsoft Defender XDR、Defender for Office 365、Microsoft Intune
- Microsoft Entra ID、Connect、Directory、Groups、Administrative Unit、
  Microsoft 365 Admin Center、Microsoft Search、Microsoft Graph
- Windows 365

Microsoft 365のシンボルとPower Platform / Dynamics 365の専用製品アイコンは
別種として表示・分類する。

## 意味・互換性

- FabricのAction、Database Hub、Feature Store、Managed Identity、PostgreSQL
  Database、Rayfinは描画専用conceptであり、AI/MCP生成やcapacity課金対象にしない。
  Rayfinの製品名・用途は未確認のため「未分類シンボル」と表示し、これらに
  未確認の料金区分・cost rangeを付与しない。
- Fabric Appsは製品itemとして表示し、容量を消費する。Power BI Appの旧パスと
  サービス解決は別に維持する。
- Function Setは旧図解決用の互換mapと新規サービス推薦を分け、容量消費の従来意味を
  変えない。
- Entra追加マッピング、Azureの既存サービス名、Fabric/M365の既存サービスキー、
  別名、pricingデータ、保存済みパスを維持する。
- `ICON-LICENSING.md` と `FABRIC-ICONS.md` に出典、利用条件、帰属、同期方法を記載。

## manifestと再現コマンド

- `manifest.json`: Azure V25全721 SVG、カテゴリ、archive bytesとLF正規化後の
  SHA-256、V24比較結果。
- `fabric-manifest.json`: Fabric全117 palette entry、現行88 asset比較、article 101
  件、各SVGのハッシュ。
- `entra-manifest.json`: Entra全14 SVGの対応状況とハッシュ。
- `microsoft-product-manifest.json`: Power Platform 8、Dynamics 365 16、
  Microsoft 365 symbol 205、合計229件。

```bash
npm run icons:sync
npm run icons:check
node mcp-server/scripts/sync-icon-map.mjs
npx tsx --test tests/iconSourceInventory.test.ts tests/fabricIcons.test.ts tests/fabricIconSync.test.ts tests/microsoftProductIcons.test.ts
node scripts/test-fabric-icons.mjs
npm run lint
npm run build
npm run mcp:build
```

## 検証結果

| コマンド | 結果 |
| --- | --- |
| `npm run icons:sync` | 成功。Azure、Fabric、Entra、Microsoft製品アイコンを固定pinから同期。既存アセット変更なし。 |
| `npm run icons:check` | 成功。Azure V25 721件、Fabric palette 117件（現行88・Legacy15・UI13・local1）、Entra 14件、Microsoft製品229件を照合。 |
| `npm run test:icons` | 成功。Azure 721、Fabric 117とarticle 101、Microsoft製品229の全SVGを検証。 |
| `npx tsx --test tests/iconSourceInventory.test.ts tests/fabricIcons.test.ts tests/fabricIconSync.test.ts tests/microsoftProductIcons.test.ts` | 成功。25/25件。全件網羅、旧パス・Function Setの意味、conceptのAI/MCP/capacity/pricing除外、M365のsymbol分類を確認。 |
| `npm run test:unit` | 成功。1246/1246件。初回の`npm ci --ignore-scripts`で省略されたリポジトリ標準postinstall（image-size parser safeguards）を適用後に実行。 |
| `node scripts/test-fabric-icons.mjs` | 成功。本番Vite build相当でFabric 117 SVG（data URL 100、native raster UI 13）、保存済みパス、lazy-load中のscroll維持、描画分類、16 canvas nodes、PNG exportとM365概念シンボル表示を確認。 |
| `npm run lint` / `npm run build` / `npm run mcp:build` | すべて成功。MCP buildはicon map 405件、embedded SVG 398件（missing 0）を生成。 |
| `node scripts/emulate-docker-context.mjs <temp path>` | 成功。Docker context 2227ファイルに4 manifestsと新規SVGが含まれることを確認し、temp出力は削除。 |

Azure V25の更新SVGは配布元archiveのhashとLF正規化hashを別々にmanifestへ記録し、
Gitのline-ending変換後も再現可能な照合にした。

## 本番反映状態

Azureリソース、権限、設定、revision、環境変数は変更していない。本番反映は、
保護された通常のPR/merge経路とproduction Environmentのユーザー承認後にのみ行う。
承認や必須チェックの迂回、手動workflow dispatchは行わない。
