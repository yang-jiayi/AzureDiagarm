# Ignite 2026 アイコン追加レビュー

## 公式発表と素材の確認

[Microsoft Ignite 2026 の Azure 発表](https://azure.microsoft.com/en-us/blog/microsoft-ignite-2026-whats-new-in-azure-the-platform-purpose-built-for-ai/)
（2026年11月17日公開）では、Fabric Apps、Microsoft IQ、Azure HorizonDB、
Fabric HorizonDB が発表された。Microsoft IQ は Work IQ、Fabric IQ、Foundry IQ、
Web IQ などを束ねる enterprise context management platform と説明されている。

[Microsoft IQ の公式ページ](https://www.microsoft.com/ai/microsoft-iq)には各IQの
48pxの画像が掲載されているが、確認できたのは製品ページ上のマーケティング画像で、
図解アイコンとして再配布・利用できる旨の条件は見当たらなかった。これらの画像や
他製品のロゴはコピーしていない。公式発表は製品名を確定する情報源としてのみ記録し、
Microsoft IQ、Work IQ、Fabric IQ、Web IQには、製品ロゴではない
ローカル作成の汎用概念記号を個別に作成した。Foundry IQはAzure V25に
公式の `Azure AI Foundry IQ` アイコンがあるため、その同一製品のアイコンを再利用し、
パレット名を現行のFoundry IQにする。Azure AI Searchの既存サービスキーとアイコンは
互換性のため保持する。4つの代替記号はパレットに「非公式の概念記号
(製品ロゴではありません)」と表示する。ファイルとSHA-256は
`ignite-manifest.json`、入力台帳は `../scripts/ignite-icon-sources.json` に記録する。

Microsoft Fabric の[公式アイコン案内](https://learn.microsoft.com/en-us/fabric/fundamentals/icons)
は、公式アイコンをarchitecture diagrams、training、slides、documentationに使用
できるとしている。Fabric Apps専用の `fabric_apps_48_item.svg` は既存の
`@fabric-msft/svg-icons` 12.0.0 inventory に収録済みのため再取得・重複追加はせず、
現行パレット・service map・MCP/export mappingで再利用する。Power BI Appsの既存
`apps_48_item.svg`、LegacyのApp Development、Org Appは別IDのまま扱い、Fabric Appsの
別名に統合しない。

Azure HorizonDBはAzure V25公式アーカイブの専用アイコンをサービスマップに登録した。
Fabric HorizonDBには12.0.0公式Fabric packageの `postgre_sql_database_48_item.svg`
を汎用PostgreSQL記号として再利用し、専用HorizonDBロゴではない旨をパレットに表示する。
Microsoft IQ系5項目とFabric HorizonDBには確認できる料金メーターを結び付けず、
価格・Fabric capacity消費を推測しない。

## 追加・再利用・見送り

| 対象 | 対応 |
| --- | --- |
| Fabric Apps | 公式12.0.0 iconを再利用。既存のservice key、Fabric capacity意味、Power BI Appsとの差を維持 |
| Azure HorizonDB | Azure V25の製品専用公式アイコンを新しいサービス名で利用 |
| Fabric HorizonDB | Fabric PostgreSQLの公式汎用記号を再利用。専用ロゴだとは表示しない |
| Microsoft IQ、Work IQ、Fabric IQ、Web IQ | 製品ページ画像は再配布条件を確認できないため不使用。互いに異なるローカル作成の汎用記号を「非公式」と明記して追加 |
| Foundry IQ | Azure V25に同一製品の公式 `Azure AI Foundry IQ` アイコンがあるため再利用。新しい別記号は追加しない |
| Microsoft IQのマーケティング画像、公式と確認できない個別IQロゴ | 見送り。製品ページ画像の利用条件を図解アイコン利用許諾とはみなさない |

```text
npm run icons:sync:ignite
npm run icons:check:ignite
```

ローカルの概念記号は同梱SVGの安全性とSHA-256をmanifestに照合する。Microsoft公式
配布アセットを示すものではなく、製品ロゴとして扱ってはならない。

## 現行main統合後のUI・動作確認

現行 `origin/main` の `a2581b58ce4b65c5390e5778d86d536f3645a825` を基点に統合した。
仮想化の行高（grid 128px、touch grid 164px、list 72px）は変更していない。
4つの非公式IQ概念記号は、英語 `Unofficial` / 日本語 `非公式` の短いバッジを
常時表示し、製品ロゴでないことの全文はtooltipと `aria-describedby` で伝える。
Gridではアイコン、名称、バッジを別行に配置してカード内に収め、Listでは横並びを保つ。

実ブラウザのPlaywright確認では、英語・日本語それぞれのGrid/Listでバッジ、名称、
アイコン、操作ボタンが重ならずカード内に収まることを確認した。検索からMicrosoft IQ、
Fabric Apps、Azure HorizonDB、Fabric HorizonDB（汎用PostgreSQL概念記号）を実際に
キャンバスへ追加し、PNGをダウンロードできることも確認した。画面キャプチャは
セッション成果物 `files/ignite-iq-palette.png` に保存した。

| 検証コマンド | 結果 |
| --- | --- |
| `npm run test:e2e -- tests/e2e/ignite-icon-palette.spec.ts` | 成功、2 tests passed。英日×Grid/List、検索・追加、PNG exportを確認 |
| `npx tsx --test tests/igniteIcons.test.ts tests/fabricIcons.test.ts` | 成功、9 tests passed |
| `npm run mcp:build` | 成功。現行カタログから412件を生成、405 SVGを埋め込み、欠落0件 |
| `npm run icons:check` | 成功。Azure V25 721件、Fabric 117件、Entra 14件、Microsoft 229件、Ignite hash checkを確認 |
| `npm run build` | 成功 |
| `npm run lint` | 成功 |
