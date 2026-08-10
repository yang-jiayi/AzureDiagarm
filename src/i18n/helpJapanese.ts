// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Japanese translations for the Help & Learn onboarding surface (GuidedHelpPanel)
// and other help-related copy. English source strings are the keys; every value is
// a natural, technical translation in です・ます調 with half-width spacing around
// Latin/Azure product names and Azure service names kept in their official English form.
//
// This module is the single source of truth for these translations. GuidedHelpPanel
// imports it directly so the panel renders correctly regardless of the shared
// dictionary, and it is also spread into `exactJapanese` in LanguageContext so the
// same keys resolve through `t()` / `translate()` app-wide.

export const helpJapanese: Record<string, string> = {
  // --- Panel chrome ---
  'Help and Learn': 'ヘルプと学習',
  'Help & Learn': 'ヘルプと学習',
  'From first prompt to architecture handoff': '最初のプロンプトからアーキテクチャの引き継ぎまで',
  'Close': '閉じる',
  'Close help': 'ヘルプを閉じる',
  'Help sections': 'ヘルプのセクション',
  'First tour': '最初のツアー',

  // --- Section navigation labels ---
  'Start Here': 'ここから開始',
  'Choose a Path': '進め方を選ぶ',
  'Create & Refine': '作成と調整',
  'Assess': '評価',
  'Deliver': '共有・引き継ぎ',
  'Prompt Lab': 'プロンプト ラボ',
  'FAQ & Safety': 'FAQ と安全な利用',

  // --- Quick start ---
  'Your first 10 minutes': '最初の10分',
  'Build confidence by completing one full loop': 'ひと通りの流れを完了して操作に慣れましょう',
  'Create, correct, validate, inspect cost, and export. Mark each task as you try it—the checklist is saved in this browser.': '作成、修正、検証、コストの確認、エクスポートを行います。試した項目にチェックを付けてください。チェックリストはこのブラウザーに保存されます。',
  '1. Create': '1. 作成',
  'Choose Guided Chat, Generate Diagram, or Import Existing.': 'ガイド付きチャット、図を生成、既存のものをインポートのいずれかを選択します。',
  '2. Refine': '2. 調整',
  'Continue in Guided Chat or edit nodes, groups, and connections directly.': 'ガイド付きチャットで続けるか、ノード、グループ、接続を直接編集します。',
  '3. Validate & Improve': '3. 検証と改善',
  'Review findings, apply selected changes, and revalidate when the design changes.': '指摘事項を確認し、選択した変更を適用し、設計が変わったら再検証します。',
  '4. Share or Build': '4. 共有またはビルド',
  'Export a review artifact or create deployment guidance when the design is ready.': '設計が固まったら、レビュー用の成果物をエクスポートするか、デプロイ手順を作成します。',
  'Important:': '重要:',
  'AI output is a starting hypothesis. Correct it with domain experts, validate assumptions, and review generated deployment content before use.': 'AI の出力は出発点となる仮説です。利用前に、専門家とともに内容を修正し、前提を検証し、生成されたデプロイ内容を確認してください。',
  'Choose how you want to start': '開始方法を選ぶ',

  // --- Choose a path ---
  'Choose a starting point': '開始地点を選ぶ',
  'You do not need to begin with a perfect prompt': '完璧なプロンプトから始める必要はありません',
  'Start from an idea': 'アイデアから始める',
  'Recommended for first-time users': '初めての方におすすめ',
  'Open Guided Chat': 'ガイド付きチャットを開く',
  'Describe the outcome and constraints': '成果と制約を説明する',
  'Create and refine in one conversation': '一つの会話で作成と調整を行う',
  'Start from a brief or sketch': '要件やスケッチから始める',
  'Detailed requirements or image': '詳細な要件または画像',
  'Open Generate Diagram': '図の生成を開く',
  'Choose Topology, Blueprint, or both': 'Topology、Blueprint、またはその両方を選ぶ',
  'Review, then continue in Guided Chat or on canvas': '確認後、ガイド付きチャットまたはキャンバスで続ける',
  'Start from current estate': '既存の環境から始める',
  'IaC or live Azure': 'IaC または稼働中の Azure',
  'Import Bicep, Terraform, or ARM': 'Bicep、Terraform、または ARM をインポートする',
  'Or use Import from Azure': 'または Azure からのインポートを使う',
  'Correct inferred relationships on canvas': '推定された関係をキャンバス上で修正する',
  'Prepare a review or workshop': 'レビューやワークショップを準備する',
  'Customer-ready flow': '顧客向けの流れ',
  'Create and correct the concept': 'コンセプトを作成して修正する',
  'Validate and improve iteratively': '検証と改善を繰り返す',
  'Share a review artifact or build deployment guidance': 'レビュー用の成果物を共有するか、デプロイ手順を作成する',

  // --- Create & refine ---
  'Create and refine': '作成と調整',
  'Move between AI and direct canvas editing': 'AI とキャンバスの直接編集を行き来する',
  'Use AI for acceleration, then use the canvas for precision. Targeted follow-up requests preserve your existing manual layout.': 'AI で作業を加速し、キャンバスで細部を仕上げます。対象を絞った追加依頼では、既存の手動レイアウトが維持されます。',
  'Guided Chat': 'ガイド付きチャット',
  'Best for conversational creation and ongoing refinement. Build from empty or change the current canvas in one thread; existing manual positions are retained during modifications.': '会話形式での作成や継続的な調整に最適です。空の状態から作成することも、現在のキャンバスを一つのスレッドで変更することもできます。変更中も既存の手動配置は維持されます。',
  'Generate Diagram': '図を生成',
  'Best when you have detailed requirements, an existing sketch, or need explicit output controls. Choose Topology, Blueprint, or Both, then hand off to Guided Chat or canvas review.': '詳細な要件や既存のスケッチがある場合、または出力を明示的に制御したい場合に最適です。Topology、Blueprint、Both のいずれかを選び、その後ガイド付きチャットやキャンバスでの確認に引き継ぎます。',
  'Import': 'インポート',
  'Reconstruct a diagram image, parse Bicep/Terraform/ARM, or sign in to reverse-engineer a live Azure resource group.': '図の画像を再構成したり、Bicep/Terraform/ARM を解析したり、サインインして稼働中の Azure リソース グループをリバース エンジニアリングしたりできます。',
  'Compare Models': 'モデル比較',
  'Run one prompt across several models, inspect latency/tokens/topology differences, and apply the result you prefer.': '一つのプロンプトを複数のモデルで実行し、待機時間・トークン・トポロジの違いを確認して、好みの結果を適用します。',
  'Edit on canvas': 'キャンバスで編集',
  'Drag services, resize groups, edit labels, reconnect edges, align selections, and choose a layout preset or edge style.': 'サービスのドラッグ、グループのサイズ変更、ラベルの編集、エッジの再接続、選択範囲の整列ができ、レイアウト プリセットやエッジ スタイルを選べます。',
  'Version History': 'バージョン履歴',
  'A snapshot is saved before AI regeneration. Save named checkpoints and restore prior versions when an experiment does not work.': 'AI による再生成の前にスナップショットが保存されます。名前付きのチェックポイントを保存でき、試行がうまくいかない場合は以前のバージョンを復元できます。',

  // --- Assess ---
  'Turn a diagram into a review conversation': '図をレビューのための対話につなげる',
  'Validation and cost are decision aids. They expose assumptions and tradeoffs; they do not replace sizing, security review, or architecture approval.': '検証とコストは意思決定を助けるものです。前提やトレードオフを明らかにしますが、サイジング、セキュリティ レビュー、アーキテクチャの承認に代わるものではありません。',
  'Well-Architected validation': 'Well-Architected 検証',
  'Review Cost Optimization, Operational Excellence, Performance Efficiency, Reliability, and Security. Apply selected recommendations, review the resulting iteration, and revalidate after material changes.': 'コスト最適化、オペレーショナル エクセレンス、パフォーマンス効率、信頼性、セキュリティを確認します。選択した推奨事項を適用し、生成された反復結果を確認して、大きな変更の後に再検証します。',
  'Compare Validation': '検証比較',
  'Ask multiple models to review the same architecture, compare findings, and use consensus to separate recurring gaps from model-specific opinions.': '複数のモデルに同じアーキテクチャをレビューさせ、指摘事項を比較します。一致度を手がかりに、繰り返し現れる課題とモデル固有の意見を切り分けます。',
  'Cost and region': 'コストとリージョン',
  'Inspect per-service monthly estimates across eight regions and switch between PAYG and 1-year savings. Usage-based values remain indicative.': '8 つのリージョンでサービスごとの月額見積もりを確認し、従量課金と 1 年間の割引を切り替えられます。使用量ベースの値はあくまで目安です。',
  'Validation timing': '検証のタイミング',
  'Review and refine the generated concept first, then use the Validate & Improve journey stage before sharing or building. Revalidate after material changes.': 'まず生成されたコンセプトを確認・調整し、共有やビルドの前に「検証と改善」の段階を使用します。大きな変更の後は再検証してください。',

  // --- Deliver ---
  'Choose an output for the next person': '次の担当者に合わせて出力を選ぶ',
  'Export based on what the recipient needs to do next: present, edit, review, estimate, or continue implementation planning.': '受け取る人が次に行うこと（プレゼン、編集、レビュー、見積もり、実装計画の継続）に応じてエクスポート形式を選びます。',
  'Editable formats': '編集可能な形式',
  'Use Visio (VSDX), Draw.io, JSON, or interactive HTML when another person needs to continue editing.': '別の担当者が編集を続ける場合は、Visio (VSDX)、Draw.io、JSON、またはインタラクティブな HTML を使用します。',
  'Presentation formats': 'プレゼンテーション形式',
  'Export PNG, SVG, a PowerPoint slide, or a customer deck. Use Export background to choose Plain (recommended), Dots, or Grid without changing the editing canvas.': 'PNG、SVG、PowerPoint スライド、または顧客向け資料をエクスポートします。「Export background」で、編集用キャンバスを変えずに Plain（推奨）、Dots、Grid を選べます。',
  'Workflow outputs': 'ワークフロー出力',
  'Export a Markdown narrative or animated workflow, and use Narrate when the Speech presenter is available.': 'Markdown の説明文やアニメーション化したワークフローをエクスポートでき、音声プレゼンターが利用できる場合は「Narrate」を使用します。',
  'Deployment Guide': 'デプロイ ガイド',
  'Generate a Microsoft Learn-grounded runbook and Bicep starters. Review all commands, sizing, identities, and safeguards before deployment.': 'Microsoft Learn に基づく手順書と Bicep のスターターを生成します。デプロイ前に、すべてのコマンド、サイジング、ID、安全策を確認してください。',
  'Cost package': 'コスト パッケージ',
  'Download CSV or the all-formats ZIP with summaries, analysis, JSON, and multi-region comparison.': 'CSV、または概要・分析・JSON・複数リージョン比較を含む全形式の ZIP をダウンロードします。',
  'Demo mode': 'デモ モード',
  'Use Focus, Hide Toolbar, collapse groups, mini-map navigation, and Fit to view to present large diagrams clearly.': 'フォーカス、ツールバーの非表示、グループの折りたたみ、ミニマップ操作、全体表示を使って、大きな図を見やすく提示します。',

  // --- Prompt Lab ---
  'Describe intent and constraints—not a shopping list': '欲しいものの一覧ではなく、目的と制約を説明する',
  'A useful prompt names the outcome, users, data, existing investments, and non-functional constraints. You can leave unknowns explicit.': '有用なプロンプトでは、成果、ユーザー、データ、既存資産、非機能要件を明示します。不明な点は不明のままはっきり記載してかまいません。',
  'Copied': 'コピーしました',
  'Copy template': 'テンプレートをコピー',
  'Quick examples': 'クイック例',
  'Best follow-ups are specific: “keep existing positions,” “use private endpoints for data services,” “show a pilot under $500/month,” or “replace App Service with Container Apps.”': '効果的な追加依頼は具体的です。例:「既存の配置を保持する」「データ サービスにプライベート エンドポイントを使う」「月額 500 ドル未満のパイロットを示す」「App Service を Container Apps に置き換える」。',

  // --- FAQ & responsible use ---
  'FAQ and responsible use': 'FAQ と責任ある利用',
  'Know what the tool does—and what still needs review': 'ツールができることと、レビューが必要なことを理解する',
  'Which model should I use?': 'どのモデルを使えばよいですか？',
  'Use the selected default for most work. Compare models when the architecture is consequential or outputs vary. Higher reasoning can improve complex designs but usually takes longer.': 'ほとんどの作業では選択済みの既定モデルを使用します。アーキテクチャの重要度が高い場合や出力にばらつきがある場合は、モデルを比較してください。推論を強めると複雑な設計の品質が向上することがありますが、通常は時間が長くなります。',
  'What is the difference between Guided Chat and Generate Diagram?': 'ガイド付きチャットと図の生成の違いは何ですか？',
  'Guided Chat is best for conversational creation and repeated refinement. Generate Diagram is best for detailed prompts, uploaded sketches, model selection, and choosing Topology or Blueprint output. Both create an editable result and can continue in Guided Chat.': 'ガイド付きチャットは会話形式の作成や繰り返しの改良に最適です。図の生成は、詳細なプロンプト、アップロードしたスケッチ、モデルの選択、トポロジまたはブループリント出力の選択に最適です。どちらも編集可能な結果を作成し、その後ガイド付きチャットで続けることができます。',
  'How do I remove the dots from an export?': 'エクスポートからドットを消すには？',
  'Open Export and set Export background to Plain (recommended). You can also choose Dots or Grid. This affects visual exports only; the editing canvas remains dotted.': 'エクスポートを開き、「Export background」を Plain（推奨）に設定します。Dots や Grid も選べます。これは表示用のエクスポートにのみ影響し、編集用キャンバスはドット表示のままです。',
  'How do I correct an AI result?': 'AI の結果を修正するには？',
  'Use Chat for a targeted change, then edit directly on canvas. Existing positions are preserved during refinements. Version History lets you restore an earlier state.': 'チャットで対象を絞って変更し、その後キャンバス上で直接編集します。調整の間も既存の配置は維持されます。バージョン履歴を使うと、以前の状態を復元できます。',
  'Can I import existing infrastructure?': '既存のインフラストラクチャをインポートできますか？',
  'Yes. Import Bicep, Terraform, ARM, an architecture image, or a live Azure resource group. Review inferred connections and unsupported resources.': 'はい。Bicep、Terraform、ARM、アーキテクチャ画像、稼働中の Azure リソース グループをインポートできます。推定された接続とサポート対象外のリソースを確認してください。',
  'Are the costs authoritative?': 'コストは確定値ですか？',
  'No. They are indicative catalog-based estimates. Confirm SKU, quantity, usage, discounts, networking, support, and regional availability in the Azure Pricing Calculator.': 'いいえ。カタログに基づく目安の見積もりです。SKU、数量、使用量、割引、ネットワーク、サポート、リージョンでの提供状況を Azure 料金計算ツールで確認してください。',
  'Does a WAF score approve the design?': 'WAF スコアで設計が承認されますか？',
  'No. It is a structured review aid based on visible topology and model context. Validate findings with architects, security, operations, and workload owners.': 'いいえ。表示されているトポロジとモデルのコンテキストに基づく、体系的なレビュー補助にすぎません。指摘事項は、アーキテクト、セキュリティ、運用、ワークロード所有者とともに検証してください。',
  'Can I deploy the generated Bicep directly?': '生成された Bicep をそのままデプロイできますか？',
  'Treat it as starter IaC. Review API versions, identities, network controls, naming, policy, sizing, dependencies, and destructive operations before deployment.': 'スターターの IaC として扱ってください。デプロイ前に、API バージョン、ID、ネットワーク制御、命名、ポリシー、サイジング、依存関係、破壊的な操作を確認してください。',
  'What information should I avoid entering?': '入力を避けるべき情報は何ですか？',
  'Do not enter passwords, keys, tokens, regulated personal data, confidential customer content, or production data unless your organization has explicitly approved that use.': '組織が明示的に承認していない限り、パスワード、キー、トークン、規制対象の個人データ、顧客の機密コンテンツ、本番データを入力しないでください。',
  'Trusted references': '信頼できる参考資料',
  'From Prompt to Production—AADB overview': 'From Prompt to Production — AADB の概要',
  'Azure Well-Architected Framework': 'Azure Well-Architected Framework',
  'Azure Pricing Calculator': 'Azure 料金計算ツール',
  'Azure Architecture Center': 'Azure アーキテクチャ センター',
  'Still stuck or found something wrong? Close Help and use the Feedback button in the lower-right corner.': '解決しない場合や問題を見つけた場合は、ヘルプを閉じて、右下の「フィードバック」ボタンをご利用ください。',

  // --- Example prompts ---
  'Internal RAG assistant grounded on SharePoint and policy documents, available in Teams, secured with Entra ID, with citations and feedback telemetry.': 'SharePoint と社内規程文書を根拠とし、Teams から利用でき、Entra ID で保護された社内 RAG アシスタント。引用とフィードバックの計測を備える。',
  'Event-driven order processing at 50K orders/hour using API Management, Service Bus, Functions, Cosmos DB, Key Vault, and Application Insights.': 'API Management、Service Bus、Functions、Cosmos DB、Key Vault、Application Insights を使用し、1 時間あたり 5 万件を処理するイベント駆動型の注文処理。',
  'Import and modernize a three-tier application into Container Apps with private connectivity, managed identity, Azure SQL, Redis, and Front Door with WAF.': '3 層アプリケーションをインポートし、プライベート接続、マネージド ID、Azure SQL、Redis、WAF 付き Front Door を備えた Container Apps へモダナイズする。',
  'AI Discovery Cards workshop concept: reduce claims triage time using Document Intelligence, anomaly detection, human review, Fabric analytics, and D365 integration.': 'AI Discovery Cards ワークショップの構想: Document Intelligence、異常検知、人によるレビュー、Fabric 分析、D365 連携を使って保険金請求のトリアージ時間を短縮する。',
};

export default helpJapanese;
