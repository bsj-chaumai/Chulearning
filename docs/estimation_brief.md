# 見積もりブリーフ / Brief — 更新（2026-07-09）

> **方針確定:** 金額・採算ステップはスキップ。成果物は **状況整理の提案ドキュメント（提案）** ＋ **WBS** ＋ **改善提案**。言語は **日本語 ⇄ ベトナム語**。

## 対象案件（モック原本より確定）

- **案件名:** WDP 稼働分析セットアップ（Operation Analysis Setup）改善
- **背景:** 旧UIはデータソースを後から選び、ソースごとに詳細設定画面が分かれ、カラムは手入力 → 学習コスト・入力ミス・画面分散が課題
- **ゴール:** 4データソースを **1ウィザード** に統合。ファイル選択直後に FE（ルールベース・AI不使用）がカラムを自動提案し、確認後に一括保存
- **対象外（本ドキュメント）:** 金額見積、Lipstickk Store（同一repoの別プロダクト）

## 提供資料（すべて反映済み）

| 資料 | パス |
|---|---|
| 稼働分析セットアップ・モック（原本） | `docs/mockups/operation-setup-wizard-mockup.html`（origin: `main` の `operation-setup-wizard-mockup (3).html`） |
| 提案ドキュメント（本成果物） | `docs/proposal_operation_setup.html` / `.md` |
| Estimation Kit | リポジトリ展開済み（金額ステップ未使用） |
| Lipstickk prototype | `docs/source_materials/lipstickk_store_index.html`（別案件・参考） |

## 出力

- 提案 HTML（JA⇄VI 切替、モック埋め込み）
- 提案 Markdown 要約
- モック原本を docs に固定

- ブリーフ確定日: 2026-07-09
- 承認: 方針承認済み / 提案内容のレビュー待ち
