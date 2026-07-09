# AI-Driven Estimation & Proposal Kit

> 「この案件を見積もって／提案資料を作って」と頼めば、**誰でも（非エンジニアも）・AI でも、同じ構成・同じ品質の見積もり**を、**積み上げ根拠つき・抜け漏れなし**で、bravesoft 統一フォーマットで作れる「手順書」と「テンプレート」のセット。

![Version](https://img.shields.io/badge/version-0.1.1-0066ff) ![Output](https://img.shields.io/badge/output-Excel%20%7C%20HTML%20%7C%20PDF%20%7C%20PPTX-00bbff) ![For](https://img.shields.io/badge/for-非エンジニアOK-16a34a) ![License](https://img.shields.io/badge/License-MIT-yellow) ![Status](https://img.shields.io/badge/status-experimental-f59e0b)

---

## 👉 まずはここを開く

視覚的な使い方ガイド **`index.html`** をブラウザで開けば、5 分で全体像が掴めます（図解 + プロンプト定型集 + FAQ）。
完成イメージは **`02_Estimation_Template/example_estimate.md`**（お手本の見積もり）。

---

## このキットは何か

見積もり・提案づくりを AI（Claude Code / Cursor / Codex 等）と組んで**標準化**するためのドキュメント集です。
AI に渡すだけで、誰がやっても **同じ構成・同じ品質・積み上げ根拠つき・抜け漏れなし**の見積もりができあがります。
姉妹キット **AI-Driven Research & Reporting Kit / iOS Development Kit** と同じ運用思想（入口ファイルによる自動起動、番号付き Step、AI↔人間の通信テンプレ、横断記憶ログ、公式デザイン）を、見積もり・提案領域に翻案しています。

### 背景 — なぜ作ったか

社内には見積もりの作り方が散在しています（BrSE の旧テンプレ／営業がざっと作る概算／個人の独自テンプレ）。**全社共通の現行テンプレが無く、人によって項目も精度もバラバラ**でした。社内6スタイル＋現場の声を分析し、**全社共通で必要な構成を 7 レイヤー**に整理して、それを「誰でも・AI でも同じ品質で」生み出す仕組みにしたのが本キットです。

### 解決する課題

| 課題 | 本キットの解決策 |
|---|---|
| 見積もり項目・精度が人によってバラバラ | 統一フォーマット（7 レイヤー）＋ 番号付き Step（Rule 4） |
| 相場・記憶・どんぶり勘定で金額を出す | 機能分解 → 工数(人日)×単価の積み上げ＋根拠（Rule 1） |
| PM・テスト・ストア申請・非機能が抜ける | フェーズ×カテゴリで逆引き＋非機能チェックシート（Rule 2） |
| バッファ認識が人によって違う | バッファ・管理費・体制維持費を見える化（Rule 5） |
| 前提・対象外が曖昧で後で揉める | 見積根拠・対応範囲・対象外・契約形態を固定（Rule 3） |
| スケジュール作成が二度手間 | 見積明細から WBS/ガントを派生（Step 06） |
| 値引き交渉に根拠で対抗できない | 各項目に工数根拠 → 見積根拠資料（Step 08） |
| 提案書づくりに時間が取られる | 提案書 PPTX/Word を一括出力（Step 08） |
| 受けるべきでない案件に突っ込む | Go/No-Go アセスメント（Step 02） |
| 原価・利益率が混じった成果物を外に出す | 社内用と提出用を分離（Rule 4 / Rule 10） |

---

## 6 つの約束（このキットの核 = Rule 1〜6）

1. **積み上げ厳守**（機能分解 → 工数×単価、根拠を残す。No Basis, No Number）
2. **抜け漏れゼロ**（フェーズ×カテゴリで網羅、非機能を必ず通す）
3. **前提・スコープを固定**（見積根拠／対応範囲／対象外／契約形態／環境／納品物）
4. **統一フォーマット & 内部/提出の分離**（1 つの正本から社内用と提出用を派生）
5. **バッファ・管理費・採算を明示**（盛りと儲けを見える化、利益率を確認）
6. **非エンジニア・ファースト / 出力は選べる**（Excel/HTML/PDF/PPTX）

残り（Rule 7〜15）は、この約束を毎回確実に回すための運用・安全ルール（詳細は [`00_AI_Instructions/workflow_guide.md`](00_AI_Instructions/workflow_guide.md)）。

---

## 全社共通の必須構成（7 レイヤー）

| # | レイヤー | 主な中身 |
|---|---|---|
| ① | メタ・表紙 | 宛先 / 案件名 / 作成日 / 見積有効期限 / 見積精度 / 版管理 |
| ② | 前提条件・スコープ | 見積根拠 / 対応範囲 / 対象外 / 契約形態 / 環境 |
| ③ | 納入物・納品物 | 計画書 / 要件定義書 / 設計書 / 画面仕様書 / テスト報告書 |
| ④ | **見積明細（核）** | フェーズ × カテゴリ × 項目 × 工数(人日) × 単価 |
| ⑤ | 金額構成・集計 | 小計 / 管理費 / バッファ / 値引き / 税 / 合計（社内・社外） |
| ⑥ | スケジュール・体制 | WBS・ガント / 役割別アサイン / Phase・STEP 分割 |
| ⑦ | 提案・根拠・品質 | 提出用 / 見積根拠 / 提出前チェック / 人レビュー |

---

## ディレクトリ構造

```
ai-driven-estimation-kit/
├── README.md                            ← 本ファイル
├── index.html                           ← 視覚的な使い方ガイド（必読）
├── DESIGN.md                            ← bravesoft 公式デザインシステム（ブランドの正本）
├── CHANGELOG.md / LICENSE / VERSION
├── CLAUDE.md / AGENTS.md / .cursor/rules/  ← 各 AI ツールの起動入口（Rule 14 ブートストラップ）
├── 00_AI_Instructions/
│   ├── workflow_guide.md                ← 規約の正本（Rule 1〜15 + 通信テンプレ + 9 ステップ）
│   ├── ai_context_log.md                ← AI セッション横断記憶ログ
│   ├── output_formats.md                ← Excel / HTML / MD / PDF / PPTX 出力ガイド
│   └── pricing_and_rates.md             ← 標準単価・係数・管理費率の共通前提
├── 01_Estimation_Steps/                 ← メインフロー（案件のたびに使う Step 01〜09）
├── 02_Estimation_Template/              ← 統一フォーマット（正本 / 社内用 / 提出用）+ お手本
├── 03_Assessment_and_Checklists/        ← Go/No-Go・ヒアリング・非機能・前提/スコープ・納品物・提出前チェック
├── 04_Quick_Estimation/                 ← 軽量見積もりモード（「ざっくりいくら？」）
├── docs/
│   ├── estimation_brief.md              ← Step 01 で生成（見積もりの前提）
│   ├── go_no_go_report.md               ← Step 02 で生成（案件判断）
│   └── estimates/                       ← 生成した見積もり/提案の置き場
└── assets/                              ← ロゴ・brand.css（会社統一の正本スタイル）
```

---

## 5 分クイックスタート

1. **このキットを、見積もりたい案件のフォルダ（または専用フォルダ）のルートに丸ごと配置**する（`CLAUDE.md` / `AGENTS.md` / `.cursor/rules/` の入口ファイルを必ず含める＝ Rule 14 のブートストラップ）。
2. そのフォルダで **Claude Code / Cursor / Codex 等**を起動（入口ファイルが自動読込され、初回プロンプトは不要）。
3. **「○○の案件を見積もって」**と一声かける（Rule 14 によりキットの読込・状態同期は自動）。
4. AI が背景・案件概要・対象 OS・契約形態・予算感・締切・機密区分などをヒアリング（迷ったら「お任せ」で OK）。
5. AI が **アセスメント(Go/No-Go) → 機能分解 → 工数積み上げ → 金額&採算 → スケジュール → クロスチェック → 提案ドキュメント → 出力**の順で進める。
6. 仕上げに **Excel / HTML / PDF / PPTX** など、欲しい形式を指定（社内用と提出用は分けて出る）。

> 「ざっくりいくら？」程度なら `04_Quick_Estimation/`（同じ原則を最小手順で）。

---

## 出力フォーマット（使う人が選ぶ）

| 形式 | 向く用途 | 備考 |
|---|---|---|
| **Excel** | 計算・社内検討・既存運用との親和 | 工数集計・ON/OFF 再計算に強い |
| **Markdown** | 編集の正本・Notion/GitHub 共有 | 明細の構造化データ（正本） |
| **HTML** | 社内共有・図が綺麗・印刷で PDF | 提出用レイアウト |
| **PDF** | 顧客提出・清書 | HTML を印刷 → PDF |
| **PPTX** | 説明会・提案プレゼン | 提案書テンプレを複製して流し込む |

詳細は [`00_AI_Instructions/output_formats.md`](00_AI_Instructions/output_formats.md)。

---

## ブランド配色について

統一フォーマットの配色・ロゴ・フォントは、**bravesoft 公式デザインシステム（[DESIGN.md](DESIGN.md)）** に準拠しています（primary `#0066FF` / ink `#090A0A` / navy `#253A58` / Noto Sans JP ＋ Roboto Mono / 共通ロゴ `assets/bravesoft-logo.png`）。全形式（Excel / HTML / Markdown / PDF / PPTX / Word）を同じブランドで統一して出します。

---

## GitHub で管理する

- **正本は Markdown（明細の構造化データ）とテンプレート**。Excel / HTML / PDF / PPTX は派生物です。規約を変えるときの正本は [`00_AI_Instructions/workflow_guide.md`](00_AI_Instructions/workflow_guide.md)。
- **機密情報はコミットしない**（Rule 10）。[`.gitignore`](.gitignore) で実成果物（`docs/estimates/`）・持ち込み資料（`private/`）・確定前ブリーフ・`*.secret` / `*.key`・`.claude/settings.local.json` を除外済み。共有したい見積もりだけ `git add -f docs/estimates/<file>` で明示追加します。
- 変更は **ブランチ + Pull Request** 推奨。テンプレ・ルールを変えたら [CHANGELOG.md](CHANGELOG.md) に追記し、AI セッションの決定事項は [`ai_context_log.md`](00_AI_Instructions/ai_context_log.md) に残します（追記主義）。

---

## バージョン状態

**v0.1.1 (experimental)** — 規約・入口・標準単価・9 Step・テンプレート・チェックリスト・視覚ガイドを整備済み。実案件 **2 件で検証**（アプリ新規／既存改修）。**見積もり Excel を JSON から自動生成**（`05_Tools/generate_estimate_excel.py`）: 見積サマリ（案件概要＋金額）／見積書（明細・社内列は折りたたみ）／スケジュール（動的ガント）／**ヒアリングシート（観点チェック・AI事前記入＋人は✓）**／非機能要件チェックを、bravesoft ブランドで出力。**提案書 PPTX も同じ JSON から自動生成**（`05_Tools/generate_proposal_pptx.py`・明るい図解デッキ）。提出物は `estimate_submission.html` → PDF が標準。実運用での本番投入はこれから。詳細は [CHANGELOG.md](CHANGELOG.md)。

## ライセンス

[MIT License](LICENSE) — 商用利用可、改変自由、保証なし。

## クレジット

設計・整備: bravesoft Inc. ／ AI コラボレーション: Claude (Anthropic), Cursor
姉妹キット: **AI-Driven Research & Reporting Kit** / **AI-Driven iOS Development Kit**（運用思想・通信テンプレを共通化）
