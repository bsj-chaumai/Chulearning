# 提案書（状況整理）— WDP 稼働分析セットアップ改善

> **推奨の見方:** [`proposal_operation_setup.html`](./proposal_operation_setup.html)（日本語 ⇄ Tiếng Việt）  
> **モック原本:** [`mockups/operation-setup-wizard-mockup.html`](./mockups/operation-setup-wizard-mockup.html)

---

## 一言 / Một câu

稼働分析セットアップを、旧UIの「方式ごと・手入力・複数画面」から、**4データソース対応の1ウィザード**（FEルールでカラム自動提案 → 確認 → 一括 `POST /operation-analysis/setup`）へ改善する提案。金額見積は対象外。

Setup phân tích vận hành: từ UI cũ (nhiều màn, nhập tay) → **1 wizard / 4 nguồn** (FE đề xuất cột theo rule → xác nhận → lưu 1 lần). Không tính tiền.

---

## 共有済みインプット / Thông tin đã gửi

| 資料 | 役割 |
|---|---|
| `operation-setup-wizard-mockup (3).html` (GitHub main) | **主資料** — 本提案の中核 |
| `index.html` (Lipstickk) | 同一repoの別プロダクト（本提案の対象外） |
| `Archive.zip` (Estimation Kit) | 手順枠のみ。金額ステップはスキップ |
| 利用者方針 | コストスキップ / WBS+改善 / VI⇄JA |

---

## TO-BE 5ステップ

1. データソース選択（Wifi CSV / 入退室Excel / 手入力Excel / Meraki）
2. ファイルアップロード or Meraki連携
3. カラム設定（FE自動提案 → 手動変更可）
4. 集計条件（曜日・時間帯・祝日・タッチポイント）
5. 確認 → `POST /operation-analysis/setup`（method / input-setting / target-dates-slots → S3 → Glue）

---

## WBS（Must中心）

W1 方式選択 · W2 二重アップロード · W3 FEヘッダー解析 · W4 自動提案 · W5 手動オーバーライド · W6 Timestamp検出 · W7 入退室固有 · W8 Meraki UI · W9 集計条件 · W10 一括保存 · W12 進捗ナビ  
Want: W11 サンプル導線 · W13 旧UI比較ヘルプ

---

## 次アクション

1. 本提案（4方式・5ステップ・API）の内容確認  
2. OK → 「モックで実装方針を確定」  
3. NG → 修正点のみ指摘
