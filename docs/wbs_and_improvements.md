# WBS & Đề xuất cải tiến / 機能分解と改善提案

> **Cách xem khuyến nghị:** mở [`wbs_and_improvements.html`](./wbs_and_improvements.html) (nút **Tiếng Việt / 日本語**).  
> **Mockup nhập liệu:** [`mockups/input_flow_optimized.html`](./mockups/input_flow_optimized.html)  
> **Prototype AS-IS:** [`source_materials/lipstickk_store_index.html`](./source_materials/lipstickk_store_index.html)

---

## Ghi chú quan trọng / 重要メモ

- Đã **bỏ qua** các bước tính chi phí / tiền bạc theo yêu cầu.
- Phần “thông tin đầu vào” và **mockup gốc** trong tin nhắn **chưa vào được môi trường** → WBS dựa trên prototype hiện có; mockup là **bản tham chiếu** do AI dựng. Gửi lại file gốc để gắn kèm / thay thế.
- コスト算出はスキップ。メッセージ内インプット本文・原本モックは未着。現行プロトタイプ＋参照モックで作成。

---

## AS-IS (tóm tắt / 要約)

| Module | Đã có / 既存 | Hạn chế / 課題 |
|---|---|---|
| Giá sỉ / 卸 | JPY→VND, 3 hệ số, lưu list | Nhập lặp tỷ giá; không sửa; không SKU |
| Lẻ / 小売 | Vốn + giá Shopee, lãi ~25.4% | Công thức lệch tab chi tiết |
| Phí chi tiết / 手数料 | 14%+3k+VX4%+TT4.9%+thuế | Không gắn SP đã lưu; hard-code |
| Tech | HTML + localStorage | Đổi máy mất data; 3 tab trùng nhập |

---

## WBS (ưu tiên / 優先度)

### F0 Nền tảng / 基盤
| ID | Hạng mục | VI | JA | Prio |
|---|---|---|---|---|
| F0.1 | Product master | 1 SP = tên/ảnh/SKU/kênh/vốn/giá/hệ số/phí | 1商品マスタ | Must |
| F0.2 | Cài đặt chung | Tỷ giá, ship/pack, hệ số sỉ, profile phí | 共通設定 | Must |
| F0.3 | Lưu bền vững | Cloud DB hoặc export/import | 永続化 | Must |
| F0.4 | UI VI/JA | Toggle ngôn ngữ | UI多言語 | Want |
| F0.5 | Auth | Login cửa hàng | 基本認証 | Want |

### F1 Wizard nhập liệu (theo mock) / 入力ウィザード
| ID | Hạng mục | Prio |
|---|---|---|
| F1.1 | Step1 thông tin SP | Must |
| F1.2 | Step2 chi phí & giá (JPY/VND) | Must |
| F1.3 | Preview lãi realtime | Must |
| F1.4 | Step3 lưu 1 lần → sỉ+lẻ | Must |
| F1.5 | Sửa / xoá / nhân bản | Must |
| F1.6 | Validate + làm tròn 1000đ + cảnh báo lỗ | Must |

### F2 Giá sỉ / 卸
| ID | Hạng mục | Prio |
|---|---|---|
| F2.1 | Giữ công thức 3 bậc | Must |
| F2.2 | Danh sách + tìm kiếm | Must |
| F2.3 | Override hệ số theo SP | Want |
| F2.4 | Export bảng giá PDF/ảnh | Want |

### F3 Lẻ & phí Shopee / 小売・手数料
| ID | Hạng mục | Prio |
|---|---|---|
| F3.1 | Thống nhất công thức (bỏ 25.4% xấp xỉ) | Must |
| F3.2 | Breakdown gắn SP đã lưu | Must |
| F3.3 | Profile phí có ngày hiệu lực | Want |
| F3.4 | Mô phỏng đổi giá bán | Want |
| F3.5 | Shopee API | Out (Phase 3) |

### F4 Báo cáo / 運用
| ID | Hạng mục | Prio |
|---|---|---|
| F4.1 | Search / filter / sort | Must |
| F4.2 | Export CSV/Excel | Must |
| F4.3 | Dashboard biên lãi | Want |
| F4.4 | Nhắc cập nhật tỷ giá | Want |
| F4.5 | Kho / đơn hàng | Out |
| F4.6 | App native | Out |

### F5 QA kỹ thuật / 技術QA
| ID | Hạng mục | Prio |
|---|---|---|
| F5.1 | Responsive mobile web | Must |
| F5.2 | Unit test công thức | Must |
| F5.3 | Migrate localStorage cũ | Want |
| F5.4 | Backup/restore | Want |

---

## Đề xuất ưu tiên / 改善提案（優先順）

1. **Gộp luồng nhập liệu (Must)** — đúng mockup wizard  
2. **Cài đặt chung tỷ giá/phí/hệ số (Must)**  
3. **Một engine tính lãi (Must)**  
4. **Cloud + search/export (Must)**  
5. **Sim giá bán (Want)**  
6. **PDF bảng sỉ (Want)**  
7. **UI VI/JA (Want)**  
8. **Shopee API / kho (Out → sau)**

---

## Roadmap

| Phase | Nội dung | Kết quả |
|---|---|---|
| 0 | Chốt WBS + gắn mockup gốc | Phạm vi rõ |
| 1 MVP | F0.1–0.3, F1.*, F2.1–2.2, F3.1–3.2, F4.1–4.2, F5.1–5.2 | Dùng hàng ngày |
| 2 | F0.4–0.5, F2.3–2.4, F3.3–3.4, F4.3–4.4 | Vận hành & báo cáo |
| 3 | F3.5, F4.5… | Tích hợp sàn/kho |

---

## Ngoài phạm vi tạm thời / 当面の対象外

- Tính chi phí dự án / 金額見積
- Shopee Open API
- Kho & kế toán đầy đủ
- Native app
