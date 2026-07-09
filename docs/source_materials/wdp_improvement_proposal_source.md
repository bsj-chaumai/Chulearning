# WORKPLACE DATA PLATFORM (WDP) - ĐỀ XUẤT CẢI TIẾN HỆ THỐNG

## 1. Tầm Nhìn Chiến Lược & Mục Tiêu (Vision & Objectives)
Chuyển đổi hệ thống WDP từ một công cụ hiển thị biểu đồ tĩnh thành một **Hệ thống SaaS AI-driven thông minh**.
* **Đối với Khách hàng (Tenant Admin):** Tối ưu hóa trải nghiệm người dùng (UX), trực quan hóa quy trình xử lý (Flow) và giảm thiểu tối đa các thao tác thủ công.
* **Đối với Hoạt động Kinh doanh của Vis:** Tạo ra tính năng mũi nhọn (USP) vượt trội so với đối thủ, thúc đẩy tỷ lệ chuyển đổi (CVR) của bản SaaS và nâng cao năng suất đột phá cho đội ngũ tư vấn (XD).

---

## 2. Các Vấn Đề Hiện Tại (Pain Points)

### 2.1. 製品・技術面の課題（Vấn đề Sản phẩm & Kỹ thuật）
* **手作業の負担（Gánh nặng thao tác thủ công）：** Khách hàng mất quá nhiều thời gian để chỉnh sửa định dạng file Excel/CSV trước khi upload dữ liệu nhân viên và lịch sử ra vào.
* **データの可視化止まり（Chỉ dừng lại ở mức hiển thị）：** WDP hiện chỉ vẽ biểu đồ (xu hướng, tỷ lệ lên công ty, kết quả survey). Việc đọc hiểu, phân tích số liệu hoàn toàn phụ thuộc vào con người.
* **技術的負債（Nợ kỹ thuật）：** Source code còn nhiều vùng bị **hard-code**, tiềm ẩn rủi ro sai lệch dữ liệu và là rào cản lớn khi muốn tích hợp AI ngay lập tức.

### 2.2. ビジネス・市場面の課題（Vấn đề Thị trường & Kinh doanh SaaS）
* **認知拡大の遅れ（Nhận diện thương hiệu chậm）：** Đang bị đi sau các đối thủ cùng ngành (các công ty HR-tech đã ứng dụng AI phân tích chuyên sâu trước Vis).
* **強みの不足（Thiếu tính năng mũi nhọn）：** Đội ngũ kinh doanh (Sales) chưa có một tính năng mang tính đột phá ("Wow-effect") để làm vũ khí cạnh tranh trên thị trường.
* **成約率の課題（Tỷ lệ chuyển đổi thấp）：** Nhu cầu thị trường rất cao nhưng tỷ lệ khách hàng xuống tiền mua bản SaaS trả phí bị nghẽn do rào cản từ việc thiết lập ban đầu quá phức tạp.

### 2.3. 運用・組織面の課題（Vấn đề Vận hành & Nhân sự nội bộ Vis）
* **コンサル業務の高負荷（Quá tải mảng tư vấn）：** Các chuyên gia tư vấn (Consultant) của Vis phải tự nhìn biểu đồ WDP rồi làm báo cáo thủ công -> Gây tốn thời gian, khó mở rộng quy mô dịch vụ (Scale-up).
* **内製化における技術力の壁（Rào cản công nghệ khi In-house）：** Định hướng muốn in-house hóa đội ngũ IT, nhưng nếu chỉ tuyển kỹ sư thông thường thì khó tự nghiên cứu và phát triển các công nghệ cao như AI để đấu với đối thủ.

---

## 3. Giải Pháp & Định Hướng Cải Tiến Cụ Thể (To-Be)

### Giai đoạn 1: Tối ưu hóa Luồng Nhập liệu (Input & Data Cleansing)
*(Áp dụng tại Top Page, khu vực `稼働分析` và `空間分析`)*
* **Thực trạng cũ:** Người dùng mất thời gian format file Excel đầu vào (`出社滞在情報` và `従業員情報`). Khi file sai format, hệ thống tự động từ chối xử lý nhưng không hiển thị bất kỳ thông báo lỗi nào, gây ức chế và khó hiểu cho người dùng. Thao tác rời rạc, lặp đi lặp lại.
* **Giải pháp cải tiến:**
    * **Hợp nhất quy trình (Step-by-step Popup):** Gom toàn bộ khu vực upload dữ liệu vào một Popup tập trung, hướng dẫn người dùng thao tác tường minh theo từng bước trực quan.
    * **Giao diện Sửa lỗi trực tiếp (Error UI):** Khi phát hiện file sai định dạng, hệ thống sẽ hiện thông báo lỗi và đề xuất sửa lỗi tương ứng cho người dùng.

### Giai đoạn 2: Quản lý và Vận hành Survey Thông minh
* **Thực trạng cũ:** Thiết lập survey thủ công. Muốn kiểm tra ai chưa làm, Tenant Admin phải Export file CSV rồi lọc từng dòng bằng mắt trên Excel, sau đó tự liên hệ nhắc nhở bên ngoài hệ thống.
* **Giải pháp cải tiến:**
    * **Dashboard Trạng thái Real-time:** Hiển thị biểu đồ tiến độ trực quan (Tỷ lệ Đã làm / Chưa làm) ngay trên Web, loại bỏ hoàn toàn việc export CSV chỉ để check trạng thái.
    * **Hệ thống Nhắc nhở Tự động (Automation Queue):** Tích hợp tính năng tự động gửi mail nhắc nhở định kỳ (Auto-remind) hoặc thiết lập nút "Nhắc nhở nhanh 1-chạm" gửi đến những user chưa hoàn thành survey (áp dụng cho luồng survey qua mail).

### Giai đoạn 3: Phân tích Đầu ra (Output Analytics) & 3 Mức Độ Ứng Dụng AI
*(Áp dụng tại `稼働分析` và kết quả `サーベイ結果`)*
* **Thực trạng cũ:** Hệ thống chỉ cung cấp biểu đồ tĩnh. Người dùng và Consultant vẫn phải tự nhìn, tự phân tích bằng mắt để đưa ra giải pháp không gian.
* **Đề xuất 3 mức độ ứng dụng AI để thảo luận:**

#### Mức độ 1 — AI Insight Card (Thấp)
* **Chi tiết:** Thêm một thẻ text phân tích kèm biểu đồ nhỏ (mini-chart) ngay dưới mỗi biểu đồ hiện tại. Khách hàng vẫn xem biểu đồ truyền thống nhưng có thêm AI "đọc hiểu" và viết nhận xét tự động.

#### Mức độ 2 — Workplace Brief (Trung bình - Phân tích chéo Cross-domain)
* **Chi tiết:** Xây dựng một trang Dashboard tổng hợp liên kết toàn bộ dữ liệu từ Vận hành (Operation) + Khảo sát (Survey) + Không gian (Space). Hàng tháng, AI tự động xuất bản một "Bản tin văn phòng" bao gồm: Điểm sức khỏe văn phòng (Health Score), các insight phân tích chéo và đề xuất hành động cụ thể.

#### Mức độ 3 — AI Chat (Cao - Hỏi đáp dữ liệu thông minh)
* **Chi tiết:** Tích hợp tính năng Chatbot tương tác dữ liệu (Ứng dụng Text-to-SQL bảo mật). Người dùng nhập câu hỏi bằng tiếng Nhật tự nhiên (Ví dụ: *"Tháng trước phòng nào đi làm trực tiếp ít nhất?"*), AI tự bốc dữ liệu, xử lý câu lệnh và trả về câu trả lời bằng văn bản kèm biểu đồ tương ứng ngay trong khung chat mà không cần đến sự hỗ trợ của Consultant.