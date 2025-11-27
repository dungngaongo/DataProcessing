# Excel Execution (Flask + HTMX)

Ứng dụng web cho phép import Excel nhiều sheet, chỉnh sửa trực tiếp dữ liệu theo giao diện giống bảng tính, liên kết dự án giữa các sheet và xuất lại ra Excel. Tập trung vào thao tác nhanh với file lớn (nhiều cột, nhiều dòng) dùng HTMX để cập nhật partial thay vì reload toàn trang.

## Các Sheet & Cột Chuẩn

### Sheet "Sizing"

```
STT | Mã PYC | Đơn vị | Đầu mối tạo PYC | Đầu mối xử lý | Status | Thời điểm đẩy y/c | Thời gian hoàn thành theo KPI | Thời gian hoàn thành ký PNX và đóng y/c | Thời gian ký bản chốt sizing | Tên dự án - Mục đích sizing | Ghi chú
```

### Sheet "Cấp phát tài nguyên"

```
STT | Dự án | Đơn vị | Đầu mối y/c | Đầu mối P.HT | Mã SR | Tiến độ, vướng mắc, đề xuất | Thời gian tiếp nhận y/c | Timeline thực hiện theo GNOC | Thời gian hoàn thành | Hoàn thành
```

## Tính Năng Hiện Tại

- Import Excel: Đọc hai sheet theo tên (thiếu sheet thì khởi tạo rỗng). Tự thêm cột thiếu để giữ thứ tự chuẩn.
- Hiển thị bảng giống Excel: Header sticky, scroll ngang/dọc toàn màn hình.
- Chỉnh sửa ô trực tiếp: `contenteditable`, tự lưu (gửi AJAX/HTMX JSON) khi blur ô.
- Thêm dòng: Nút `+` xuất hiện dưới mỗi dòng khi hover, thêm ngay dòng mới phía dưới.
- Xóa dòng: Nút `×` cạnh nút `+` để xóa dòng và đánh lại STT.
- Đánh số STT tự động: Cập nhật lại khi thêm/xóa/sắp xếp (hiện chỉ thêm/xóa).
- Mapping dự án: Click ô "Tên dự án - Mục đích sizing" để chuyển sang sheet "Cấp phát tài nguyên" và highlight dòng có cột "Dự án" chứa hoặc trùng tên (tìm gần đúng, không phân biệt hoa thường).
- Highlight 2.5s: Dòng được outline + nền xanh nhạt rồi tự trả về trạng thái bình thường.
- Lưu cache JSON: Mọi thay đổi (import, sửa ô, thêm/xóa) ghi vào `cache/data_store.json`. Khởi động lại server sẽ load lại dữ liệu trước đó.
- Xuất Excel: Nút `Export Excel` tạo file bao gồm cả hai sheet với thứ tự cột chuẩn.
- Làm sạch dữ liệu trống: Các giá trị `NaN`, `NaT`, `None` hiển thị rỗng, tránh gây nhiễu.
- Không reload toàn trang: Chỉ vùng bảng thay đổi, giữ trạng thái focus người dùng tốt hơn.

## Cài Đặt & Chạy

```powershell
pip install -r requirements.txt
python app.py
```

Mặc định chạy ở: `http://127.0.0.1:5000`

## Quy Trình Dữ Liệu

1. Người dùng chọn file Excel và bấm Import.
2. Server đọc hai sheet bằng `pandas`, chuẩn hoá cột, làm sạch giá trị trống.
3. Chuyển DataFrame thành list dict lưu trong `data_store` và cache JSON.
4. Giao diện hiển thị bảng với các ô editable. Khi người dùng sửa ô, JS gửi yêu cầu `/update-cell` để cập nhật.
5. Thêm/xóa dòng gọi `/add-row` hoặc `/delete-row` trả về partial HTML sheet mới.
6. Xuất file gọi `/export` dựng workbook mới từ `data_store`.

## Danh Sách Endpoint

| Method | Route                               | Mô tả                                    |
| ------ | ----------------------------------- | ------------------------------------------ |
| GET    | `/`                               | Trang chính + 2 bảng                     |
| POST   | `/import`                         | Import file Excel, cập nhật bảng        |
| POST   | `/update-cell`                    | Cập nhật 1 ô (JSON)                     |
| POST   | `/add-row/<sheet>/<after_index>`  | Thêm dòng sau chỉ số cho trước       |
| POST   | `/delete-row/<sheet>/<row_index>` | Xóa dòng chỉ định                     |
| GET    | `/export`                         | Tải file Excel mới                       |
| GET    | `/sheet/<name>`                   | Partial render sheet (dùng nội bộ HTMX) |

## Kiến Trúc & Thư Mục

```
excel-execution/
	app.py               # Flask app + endpoints + cache
	templates/           # base.html, index.html, tables.html, sheet.html
	static/              # main.js (logic edit, mapping), style.css
	uploads/             # Lưu file import và export
	cache/data_store.json# Cache dữ liệu hiện tại
```

## Ghi Chú Kỹ Thuật

- `pandas` + `openpyxl` để đọc/ghi Excel.
- Giới hạn upload: 200MB (`MAX_CONTENT_LENGTH`).
- Dùng HTMX `hx-post` + `hx-target` để thay thế phần bảng.
- Lưu cache dưới dạng JSON giúp khởi động lại không mất dữ liệu.
- Làm sạch dữ liệu với filter Jinja `blanknan` và hàm Python `sanitize_rows`.

## Mẹo Hiệu Năng

- Nếu file rất lớn: có thể phân trang hoặc lazy render (limit số dòng đầu, thêm nút tải thêm).
- Sử dụng WebSocket hoặc Server-Sent Events nếu muốn phản hồi thời gian thực cho nhiều người dùng.
- Có thể chuyển sang streaming đọc Excel chunk nếu kích thước cực lớn.

## Hướng Mở Rộng

- Undo/redo các thao tác.
- Mapping ngược từ Cấp phát tài nguyên → Sizing.
- Tìm kiếm & lọc cột ngay trên giao diện.
- Phân quyền / đăng nhập (Flask-Login / JWT).
- Xuất thêm định dạng CSV, JSON.
- Thống kê nhanh (đếm status, tiến độ) hiển thị dạng dashboard.
- Validate định dạng ngày, mã PYC, mã SR trước khi lưu.

## Troubleshooting

| Vấn đề                | Nguyên nhân                          | Cách xử lý                                                           |
| ------------------------ | -------------------------------------- | ----------------------------------------------------------------------- |
| Không thấy bảng       | Cache rỗng hoặc CSS không load      | Hard refresh (Ctrl+Shift+R), kiểm tra Network tab                      |
| Hiện "nan"              | Filter chưa áp dụng                 | Đảm bảo đã refresh sau cập nhật mới nhất                       |
| Không highlight dự án | Chuỗi không khớp                    | Kiểm tra viết sai chính tả hoặc khác dấu, dùng chứa một phần |
| Export trống            | Chưa import hoặc xóa hết dữ liệu | Thêm dòng rồi export lại                                            |

## License

Nội bộ (private). Không phân phối bên ngoài nếu chưa được phép.

& D:\workspace\DataProcessing\.venv\Scripts\python.exe d:/workspace/DataProcessing/excel-execution/app.py

Invoke-RestMethod -Method Post -Uri http://127.0.0.1:5000/trigger-whatsapp-alert


List test:

- import
- export
- đúng dang dữ liệu
- auto fill kpi, tiến độ
- insert/delete/edit row, column
- mapping between sheet
- thiet lap canh bao gui ve tin nhan WhatsApp

to do:

- chuan hoa du lieu by id, doi lai mapping bang id
- xuat bao cao (can lam ro)
- bang total
- auto calculate o tinh
- them cac sheet khac
