# Render bản V3 từ `openapi.yaml`: hiện trạng và việc cần làm

Ngày đo: 2026-11-10.
- Spec: `GoPaperless-Workflow-API` nhánh `master` @ `76e49d483`, file `docs/api/openapi.yaml` (OpenAPI 3.1, 288 operation, 346 schema).
- Tài liệu: `formatted/GoPaperless Workflow RESTful API Specification V3 - 20260727.docx` (265 endpoint).

Báo cáo chi tiết từng endpoint: `reports/openapi_vs_v3_20261110.md`. Báo cáo này đã
được chạy lại trên spec của nhánh backend `claude/blissful-keller-9jp18x` (tách từ `dev`), tức là
sau khi đã có các bổ sung ở mục 0.

## 0. Tiến độ

Đã làm ở backend (nhánh `claude/blissful-keller-9jp18x`), `OpenApiSpecConsistencyTest` +
`ApiContractSnapshotTest` vẫn pass:

| Mục | Trạng thái | Kết quả |
|---|---|---|
| 4.1 `x-doc-title` | Xong (52 operation) | Số tên khác nhau giảm từ 96 xuống 2 |
| 4.2 `x-doc-sections` (tiêu đề, path) | Xong: `authenticate` (2 section), add/update field (2 × 20 section) | Mỗi section ra một bản ghi riêng có đúng tên và path. **Chưa** có schema riêng cho từng loại field |
| 4.5 object có tên | Mới làm `FieldDimension` (`x-doc-type: dimension`) | Số type request lệch giảm từ 65 xuống 25 |
| 4.3 `required` của response | Chưa làm | Cần đọc serializer để biết field nào luôn được trả về; không suy đoán |
| 4.4 `x-presence` (RC) | Chưa làm | — |
| 4.6, 4.7 | Chưa làm | 4.7 cần owner chốt |

Hai tên còn lệch là do ghép cặp, cần owner xem:
- `Share Workflow`: tài liệu có **hai section cho cùng một route** `POST /v1/packages/{id}/workflow/share`.
- `Get Workflow Reminders` (`GET /v1/packages/{id}/workflow/reminders`): đường dẫn này khớp với
  regex `x-route` của `getAllWorkflows`. Có thể route reminders đang bị Get All Workflow nuốt mất
  (tương tự lỗi Get Fields mà FEATURE-111 đã ghi nhận). Cần kiểm tra lại trong code.

## 1. Input để ra được file V3 hôm nay

`fix_docx_format.py` **không sinh nội dung**. Nó lấy `unpacked_docx/` cộng với
`scripts/GoPaperless_DOCX_Format_Template.md`, chuẩn hóa định dạng rồi đóng gói lại.
Nội dung của V3 là do người viết tay vào DOCX.

Lưu ý: `unpacked_docx/` hiện khớp với bản `V2-Formatted-New`, **không phải V3**
(nó thiếu History 2.6 → 3.0 và Version 3.0.260727). Muốn render lại V3 thì phải giải
nén V3 vào `unpacked_docx/` trước.

## 2. Pipeline đích

```
openapi.yaml ──openapi_to_contract.py──► contract/from_openapi.json ──┐
                                                                     ├─ compare_openapi_doc.py ─► báo cáo lệch
V3.docx ─────extract_doc_contract.py───► contract/from_doc.json ─────┘
                                                                     
from_openapi.json + V3.docx (làm seed) ──[renderer: CHƯA CÓ]──► unpacked_docx/ ──fix_docx_format.py──► V3 mới
```

Hai bước đầu đã có và chạy được:

```bash
python scripts/openapi_to_contract.py <đường_dẫn>/openapi.yaml          # -> contract/from_openapi.json
python scripts/extract_doc_contract.py "formatted/...V3 - 20260727.docx" -o contract/from_doc.json
python scripts/compare_openapi_doc.py contract/from_openapi.json contract/from_doc.json -o reports/<tên>.md
```

**Spec không thể thay thế toàn bộ DOCX.** Những phần không phải endpoint (bìa,
History, Introduction, chương REQUEST/RESPONSE OBJECT và RELATIVE OBJECT, phần mã
lỗi) không có trong spec. Vì vậy renderer phải dùng V3 làm **seed**: giữ nguyên các
phần đó và chỉ sinh lại nội dung các section endpoint (Heading 3/4) từ spec.

## 3. Mức khớp hiện tại

| | Số lượng |
|---|---|
| Endpoint ghép cặp được (theo method + `x-route`) | 229 / 265 |
| Chỉ có trong tài liệu | 36 |
| Chỉ có trong spec (chưa có trong tài liệu) | 100 |
| Endpoint đã ghép mà **không lệch field nào** | 18 |
| Tên endpoint khác nhau | 96 |
| Response field: M/O khác nhau | 259 (doc ghi `M`, spec để `O`) |
| Request field: M/O khác nhau | 182 |
| Header `Content-Type` có trong doc nhưng không có trong spec | 106 |

Hàng lỗi chung ở cuối mỗi bảng response (`error_description`, `error_code`,
`message` với presence `M/O`) là quy ước trình bày, nên đã bị loại khỏi phép so sánh.

## 4. Cần update gì trong `openapi.yaml` (phía backend)

Xếp theo mức ảnh hưởng tới việc render. Các mục dùng extension `x-doc-*` không
làm thay đổi contract của API, chỉ phục vụ tài liệu.

1. **Tên section tài liệu** → `x-doc-title` trên mỗi operation. `summary` đang là
   `apiName` trong code (`Create Package`, `PreLogin`, `Login SSO`), còn tài liệu dùng
   tên khác (`Add Document Package`, `Pre Login Authentication`, ...). Có 96 chỗ khác nhau.
2. **Một operation tương ứng nhiều section** → `x-doc-sections`.
   - `POST /v1/authenticate`: tài liệu tách thành 2 section (Client Credentials và Password).
   - `POST|PUT .../fields/{fieldType}`: tài liệu có **một section cho mỗi loại field**
     (text, checkbox, signature, vneid, ...; 37 section), nhưng spec chỉ có một
     schema `FieldAttributeRequest`. Cần một schema cho mỗi loại field (dạng `oneOf`
     + `discriminator` theo `fieldType`, hoặc danh sách `x-doc-sections`, mỗi mục một
     `fieldType` và một `$ref` schema). Đây là nguồn chính của 502 request field "chỉ có trong doc".
3. **`required` cho response schema.** Chỉ 13 trên 78 response schema có `required`,
   nên 259 field mà doc ghi `M` đang ra `O`. Cần khai báo `required` cho các field luôn được trả về.
4. **Bắt buộc có điều kiện (RC).** Tài liệu dùng `RC (When ...)` ở 70 chỗ. OpenAPI
   không có khái niệm này, nên cần thêm `x-presence: "RC (When operation_mode is S)"`
   trên property.
5. **Object lồng phải là component có tên, kèm tên hiển thị.** Ví dụ `dimension`
   đang là object inline (40 chỗ ra `Object`), trong khi doc hiển thị `dimension` và
   link tới chương RELATIVE OBJECT. Cần tách thành `$ref` tới schema có tên, cộng
   `x-doc-type: dimension` và `x-doc-object: <tên Heading 2 trong RELATIVE OBJECT>`
   để renderer tạo hyperlink bookmark đúng.
6. **Thứ tự section** → `x-doc-order` (hoặc giữ đúng thứ tự trong file yaml). Thứ tự
   tag đã khớp với các Heading 2.
7. **Body lỗi khác nhau.** Tài liệu ghi `error_description`, `error_code`, `message`;
   spec `ErrorResponse` ghi `error`, `error_description`, `error_code`, `remark`. Cần chốt bên nào đúng.

## 5. Cần quyết định (không tự sửa)

Theo quy tắc của dự án, mỗi chỗ code và tài liệu mâu thuẫn nhau đều phải do chủ dự án quyết định:

- **36 endpoint chỉ có trong tài liệu** (danh sách ở báo cáo chi tiết). Ví dụ
  `POST /v1/authenticate/otp` không có route trong code; các route workflow
  `permissions`/`authentication`/`reminders` thì có thể chỉ khác path. Mỗi endpoint
  cần chốt: xóa khỏi tài liệu, hay sửa spec/code.
- **100 operation chỉ có trong spec**: có thêm vào V3 không? Báo cáo FEATURE-111 ghi
  có 89 operation không có trong DOCX V3.
- **Header `Content-Type` trên route không có body** (106 chỗ): giữ theo tài liệu hay bỏ theo code?
- **Các khác biệt mà FEATURE-111 đã liệt kê** (đổi tên field, wrapper của list, enum,
  method/path, public hay cần xác thực): spec đang đi theo code.

## 6. Cần làm phía Agentic-Spec-Engine

1. Đọc thêm các extension `x-doc-*` ở mục 4 trong `openapi_to_contract.py`.
2. Viết **renderer**: với mỗi endpoint, clone section mẫu cùng hình dạng trong V3
   (GET không body, POST có body, response rỗng/có dữ liệu), điền tiêu đề, khung
   sample (lấy từ `example`), 4 bảng thuộc tính, rồi đặt lại cho duy nhất bookmark,
   `wp:docPr/@id` và VML id theo CLAUDE.md. Những phần không phải endpoint lấy nguyên từ seed V3.
3. Kiểm chứng: render lại từ spec và so với V3 trên cả 3 trục (text, format, structure).
   Khi `compare_openapi_doc.py` báo 0 lệch, sau khi đã trừ các mục owner chọn giữ
   theo tài liệu, thì bản render chính là V3 chuẩn.

Đề xuất thứ tự làm: mục 4.1 + 4.3 + 4.5 trước (dễ, và giảm được phần lớn các chỗ lệch),
rồi 4.2, sau đó mới tới renderer.
