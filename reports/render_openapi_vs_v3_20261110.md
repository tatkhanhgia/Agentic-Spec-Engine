# Thử render V3 từ `openapi.yaml` mới nhất (2026-11-10)

- Spec: `GoPaperless-Workflow-API` `dev` @ `3316e775a` (đã có PR #439: `x-doc-title`, `x-doc-sections`, `FieldDimension`).
- Seed: `formatted/GoPaperless Workflow RESTful API Specification V3 - 20260727.docx` (bản đã chốt).
- Kết quả: `formatted/GoPaperless Workflow RESTful API Specification V3 - rendered from OpenAPI.docx`.
- Ảnh so sánh trang bảng thuộc tính của Client Credentials Authentication (trái V3, phải bản render): `reports/render_openapi_vs_v3_20261110.png`.

```bash
python scripts/openapi_to_contract.py <openapi.yaml> -o contract/from_openapi.json
python scripts/render_from_openapi.py "formatted/...V3 - 20260727.docx" contract/from_openapi.json "formatted/...V3 - rendered from OpenAPI.docx"
python scripts/check_frame_format.py "formatted/...V3 - rendered from OpenAPI.docx"
```

## Renderer làm gì

`scripts/render_from_openapi.py` chỉ sinh lại **các dòng dữ liệu của 5 loại bảng thuộc tính**
(Header Request, Path, Request, Response, Header Response) cho những endpoint ghép được với spec.
Mọi thứ khác lấy nguyên từ V3: bìa, History, Introduction, heading, khung Sample Request/Response,
hai chương object và style bảng.

- Field đã có trong V3 thì giữ dòng cũ, nên link Type tới object vẫn còn. Renderer ghi lại
  name, presence và description; type chỉ ghi lại khi ô đó là chữ thường, không có link.
- Field mới thì clone một dòng chữ thường của cùng bảng, hoặc của bảng cùng loại nếu bảng đó trống.
- Ba dòng lỗi chung ở cuối bảng Response được giữ nguyên.
- Section ghép theo regex mà tên không khớp thì không render (Get Workflow Reminders, section
  Share Workflow trùng).

## Kiểm tra

| Trục | V3 | Bản render |
|---|---|---|
| Định dạng (`check_frame_format.py`) | 518 lỗi có sẵn | 518, **không thêm lỗi nào** |
| Heading | — | giống hệt |
| 1598 khung sample | — | giống hệt (chưa render từ `example`) |
| Bảng | 1033 | 1033 |
| Hyperlink | 864 (14 hỏng có sẵn) | 705 (14 hỏng có sẵn, không thêm) |
| Số trang PDF (LibreOffice) | 715 | 654 |

Có 159 hyperlink mất đi vì nằm trong các dòng mà spec không còn, hoặc trong ô description bị ghi lại.

## Nội dung thay đổi so với V3

229/265 section ghép được, 227 section được render, 38 endpoint không đổi gì.

| Bảng | Thêm | Bỏ | Đổi M/O | Đổi type | Đổi description |
|---|---|---|---|---|---|
| Header Request | 84 | 127 | 3 | 9 | 710 |
| Path | 33 | 33 | 16 | 7 | 239 |
| Request | 285 | 502 | 182 | 16 | 541 |
| Response | 127 | 156 | 259 | 33 | 302 |

Ví dụ cho thấy rõ spec còn thiếu gì:

- **Client Credentials Authentication**
  - Request: từ 4 field thành 8. Spec chỉ có một schema dùng chung cho mọi `grant_type`, nên
    section này kéo thêm cả `username`, `password`, `refresh_token`, `scope`.
    → Cần thêm schema riêng cho từng section trong `x-doc-sections`.
  - Response: `refresh_token_expires_in` đổi thành `refresh_expires_in`. Bản render đúng theo code,
    nhưng **khung Sample Response vẫn ghi tên cũ**, nên bảng và sample đang mâu thuẫn nhau.
    → Cần render khung sample từ `example`.
  - Response: `refresh_token` đổi từ M sang O. Spec ghi field này chỉ có khi refresh token được cấp (`remember_me_enabled`).
- **Add TextBox Field**
  - Request: V3 có 18 field, bản render còn 9 và tất cả đều là O. Spec mới chỉ có các key chung
    (`FieldAttributeRequest`), chưa có key riêng của từng loại field, và cũng chưa khai báo `required`.
- **Description**: 2.000+ chỗ đổi câu chữ, từ văn của tài liệu (theo SigningHub) sang văn của spec.
  Spec viết đúng theo code nhưng ngắn hơn. Header `Authorization` hiện chỉ ghi "Bearer access token.".

## Kết luận

Định dạng thì dùng được ngay: renderer không làm vỡ gì. Về nội dung, bản render **chưa thay được V3**:

1. Nhiều field từ M thành O vì spec thiếu `required`, ở cả response lẫn request.
2. Nhóm Add/Update Field mất key riêng của từng loại field.
3. Khung sample chưa render từ spec, nên sample và bảng có thể mâu thuẫn nhau.
4. Văn mô tả đổi giọng hoàn toàn. Owner cần chọn: giữ văn của V3 khi field không đổi, hay theo spec.

Khi spec bổ sung đủ mục 1–2 và renderer làm thêm mục 3, chạy lại `render_from_openapi.py` sẽ ra bản thay được V3.
