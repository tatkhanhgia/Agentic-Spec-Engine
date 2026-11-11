# Render V3 từ `openapi.yaml` (lần 3, 2026-11-11)

- Spec: nhánh `GoPaperless-Workflow-API` `claude/blissful-keller-9jp18x`, dựa trên `dev` (đã có #439, #440) và
  thêm phần tách schema request, `x-doc-title` cho route đổi đường dẫn, `x-doc-status`.
- Seed: `formatted/GoPaperless Workflow RESTful API Specification V3 - 20260727.docx`.
- Kết quả: `formatted/GoPaperless Workflow RESTful API Specification V3 - rendered from OpenAPI.docx`.

```bash
python scripts/openapi_to_contract.py <openapi.yaml> -o contract/from_openapi.json
python scripts/render_from_openapi.py "formatted/...V3 - 20260727.docx" contract/from_openapi.json "formatted/...V3 - rendered from OpenAPI.docx"
python scripts/check_frame_format.py "formatted/...V3 - rendered from OpenAPI.docx"
```

Khi mở file trong Word, chọn **Yes** ở hộp thoại hỏi cập nhật field để mục lục có thêm các section mới
(renderer đã bật `updateFields`).

## Quyết định của owner đã áp dụng

| # | Quyết định | Cách làm |
|---|---|---|
| 1 | Bỏ `Content-Type` ở GET/DELETE | Adapter sinh dòng `Content-Type` từ `requestBody` (OpenAPI bỏ qua header param tên `Content-Type`). API không có body thì không có dòng này trong bảng và trong khung Sample Request. API có body ghi đúng định dạng (`application/json` hoặc `application/octet-stream`). |
| 2 | Section V3 chưa có trong code được ghi chú trên tên API | 36 tiêu đề có thêm ` (Not implemented yet)`: 32 section không có route nào, Get Package Verification, Get Workflow Reminders, và 2 route TSE mà code đã tắt (`x-doc-status: not-implemented`). Ba route chỉ đổi đường dẫn (Certificate Profile of Connector, Update Enterprise Electronic Seals, Update Address Book) vẫn giữ section, URL trong sample được sửa theo code. |
| 3 | Tạo section cho API mới | **88 section mới**, mỗi section đặt ở cuối nhóm (Heading 2) của nó. Mỗi section clone một section đã render có cùng method và cùng dạng có body / không body, ưu tiên cùng nhóm. Nội dung sinh từ spec: tiêu đề, mô tả, URL, header, khung sample (lấy từ `example`) và 5 loại bảng. Khung lỗi của section mẫu bị bỏ, id được đánh lại cho khỏi trùng. Nhiều route cùng tên thì gộp thành một section và ghi các đường dẫn còn lại trong mô tả. |
| 4 | Tách schema request/response | 7 schema request riêng trong spec (`CreateCustomFieldRequest`, `PackageNameRequest`, `DocumentPackageCreated`, `RenameDocumentRequest`, `DeclineWorkflowRequest`, `CreateCategoryRequest`, `CreateContactRequest`). 7 schema còn lại không bắt buộc field nào ở cả hai chiều nên vẫn dùng chung. |
| 5 | Lỗi code thành issue | tatkhanhgia/GoPaperless-Workflow-API#444 (Get Fields không tới được), tatkhanhgia/GoPaperless-Workflow-API#445 (JSON không hợp lệ). `TermsOfServiceItem` không phải lỗi: 4 key null là cố ý, để tương thích với SigningHub. |

## Kiểm tra

| Trục | V3 | Bản render |
|---|---|---|
| Định dạng (`check_frame_format.py`) | 518 lỗi có sẵn | 518, **không thêm lỗi nào** |
| Số endpoint trong tài liệu | 265 | 353 (265 + 88) |
| `wp:docPr` / VML `spid` trùng | 0 | 0 |
| Bookmark trùng / link hỏng | 19 / 14 (có sẵn) | 19 / 14 (không thêm) |
| Ghép với spec | 229 | 320 |

Bảng của section mới lấy độ rộng cột theo lưới của bảng, nên không chép lại lỗi độ rộng mà một số bảng V3
đang có (mục H).

## Còn lại

- **13 route chỉ có trong spec**: các route trùng tên đã gộp vào một section, đường dẫn của chúng được ghi trong
  mô tả của section đó.
- **12 bảng thiếu trong section mới**: section mẫu không có bảng loại đó, chủ yếu là bảng Header Response cho
  `x-total-records`. Muốn có thì cần một section mẫu chứa bảng ấy.
- **Nội dung V3 khác code**: field V3 ghi M nhưng code không bắt buộc, field đã đổi tên, tên field trong body lỗi.
  Renderer theo code; báo cáo chi tiết từng endpoint ở `reports/openapi_vs_v3_20261110.md`.
