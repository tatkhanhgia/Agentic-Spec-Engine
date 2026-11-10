# Render V3 từ `openapi.yaml` (2026-11-10, lần 2)

- Spec: nhánh `GoPaperless-Workflow-API` `claude/blissful-keller-9jp18x` (PR #440), dựa trên `dev` và thêm
  `required` cho 172 schema cùng schema riêng cho từng loại field.
- Seed: `formatted/GoPaperless Workflow RESTful API Specification V3 - 20260727.docx` (bản đã chốt).
- Kết quả: `formatted/GoPaperless Workflow RESTful API Specification V3 - rendered from OpenAPI.docx`.

```bash
python scripts/openapi_to_contract.py <openapi.yaml> -o contract/from_openapi.json
python scripts/render_from_openapi.py "formatted/...V3 - 20260727.docx" contract/from_openapi.json "formatted/...V3 - rendered from OpenAPI.docx"
python scripts/check_frame_format.py "formatted/...V3 - rendered from OpenAPI.docx"
```

## Renderer làm gì

- **Bảng thuộc tính** (Header Request, Path, Request, Response, Header Response): danh sách field, M/O và
  type lấy từ spec.
  - Field đã có trong V3 thì giữ dòng cũ: giữ **description của V3** (owner đã chọn) và giữ link Type tới object.
  - Field mới thì lấy description từ spec.
- **Khung Sample Request / Response 2xx**: tập key cấp 1 lấy theo spec.
  - Key đã có trong V3 thì giữ giá trị V3.
  - Key mới lấy `example` của property trong spec.
  - Key spec không còn thì bỏ.
  - Cả hai bản XML (DrawingML và VML) đều được ghi giống nhau.
- **Giữ nguyên từ V3**: bìa, History, Introduction, heading, hai chương object, các khung lỗi, và những
  section không ghép được với spec.
- **Presence**: `M` nếu field nằm trong `required`; `RC` nếu field chỉ bắt buộc trong một nhánh `anyOf`
  (ví dụ Add Signature: cần `page`+`dimension` hoặc `keyword_placement`); còn lại là `O`.

## Kiểm tra

| Trục | V3 | Bản render |
|---|---|---|
| Định dạng (`check_frame_format.py`) | 518 lỗi có sẵn | 518, **không thêm lỗi nào** |
| Heading | — | giống hệt |
| Hai bản XML của mỗi khung sample | — | 572/572 khớp nhau |
| Khung sample được ghi lại | — | 118 khung (236 bản XML); 104 khung đã khớp sẵn |

## So với V3

Có 54/265 endpoint giống hệt V3 về bảng. Lần render trước là 38 endpoint. Trong 3.778 dòng thuộc tính
của V3, khoảng 88% vẫn còn.

| | Lần 1 (spec trên `dev`) | Lần 2 (spec PR #440) |
|---|---|---|
| Response: field lệch M/O | 259 | **105** |
| Request: field lệch M/O | 182 | **152** |
| Request: field V3 có mà spec không có | 502 | **154** |
| Add TextBox Field: số field request | 9 (tất cả O) | 23 (`field_name`, `page`, `dimension` là M) |
| Sample vs bảng mâu thuẫn (vd `refresh_token_expires_in`) | có | không: sample đã theo spec |

### Phần lệch còn lại là chỗ V3 khác code, không phải lỗi render

- **Header `Content-Type` trên route không có body** (104 chỗ): V3 luôn ghi, spec chỉ khai báo khi có body.
- **Field V3 ghi M nhưng code không bắt buộc** (khoảng 84 dòng request, 90 dòng response).
  - Ví dụ: `refresh_token` chỉ có khi `remember_me_enabled`; `value` của các field type ("RC (Required attribute
    for processing.)") không bị validate.
  - Field response giữ M chỉ khi serializer luôn ghi key đó.
- **Field V3 có mà code không có**, hoặc đã đổi tên:
  - `read_only`/`place_holder` trên một số loại field: class Java không có các key này.
  - `refresh_token_expires_in` đổi thành `refresh_expires_in`.
  - Tham số path `pageNo`/`recordsPerPage` bị V3 liệt kê vào bảng Request.
- **Field code có mà V3 thiếu**: `field_security`, `lock_field_enabled`, `rotate`, `type`, `renamed_as` trên
  các field type; header `x-total-records`, `x-enterprise-mode`, `x-sort-by`, `x-asc`.
- **36 section chỉ có trong V3** (ví dụ `POST /v1/authenticate/otp`, Get Workflow Reminders): không có route
  trong code, nên renderer để nguyên.

## Cần owner quyết định

1. Header `Content-Type` trên route không có body: theo V3 (luôn ghi) hay theo spec (chỉ ghi khi có body)?
2. 36 section chỉ có trong V3: xóa, hay làm route trong code?
3. 102 operation chỉ có trong spec: có thêm section vào V3 không? Renderer hiện chưa tạo section mới.
4. 14 schema mà request và response bắt buộc khác nhau (`CustomFieldResponse`, `DocumentPackage`,
   `Document`, `FieldDimension`, `AddCategoryRequest`...): tách thành schema request riêng và schema response
   riêng?
5. Lỗi code agent phát hiện, chưa sửa:
   - `ResponseMessageService` ghi giá trị null thành `"key":` với phần giá trị bỏ trống, tức JSON không hợp lệ
     (ảnh hưởng `AddIdentityResponse`).
   - `TermsOfServiceItem` luôn trả 4 key null.
   - Get Fields không bao giờ tới được, vì regex của Get Field Types bắt trước.

Khi các mục 1–4 được chốt, chạy lại 3 lệnh ở trên là ra bản thay được V3.
