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

## Bổ sung 2026-11-11: object và body lỗi

- **Chương RELATIVE OBJECT có thêm 86 mục object** cho các kiểu dữ liệu mà bảng mới dùng tới (ví dụ
  `ElectronicSealResponse`, `FieldSecurity`, `ProfileServicePlan`).
  - Mỗi mục gồm Heading 2 và bảng thuộc tính, clone từ một mục object có sẵn và đặt ở cuối chương.
  - Object lồng nhau cũng được tạo mục riêng.
  - Cột Type của dòng mới trỏ tới mục đó bằng hyperlink + bookmark, đúng quy ước của tài liệu
    (`List<` + link + `>`, style Hyperlink màu đen). Có 180 link mới, không link nào hỏng hay rỗng.
  - Nếu V3 đã link cùng nhãn tới đúng một object thì dùng lại anchor đó.
- **Body lỗi:** báo cáo trước ghi nhầm rằng code trả `error`/`remark`. Thực tế code có 3 định dạng
  (`FormatSelector`), và cấu hình phát hành dùng `NEW` = `error_code`, `error_description`, `message`, đúng như
  V3. Spec đã sửa: `ErrorResponse` mô tả NEW, thêm `LegacyErrorResponse` và `HybridErrorResponse`. Ba dòng lỗi
  trong bảng Response của tài liệu giữ nguyên.

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

## Bổ sung 2026-11-12: section mới đủ khung lỗi và đủ bảng

- **Khung sample lỗi:** 80 section mới (mọi API nhận access token) có thêm khung lỗi dùng chung của V3
  (`Http Code: 400`, token không thuộc hệ thống), là khung mà 502/532 section V3 đang dùng. API không cần token
  (webhook, đăng ký) không có khung này.
- **Bảng thiếu:** khi section mẫu không có một loại bảng, renderer lấy bảng loại đó từ seed và đặt đúng thứ tự
  (Header Request, Path, Request, Header Response, Response). 12 bảng được thêm, không còn section mới nào
  thiếu bảng.
- **Dòng `Content-type` của Sample Response** trong section mới theo đúng media type của spec
  (`application/octet-stream`, `application/pdf`...); response không có body thì bỏ dòng này.
- **Tên mục object:** giữ tên schema. Chương RELATIVE OBJECT của V3 cũng đặt tên kiểu này
  (`NextSigner`, `SignatureFieldAttribute`, `UserOfEnterprise`). Muốn tên khác thì đặt `x-doc-type` trên schema.

## Bổ sung 2026-11-12: phần đầu tài liệu lấy từ `doc-meta.yaml`

Backend có thêm `docs/api/doc-meta.yaml`, đặt cạnh `openapi.yaml`. Renderer đọc file này (`--meta`) để dựng:

| Phần | Khoá trong file |
|---|---|
| Trang bìa: tên, phụ đề, `Version 3.1.261112`, địa danh và năm | `document`, dòng cuối của `history.entries` |
| INTRODUCTION, Target | `introduction.overview`, `introduction.target` |
| Bảng Abbreviation | `abbreviations` |
| Bảng System Connection Information | `connection` |
| Bảng History (mọi dòng) | `history.entries` |

- Seed chỉ còn cho mượn hình thức của các phần này. Ô nào nội dung không đổi thì giữ nguyên từng run.
- Ràng buộc, kiểm tra trước khi ghi file (`scripts/doc_meta.py`): version `major.minor` tăng dần; ngày
  `YYYYMMDD` có thật và không lùi; có tác giả; dòng mới hơn `3.0` tối đa 6 dòng, mỗi dòng tối đa 200 ký tự
  (các dòng cũ giữ như đã phát hành).
- **Thêm dòng History khi API đổi** (`scripts/propose_history.py`): mỗi version có một bản chụp API ở
  `contract/openapi_history/<version>.json`. Script so spec hiện tại với bản chụp của dòng cuối, soạn dòng mới
  (version + 0.1, ngày hôm nay, liệt kê API thêm / bỏ / sửa trong giới hạn ký tự); `--write` ghi vào cuối
  `doc-meta.yaml` và lưu bản chụp mới. Sửa lời văn trong file nếu cần.
- Renderer từ chối render khi API đã khác bản chụp của dòng History cuối (bỏ qua bằng `--allow-stale-history`).

```bash
python scripts/openapi_to_contract.py <openapi.yaml> -o contract/from_openapi.json
python scripts/propose_history.py contract/from_openapi.json <doc-meta.yaml> --write   # khi API đổi
python scripts/render_from_openapi.py <seed.docx> contract/from_openapi.json <out.docx> --meta <doc-meta.yaml>
```

## Bổ sung 2026-11-12: mô tả field lấy từ file mô tả, không lấy từ seed

- `config/doc-descriptions.yaml` giữ lời văn của tài liệu cho 2.148 mô tả ở 230 section (sinh một lần bằng
  `scripts/extract_doc_descriptions.py` từ V3, sau đó sửa tay). Khoá: nhóm API > section > loại bảng > field.
- Renderer (`--descriptions`): field có mục trong file thì dùng mục đó, không có thì dùng mô tả trong
  `openapi.yaml`. Xoá một mục là trả field đó về cho spec.
- Mô tả trong `openapi.yaml` **không bị sửa**: chúng đã được đối chiếu với code, còn lời văn V3 thì chưa, nên
  chép V3 đè lên spec sẽ làm spec kém chính xác. Hai nguồn nằm riêng: spec đúng theo code, file mô tả giữ lời
  văn tài liệu.
- So với bản render trước: 35 ô khác. 10 ô V3 để trống nay có mô tả của spec; 25 ô thuộc section mới trùng tên
  với một section V3 nay dùng lời văn của section đó.

```bash
python scripts/render_from_openapi.py <seed.docx> contract/from_openapi.json <out.docx> \
    --meta <doc-meta.yaml> --descriptions config/doc-descriptions.yaml
```

## Bổ sung 2026-11-12: chương object sinh từ `openapi.yaml`

- Renderer xoá toàn bộ mục của hai chương REQUEST/RESPONSE OBJECT và RELATIVE OBJECT trong seed (122 mục) rồi
  dựng lại từ spec: mỗi schema mà một field trong tài liệu trỏ tới (trực tiếp hoặc qua object lồng nhau) có một
  mục. Kết quả: **124 mục** (6 + 118), 371 link từ cột Type.
- **83 mục V3 không API nào dùng đã bỏ** theo quyết định của owner.
- `config/doc-objects.yaml` (`--objects`): tên và chương trong tài liệu của 39 schema ghép được với mục V3
  (ví dụ `PackageListItem` → `Package`, `PackageDetailsOwner` → `Owner(Package Details)`). Schema không có trong
  file dùng tên schema và nằm ở RELATIVE OBJECT.
- Lời văn V3 của các object đó (197 mô tả) nằm trong `config/doc-descriptions.yaml`, mục `objects`.
- **Cột Type theo spec:** field trỏ tới schema thì link tới mục của schema đó, nhãn là tên mục
  (`List<Package>`). 34 field V3 từng link tới object nhưng spec không khai báo schema nay hiện kiểu thuần.
- **Link tới chương ENUMERATION** nằm trong mô tả ("Possible values are here"). File mô tả ghi chúng dạng
  `[here](#WorkflowMode)`, renderer dựng lại thành link. Còn 46 link (V3 có 64; 18 link nằm trong các mục
  object đã bỏ).
- 32 link trong các section không render (chưa triển khai) trỏ tới object đã bỏ được chuyển thành chữ thường,
  nên tài liệu không còn link hỏng trong bảng (V3 có 14).
- Chương **ENUMERATION vẫn lấy từ seed** (18 mục). Bảng của nó có cột mô tả cho từng giá trị, mà spec chỉ có
  danh sách giá trị; muốn sinh từ spec cần đặt tên enum thành schema và thêm mô tả từng giá trị trong
  `openapi.yaml`.
- Kiểm tra: `check_frame_format.py` 515 (V3: 518, ba lỗi cũ nằm trong mục object đã bỏ); không có bookmark lẻ
  đôi, id trùng hay link hỏng trong bảng; xem chương object qua PDF của LibreOffice.

```bash
python scripts/render_from_openapi.py <seed.docx> contract/from_openapi.json <out.docx> \
    --meta <doc-meta.yaml> --descriptions config/doc-descriptions.yaml --objects config/doc-objects.yaml
```

## Bổ sung 2026-11-12: chương ENUMERATION đọc enum có tên từ `openapi.yaml`

- Spec (backend, PR #467) có 13 enum được đặt tên, mỗi enum kèm `x-enumDescriptions`: mô tả từng giá trị,
  lấy từ `EnumLabelRegistry` trong code. 29 chỗ dùng trỏ tới chúng.
- Adapter xuất `enums` (giá trị + mô tả) và đánh dấu field dùng enum (`enum`). Kiểu của field vẫn hiện là
  kiểu thuần (`String`).
- Renderer dựng mục cho từng enum trong chương ENUMERATION:
  - 3 enum điền vào mục V3 có sẵn theo `config/doc-objects.yaml` (`enums`): `ParticipantType` → Role,
    `ProcessTypeOfWorkflow` → WorkflowMode, `SignaturePosition` → AnnotationPosition. Bảng của ba mục này
    nay theo code: Role có thêm `PLACEHOLDER`, WorkflowMode có thêm `EXCEL`, lời mô tả theo code.
  - 10 enum còn lại có mục mới ở cuối chương, đặt tên theo schema.
- Field dùng enum mà mô tả chưa có link thì được thêm câu "Possible values are here" link tới mục đó.
  Link tới chương ENUMERATION trong bảng: 68 (trước: 46).
- 15 mục ENUMERATION còn lại của V3 **giữ nguyên từ seed**: spec chưa khai báo các enum đó (ví dụ
  `LevelOfAssurance`, `LockAction`, `Color`, `FontName`), cần đối chiếu code trước khi đưa vào spec.

## Bổ sung 2026-11-12: chương ENUMERATION theo code hoàn toàn (cách A của owner)

- Spec (backend, PR #469) đặt tên thêm 8 enum: tổng cộng 21 enum có tên.
- Chương ENUMERATION nay chỉ gồm các enum spec khai báo: **21 mục** (11 mục V3 được điền lại theo code, 10 mục
  mới). Tên mục V3 ↔ tên enum nằm ở `config/doc-objects.yaml`, mục `enums`.

| Mục V3 | Enum trong spec | Thay đổi |
|---|---|---|
| WorkflowStatus | `WorkflowCriteria` | 13 giá trị theo code; 8 giá trị giữ mô tả V3, 5 giá trị mới chưa có mô tả |
| WorkflowState | `WorkflowStatus` | mô tả theo code |
| LevelOfAssurance | `LevelOfAssurance` | 5 giá trị theo code, chưa có mô tả (giá trị V3 không phải của code) |
| DocumentType | `CredentialDocumentType` | giá trị viết gạch dưới theo code (`PERSONAL_ID`), mô tả V3 |
| HashingAlgorithm | `TsaHashAlgorithm` | bỏ `SHA1` |
| AppearanceDesign, AppearanceAlignment, DocumentTypeForEID | cùng tên / `AppearanceSignatureAlignment` / `IdentityDocumentType` | giữ nguyên giá trị và mô tả V3 |

- **7 mục V3 bị bỏ** vì spec không khai báo enum cho chúng: LockAction, LockPermission, SigningAlgorithm,
  FieldTypes, AuthorizeMode, Color, FontName.
- **Link tới enum chỉ giữ khi field khai báo đúng enum đó.** 27 mô tả lấy từ V3 có câu "Possible values are
  here" nhưng field không khai báo enum: chữ "here" còn lại, không còn là link. Muốn gọn thì sửa lời văn trong
  `config/doc-descriptions.yaml`, hoặc khai báo enum cho field trong spec.
- Link tới chương ENUMERATION trong bảng: 53. Không có link hỏng trong bảng; `check_frame_format.py` 515.

## Còn lại

- **13 route chỉ có trong spec**: các route trùng tên đã gộp vào một section, đường dẫn của chúng được ghi trong
  mô tả của section đó.
- **301 thuộc tính response chưa xác định bắt buộc hay không** (65 schema): giá trị lấy từ cột DB hoặc dịch vụ
  ngoài, repo không có schema DB nên không chứng minh được cột có NULL hay không. Spec để tuỳ chọn (O), là
  lựa chọn an toàn. Cần quyền đọc schema DB để chốt.
- **Chưa mở bằng Microsoft Word thật**: mới kiểm tra qua PDF của LibreOffice.
- **Nội dung V3 khác code**: field V3 ghi M nhưng code không bắt buộc, field đã đổi tên.
  Renderer theo code; báo cáo chi tiết từng endpoint ở `reports/openapi_vs_v3_20261110.md`.
