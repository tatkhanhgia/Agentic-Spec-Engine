# OpenAPI as the content source for rendering the specification

Status: proposal. Nothing here is implemented yet.

## Why the current inputs are not enough

`fix_docx_format.py` only formats content that already exists in `unpacked_docx/`;
it never generates sections. Rendering a document from the backend needs a
document-shaped contract, and today no backend source provides one:

| Needed by the document | `contract/from_code.json` | backend `api-contract-snapshot.json` |
|---|---|---|
| Method, path, endpoint name | yes | no |
| Request/response fields + types | no | yes (`jsonName`, `type` only) |
| Endpoint ↔ DTO mapping | no | no |
| Mandatory/optional (M/O) | no | no |
| Field descriptions | no | no |
| Headers, path params | no | no |
| Sample request/response | no | no |

Measured on 2026-10-19: `extract_code_contract.py` finds 273 endpoints in 20
controllers; the snapshot holds 200 entries (93 request, 52 response,
55 serializer) with 1,520 fields.

The target contract shape is `contract/from_doc.json` (name, category, level,
method, path, headers, path_params, request_fields, response_fields with
type/presence/description). It is framework-neutral and can be the common input
for every project.

## Proposed pipeline

```
openapi.yaml ─[adapter: OpenAPI → from_doc.json shape]─► renderer ─► fix_docx_format.py ─► DOCX
```

- Adapter: thin format conversion. Not written yet.
- Renderer (contract → DOCX by cloning template sections): not written yet.
  Validate it by rendering GoPaperless's own `from_doc.json` and diffing against
  the current V3 document.

## Prompt for the backend agent (GoPaperless-Workflow-API)

The backend is pure servlet (no JAX-RS/Spring allowed by its governance), so the
spec is hand-written and guarded by a consistency test.

```markdown
## Nhiệm vụ: Xây dựng OpenAPI 3.1 spec đầy đủ cho GoPaperless Workflow API

### Bối cảnh
Spec này là nguồn dữ liệu để tự động render tài liệu "GoPaperless Workflow RESTful API
Specification" (DOCX) ở project Agentic-Spec-Engine. Mọi trường tài liệu cần đều phải có
trong spec; thiếu trường nào thì cột đó trong tài liệu sẽ trống.

### Ràng buộc kiến trúc (bắt buộc)
- KHÔNG thêm JAX-RS, Spring, swagger-core annotation hay thay đổi routing servlet.
  Spec là file tĩnh, không sinh từ annotation runtime.
- KHÔNG đổi hành vi production để phục vụ spec (GOVERNANCE "Verification must not
  distort production design").
- Tuân theo docs/ai-memory/ (PROJECT_MEMORY, GOVERNANCE, CONVENTIONS). Đây là task LARGE:
  lập plan theo Plan Lifecycle trước khi viết.

### Vị trí file
- `docs/api/openapi.yaml` (một file gốc, có thể tách `components/` bằng `$ref` nếu quá lớn).

### Nguồn dữ liệu (theo thứ tự tin cậy)
1. Route: registry `validURLs` + `web.xml` của từng Controller — nguồn duy nhất cho
   method/path. Không thêm route không có trong registry.
2. Field/type: DTO request/response và serializer — đối chiếu với
   `src/test/resources/api-contract-snapshot.json`.
3. Mapping endpoint → DTO: đọc nhánh `doGet/doPost/doPut/doDelete` → Service → class
   Request được parse và Response/serializer được trả về.
4. Mô tả, M/O, sample: tài liệu DOCX hiện tại (bản V3 - 20260727) và Postman collection.
   Không suy đoán: field nào không xác định được M/O hoặc ý nghĩa thì đánh dấu
   `x-needs-review: true` kèm lý do, KHÔNG bịa.

### Yêu cầu nội dung cho MỖI operation
- `operationId` (camelCase, duy nhất), `summary` = đúng tên API như `loggingObject.apiName`.
- `description`: 1–3 câu, văn phong mô tả hành vi.
- `tags`: đúng một tag = category tài liệu (Heading 2), ví dụ "Authenticate - Log in".
- `x-controller`: tên Controller xử lý.
- `parameters`: header (Authorization, x-language-name, Content-Type…), path, query —
  mỗi cái có `description`, `required`, `schema`, `example`.
  Header dùng chung khai báo một lần trong `components.parameters` rồi `$ref`.
- `security`: khai báo scheme Bearer; endpoint public (authenticate…) ghi `security: []`.
- `requestBody` (nếu có): `$ref` tới schema + ít nhất một `example` đầy đủ, thực tế.
- `responses`:
  - `200` (hoặc mã thực tế): `$ref` schema + `example` thành công.
  - Lỗi: hệ thống trả lỗi trong body (`code`/`sub_code`), không theo HTTP status.
    Khai báo extension `x-error-codes: [{code, sub_code, constant, message_en}]`
    lấy từ hằng `*ErrorSubCode` thực sự được dùng trên luồng đó. Không bịa mã.

### Yêu cầu cho MỖI schema (components.schemas)
- Tên schema = tên class DTO.
- Mỗi property: `type` (+`format`), `description`, `example`.
- `required: [...]` phản ánh đúng validation thực tế trong code (Service/Facade check null).
- Enum Java → `enum` liệt kê đúng giá trị serialize ra JSON.
- Object lồng → `$ref`, không inline lặp lại. `List<X>` → `type: array, items: $ref X`.
- Field có trong DTO nhưng serializer không bao giờ ghi ra thì KHÔNG đưa vào response schema.

### Kiểm chứng (bắt buộc, chống lệch spec–code)
Viết JUnit test `OpenApiSpecConsistencyTest` (cùng package với `ApiContractSnapshotTest`):
1. Mọi route trong `validURLs` có operation tương ứng trong spec, và ngược lại.
2. Mọi schema request/response khớp field + type với `api-contract-snapshot.json`.
3. Spec hợp lệ OpenAPI 3.1 (dùng thư viện parser, chỉ scope test).
4. Lint: mọi operation có summary/description/tags; mọi property có description/example;
   báo danh sách `x-needs-review` (không fail, chỉ report).
Chạy: `mvn test -Dtest=OpenApiSpecConsistencyTest,ApiContractSnapshotTest`.

### Cách triển khai
- Làm theo từng Controller/category, mỗi category một commit.
- Sau mỗi category chạy test ở trên.
- Báo cáo cuối: số operation / tổng route, số schema, số field thiếu description/example,
  danh sách `x-needs-review`, các điểm code–docs mâu thuẫn phát hiện được
  (liệt kê, KHÔNG tự sửa code hay tài liệu).

### Không làm
- Không sửa logic, DTO, error code để "cho khớp spec".
- Không push/PR khi chưa được duyệt.
```

Before sending: source 4 needs access to the V3 DOCX (this repo) and the Postman
collection (currently only on the owner's machine). Without them, the agent
should leave descriptions empty and mark them `x-needs-review`.
