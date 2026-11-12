# Handoff: render DOCX cho release Gopaperless trên self-hosted runner

Ngày bàn giao: 2026-11-12 (Asia/Bangkok).

## Kết luận và phạm vi

Đã kiểm tra pipeline phát hành và chạy thử tạo DOCX từ tag `v4.1.1` trên PC hiện tại. Tạo file thành công, nhưng quality gate chưa đạt và chưa kiểm chứng thực thi dưới tài khoản service của runner. Chưa đủ điều kiện bật tự động xuất bản tài liệu.

Mục tiêu: mỗi release Gopaperless dùng OpenAPI của chính commit phát hành để tạo một DOCX mới, kiểm tra chất lượng, rồi gắn vào release cùng WAR và checksum. Không cần xây web server cho renderer.

Trong giai đoạn khảo sát và thử nghiệm, người dùng yêu cầu chỉ thao tác local: không push, dispatch workflow hay thay đổi release. Các giới hạn đó đã được tuân thủ. Sau thử nghiệm, người dùng cho phép tạo handoff này và push nhánh/PR lên Agentic-Spec-Engine để nghiên cứu tiếp. Quyền mới chỉ áp dụng cho bàn giao; không triển khai workflow hay sửa release Gopaperless.

Repository nhận PR này là public. Không kèm OpenAPI từ repo private, DOCX thử, dữ liệu contract, cấu hình triển khai hoặc credential. Những đầu ra thử nghiệm nằm trên PC, không thể tải chúng từ PR này.

## Trạng thái đã xác minh

### Runner và máy

- Runner `pccompany` đã được đăng ký và online khi khảo sát; không cần cài thêm runner để bắt đầu.
- Runner chạy dạng Windows Service, tự khởi động, tài khoản `NETWORK SERVICE`.
- Biến repository `CI_RUNNER=self-hosted` đang chọn runner này.
- Máy có Git, GitHub CLI và Python 3.12.10. Python hệ thống thiếu `lxml` trong preflight.
- Chế độ sleep khi cắm điện đã tắt. Ổ hệ thống còn khoảng 5,8 GiB lúc khảo sát; cần kiểm tra lại dung lượng trước khi bổ sung cache/dependencies.
- Lần thử dùng Python đi kèm ứng dụng, có `lxml`, `PyYAML`, `python-docx`; đây chưa phải môi trường CI được provision độc lập.
- Phiên chạy thử thuộc tài khoản người dùng tương tác, không có vai trò administrator. Không thử chuyển sang tài khoản service hay thay đổi service đang chạy.
- Khi tắt PC, runner sẽ offline và không nhận job cho tới khi máy/service hoạt động lại. Trạng thái online, dung lượng và biến CI phải kiểm tra lại ở lần tiếp tục.

### Pipeline Gopaperless

Đã đọc workflow remote ở nhánh mặc định `dev` và kiểm tra các bước release tương ứng trên `master`.

`ci-linux.yml`:

1. Push `master` kích hoạt build/test; tên workflow có Linux nhưng runner được chọn bằng `CI_RUNNER`.
2. Chạy `ApiContractSnapshotTest` và `OpenApiSpecConsistencyTest`, rồi package WAR và kiểm tra tập thư viện.
3. Nhánh self-hosted/master có bước chuẩn bị cấu hình runtime và rebuild WAR. Tài liệu không cần đọc các cấu hình đó.
4. `scripts/next-release-version.sh` tính version theo Conventional Commits: feat/minor, fix/patch, breaking/major; docs/chore/ci đơn thuần có thể không tạo release.
5. Bước `Auto-release from Conventional Commits` gọi `gh release create`, gắn WAR và SHA-256 vào release, target là commit đang build.
6. Nếu release đã tồn tại, logic hiện tại bỏ qua. Cần thiết kế riêng khả năng bổ sung DOCX thiếu khi retry.
7. Nhánh sự kiện `release: published` build/upload WAR cho release thủ công hoặc do actor khác tạo. Auto-release bằng `GITHUB_TOKEN` không kích hoạt tiếp nhánh này.
8. Job `checks` và `build-and-test` hiện độc lập; publish nằm trong `build-and-test`. Khi tích hợp, publish cần phụ thuộc tất cả gate bắt buộc.

`deploy-self-hosted.yml` phụ trách deploy `dev`/`prd`, tách khỏi luồng release tài liệu. Không thay đổi workflow deploy trong công việc này.

Release `v4.1.1` do `github-actions[bot]` tạo, có `workflow-api-4.1.1.war` và checksum. Đã kiểm tra lại sau thử nghiệm: vẫn chỉ có hai assets ban đầu. Release này trả về `immutable=false`; không suy ra chính sách đó áp dụng cho mọi release tương lai. Repo còn có release `vendor-libs-*`, cần loại khỏi luồng tài liệu sản phẩm.

## Renderer hiện có

- `scripts/openapi_to_contract.py`: OpenAPI YAML → contract JSON.
- `scripts/render_from_openapi.py`: seed DOCX + contract → DOCX mới.
- `scripts/check_frame_format.py`: gate định dạng bằng exit code.
- `scripts/extract_doc_contract.py` và `scripts/compare_openapi_doc.py`: đối chiếu nội dung.
- Seed đã dùng: `formatted/GoPaperless Workflow RESTful API Specification V3 - 20260727.docx`.
- Renderer giữ bố cục, nội dung ngoài endpoint và một phần nội dung seed; không dựng toàn bộ nội dung chỉ từ YAML. Endpoint không còn được phục vụ có thể được giữ lại với nhãn `Not implemented yet`.
- Renderer hỗ trợ version/ngày/History/tác giả, nhưng `--doc-version` chỉ nhận major.minor. Không truyền trực tiếp `4.1.1`; cần quy tắc version tài liệu riêng và ghi đủ tag phần mềm trong History.
- `requirements.txt` hiện dùng giới hạn `>=`, chưa khóa phiên bản cho môi trường CI tái lập.
- Mục lục hiện được yêu cầu cập nhật field khi mở bằng Word; chưa chứng minh số trang/mục lục được cập nhật sẵn trong artifact.

## Lần thử v4.1.1

### Nguồn và cách chạy

| Thành phần | Định danh |
|---|---|
| Tag backend | `v4.1.1` |
| Commit backend | `f31e32cb1cce1bdace1f11d1955c8e7617fba1c2` |
| File nguồn | `docs/api/openapi.yaml` tại commit trên |
| Git blob OpenAPI | `4dea1aefd18c51095ce6bf70957bda16f020213e` |
| Commit renderer | `044b30a2461e39c8aea41a4c1d4ea9929646a739` |
| Version tài liệu | `3.1`, ngày `20261112`; History ghi tag `v4.1.1` và commit nguồn |

Script local dùng `git archive` lấy scripts/seed đã commit để không trộn thay đổi working tree. OpenAPI được đọc qua GitHub API ở commit cố định; chỉ có các thao tác đọc remote. Contract và DOCX được sinh trong thư mục riêng. Có checksum cho đầu ra.

**Đây là local process trên cùng PC với runner, không phải một GitHub Actions run, và không phải phép thử dưới NETWORK SERVICE.** Không coi thử nghiệm này là đã xác nhận runner service có quyền đọc dependencies hoặc thực thi renderer.

### Kết quả chất lượng

| Kiểm tra | Kết quả |
|---|---|
| Tạo DOCX | Thành công, 3.433.016 byte |
| ZIP/XML | ZIP integrity hợp lệ, document XML parse được |
| Gate định dạng seed | 518 lỗi |
| Gate định dạng output | 518 lỗi, exit code 1 |
| So sánh lỗi raw | 55 finding thêm, 55 finding mất do tên heading đổi |
| So sánh sau chuẩn hóa đúng hậu tố ` (Not implemented yet)` | 0 lỗi mới, 0 lỗi được sửa |
| Tên bookmark trùng | 19, cùng tập với seed |
| Link nội bộ hỏng | 14 lượt, cùng tập với seed |
| Spec records sau mở rộng | 329 |
| Endpoint trích xuất từ DOCX | 353 |
| Section endpoint thêm mới | 88 |
| Section object thêm mới | 86 |
| Section đánh dấu chưa triển khai | 36 |
| Frame bị bỏ qua vì seed không phải JSON | 76 |

Báo cáo so sánh ghi nhận 320 cặp, 33 doc-only, 13 spec-only. Pairing không phải ánh xạ một-một; không tính tỷ lệ coverage bằng cách trừ các số này. Cần phân tích route gộp, alias và endpoint legacy trước khi kết luận thiếu API.

Có 77 chênh lệch kiểu request field và 62 chênh lệch kiểu response field. Một phần có thể là nhãn hiển thị/hyperlink alias, chưa được xác nhận ngữ nghĩa. Chỉ số `313 paired with no field difference` trong comparator không đưa type differences vào danh sách issues từng endpoint; không được hiểu là 313 endpoint khớp hoàn toàn.

**Visual QA chưa hoàn tất:** đã gọi `render_docx.py` của bộ công cụ Documents, nhưng thất bại với `FileNotFoundError: LibreOffice soffice.exe was not found on PATH`. Chưa tạo ảnh trang, chưa duyệt bố cục. DOCX là bản thử cần kiểm tra, không phải bản đủ điều kiện phát hành.

### Checksum SHA-256

```text
OpenAPI: 561d57b8376ab17bc5732b82222d974af7144a30eeca84c80181063b038f0c21
Seed:    f73a7202fbca57c5befd1a6654325d0c01f765e83d200ff146b1bf49778bea9f
DOCX:    2dd105eb56939fd00fe8a8fd4b926394531e3ca0bfe98907c2779c09ffd4705f
```

### Các file chỉ có trên PC

Trong checkout Agentic-Spec-Engine, thư mục untracked `formatted/dry-run-v4.1.1/` chứa:

- `GoPaperless-API-v4.1.1.docx` và `.sha256`.
- `REPORT.md`, `manifest.json`, `quality.json`, `render-stats.json`.
- `seed-format.txt`, `output-format.txt`, `comparison.md`, `comparison.log`.
- `openapi.yaml`, `contract.json`, `doc-contract.json`.
- `engine.zip`, `engine/`: snapshot renderer và seed.
- `run_local.py`, `summarize.py`: driver và tổng hợp kết quả thử. Driver từ chối ghi đè lần chạy đã có manifest.

Không xóa các file này khi dọn workspace. Các file untracked đã tồn tại trước thử nghiệm cũng được giữ nguyên. Không có sửa đổi tracked vào renderer hay workflow trong giai đoạn thử.

## Thiết kế tích hợp đề xuất, chưa triển khai

```text
Push master
  → checks + build/test + OpenAPI gate + WAR
  → xác định có release và version/tag
  → checkout renderer ở commit cố định, seed cố định
  → OpenAPI của commit phát hành → contract → DOCX
  → quality gates + checksum
  → publish WAR + DOCX + checksum
```

Ưu tiên tích hợp trực tiếp vào job/pipeline release hiện có thay vì chỉ tạo listener `release: published`. Chỉ render khi version resolver xác định thực sự cần phát hành. Chặn publish nếu DOCX không đạt gate. Nếu cần tính nguyên khối, tạo draft, upload đầy đủ rồi publish.

- Workflow thuộc repo Gopaperless; Agentic-Spec-Engine là công cụ checkout vào thư mục con, pin commit, `persist-credentials: false`.
- Engine public nên không cần credential đọc repo private bổ sung. Upload assets dùng token của repo Gopaperless với quyền cần thiết; không nhúng token trong script/log.
- Chỉ xử lý tag sản phẩm `vX.Y.Z`, không xử lý vendor release.
- Dùng thư mục output riêng mỗi lần chạy; không dùng checkout cá nhân làm CI workspace.
- Ghi provenance: commit backend, commit renderer, seed hash, dependency versions, tag và ngày tài liệu.
- Retry phân biệt release hoàn chỉnh, release thiếu tài liệu và asset đã tồn tại; không thay WAR khi chỉ chạy bù DOCX.
- Hỗ trợ manual dry run chỉ tạo artifact; mặc định không publish. Chạy bù dùng đúng tag, không lấy OpenAPI từ nhánh mới nhất.
- Không mặc định dùng output lần trước làm seed lần sau: tránh tích lũy nội dung/history và mất tính tái lập.
- Với nhiều push master gần nhau, serialize đoạn quyết định version/publish và kiểm tra lại trạng thái tag để tránh race.
- Kiểm tra lại timeout job 30 phút hiện có sau khi đo thời gian render/QA; tránh hủy job đang upload.
- Không đưa cấu hình runtime, signing material hoặc WAR vào artifact chẩn đoán tài liệu.

## Thứ tự tiếp tục khi mở máy lại

1. Đọc handoff này và các báo cáo local. Kiểm tra runner online, dung lượng đĩa, working tree; không tự động chạy release/deploy.
2. Tạo môi trường Python CI riêng với dependencies khóa phiên bản và quyền phù hợp cho service. Thử offline/local dưới tài khoản runner bằng cơ chế quản trị máy phù hợp; chưa cần dispatch remote.
3. Sửa bản sao seed để xử lý 518 lỗi định dạng và lỗi bookmark/link, chạy lại gate. Không bỏ qua exit code hoặc chỉ so tổng số lỗi.
4. Rà soát type alias, endpoint ghép/gộp, 76 frame không phải JSON, chính sách giữ/bỏ endpoint legacy và version tài liệu.
5. Chuẩn bị môi trường render trang, kiểm tra bố cục/mục lục. Thiếu visual QA là một điều kiện chưa hoàn thành, không phải pass.
6. Chạy lại v4.1.1 trong thư mục output mới với snapshot nguồn cố định; giữ lần thử này làm đối chứng.
7. Khi dry run đạt và được yêu cầu triển khai tiếp, sửa workflow Gopaperless qua PR riêng; thử chế độ chỉ tạo artifact trước, sau đó mới bật gắn vào release.

Tiêu chí hoàn tất: chạy được dưới service account, nguồn đúng tag, quality gate rõ ràng, bố cục đã duyệt, log/provenance đủ, retry an toàn, và release chỉ publish khi tất cả kiểm tra bắt buộc thành công.

## Tài liệu tham khảo

- [GitHub self-hosted runners](https://docs.github.com/en/actions/reference/runners/self-hosted-runners): runner chủ động kết nối HTTPS; không cần mở inbound port/ngrok chỉ để nhận job.
- [Triggering a workflow](https://docs.github.com/en/actions/how-tos/write-workflows/choose-when-workflows-run/trigger-a-workflow): giới hạn sự kiện sinh bởi GITHUB_TOKEN.
- [GITHUB_TOKEN](https://docs.github.com/en/actions/concepts/security/github_token): phạm vi token theo repository.
- Báo cáo nền trong engine: `reports/render_openapi_vs_v3_20261110.md` và `docs/openapi-to-v3-render-plan.md`. Một số phần kế hoạch cũ mô tả renderer chưa có; ưu tiên mã hiện tại và kết quả thực đo ở trên.
