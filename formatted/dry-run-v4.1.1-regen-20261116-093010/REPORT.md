# Dry-run v4.1.1 (REGENERATE) — 2026-11-16

Tái tạo lại lần chạy thử v4.1.1 đã mô tả trong
`docs/handoff-self-hosted-docx-release-20261112.md`. Bản `formatted/dry-run-v4.1.1/`
gốc đã bị xóa nhầm ngày 2026-11-16; thư mục này là bản dựng lại từ **đúng 3 input đã pin**.

## Kết luận
- **Tái lập trung thực.** 2/3 input khớp checksum nguyên văn với handoff; toàn bộ chỉ số
  chủ chốt (format gate, comparison, số section mới) **trùng khít** bảng "Lần thử v4.1.1".
- **Vẫn KHÔNG đủ điều kiện phát hành**: `check_frame_format` exit 1, 518 lỗi format thừa kế
  từ seed, visual QA chưa làm. Khớp đúng trạng thái "chưa production-ready" của handoff.

## Đầu vào đã pin (bytes-identical với handoff)
| Thành phần | Định danh | Khớp handoff |
|---|---|---|
| Backend tag / commit | `v4.1.1` / `f31e32cb…` | — |
| openapi.yaml SHA-256 | `561d57b8…f0c21` | ✅ |
| Seed V3 SHA-256 | `f73a7202…bea9f` | ✅ |
| Renderer commit | `044b30a2…646a739` | ✅ |
| Output DOCX SHA-256 | `2f272552…f98a` | ❌ (bản gốc `2dd105eb…`; lệch do mtime zip + core.xml + dòng History) |

## Lệnh tái lập
```bash
# 1) fetch openapi.yaml ở commit backend đã pin
gh api -H "Accept: application/vnd.github.raw+json" \
  "repos/tatkhanhgia/GoPaperless-Workflow-API/contents/docs/api/openapi.yaml?ref=f31e32cb1cce1bdace1f11d1955c8e7617fba1c2" \
  > openapi.yaml

# 2) engine snapshot @ 044b30a
git archive 044b30a scripts requirements.txt | tar -x -C engine/

# 3) contract -> render -> gate -> compare
python engine/scripts/openapi_to_contract.py openapi.yaml -o contract.json
python engine/scripts/render_from_openapi.py "<seed V3>" contract.json GoPaperless-API-v4.1.1.docx \
  --doc-version 3.1 --doc-date 20261112 --author GIATK
python engine/scripts/check_frame_format.py GoPaperless-API-v4.1.1.docx
python engine/scripts/extract_doc_contract.py GoPaperless-API-v4.1.1.docx -o doc-contract.json
python engine/scripts/compare_openapi_doc.py contract.json doc-contract.json -o comparison.md
```

## Số liệu tái lập
**Format gate:** seed 518 / output 518 / exit 1 — 0 lỗi mới so với seed.
**Comparison:** spec 329 · doc 353 · paired 320 · doc-only 33 · spec-only 13 · paired-không-lệch-field 313 · request type diff 77 · response type diff 62.
**Render:** section endpoint mới 88 · object mới 86 · not-implemented 36 · frame bỏ qua (seed không JSON) 76.

## File trong thư mục
`openapi.yaml` · `contract.json` · `doc-contract.json` · `GoPaperless-API-v4.1.1.docx` (+ `.sha256`)
· `seed-format.txt` · `output-format.txt` · `comparison.md` · `comparison.log`
· `manifest.json` · `quality.json` · `render-stats.json` · `REPORT.md` · `engine/` · `engine.zip`

## Khác biệt so với bản gốc
Bản gốc có thêm driver tự viết `run_local.py` / `summarize.py` — **mất vĩnh viễn** (không pin ở
đâu, không tái tạo tự động). Các artifact còn lại ở trên thay thế đầy đủ cho mục đích đối chứng.
