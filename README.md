# Dữ liệu văn bản pháp luật – Trợ lý AI pháp lý Thuế – Kế toán

Phần **chuẩn bị dữ liệu**: thu thập, kiểm kê, OCR, tách cấu trúc, chunking và đối chiếu với văn bản gốc cho bộ 65 văn bản pháp luật thuế – kế toán (cập nhật 8/2026). Đây là nền cho WP1 (Parser), WP2 (Retrieval), WP3 (RAG), WP4 (Temporal), WP5 (Knowledge Graph).

## Kết quả chính

| Hạng mục | Kết quả |
|---|---|
| Văn bản | 65/65 (đã mở trang 1 bản gốc xác nhận đúng số hiệu từng văn bản) |
| Định dạng gốc | 1 Word gốc · 2 PDF convert từ Word · 7 PDF có text · 55 PDF chụp hình (3.421 trang, đã OCR) |
| Cấu trúc | 2.335/2.355 Điều (99,2%) · 8.274 Khoản · 6.719 Điểm · 792 phụ lục/biểu mẫu |
| Chunk | 9.833 chunk legal-structure (≤ 512 token bge-m3) · 4.524 chunk fixed-token |
| Đối chiếu bản gốc | 99,95% chunk khớp nguyên văn · 99,6% đúng trang · mẫu 30 chunk: 30/30 đúng nội dung, 28/30 đúng cấu trúc |
| OCR | lỗi ký tự trung vị 2,2% · độ tin cậy trung bình 88,5/100 |

## Cấu trúc thư mục

```
raw/                 văn bản gốc (10 nhóm) + danh mục + dữ liệu đã đối chiếu tay
src/                 code xử lý dữ liệu, schema.py, tim_kiem.py
tests/               bộ kiểm thử parser (pytest)
docs/                HUONG_DAN_NHOM_2.md – mô tả dữ liệu cho nhóm cài đặt
output/
  Kiem_ke_van_ban.xlsx                 file Excel tổng hợp (12 sheet)
  Bao_cao_cau_truc_va_chunking.docx    báo cáo
  text/                                text từng văn bản (OCR / lớp chữ PDF)
  structure/                           văn bản có cấu trúc (JSON schema WP1)
  chunks/                              legal_structure_chunks.jsonl, fixed_token_chunks.jsonl
  quan_he_van_ban.json                 quan hệ AMENDS / REPLACES / GUIDES ...
  schema/                              JSON Schema
  doi_chieu/                           kết quả đối chiếu từng chunk
```

## Chạy

Xem `HUONG_DAN.txt`. Tóm tắt:

```bash
pip install -r requirements.txt
python src/kiem_ke.py && python src/trich_text.py && python src/tach_van_ban.py
python src/chunking.py && python src/doi_chieu.py
python -m pytest tests -v
```

## Git

Bấm đúp `tao_git.bat` (cần cài Git) để tạo repo và commit đầu tiên, rồi:

```bash
git remote add origin <link repo của nhóm>
git push -u origin main
```
