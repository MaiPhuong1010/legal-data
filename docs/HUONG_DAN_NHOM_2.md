# Hướng dẫn dùng dữ liệu cho Nhóm 2

Tài liệu này mô tả dữ liệu phần chuẩn bị đã tạo ra, để nhóm 2 cài đặt **WP2 (Retrieval)**, **WP3 (RAG)**, **WP4 (Temporal)** và **WP5 (Knowledge Graph)** mà không phải xử lý lại văn bản gốc.

## 1. Các file cần dùng

| File | Nội dung | Dùng cho |
|---|---|---|
| `output/chunks/legal_structure_chunks.jsonl` | **9.833 chunk** theo cấu trúc luật, mỗi dòng 1 JSON, mỗi chunk ≤ 512 token (bge-m3) | WP2 tạo vector / BM25, WP3 dựng context |
| `output/chunks/fixed_token_chunks.jsonl` | 4.524 chunk cắt cứng 512 token, overlap 50 | WP3 baseline để so sánh |
| `output/structure/<document_id>.json` | 65 văn bản có cấu trúc theo schema WP1: metadata + `legal_units` (Điều/Khoản/Điểm/Phụ lục) | Tra nguyên văn, dựng context lân cận, WP1 |
| `output/quan_he_van_ban.json` | Quan hệ giữa văn bản: AMENDS, REPLACES, PARTIALLY_REPEALS, GUIDES, REFERENCES | WP5 dựng đồ thị, WP4 version chain |
| `output/schema/*.schema.json` + `src/schema.py` | JSON Schema chuẩn (Pydantic) của mọi file trên | Kiểm tra dữ liệu khi nạp |
| `output/Kiem_ke_van_ban.xlsx` | Số liệu kiểm kê, thống kê token, đối chiếu chất lượng | Báo cáo |

## 2. Một chunk gồm những gì

```json
{
  "chunk_id": "ND254_2026_D9_K2",
  "chunk_type": "clause",
  "document_no": "254/2026/NĐ-CP",
  "document_type": "Nghị định",
  "chapter": "II", "article": 9, "clause": 2, "point": null,
  "unit_ids": ["ND254_2026_D9_K2"],
  "parent_id": "ND254_2026_D9",
  "context_header": "[Nghị định 254/2026/NĐ-CP > Chương II ... > Điều 9. Thời điểm lập hóa đơn > Khoản 2]",
  "content": "2. Thời điểm lập hóa đơn đối với cung cấp dịch vụ là ...",
  "text_for_embedding": "<context_header>\n<content>",
  "token_count": 161,
  "page": 16, "page_end": 16,
  "effective_from": "2026-07-01", "effective_to": null, "status": "active",
  "is_confident": true,
  "issuer": "Chính phủ", "issued_date": "2026-06-30",
  "amended_by": [], "replaced_by": [],
  "source_file": "raw/II. Hoa don dien tu/ND254_2026.pdf", "source_url": "https://..."
}
```

- **Tạo vector:** dùng trường `text_for_embedding` (đã có header ngữ cảnh). Không embed riêng `content`.
- **Trích dẫn (WP3/WP8):** `unit_id` = `chunk_id` (bỏ hậu tố `_partN` nếu có), `document_no`, `article`, `clause`, `point`. `quote_text` lấy từ `content`.
- **`chunk_type`:** `article` (cả Điều), `clause`, `point`, `article_intro`/`clause_intro` (câu dẫn), `preamble`, `appendix`, `table`. Hậu tố `_split` nghĩa là đơn vị quá dài nên đã cắt nhỏ, khi đó `chunk_id` có thêm `_part1`, `_part2`…
- **`is_confident = false`** (khoảng 1,8%): văn bản OCR có nhiều ký tự lạ hoặc số thứ tự bị nhảy. Nên ưu tiên thấp hoặc cho người kiểm duyệt.

## 3. Lọc theo hiệu lực (WP4)

Mỗi chunk có `effective_from`, `effective_to`, `status`. Một số Điều có **hiệu lực riêng** khác toàn văn bản (ví dụ Luật 108/2025/QH15: toàn luật từ 01/7/2026, nhưng Điều 13 từ 01/01/2026). Trường hợp này đã ghi ở mức chunk, kèm `effective_note` trong file structure.

```python
def con_hieu_luc(c, ngay):          # ngay = "2026-08-15"
    if c["effective_from"] and c["effective_from"] > ngay: return False
    if c["effective_to"] and c["effective_to"] < ngay: return False
    return c["status"] != "replaced"
```

Văn bản bị sửa đổi nằm trong `amended_by` (ví dụ Nghị định 125/2020 được sửa bởi 310/2025 và 291/2026). Việc áp dụng nội dung sửa đổi vào từng Khoản là phần việc của WP4.

## 4. Đồ thị quan hệ (WP5)

`output/quan_he_van_ban.json` là danh sách `{source, relation, target, method, verified, evidence}`:

- `verified = true`: lấy từ danh mục của thầy hoặc đã đối chiếu bản gốc (65 quan hệ).
- `verified = false`: trích tự động bằng regex. Nên kiểm tra trước khi dùng, bằng chứng nằm ở trường `evidence`.
- Cạnh `CONTAINS` (Văn bản → Điều → Khoản → Điểm) dựng từ `parent_id` trong file structure.

## 5. Chạy thử

```bash
pip install -r requirements.txt rank-bm25
python src/tim_kiem.py "thời điểm lập hóa đơn đối với cung cấp dịch vụ" --ngay 2026-08-15
# -> kết quả top 1: ND254_2026_D9_K2 (đúng ví dụ trong đề)
python src/tim_kiem.py --tao-vector      # tạo vector bge-m3 (cần PyTorch + sentence-transformers)
```

Kiểm tra dữ liệu: `python -m pytest tests -v` và `python src/doi_chieu.py`.

## 6. Hạn chế cần biết

- 55/65 văn bản là bản scan nên phải OCR. Tỷ lệ lỗi ký tự trung vị khoảng 2,2% (lỗi từ khoảng 4,7%). Trang có bảng biểu lỗi nhiều hơn. Lỗi chính tả OCR ảnh hưởng BM25 nhiều hơn tìm kiếm vector.
- 20/2.355 Điều chưa nhận diện được (tiêu đề bị OCR làm hỏng). Nội dung của các Điều này nằm chung với Điều đứng trước.
- Bảng biểu trong văn bản bị trải phẳng thành dòng chữ, không giữ được cột.
- TT133/2016 là bản Word lấy từ nguồn không chính thức (đã kiểm tra số hiệu, ngày ban hành); Word không có số trang nên page = 1.
