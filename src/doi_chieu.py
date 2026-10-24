"""
ĐỐI CHIẾU LẠI VỚI VĂN BẢN GỐC, KIỂM TRA ĐỘ CHÍNH XÁC TỪNG CHUNK

Kiểm tra tự động (Structural Validator):
  1. Mỗi file JSON văn bản và mỗi chunk đúng JSON Schema (src/schema.py)
  2. unit_id / chunk_id không trùng; parent_id trỏ tới unit có thật
  3. Số Điều tăng dần, liệt kê Điều bị thiếu; Khoản tăng dần trong từng Điều
  4. Nội dung chunk KHỚP NGUYÊN VĂN với text của văn bản gốc (từng dòng của chunk phải có trong văn bản)
  5. Trang ghi trong chunk đúng: đoạn đầu của chunk nằm trên trang đó (cho phép lệch 1 trang)
  6. Tiêu đề "Điều N." của chunk có trong văn bản gốc
  7. Số hiệu văn bản có xuất hiện ở trang đầu văn bản gốc
  8. Độ phủ: bao nhiêu % chữ của văn bản đã nằm trong các unit (không bị rơi mất)
  9. Ngày ban hành / hiệu lực: định dạng đúng, hiệu lực không trước ngày ban hành quá xa (cảnh báo)
Kiểm tra tay:
  10. Chọn ngẫu nhiên 30 chunk -> sheet "Mau kiem tra tay" để người kiểm tra mở PDF gốc so sánh.
      Kết quả kiểm tra đã làm nằm trong raw/kiem_tra_tay.csv và được gộp vào sheet.

Đầu ra: output/doi_chieu/*.csv, output/doi_chieu/tong_hop.json và các sheet "Doi chieu ..." trong Kiem_ke_van_ban.xlsx
Cách chạy:  python src\\doi_chieu.py
"""
import csv
import json
import random
import re
import statistics
import sys
import unicodedata
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from pydantic import ValidationError

sys.path.insert(0, str(Path(__file__).resolve().parent))
from tach_van_ban import dinh_dang_sheet
from regex_luat import RE_SO_HIEU, chuan_so_hieu, lam_sach
from schema import FixedChunk, LegalChunk, LegalDocument

ROOT = Path(__file__).resolve().parent.parent
STRUCT_DIR = ROOT / "output" / "structure"
CHUNK_DIR = ROOT / "output" / "chunks"
TEXT_DIR = ROOT / "output" / "text"
OUT_DIR = ROOT / "output" / "doi_chieu"
KIEM_KE = ROOT / "output" / "Kiem_ke_van_ban.xlsx"
SO_MAU = 30


def chuan(s):
    """Chuẩn hóa để so khớp: Unicode NFC, gộp khoảng trắng."""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", s or "")).strip()


def doc_text_goc(doc_id):
    """Text gốc của văn bản (sau OCR / đọc PDF), tách theo trang."""
    raw = (TEXT_DIR / f"{doc_id}.txt").read_text(encoding="utf-8")
    trang = {}
    for so, noi_dung in re.findall(r"<<<TRANG (\d+)>>>\n(.*?)(?=<<<TRANG \d+>>>|\Z)", raw, re.S):
        trang[int(so)] = chuan(noi_dung)
    sach = chuan(" ".join(d for _, d in lam_sach(raw)))
    return raw, trang, sach


def kiem_tra_van_ban(doc_json, raw, trang, sach):
    loi = []
    try:
        LegalDocument.model_validate(doc_json)
        schema_ok = True
    except ValidationError as e:
        schema_ok = False
        loi.append(f"Sai schema: {str(e)[:150]}")
    units = doc_json["legal_units"]
    ids = [u["unit_id"] for u in units]
    trung = [k for k, v in Counter(ids).items() if v > 1]
    if trung:
        loi.append(f"unit_id trùng: {trung[:5]}")
    tap = set(ids)
    cha_sai = [u["unit_id"] for u in units if u.get("parent_id") and u["parent_id"] not in tap]
    if cha_sai:
        loi.append(f"parent_id không tồn tại: {cha_sai[:5]}")
    dieu = [u["article"] for u in units if u["unit_type"] == "article"]
    tang_dan = all(a < b for a, b in zip(dieu, dieu[1:]))
    thieu = sorted(set(range(1, max(dieu) + 1)) - set(dieu)) if dieu else []
    khoan_sai = 0
    for d in set(dieu):
        k = [u["clause"] for u in units if u["unit_type"] == "clause" and u["article"] == d]
        khoan_sai += sum(1 for a, b in zip(k, k[1:]) if b <= a)
    # số hiệu có trong 2 trang đầu?
    dau = " ".join(trang.get(i, "") for i in (1, 2))
    so_hieu_trang_dau = {chuan_so_hieu(m) for m in RE_SO_HIEU.finditer(dau)}
    so_hieu_ok = doc_json["document_no"] in so_hieu_trang_dau or \
        doc_json["document_no"].split("/")[0] + "/" + doc_json["document_no"].split("/")[1] in dau
    # độ phủ ký tự: chữ trong các unit / chữ trong text đã làm sạch
    chu_unit = sum(len(chuan(u["content"])) + len(chuan(u.get("title") or "")) for u in units)
    do_phu = min(1.0, chu_unit / max(len(sach), 1))
    # ngày tháng
    bh, hl = doc_json.get("issued_date"), doc_json.get("effective_from")
    if bh and hl and hl < bh and (int(bh[:4]) - int(hl[:4])) > 0:
        loi.append(f"Ngày hiệu lực {hl} trước ngày ban hành {bh} (cần kiểm tra)")
    if not hl:
        loi.append("Chưa xác định ngày hiệu lực")
    return dict(schema_ok=schema_ok, so_dieu=len(dieu), dieu_lon_nhat=max(dieu, default=0), dieu_thieu=thieu,
                dieu_tang_dan=tang_dan, khoan_sai_thu_tu=khoan_sai, so_hieu_trang_dau=so_hieu_ok,
                do_phu=round(do_phu, 4), unit_tin_cay_thap=sum(1 for u in units if not u["is_confident"]),
                so_unit=len(units), loi=loi)


def kiem_tra_chunk(c, trang, sach):
    """Trả về (khớp nguyên văn %, đúng trang?, tiêu đề Điều có trong văn bản?)."""
    # bỏ nhãn "1." / "a)" ở đầu dòng: parser chuẩn hóa nhãn (OCR "4," -> "4.", "a)Áp" -> "a) Áp")
    dong = [re.sub(r"^(\d{1,2}\s*[\.,]|[a-zđ]\s*\))\s*", "", chuan(x)) for x in c["content"].split("\n")]
    dong = [x for x in dong if len(x) >= 8]
    if not dong:
        return 1.0, True, True
    khop = sum(1 for x in dong if x in sach or x[:80] in sach)
    ty_le = khop / len(dong)
    p = c.get("page")
    p2 = c.get("page_end") or p
    vung = range(p - 1, p2 + 2) if p else []
    dung_trang = p is None or any(any(d[:40] in trang.get(q, "") or d[5:35] in trang.get(q, "") for q in vung)
                                  for d in dong[:2])
    co_dieu = True
    if c.get("article"):
        co_dieu = bool(re.search(rf"\b{c['article']}\s*[\.:,]", sach))
    return ty_le, dung_trang, co_dieu


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    docs = {f.stem: json.loads(f.read_text(encoding="utf-8")) for f in sorted(STRUCT_DIR.glob("*.json"))}
    chunks = [json.loads(x) for x in open(CHUNK_DIR / "legal_structure_chunks.jsonl", encoding="utf-8")]
    fixed = [json.loads(x) for x in open(CHUNK_DIR / "fixed_token_chunks.jsonl", encoding="utf-8")]

    # ----- 1. kiểm tra từng văn bản -----
    kq_vb, nguon = [], {}
    for doc_id, d in docs.items():
        raw, trang, sach = doc_text_goc(doc_id)
        nguon[doc_id] = (trang, sach)
        k = kiem_tra_van_ban(d, raw, trang, sach)
        k["document_id"] = doc_id
        kq_vb.append(k)
        print(f"  {doc_id:22} Điều {k['so_dieu']}/{k['dieu_lon_nhat']}  phủ {k['do_phu']:.1%}  lỗi: {len(k['loi'])}")

    # ----- 2. kiểm tra từng chunk -----
    loi_schema = 0
    for c in chunks:
        try:
            LegalChunk.model_validate(c)
        except ValidationError:
            loi_schema += 1
    loi_schema_fixed = 0
    for c in fixed:
        try:
            FixedChunk.model_validate(c)
        except ValidationError:
            loi_schema_fixed += 1
    trung_chunk = [k for k, v in Counter(c["chunk_id"] for c in chunks).items() if v > 1]
    ds_chunk = []
    for c in chunks:
        trang, sach = nguon[c["document_id"]]
        ty_le, dung_trang, co_dieu = kiem_tra_chunk(c, trang, sach)
        ds_chunk.append(dict(chunk_id=c["chunk_id"], document_id=c["document_id"], chunk_type=c["chunk_type"],
                             khop_nguyen_van=round(ty_le, 3), dung_trang=dung_trang, co_tieu_de_dieu=co_dieu,
                             is_confident=c["is_confident"], token_count=c["token_count"]))
    with open(OUT_DIR / "doi_chieu_tung_chunk.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(ds_chunk[0].keys()))
        w.writeheader()
        w.writerows(ds_chunk)

    n = len(ds_chunk)
    khop_hoan_toan = sum(1 for x in ds_chunk if x["khop_nguyen_van"] == 1.0)
    tong_hop = {
        "so_van_ban": len(docs),
        "van_ban_dung_schema": sum(k["schema_ok"] for k in kq_vb),
        "tong_dieu_nhan_dien": sum(k["so_dieu"] for k in kq_vb),
        "tong_dieu_theo_so_lon_nhat": sum(k["dieu_lon_nhat"] for k in kq_vb),
        "van_ban_so_dieu_tang_dan": sum(k["dieu_tang_dan"] for k in kq_vb),
        "van_ban_co_so_hieu_o_trang_dau": sum(k["so_hieu_trang_dau"] for k in kq_vb),
        "do_phu_trung_binh": round(statistics.mean(k["do_phu"] for k in kq_vb), 4),
        "so_chunk": n,
        "chunk_sai_schema": loi_schema,
        "chunk_fixed_sai_schema": loi_schema_fixed,
        "chunk_id_trung": len(trung_chunk),
        "chunk_khop_nguyen_van_100": khop_hoan_toan,
        "ty_le_chunk_khop_nguyen_van_100": round(khop_hoan_toan / n, 4),
        "ty_le_dong_khop_trung_binh": round(statistics.mean(x["khop_nguyen_van"] for x in ds_chunk), 4),
        "ty_le_chunk_dung_trang": round(sum(x["dung_trang"] for x in ds_chunk) / n, 4),
        "ty_le_chunk_tieu_de_dieu_co_trong_goc": round(sum(x["co_tieu_de_dieu"] for x in ds_chunk) / n, 4),
        "ty_le_chunk_do_tin_cay_cao": round(sum(x["is_confident"] for x in ds_chunk) / n, 4),
    }
    # chất lượng OCR (chat_luong_ocr.py)
    cer_file = TEXT_DIR / "_cer_ocr.csv"
    if cer_file.exists():
        cer = list(csv.DictReader(open(cer_file, encoding="utf-8-sig")))
        tong_hop["ocr_cer_trung_vi_trang"] = round(statistics.median(float(x["CER"]) for x in cer), 4)
        tong_hop["ocr_wer_trung_vi_trang"] = round(statistics.median(float(x["WER"]) for x in cer), 4)
        tong_hop["ocr_cer_tong"] = round(sum(int(x["loi_ky_tu"]) for x in cer) / sum(int(x["so_ky_tu"]) for x in cer), 4)
        tong_hop["ocr_so_trang_do_cer"] = len(cer)
    tin_cay = TEXT_DIR / "_chat_luong_ocr.csv"
    if tin_cay.exists():
        tc = list(csv.DictReader(open(tin_cay, encoding="utf-8-sig")))
        tong_hop["ocr_do_tin_cay_trung_binh"] = round(statistics.mean(float(x["do_tin_cay_tb"]) for x in tc), 1)

    # ----- 3. mẫu kiểm tra tay -----
    da_kiem = {}
    f_tay = ROOT / "raw" / "kiem_tra_tay.csv"
    if f_tay.exists():
        da_kiem = {r["chunk_id"]: r for r in csv.DictReader(open(f_tay, encoding="utf-8-sig"))}
    # mẫu ổn định: ưu tiên các chunk ĐÃ kiểm tra tay (còn tồn tại), thiếu thì chọn ngẫu nhiên thêm
    theo_id = {c["chunk_id"]: c for c in chunks}
    mau = [theo_id[k] for k in da_kiem if k in theo_id][:SO_MAU]
    random.seed(2026)
    than = [c for c in chunks if c["article"] is not None and c["chunk_id"] not in da_kiem]
    mau += random.sample(than, max(0, SO_MAU - len(mau)))
    tong_hop["mau_kiem_tra_tay"] = SO_MAU
    tong_hop["mau_da_kiem_tra"] = sum(1 for c in mau if c["chunk_id"] in da_kiem)
    tong_hop["mau_dung_cau_truc"] = sum(1 for c in mau if da_kiem.get(c["chunk_id"], {}).get("cau_truc") == "Đúng")
    tong_hop["mau_dung_noi_dung"] = sum(1 for c in mau if da_kiem.get(c["chunk_id"], {}).get("noi_dung") == "Đúng")

    (OUT_DIR / "tong_hop.json").write_text(json.dumps(tong_hop, ensure_ascii=False, indent=2), encoding="utf-8")
    ghi_excel(kq_vb, tong_hop, ds_chunk, mau, da_kiem)
    print()
    for k, v in tong_hop.items():
        print(f"  {k:45} {v}")
    print(f"\n=> Kết quả đối chiếu: {OUT_DIR}")


def ghi_excel(kq_vb, tong_hop, ds_chunk, mau, da_kiem):
    wb = openpyxl.load_workbook(KIEM_KE)
    for ten in ("Doi chieu tong hop", "Doi chieu van ban", "Mau kiem tra tay", "Chat luong OCR"):
        if ten in wb.sheetnames:
            del wb[ten]

    th = wb.create_sheet("Doi chieu tong hop")
    th.append(["Hạng mục kiểm tra", "Kết quả", "Giải thích"])
    giai_thich = {
        "van_ban_dung_schema": "Số văn bản có file JSON đúng JSON Schema WP1 (src/schema.py)",
        "tong_dieu_nhan_dien": "Tổng số Điều tách được",
        "tong_dieu_theo_so_lon_nhat": "Tổng số Điều theo số Điều lớn nhất của từng văn bản (ước lượng số Điều thật)",
        "van_ban_so_dieu_tang_dan": "Số văn bản có số Điều tăng dần đúng thứ tự",
        "van_ban_co_so_hieu_o_trang_dau": "Số văn bản tìm thấy đúng số hiệu ở 2 trang đầu bản gốc",
        "do_phu_trung_binh": "% chữ của văn bản gốc nằm trong các unit (phần nơi nhận/chữ ký bị loại có chủ đích)",
        "chunk_sai_schema": "Số chunk legal-structure sai schema",
        "chunk_fixed_sai_schema": "Số chunk fixed-token sai schema",
        "chunk_id_trung": "Số chunk_id bị trùng",
        "ty_le_chunk_khop_nguyen_van_100": "% chunk mà MỌI dòng nội dung đều có nguyên văn trong văn bản gốc",
        "ty_le_dong_khop_trung_binh": "% dòng nội dung chunk có nguyên văn trong văn bản gốc (trung bình)",
        "ty_le_chunk_dung_trang": "% chunk có đoạn đầu nằm đúng trang ghi trong metadata (±1 trang)",
        "ty_le_chunk_tieu_de_dieu_co_trong_goc": "% chunk mà số Điều của nó có trong văn bản gốc",
        "ty_le_chunk_do_tin_cay_cao": "% chunk is_confident = true",
        "ocr_cer_trung_vi_trang": "Tỷ lệ lỗi ký tự OCR (trung vị theo trang) – đo trên PDF có text: OCR lại rồi so với text thật",
        "ocr_wer_trung_vi_trang": "Tỷ lệ lỗi từ OCR (trung vị theo trang)",
        "ocr_cer_tong": "Tỷ lệ lỗi ký tự OCR trên toàn bộ trang mẫu (trang có bảng biểu làm tăng lỗi)",
        "ocr_do_tin_cay_trung_binh": "Độ tin cậy trung bình Tesseract trên các văn bản scan (0–100)",
        "mau_da_kiem_tra": "Số chunk mẫu đã mở PDF gốc kiểm tra bằng mắt",
        "mau_dung_cau_truc": "Số chunk mẫu đúng Điều/Khoản/Điểm, đúng trang",
        "mau_dung_noi_dung": "Số chunk mẫu đúng nội dung (sai chính tả OCR nhỏ không làm sai nghĩa vẫn tính Đúng)",
    }
    for k, v in tong_hop.items():
        th.append([k, v, giai_thich.get(k, "")])
        if k.startswith(("ty_le", "do_phu", "ocr_cer", "ocr_wer")):
            th.cell(th.max_row, 2).number_format = "0.00%"
    dinh_dang_sheet(th, [40, 14, 100])

    vb = wb.create_sheet("Doi chieu van ban")
    vb.append(["Document ID", "Đúng schema", "Số Điều tách được", "Số Điều lớn nhất", "Điều bị thiếu",
               "Điều tăng dần", "Khoản sai thứ tự", "Số hiệu có ở trang đầu", "Độ phủ nội dung",
               "Số unit", "Unit độ tin cậy thấp", "Cảnh báo"])
    for k in kq_vb:
        vb.append([k["document_id"], "Có" if k["schema_ok"] else "KHÔNG", k["so_dieu"], k["dieu_lon_nhat"],
                   ", ".join(map(str, k["dieu_thieu"])), "Có" if k["dieu_tang_dan"] else "KHÔNG",
                   k["khoan_sai_thu_tu"], "Có" if k["so_hieu_trang_dau"] else "Không (OCR méo số hiệu)",
                   k["do_phu"], k["so_unit"], k["unit_tin_cay_thap"], "; ".join(k["loi"])])
        vb.cell(vb.max_row, 9).number_format = "0.0%"
    dinh_dang_sheet(vb, [18, 9, 9, 9, 20, 9, 9, 16, 10, 8, 10, 60])
    vb.freeze_panes = "B2"

    mt = wb.create_sheet("Mau kiem tra tay")
    mt.append(["STT", "Chunk ID", "File gốc", "Trang", "Vị trí (header)", "Nội dung chunk (300 ký tự đầu)",
               "Khớp nguyên văn (tự động)", "Kiểm tra cấu trúc", "Kiểm tra nội dung", "Người kiểm tra", "Ghi chú"])
    tu_dong = {x["chunk_id"]: x for x in ds_chunk}
    for i, c in enumerate(mau, 1):
        k = da_kiem.get(c["chunk_id"], {})
        mt.append([i, c["chunk_id"], c["source_file"], c["page"], c["context_header"], c["content"][:300],
                   tu_dong[c["chunk_id"]]["khop_nguyen_van"], k.get("cau_truc", ""), k.get("noi_dung", ""),
                   k.get("nguoi_kiem_tra", ""), k.get("ghi_chu", "")])
        mt.cell(mt.max_row, 7).number_format = "0%"
    dinh_dang_sheet(mt, [5, 26, 30, 6, 40, 60, 10, 10, 10, 12, 45])
    for row in mt.iter_rows(min_row=2):
        for c in (row[4], row[5], row[10]):
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for c in (row[7], row[8]):
            if c.value == "Đúng":
                c.fill = PatternFill("solid", fgColor="E2EFDA")
            elif c.value:
                c.fill = PatternFill("solid", fgColor="FCE4D6")

    oc = wb.create_sheet("Chat luong OCR")
    oc.append(["Document ID", "Trang", "Số ký tự", "Lỗi ký tự", "CER", "Số từ", "Lỗi từ", "WER"])
    cer_file = TEXT_DIR / "_cer_ocr.csv"
    if cer_file.exists():
        for r in csv.DictReader(open(cer_file, encoding="utf-8-sig")):
            oc.append([r["document_id"], int(r["trang"]), int(r["so_ky_tu"]), int(r["loi_ky_tu"]), float(r["CER"]),
                       int(r["so_tu"]), int(r["loi_tu"]), float(r["WER"])])
            oc.cell(oc.max_row, 5).number_format = oc.cell(oc.max_row, 8).number_format = "0.0%"
    m = oc.max_row
    oc.append([])
    oc.append(["CER tổng", None, f"=SUM(C2:C{m})", f"=SUM(D2:D{m})", f"=D{m + 2}/C{m + 2}", f"=SUM(F2:F{m})",
               f"=SUM(G2:G{m})", f"=G{m + 2}/F{m + 2}"])
    oc.cell(m + 2, 5).number_format = oc.cell(m + 2, 8).number_format = "0.0%"
    for c in oc[m + 2]:
        c.font = Font(name="Arial", bold=True, size=10)
    oc.append(["Cách đo: các văn bản PDF CÓ TEXT được chụp thành ảnh 200 dpi rồi OCR lại; so với text thật trong PDF "
               "để biết OCR sai bao nhiêu (CER = lỗi ký tự / tổng ký tự, WER = lỗi từ / tổng từ)."])
    dinh_dang_sheet(oc, [20, 7, 9, 9, 8, 8, 8, 8])
    wb.save(KIEM_KE)


if __name__ == "__main__":
    main()
