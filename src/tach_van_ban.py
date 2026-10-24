"""
TÁCH CẤU TRÚC VĂN BẢN + ĐẾM TOKEN

Đầu vào : output/text/*.txt (kết quả của trich_text.py), output/Kiem_ke_van_ban.xlsx (kết quả của kiem_ke.py)
Đầu ra  : output/structure/<document_id>.json  – văn bản dạng có cấu trúc theo JSON schema WP1
          Sheet "Thong ke token" và "Phan bo token" trong output/Kiem_ke_van_ban.xlsx

Cách chạy (đứng ở thư mục legal-data):
    python src\\tach_van_ban.py
"""
import json
import re
import statistics
import unicodedata
import sys
from collections import Counter
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, str(Path(__file__).resolve().parent))
from regex_luat import co_quan, dem_token, doc_tokenizer, tach_cau_truc, tim_hieu_luc, tim_ngay, trich_quan_he
from metadata import doc_bo_sung, doc_chat_luong_ocr, doc_trich_yeu, gop_quan_he, hieu_luc_rieng

ROOT = Path(__file__).resolve().parent.parent
TEXT_DIR = ROOT / "output" / "text"
STRUCT_DIR = ROOT / "output" / "structure"
KIEM_KE = ROOT / "output" / "Kiem_ke_van_ban.xlsx"
NGUONG = 512                       # ngưỡng token để đánh giá (giới hạn khuyến nghị cho 1 chunk)
HOM_NAY = date.today().isoformat()


def phan_vi(ds, p):
    if not ds:
        return 0
    ds = sorted(ds)
    k = (len(ds) - 1) * p / 100
    f = int(k)
    return round(ds[f] + (ds[min(f + 1, len(ds) - 1)] - ds[f]) * (k - f))


def toan_van(unit, con, co_nhan=False):
    """Nguyên văn đầy đủ của 1 unit = nội dung của nó + của tất cả unit con (kèm số Khoản "1." / Điểm "a)")."""
    nhan = {"clause": f"{unit.get('clause')}. ", "point": f"{unit.get('point')}) "}.get(unit["unit_type"], "") if co_nhan else ""
    phan = [unit.get("title") if unit["unit_type"] == "article" else None, nhan + unit["content"]]
    for c in con.get(unit["unit_id"], []):
        phan.append(toan_van(c, con, True))
    return "\n".join(p for p in phan if p and p.strip())


THU_TU = ["unit_id", "unit_type", "parent_id", "part", "chapter", "chapter_title", "section", "section_title",
          "article", "clause", "point", "title", "content", "page", "has_table", "token_count", "token_count_full",
          "effective_from", "effective_to", "status", "effective_note", "is_confident"]


def sap_xep(u):
    return {k: u.get(k) for k in THU_TU if k in u}


def doc_danh_muc():
    ws = openpyxl.load_workbook(KIEM_KE)["Danh muc van ban"]
    info = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        stt, nhom, so_hieu, loai, ten, ten_file, thu_muc, dinh_dang, so_trang, _, tinh_trang, link, note = r[:13]
        if tinh_trang != "Có file":
            continue
        doc_id = Path(ten_file.split("+")[0].strip()).stem
        info.setdefault(doc_id, dict(nhom=[], so_hieu=so_hieu, loai=loai, ten=ten, link=link,
                                     dinh_dang=dinh_dang, so_trang=so_trang,
                                     files=[f"{thu_muc}/{x.strip()}".replace("\\", "/") for x in ten_file.split("+")]))
        info[doc_id]["nhom"].append(nhom)
    return info


def main():
    STRUCT_DIR.mkdir(parents=True, exist_ok=True)
    print("Tokenizer:", doc_tokenizer()[0])
    danh_muc = doc_danh_muc()
    bo_sung = doc_bo_sung()
    trich_yeu = doc_trich_yeu()
    tin_cay_ocr, cach_doc = doc_chat_luong_ocr()
    bang, tat_ca_dieu, tat_ca_khoan = [], [], []
    docs, tu_dong, meta_rows = {}, {}, []

    # ---------- Lượt 1: tách cấu trúc từng văn bản ----------
    for doc_id, dm in danh_muc.items():
        f = TEXT_DIR / f"{doc_id}.txt"
        if not f.exists():
            print(f"  !! chưa có text của {doc_id} (chạy trich_text.py trước)")
            continue
        text = f.read_text(encoding="utf-8")
        units, info = tach_cau_truc(text, doc_id)

        # metadata: tự động bằng regex, sau đó ghi đè bằng giá trị đã đối chiếu với văn bản gốc (nếu có)
        ban_hanh, hieu_luc = tim_ngay(text, dm["so_hieu"])
        nguon_bh = "tự động (regex)" if ban_hanh else "không tìm thấy"
        nguon_hl = "tự động (regex)" if hieu_luc else "không tìm thấy"
        ty = trich_yeu.get(doc_id, {})
        if ty.get("ngay_ban_hanh_ban_goc"):
            if ban_hanh and ban_hanh != ty["ngay_ban_hanh_ban_goc"]:
                nguon_bh = f"đối chiếu trang 1 bản gốc (regex đọc sai: {ban_hanh})"
            else:
                nguon_bh = "đối chiếu trang 1 bản gốc"
            ban_hanh = ty["ngay_ban_hanh_ban_goc"]
        if (doc_id, "issued_date") in bo_sung:
            ban_hanh, ghi = bo_sung[(doc_id, "issued_date")]
            nguon_bh = "đối chiếu bản gốc: " + ghi
        if (doc_id, "effective_from") in bo_sung:
            hieu_luc, ghi = bo_sung[(doc_id, "effective_from")]
            nguon_hl = "đối chiếu bản gốc: " + ghi
        _, bang_chung_hl = tim_hieu_luc(re.sub(r"\s+", " ", unicodedata.normalize("NFC", text)), ban_hanh)
        trang_thai = "unknown" if not hieu_luc else ("active" if hieu_luc <= HOM_NAY else "not_yet_effective")
        ten_day_du, quan_he = trich_quan_he(dm["so_hieu"], units)
        tu_dong[dm["so_hieu"]] = quan_he

        con = {}
        for u in units:
            if u.get("parent_id"):
                con.setdefault(u["parent_id"], []).append(u)
        rieng = hieu_luc_rieng(units, hieu_luc)          # hiệu lực riêng của một số Điều (unit-level)
        for u in units:
            u["token_count"] = dem_token(u["content"])
            u["effective_from"] = hieu_luc
            u["effective_to"] = None
            u["status"] = trang_thai
            if u.get("article") in rieng:
                ngay, cau = rieng[u["article"]]
                u["effective_from"] = ngay
                u["effective_note"] = "Hiệu lực riêng theo Điều hiệu lực của văn bản: " + cau
                u["status"] = "active" if ngay <= HOM_NAY else "not_yet_effective"

        dieu = [u for u in units if u["unit_type"] == "article"]
        khoan = [u for u in units if u["unit_type"] == "clause"]
        tok_dieu = [dem_token(toan_van(u, con)) for u in dieu]
        tok_khoan = [dem_token(toan_van(u, con)) for u in khoan]
        for u, t in zip(dieu, tok_dieu):
            u["token_count_full"] = t
        for u, t in zip(khoan, tok_khoan):
            u["token_count_full"] = t
        tat_ca_dieu += tok_dieu
        tat_ca_khoan += tok_khoan

        dem = Counter(u["unit_type"] for u in units)
        so_chuong = len({u.get("chapter") for u in dieu if u.get("chapter")})
        so_muc = len({(u.get("chapter"), u.get("section")) for u in dieu if u.get("section")})
        tong_token = sum(u["token_count"] + dem_token(u.get("title") or "") for u in units)
        khong_chac = sum(1 for u in units if not u["is_confident"])
        la_ocr = "scan" in (dm["dinh_dang"] or "")
        tc = tin_cay_ocr.get(doc_id, {})

        docs[dm["so_hieu"]] = {
            "document_id": doc_id,
            "document_no": dm["so_hieu"],
            "document_type": dm["loai"],
            "title": ty.get("trich_yeu_ban_goc") or dm["ten"],
            "title_in_document": ten_day_du,
            "issuer": co_quan(dm["so_hieu"]),
            "issued_date": ban_hanh,
            "effective_from": hieu_luc,
            "effective_to": None,
            "status": trang_thai,
            "groups": dm["nhom"],
            "source_url": dm["link"],
            "source_file": dm["files"],
            "source_format": dm["dinh_dang"],
            "metadata_source": {"issued_date": nguon_bh, "effective_from": nguon_hl,
                                "effective_evidence": bang_chung_hl},
            "extraction": {
                "method": "OCR" if la_ocr else "PDF text layer",
                "ocr_engine": "Tesseract 5 (LSTM), ngôn ngữ vie" if la_ocr else None,
                "ocr_dpi": 200 if la_ocr else None,
                "ocr_mean_confidence": float(tc["do_tin_cay_tb"]) if tc.get("do_tin_cay_tb") else None,
                "ocr_low_confidence_word_ratio": float(tc["ty_le_tu_tin_cay_thap"]) if tc.get("ty_le_tu_tin_cay_thap") else None,
                "ocr_sampled_pages": int(tc["so_trang_mau"]) if tc.get("so_trang_mau") else None,
                "pages": dm["so_trang"],
                "tokenizer": doc_tokenizer()[0],
            },
            "parse_warnings": info["canh_bao"],
            "legal_units": [sap_xep(u) for u in units],
        }
        bang.append([
            doc_id, dm["so_hieu"], dm["loai"], "OCR" if la_ocr else "Text",
            dm["so_trang"], so_chuong, so_muc, dem["article"], dem["clause"], dem["point"], dem["appendix"],
            "Có" if any(u["has_table"] for u in units) else "", tong_token,
            round(statistics.mean(tok_dieu)) if tok_dieu else None, max(tok_dieu, default=None),
            phan_vi(tok_dieu, 95) if tok_dieu else None,
            round(sum(t > NGUONG for t in tok_dieu) / len(tok_dieu), 3) if tok_dieu else None,
            max(tok_khoan, default=None),
            round(sum(t > NGUONG for t in tok_khoan) / len(tok_khoan), 3) if tok_khoan else None,
            khong_chac, ban_hanh, hieu_luc, "; ".join(info["canh_bao"][:3]),
        ])
        print(f"  {doc_id:22} Điều={dem['article']:4} Khoản={dem['clause']:4} Điểm={dem['point']:4} "
              f"PL={dem['appendix']:3} token={tong_token:7}  hiệu lực={hieu_luc}")

    # ---------- Lượt 2: quan hệ giữa các văn bản (WP4/WP5) ----------
    quan_he = gop_quan_he(docs, tu_dong)
    for d in docs.values():
        if d["status"] == "replaced":
            for u in d["legal_units"]:
                u["status"], u["effective_to"] = "replaced", d["effective_to"]
        thu_tu = ["document_id", "document_no", "document_type", "title", "title_in_document", "issuer",
                  "issued_date", "effective_from", "effective_to", "status", "amends", "amended_by", "replaces",
                  "replaced_by", "partially_repeals", "partially_repealed_by", "guides", "guided_by", "groups",
                  "source_url", "source_file", "source_format", "metadata_source", "extraction", "parse_warnings",
                  "legal_units"]
        d2 = {k: d.get(k) for k in thu_tu}
        (STRUCT_DIR / f"{d['document_id']}.json").write_text(json.dumps(d2, ensure_ascii=False, indent=2), encoding="utf-8")
        tc, cd = tin_cay_ocr.get(d["document_id"], {}), cach_doc.get(d["document_id"], {})
        meta_rows.append([d["document_id"], d["document_no"], d["document_type"], d["title_in_document"] or d["title"],
                          d["issuer"], d["issued_date"], d["metadata_source"]["issued_date"],
                          d["effective_from"], d["metadata_source"]["effective_from"],
                          (d["metadata_source"]["effective_evidence"] or "")[:200], d["status"],
                          ", ".join(d["amends"]), ", ".join(d["replaces"]), ", ".join(d["guides"]),
                          ", ".join(d["amended_by"]), ", ".join(d["replaced_by"]), ", ".join(d["guided_by"]),
                          d["source_format"], d["extraction"]["method"], d["extraction"]["ocr_engine"],
                          d["extraction"]["ocr_dpi"], d["extraction"]["ocr_mean_confidence"],
                          d["extraction"]["ocr_low_confidence_word_ratio"], cd.get("thoi_gian_giay")])
    (ROOT / "output" / "quan_he_van_ban.json").write_text(json.dumps(quan_he, ensure_ascii=False, indent=2), encoding="utf-8")

    ghi_excel(bang, tat_ca_dieu, tat_ca_khoan)
    ghi_excel_metadata(meta_rows, quan_he)
    print(f"\n=> Đã lưu {len(docs)} file JSON vào {STRUCT_DIR}")
    print(f"=> Quan hệ văn bản: output/quan_he_van_ban.json ({len(quan_he)} quan hệ)")
    print(f"=> Đã thêm sheet 'Thong ke token', 'Phan bo token', 'Metadata van ban', 'Quan he van ban' vào {KIEM_KE}")


def ghi_excel_metadata(meta_rows, quan_he):
    wb = openpyxl.load_workbook(KIEM_KE)
    for ten in ("Metadata van ban", "Quan he van ban"):
        if ten in wb.sheetnames:
            del wb[ten]
    ws = wb.create_sheet("Metadata van ban")
    ws.append(["Document ID", "Số hiệu", "Loại VB", "Tên văn bản (trích từ văn bản)", "Cơ quan ban hành",
               "Ngày ban hành", "Nguồn ngày ban hành", "Ngày hiệu lực", "Nguồn ngày hiệu lực",
               "Câu làm bằng chứng hiệu lực", "Trạng thái", "Sửa đổi VB", "Thay thế VB", "Hướng dẫn VB",
               "Bị sửa đổi bởi", "Bị thay thế bởi", "Được hướng dẫn bởi", "Định dạng gốc", "Cách lấy text",
               "OCR engine", "OCR DPI", "OCR độ tin cậy TB (0-100)", "OCR % từ tin cậy thấp (<60)", "Thời gian xử lý (giây)"])
    for r in meta_rows:
        ws.append(r)
    dinh_dang_sheet(ws, [18, 16, 9, 45, 14, 11, 30, 11, 30, 45, 12, 22, 22, 22, 22, 18, 22, 22, 13, 18, 7, 11, 11, 9])
    for row in ws.iter_rows(min_row=2):
        row[22].number_format = "0.0%"
        for c in (row[3], row[6], row[8], row[9]):
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    qh = wb.create_sheet("Quan he van ban")
    qh.append(["Văn bản nguồn", "Quan hệ", "Văn bản đích", "Cách xác định", "Đã xác minh", "Bằng chứng / ghi chú"])
    for r in sorted(quan_he, key=lambda r: (not r["verified"], r["relation"] == "REFERENCES", r["source"])):
        qh.append([r["source"], r["relation"], r["target"], r["method"], "Có" if r["verified"] else "Chưa",
                   (r["evidence"] or "")[:250]])
    dinh_dang_sheet(qh, [17, 18, 17, 34, 10, 90])
    qh.freeze_panes = "A2"
    qh.auto_filter.ref = qh.dimensions
    wb.save(KIEM_KE)


def dinh_dang_sheet(ws, do_rong, dong_tieu_de=1):
    vien = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
    for c in ws[dong_tieu_de]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=dong_tieu_de + 1):
        for c in row:
            if c.value is not None and not (c.font and c.font.bold):
                c.font = Font(name="Arial", size=10)
            c.border = vien if c.value is not None else Border()
    for i, w in enumerate(do_rong, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


def ghi_excel(bang, tok_dieu, tok_khoan):
    wb = openpyxl.load_workbook(KIEM_KE)
    for ten in ("Thong ke token", "Phan bo token"):
        if ten in wb.sheetnames:
            del wb[ten]

    # ----- Sheet 2: thống kê từng văn bản -----
    ws = wb.create_sheet("Thong ke token")
    cot = ["Document ID", "Số hiệu", "Loại VB", "Nguồn text", "Số trang", "Số Chương", "Số Mục", "Số Điều",
           "Số Khoản", "Số Điểm", "Số Phụ lục", "Có bảng", "Tổng token", "Token TB / Điều", "Token max / Điều",
           "P95 token / Điều", f"% Điều > {NGUONG}", "Token max / Khoản", f"% Khoản > {NGUONG}",
           "Số unit độ tin cậy thấp", "Ngày ban hành", "Ngày hiệu lực", "Cảnh báo khi tách"]
    ws.append(cot)
    for r in bang:
        ws.append(r)
    n = len(bang) + 1
    ws.append(["TỔNG CỘNG", None, None, None, f"=SUM(E2:E{n})", f"=SUM(F2:F{n})", f"=SUM(G2:G{n})",
               f"=SUM(H2:H{n})", f"=SUM(I2:I{n})", f"=SUM(J2:J{n})", f"=SUM(K2:K{n})", None, f"=SUM(M2:M{n})",
               round(statistics.mean(tok_dieu)) if tok_dieu else None, max(tok_dieu, default=None),
               phan_vi(tok_dieu, 95), round(sum(t > NGUONG for t in tok_dieu) / max(len(tok_dieu), 1), 3),
               max(tok_khoan, default=None), round(sum(t > NGUONG for t in tok_khoan) / max(len(tok_khoan), 1), 3),
               f"=SUM(T2:T{n})"])
    for c in ws[n + 1]:
        c.font = Font(name="Arial", bold=True, size=10)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    dinh_dang_sheet(ws, [20, 18, 10, 8, 8, 8, 7, 7, 8, 8, 8, 7, 10, 9, 9, 9, 9, 9, 9, 10, 11, 11, 40])
    for row in ws.iter_rows(min_row=2, min_col=17, max_col=19):
        for c in row:
            if c.column in (17, 19):
                c.number_format = "0.0%"
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:W{n}"
    ws.cell(n + 3, 1, f"Token đếm bằng tokenizer: {doc_tokenizer()[0]}. Token của Điều/Khoản = toàn văn kể cả Khoản/Điểm con.")
    ws.cell(n + 4, 1, "Dòng TỔNG CỘNG: các cột TB / max / P95 / % tính trên toàn bộ Điều, Khoản của mọi văn bản (do script tính).")
    ws.cell(n + 5, 1, "Ngày ban hành / hiệu lực được trích tự động bằng regex, cần kiểm tra lại với văn bản gốc.")

    # ----- Sheet 3: phân bố độ dài -----
    pb = wb.create_sheet("Phan bo token")
    pb.append(["Khoảng token", "Số Điều", "% Điều", "Số Khoản", "% Khoản"])
    moc = [0, 128, 256, 512, 1024, 2048, 4096, 10 ** 9]
    for a, b in zip(moc, moc[1:]):
        nhan = f"{a + 1}–{b}" if b < 10 ** 9 else f"> {a}"
        r = pb.max_row + 1
        pb.append([nhan, sum(a < t <= b for t in tok_dieu), f"=B{r}/SUM(B$2:B$8)",
                   sum(a < t <= b for t in tok_khoan), f"=D{r}/SUM(D$2:D$8)"])
    pb.append(["Tổng", "=SUM(B2:B8)", None, "=SUM(D2:D8)", None])
    pb.append([])
    pb.append(["Thống kê", "Điều", None, "Khoản"])
    for ten, f in [("Trung bình", lambda d: round(statistics.mean(d)) if d else 0),
                   ("Trung vị (P50)", lambda d: phan_vi(d, 50)), ("P90", lambda d: phan_vi(d, 90)),
                   ("P95", lambda d: phan_vi(d, 95)), ("Lớn nhất", lambda d: max(d, default=0))]:
        pb.append([ten, f(tok_dieu), None, f(tok_khoan)])
    dinh_dang_sheet(pb, [18, 10, 10, 10, 10])
    for row in pb.iter_rows(min_row=2, max_row=8):
        row[2].number_format = row[4].number_format = "0.0%"
    for c in pb[11]:
        c.font = Font(name="Arial", bold=True, size=10)
    wb.save(KIEM_KE)


if __name__ == "__main__":
    main()
