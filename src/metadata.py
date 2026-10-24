"""
Metadata văn bản (dùng cho tach_van_ban.py):
  - Đọc giá trị đã đối chiếu thủ công với văn bản gốc (raw/bo_sung_metadata.csv)
  - Đọc quan hệ văn bản đã xác minh (raw/quan_he_danh_muc.csv) và gộp với quan hệ trích tự động
  - Suy ra amended_by / replaced_by / guided_by ... (metadata bắt buộc của WP4) và hiệu lực riêng từng Điều
  - Đọc kết quả đo chất lượng OCR (chat_luong_ocr.py)
"""
import csv
import re
import unicodedata
from datetime import date, timedelta
from pathlib import Path

from regex_luat import RE_HL, RE_NGAY, _ngay, tim_so_hieu

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "raw"
TEXT_DIR = ROOT / "output" / "text"

CHIEU_NGUOC = {"AMENDS": "amended_by", "REPLACES": "replaced_by", "GUIDES": "guided_by",
               "PARTIALLY_REPEALS": "partially_repealed_by"}
CHIEU_XUOI = {"AMENDS": "amends", "REPLACES": "replaces", "GUIDES": "guides",
              "PARTIALLY_REPEALS": "partially_repeals"}


def _doc_csv(f):
    if not f.exists():
        return []
    with open(f, encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


def doc_bo_sung():
    """{(document_id, truong): (gia_tri hoặc None, nguồn)} – giá trị đọc trực tiếp từ văn bản gốc."""
    return {(r["document_id"], r["truong"]): (r["gia_tri"] or None, r["nguon_doi_chieu"])
            for r in _doc_csv(RAW / "bo_sung_metadata.csv")}


def doc_trich_yeu():
    """{document_id: dòng} – tên (trích yếu) và ngày ban hành đọc trực tiếp từ trang 1 bản gốc."""
    return {r["document_id"]: r for r in _doc_csv(RAW / "trich_yeu_doi_chieu.csv")}


def doc_chat_luong_ocr():
    tin_cay = {r["document_id"]: r for r in _doc_csv(TEXT_DIR / "_chat_luong_ocr.csv")}
    cach_doc = {r["document_id"]: r for r in _doc_csv(TEXT_DIR / "_bao_cao_b7.csv")}
    return tin_cay, cach_doc


def hieu_luc_rieng(units, hieu_luc_chung):
    """Hiệu lực riêng của từng Điều: trong Điều "Hiệu lực thi hành", câu dạng
    "Quy định tại Điều 13 ... của Luật này có hiệu lực thi hành từ ngày 01 tháng 01 năm 2026"
    -> {13: ("2026-01-01", câu bằng chứng)}."""
    ket_qua = {}
    for u in units:
        if u["unit_type"] != "article" or not re.search(r"(?i)hiệu\s+l|kho\S{1,3}n\s+thi\s+h", u.get("title") or ""):
            continue
        phan = [x["content"] for x in units if x["unit_id"] == u["unit_id"] or x["unit_id"].startswith(u["unit_id"] + "_")]
        noi_dung = re.sub(r"\s+", " ", unicodedata.normalize("NFC", " ".join(phan)))
        for cau in re.split(r"(?<=[\.;])\s+", noi_dung):
            m = re.search(RE_HL + RE_NGAY, cau, re.I)
            if not m:
                continue
            ngay = _ngay(m)
            truoc = cau[:m.start()]
            if not ngay or ngay == hieu_luc_chung or tim_so_hieu(truoc):
                continue      # câu nói về văn bản khác, hoặc trùng hiệu lực chung
            for nhom in re.findall(r"(?:Điều|điều)\s+((?:\d{1,3}\s*(?:,|và)\s*)*\d{1,3})", truoc):
                for so in re.findall(r"\d{1,3}", nhom):
                    ket_qua[int(so)] = (ngay, cau[:250])
    return ket_qua


def gop_quan_he(docs, tu_dong):
    """docs: {document_no: doc_json}; tu_dong: {document_no: [ {relation, target, evidence} ]}.
    Trả về list quan hệ cuối cùng (đã xác minh ưu tiên hơn tự động) và gắn trường amends/amended_by... vào docs."""
    tat_ca = [r for r in _doc_csv(RAW / "quan_he_danh_muc.csv") if r.get("target")]
    bo_qua = {(r["source"], r["target"]) for r in tat_ca if r["relation"] == "BO_QUA"}   # quan hệ tự động đã kiểm tra là SAI
    xac_minh = [r for r in tat_ca if r["relation"] != "BO_QUA"]
    cap_da_xm = {(r["source"], r["target"]) for r in xac_minh}
    ket_qua = []
    for r in xac_minh:
        ket_qua.append(dict(source=r["source"], relation=r["relation"], target=r["target"],
                            method="xác minh (danh mục / văn bản gốc)", verified=True, evidence=r["ghi_chu"]))
    for nguon, ds in tu_dong.items():
        for r in ds:
            if r["relation"] == "REFERENCES":
                continue
            if (nguon, r["target"]) in bo_qua:
                continue
            if (nguon, r["target"]) in cap_da_xm:
                # tự động trùng với bản đã xác minh -> chỉ ghi nhận là khớp
                for k in ket_qua:
                    if k["source"] == nguon and k["target"] == r["target"] and k["relation"] == r["relation"]:
                        k["method"] += " + tự động khớp"
                continue
            ket_qua.append(dict(source=nguon, relation=r["relation"], target=r["target"],
                                method="tự động (regex) – cần kiểm tra", verified=False, evidence=r["evidence"]))
    # REFERENCES: chỉ giữ văn bản có trong bộ dữ liệu (để dựng đồ thị WP5)
    for nguon, ds in tu_dong.items():
        for r in ds:
            if r["relation"] == "REFERENCES" and r["target"] in docs:
                ket_qua.append(dict(source=nguon, relation="REFERENCES", target=r["target"],
                                    method="tự động (regex)", verified=False, evidence=r["evidence"]))

    for d in docs.values():
        for k in list(CHIEU_XUOI.values()) + list(CHIEU_NGUOC.values()):
            d[k] = []
    for r in ket_qua:
        if r["relation"] not in CHIEU_XUOI:
            continue
        if r["source"] in docs:
            docs[r["source"]][CHIEU_XUOI[r["relation"]]].append(r["target"])
        if r["target"] in docs:
            docs[r["target"]][CHIEU_NGUOC[r["relation"]]].append(r["source"])
    # văn bản bị thay thế toàn bộ bởi một văn bản trong bộ dữ liệu -> effective_to, status
    for d in docs.values():
        for moi in d["replaced_by"]:
            hl = docs[moi].get("effective_from")
            if hl:
                d["effective_to"] = (date.fromisoformat(hl) - timedelta(days=1)).isoformat()
                if hl <= date.today().isoformat():
                    d["status"] = "replaced"
    return ket_qua
