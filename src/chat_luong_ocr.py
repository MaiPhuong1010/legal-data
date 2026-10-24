"""
ĐO CHẤT LƯỢNG OCR (metadata OCR)

Hai phép đo:
  1) Độ tin cậy OCR của từng văn bản scan: Tesseract cho mỗi từ một điểm tin cậy 0–100.
     Lấy mẫu tối đa 8 trang trải đều trong văn bản -> điểm trung bình, % từ có độ tin cậy thấp (< 60).
  2) Tỷ lệ lỗi ký tự (CER) / lỗi từ (WER) của OCR: lấy các văn bản PDF CÓ TEXT, chụp trang thành ảnh,
     OCR lại rồi so với text gốc trong PDF (text gốc = đáp án đúng). Cho biết OCR sai khoảng bao nhiêu %.

Đầu ra: output/text/_chat_luong_ocr.csv  và  output/text/_cer_ocr.csv
Cách chạy:  python src\\chat_luong_ocr.py     (cần Tesseract như trich_text.py; đã có kết quả thì bỏ qua)
"""
import csv
import io
import os
import re
import sys
import unicodedata
import warnings
from multiprocessing import Pool
from pathlib import Path

warnings.filterwarnings("ignore")
import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parent))
from trich_text import DPI, OUT_DIR, TESSDATA_DIR, TESSERACT_EXE, danh_sach_van_ban

SO_TRANG_MAU = 8
NGUONG_THAP = 60
OUT_CONF = OUT_DIR / "_chat_luong_ocr.csv"
OUT_CER = OUT_DIR / "_cer_ocr.csv"


def _tess():
    import pytesseract
    if TESSERACT_EXE:
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
    return pytesseract


def _anh(trang):
    from PIL import Image
    pix = trang.get_pixmap(dpi=DPI, colorspace=pymupdf.csGRAY)
    return Image.open(io.BytesIO(pix.tobytes("png")))


def _cfg():
    return f'--tessdata-dir "{TESSDATA_DIR}"' if TESSDATA_DIR else ""


def trang_mau(n, k=SO_TRANG_MAU):
    if n <= k:
        return list(range(n))
    return sorted({round(i * (n - 1) / (k - 1)) for i in range(k)})


def do_tin_cay(vb):
    pt = _tess()
    diem, thap = [], 0
    trang_da_do = 0
    tat_ca = [(f, i) for f in vb["files"] for i in range(pymupdf.open(f).page_count)]
    for idx in trang_mau(len(tat_ca)):
        f, i = tat_ca[idx]
        d = pt.image_to_data(_anh(pymupdf.open(f)[i]), lang="vie", config=_cfg(), output_type=pt.Output.DICT)
        for w, c in zip(d["text"], d["conf"]):
            c = float(c)
            if w.strip() and c >= 0:
                diem.append(c)
                thap += c < NGUONG_THAP
        trang_da_do += 1
    tb = round(sum(diem) / len(diem), 1) if diem else None
    ty_le = round(thap / len(diem), 4) if diem else None
    print(f"  {vb['doc_id']:22} tin cậy TB={tb}  từ <{NGUONG_THAP}: {ty_le}", flush=True)
    return [vb["doc_id"], trang_da_do, len(diem), tb, ty_le]


def chuan_hoa(s):
    s = unicodedata.normalize("NFC", s)
    s = re.sub(r"(Người ký|Ký bởi|Email|Cơ quan|Thời gian ký)\s*:.*", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def cer_trang(viec):
    from rapidfuzz.distance import Levenshtein
    doc_id, f, i = viec
    trang = pymupdf.open(f)[i]
    goc = chuan_hoa(trang.get_text())
    if len(goc) < 300:
        return None
    ocr = chuan_hoa(_tess().image_to_string(_anh(trang), lang="vie", config=_cfg()))
    loi_ky_tu = Levenshtein.distance(goc, ocr)
    tu_goc, tu_ocr = goc.split(), ocr.split()
    loi_tu = Levenshtein.distance(tu_goc, tu_ocr)
    return [doc_id, i + 1, len(goc), loi_ky_tu, round(loi_ky_tu / len(goc), 4), len(tu_goc), loi_tu,
            round(loi_tu / len(tu_goc), 4)]


if __name__ == "__main__":
    os.environ.setdefault("OMP_THREAD_LIMIT", "1")
    ds = danh_sach_van_ban()
    cpu = max(1, os.cpu_count() or 2)

    # chỉ đo các văn bản scan CHƯA có kết quả (thêm văn bản mới thì chạy lại sẽ đo bổ sung)
    cu = []
    if OUT_CONF.exists():
        with open(OUT_CONF, encoding="utf-8-sig") as f:
            cu = list(csv.reader(f))[1:]
    da_do = {r[0] for r in cu}
    scan = [v for v in ds if "scan" in (v["dinh_dang"] or "") and v["doc_id"] not in da_do]
    if scan:
        print(f"1) Đo độ tin cậy OCR cho {len(scan)} văn bản scan (mẫu {SO_TRANG_MAU} trang/văn bản)...")
        with Pool(cpu) as p:
            kq = p.map(do_tin_cay, scan, chunksize=1)
        with open(OUT_CONF, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["document_id", "so_trang_mau", "so_tu", "do_tin_cay_tb", "ty_le_tu_tin_cay_thap"])
            w.writerows(sorted(cu + [[str(x) for x in r] for r in kq]))
    else:
        print("1) Độ tin cậy OCR: đã đo đủ - bỏ qua")

    if not OUT_CER.exists():
        text = [v for v in ds if "scan" not in (v["dinh_dang"] or "") and str(v["files"][0]).lower().endswith(".pdf")]
        viec = []
        for v in text:
            n = pymupdf.open(v["files"][0]).page_count
            viec += [(v["doc_id"], v["files"][0], i) for i in trang_mau(n, 6)]
        print(f"2) Đo CER: OCR lại {len(viec)} trang của {len(text)} văn bản có text rồi so với text gốc...")
        with Pool(cpu) as p:
            kq = [r for r in p.map(cer_trang, viec, chunksize=1) if r]
        with open(OUT_CER, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["document_id", "trang", "so_ky_tu", "loi_ky_tu", "CER", "so_tu", "loi_tu", "WER"])
            w.writerows(kq)
        tong = sum(r[2] for r in kq)
        print(f"   CER chung = {sum(r[3] for r in kq) / tong:.2%}, WER chung = {sum(r[6] for r in kq) / sum(r[5] for r in kq):.2%}")
    else:
        print("2) Đã có", OUT_CER.name, "- bỏ qua")
    print("=> Xong")
