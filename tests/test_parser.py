"""
TEST SUITE kiểm thử bộ tách cấu trúc (theo yêu cầu WP1: kiểm thử theo từng loại văn bản Luật / Nghị định / Thông tư)

Chạy (đứng ở thư mục legal-data):
    pip install pytest
    python -m pytest tests -v
"""
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
from regex_luat import tach_cau_truc, tim_ngay, trich_quan_he  # noqa: E402

STRUCT = ROOT / "output" / "structure"


def loai(units, t):
    return [u for u in units if u["unit_type"] == t]


# ---------------- 1. Kiểm thử đơn vị trên đoạn văn mẫu ----------------
LUAT = """<<<TRANG 1>>>
QUỐC HỘI
Luật số: 48/2024/QH15
LUẬT
THUẾ GIÁ TRỊ GIA TĂNG
Căn cứ Hiến pháp nước Cộng hòa xã hội chủ nghĩa Việt Nam;
Chương I
NHỮNG QUY ĐỊNH CHUNG
Điều 1. Phạm vi điều chỉnh
Luật này quy định về đối tượng chịu thuế.
Điều 2. Người nộp thuế
1. Tổ chức, hộ, cá nhân sản xuất, kinh doanh hàng hóa.
2. Tổ chức, cá nhân nhập khẩu hàng hóa chịu thuế gồm:
a) Hàng hóa nhập khẩu;
b) Dịch vụ nhập khẩu;
đ) Điểm sai thứ tự không được nhận
<<<TRANG 2>>>
Chương II
CĂN CỨ TÍNH THUẾ
Điều 3. Hiệu lực thi hành
1. Luật này có hiệu lực thi hành từ ngày 01 tháng 7 năm 2025.
2. Luật Thuế giá trị gia tăng số 13/2008/QH12 hết hiệu lực kể từ ngày Luật này có hiệu lực thi hành.
Luật này được Quốc hội nước Cộng hòa xã hội chủ nghĩa Việt Nam khóa XV, kỳ họp thứ 8 thông qua ngày 26 tháng 11 năm 2024.
"""


def test_luat_tach_dung_dieu_khoan_diem():
    units, info = tach_cau_truc(LUAT, "L48_2024_QH15")
    assert [u["article"] for u in loai(units, "article")] == [1, 2, 3]
    assert [u["unit_id"] for u in loai(units, "clause")][:2] == ["L48_2024_QH15_D2_K1", "L48_2024_QH15_D2_K2"]
    assert [u["point"] for u in loai(units, "point")] == ["a", "b"]          # "đ)" sai thứ tự -> không nhận
    assert loai(units, "article")[2]["chapter"] == "II"
    assert loai(units, "point")[0]["parent_id"] == "L48_2024_QH15_D2_K2"


def test_luat_ngay_ban_hanh_va_hieu_luc():
    assert tim_ngay(LUAT, "48/2024/QH15") == ("2024-11-26", "2025-07-01")


def test_luat_quan_he_thay_the():
    units, _ = tach_cau_truc(LUAT, "L48_2024_QH15")
    _, qh = trich_quan_he("48/2024/QH15", units)
    assert {"relation": "REPLACES", "target": "13/2008/QH12"}.items() <= qh[0].items()


NGHI_DINH_SUA_DOI = """<<<TRANG 1>>>
CHÍNH PHỦ
Số: 310/2025/NĐ-CP Hà Nội, ngày 02 tháng 12 năm 2025
NGHỊ ĐỊNH
Sửa đổi, bổ sung một số điều của Nghị định số 125/2020/NĐ-CP
Căn cứ Luật Tổ chức Chính phủ;
Điều 1. Sửa đổi, bổ sung một số điều của Nghị định số 125/2020/NĐ-CP
1. Sửa đổi, bổ sung khoản 2 Điều 6 như sau:
“2. Vi phạm hành chính có quy mô lớn trong các trường hợp sau:
a) Thực hiện một hành vi vi phạm;
Điều 7. Điều này nằm trong đoạn trích, không phải Điều của Nghị định
b) Thực hiện hành vi trốn thuế.”.
2. Bãi bỏ khoản 3 Điều 7.
Điều 2. Hiệu lực thi hành
Nghị định này có hiệu lực thi hành từ ngày 16 tháng 01 năm 2026.
"""


def test_nghi_dinh_sua_doi_khong_nham_dieu_trong_trich_dan():
    units, _ = tach_cau_truc(NGHI_DINH_SUA_DOI, "ND310_2025")
    assert [u["article"] for u in loai(units, "article")] == [1, 2]
    k1 = next(u for u in units if u["unit_id"] == "ND310_2025_D1_K1")
    assert "Điều 7. Điều này nằm trong đoạn trích" in k1["content"]      # giữ nguyên trong Khoản sửa đổi
    assert "ND310_2025_D1_K2" in [u["unit_id"] for u in units]


def test_nghi_dinh_quan_he_sua_doi_va_ngay():
    units, _ = tach_cau_truc(NGHI_DINH_SUA_DOI, "ND310_2025")
    _, qh = trich_quan_he("310/2025/NĐ-CP", units)
    assert ("AMENDS", "125/2020/NĐ-CP") in [(r["relation"], r["target"]) for r in qh]
    assert tim_ngay(NGHI_DINH_SUA_DOI, "310/2025/NĐ-CP") == ("2025-12-02", "2026-01-16")


THONG_TU_OCR = """<<<TRANG 1>>>
BỘ TÀI CHÍNH
Số: 90/2026/TT-BTC Hà Nội, ngày 30 tháng 6 năm 2026
THÔNG TƯ
Quy định về đăng ký thuế
Căn cứ Luật Quản lý thuế số 108/2025/QH15;
Chương J
QUY ĐỊNH CHUNG
Điều 1. Phạm vi điều chỉnh
Thông tư này quy định về đăng ký thuế.
Đ1eu 2. Đối tượng áp dụng
l. Người nộp thuế theo quy định.
2. Cơ quan thuế.
Điều 3. Điều 14, Điều 15 Nghị định này là dòng dẫn chiếu bị xuống dòng
s _ Chương H i
QUẢN LÝ ĐĂNG KÝ THUẾ
Điều 3. Hồ sơ đăng ký thuế
Nội dung.
Điều 4. Hiệu lực thi hành
1. Thông tư mỵ có hiệu lực thị hành kể từ ngày 01 tháng 7 năm 2026, thay thế Thông tư số 86/2024/TT-BTC.
Nơi nhận:
- Như trên;
PHỤ LỤC I
MẪU TỜ KHAI
STT Chỉ tiêu
"""


def test_thong_tu_chiu_loi_ocr():
    units, info = tach_cau_truc(THONG_TU_OCR, "TT90_2026_BTC")
    dieu = loai(units, "article")
    assert [u["article"] for u in dieu] == [1, 2, 3, 4]                     # "Đ1eu 2." vẫn nhận; dòng dẫn chiếu bị loại
    assert dieu[0]["chapter"] == "I" and dieu[2]["chapter"] == "II"           # "Chương J", "Chương H i" -> I, II
    assert "TT90_2026_BTC_D2_K1" in [u["unit_id"] for u in units]             # "l." -> Khoản 1
    pl = loai(units, "appendix")
    assert len(pl) == 1 and pl[0]["has_table"]
    assert tim_ngay(THONG_TU_OCR, "90/2026/TT-BTC")[1] == "2026-07-01"


# ---------------- 2. Kiểm thử trên dữ liệu thật đã tách (output/structure) ----------------
can_du_lieu = pytest.mark.skipif(not STRUCT.exists(), reason="chưa chạy tach_van_ban.py")


def doc(doc_id):
    return json.loads((STRUCT / f"{doc_id}.json").read_text(encoding="utf-8"))


@can_du_lieu
@pytest.mark.parametrize("doc_id,so_dieu", [("L48_2024_QH15", 18), ("L41_2024_QH15", 141), ("L50_2024_QH15", 37),
                                            ("ND254_2026", 45), ("ND252_2026", 76), ("TT89_2026_BTC", 101),
                                            ("TT90_2026_BTC", 38), ("TT99_2025_BTC", 31)])
def test_so_dieu_dung(doc_id, so_dieu):
    assert len([u for u in doc(doc_id)["legal_units"] if u["unit_type"] == "article"]) == so_dieu


@can_du_lieu
def test_nd254_dieu9_khoan2_dung_noi_dung():
    u = next(u for u in doc("ND254_2026")["legal_units"] if u["unit_id"] == "ND254_2026_D9_K2")
    assert u["content"].startswith("Thời điểm lập hóa đơn đối với cung cấp dịch vụ")
    assert u["parent_id"] == "ND254_2026_D9" and u["chapter"] == "II"


@can_du_lieu
@pytest.mark.parametrize("doc_id,ngay", [("L108_2025_QH15", "2026-07-01"), ("L48_2024_QH15", "2025-07-01"),
                                         ("ND254_2026", "2026-07-01"), ("L45_2019_QH14", "2021-01-01"),
                                         ("L59_2020_QH14", "2021-01-01")])
def test_ngay_hieu_luc(doc_id, ngay):
    assert doc(doc_id)["effective_from"] == ngay


@can_du_lieu
def test_hieu_luc_rieng_tung_dieu_l108():
    d13 = next(u for u in doc("L108_2025_QH15")["legal_units"] if u["unit_id"] == "L108_2025_QH15_D13")
    assert d13["effective_from"] == "2026-01-01"                             # Điều 13 có hiệu lực sớm hơn toàn Luật


@can_du_lieu
def test_quan_he_nguoc():
    assert "310/2025/NĐ-CP" in doc("ND125_2020")["amended_by"]
    assert "254/2026/NĐ-CP" in doc("L108_2025_QH15")["guided_by"]
