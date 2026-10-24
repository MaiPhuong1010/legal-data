"""
Thư viện dùng chung cho tach_van_ban.py và chunking.py:
  - doc_tokenizer(): bộ đếm token bge-m3
  - lam_sach(): bỏ dòng rác (chữ ký số, số trang, tiêu đề Công báo...)
  - tach_cau_truc(): tách văn bản thành Chương / Mục / Điều / Khoản / Điểm / Phụ lục bằng regex
"""
import os
import re
import unicodedata
from datetime import date

# ======================= 1. BỘ ĐẾM TOKEN =======================
_tok = None


def doc_tokenizer():
    """Ưu tiên tokenizer BAAI/bge-m3 (đã tải về máy ở bước kiểm tra môi trường).
    Nếu không có mạng/không có cache thì dùng file sentencepiece (biến môi trường BGE_SPM)."""
    global _tok
    if _tok:
        return _tok
    spm_path = os.environ.get("BGE_SPM")
    if spm_path:
        import sentencepiece as spm
        sp = spm.SentencePieceProcessor(model_file=spm_path)
        _tok = ("sentencepiece (cùng bộ từ vựng bge-m3)", lambda s: sp.encode(s), lambda ids: sp.decode(ids))
        return _tok
    import warnings
    warnings.filterwarnings("ignore")
    from transformers import AutoTokenizer
    tk = AutoTokenizer.from_pretrained("BAAI/bge-m3")
    _tok = ("BAAI/bge-m3", lambda s: tk.encode(s, add_special_tokens=False), lambda ids: tk.decode(ids))
    return _tok


def dem_token(s):
    return len(doc_tokenizer()[1](s)) if s else 0


def ma_hoa(s):
    return doc_tokenizer()[1](s)


def giai_ma(ids):
    return doc_tokenizer()[2](ids)


# ======================= 2. LÀM SẠCH TEXT =======================
DONG_RAC = [
    r"^(Người ký|Ký bởi|Email|Cơ quan|Thời gian ký)\s*:",
    r"^CÔNG BÁO\s*/",
    r"^\d{1,4}$",                                   # số trang
    r"^\d{1,4}\s+CÔNG BÁO",
    r"^CÔNG THÔNG TIN ĐIỆN TỬ CHÍNH PHỦ",
    r"^(ĐẾN|Giờ|Ngày)\s*:?\s*$",
    r"^[\W_]{1,4}$",                                # dòng chỉ có 1-4 ký hiệu lạ (vết mực, con dấu)
]
DONG_RAC = [re.compile(p) for p in DONG_RAC]
TRANG = re.compile(r"^<<<TRANG (\d+)>>>$")


def lam_sach(text):
    """Trả về list (số_trang, dòng) đã bỏ rác."""
    text = unicodedata.normalize("NFC", text)
    ket_qua, trang = [], 1
    for dong in text.splitlines():
        dong = dong.replace(" ", " ").strip()
        m = TRANG.match(dong)
        if m:
            trang = int(m.group(1))
            continue
        if not dong or any(p.search(dong) for p in DONG_RAC):
            continue
        # OCR hay dính vệt bẩn đầu dòng: "_ .1. Quy định..." -> "1. Quy định..."
        dong = re.sub(r"^[_\-–\.,·•'`~|\s]+(?=[\w“\"(])", "", dong)
        if not dong:
            continue
        ket_qua.append((trang, re.sub(r"\s+", " ", dong)))
    return ket_qua


# ======================= 3. REGEX NHẬN DIỆN CẤU TRÚC =======================
# OCR hay đọc sai: "Điều" -> "Điểu", "Diều"; "I" -> "J", "l", "1"
RE_PHAN = re.compile(r"^(?:Phần|PHẦN)\s+(?:thứ|THỨ)?\s*([A-Za-zÀ-ỹ0-9IVX]+)\s*$")
RE_CHUONG = re.compile(r"^(?:Chương|CHƯƠNG|Chuơng|Chươmg)\s+([IVXLCJl1]+|\d+)\s*[\.:]?\s*(.*)$")
RE_MUC = re.compile(r"^(?:Mục|MỤC)\s*(\d+|[IVX]+)\s*[\.:]?\s*(.*)$")
RE_CHUONG_MO = re.compile(r"^(?:Ch\w{1,4}ng|CH\w{1,4}NG)\s*(.{0,8})$")   # OCR méo: "Chương HI \\", "ChươngH i Ì"
RE_DIEU_MO = re.compile(r"^.{0,2}[ĐDÐ]\S{1,6}\s+\.?\s*(\d{1,3})\s*[\.:,\-]+\s*(.*)$")  # "Đ1tzu 3.", "Điề_l.l 67."
RE_DIEU = re.compile(r"^(?:Điều|Điểu|Diều|Ðiều|Điêu|Điếu)\s*(\d{1,3})\s*[\.:,]\s*(.*)$")
RE_KHOAN = re.compile(r"^(\d{1,2})\s*[\.,]\s+(\S.*)$")
RE_KHOAN_1_OCR = re.compile(r"^[lI|]\s*[\.,]\s+([A-ZĐÀ-Ỹ].*)$")
RE_DIEM = re.compile(r"^([a-zđ])\s*\)\s*(.*)$")
RE_PHU_LUC = re.compile(r"^(?:PHỤ LỤC|Phụ lục|PHU LUC)\b\s*([IVXLC\d]*)\b(.*)$")
RE_DAU_PHU_LUC = re.compile(r"^(PH\S{0,2}\s*L\S{0,2}C\b|Phụ\s+lục\s+[IVXLC\d]+|Mẫu\s*s\S{0,2}\s*:?\s*\d|DANH\s+M\S{0,2}C\b|BIỂU\s+MẪU)")
RE_NOI_NHAN = re.compile(r"^Nơi nhận\s*:?")
MO_NGOAC = ("“", '"', "„", "«")
DONG_NGOAC = re.compile(r"[”\"»]\s*[\.;,]?\s*$")

# thứ tự Điểm trong văn bản Việt Nam (không có f, j, w, z; có đ)
CHU_DIEM = list("abcdđeghiklmnopqrstuvxy")


def so_la_ma(s):
    s = s.upper().replace("J", "I").replace("L", "I").replace("1", "I")
    if s.isdigit():
        return s
    gt = {"I": 1, "V": 5, "X": 10, "C": 100}
    tong, truoc = 0, 0
    for c in reversed(s):
        v = gt.get(c, 0)
        tong = tong - v if v < truoc else tong + v
        truoc = max(truoc, v)
    roman = ""
    for v, k in [(100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]:
        while tong >= v:
            roman += k
            tong -= v
    return roman or s


def la_ma_sang_so(s):
    if not s or not re.fullmatch(r"[IVXLC]+", s):
        return int(s) if s and s.isdigit() else None
    gt = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100}
    tong, truoc = 0, 0
    for c in reversed(s):
        v = gt[c]
        tong = tong - v if v < truoc else tong + v
        truoc = max(truoc, v)
    return tong


def so_sang_la_ma(n):
    r = ""
    for v, k in [(100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]:
        while n >= v:
            r += k
            n -= v
    return r


def diem_ke_tiep(truoc):
    """Chữ cái điểm mong đợi tiếp theo."""
    if truoc is None:
        return "a"
    i = CHU_DIEM.index(truoc) if truoc in CHU_DIEM else -1
    return CHU_DIEM[i + 1] if 0 <= i < len(CHU_DIEM) - 1 else None


def ty_le_rac(s):
    """Tỷ lệ ký tự lạ (dấu hiệu OCR kém)."""
    if not s:
        return 0
    la = sum(1 for c in s if not (c.isalnum() or c.isspace() or c in ".,;:()-–/%“”\"'‘’?!…[]+*&=<>§°"))
    return la / len(s)


def noi_dong(cu, moi):
    if not cu:
        return moi
    return cu + ("\n" if cu.endswith((":", ";", ".")) else " ") + moi


# ======================= 4. TÁCH CẤU TRÚC =======================
class Unit(dict):
    pass


def tach_cau_truc(text, doc_id):
    """Trả về (legal_units, thong_tin) – legal_units là list dict theo thứ tự xuất hiện."""
    dong = lam_sach(text)
    units = []
    ctx = dict(phan=None, chuong=None, chuong_ten=None, muc=None, muc_ten=None)
    dieu = khoan = diem = None
    so_dieu_truoc = 0
    trong_ngoac = False
    vung = "mo_dau"            # mo_dau -> than -> cuoi (sau Nơi nhận) -> phu_luc
    canh_bao = []
    cho_ten_chuong = cho_ten_muc = False

    def tao(loai, **kw):
        u = Unit(unit_type=loai, content="", **kw)
        units.append(u)
        return u

    mo_dau = tao("preamble", unit_id=f"{doc_id}_MODAU", title="Phần mở đầu (căn cứ ban hành)",
                 article=None, clause=None, point=None, page=1, parent_id=None)
    hien_tai = mo_dau
    so_pl = 0
    so_chuong = [0]
    so_dong_cuoi = [0]

    def ten_hop_le(ten):
        """Tên Điều phải bắt đầu bằng chữ in hoa và không phải là câu dẫn chiếu ("Điều 14, Điều 15...")."""
        ten = ten.strip()
        return bool(ten) and ten[0].isupper() and not la_dan_chieu(ten)

    def la_dan_chieu(ten):
        return bool(re.match(r"^(Điều|khoản|Khoản|điểm|Điểm)\s*\d", ten.strip()))

    # danh sách các dòng "Điều N. Tên" rõ ràng – dùng để kiểm tra trước khi chấp nhận Điều bị nhảy số
    ung_vien = [(i, int(m.group(1))) for i, (_, s) in enumerate(dong)
                for m in [RE_DIEU.match(s) or RE_DIEU_MO.match(s)] if m and ten_hop_le(m.group(2))]

    def xac_nhan_nhay(i, n):
        sau = [k for j, k in ung_vien if j > i][:3]
        return any(k in (n + 1, n + 2) for k in sau[:2]) or not sau

    for i, (trang, s) in enumerate(dong):
        # ---------- vùng phụ lục / biểu mẫu ----------
        m = RE_PHU_LUC.match(s)
        if not m and vung in ("cuoi", "phu_luc") and len(s) < 90:
            m = RE_DAU_PHU_LUC.match(s)                 # "Mẫu số 01/CNKD", "PHU LỤC" (OCR), "DANH MỤC ..."
        if not m and vung == "cuoi":
            so_dong_cuoi[0] += 1
            if so_dong_cuoi[0] > 40:                    # khối nơi nhận + chữ ký đã qua, phần sau là biểu mẫu
                m = True
        if m and (vung in ("cuoi", "phu_luc") or (s.isupper() and len(s) < 80)) and not trong_ngoac:
            vung = "phu_luc"
            so_pl += 1
            ten = s
            hien_tai = tao("appendix", unit_id=f"{doc_id}_PL{so_pl}", title=ten, article=None, clause=None,
                           point=None, page=trang, parent_id=None, chapter=None, section=None)
            continue
        if vung == "phu_luc":
            hien_tai["content"] = noi_dong(hien_tai["content"], s)
            if s.startswith("STT") or s.startswith("TT "):
                hien_tai["has_table"] = True
            continue
        if RE_NOI_NHAN.match(s) and vung == "than":
            vung = "cuoi"
            continue
        if vung == "cuoi":
            continue

        # ---------- trong đoạn trích dẫn (văn bản sửa đổi "...như sau: “Điều 5. ...”") ----------
        if trong_ngoac:
            m_thoat = RE_DIEU.match(s)
            if m_thoat and int(m_thoat.group(1)) == so_dieu_truoc + 1 and not s.startswith(MO_NGOAC):
                trong_ngoac = False          # gặp Điều kế tiếp của chính văn bản -> đoạn trích đã kết thúc
            else:
                hien_tai["content"] = noi_dong(hien_tai["content"], s)
                if DONG_NGOAC.search(s):
                    trong_ngoac = False
                continue
        # chỉ coi là đoạn trích dẫn khi câu trước là "... như sau:" (văn bản sửa đổi, bổ sung)
        if s.startswith(MO_NGOAC) and hien_tai is not mo_dau and \
                re.search(r"(như sau|sau đây)\s*:\s*$", hien_tai["content"]) and len(s) > 3:
            hien_tai["content"] = noi_dong(hien_tai["content"], s)
            trong_ngoac = not DONG_NGOAC.search(s[1:])
            continue

        # ---------- tên Chương / Mục nằm ở dòng kế tiếp ----------
        if cho_ten_chuong and s.upper() == s and not RE_DIEU.match(s):
            ctx["chuong_ten"] = (ctx["chuong_ten"] + " " + s).strip() if ctx["chuong_ten"] else s
            continue
        if cho_ten_muc and s.upper() == s and not RE_DIEU.match(s):
            ctx["muc_ten"] = (ctx["muc_ten"] + " " + s).strip() if ctx["muc_ten"] else s
            continue
        cho_ten_chuong = cho_ten_muc = False

        m = RE_PHAN.match(s)
        if m and len(s) < 30:
            ctx["phan"] = m.group(1)
            continue
        # bỏ vệt bẩn OCR trước chữ "Chương" ở dòng ngắn: "s _ Chương H i", ": ChươngIV _", "Ñ Mục3"
        s_ch = re.sub(r"^.{0,5}?(?=Ch\w{1,4}ng|CH\w{1,4}NG|Mục|MỤC)", "", s) if len(s) <= 26 else s
        m = RE_CHUONG.match(s_ch)
        if m and re.sub(r"[^A-Za-zÀ-ỹ]", "", m.group(2) or "") and not m.group(2).isupper():
            m = None          # "Chương VII của Bộ luật..." là câu văn bị xuống dòng, không phải tiêu đề Chương
        m_mo = RE_CHUONG_MO.match(s_ch) if len(s_ch) <= 20 else None
        if m or m_mo:
            # Chương luôn tăng dần 1. Số La Mã đọc được rõ ràng thì tin; đọc không ra (OCR méo) thì suy ra = Chương trước + 1
            truoc = so_chuong[0]
            gt = la_ma_sang_so(so_la_ma(m.group(1))) if m else None
            if gt is not None and truoc < gt <= truoc + 2:
                pass
            elif gt is not None and gt <= truoc:
                gt = None                                  # trùng / lùi số -> dòng dẫn chiếu, bỏ qua
            elif m_mo and (m is None or not m.group(2)):
                gt = truoc + 1
                canh_bao.append(f"Chương OCR '{s}' -> suy ra Chương {so_sang_la_ma(gt)}")
            else:
                gt = None
            if gt:
                so_chuong[0] = gt
                ten_ch = re.sub(r"^[^A-Za-zÀ-ỹ]+", "", m.group(2)).strip() if m and m.group(2) else ""
                ctx.update(chuong=so_sang_la_ma(gt), chuong_ten=ten_ch or None, muc=None, muc_ten=None)
                cho_ten_chuong = True
                vung = "than"
                continue
        m = RE_MUC.match(s_ch)
        if m and vung == "than":
            ctx.update(muc=m.group(1), muc_ten=m.group(2) or None)
            cho_ten_muc = True
            continue

        # ---------- Điều ----------
        m = RE_DIEU.match(s)
        if not m:
            m = RE_DIEU_MO.match(s)                     # chỉ chấp nhận Điều "méo" nếu đúng số kế tiếp (±1)
            if m and not (so_dieu_truoc < int(m.group(1)) <= so_dieu_truoc + 2 and len(m.group(2)) >= 3):
                m = None
        if m and la_dan_chieu(m.group(2)):
            m = None
        if m:
            n = int(m.group(1))
            # phải tăng dần; nếu nhảy số (OCR làm mất Điều trước) thì Điều sau đó phải nối tiếp đúng
            if n == so_dieu_truoc + 1 or (so_dieu_truoc + 1 < n <= so_dieu_truoc + 4
                                          and ten_hop_le(m.group(2)) and xac_nhan_nhay(i, n)):
                if n != so_dieu_truoc + 1:
                    canh_bao.append(f"Nhảy từ Điều {so_dieu_truoc} sang Điều {n}")
                so_dieu_truoc = n
                vung = "than"
                ten = m.group(2).strip()
                # tên Điều bị xuống dòng: dòng sau bắt đầu chữ thường thì nối vào
                j = i + 1
                while j < len(dong) and ten and not ten.endswith((".", ":")) and dong[j][1][:1].islower() \
                        and not RE_DIEM.match(dong[j][1]) \
                        and not RE_KHOAN_1_OCR.match(dong[j][1]) and j - i <= 2:
                    ten += " " + dong[j][1]
                    j += 1
                dieu = tao("article", unit_id=f"{doc_id}_D{n}", title=f"Điều {n}. {ten}".strip(),
                           article=n, clause=None, point=None, page=trang, parent_id=None,
                           chapter=ctx["chuong"], chapter_title=ctx["chuong_ten"],
                           section=ctx["muc"], section_title=ctx["muc_ten"], part=ctx["phan"],
                           gap=n != so_dieu_truoc)
                dieu["_bo_qua"] = j - i - 1       # số dòng đã nối vào tên Điều
                khoan = diem = None
                hien_tai = dieu
                continue
        if dieu is not None and dieu.get("_bo_qua"):
            dieu["_bo_qua"] -= 1
            continue

        # ---------- Khoản ----------
        m = RE_KHOAN.match(s)
        if not m and dieu is not None and khoan is None:
            m = RE_KHOAN_1_OCR.match(s)                 # OCR đọc "1." thành "l." / "I."
            if m:
                m = re.match(r"^(1)\s*[\.,]\s+(\S.*)$", "1." + s[2:])
        if m and dieu is not None:
            n = int(m.group(1))
            ky_vong = (khoan["clause"] + 1) if khoan else 1
            if n == ky_vong or (n == ky_vong + 1 and len(m.group(2)) > 15 and m.group(2)[:1].isupper()):
                if n != ky_vong:
                    canh_bao.append(f"Điều {dieu['article']}: thiếu Khoản {ky_vong}")
                khoan = tao("clause", unit_id=f"{dieu['unit_id']}_K{n}", title=None, article=dieu["article"],
                            clause=n, point=None, page=trang, parent_id=dieu["unit_id"],
                            chapter=dieu["chapter"], section=dieu["section"])
                khoan["content"] = m.group(2)
                khoan["gap"] = n != ky_vong
                diem = None
                hien_tai = khoan
                continue

        # ---------- Điểm ----------
        m = RE_DIEM.match(s)
        if m and dieu is not None:
            c = m.group(1)
            ky_vong = diem_ke_tiep(diem["point"] if diem else None)
            if c == ky_vong or (ky_vong == "đ" and c == "d"):   # OCR hay đọc "đ)" thành "d)"
                c = ky_vong
                cha = khoan or dieu
                diem = tao("point", unit_id=f"{cha['unit_id']}_{'P' + c.replace('đ', 'dd')}", title=None,
                           article=dieu["article"], clause=khoan["clause"] if khoan else None, point=c,
                           page=trang, parent_id=cha["unit_id"], chapter=dieu["chapter"], section=dieu["section"])
                diem["content"] = m.group(2)
                hien_tai = diem
                continue

        # ---------- dòng nội dung thường ----------
        if s.startswith("STT"):
            hien_tai["has_table"] = True
        hien_tai["content"] = noi_dong(hien_tai["content"], s)

    # dọn: bỏ trường tạm, tính độ tin cậy
    for u in units:
        u.pop("_bo_qua", None)
        u["content"] = u["content"].strip()
        u["has_table"] = bool(u.get("has_table"))
        u["is_confident"] = ty_le_rac(u["content"]) < 0.03 and not u.pop("gap", False)
    units = [u for u in units if u["content"] or u["unit_type"] in ("article",)]
    return units, dict(canh_bao=canh_bao, so_dieu_cuoi=so_dieu_truoc)


# ======================= 5. METADATA VĂN BẢN =======================
CO_QUAN = {"QH": "Quốc hội", "NĐ-CP": "Chính phủ", "TT-BTC": "Bộ Tài chính", "TT-BNV": "Bộ Nội vụ",
           "TT-BYT": "Bộ Y tế", "TT-BNNMT": "Bộ Nông nghiệp và Môi trường"}
RE_NGAY = r"ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})"


def co_quan(so_hieu):
    for k, v in CO_QUAN.items():
        if k in so_hieu:
            return v
    return None


def _ngay(m):
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1))).isoformat()
    except ValueError:
        return None


RE_HL = r"có\s+hiệu\s+l\S{1,3}(?:\s+th\S{1,2}\s+\S{1,6})?(?:\s+k\S{1,2})?\s*t\S{1,2}\s*"
RE_DIEU_HL = r"Đi\S{1,4}\s*\d{1,3}\s*[\.:]\s*(?:Hiệu\s+l\S{1,3}|Đi\S{1,3}u\s+kho\S{1,3}n\s+thi\s+h\S{2,4}|Tổ chức thực hiện và hiệu)"


def tim_hieu_luc(phang, ban_hanh):
    """Ngày có hiệu lực: tìm câu "... có hiệu lực (thi hành) (kể) từ ngày ..." – ưu tiên trong Điều
    "Hiệu lực thi hành" / "Điều khoản thi hành". Lấy câu ĐẦU TIÊN (khoản 1 = hiệu lực chung của văn bản;
    các khoản sau thường là hiệu lực riêng của từng Điều). Trả về (ngày, câu làm bằng chứng)."""
    vts = [m.start() for m in re.finditer(RE_DIEU_HL, phang)]
    vung = [phang[vts[-1]:vts[-1] + 3000]] if vts else []
    vung.append(phang)
    for v in vung:
        ung_vien = []
        def chu_ngu(m):
            # câu có chủ ngữ "Luật này / Nghị định này / Thông tư này ..." được ưu tiên (0) hơn câu khác (1)
            truoc = v[max(0, m.start() - 25):m.start()]
            return 0 if re.search(r"(Luật|luật|định|tư|quyết)\s+\S{2,4}\s*$", truoc) else 1

        for m in re.finditer(RE_HL + RE_NGAY, v, re.I):
            ung_vien.append((chu_ngu(m), m.start(), _ngay(m), m.group(0)))
        for m in re.finditer(RE_HL + r"ngày\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", v, re.I):
            ung_vien.append((chu_ngu(m), m.start(), _ngay(m), m.group(0)))
        for m in re.finditer(RE_HL + r"ngày\s*(ký|được\s+thông\s+qua)", v, re.I):
            ung_vien.append((chu_ngu(m), m.start(), ban_hanh, m.group(0)))
        ung_vien = [u for u in ung_vien if u[2]]
        if ung_vien:
            u = min(ung_vien)
            if u[0] == 0 or v is phang:
                truoc = v[max(0, u[1] - 30):u[1]]
                return u[2], (truoc + u[3]).strip()
    return None, None


def tim_ngay(text, so_hieu):
    phang = re.sub(r"\s+", " ", unicodedata.normalize("NFC", text))
    ban_hanh = hieu_luc = None
    if "QH" in so_hieu:
        m = re.search(r"th\S{1,3}ng\s+qua\s*" + RE_NGAY, phang, re.I)
    else:
        # "Hà Nội, ngày 30 tháng 6 năm 2026" ở đầu văn bản (OCR hay méo: "Hà Nộii", "Hả Nội")
        m = re.search(r"H\S{1,2}\s*N\S{1,3}i\s*,?\s*" + RE_NGAY, phang[:6000])
    if m:
        ban_hanh = _ngay(m)
    hieu_luc, bang_chung = tim_hieu_luc(phang, ban_hanh)
    return ban_hanh, hieu_luc


# ======================= 6. QUAN HỆ GIỮA CÁC VĂN BẢN (cho WP4, WP5) =======================
# Số hiệu văn bản: 254/2026/NĐ-CP, 108/2025/QH15, 91/2026/TT-BTC ... (OCR hay chèn dấu cách: "NĐ- CP")
RE_SO_HIEU = re.compile(r"\b(\d{1,3})\s*/\s*(\d{4})\s*/\s*(QH\s*\d{2}|NĐ\s*-\s*CP|TT\s*-\s*[A-ZĐ]{2,6}|QĐ\s*-\s*TTg|NQ\s*-\s*[A-Z]{2,6})")


def chuan_so_hieu(m):
    hau_to = re.sub(r"\s+", "", m.group(3))
    return f"{m.group(1)}/{m.group(2)}/{hau_to}"


def tim_so_hieu(text):
    return [chuan_so_hieu(m) for m in RE_SO_HIEU.finditer(text)]


def ten_van_ban_day_du(preamble):
    """Trích tên đầy đủ của văn bản ở phần đầu: sau chữ LUẬT / NGHỊ ĐỊNH / THÔNG TƯ, trước 'Căn cứ'."""
    phang = re.sub(r"\s+", " ", preamble)
    m = re.search(r"\b(BỘ LUẬT|LUẬT|NGHỊ ĐỊNH|THÔNG TƯ)\b\s+(.{10,400}?)\s+(?:C\S{0,3}n\s+c\S{1,2}\b|\S{0,4}n\s+cứ\b|Theo đề nghị)", phang)
    return (m.group(1) + " " + m.group(2)).strip() if m else None


def cap_van_ban(so_hieu):
    """Cấp hiệu lực: Luật/Nghị quyết QH = 3, Nghị định/QĐ-TTg = 2, Thông tư = 1."""
    if "/QH" in so_hieu or "NQ-" in so_hieu:
        return 3
    if "NĐ-CP" in so_hieu or "QĐ-TTg" in so_hieu:
        return 2
    return 1


def hop_le(rel, nguon, dich):
    """Ràng buộc thứ bậc: chỉ văn bản cùng cấp mới sửa đổi/thay thế nhau; văn bản cấp dưới mới hướng dẫn cấp trên."""
    if rel in ("AMENDS", "REPLACES", "PARTIALLY_REPEALS"):
        return cap_van_ban(nguon) == cap_van_ban(dich)
    if rel == "GUIDES":
        return cap_van_ban(nguon) < cap_van_ban(dich)
    return True


def trich_quan_he(so_hieu, units):
    """Trích tự động quan hệ với văn bản khác. Trả về (tên văn bản, list {relation, target, evidence}).
    relation: AMENDS (sửa đổi) / REPLACES (thay thế toàn bộ) / PARTIALLY_REPEALS (bãi bỏ một phần)
              / GUIDES (hướng dẫn) / REFERENCES (có nhắc tới)."""
    preamble = next((u["content"] for u in units if u["unit_type"] == "preamble"), "")
    ten = ten_van_ban_day_du(preamble) or ""
    ket_qua, da_co = [], set()

    def them(rel, target, bang_chung):
        if target != so_hieu and (rel, target) not in da_co and hop_le(rel, so_hieu, target):
            da_co.add((rel, target))
            ket_qua.append({"relation": rel, "target": target, "evidence": bang_chung[:300]})

    # 1) Tên văn bản: "sửa đổi, bổ sung ..." -> AMENDS ; "quy định chi tiết / hướng dẫn ..." -> GUIDES
    ten_thuong = ten.lower()
    la_sua_doi = bool(re.search(r"s\S{1,2}a\s+đ\S{1,2}i|b\S{1,2}\s+sung", ten_thuong))
    la_huong_dan = bool(re.search(r"hướng\s+dẫn|quy\s+định\s+chi\s+ti|quy\s+định\s+chỉ\s+ti", ten_thuong))
    for sh in tim_so_hieu(ten)[:3]:
        if la_sua_doi:
            them("AMENDS", sh, ten)
        elif la_huong_dan:
            them("GUIDES", sh, ten)
    # 2) Điều hiệu lực / chuyển tiếp: xét TỪNG CÂU có "thay thế" hoặc "hết hiệu lực"
    for u in units:
        if u["unit_type"] != "article" or not re.search(r"(?i)hiệu lực|thi hành|chuyển tiếp", u.get("title") or ""):
            continue
        noi_dung = " ".join(x["content"] for x in units if x["unit_id"] == u["unit_id"]
                            or x["unit_id"].startswith(u["unit_id"] + "_"))
        for cau in re.split(r"(?<=[\.;])\s+|\n", re.sub(r"[ \t]+", " ", noi_dung)):
            if "chấm dứt hiệu lực mã số" in cau:
                continue
            m = re.search(r"thay\s+thế(?:\s+cho)?(?:\s+các)?\s*:?\s*(.*)", cau)
            if m and not re.match(r"\s*(mẫu|Mẫu|Phụ lục|phụ lục|biểu)", m.group(1)):
                for sh in tim_so_hieu(m.group(1))[:1]:
                    them("REPLACES", sh, cau)
            if re.search(r"hết\s+hiệu\s+lực", cau):
                truoc = re.split(r"hết\s+hiệu\s+lực", cau)[0]
                sh = tim_so_hieu(truoc)
                if not sh:
                    continue
                # câu bắt đầu bằng Điểm/Khoản/Điều/Mục/Chương/Phụ lục -> chỉ bãi bỏ một phần
                mot_phan = re.match(r"\s*\W*\s*(Điểm|điểm|Khoản|khoản|Điều|điều|Mục|Chương|Phụ lục|các điều|Các điều|\S{0,3}\s*Mục)", truoc)
                them("PARTIALLY_REPEALS" if mot_phan else "REPLACES", sh[0], cau)
    # 3) Mọi văn bản được nhắc tới -> REFERENCES (kèm số lần)
    dem = {}
    for u in units:
        for sh in tim_so_hieu(u["content"]):
            dem[sh] = dem.get(sh, 0) + 1
    for sh, n in sorted(dem.items(), key=lambda x: -x[1]):
        if sh != so_hieu and not any(r["target"] == sh for r in ket_qua):
            ket_qua.append({"relation": "REFERENCES", "target": sh, "evidence": f"được nhắc {n} lần"})
    return ten, ket_qua
