import importlib

for ten in ["docx", "pdfplumber", "pymupdf", "pandas", "openpyxl", "pydantic"]:
    importlib.import_module(ten)          # báo lỗi ngay nếu thư viện nào chưa cài
print("1. Thu vien: OK")

from transformers import AutoTokenizer
print("2. Dang tai tokenizer bge-m3 (lan dau hoi lau)...")
tok = AutoTokenizer.from_pretrained("BAAI/bge-m3")

cau = "Điều 9. Thời điểm lập hóa đơn đối với bán hàng hóa là thời điểm chuyển giao quyền sở hữu."
so_token = len(tok.encode(cau, add_special_tokens=False))
print("3. Cau mau:", cau)
print("   So token:", so_token)
print("=> Moi truong da san sang!")
