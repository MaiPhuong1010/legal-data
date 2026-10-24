"""
VÍ DỤ CHO NHÓM 2: đọc chunk, lọc theo ngày hiệu lực, tìm kiếm BM25, (tùy chọn) tạo vector bge-m3.

    pip install rank-bm25
    python src\\tim_kiem.py "thời điểm lập hóa đơn đối với cung cấp dịch vụ" --ngay 2026-08-15

Tạo vector (cần máy có PyTorch, ~2.3 GB model):
    pip install sentence-transformers
    python src\\tim_kiem.py --tao-vector
"""
import argparse
import json
import re
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CHUNKS = ROOT / "output" / "chunks" / "legal_structure_chunks.jsonl"


def doc_chunks(path=CHUNKS):
    with open(path, encoding="utf-8") as f:
        return [json.loads(dong) for dong in f]


def con_hieu_luc(c, ngay):
    """Lọc theo thời gian (vai trò WP4): chunk áp dụng được tại `ngay` (YYYY-MM-DD)."""
    if c["effective_from"] and c["effective_from"] > ngay:
        return False
    if c["effective_to"] and c["effective_to"] < ngay:
        return False
    return c["status"] != "replaced"


def tach_tu(s):
    s = unicodedata.normalize("NFC", s.lower())
    return re.findall(r"\w+", s)


def tim_bm25(chunks, cau_hoi, k=5):
    from rank_bm25 import BM25Okapi
    bm25 = BM25Okapi([tach_tu(c["text_for_embedding"]) for c in chunks])
    diem = bm25.get_scores(tach_tu(cau_hoi))
    top = sorted(range(len(chunks)), key=lambda i: -diem[i])[:k]
    return [(diem[i], chunks[i]) for i in top]


def tao_vector(chunks, out=ROOT / "output" / "chunks" / "vectors_bge_m3.npy"):
    """Tạo embedding dense bằng bge-m3 cho từng chunk (đầu vào cho Qdrant / pgvector của WP2)."""
    import numpy as np
    from sentence_transformers import SentenceTransformer
    model = SentenceTransformer("BAAI/bge-m3")
    vec = model.encode([c["text_for_embedding"] for c in chunks], batch_size=16, normalize_embeddings=True,
                       show_progress_bar=True)
    np.save(out, vec)
    (out.with_suffix(".ids.json")).write_text(json.dumps([c["chunk_id"] for c in chunks]), encoding="utf-8")
    print("=> Đã lưu", vec.shape, "vào", out)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("cau_hoi", nargs="?", default="thời điểm lập hóa đơn đối với cung cấp dịch vụ")
    ap.add_argument("--ngay", default=None, help="Ngày phát sinh nghiệp vụ YYYY-MM-DD (lọc hiệu lực)")
    ap.add_argument("--tao-vector", action="store_true")
    a = ap.parse_args()
    ds = doc_chunks()
    if a.tao_vector:
        tao_vector(ds)
        raise SystemExit
    if a.ngay:
        ds = [c for c in ds if con_hieu_luc(c, a.ngay)]
    print(f"Tìm trong {len(ds)} chunk: {a.cau_hoi!r}\n")
    for diem, c in tim_bm25(ds, a.cau_hoi):
        print(f"[{diem:6.2f}] {c['chunk_id']}  (hiệu lực {c['effective_from']}, trang {c['page']})")
        print(f"         {c['context_header']}")
        print(f"         {c['content'][:160]}...\n")
