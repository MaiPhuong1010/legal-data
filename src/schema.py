"""
JSON SCHEMA CHUẨN CỦA DỮ LIỆU ĐẦU RA (định nghĩa bằng Pydantic, theo đề WP1 / WP3 / WP4 / WP5)

Nhóm 2 dùng file này để đọc và kiểm tra dữ liệu:
    from schema import LegalDocument, LegalChunk
    doc = LegalDocument.model_validate_json(open("output/structure/ND254_2026.json", encoding="utf-8").read())

Xuất JSON Schema ra output/schema/*.json:
    python src\\schema.py
"""
import json
from pathlib import Path
from typing import List, Literal, Optional

from pydantic import BaseModel, ConfigDict, Field

UnitType = Literal["preamble", "article", "clause", "point", "appendix"]
Status = Literal["active", "not_yet_effective", "replaced", "expired", "unknown"]
Relation = Literal["AMENDS", "REPLACES", "PARTIALLY_REPEALS", "GUIDES", "REFERENCES"]


class LegalUnit(BaseModel):
    """1 đơn vị pháp lý: Điều / Khoản / Điểm / Phụ lục / Phần mở đầu (schema legal_units của WP1)."""
    model_config = ConfigDict(extra="forbid")
    unit_id: str = Field(description="Mã duy nhất, vd ND254_2026_D9_K2_Pa")
    unit_type: UnitType
    parent_id: Optional[str] = Field(None, description="unit_id của đơn vị cha (Khoản -> Điều, Điểm -> Khoản)")
    part: Optional[str] = None
    chapter: Optional[str] = Field(None, description="Số Chương (La Mã)")
    chapter_title: Optional[str] = None
    section: Optional[str] = Field(None, description="Số Mục")
    section_title: Optional[str] = None
    article: Optional[int] = None
    clause: Optional[int] = None
    point: Optional[str] = None
    title: Optional[str] = Field(None, description="Tên Điều / tên Phụ lục")
    content: str = Field(description="Nguyên văn nội dung riêng của đơn vị (không gồm đơn vị con)")
    page: Optional[int] = Field(None, description="Trang bắt đầu trong file gốc")
    has_table: bool = False
    token_count: int
    token_count_full: Optional[int] = Field(None, description="Token của cả đơn vị kể cả đơn vị con (Điều, Khoản)")
    effective_from: Optional[str] = None
    effective_to: Optional[str] = None
    status: Status
    effective_note: Optional[str] = Field(None, description="Có khi Điều này có hiệu lực riêng khác toàn văn bản")
    is_confident: bool = Field(description="false = OCR nhiều ký tự lạ hoặc số thứ tự bị nhảy -> cần người kiểm duyệt")


class LegalDocument(BaseModel):
    """1 văn bản có cấu trúc (output/structure/<document_id>.json)."""
    model_config = ConfigDict(extra="forbid")
    document_id: str
    document_no: str
    document_type: Literal["Luật", "Bộ luật", "Nghị định", "Thông tư", "Khác"]
    title: Optional[str]
    title_in_document: Optional[str]
    issuer: Optional[str]
    issued_date: Optional[str]
    effective_from: Optional[str]
    effective_to: Optional[str]
    status: Status
    amends: List[str] = []
    amended_by: List[str] = []
    replaces: List[str] = []
    replaced_by: List[str] = []
    partially_repeals: List[str] = []
    partially_repealed_by: List[str] = []
    guides: List[str] = []
    guided_by: List[str] = []
    groups: List[str]
    source_url: Optional[str]
    source_file: List[str]
    source_format: Optional[str]
    metadata_source: dict
    extraction: dict
    parse_warnings: List[str]
    legal_units: List[LegalUnit]


class LegalChunk(BaseModel):
    """1 chunk phương án Legal-structure (output/chunks/legal_structure_chunks.jsonl)."""
    model_config = ConfigDict(extra="forbid")
    chunk_id: str
    strategy: Literal["legal_structure"]
    chunk_type: str
    document_id: str
    document_no: str
    document_type: str
    chapter: Optional[str]
    section: Optional[str]
    article: Optional[int]
    clause: Optional[int]
    point: Optional[str]
    unit_ids: List[str]
    parent_id: Optional[str]
    context_header: str
    content: str
    text_for_embedding: str = Field(description="= context_header + xuống dòng + content -> đem đi tạo vector")
    token_count: int = Field(le=512)
    page: Optional[int] = Field(description="Trang bắt đầu trong file gốc")
    page_end: Optional[int] = Field(description="Trang kết thúc trong file gốc")
    has_table: bool
    effective_from: Optional[str]
    effective_to: Optional[str]
    status: Status
    is_confident: bool
    document_title: Optional[str]
    issuer: Optional[str]
    issued_date: Optional[str]
    amended_by: List[str]
    replaced_by: List[str]
    source_file: Optional[str]
    source_url: Optional[str]
    extraction_method: Optional[str]


class FixedChunk(BaseModel):
    """1 chunk phương án Fixed-token (để so sánh)."""
    model_config = ConfigDict(extra="forbid")
    chunk_id: str
    strategy: Literal["fixed_token"]
    document_id: str
    document_no: str
    document_type: str
    chunk_index: int
    token_start: int
    token_end: int
    articles_covered: List[int]
    starts_at_unit_boundary: bool
    content: str
    text_for_embedding: str
    token_count: int = Field(le=512)
    effective_from: Optional[str]
    effective_to: Optional[str]
    status: Status
    document_title: Optional[str]
    source_file: Optional[str]


class DocumentRelation(BaseModel):
    """1 quan hệ giữa 2 văn bản (output/quan_he_van_ban.json) – đầu vào cho WP4 / WP5."""
    model_config = ConfigDict(extra="forbid")
    source: str
    relation: Relation
    target: str
    method: str
    verified: bool
    evidence: Optional[str]


if __name__ == "__main__":
    out = Path(__file__).resolve().parent.parent / "output" / "schema"
    out.mkdir(parents=True, exist_ok=True)
    for model in (LegalUnit, LegalDocument, LegalChunk, FixedChunk, DocumentRelation):
        (out / f"{model.__name__}.schema.json").write_text(
            json.dumps(model.model_json_schema(), ensure_ascii=False, indent=2), encoding="utf-8")
    print("=> Đã xuất JSON Schema vào", out)
