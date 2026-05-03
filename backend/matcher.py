# matcher.py
# -------------------------------------------------------
# 기존 후보기술 리스트와 문서 매칭 엔진
# -------------------------------------------------------

import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Dict, Optional, Tuple

from rapidfuzz import fuzz, process

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

from metadata_extractor import DocMeta


# ============================================================
# Data structures
# ============================================================
@dataclass
class RefItem:
    """기존 후보기술 리스트 항목."""
    index: int = 0              # 기존 순번
    title: str = ""             # 기술명
    doi: str = ""               # DOI (있는 경우)
    patent_no: str = ""         # 특허번호 (있는 경우)
    category: str = ""          # 분류/카테고리
    extra: Dict = field(default_factory=dict)  # 추가 정보


@dataclass
class MatchResult:
    """매칭 결과."""
    doc_meta: DocMeta           # 입력 문서
    ref_item: Optional[RefItem] = None  # 매칭된 기존 항목
    score: float = 0.0          # 매칭 점수 (0-100)
    match_type: str = "unmatched"  # exact_doi, exact_patent, fuzzy, unmatched
    needs_review: bool = False  # 60% 미만 수동 검토 필요
    new_index: int = 0          # 변환 후 번호

    @property
    def status_label(self) -> str:
        if self.match_type == "exact_filename":
            return "✅ 파일명 정확 매칭"
        elif self.match_type == "fuzzy_filename":
            return "🟢 파일명 유사 매칭"
        elif self.match_type.startswith("exact"):
            return "✅ 정확 매칭"
        elif self.match_type == "fuzzy" and self.score >= 80:
            return "🟢 자동 매칭"
        elif self.match_type == "fuzzy" and self.score >= 60:
            return "🟡 검토 필요"
        elif self.match_type == "fuzzy" and self.score >= 30:
            return "🟠 낮은 유사도"
        elif self.needs_review:
            return "🔴 수동 확인"
        else:
            return "⚪ 미매칭(신규)"


# ============================================================
# Reference List Loading
# ============================================================
def load_reference_list(file_path: str, title_col: str = None,
                        doi_col: str = None, patent_col: str = None,
                        index_col: str = None, category_col: str = None
                        ) -> List[RefItem]:
    """
    엑셀/CSV에서 기존 후보기술 리스트를 로드합니다.

    컬럼명을 자동 감지하거나 명시적으로 지정할 수 있습니다.
    자동 감지 키워드:
      - title: 기술명, 제목, title, name, 명칭
      - doi: doi
      - patent: 특허, patent, 출원번호, 등록번호
      - index: 번호, no, index, #, 순번
      - category: 분류, 카테고리, category, type
    """
    ext = Path(file_path).suffix.lower()

    if ext in ('.xlsx', '.xls'):
        if not HAS_PANDAS:
            raise RuntimeError("pandas 필요: pip install pandas openpyxl")
        df = pd.read_excel(file_path, engine='openpyxl')
    elif ext == '.csv':
        if not HAS_PANDAS:
            raise RuntimeError("pandas 필요: pip install pandas")
        df = pd.read_csv(file_path, encoding='utf-8-sig')
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext}")

    # 컬럼 자동 감지
    cols = [str(c).strip() for c in df.columns]

    def _find_col(keywords, explicit=None):
        if explicit and explicit in cols:
            return explicit
        for kw in keywords:
            for c in cols:
                if kw in c.lower():
                    return c
        return None

    title_c = _find_col(['기술명', '기술 명', '제목', 'title', 'name', '명칭', '발명'], title_col)
    doi_c = _find_col(['doi'], doi_col)
    patent_c = _find_col(['특허', 'patent', '출원번호', '등록번호', '출원'], patent_col)
    index_c = _find_col(['개요서 번호', '개요서번호', '이전 순번', '이전순번', '기존 번호', '기존번호', '번호', 'no', 'index', '#', '순번', '순서'], index_col)
    cat_c = _find_col(['분류', '카테고리', 'category', 'type', '유형'], category_col)
    filename_c = _find_col(['파일명', 'filename', '원문 파일명', '원문파일명', 'file_name', '파일 이름'])

    items = []
    for i, row in df.iterrows():
        extra = {}
        if filename_c and pd.notna(row.get(filename_c)):
            extra['filename'] = str(row[filename_c]).strip()
        item = RefItem(
            index=int(row[index_c]) if index_c and pd.notna(row.get(index_c)) else i + 1,
            title=str(row[title_c]).strip() if title_c and pd.notna(row.get(title_c)) else "",
            doi=str(row[doi_c]).strip() if doi_c and pd.notna(row.get(doi_c)) else "",
            patent_no=str(row[patent_c]).strip() if patent_c and pd.notna(row.get(patent_c)) else "",
            category=str(row[cat_c]).strip() if cat_c and pd.notna(row.get(cat_c)) else "",
            extra=extra,
        )
        # 빈 행 스킵
        if item.title or item.doi or item.patent_no or extra.get('filename'):
            items.append(item)

    return items


# ============================================================
# Matching Logic
# ============================================================
def _normalize_id(s: str) -> str:
    """식별자 정규화 (공백, 하이픈 제거)."""
    return re.sub(r'[\s\-/]', '', s).lower().strip()


def match_by_doi(doc: DocMeta, ref_list: List[RefItem]) -> Optional[Tuple[RefItem, float]]:
    """DOI로 정확 매칭합니다."""
    if not doc.doi:
        return None
    doc_doi = _normalize_id(doc.doi)
    for ref in ref_list:
        if ref.doi and _normalize_id(ref.doi) == doc_doi:
            return ref, 100.0
    return None


def match_by_patent_no(doc: DocMeta, ref_list: List[RefItem]) -> Optional[Tuple[RefItem, float]]:
    """특허번호로 정확 매칭합니다."""
    if not doc.patent_no:
        return None
    doc_pno = _normalize_id(doc.patent_no)
    for ref in ref_list:
        if ref.patent_no and _normalize_id(ref.patent_no) == doc_pno:
            return ref, 100.0
    return None


def fuzzy_match_title(doc_title: str, ref_title: str) -> float:
    """두 제목 간 유사도를 계산합니다 (0-100).
    
    한국어 기술명 ↔ 영어 논문 제목처럼 언어가 다를 때를 위해
    영문 키워드 교집합 보너스 점수를 추가 적용합니다.
    """
    if not doc_title or not ref_title:
        return 0.0
    # 여러 매칭 전략의 가중 평균
    ratio = fuzz.ratio(doc_title, ref_title)
    token_sort = fuzz.token_sort_ratio(doc_title, ref_title)
    token_set = fuzz.token_set_ratio(doc_title, ref_title)
    partial = fuzz.partial_ratio(doc_title, ref_title)
    # 가중 평균 (token_set에 높은 가중치)
    base_score = (ratio * 0.15 + token_sort * 0.25 + token_set * 0.35 + partial * 0.25)

    # ── 한-영 혼합 키워드 교집합 보너스 ──────────────────────────
    # 영문 단어(4글자 이상)를 추출해서 교집합 비율로 보너스 산출
    def _eng_keywords(s: str) -> set:
        tokens = re.findall(r'[a-zA-Z]{4,}', s.lower())
        return set(tokens)
    kw1 = _eng_keywords(doc_title)
    kw2 = _eng_keywords(ref_title)
    if kw1 and kw2:
        overlap = len(kw1 & kw2)
        union = len(kw1 | kw2)
        kw_bonus = (overlap / union) * 40.0  # 최대 +40점
        base_score = min(100.0, base_score + kw_bonus)

    return base_score


def match_by_filename(doc: DocMeta, ref_list: List[RefItem]) -> Optional[Tuple[RefItem, float]]:
    """파일명 stem으로 ref_list의 파일명 컬럼과 매칭합니다.
    ref_item.extra['filename']이 설정된 경우에만 동작합니다.
    """
    doc_stem = Path(doc.filename).stem.lower().strip()
    if not doc_stem:
        return None
    best_ref = None
    best_score = 0.0
    for ref in ref_list:
        ref_filename = (ref.extra.get('filename') or '').strip()
        if not ref_filename:
            continue
        ref_stem = Path(ref_filename).stem.lower().strip()
        # 정확히 같은 파일명
        if doc_stem == ref_stem:
            return ref, 100.0
        # fuzzy
        score = fuzz.token_set_ratio(doc_stem, ref_stem)
        if score > best_score:
            best_score = score
            best_ref = ref
    if best_score >= 80.0 and best_ref:
        return best_ref, float(best_score)
    return None


def match_by_title(doc: DocMeta, ref_list: List[RefItem],
                   threshold: float = 30.0) -> Optional[Tuple[RefItem, float]]:
    """제목 유사도로 매칭합니다 (최고 점수 반환)."""
    doc_title = doc.title
    if not doc_title:
        # 파일명을 대체 제목으로 사용
        doc_title = Path(doc.filename).stem

    best_ref = None
    best_score = 0.0

    for ref in ref_list:
        score = fuzzy_match_title(doc_title, ref.title)
        if score > best_score:
            best_score = score
            best_ref = ref

    if best_score >= threshold and best_ref:
        return best_ref, best_score
    return None


# ============================================================
# Main Matching Pipeline
# ============================================================
def match_documents(doc_metas: List[DocMeta],
                    ref_list: List[RefItem],
                    auto_threshold: float = 80.0,
                    review_threshold: float = 60.0
                    ) -> List[MatchResult]:
    """
    입력 문서 리스트를 기존 리스트와 매칭합니다.

    매칭 우선순위:
      1. DOI 정확 매칭
      2. 특허번호 정확 매칭
      3. 제목 Fuzzy Matching

    Args:
        doc_metas: 입력 문서 메타데이터 리스트
        ref_list: 기존 후보기술 리스트
        auto_threshold: 자동 매칭 승인 기준 (default: 80%)
        review_threshold: 수동 검토 기준 (default: 60%)

    Returns:
        List[MatchResult]: 매칭 결과 리스트
    """
    results = []
    used_refs = set()  # 이미 매칭된 ref의 index

    for doc in doc_metas:
        result = MatchResult(doc_meta=doc)

        # Step 1: DOI 매칭
        doi_match = match_by_doi(doc, ref_list)
        if doi_match and doi_match[0].index not in used_refs:
            result.ref_item = doi_match[0]
            result.score = doi_match[1]
            result.match_type = "exact_doi"
            result.needs_review = False
            used_refs.add(doi_match[0].index)
            results.append(result)
            continue

        # Step 2: 특허번호 매칭
        patent_match = match_by_patent_no(doc, ref_list)
        if patent_match and patent_match[0].index not in used_refs:
            result.ref_item = patent_match[0]
            result.score = patent_match[1]
            result.match_type = "exact_patent"
            result.needs_review = False
            used_refs.add(patent_match[0].index)
            results.append(result)
            continue

        # Step 2-B: 파일명 매칭 (엑셀에 파일명 컬럼이 있을 때)
        available_refs_fn = [r for r in ref_list if r.index not in used_refs]
        fn_match = match_by_filename(doc, available_refs_fn)
        if fn_match and fn_match[0].index not in used_refs:
            result.ref_item = fn_match[0]
            result.score = fn_match[1]
            result.match_type = "exact_filename" if fn_match[1] >= 100.0 else "fuzzy_filename"
            result.needs_review = fn_match[1] < auto_threshold
            if fn_match[1] >= auto_threshold:
                used_refs.add(fn_match[0].index)
            results.append(result)
            continue

        # Step 3: 제목 Fuzzy Matching
        # 이미 매칭된 ref는 제외
        available_refs = [r for r in ref_list if r.index not in used_refs]
        title_match = match_by_title(doc, available_refs)

        if title_match:
            ref, score = title_match
            result.ref_item = ref
            result.score = score
            result.match_type = "fuzzy"

            if score >= auto_threshold:
                result.needs_review = False
                used_refs.add(ref.index)
            elif score >= review_threshold:
                result.needs_review = True
                # used_refs에 추가하지 않음 (검토 후 확정)
            else:
                result.needs_review = True
        else:
            result.match_type = "unmatched"
            result.needs_review = True

        results.append(result)

    return results


# ============================================================
# 결과를 DataFrame으로 변환 (UI 표시용)
# ============================================================
def results_to_dataframe(results: List[MatchResult]) -> 'pd.DataFrame':
    """매칭 결과를 pandas DataFrame으로 변환합니다."""
    if not HAS_PANDAS:
        raise RuntimeError("pandas 필요")

    rows = []
    for r in results:
        rows.append({
            "파일명": r.doc_meta.filename,
            "문서유형": r.doc_meta.doc_type,
            "추출 제목": r.doc_meta.title[:60] if r.doc_meta.title else "",
            "DOI": r.doc_meta.doi,
            "특허번호": r.doc_meta.patent_no,
            "매칭 기존번호": r.ref_item.index if r.ref_item else "-",
            "매칭 기술명": r.ref_item.title[:40] if r.ref_item else "-",
            "일치도(%)": f"{r.score:.1f}",
            "매칭방식": r.match_type,
            "상태": r.status_label,
            "새 번호": r.new_index if r.new_index else "-",
        })

    return pd.DataFrame(rows)
