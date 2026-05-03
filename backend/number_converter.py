# number_converter.py
# -------------------------------------------------------
# 번호 변환 매핑 모듈
# -------------------------------------------------------

import os
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

from matcher import MatchResult


# ============================================================
# Data structures
# ============================================================
@dataclass
class ConvertedDoc:
    """번호 변환 결과."""
    match_result: MatchResult
    old_index: int = 0
    new_index: int = 0
    note: str = ""   # 비고 (변환 근거)


# ============================================================
# 매핑 테이블 로드/저장
# ============================================================
def load_mapping_table(file_path: str) -> Dict[int, int]:
    """
    번호 변환 매핑 테이블을 로드합니다.

    형식 (Excel/CSV):
      old_index | new_index | 비고(선택)

    형식 (JSON):
      {"1": 5, "2": 3, ...}

    Returns:
        Dict[int, int]: {기존번호: 새번호}
    """
    ext = Path(file_path).suffix.lower()
    mapping = {}

    if ext == '.json':
        with open(file_path, 'r', encoding='utf-8') as f:
            raw = json.load(f)
        for k, v in raw.items():
            mapping[int(k)] = int(v)

    elif ext in ('.xlsx', '.xls', '.csv'):
        if not HAS_PANDAS:
            raise RuntimeError("pandas 필요: pip install pandas openpyxl")

        if ext == '.csv':
            df = pd.read_csv(file_path, encoding='utf-8-sig')
        else:
            df = pd.read_excel(file_path, engine='openpyxl')

        cols = [str(c).strip().lower() for c in df.columns]

        # 컬럼 자동 감지
        old_col = None
        new_col = None
        for i, c in enumerate(cols):
            if any(kw in c for kw in ['기존', 'old', '원래', '변환전', '이전']):
                old_col = df.columns[i]
            elif any(kw in c for kw in ['새', 'new', '변환후', '이후', '신규']):
                new_col = df.columns[i]

        # 자동 감지 실패 시 첫 두 컬럼 사용
        if old_col is None and len(df.columns) >= 2:
            old_col = df.columns[0]
        if new_col is None and len(df.columns) >= 2:
            new_col = df.columns[1]

        if old_col is None or new_col is None:
            raise ValueError("매핑 테이블에서 기존번호/새번호 컬럼을 찾을 수 없습니다.")

        for _, row in df.iterrows():
            old_val = row[old_col]
            new_val = row[new_col]
            if pd.notna(old_val) and pd.notna(new_val):
                mapping[int(old_val)] = int(new_val)
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext}")

    return mapping


def save_mapping_table(mapping: Dict[int, int], output_path: str):
    """매핑 테이블을 저장합니다."""
    ext = Path(output_path).suffix.lower()

    if ext == '.json':
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump({str(k): v for k, v in mapping.items()}, f,
                      ensure_ascii=False, indent=2)

    elif ext in ('.xlsx', '.csv'):
        if not HAS_PANDAS:
            raise RuntimeError("pandas 필요")
        df = pd.DataFrame([
            {"기존번호": k, "새번호": v}
            for k, v in sorted(mapping.items())
        ])
        if ext == '.csv':
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
        else:
            df.to_excel(output_path, index=False, engine='openpyxl')
    else:
        raise ValueError(f"지원하지 않는 형식: {ext}")


# ============================================================
# 번호 변환 적용
# ============================================================
def apply_conversion(match_results: List[MatchResult],
                     mapping: Optional[Dict[int, int]] = None
                     ) -> List[ConvertedDoc]:
    """
    매칭 결과에 번호 변환을 적용합니다.

    매핑이 없으면 기존 번호를 유지합니다.
    미매칭 문서에는 새 번호를 자동 부여합니다.

    Args:
        match_results: 매칭 결과 리스트
        mapping: {기존번호: 새번호} 매핑 (없으면 순차 번호)

    Returns:
        List[ConvertedDoc]: 변환 결과
    """
    converted = []
    used_new_indices = set()

    if mapping:
        used_new_indices = set(mapping.values())

    # 매핑이 있는 경우 먼저 처리
    for mr in match_results:
        old_idx = mr.ref_item.index if mr.ref_item else 0
        new_idx = 0
        note = ""

        if mapping and old_idx in mapping:
            new_idx = mapping[old_idx]
            note = f"매핑 적용: {old_idx} → {new_idx}"
        elif mr.ref_item and mr.match_type != "unmatched":
            new_idx = old_idx  # 매핑 없으면 기존 번호 유지
            note = "기존 번호 유지"

        mr.new_index = new_idx

        converted.append(ConvertedDoc(
            match_result=mr,
            old_index=old_idx,
            new_index=new_idx,
            note=note,
        ))

    # 미매칭(신규) 문서에 새 번호 부여
    max_idx = max(used_new_indices) if used_new_indices else 0
    all_indices = {c.new_index for c in converted if c.new_index > 0}
    max_idx = max(max_idx, max(all_indices) if all_indices else 0)

    next_idx = max_idx + 1
    for c in converted:
        if c.new_index == 0:
            c.new_index = next_idx
            c.match_result.new_index = next_idx
            c.note = f"신규 번호 부여: {next_idx}"
            next_idx += 1

    # 새 번호 기준으로 정렬
    converted.sort(key=lambda c: c.new_index)

    return converted


# ============================================================
# 변환 리포트 생성
# ============================================================
def generate_conversion_report(converted: List[ConvertedDoc],
                               output_path: str):
    """
    번호 변환 결과를 리포트로 저장합니다.

    Args:
        converted: 변환 결과 리스트
        output_path: 출력 파일 경로 (.xlsx 또는 .csv)
    """
    if not HAS_PANDAS:
        raise RuntimeError("pandas 필요")

    rows = []
    for c in converted:
        mr = c.match_result
        rows.append({
            "새 번호": c.new_index,
            "기존 번호": c.old_index if c.old_index else "-",
            "파일명": mr.doc_meta.filename,
            "문서유형": mr.doc_meta.doc_type,
            "추출 제목": mr.doc_meta.title[:80] if mr.doc_meta.title else "",
            "매칭 기술명": mr.ref_item.title[:60] if mr.ref_item else "-",
            "일치도(%)": f"{mr.score:.1f}",
            "매칭방식": mr.match_type,
            "상태": mr.status_label,
            "비고": c.note,
        })

    df = pd.DataFrame(rows)
    ext = Path(output_path).suffix.lower()

    if ext == '.csv':
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
    else:
        df.to_excel(output_path, index=False, engine='openpyxl')

    return df


# ============================================================
# 변환 상태 저장/로드 (수정 실행 모드용)
# ============================================================
def save_conversion_state(converted: List[ConvertedDoc], output_path: str):
    """변환 상태를 JSON으로 저장합니다 (수정 모드 재사용)."""
    state = []
    for c in converted:
        state.append({
            "filename": c.match_result.doc_meta.filename,
            "source_path": c.match_result.doc_meta.source_path,
            "doc_type": c.match_result.doc_meta.doc_type,
            "old_index": c.old_index,
            "new_index": c.new_index,
            "match_type": c.match_result.match_type,
            "score": c.match_result.score,
            "ref_title": c.match_result.ref_item.title if c.match_result.ref_item else "",
            "note": c.note,
        })
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def load_conversion_state(state_path: str) -> List[dict]:
    """이전 변환 상태를 로드합니다."""
    with open(state_path, 'r', encoding='utf-8') as f:
        return json.load(f)
