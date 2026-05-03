# metadata_extractor.py
# -------------------------------------------------------
# PDF / 특허 메타데이터 추출 모듈
# -------------------------------------------------------

import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, List

import fitz  # PyMuPDF

try:
    import requests
    from bs4 import BeautifulSoup
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False


@dataclass
class DocMeta:
    """문서 메타데이터 컨테이너."""
    filename: str = ""
    doc_type: str = "paper"       # "paper" or "patent"
    title: str = ""
    doi: str = ""
    patent_no: str = ""
    institution: str = ""
    year: str = ""
    full_text: str = ""
    source_type: str = "pdf"      # "pdf" or "url"
    source_path: str = ""         # 원본 파일 경로 또는 URL
    figure_urls: List[str] = field(default_factory=list)  # 도면 이미지 URL 목록


# ============================================================
# DOI 추출
# ============================================================
_DOI_PATTERN = re.compile(r'\b(10\.\d{4,}/[^\s,;\]}>\"\']+)', re.IGNORECASE)


def _extract_doi(text: str) -> str:
    """텍스트에서 DOI를 추출합니다."""
    m = _DOI_PATTERN.search(text)
    if m:
        doi = m.group(1).rstrip(".")
        return doi
    return ""


# ============================================================
# 특허번호 추출
# ============================================================
_PATENT_PATTERNS = [
    # 한국 특허
    re.compile(r'(?:출원번호|출원\s*번호)\s*[:\s]*(\d{2}-\d{4}-\d{7})', re.IGNORECASE),
    re.compile(r'(?:등록번호|등록\s*번호)\s*[:\s]*(\d{2}-\d{4}-\d{7})', re.IGNORECASE),
    re.compile(r'\b(KR\s*\d{10,}[A-Z]?\d*)\b', re.IGNORECASE),
    re.compile(r'\b(\d{2}-\d{4}-\d{7})\b'),
    # 미국 특허
    re.compile(r'\b(US\s*\d{7,11}\s*[A-Z]?\d*)\b', re.IGNORECASE),
    re.compile(r'\b(US\s*\d{4}/\d{7,})\b', re.IGNORECASE),
    # 유럽/PCT
    re.compile(r'\b(EP\s*\d{7,})\b', re.IGNORECASE),
    re.compile(r'\b(WO\s*\d{4}/?\d{6,})\b', re.IGNORECASE),
    # 일본
    re.compile(r'\b(JP\s*\d{4}-\d{6,})\b', re.IGNORECASE),
]

_PATENT_KEYWORDS = [
    "claim", "patent", "출원", "특허", "청구항", "발명",
    "등록", "공개", "실용신안", "invention", "applicant"
]


def _extract_patent_no(text: str) -> str:
    """텍스트에서 특허번호를 추출합니다."""
    for pat in _PATENT_PATTERNS:
        m = pat.search(text)
        if m:
            return m.group(1).strip()
    return ""


def _detect_doc_type(text: str, filename: str = "") -> str:
    """특허/논문 유형을 판별합니다."""
    combined = (text[:5000] + " " + filename).lower()
    score = sum(1 for kw in _PATENT_KEYWORDS if kw in combined)
    return "patent" if score >= 2 else "paper"


# ============================================================
# 제목 추출 (파일명 + 텍스트)
# ============================================================
def _extract_title_from_filename(filename: str) -> str:
    """파일명에서 제목 후보를 추출합니다."""
    stem = Path(filename).stem
    # DOI 형식의 파일명인 경우 제목으로 부적절
    if _DOI_PATTERN.match(stem) or re.match(r'^[\d\-/]+$', stem):
        return ""
    # 출원번호 형식의 파일명
    if re.match(r'^\d{2}-\d{4}-\d{7}$', stem):
        return ""
    # US 특허번호 형식
    if re.match(r'^[A-Z]{2}\d{7,}', stem, re.IGNORECASE):
        return ""
    return stem


def _extract_title_from_text(text: str, doc_type: str) -> str:
    """텍스트 첫 부분에서 제목을 추출합니다."""
    lines = text.split('\n')
    # 빈 줄이 아닌 첫 몇 줄에서 제목 후보 탐색
    candidates = []
    for line in lines[:30]:
        line = line.strip()
        if not line or len(line) < 5:
            continue
        # 너무 짧거나 메타데이터성 줄은 스킵
        if len(line) > 200:
            continue
        if any(line.lower().startswith(p) for p in [
            "abstract", "keywords", "doi:", "http", "www.",
            "received", "accepted", "published", "copyright",
            "journal", "vol.", "volume", "issue"
        ]):
            continue
        candidates.append(line)
        if len(candidates) >= 3:
            break
    # 가장 긴 후보를 제목으로 선택 (일반적으로 제목이 가장 김)
    if candidates:
        return max(candidates, key=len)
    return ""


# ============================================================
# 연도 추출
# ============================================================
_YEAR_PATTERN = re.compile(r'\b(19\d{2}|20[0-3]\d)\b')


def _extract_year(text: str) -> str:
    """텍스트에서 출판/출원 연도를 추출합니다."""
    # 앞쪽 텍스트에서 연도 검색 (보통 앞부분에 있음)
    years = _YEAR_PATTERN.findall(text[:3000])
    if years:
        # 가장 최근 연도 반환
        return max(years, key=int)
    return ""


# ============================================================
# 기관 추출
# ============================================================
_INSTITUTION_PATTERNS = [
    re.compile(r'(?:university|univ\.|대학교|대학|institute|연구원|연구소|corporation|corp\.|inc\.|주식회사|㈜)\s*[\w\s]*', re.IGNORECASE),
]


def _extract_institution(text: str) -> str:
    """텍스트에서 기관명을 추출합니다."""
    for pat in _INSTITUTION_PATTERNS:
        m = pat.search(text[:5000])
        if m:
            inst = m.group(0).strip()
            if len(inst) > 5:
                return inst[:100]
    return ""


# ============================================================
# PDF 메타데이터 추출 (메인 함수)
# ============================================================
def extract_pdf_metadata(pdf_path: str, max_pages: int = 20) -> DocMeta:
    """
    PDF 파일에서 메타데이터를 추출합니다.

    Args:
        pdf_path: PDF 파일 경로
        max_pages: 텍스트 추출할 최대 페이지 수

    Returns:
        DocMeta: 추출된 메타데이터
    """
    pdf_path = str(pdf_path)
    filename = os.path.basename(pdf_path)

    with fitz.open(pdf_path) as doc:
        # 전체 텍스트 추출
        pages = min(max_pages, doc.page_count)
        chunks = []
        for i in range(pages):
            try:
                chunks.append(doc.load_page(i).get_text("text"))
            except Exception:
                continue
        full_text = "\n".join(chunks)
        full_text = re.sub(r'\n{3,}', '\n\n', full_text).strip()

        # PyMuPDF 메타데이터
        pdf_meta = doc.metadata or {}

    # DOI 추출 (텍스트 + 파일명)
    doi = _extract_doi(full_text) or _extract_doi(filename)

    # 특허번호 추출
    patent_no = _extract_patent_no(full_text) or _extract_patent_no(filename)

    # 문서 유형 판별 (특허번호, DOI 유무 우선 판단)
    if patent_no:
        doc_type = "patent"
    elif doi:
        doc_type = "paper"
    else:
        doc_type = _detect_doc_type(full_text, filename)

    # 제목 추출
    title = (pdf_meta.get("title", "").strip() or
             _extract_title_from_filename(filename) or
             _extract_title_from_text(full_text, doc_type))

    # 연도 추출
    year = _extract_year(full_text)

    # 기관 추출
    institution = _extract_institution(full_text)

    return DocMeta(
        filename=filename,
        doc_type=doc_type,
        title=title,
        doi=doi,
        patent_no=patent_no,
        institution=institution,
        year=year,
        full_text=full_text,
        source_type="pdf",
        source_path=pdf_path,
    )


# ============================================================
# WIPS 특허 API 헬퍼
# ============================================================
_WIPS_BASE = "https://sd.wips.co.kr/wipslink"
_WIPS_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "X-Requested-With": "XMLHttpRequest",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
}


def _html_to_text(html_str: str) -> str:
    """HTML 태그를 제거하고 텍스트만 추출합니다."""
    # HTML 태그 제거
    text = re.sub(r'<[^>]+>', ' ', html_str)
    # 연속 공백/줄바꿈 정리
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def _wips_get_skey(url: str) -> str:
    """URL에서 skey 파라미터를 추출합니다."""
    m = re.search(r'skey=([\d]+)', url)
    return m.group(1) if m else ""


def _wips_fetch_main(session: "requests.Session", skey: str) -> tuple:
    """
    WIPS 메인 페이지를 가져와 제목과 기본 메타데이터를 추출합니다.
    Returns: (title, patent_no, applicant, year, summary_text)
    """
    url = f"{_WIPS_BASE}/api/dkrdshtm.wips"
    r = session.get(url, params={"skey": skey}, timeout=30)
    r.encoding = 'utf-8'
    soup = BeautifulSoup(r.text, 'html.parser')

    # 제목 추출 (파우치형 전지셀의 제조장치 및 제조방법 + 영문)
    title = ""
    title_tag = soup.find('title')
    if title_tag:
        raw = title_tag.text.strip()
        # '상세보기' 부분 제거
        title = re.sub(r'\s*상세보기.*$', '', raw).strip()

    # 한국어 제목 찾기 (본문 내)
    for selector in ['.doc_title', 'h3.tit', '.tit_doc']:
        el = soup.select_one(selector)
        if el and len(el.get_text(strip=True)) > 5:
            title = el.get_text(strip=True)
            break

    # 출원번호/특허번호
    text_all = soup.get_text()
    patent_no = _extract_patent_no(text_all)
    year = _extract_year(text_all)
    institution = _extract_institution(text_all)

    return title, patent_no, institution, year


def _wips_fetch_abstract(session: "requests.Session", skey: str) -> str:
    """
    문헌전체 탭 (요약 포함) 내용을 가져옵니다.
    """
    try:
        r = session.post(
            f"{_WIPS_BASE}/api/findDkrDocInfo.wips",
            data={"skey": skey, "tabGb": "AB"},
            timeout=30,
        )
        r.encoding = 'utf-8'
        return _html_to_text(r.text)
    except Exception as e:
        print(f"[WARN] WIPS abstract fetch failed: {e}")
        return ""


def _wips_fetch_claims(session: "requests.Session", skey: str) -> str:
    """
    청구항 탭 (clList) 내용을 가져옵니다.
    """
    try:
        r = session.post(
            f"{_WIPS_BASE}/doc/docContJson.wips",
            data={"skey": skey, "tabGb": "CL"},
            timeout=30,
        )
        r.encoding = 'utf-8'
        j = r.json()
        cl_list = j.get("clList", [])
        parts = []
        for item in cl_list:
            cl_num = item.get("clNum", "")
            cl_text = _html_to_text(item.get("cl", ""))
            if cl_text:
                parts.append(f"[청구항 {cl_num}] {cl_text}")
        return "\n\n".join(parts)
    except Exception as e:
        print(f"[WARN] WIPS claims fetch failed: {e}")
        return ""


def _wips_fetch_description(session: "requests.Session", skey: str) -> str:
    """
    발명의 설명 탭 (descList) 내용을 가져옵니다.
    """
    try:
        r = session.post(
            f"{_WIPS_BASE}/doc/docContJson.wips",
            data={"skey": skey, "tabGb": "DS"},
            timeout=60,
        )
        r.encoding = 'utf-8'
        j = r.json()
        desc_list = j.get("descList", [])
        parts = []
        for item in desc_list:
            html = item.get("dtlDesc", "")
            if html:
                parts.append(_html_to_text(html))
        return "\n\n".join(parts)
    except Exception as e:
        print(f"[WARN] WIPS description fetch failed: {e}")
        return ""


def _wips_fetch_drawing_urls(session: "requests.Session", skey: str) -> List[str]:
    """
    도면 이미지 URL 목록을 가져옵니다.
    Returns: List of image URLs (e.g. img4.wipson.com/...)
    """
    try:
        r = session.post(
            f"{_WIPS_BASE}/dkr/findDrwImage.wips",
            data={"skey": skey},
            timeout=20,
        )
        r.encoding = 'utf-8'
        j = r.json()
        drw_list = j.get("drwImgList", [])
        urls = []
        for item in drw_list:
            u = item.get("drwUrl", "")
            if u:
                urls.append(u)
        return urls
    except Exception as e:
        print(f"[WARN] WIPS drawing fetch failed: {e}")
        return []


# ============================================================
# 특허 URL 메타데이터 추출
# ============================================================
def extract_patent_from_url(url: str) -> DocMeta:
    """
    특허 URL(WIPS 등)에서 메타데이터를 추출합니다.
    WIPS의 경우 문헌전체(요약), 청구항, 발명의 설명, 도면 탭을 모두 수집합니다.

    Args:
        url: 특허 상세 페이지 URL
            예: https://sd.wips.co.kr/wipslink/api/dkrdshtm.wips?skey=XXXXX

    Returns:
        DocMeta: 추출된 메타데이터 (figure_urls에 도면 이미지 URL 포함)
    """
    if not HAS_REQUESTS:
        raise RuntimeError("requests/beautifulsoup4 필요: pip install requests beautifulsoup4")

    skey = _wips_get_skey(url)
    stem = skey if skey else "patent"

    # WIPS 판별: skey 파라미터 또는 wips 도메인
    is_wips = bool(skey) and ("wips" in url.lower() or "wipson" in url.lower())

    if is_wips:
        return _extract_wips_patent(url, skey, stem)
    else:
        return _extract_generic_patent(url, stem)


def _extract_wips_patent(url: str, skey: str, stem: str) -> DocMeta:
    """WIPS 특허 페이지에서 모든 탭 내용을 수집합니다."""
    session = requests.Session()
    session.headers.update(_WIPS_HEADERS)
    session.headers["Referer"] = f"{_WIPS_BASE}/api/dkrdshtm.wips?skey={skey}"

    print(f"[INFO] WIPS 특허 크롤링 시작: skey={skey}")

    # 1. 메인 페이지 (제목, 특허번호, 기관, 연도)
    title, patent_no, institution, year = _wips_fetch_main(session, skey)
    print(f"[INFO] 제목: {title[:50] if title else 'N/A'}")

    # 2. 문헌전체 (요약/초록)
    abstract_text = _wips_fetch_abstract(session, skey)
    print(f"[INFO] 문헌전체 텍스트: {len(abstract_text)}자")

    # 3. 청구항
    claims_text = _wips_fetch_claims(session, skey)
    print(f"[INFO] 청구항 텍스트: {len(claims_text)}자")

    # 4. 발명의 설명
    desc_text = _wips_fetch_description(session, skey)
    print(f"[INFO] 발명의 설명 텍스트: {len(desc_text)}자")

    # 5. 도면 URL 목록
    figure_urls = _wips_fetch_drawing_urls(session, skey)
    print(f"[INFO] 도면 이미지: {len(figure_urls)}개")

    # 전체 텍스트 조합
    sections = []
    if title:
        sections.append(f"[제목]\n{title}")
    if abstract_text:
        sections.append(f"[문헌전체 / 요약]\n{abstract_text}")
    if claims_text:
        sections.append(f"[청구항]\n{claims_text}")
    if desc_text:
        sections.append(f"[발명의 설명]\n{desc_text}")

    full_text = "\n\n" + ("=" * 60) + "\n\n".join(sections)
    full_text = re.sub(r'\n{4,}', '\n\n\n', full_text).strip()

    if not full_text or len(full_text) < 100:
        raise RuntimeError(f"WIPS에서 충분한 텍스트를 추출하지 못함 (skey={skey})")

    # 제목이 없으면 텍스트에서 추출 시도
    if not title:
        title = _extract_title_from_text(full_text, "patent")
    # 특허번호 보완
    if not patent_no:
        patent_no = _extract_patent_no(full_text)
    # 연도 보완
    if not year:
        year = _extract_year(full_text)

    return DocMeta(
        filename=f"patent_{stem}",
        doc_type="patent",
        title=title,
        doi="",
        patent_no=patent_no,
        institution=institution,
        year=year,
        full_text=full_text,
        source_type="url",
        source_path=url,
        figure_urls=figure_urls,
    )


def _extract_generic_patent(url: str, stem: str) -> DocMeta:
    """WIPS 이외 특허 URL에서 일반 HTML 스크레이핑으로 메타데이터를 추출합니다."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
    }

    resp = requests.get(url, headers=headers, timeout=30)
    resp.encoding = 'utf-8'
    soup = BeautifulSoup(resp.text, 'html.parser')

    parts = []
    for tag in soup.find_all(['div', 'td', 'span', 'p']):
        text = tag.get_text(strip=True)
        if len(text) > 30:
            parts.append(text)

    # 중복 제거
    seen = set()
    unique_parts = []
    for p in parts:
        key = p[:200]
        if key not in seen:
            seen.add(key)
            unique_parts.append(p)

    full_text = "\n\n".join(unique_parts)
    full_text = re.sub(r'\n{3,}', '\n\n', full_text).strip()

    if len(full_text) < 100:
        raise RuntimeError(f"URL에서 충분한 텍스트를 추출하지 못함 (len={len(full_text)})")

    patent_no = _extract_patent_no(full_text)
    title = _extract_title_from_text(full_text, "patent")
    year = _extract_year(full_text)
    institution = _extract_institution(full_text)

    return DocMeta(
        filename=f"patent_{stem}",
        doc_type="patent",
        title=title,
        doi="",
        patent_no=patent_no,
        institution=institution,
        year=year,
        full_text=full_text,
        source_type="url",
        source_path=url,
    )


# ============================================================
# 폴더 스캔
# ============================================================
def scan_input_folder(folder_path: str, max_pages: int = 20) -> list:
    """
    입력 폴더의 모든 PDF 파일에서 메타데이터를 추출합니다.

    Args:
        folder_path: PDF 파일이 있는 폴더 경로
        max_pages: PDF당 최대 페이지 수

    Returns:
        list[DocMeta]: 추출된 메타데이터 리스트
    """
    folder = Path(folder_path)
    if not folder.is_dir():
        raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {folder_path}")

    results = []
    pdfs = sorted(folder.glob("*.pdf"), key=lambda p: p.name.lower())

    for pdf_path in pdfs:
        try:
            meta = extract_pdf_metadata(str(pdf_path), max_pages=max_pages)
            results.append(meta)
        except Exception as e:
            print(f"[WARN] {pdf_path.name} 메타데이터 추출 실패: {e}")
            results.append(DocMeta(
                filename=pdf_path.name,
                source_type="pdf",
                source_path=str(pdf_path),
            ))

    return results
