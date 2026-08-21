#!/usr/bin/env python3
"""
소장도서 xlsx -> data/books.json 변환 스크립트

사용법:
  1. 도서관리 프로그램에서 소장도서 목록을 xlsx로 내려받는다.
     (열 순서: No, 등록번호, 자료명, 저자, 출판사, 출판년도, 청구기호, 자료상태, 소장처, 가격, ISBN)
  2. 그 파일을 data/collection_source.xlsx 로 덮어쓴다.
  3. 터미널에서 실행: python3 scripts/build_books.py
  4. data/books.json 이 새로 생성되면, 그 파일과 index.html 을 함께 커밋/배포한다.

코드 수정 없이 소장도서 목록만 반복 갱신할 수 있도록 만든 스크립트입니다.
"""
import json
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl이 필요합니다. 먼저 'pip install openpyxl'을 실행하세요.")

ROOT = Path(__file__).resolve().parent.parent
SOURCE_XLSX = ROOT / "data" / "collection_source.xlsx"
OUTPUT_JSON = ROOT / "data" / "books.json"

# 대출 불가한(더 이상 서가에 없는) 자료상태는 책장 구경하기/검색에서 제외한다.
EXCLUDED_STATUSES = {"분실", "파손", "가치상실"}

# 청구기호 앞자리(KDC 대분류 0~9) -> 화면에 보여줄 주제 카테고리
KDC_CATEGORY = {
    0: "백과사전",
    1: "철학",
    2: "종교",
    3: "사회",
    4: "자연과학",
    5: "기술",
    6: "예술체육",
    7: "언어",
    8: "문학",
    9: "역사",
}

ISBN10_RE = re.compile(r"^\d{9}[\dXx]$")
ISBN13_RE = re.compile(r"^\d{13}$")


def classify(call_number: str, location: str, title: str) -> str:
    """청구기호(KDC) + 소장처 + 제목 키워드로 13개 주제 중 하나를 정한다."""
    # 1) 그림책서가에 있으면 KDC와 무관하게 '그림책'으로 분류한다.
    if location == "그림책서가":
        return "그림책"

    m = re.search(r"(\d+(?:\.\d+)?)", call_number or "")
    if not m:
        return "기타"
    kdc_main = int(float(m.group(1)) // 100)

    # 2) 제목에 '도감' 또는 '사전'이 있으면 총류(백과사전)를 제외한 나머지 분류에서
    #    '도감·사전'으로 우선 분류한다. (동물도감/식물도감/각종 사전류가
    #    자연과학·언어 등 여러 KDC에 흩어져 있어 탐색 화면에서 따로 모아준다.)
    if kdc_main != 0 and title and ("도감" in title or "사전" in title):
        return "도감·사전"

    return KDC_CATEGORY.get(kdc_main, "기타")


def normalize_isbn(raw) -> str:
    if raw is None:
        return ""
    s = str(raw).strip().replace("-", "")
    if ISBN10_RE.match(s) or ISBN13_RE.match(s):
        return s
    return ""


def main():
    if not SOURCE_XLSX.exists():
        sys.exit(f"원본 파일을 찾을 수 없습니다: {SOURCE_XLSX}")

    wb = openpyxl.load_workbook(SOURCE_XLSX, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    next(rows)  # 헤더 건너뛰기

    books = []
    for row in rows:
        if row is None or len(row) < 11:
            continue
        (_no, reg_no, title, author, publisher, year,
         call_number, status, location, _price, isbn) = row[:11]

        if not title or not reg_no:
            continue
        if status in EXCLUDED_STATUSES:
            continue

        title = str(title).strip()
        call_number = str(call_number or "").strip()
        location = str(location or "").strip()

        books.append({
            "id": str(reg_no).strip(),
            "title": title,
            "author": str(author or "").strip(),
            "publisher": str(publisher or "").strip(),
            "year": str(year or "").strip(),
            "callNumber": call_number,
            "status": str(status or "").strip(),
            "location": location,
            "isbn": normalize_isbn(isbn),
            "category": classify(call_number, location, title),
        })

    OUTPUT_JSON.write_text(
        json.dumps(books, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )

    from collections import Counter
    counts = Counter(b["category"] for b in books)
    print(f"총 {len(books)}권 -> {OUTPUT_JSON}")
    for cat, n in sorted(counts.items(), key=lambda x: -x[1]):
        print(f"  {cat}: {n}권")


if __name__ == "__main__":
    main()
