"""
Wikipedia 자동 검색 & 요약 스크립트
사용법: python wiki_search.py <검색어>
예시:  python wiki_search.py 인공지능
"""

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import re
from playwright.sync_api import sync_playwright


def clean_text(text: str) -> str:
    """불필요한 공백·각주 번호 제거"""
    text = re.sub(r'\[\d+\]', '', text)          # [1], [2] 각주 제거
    text = re.sub(r'\s+', ' ', text)             # 연속 공백 정리
    return text.strip()


def summarize(paragraphs: list[str], max_chars: int = 1500) -> str:
    """단락 리스트를 max_chars 이내로 합쳐 요약문 반환"""
    result = []
    total = 0
    for p in paragraphs:
        if total + len(p) > max_chars:
            break
        result.append(p)
        total += len(p)
    return '\n\n'.join(result)


def wiki_search(query: str) -> dict:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # ── 1. Wikipedia 접속 ──
        print(f"[1/4] Wikipedia 접속 중...")
        page.goto("https://www.wikipedia.org")
        page.wait_for_load_state("domcontentloaded")

        # ── 2. 검색어 입력 및 검색 ──
        print(f"[2/4] '{query}' 검색 중...")
        search_box = page.get_by_role("searchbox", name="Search Wikipedia")
        search_box.fill(query)
        search_box.press("Enter")
        page.wait_for_load_state("domcontentloaded")

        result_url = page.url
        result_title = page.title()

        # ── 3. 검색 결과 페이지 처리 ──
        # 직접 문서로 이동한 경우 vs 검색 목록이 나온 경우
        if "/wiki/" in result_url:
            print(f"[3/4] 문서 직접 이동됨: {result_title}")
        else:
            # 검색 결과 목록에서 첫 번째 링크 클릭
            print("[3/4] 검색 결과 목록에서 첫 번째 항목 클릭...")
            first_result = page.locator(".mw-search-result-heading a").first
            first_result.wait_for()
            result_title = first_result.inner_text()
            first_result.click()
            page.wait_for_load_state("domcontentloaded")
            result_url = page.url

        # ── 4. 본문 내용 추출 ──
        print(f"[4/4] 본문 추출 중...")

        # 제목
        title = page.locator("#firstHeading").inner_text()

        # 섹션 목차
        headings = page.locator("#mw-content-text h2, #mw-content-text h3").all()
        sections = [clean_text(h.inner_text().replace('[편집]', '')) for h in headings]
        sections = [s for s in sections if s and s not in ('목차', '참고 문헌', '외부 링크', '각주')]

        # 본문 단락 (상위 10개, 30자 이상)
        paragraphs = page.locator("#mw-content-text p").all()
        raw_paras = []
        for para in paragraphs[:20]:
            text = clean_text(para.inner_text())
            if len(text) > 30:
                raw_paras.append(text)

        summary = summarize(raw_paras)

        browser.close()

        return {
            "query":    query,
            "title":    title,
            "url":      result_url,
            "sections": sections[:10],
            "summary":  summary,
        }


def print_result(result: dict):
    divider = "=" * 60
    print(f"\n{divider}")
    print(f"  검색어  : {result['query']}")
    print(f"  제목    : {result['title']}")
    print(f"  URL     : {result['url']}")
    print(divider)

    print("\n[주요 섹션 목차]")
    for i, s in enumerate(result['sections'], 1):
        print(f"  {i:2}. {s}")

    print(f"\n[본문 요약]\n")
    print(result['summary'])
    print(f"\n{divider}\n")


if __name__ == "__main__":
    query = sys.argv[1] if len(sys.argv) > 1 else "인공지능"
    result = wiki_search(query)
    print_result(result)
