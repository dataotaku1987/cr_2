# -*- coding: utf-8 -*-
"""
Streamlit: Google Scholar 자동화 (UC 사용)
- 검색어 입력, 페이지 수/스크롤 횟수 지정
- <h3 class="gs_rt"> 제목 수집
- 인용 부모 div: <div class="gs_fl gs_flb"> 우선 탐색 → 보조 탐색 포함
- 원하는 페이지 수만큼 순차 이동(start=0,10,20...)
- 인용수 내림차순 정렬 후 미리보기 + 엑셀 다운로드
- UC Windows 종료 에러(WinError 6) 방지를 위한 소멸자 무력화 + atexit 안전 종료 포함
"""

import re
import io
import time
import atexit
import weakref
import urllib.parse
from datetime import datetime

import streamlit as st
import pandas as pd
import undetected_chromedriver as uc

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

# ---------------------------------------------------------
# 기본 환경 설정 (필요 시 사이드바에서 변경)
# ---------------------------------------------------------
DEFAULT_CHROME_MAJOR = 140     # 설치된 크롬 메이저 버전(예: 140.x → 140)
DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/139.0.0.0 Safari/537.36"
)
URL_TEMPLATE = "https://scholar.google.co.kr/scholar?hl=ko&as_sdt=0%2C5&q={query}&btnG="

# UC 소멸자 무력화 (Windows 종료 시 WinError 6 방지)
try:
    uc.Chrome.__del__ = lambda self: None
except Exception:
    pass


def wait(sec: float = 1.2):
    time.sleep(sec)


def build_driver(chrome_major: int, user_agent: str, headless: bool):
    """
    undetected-chromedriver 기반 드라이버 생성.
    - headless 옵션 선택 가능
    - User-Agent 적용
    - 한국어 선호, 자동화 흔적 일부 감소 옵션
    - atexit로 안전 종료 훅 등록
    """
    chrome_args = [
        "--lang=ko-KR",
        f"--user-agent={user_agent}",
        "--disable-blink-features=AutomationControlled",
    ]
    if not headless:
        chrome_args.append("--start-maximized")

    driver = uc.Chrome(
        version_main=chrome_major,
        headless=headless,
        use_subprocess=True,
        arguments=chrome_args,
    )

    # 안전 종료 훅: 프로세스가 남지 않도록 종료 시도
    dref = weakref.ref(driver)

    def _safe_quit():
        d = dref()
        if d is None:
            return
        try:
            d.quit()
        except Exception:
            pass

    atexit.register(_safe_quit)
    return driver


def wait_if_captcha(driver, prompt_on_streamlit: bool = True):
    """
    Scholar의 로봇/보안 확인 화면 감지 시 사용자 수동 해결 대기.
    (자동 우회 코드는 제공하지 않습니다.)
    """
    lower = driver.page_source.lower()
    hints = ["captcha", "로봇이", "자동화", "보안", "verify you are a human", "i'm not a robot", "문자를 입력"]
    if any(h in lower for h in hints):
        if prompt_on_streamlit:
            st.warning("보안 확인(CAPTCHA) 화면이 감지되었습니다. **브라우저에서 확인을 완료**하신 다음, 아래 버튼을 눌러 진행해주세요.")
            st.stop()  # 사용자가 확인 후 다시 '수집 재개' 버튼을 눌러 rerun하도록 유도
        return True
    return False


def open_with_query(driver, query: str):
    """
    검색어를 URL에 넣어 접속 + 검색창에 재입력(엔터)로 안정성 보완
    """
    encoded = urllib.parse.quote_plus(query)
    url = URL_TEMPLATE.format(query=encoded)
    driver.get(url)

    wait_if_captcha(driver)

    WebDriverWait(driver, 15).until(
        EC.any_of(
            EC.presence_of_element_located((By.NAME, "q")),
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.gs_ri")),
        )
    )

    # 검색창 재입력 → 엔터
    try:
        box = driver.find_element(By.NAME, "q")
        box.clear()
        box.send_keys(query)
        box.send_keys(Keys.ENTER)
        wait(1.0)
        wait_if_captcha(driver)
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.gs_ri"))
        )
    except Exception:
        pass


def scroll_page(driver, times: int = 2, pause: float = 1.2):
    for _ in range(times):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        wait(pause)


def go_to_page(driver, query: str, page_index: int):
    """
    N번째 페이지로 직접 이동 (1페이지 기준 → start=0, 10, 20, ...)
    """
    start = (page_index - 1) * 10
    encoded = urllib.parse.quote_plus(query)
    url = f"https://scholar.google.co.kr/scholar?start={start}&q={encoded}&hl=ko&as_sdt=0,5"
    driver.get(url)
    wait_if_captcha(driver)
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.gs_ri"))
    )
    wait(1.0)


def parse_citations_text(text: str) -> int:
    """
    '13회 인용' 혹은 'Cited by 13' 같은 문자열에서 숫자만 추출하여 정수 반환.
    """
    if not text:
        return 0
    m = re.search(r"(\d+)", text)
    return int(m.group(1)) if m else 0


def collect_page_items(driver, page_idx: int) -> list[dict]:
    """
    현재 페이지에서 제목 + 인용수 수집 (부모 div.gs_fl.gs_flb 우선)
    - title: 논문 제목
    - citations: 인용 수(정수)
    - page: 페이지 번호
    - rank: 페이지 내 순번(1부터)
    """
    items = []
    cards = driver.find_elements(By.CSS_SELECTOR, "div.gs_ri")
    rank = 0

    for card in cards:
        # 1) 제목
        title = ""
        try:
            h3 = card.find_element(By.CSS_SELECTOR, "h3.gs_rt")
            try:
                a = h3.find_element(By.CSS_SELECTOR, "a")
                title = a.text.strip()
            except NoSuchElementException:
                title = h3.text.strip()
        except NoSuchElementException:
            title = ""

        # 2) 인용수
        citations = 0
        try:
            # (a) 우선 부모: div.gs_fl.gs_flb
            try:
                fl_container = card.find_element(By.CSS_SELECTOR, "div.gs_fl.gs_flb")
            except NoSuchElementException:
                fl_container = None

            cite_link = None
            if fl_container:
                try:
                    cite_link = fl_container.find_element(By.CSS_SELECTOR, "a[href*='cites=']")
                except NoSuchElementException:
                    cite_link = None
                if not cite_link:
                    try:
                        cite_link = fl_container.find_element(
                            By.XPATH, ".//a[contains(., '인용') or contains(., 'Cited by')]"
                        )
                    except NoSuchElementException:
                        cite_link = None

            # (b) 보조: 부모가 없거나 실패하면 결과 블록(ancestor)에서 재탐색
            if not cite_link:
                try:
                    parent = card.find_element(
                        By.XPATH,
                        "./ancestor-or-self::div[contains(@class,'gs_r') or contains(@class,'gs_scl')][1]"
                    )
                    try:
                        cite_link = parent.find_element(By.CSS_SELECTOR, "a[href*='cites=']")
                    except NoSuchElementException:
                        cite_link = None
                    if not cite_link:
                        try:
                            cite_link = parent.find_element(
                                By.XPATH, ".//a[contains(., '인용') or contains(., 'Cited by')]"
                            )
                        except NoSuchElementException:
                            cite_link = None
                except Exception:
                    cite_link = None

            if cite_link:
                citations = parse_citations_text(cite_link.text.strip())
        except Exception:
            citations = 0

        if title:
            rank += 1
            items.append({
                "page": page_idx,
                "rank": rank,
                "title": title,
                "citations": int(citations),
            })

    return items


def run_scrape(query: str, total_pages: int, scroll_count: int,
               chrome_major: int, user_agent: str, headless: bool):
    """
    전체 수집 파이프라인 실행 → DataFrame 반환
    """
    driver = build_driver(chrome_major, user_agent, headless)
    results: list[dict] = []
    try:
        # 첫 페이지
        open_with_query(driver, query)
        scroll_page(driver, times=scroll_count)
        results.extend(collect_page_items(driver, page_idx=1))

        # 이후 페이지
        for p in range(2, total_pages + 1):
            go_to_page(driver, query, p)
            scroll_page(driver, times=scroll_count)
            results.extend(collect_page_items(driver, page_idx=p))

        df = pd.DataFrame(results)
        if df.empty:
            return df

        df["citations"] = pd.to_numeric(df["citations"], errors="coerce").fillna(0).astype(int)
        df_sorted = df.sort_values(["citations", "page", "rank"], ascending=[False, True, True]).reset_index(drop=True)
        return df_sorted

    finally:
        try:
            driver.quit()
        except Exception:
            pass


def df_to_excel_bytes(df: pd.DataFrame, query: str) -> bytes:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_query = re.sub(r"[\\/:*?\"<>|]", "_", query)
    file_name = f"scholar_results_{safe_query}_{ts}.xlsx"

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="results")
    bio.seek(0)
    return file_name, bio.read()


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="Google Scholar 자동화 수집기", layout="wide")
st.title("Google Scholar 자동화 수집기 (UC + Selenium)")

with st.sidebar:
    st.subheader("설정")
    query = st.text_input("검색어", value="AI 에이전트", help="예: LLM, 멀티모달, AI 에이전트 등")
    total_pages = st.number_input("총 페이지 수", min_value=1, max_value=50, value=3, step=1)
    scroll_count = st.number_input("페이지당 스크롤 횟수", min_value=0, max_value=10, value=2, step=1)

    st.markdown("---")
    chrome_major = st.number_input("Chrome 메이저 버전", min_value=80, max_value=200,
                                   value=DEFAULT_CHROME_MAJOR, step=1,
                                   help="설치된 Chrome 버전의 메이저 숫자. 예: 140.x → 140")
    user_agent = st.text_input("User-Agent", value=DEFAULT_USER_AGENT)
    headless = st.checkbox("헤드리스(브라우저 미표시) 모드", value=False,
                           help="체크 시 브라우저 창을 띄우지 않습니다. CAPTCHA 처리에 불리할 수 있습니다.")

    st.markdown("---")
    st.caption("⚠️ CAPTCHA(보안 확인)가 나타나면 브라우저에서 직접 확인을 완료하세요. 자동 우회는 제공되지 않습니다.")

col1, col2 = st.columns([1, 2])

with col1:
    run_btn = st.button("수집 시작", type="primary")
    st.write("")

with col2:
    placeholder = st.empty()

if run_btn:
    if not query.strip():
        st.error("검색어를 입력해주세요.")
        st.stop()

    progress = st.progress(0, text="초기화 중...")
    try:
        progress.progress(5, text="드라이버 준비...")
        # 실제 수집
        df_sorted = run_scrape(
            query=query,
            total_pages=int(total_pages),
            scroll_count=int(scroll_count),
            chrome_major=int(chrome_major),
            user_agent=user_agent.strip(),
            headless=bool(headless),
        )
        progress.progress(90, text="정렬 및 정리 중...")

        if df_sorted is None or df_sorted.empty:
            progress.progress(100, text="완료")
            st.warning("수집된 결과가 없습니다.")
        else:
            progress.progress(100, text="완료")
            st.success(f"총 {len(df_sorted)}건 수집 및 정렬 완료")

            st.subheader("미리보기")
            st.dataframe(df_sorted, use_container_width=True, height=480)

            # 다운로드
            fname, xbytes = df_to_excel_bytes(df_sorted, query)
            st.download_button(
                label="엑셀 다운로드",
                data=xbytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.info("참고: 정렬 기준 = 인용수 내림차순 → 동일 시 페이지/랭크 오름차순")

    except Exception as e:
        st.error(f"에러가 발생했습니다: {e}")
    finally:
        progress.empty()

# 도움말
with st.expander("도움말 / 트러블슈팅"):
    st.markdown(
        """
- **브라우저 표시 권장:** CAPTCHA가 자주 발생하므로 `헤드리스 모드`를 끄고 진행하는 것을 권장합니다.
- **Chrome 메이저 버전이 맞지 않을 때:** 사이드바에서 **Chrome 메이저 버전** 숫자를 설치된 크롬과 맞춰주세요. (예: 140.x → 140)
- **CAPTCHA 발생 시:** 브라우저에서 보안 확인을 완료한 뒤, 사이드바 설정을 변경하지 말고 다시 **수집 시작**을 눌러 재시도하세요.
- **결과가 비어있는 경우:**
  - 검색어를 더 구체화하거나,
  - 페이지 수/스크롤 횟수를 늘려보세요.
- **권장 실행 환경:** 로컬 PC(Windows)에서 실행. Streamlit Cloud 등 원격/컨테이너 환경에서는 브라우저 표시가 제한될 수 있습니다.
        """
    )
