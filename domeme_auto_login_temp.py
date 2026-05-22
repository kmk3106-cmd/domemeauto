# -*- coding: utf-8 -*-
"""
[1] 도매매 상품검색후 마이박스 추출
- 총작업 순서: 1.도매매 상품검색·마이박스 → 2.Run_all_steps(_최종.xlsx 생성) → 3.스피드고 마이박스 엑셀업로드 및 전송
- 1~6번사업자 루프: 도매매 로그인 → 키워드 검색 → 마이박스담기 → 스피드고 엑셀 다운로드 → run_all_steps → _최종.xlsx 스피드고 업로드·전송
- 주차별 키워드 분리로 겹침 방지
- 일반 브라우저(launch_persistent_context, 시크릿 아님) + domeme_browser_profile
- 카테고리 다양화, 제외: 화장품/식품/영유아/브랜드/명품/골프/낚시/서적·DVD·CD음반(정가정책)

사용법:
  FAST_MODE=1: 대기시간 축소. 느리면 FAST_MODE=0으로 안정 모드.
  CHROME_PROCESS_DIAG=1: 시작 시 Chrome 프로세스 cmdline 진단 출력(기본 생략·시작 빠름).
  Phase 1 Chrome·도매매 첫 접속은 **Phase 2 와 동일**하게 동작한다: Profile 67 전용 복사본 user_data_dir,
  동일 `launch_persistent_context` 인자, `new_page` 후 `goto(..., commit)` → `domcontentloaded` 대기.
  (끄기: `PHASE1_LAUNCH_LIKE_PHASE2=0` — 예전 전체 User Data 복사본·별도 launch 인자)
  로그인도 **test_speedgo_upload_1번** 동선을 먼저 시도하고, 실패 시 mem_formLogin 폴백.
  Chrome 창이 뒤에 있으면 로딩이 느려질 수 있음 → bring_to_front 로 전면 표시한다.
  1. pip install playwright
  2. playwright install chromium
  3. 아래 ACCOUNTS, PASSWORD 수정 후 실행
     python domeme_auto_login_temp.py

  회차·사업자 자동 선택(동일 주차 재실행 시):
    마이박스 저장 경로 …/{yy년M월w주차}/{N}회차/{M}번사업자 폴더가 없는 가장 앞 회차(N)와,
    그 회차에서 폴더가 없는 사업자(M)만 실행한다. 1회차 6명 모두 있으면 2회차로 넘어간다.
    WEEK_RUN=N 을 주면 N회차만 검사해 빈 사업자만 돌린다. .week_run_state 는 이번 실행 회차와 동기화된다.

  프로필:
    - 기본: 자동화 전용 프로필(domeme_browser_profile)
    - 실제 Chrome 프로필: USE_REAL_CHROME_PROFILE=1 (Chrome 종료 필요)
    - Chrome 실행 중: USE_REAL_CHROME_PROFILE=1 USE_REAL_CHROME_PROFILE_COPY=1 (프로필 복사)
    - Phase 1·2: `_chrome_persistent_launch_kw_phase2_identical` 로 launch 인자 통일, `--profile-directory` 는
      `PHASE2_CHROME_PROFILE_DIR`(기본 Profile 67). Phase1 은 기본적으로 Phase2 와 같은 Profile67 복사 경로 사용.
      예전 Phase1 전용 launch: `PHASE1_LAUNCH_LIKE_PHASE2=0`
    - 첫 실행만 프로필 오류 시: CHROME_PROFILE_COPY_FORCE=1 로 전체 재복사 후 재실행
    - 프로필 복사본 사용 시 Chrome 우측 프로필 느낌표는 동기화 불일치인 경우가 많음(도매매 동작과 무관할 수 있음)
    - 엑셀 다운로드 직전/직후 popup, navigation, response, dialog 로그는
      profile_compare_log.txt 에 추가됨 (PROFILE_COMPARE_LOG=0 으로 비활성화 가능)

보안: 이 파일에 계정정보가 있습니다. git 커밋하지 마세요.
"""

import base64
import gc
import json
import os
import re
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urlparse

# stdout/stderr 를 UTF-8·치환모드로 고정: 콘솔 코드페이지(cp949 등)에서
# 한글·치환문자(�) 출력 시 print() 가 UnicodeEncodeError 로 죽는 것을 방지.
for _stream_name in ("stdout", "stderr"):
    try:
        getattr(sys, _stream_name).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# .env 로드 (이 파일과 같은 폴더). python-dotenv 미설치 시 OS 환경변수만 사용.
try:
    from dotenv import load_dotenv

    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass

# ============== 계정 설정 (.env 에서 로드, 커밋 금지) ==============
# DOMEME_ACCOUNTS: 쉼표 구분 (순서 = 1~6번 사업자)
ACCOUNTS = [a.strip() for a in os.environ.get("DOMEME_ACCOUNTS", "").split(",") if a.strip()]
PASSWORD = os.environ.get("DOMEME_PASSWORD", "")  # 대소문자 구분

# Naver DataLab API (쇼핑인사이트)
NAVER_CLIENT_ID = os.environ.get("NAVER_CLIENT_ID", "")
NAVER_CLIENT_SECRET = os.environ.get("NAVER_CLIENT_SECRET", "")
# =========================================================

# 도매매 DB 메인 — work_page.goto 에 그대로 넘기는 고정 문자열(다른 상수와 동일해야 함)
_DOMEME_GOTO_URL = "https://domemedb.domeggook.com/index/"
DOMEME_URL = _DOMEME_GOTO_URL
# work_page 전용: 요구사항 timeout 30000ms
_DOMEME_WORK_PAGE_GOTO_TIMEOUT_MS = 30_000
# 실제 사이트 헤더「로그인」과 동일: domeme 호스트의 회원 로그인 폼 (?back = 로그인 후 복귀 URL base64)
DOMEME_MEMBER_LOGIN_BASE = "https://domeme.domeggook.com/main/member/mem_formLogin.php"
_DOMEME_LOGIN_BACK_DEFAULT = "https://domemedb.domeggook.com/index"
DOMEME_MEMBER_LOGIN_ENTRY = (
    f"{DOMEME_MEMBER_LOGIN_BASE}?back="
    + quote(base64.b64encode(_DOMEME_LOGIN_BACK_DEFAULT.encode("utf-8")).decode("ascii"))
)
# 구버전/보조: domemedb 의 ?login=pc (폼이 다를 수 있음)
DOMEME_LOGIN_URL = "https://domemedb.domeggook.com/index/?login=pc"
# 첫 탭이 blank에 묶일 때 쓰는 보조 URL
DOMEME_URL_ROOT = "https://domemedb.domeggook.com/"
SPEEDGO_URL = "https://speedgo.domeggook.com/"


def _domeme_member_login_url(back_href=None) -> str:
    """헤더 `<a href=\"...mem_formLogin.php?back=...\">로그인</a>` 과 동일 형식."""
    back = (back_href or _DOMEME_LOGIN_BACK_DEFAULT).strip()
    b64 = base64.b64encode(back.encode("utf-8")).decode("ascii")
    return f"{DOMEME_MEMBER_LOGIN_BASE}?back={quote(b64)}"


def _url_is_domeme_session(u: str) -> bool:
    """작업 가능한 도매매(domemedb)·회원로그인(domeme) URL. speedgo 는 제외(오탐 방지)."""
    t = (u or "").strip().lower()
    if not t or t == "about:blank" or t.startswith(("chrome", "devtools:", "edge:", "brave:")):
        return False
    if "speedgo.domeggook" in t:
        return False
    if "domemedb.domeggook" in t:
        return True
    if "domeme.domeggook.com" in t:
        return True
    if "domeggook.com" in t and any(
        k in t for k in ("mem_formlogin", "/member/", "login=pc", "domemedb", "domeme.")
    ):
        return True
    return False


def _domeme_url_is_chrome_or_devtools(url: str) -> bool:
    u = (url or "").strip().lower()
    return u.startswith("chrome://") or u.startswith("devtools:")


def _domeme_is_internal_chrome_tab_url(url: str) -> bool:
    """닫아도 되는 내부 UI 탭(chrome://, devtools, omnibox). about:blank 은 단독 탭일 수 있어 여기서 제외."""
    u = (url or "").strip().lower()
    return _domeme_url_is_chrome_or_devtools(u) or "omnibox" in u


def _domeme_log_browser_context_state(context, log_prefix: str, label: str, work_page=None) -> bool:
    """
    context / browser / pages 상태 로그. context 가 닫혔거나 pages 를 못 읽으면 False.
    """
    print(f"{log_prefix}[ctx상태:{label}]", flush=True)
    try:
        br = context.browser
        if br is not None:
            try:
                print(f"  browser.is_connected()={br.is_connected()}", flush=True)
            except Exception as e:
                print(f"  browser 연결 확인 실패: {e!r}", flush=True)
    except Exception as e:
        print(f"  context.browser 접근 실패: {e!r}", flush=True)
    try:
        ctx_closed = context.is_closed()
    except AttributeError:
        ctx_closed = False
    except Exception:
        ctx_closed = "?"
    print(f"  context.is_closed()={ctx_closed}", flush=True)
    if ctx_closed is True:
        print(f"{log_prefix}[ctx상태] context 가 이미 닫혀 있음 → new_page 불가", flush=True)
        return False
    try:
        plist = list(context.pages)
    except Exception as e:
        print(f"  context.pages 읽기 실패: {e!r}", flush=True)
        return False
    print(f"  pages 개수={len(plist)}", flush=True)
    for i, p in enumerate(plist):
        try:
            ic = p.is_closed()
        except Exception:
            ic = "?"
        try:
            u = repr(p.url)[:160]
        except Exception as e:
            u = str(e)
        print(f"  page[{i}] id={id(p)} is_closed={ic} url={u}", flush=True)
    if work_page is not None:
        try:
            wu = repr(work_page.url)[:160]
            wc = work_page.is_closed()
        except Exception as e:
            wu = str(e)
            wc = "?"
        print(f"  work_page id={id(work_page)} is_closed={wc} url={wu}", flush=True)
    return True


def _domeme_close_internal_pages(context, *, work_page=None, log_prefix: str = "") -> None:
    """
    chrome://·devtools·omnibox 만 정리. work_page 절대 닫지 않음.
    pages < 2 이면 스킵(마지막 탭 보호). 닫은 뒤 0페이지가 되면 close 금지.
    """
    try:
        plist = []
        for p in list(context.pages):
            try:
                if p.is_closed():
                    continue
            except Exception:
                continue
            plist.append(p)
    except Exception as e:
        print(f"{log_prefix}[탭정리] pages 수집 실패: {e!r}", flush=True)
        return

    if len(plist) < 2:
        print(
            f"{log_prefix}[탭정리] 스킵: 열린 탭 {len(plist)}개(<2) — 닫아서 0개 되는 것 방지",
            flush=True,
        )
        return

    work_id = id(work_page) if work_page is not None else None
    max_pass = 20
    for _ in range(max_pass):
        try:
            alive = [p for p in list(context.pages) if not p.is_closed()]
        except Exception as e:
            print(f"{log_prefix}[탭정리] pages 재조회 실패: {e!r}", flush=True)
            break
        if len(alive) < 2:
            print(f"{log_prefix}[탭정리] 종료: 열린 탭 {len(alive)}개(<2)", flush=True)
            break
        victim = None
        victim_url = ""
        for p in alive:
            if work_id is not None and id(p) == work_id:
                continue
            try:
                u = (p.url or "").lower()
            except Exception:
                continue
            if _domeme_is_internal_chrome_tab_url(u):
                victim = p
                victim_url = u
                break
        if victim is None:
            break
        if len(alive) <= 1:
            break
        if len(alive) - 1 < 1:
            print(f"{log_prefix}[탭정리] 스킵: 닫으면 탭 0개", flush=True)
            break
        try:
            victim.close()
            print(f"{log_prefix}[탭정리] 내부 UI 탭 닫음: {victim_url[:100]}", flush=True)
        except Exception as e:
            print(f"{log_prefix}[탭정리] close 실패: {e!r}", flush=True)
            break
    time.sleep(0.12)


def _domeme_discard_all_context_pages(context, log_prefix: str = "") -> None:
    """[사용 금지] work_page 생성 전에 호출하면 pages=0 이 되어 new_page 가 실패할 수 있음. 유지는 호환·수동 디버그용."""
    for p in list(context.pages):
        try:
            if p.is_closed():
                continue
        except Exception:
            continue
        try:
            u = repr(p.url)[:160]
        except Exception:
            u = "?"
        try:
            print(f"{log_prefix}[기존탭 일괄 닫음] {u}", flush=True)
            p.close()
        except Exception as e:
            print(f"{log_prefix}[탭 닫기 실패] {e!r}", flush=True)
    time.sleep(0.2)


def _domemedb_work_host_ok(url: str) -> bool:
    """작업 성공 URL: domemedb 호스트 포함(요구사항)."""
    return "domemedb.domeggook.com" in (url or "").lower()


def _domeme_focus_window_for_navigation(page) -> None:
    """전면 표시만 수행(주소창·키보드 단축키·pyautogui 미사용)."""
    try:
        page.bring_to_front()
    except Exception:
        pass
    time.sleep(0.15)


def _domeme_url_is_blankish(url: str) -> bool:
    u = (url or "").strip().lower()
    return u in ("", "about:blank")


def _domeme_debug_dump_pages(context, log_prefix: str = "") -> None:
    """(1)(2) 연결된 모든 Page와 각 URL을 출력 — 잘못된 page 바인딩 디버깅."""
    try:
        pages = list(context.pages)
    except Exception as e:
        print(f"{log_prefix}[페이지목록] context.pages 읽기 실패: {e!r}", flush=True)
        return
    print(f"{log_prefix}[페이지목록] 총 {len(pages)}개 Page", flush=True)
    for i, p in enumerate(pages):
        try:
            closed = p.is_closed()
        except Exception:
            closed = "?"
        try:
            u = (p.url or "").strip()
        except Exception as e:
            u = f"<url 읽기 실패: {e!r}>"
        blank = _domeme_url_is_blankish(u)
        note = "  ← new_tab 직후면 about:blank 은 정상일 수 있음(참고만)" if blank else ""
        print(
            f"{log_prefix}  page[{i}] is_closed={closed} url={repr(u)}{note}",
            flush=True,
        )


def _domemedb_goto_target_reached(url: str) -> bool:
    """도매매/도매매DB 진입 성공: URL에 domeggook.com(및 domemedb 하위) 포함."""
    return "domeggook.com" in (url or "").lower()


def _domeme_print_page_nav_diagnostics(page, err, log_prefix: str) -> None:
    """이동 실패 시 url·title·예외 출력."""
    try:
        u = page.url
    except Exception as e:
        u = f"<url 읽기 실패: {e!r}>"
    try:
        t = page.title()
    except Exception as e:
        t = f"<title 읽기 실패: {e!r}>"
    print(f"{log_prefix}[실패] url={repr(u)} title={repr(t)} error={err!r}", flush=True)


def _domeme_log_navigation_assert(ok: bool, url_now: str, log_prefix: str, where: str) -> None:
    """(9) navigation 성공 여부를 로그로 명시(자동화 중단 없이 검증)."""
    if ok:
        print(
            f"{log_prefix}[검증 OK] {where}: domeggook URL 확인 url={repr(url_now)[:200]}",
            flush=True,
        )
    else:
        print(
            f"{log_prefix}[검증 FAIL] {where}: domeggook/domemedb URL 미확인 "
            f"url={repr(url_now)[:200]}",
            flush=True,
        )


def _domeme_analyze_navigation_not_applied(work_page, log_prefix: str, goto_exc) -> None:
    """
    goto 호출 후에도 domemedb 미도착일 때만. about:blank 자체를 실패 원인으로 단정하지 않는다.
    """
    print(
        f"{log_prefix}[분석] new_page() 직후 about:blank 는 **정상일 수 있음**. "
        "검증 포인트는 `work_page.goto(...)` 이후에도 URL이 domemedb 로 바뀌는지 여부.",
        flush=True,
    )
    try:
        ic = work_page.is_closed()
    except Exception:
        ic = "?"
    print(f"{log_prefix}[분석] work_page.is_closed()={ic}", flush=True)
    try:
        href = work_page.evaluate("() => location.href")
        print(f"{log_prefix}[분석] JS location.href={repr(href)}", flush=True)
    except Exception as e:
        print(f"{log_prefix}[분석] JS location.href 읽기 실패: {e!r}", flush=True)
    try:
        sess = work_page.context.new_cdp_session(work_page)
        sess.send("Runtime.evaluate", {"expression": "document.readyState", "returnByValue": True})
        print(f"{log_prefix}[분석] CDP 세션 연결됨(문서 상태 조회 시도)", flush=True)
    except Exception as e:
        print(f"{log_prefix}[분석] CDP 세션 실패: {e!r}", flush=True)
    if goto_exc is not None:
        print(f"{log_prefix}[분석] work_page.goto 예외: {goto_exc!r}", flush=True)
    print(
        f"{log_prefix}[분석] 기타 점검: Chrome/프로필 잠금, 보안SW·SSL 가로채기, "
        "회사망 프록시, Playwright↔Chrome 버전 불일치 등.",
        flush=True,
    )


def _domeme_create_work_page_with_domemedb_goto(
    context,
    log_prefix: str,
    *,
    discard_entire_pages_list: bool,
):
    """
    순서: (1) work_page = context.new_page() 및 goto 성공 검증 → (2) 그 다음에만 내부 chrome 탭 정리.
    new_page 전에 탭을 닫아 pages=0 이 되는 구조는 금지. work_page 는 절대 close 하지 않음.

    discard_entire_pages_list: 예전 의미(일괄 닫기)는 사용하지 않음. True여도 일괄 닫기 호출 안 함.
    """
    tout = _DOMEME_WORK_PAGE_GOTO_TIMEOUT_MS
    if discard_entire_pages_list:
        print(
            f"{log_prefix}[정책] discard_entire_pages_list=True 이어도, "
            "work_page 생성 전 일괄 탭 닫기는 하지 않습니다.",
            flush=True,
        )

    if not _domeme_log_browser_context_state(context, log_prefix, "work_page 생성 전", None):
        raise RuntimeError(f"{log_prefix}BrowserContext 가 이미 닫혀 있어 new_page 를 할 수 없습니다.")

    try:
        work_page = context.new_page()
    except Exception as e:
        print(f"{log_prefix}[오류] context.new_page() 실패: {e!r}", flush=True)
        _domeme_log_browser_context_state(context, log_prefix, "new_page 예외 직후", None)
        raise

    try:
        if work_page.is_closed():
            print(f"{log_prefix}[오류] new_page 직후 work_page.is_closed()==True", flush=True)
            raise RuntimeError("work_page 가 즉시 닫힌 상태입니다.")
    except RuntimeError:
        raise
    except Exception as e:
        print(f"{log_prefix}[오류] work_page 상태 확인 실패: {e!r}", flush=True)
        raise

    _domeme_log_browser_context_state(context, log_prefix, "work_page 생성 직후", work_page)
    try:
        u_new = (work_page.url or "").strip()
    except Exception as e:
        u_new = f"<읽기 실패: {e!r}>"
    print(
        f"{log_prefix}[검증1] id(work_page)={id(work_page)} work_page.url={repr(u_new)}",
        flush=True,
    )

    try:
        work_page.bring_to_front()
    except Exception:
        pass
    _domeme_focus_window_for_navigation(work_page)
    try:
        work_page.set_default_navigation_timeout(tout)
        work_page.set_default_timeout(tout)
    except Exception:
        pass

    goto_exc = None
    try:
        work_page.goto(
            "https://domemedb.domeggook.com/index/",
            wait_until="domcontentloaded",
            timeout=tout,
        )
    except Exception as e:
        goto_exc = e

    try:
        u_after = (work_page.url or "").strip()
    except Exception as e:
        u_after = f"<읽기 실패: {e!r}>"
    print(
        f"{log_prefix}[검증2] id(work_page)={id(work_page)} work_page.url={repr(u_after)}",
        flush=True,
    )

    try:
        title_now = work_page.title()
    except Exception as e:
        title_now = f"<title 읽기 실패: {e!r}>"
    print(f"{log_prefix}[검증3] work_page.title()={repr(title_now)}", flush=True)

    time.sleep(3.0)
    try:
        u_3s = (work_page.url or "").strip()
    except Exception as e:
        u_3s = f"<읽기 실패: {e!r}>"
    print(f"{log_prefix}[검증4] 3초 후 work_page.url={repr(u_3s)}", flush=True)

    ok = _domemedb_work_host_ok(u_3s)
    _domeme_log_navigation_assert(ok, u_3s, log_prefix, "3초 후 URL 기준")
    if ok:
        _domeme_log_browser_context_state(context, log_prefix, "domemedb 성공 후·내부탭 정리 전", work_page)
        _domeme_close_internal_pages(context, work_page=work_page, log_prefix=log_prefix)
        _domeme_log_browser_context_state(context, log_prefix, "내부탭 정리 후", work_page)
        return work_page

    if _domeme_url_is_blankish(u_3s):
        _domeme_analyze_navigation_not_applied(work_page, log_prefix, goto_exc)
    else:
        print(
            f"{log_prefix}[분석] about:blank 아님·그러나 domemedb.domeggook.com 미포함: {repr(u_3s)[:220]}",
            flush=True,
        )
        _domeme_print_page_nav_diagnostics(work_page, goto_exc, log_prefix)

    print(f"{log_prefix}[복구] domemedb 미포함 → 동일 work_page 에서 work_page.goto 1회 재시도", flush=True)
    goto_exc2 = None
    try:
        work_page.goto(
            "https://domemedb.domeggook.com/index/",
            wait_until="domcontentloaded",
            timeout=tout,
        )
    except Exception as e2:
        goto_exc2 = e2
    time.sleep(1.0)
    try:
        u_retry = (work_page.url or "").strip()
    except Exception:
        u_retry = ""
    try:
        t_retry = work_page.title()
    except Exception as e:
        t_retry = f"<{e!r}>"
    print(
        f"{log_prefix}[재시도 후] work_page.url={repr(u_retry)} title={repr(t_retry)}",
        flush=True,
    )
    if _domemedb_work_host_ok(u_retry):
        _domeme_log_navigation_assert(True, u_retry, log_prefix, "재goto 후")
        _domeme_log_browser_context_state(context, log_prefix, "재goto 성공 후·내부탭 정리 전", work_page)
        _domeme_close_internal_pages(context, work_page=work_page, log_prefix=log_prefix)
        _domeme_log_browser_context_state(context, log_prefix, "내부탭 정리 후", work_page)
    elif _domeme_url_is_blankish(u_retry):
        _domeme_analyze_navigation_not_applied(work_page, log_prefix, goto_exc2 or goto_exc)
    return work_page


def _domeme_open_domemedb_like_phase2_new_page(context, *, log_prefix: str = "[시작] "):
    """
    Phase 2 직후 `new_page()` + `test_speedgo_upload_1번` 첫 이동과 동일:
    domemedb index → wait_until=commit → wait_for_load_state(domcontentloaded).
    (구버전의 domcontentloaded만·3초 대기·탭정리 등은 사용하지 않음 — 접속 실패 시 Phase2와 불일치 원인 제거)
    """
    if not _domeme_log_browser_context_state(context, log_prefix, "Phase2식 도매매 연결 전", None):
        raise RuntimeError(f"{log_prefix}BrowserContext 가 닫혀 있어 new_page 를 할 수 없습니다.")
    work_page = context.new_page()
    try:
        work_page.bring_to_front()
    except Exception:
        pass
    _domeme_focus_window_for_navigation(work_page)
    try:
        work_page.goto(_DOMEME_GOTO_URL, wait_until="commit", timeout=45000)
    except Exception as e:
        print(f"{log_prefix}[Phase2식 접속] commit goto: {e!r}", flush=True)
    time.sleep(_S(0.35, 0.9))
    try:
        work_page.wait_for_load_state("domcontentloaded", timeout=20000)
    except Exception as e:
        print(f"{log_prefix}[Phase2식 접속] domcontentloaded 대기: {e!r}", flush=True)
    try:
        u0 = (work_page.url or "").strip()
        print(f"{log_prefix}[Phase2식 접속] 1차 url={repr(u0)[:220]}", flush=True)
    except Exception:
        u0 = ""
    if not _url_is_domeme_session(u0):
        try:
            work_page.goto(_DOMEME_GOTO_URL, wait_until="domcontentloaded", timeout=120000)
            u1 = (work_page.url or "").strip()
            print(f"{log_prefix}[Phase2식 접속] dcl 재goto 후 url={repr(u1)[:220]}", flush=True)
        except Exception as e2:
            print(f"{log_prefix}[Phase2식 접속] dcl 재goto 실패: {e2!r}", flush=True)
    _domeme_log_browser_context_state(context, log_prefix, "Phase2식 도매매 연결 후", work_page)
    return work_page


def _domeme_startup_new_tab_goto_domemedb(
    context,
    *,
    log_prefix: str = "[시작] ",
):
    """Chrome 실행 직후 도매매 연결: Phase 2 와 동일한 new_page + commit/dcl 경로."""
    return _domeme_open_domemedb_like_phase2_new_page(context, log_prefix=log_prefix)


def _domeme_goto_load(work_page, url: str, log_prefix: str = "", timeout_ms: int = 300000) -> bool:
    """주소창·키보드 없이: CDP → commit → work_page.goto(load/dcl) → location.href 순."""
    if timeout_ms < 30_000:
        timeout_ms = 30_000
    _domeme_close_internal_pages(work_page.context, work_page=work_page, log_prefix=log_prefix)
    _domeme_focus_window_for_navigation(work_page)
    _domeme_debug_dump_pages(work_page.context, f"{log_prefix}[goto_load] ")

    print(f"{log_prefix}[네비] ① CDP Page.navigate → {url[:90]}", flush=True)
    if _domeme_navigate_cdp(work_page, url, log_prefix):
        try:
            u_ok = work_page.url or ""
        except Exception:
            u_ok = ""
        _domeme_log_navigation_assert(
            _url_is_domeme_session(u_ok), u_ok, log_prefix, "goto_load CDP 후"
        )
        return _url_is_domeme_session(u_ok)

    print(f"{log_prefix}[네비] ② work_page.goto commit 45s", flush=True)
    try:
        work_page.goto(url, wait_until="commit", timeout=45000)
        time.sleep(0.35)
        try:
            u = (work_page.url or "").strip().lower()
        except Exception:
            u = ""
        if u and u != "about:blank" and not u.startswith("chrome://"):
            print(f"{log_prefix}[네비] commit 후 url={repr(work_page.url)[:160]}", flush=True)
            return True
    except Exception as e:
        print(f"{log_prefix}[네비] commit 실패: {e!r}", flush=True)

    for wt in ("load", "domcontentloaded"):
        try:
            print(f"{log_prefix}[네비] ③ work_page.goto {wt} (최대 {timeout_ms // 1000}s)", flush=True)
            work_page.goto(url, wait_until=wt, timeout=timeout_ms)
            time.sleep(0.35)
            try:
                u = (work_page.url or "").strip()
            except Exception:
                u = ""
            print(f"{log_prefix}work_page.goto({wt}) 직후 work_page.url={repr(u)[:200]}", flush=True)
            if not (u.lower() in ("about:blank", "") or u.lower().startswith("chrome://")):
                return True
        except Exception as e:
            print(f"{log_prefix}goto 실패 url={url[:90]!r} wait={wt}: {e!r}", flush=True)

    try:
        work_page.evaluate("""(href) => { window.location.href = href; }""", url)
        deadline = time.time() + 45.0
        while time.time() < deadline:
            try:
                u2 = (work_page.url or "").strip().lower()
            except Exception:
                u2 = ""
            if u2 and u2 != "about:blank" and not u2.startswith("chrome://"):
                print(f"{log_prefix}location.href 반영 url={repr(work_page.url)[:160]}")
                return True
            time.sleep(0.25)
    except Exception as e:
        print(f"{log_prefix}location.href 실패: {e}")
    return False


def _domeme_navigate_cdp(work_page, url: str, log_prefix: str) -> bool:
    """CDP Page.navigate — 일부 프로필에서 Playwright goto 가 blank에 멈출 때 보조."""
    _domeme_focus_window_for_navigation(work_page)
    try:
        sess = work_page.context.new_cdp_session(work_page)
        sess.send("Page.enable", {})
        sess.send("Page.navigate", {"url": url})
        deadline = time.time() + 120.0
        while time.time() < deadline:
            try:
                if _url_is_domeme_session(work_page.url):
                    return True
            except Exception:
                pass
            time.sleep(0.35)
        print(f"{log_prefix}CDP navigate 후에도 domeggook URL 미확인(120초)")
    except Exception as e:
        print(f"{log_prefix}CDP navigate 실패: {e}")
    return False


def _domeme_navigate_address_bar(page, url: str, log_prefix: str) -> bool:
    """(4)(5) 정책상 비활성화 — 주소창·Ctrl+L·keyboard 타입 네비는 사용하지 않는다."""
    print(f"{log_prefix}[비활성] 주소창 키입력 네비 제거됨(_domeme_navigate_address_bar)", flush=True)
    return False


def _domeme_navigate_with_fallbacks(work_page, log_prefix: str = "") -> bool:
    """보조 URL들에 대해 work_page.goto(domcontentloaded) 및 _domeme_goto_load만 사용(주소창 없음)."""
    _domeme_focus_window_for_navigation(work_page)
    _domeme_debug_dump_pages(work_page.context, f"{log_prefix}[fallback] ")
    targets = [
        DOMEME_URL,
        DOMEME_MEMBER_LOGIN_ENTRY,
        DOMEME_LOGIN_URL,
        DOMEME_URL_ROOT,
    ]
    timeout_ms = max(30_000, 120_000)
    for url in targets:
        err = None
        try:
            work_page.set_default_navigation_timeout(timeout_ms)
            work_page.set_default_timeout(timeout_ms)
        except Exception:
            pass
        try:
            work_page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
        except Exception as e:
            err = e
        try:
            u = (work_page.url or "").strip()
        except Exception as e:
            u = ""
            if err is None:
                err = e
        print(f"{log_prefix}work_page.goto 직후 work_page.url={repr(u)} (목표={url[:70]}…)", flush=True)
        if _url_is_domeme_session(u):
            _domeme_log_navigation_assert(True, u, log_prefix, "navigate_with_fallbacks goto")
            return True
        _domeme_log_navigation_assert(False, u, log_prefix, "navigate_with_fallbacks goto")
        _domeme_print_page_nav_diagnostics(work_page, err, log_prefix)
        if _domeme_goto_load(work_page, url, log_prefix, timeout_ms=300000):
            try:
                u2 = work_page.url or ""
            except Exception:
                u2 = ""
            if _url_is_domeme_session(u2):
                return True
    try:
        work_page.evaluate(
            """(href) => { window.location.href = href; }""",
            DOMEME_URL,
        )
        deadline = time.time() + 120.0
        while time.time() < deadline:
            try:
                if _url_is_domeme_session(work_page.url):
                    return True
            except Exception:
                pass
            time.sleep(0.35)
        print(f"{log_prefix}location.href 후에도 domeggook URL 미확인(120초)", flush=True)
    except Exception as e:
        print(f"{log_prefix}location.href 폴백 실패: {e}", flush=True)
    if _domeme_navigate_cdp(work_page, DOMEME_URL, log_prefix):
        return True
    return False


def _domeme_attach_work_page(context):
    """context.pages 에서 고르지 않고, new_page()→goto 먼저. Phase2식 접속."""
    time.sleep(0.3)
    return _domeme_open_domemedb_like_phase2_new_page(context, log_prefix="[시작] attach→")


PROJECT_DIR = Path(__file__).resolve().parent
# 엑셀 다운로드 저장 경로: 국내위탁\마이박스\{26년3월1주차}\{1번사업자}\
EXCEL_SAVE_BASE = Path(r"C:\Users\USER\Documents\국내위탁\마이박스")
# run_all_steps 후 _최종.xlsx 를 스피드고 마이박스에 업로드 (해시태그 검색 → 200개 보기 → 엑셀업로드 → 파일선택). 전송은 별도. (0이면 생략)
ENABLE_SPEEDGO_UPLOAD_AND_SEND = os.environ.get("SPEEDGO_UPLOAD", "1").lower() in ("1", "true", "yes")
# Phase 2(엑셀 업로드) 전용 Chrome 프로필. 새 창으로 해당 프로필 사용. 기본 Profile 67
PHASE2_CHROME_PROFILE_DIR = os.environ.get("PHASE2_CHROME_PROFILE_DIR", "Profile 67").strip() or None
# Phase 1·2 공통: launch_persistent_context 의 ignore_default_args (Phase2와 동일 스택)
CHROME_PERSISTENT_IGNORE_ARGS_SHARED = [
    "--incognito", "--guest", "--off-the-record", "--bwsi", "--inprivate",
    "--enable-automation", "--no-sandbox",
    "--disable-extensions", "--disable-sync", "--disable-default-apps",
    "--disable-component-extensions-with-background-pages",
    "--disable-background-networking",
    "--disable-client-side-phishing-detection",
]
# 비시크릿 모드용 프로필. 기존 폴더 문제 의심 시 USE_FRESH_PROFILE=True
USE_FRESH_PROFILE = os.environ.get("DOMEME_FRESH_PROFILE", "").lower() in ("1", "true", "yes")
# 자동화 전용 프로필 (domeme_browser_profile)
AUTOMATION_PROFILE = PROJECT_DIR / ("fresh_chrome_profile_01" if USE_FRESH_PROFILE else "domeme_browser_profile")
# 실제 사용자 Chrome 프로필 (C:\Users\USER\AppData\Local\Google\Chrome\User Data)
REAL_CHROME_USER_DATA = Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "User Data"
# USE_REAL_CHROME_PROFILE=1 → 실제 프로필 사용 (기본). 0 → 자동화 전용 프로필
# USE_REAL_CHROME_PROFILE_COPY=1 → 프로필 복사본 (Chrome 실행 중에도 가능, 기본값)
USE_REAL_CHROME_PROFILE = os.environ.get("USE_REAL_CHROME_PROFILE", "1").lower() in ("1", "true", "yes")
USE_REAL_PROFILE_COPY = os.environ.get("USE_REAL_CHROME_PROFILE_COPY", "1").lower() in ("1", "true", "yes") if USE_REAL_CHROME_PROFILE else False
REAL_PROFILE_COPY_DIR = PROJECT_DIR / "chrome_real_profile_copy"
# 1이면 24h 재사용 무시하고 실제 프로필 전체 재복사 (첫 실행 프로필 오류·손상 의심 시)
FORCE_CHROME_PROFILE_COPY = os.environ.get("CHROME_PROFILE_COPY_FORCE", "").lower() in ("1", "true", "yes")
# 엑셀 프로필 비교 리스너: context마다 한 번만 (매 사업자마다 중복 등록 시 이벤트·루프 부담)
_EXCEL_PROFILE_LOG_CONTEXT_IDS = set()

def _resolve_user_data_dir():
    if not USE_REAL_CHROME_PROFILE:
        return AUTOMATION_PROFILE
    if USE_REAL_PROFILE_COPY:
        return REAL_PROFILE_COPY_DIR
    return REAL_CHROME_USER_DATA


def _phase1_named_profile_for_launch(user_data_base: Path) -> str | None:
    """Phase2 와 같은 프로필 폴더명을 기본으로 시도(PHASE2_CHROME_PROFILE_DIR / Profile 67). 폴더가 있을 때만 --profile-directory."""
    name = os.environ.get("PHASE1_CHROME_PROFILE_DIR", "").strip()
    if not name:
        name = (PHASE2_CHROME_PROFILE_DIR or "Profile 67").strip()
    sub = Path(user_data_base) / name
    if not sub.is_dir():
        print(
            f"[경고] Phase1: Phase2와 동일하게 쓰려면 user_data_dir 아래에 '{name}' 폴더가 필요합니다: {sub}\n"
            f"        → --profile-directory 생략. CHROME_PROFILE_COPY_FORCE=1 전체 재복사 또는 "
            f"PHASE1_CHROME_PROFILE_DIR=Default 등 복사본에 실제 있는 폴더명을 지정하세요.",
            flush=True,
        )
        return None
    return name


# 프로필 비교 로그 (USE_REAL_CHROME_PROFILE / 자동화 프로필 둘 다 실행 후 비교)
ENABLE_PROFILE_COMPARE_LOG = os.environ.get("PROFILE_COMPARE_LOG", "1").lower() in ("1", "true", "yes")
PROFILE_COMPARE_LOG_FILE = PROJECT_DIR / "profile_compare_log.txt"
# 실제 Chrome 경로 (channel 실패 시 대체). None이면 channel="chrome"만 사용
CHROME_EXECUTABLE = os.environ.get("CHROME_PATH") or r"C:\Program Files\Google\Chrome\Application\chrome.exe"
# FAST_MODE=1: 대기시간 축소 (불안정할 수 있음). 기본 0=안정 우선.
FAST_MODE = os.environ.get("FAST_MODE", "0").lower() in ("1", "true", "yes")
# CHROME_PROCESS_DIAG=1: 시작 시 Chrome cmdline 진단(psutil/PowerShell). 기본 끔(시작·첫 이동 체감 속도).
CHROME_PROCESS_DIAG = os.environ.get("CHROME_PROCESS_DIAG", "0").lower() in ("1", "true", "yes")
_S = lambda quick, slow: (quick if FAST_MODE else slow)  # _S(0.2, 1.5) → FAST면 0.2
_WAIT = "domcontentloaded" if FAST_MODE else "load"  # 페이지 로드 대기: domcontentloaded가 더 빠름


def _p1_close_stale_tabs(context, keep_pages, label: str = "") -> int:
    """사업자 전환 시 누적된 stale 탭(이전 도매매 검색결과·스피드고 마이박스·다운로드 팝업 등)을
    닫는다. keep_pages 에 포함된 page 객체와 'about:blank' keep-alive 만 유지.

    탭 누적의 부작용:
      - login/검색 click 이 stale 탭으로 빨려들어가 wrong-account 동작
      - dialog/download 이벤트가 잘못된 page 객체에 listener 가 연결되어 timeout
      - Chrome 메모리·이벤트 루프 점유 누적 → 후반 사업자 갈수록 느려짐/끊김
    """
    closed = 0
    keep_ids = {id(p) for p in keep_pages if p is not None}
    try:
        pages = list(context.pages)
    except Exception:
        return 0
    # keep-alive(about:blank) 1개는 유지: 모든 작업 탭 닫혀도 Chrome 자체가 종료되지 않게.
    kept_blank = False
    for pg in pages:
        try:
            if pg.is_closed():
                continue
            if id(pg) in keep_ids:
                continue
            try:
                u = (pg.url or "").lower().strip()
            except Exception:
                u = ""
            if not kept_blank and (u in ("about:blank", "", "chrome://newtab/") or u.startswith("chrome://")):
                kept_blank = True
                continue
            pg.close()
            closed += 1
        except Exception:
            continue
    if closed:
        print(f"{label}stale 탭 {closed}개 정리(남은 탭은 작업/keep-alive)", flush=True)
    return closed


def _domeme_work_page(context, prev_work_page, label: str):
    """context.pages[]·pages[0]·chrome 탭을 작업 대상으로 쓰지 않는다.
    이전 루프의 domemedb work_page(prev_work_page)만 재사용. 그 외는 new_page()+검증 goto.
    work_page 생성 전에는 탭 정리를 하지 않는다."""
    work_page = None
    if prev_work_page is not None:
        try:
            if not prev_work_page.is_closed():
                pu = (prev_work_page.url or "").lower()
                if "domemedb.domeggook.com" in pu and not _domeme_url_is_chrome_or_devtools(pu):
                    work_page = prev_work_page
        except Exception:
            work_page = None
    if work_page is None:
        try:
            work_page = _domeme_open_domemedb_like_phase2_new_page(
                context, log_prefix=f"{label}새탭→"
            )
        except Exception as e:
            print(f"{label}work_page 생성 실패: {e}", flush=True)
            return None
    _domeme_focus_window_for_navigation(work_page)
    try:
        u = work_page.url or ""
    except Exception:
        u = ""
    if not _url_is_domeme_session(u):
        print(f"{label}domemedb 세션이 아님 url={repr(u)[:100]} → 보조 네비", flush=True)
        if not _domeme_goto_load(work_page, DOMEME_URL, label):
            _domeme_navigate_with_fallbacks(work_page, label)
    else:
        print(f"{label}작업 탭 url={repr(u)[:120]}")
    return work_page


# 주차 계산: 1~7일=1째주, 8~14일=2째주, 15~21일=3째주, 22~28일=4째주, 29~31일=5째주
def _get_week_of_month(dt):
    return min((dt.day - 1) // 7 + 1, 5)


def _normalize_year_week_key(s: str) -> str:
    """주차 키 통일: 202505w3 → 2505w3 (2자리 연도). 비교용."""
    if not s or "w" not in s:
        return s
    part, _, week = s.partition("w")
    if not week.isdigit():
        return s
    # 202505 → 2505, 2505 → 2505
    if len(part) >= 4 and part.isdigit():
        part = part[-4:]  # 마지막 4자리 = yyMM
    return f"{part}w{week}"


# 1주 7회 실행 (하루 1회): 1~42등 키워드 (6사업자×7회). 회차별 카테고리 분리로 겹침 방지.
# 회차 선택: 실행할 때마다 1→7 순환(.week_run_state) 대신, 마이박스 저장 경로의
# `…/주차폴더/N회차/M번사업자` 폴더 미생성 기준으로 빈 곳만 채운다 (오류로 중단된 회차 재실행에 유리).
# WEEK_RUN=3 처럼 지정하면 해당 회차만 보고, 빈 사업자만 실행한다.
RUNS_PER_WEEK = 7
WEEK_RUN_STATE_FILE = PROJECT_DIR / ".week_run_state"
WEEK_KEYWORDS_FILE = PROJECT_DIR / ".week_keywords"  # 직전주차 키워드 저장 (겹치면 7~12위 사용)
KEYWORDS_PER_WEEK = 6 * RUNS_PER_WEEK  # 42


def _state_year_week_token(target_year: int, target_month: int, target_week: int) -> str:
    """`.week_run_state`(get_upload_path_from_state 등) 첫 필드 포맷. 예: 2505w3."""
    return f"{target_year % 100}{target_month:02d}w{target_week}"


def _biz_folder_path(ymw_str: str, week_run: int, rank: int) -> Path:
    """Phase 1 저장과 동일: 마이박스/{주차}/{N}회차/{M}번사업자."""
    return EXCEL_SAVE_BASE / ymw_str / f"{week_run}회차" / f"{rank}번사업자"


def _missing_ranks_for_run(ymw_str: str, week_run: int, n_accounts: int):
    """미완 사업자 판정.

    이전엔 폴더 존재 여부만 봤는데, Phase1 이 폴더만 만들고 실패한 경우(예: 쿠키 잔류로
    로그인 폼 미표시 → 다음 사업자) 빈 폴더가 '완료'로 잘못 집계되어 다음 회차로 넘어가지
    못한 채 영구히 스킵되는 문제가 있었음. 폴더 안에 ``_최종.xlsx`` 가 1개 이상 있어야
    완료로 간주한다.
    """
    miss = []
    for r in range(1, n_accounts + 1):
        bf = _biz_folder_path(ymw_str, week_run, r)
        if not bf.exists():
            miss.append(r)
            continue
        try:
            done = any(p.name.endswith("_최종.xlsx") or "회_최종" in p.name for p in bf.iterdir() if p.is_file())
        except Exception:
            done = False
        if not done:
            miss.append(r)
    return miss


def _parse_week_run_env():
    ev = os.environ.get("WEEK_RUN", "").strip()
    if not ev:
        return None
    try:
        return min(max(int(ev), 1), RUNS_PER_WEEK)
    except ValueError:
        return None


def _resolve_week_plan_from_folders(ymw_str: str, n_accounts: int, forced_week_run):
    """1~7회차 순으로, 해당 주차 폴더가 없는 가장 앞 회차와 그때 비어 있는 사업자 번호들. 없으면 None."""
    if forced_week_run is not None:
        wr = int(forced_week_run)
        miss = _missing_ranks_for_run(ymw_str, wr, n_accounts)
        if not miss:
            return None
        return wr, miss
    for wr in range(1, RUNS_PER_WEEK + 1):
        miss = _missing_ranks_for_run(ymw_str, wr, n_accounts)
        if miss:
            return wr, miss
    return None


def _write_week_run_state(year_week_token: str, week_run: int) -> None:
    try:
        WEEK_RUN_STATE_FILE.write_text(
            f"{year_week_token.strip()},{int(week_run)}", encoding="utf-8"
        )
        print(
            f"[회차 상태] {WEEK_RUN_STATE_FILE} → {year_week_token} {week_run}회차 "
            "(이번 실행 회차; get_upload_path_from_state 와 동기화)"
        )
    except Exception as e:
        print(f"[회차 상태] 저장 실패: {e}")


# 회차별 카테고리 (사진 기준): 의류·가전 등 겹치지 않게 회차마다 다른 대분류
RUN_CATEGORY = {
    1: "패션의류",
    2: "스포츠/레저",
    3: "디지털/가전",
    4: "가구/인테리어",
    5: "여가/생활편의",
    6: "생활/건강",
    7: "패션잡화",
}

# 카테고리별 키워드 풀 30개 (5주×6). 주차마다 6개씩 돌려쓰기 → 1월만 해도 니트 등 겹침 방지.
# 스포츠/레저는 WEEK_KEYWORD_CANDIDATES(월,주) 사용.
CATEGORY_DEFAULT_KEYWORDS = {
    "패션의류": [
        "니트", "원피스", "블라우스", "맨투맨", "바지", "스커트",
        "겨울코트", "패딩", "목도리", "가디건", "트렌치", "스웨터",
        "점퍼", "롱패딩", "니트가디건", "청바지", "레깅스", "후드티",
        "블레이저", "코트", "기모바지", "바람막이", "후드집업", "플리스",
        "조끼", "롱스커트", "티셔츠", "린넨바지", "슬랙스", "반바지",
    ],
    "스포츠/레저": None,  # WEEK_KEYWORD_CANDIDATES(month, week) 사용
    "디지털/가전": [
        "이어폰", "충전기", "스마트워치", "블루투스스피커", "보조배터리", "케이블",
        "스마트폰케이스", "무선이어폰", "USB허브", "거치대", "보호필름", "노트북파우치",
        "태블릿", "키보드", "마우스", "웹캠", "이어폰케이스", "무선충전기",
        "멀티포트", "케이블정리", "HDMI", "C타입", "블루투스이어폰", "스피커",
        "충전케이블", "보호케이스", "스탠드", "파우치", "갤럭시워치", "액세서리",
    ],
    "가구/인테리어": [
        "수납장", "조명", "침구", "커튼", "매트", "인테리어소품",
        "책장", "수납박스", "램프", "이불", "카펫", "화분",
        "거울", "선반", "소파커버", "테이블", "의자", "행거",
        "수납정리", "무드등", "베개", "러그", "벽선반", "옷걸이",
        "수납함", "인테리어조명", "침대커버", "러그매트", "소품", "인테리어",
    ],
    # 서적/DVD/CD음반 제외 (정가정책) - 다이어리·문구·플래너·노트 등 대체
    "여가/생활편의": [
        "캠핑용품", "텀블러", "휴대폰케이스", "휴대용품", "정리수납", "생활편의소품",
        "보냉백", "파우치", "여행용품", "생활소품", "손소독제", "이어폰케이스",
        "수납함", "휴대폰거치대", "USB선풍기", "마스크케이스", "정리함",
        "수납박스", "캠핑의자", "물병", "스티커", "메모판", "키홀더",
        "카드지갑", "케이블정리", "테이블매트", "도시락",
    ],
    # 의약품·건강보조식품 제외, 생활용품·위생·찜질 등만
    "생활/건강": [
        "마스크", "핫팩", "손난로", "손세정제", "물티슈", "생활용품",
        "발열패치", "마스크스트랩", "밴드", "반창고", "온열패치", "찜질팩",
        "수건", "목욕용품", "구강케어", "세안용품", "습도계", "체온계",
        "약보관함", "거즈", "소독솜", "안대", "목베개", "발매트",
        "세탁용품", "섬유유연제", "탈취제", "방향제", "제습제", "생활소품",
    ],
    "패션잡화": [
        "가방", "지갑", "모자", "양말", "벨트", "스카프",
        "크로스백", "백팩", "토트백", "넥타이", "장갑", "우산",
        "선글라스", "시계", "반지", "목걸이", "귀걸이", "브레이슬릿",
        "키링", "파우치", "스타킹", "넥워머", "캡모자", "비니",
        "마스크", "벨트", "지갑", "양말", "가방", "액세서리",
    ],
}
PRODUCTS_PER_SEARCH = 200
TARGET_PER_WEEK = 1000

# N+2월 N째주 키워드 1~6등 (기본): (month, week) → [1등~6등]
# 1~30등은 세분류로 자동 확장 (6×5)
WEEK_KEYWORD_CANDIDATES = {
    (1, 1): ["겨울코트", "히터", "난로", "전기요", "담요", "핫팩"],
    (1, 2): ["겨울패딩", "목도리", "장갑", "터틀넥", "방한화", "기모스타킹"],
    (1, 3): ["겨울부츠", "기모바지", "히트텍", "내복", "양말", "발열조끼"],
    (1, 4): ["겨울용품", "수면양말", "가습기", "에어워셔", "공기청정기", "전기담요"],
    (1, 5): ["겨울아우터", "다운코트", "빅사이즈패딩", "캐시미어", "니트", "방한모자"],
    (2, 1): ["발렌타인", "초콜릿", "선물상자", "카드", "꽃다발", "포장용품"],
    (2, 2): ["발렌타인선물", "초콜릿세트", "케이스", "리본", "포장지", "선물박스"],
    (2, 3): ["롱패딩", "울코트", "트렌치코트", "코트", "점퍼", "니트가디건"],
    (2, 4): ["화이트데이", "답례품", "캔디", "과자", "쿠키", "초콜릿선물"],
    (2, 5): ["봄신상", "니트", "가디건", "블라우스", "원피스", "트렌치"],
    (3, 1): ["등산용품", "등산화", "등산복", "배낭", "트레킹화", "등산스틱"],
    (3, 2): ["아웃도어", "캠핑", "텐트", "슬리핑백", "랜턴", "캠핑매트"],
    (3, 3): ["등산가방", "트레킹팩", "폴딩체어", "테이블", "버너", "아이스박스"],
    (3, 4): ["산행용품", "등산스틱", "모자", "선글라스", "물병", "등산화"],
    (3, 5): ["봄아웃도어", "자켓", "바람막이", "레이어드", "스포츠웨어", "런닝복"],
    (4, 1): ["캠핑용품", "캠핑의자", "캠핑테이블", "BBQ그릴", "아이스박스", "텐트"],
    (4, 2): ["캠핑장비", "텐트", "타프", "캔들", "파우치", "랜턴"],
    (4, 3): ["피크닉", "돗자리", "와인잔", "포크세트", "캐리어", "보냉백"],
    (4, 4): ["가정의달", "선물", "홈웨어", "이불", "침구", "쿠션"],
    (4, 5): ["봄자켓", "가디건", "트렌치", "바람막이", "블루종", "린넨자켓"],
    (5, 1): ["봄자켓", "러닝화", "등산가방", "도시락통", "텀블러", "스포츠가방"],
    (5, 2): ["밀짚모자", "썬캡", "캡모자", "파라솔", "선글라스", "양산"],
    (5, 3): ["운동화", "런닝화", "스니커즈", "트레일러", "워킹화", "실내화"],
    (5, 4): ["반팔티", "반바지", "린넨바지", "린넨", "시원한원피스", "캐주얼티"],
    (5, 5): ["피크니크", "아이스팩", "보냉가방", "물통", "도시락", "보냉팩"],
    (6, 1): ["선캡", "선글라스", "파라솔", "양산", "여름모자", "썬베드"],
    (6, 2): ["텀블러", "보냉컵", "아이스컵", "물병", "빨대", "스탠리텀블러"],
    (6, 3): ["캠핑의자", "캠핑테이블", "해먹", "썬베드", "우산", "파라솔"],
    (6, 4): ["세차용품", "세차장갑", "워시밋", "광택제", "왁스", "세차스펀지"],
    (6, 5): ["선풍기", "USB선풍기", "탁상선풍기", "무선선풍기", "미니선풍기", "서큘레이터"],
    (7, 1): ["휴가용품", "수영복", "래쉬가드", "비치웨어", "비치타올", "비치가방"],
    (7, 2): ["선캡", "파라솔", "양산", "모기퇴치", "파리채", "여름모자"],
    (7, 3): ["에어컨", "선풍기", "제습기", "제습제", "에어컨커버", "공기청정기"],
    (7, 4): ["수영용품", "수경", "핀", "플로트", "비치볼", "수영가방"],
    (7, 5): ["여름휴양", "여행가방", "캐리어", "넥쿨러", "아이스팩", "보냉백"],
    (8, 1): ["스쿨백", "가방", "파우치", "필기구", "휴대폰케이스", "휴대폰거치대"],
    (8, 2): ["볼펜", "형광펜", "스티커", "메모판", "펜케이스", "파우치"],
    (8, 3): ["운동화", "등산화", "스니커즈", "실내화", "운동복", "워킹화"],
    (8, 4): ["니트가디건", "가디건", "니트", "바람막이", "맨투맨", "후드집업"],
    (8, 5): ["MT용품", "텐트", "취사도구", "랜턴", "캠핑의자", "침낭"],
    (9, 1): ["가을자켓", "니트", "가디건", "트렌치코트", "블레이저", "맨투맨"],
    (9, 2): ["등산복", "야상", "플리스", "후드집업", "패딩조끼", "바람막이"],
    (9, 3): ["가을원피스", "니트원피스", "롱스커트", "레깅스", "부츠", "니트"],
    (9, 4): ["추석선물", "선물세트", "한과", "과일", "선물박스", "선물용품"],
    (9, 5): ["가을코트", "트렌치", "울코트", "레이어드", "스타킹", "가디건"],
    (10, 1): ["핼로윈", "코스튬", "가면", "장식", "호박", "파티용품"],
    (10, 2): ["가을코트", "패딩", "니트", "부츠", "목도리", "가디건"],
    (10, 3): ["니트", "가디건", "맨투맨", "후드", "집업", "트렌치"],
    (10, 4): ["발열조끼", "패딩", "히트텍", "내복", "양말", "기모레깅스"],
    (10, 5): ["부츠", "앵클부츠", "첼시부츠", "워커", "방한화", "울부츠"],
    (11, 1): ["가을코트", "패딩", "부츠", "히트텍", "난로", "담요"],
    (11, 2): ["겨울코트", "패딩", "목도리", "장갑", "모자", "발열조끼"],
    (11, 3): ["히터", "전기장판", "담요", "핫팩", "손난로", "전기요"],
    (11, 4): ["롱패딩", "다운코트", "숏패딩", "패딩조끼", "기모패딩", "겨울패딩"],
    (11, 5): ["겨울아우터", "다운", "패딩조끼", "울코트", "기모", "다운코트"],
    (12, 1): ["겨울코트", "패딩", "목도리", "장갑", "히터", "전기담요"],
    (12, 2): ["크리스마스", "트리", "장식", "선물", "캔들", "리스"],
    (12, 3): ["겨울용품", "히터", "담요", "전기요", "가습기", "핫팩"],
    (12, 4): ["롱패딩", "다운코트", "울코트", "숏패딩", "패딩조끼", "겨울아우터"],
    (12, 5): ["시계", "알람시계", "거울", "수납함", "홈데코", "조명"],
}

# 키워드별 세분류 후보: Naver API로 월별 트렌드 높은 키워드 선정 후 도매매 검색에 사용
# 키가 WEEK_KEYWORD_CANDIDATES에 있는 키워드와 일치할 때 API로 비교하여 최적 세분류 선택
KEYWORD_SUBCATEGORIES = {
    # 모자/캡 계열
    "밀짚모자": ["밀짚모자", "볼캡", "썬캡", "야구모자", "캡모자"],
    "썬캡": ["썬캡", "볼캡", "야구모자", "캡모자", "스냅백"],
    "캡모자": ["캡모자", "볼캡", "야구모자", "스냅백", "썬캡"],
    "선캡": ["볼캡", "썬캡", "야구모자", "캡모자", "밀짚모자"],
    "여름모자": ["볼캡", "썬캡", "밀짚모자", "야구모자", "캡모자"],
    "방한모자": ["방한모자", "비니", "털모자", "기모캡", "넥워머"],
    "모자": ["볼캡", "야구모자", "밀짚모자", "썬캡", "캡모자"],
    # 운동화/신발 계열
    "운동화": ["런닝화", "등산화", "워킹화", "스니커즈", "트레일러"],
    "등산화": ["등산화", "트레킹화", "워킹화", "등산부츠", "하이킹화"],
    "런닝화": ["런닝화", "스니커즈", "워킹화", "트레일러", "마라톤화"],
    "스니커즈": ["스니커즈", "런닝화", "캔버스화", "라이프스타일", "워킹화"],
    "워킹화": ["워킹화", "등산화", "런닝화", "스니커즈", "트레킹화"],
    "트레킹화": ["트레킹화", "등산화", "워킹화", "하이킹화", "등산부츠"],
    "실내화": ["실내화", "슬리퍼", "개그신", "쿠션슬리퍼", "EVA실내화"],
    "트레일러": ["트레일러", "트레일러닝화", "오프로드", "그립화", "아웃도어런닝"],
    "부츠": ["앵클부츠", "첼시부츠", "워커", "미들부츠", "롱부츠"],
    "앵클부츠": ["앵클부츠", "첼시부츠", "워커", "미들부츠", "단화"],
    "방한화": ["방한화", "털부츠", "기모부츠", "스노우부츠", "발열부츠"],
    "겨울부츠": ["겨울부츠", "털부츠", "기모부츠", "스노우부츠", "앵클부츠"],
    # 자켓/아우터 계열
    "봄자켓": ["봄자켓", "트렌치", "바람막이", "청자켓", "블루종"],
    "가을자켓": ["가을자켓", "트렌치", "가디건", "블레이저", "바람막이"],
    "겨울코트": ["겨울코트", "롱패딩", "다운코트", "울코트", "코트"],
    "가을코트": ["가을코트", "트렌치", "울코트", "레이어드코트", "블레이저"],
    "바람막이": ["바람막이", "경량자켓", "윈드브레이커", "방수자켓", "트레킹자켓"],
    "트렌치": ["트렌치", "트렌치코트", "레인코트", "롱자켓", "블레이저"],
    "등산복": ["등산복", "트레킹복", "아웃도어자켓", "등산바람막이", "야상"],
    "겨울패딩": ["겨울패딩", "롱패딩", "숏패딩", "다운패딩", "기모패딩"],
    "겨울아우터": ["겨울아우터", "다운코트", "패딩", "울코트", "히트텍"],
    # 선글라스/양산 계열
    "선글라스": ["선글라스", "편광선글라스", "스포츠선글라스", "드라이빙선글라스", "클립온"],
    "파라솔": ["파라솔", "양산", "접이식양산", "장대양산", "우산"],
    "양산": ["양산", "접이식양산", "장대양산", "미니양산", "파라솔"],
    "우산": ["우산", "접이식우산", "장대우산", "양산", "5단우산"],
    # 캠핑/아웃도어 계열
    "등산용품": ["등산화", "등산가방", "트레킹화", "등산스틱", "등산복"],
    "캠핑용품": ["캠핑의자", "캠핑테이블", "텐트", "랜턴", "아이스박스"],
    "캠핑장비": ["텐트", "타프", "캠핑의자", "캠핑테이블", "랜턴"],
    "등산가방": ["등산가방", "트레킹팩", "배낭", "힙색", "크로스백"],
    "텐트": ["텐트", "돔텐트", "타프", "쉘터", "비치텐트"],
    "캠핑의자": ["캠핑의자", "폴딩체어", "캠핑스툴", "체어", "접이식의자"],
    "캠핑테이블": ["캠핑테이블", "캠핑테이블세트", "롤테이블", "접이식테이블", "캠핑박스"],
    "트레킹팩": ["트레킹팩", "등산가방", "배낭", "하이킹팩", "크로스백"],
    "폴딩체어": ["폴딩체어", "캠핑의자", "캠핑스툴", "접이식의자", "체어"],
    "테이블": ["캠핑테이블", "롤테이블", "접이식테이블", "캠핑테이블세트", "테이블"],
    "버너": ["버너", "캠핑버너", "가스버너", "카세트버너", "스토브"],
    "아이스박스": ["아이스박스", "보냉백", "캠핑쿨러", "도시락", "보냉가방"],
    # 텀블러/물병 계열
    "텀블러": ["스탠리텀블러", "보냉컵", "아이스컵", "물병", "보온컵"],
    "보냉컵": ["보냉컵", "아이스컵", "텀블러", "스탠리", "보냉물병"],
    "물병": ["물병", "보냉물병", "스포츠물병", "텀블러", "보온물병"],
    # 선풍기/가전 계열
    "선풍기": ["USB선풍기", "탁상선풍기", "무선선풍기", "미니선풍기", "서큘레이터"],
    "히터": ["히터", "전기히터", "오일히터", "선풍히터", "커넥터히터"],
    "난로": ["전기난로", "오일히터", "전기히터", "히터", "발열기"],
    "에어컨": ["에어컨커버", "휴대용에어컨", "선풍기", "에어컨필터", "제습기"],
    # 겨울용품 계열
    "전기요": ["전기요", "전기장판", "발열매트", "담요", "히트매트"],
    "담요": ["담요", "담요블랭킷", "양털담요", "전기담요", "무릎담요"],
    "핫팩": ["핫팩", "발열패드", "손난로", "찜질팩", "일회용핫팩"],
    "전기담요": ["전기담요", "전기요", "발열담요", "USB담요", "담요"],
    # 가방/잡화 계열
    "가방": ["크로스백", "숄더백", "백팩", "토트백", "메신저백"],
    "스쿨백": ["스쿨백", "백팩", "크로스백", "숄더백", "가방"],
    "등산가방": ["등산가방", "트레킹팩", "배낭", "힙색", "크로스백"],
    "스포츠가방": ["스포츠가방", "헬스가방", "수영가방", "등산가방", "캐리어백"],
    "여행가방": ["캐리어", "여행가방", "더플백", "백팩", "캐리어백"],
    # 세차/선풍기/기타
    "세차용품": ["세차장갑", "워시밋", "광택제", "왁스", "세차스펀지"],
    "문구": ["노트", "볼펜", "형광펜", "다이어리", "메모장"],
    "학교용품": ["가방", "필통", "노트", "볼펜", "다이어리"],
}

# 월별 폴백 (주차 데이터 없을 때)
MONTH_KEYWORD_FALLBACK = {
    1: "겨울용품", 2: "발렌타인", 3: "등산용품", 4: "캠핑용품",
    5: "봄자켓", 6: "선캡", 7: "휴가용품", 8: "스쿨백",
    9: "가을자켓", 10: "핼로윈", 11: "가을코트", 12: "겨울코트",
}

# Naver DataLab API - 제외할 카테고리 (화장품, 식품류, 영유아, 브랜드, 명품, 골프, 낚시, 서적/DVD/CD음반 등)
# 서적/DVD/CD음반: 정가기준 판매 정책상 추천단가 전송→할인율 개별설정 필요
# cat_id 참고: 네이버쇼핑 URL cat_id
NAVER_CATEGORY_EXCLUDED = {
    "50000002",   # 화장품/미용
    "50000004",   # 서적/DVD/CD음반 (정가정책)
    "50000167",   # 식품
    "50000168",   # 영유아용품 (3세미만 포함)
    "50000169",   # 브랜드
    "50000170",   # 명품/디자이너
    "50000550",   # 골프
    "50000551",   # 낚시
}

# API 사용 시 허용 카테고리 (다양화, 제외 8개 제외)
NAVER_CATEGORY_ALLOWED = [
    "50000000",   # 패션의류
    "50000001",   # 가방/잡화
    "50000003",   # 디지털/가전
    "50001310",   # 가구/인테리어
    "50000005",   # 생활/건강(화장품 제외)
    "50000006",   # 스포츠/레저(골프·낚시 제외한 일반 스포츠)
]


def _get_keyword_from_naver_api(target_month, target_week, rank=1, candidates_override=None):
    """Naver DataLab API로 N월 트렌드(판매량) 가장 높은 세분류 키워드 반환.
    candidates_override: 키워드 후보 리스트(세분류). 없으면 WEEK_KEYWORD_CANDIDATES 사용.
    허용 카테고리만, 제외 8개 카테고리 제외."""
    cid = os.environ.get("NAVER_CLIENT_ID") or NAVER_CLIENT_ID
    secret = os.environ.get("NAVER_CLIENT_SECRET") or NAVER_CLIENT_SECRET
    if not cid or not secret:
        return None
    if candidates_override is not None:
        candidates = candidates_override if isinstance(candidates_override, list) else [candidates_override]
    else:
        candidates = WEEK_KEYWORD_CANDIDATES.get((target_month, target_week))
    if not candidates or len(candidates) < 2:
        return candidates[0] if candidates else None
    # 허용 카테고리에서 rank번째로 순환 (다양화)
    cat = NAVER_CATEGORY_ALLOWED[(rank - 1) % len(NAVER_CATEGORY_ALLOWED)]
    if cat in NAVER_CATEGORY_EXCLUDED:
        cat = NAVER_CATEGORY_ALLOWED[0]
    try:
        import urllib.request
        import json

        year = datetime.now().year
        if target_month > datetime.now().month + 2:
            year -= 1
        start = f"{year}-{target_month:02d}-01"
        end = f"{year}-{target_month:02d}-28"

        body = {
            "startDate": start,
            "endDate": end,
            "timeUnit": "week",
            "category": cat,
            "keyword": [{"name": k, "param": [k]} for k in candidates[:5]],
            "device": "",
            "gender": "",
            "ages": [],
        }
        req = urllib.request.Request(
            "https://openapi.naver.com/v1/datalab/shopping/category/keywords",
            data=json.dumps(body).encode(),
            headers={
                "X-Naver-Client-Id": cid,
                "X-Naver-Client-Secret": secret,
                "Content-Type": "application/json",
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            data = json.loads(r.read().decode())
        results = data.get("results", [])
        if not results:
            return None
        # target_month 기간의 ratio 가장 높은 키워드
        best = max(results, key=lambda r: (r.get("data") or [{}])[0].get("ratio", 0))
        return best.get("keyword", [None])[0]
    except Exception as e:
        print(f"[Naver API 실패] {e}")
        return None


def _remove_chrome_profile_lock_files(base: Path) -> int:
    """복사본/프로필 폴더에 남은 Chrome 잠금 파일 제거. 첫 실행 실패(프로필 사용 중) 완화."""
    if not base.is_dir():
        return 0
    lock_names = ("SingletonLock", "SingletonSocket", "SingletonCookie", "lockfile")
    removed = 0
    for name in lock_names:
        p = base / name
        if p.is_file():
            try:
                p.unlink()
                removed += 1
            except OSError:
                pass
    try:
        for child in base.iterdir():
            if not child.is_dir():
                continue
            if child.name.startswith("component_") or child.name in ("Crashpad", "GrShaderCache", "ShaderCache", "GPUCache"):
                continue
            for name in lock_names:
                p = child / name
                if p.is_file():
                    try:
                        p.unlink()
                        removed += 1
                    except OSError:
                        pass
            try:
                for p in child.glob("*.lock"):
                    if p.is_file():
                        try:
                            p.unlink()
                            removed += 1
                        except OSError:
                            pass
            except OSError:
                pass
    except OSError:
        pass
    if removed:
        print(f"[프로필] 잠금 파일 {removed}개 제거: {base}")
    return removed


def _copy_real_chrome_profile():
    """실제 Chrome 프로필을 복사하여 충돌 없이 사용 (Chrome 실행 중에도 가능)"""
    import subprocess

    src = REAL_CHROME_USER_DATA
    dst = REAL_PROFILE_COPY_DIR
    if not src.exists():
        print(f"[경고] 원본 프로필 없음: {src}")
        return False

    if FORCE_CHROME_PROFILE_COPY:
        print("[프로필복사] CHROME_PROFILE_COPY_FORCE=1 → 전체 재복사")

    # 복사본이 24시간 이내면 재사용 (매번 복사 시 1~5분 걸림)
    if not FORCE_CHROME_PROFILE_COPY and dst.exists() and (dst / "Default").exists():
        try:
            mtime = (dst / "Default").stat().st_mtime
            if time.time() - mtime < 24 * 3600:
                print("[프로필복사] 기존 복사본 재사용 (24h 이내)")
                return True
        except OSError:
            pass

    if dst.exists():
        shutil.rmtree(dst, ignore_errors=True)
    dst.mkdir(parents=True, exist_ok=True)

    print(f"[프로필복사] {src} → {dst}")
    print("[프로필복사] 대용량 폴더라 1~5분 걸릴 수 있습니다. 잠시만 기다려 주세요...", flush=True)

    # Windows: robocopy 먼저 시도 (잠긴 파일 복사 가능, /B=백업모드)
    if sys.platform == "win32":
        for extra in [[], ["/B"]]:  # /B 없이 시도 후, 실패 시 /B로 재시도 (관리자 권한 필요할 수 있음)
            try:
                args = ["robocopy", str(src), str(dst), "/E", "/R:1", "/W:1", "/MT:8"] + extra
                r = subprocess.run(args, capture_output=False, timeout=300)
                if r.returncode < 8 and (dst / "Default").exists():
                    print("[프로필복사] 완료")
                    return True
                if extra:
                    print(f"[프로필복사] robocopy/B 실패, 종료코드={r.returncode}")
                else:
                    print(f"[프로필복사] robocopy 종료코드={r.returncode}, /B 모드로 재시도...")
            except subprocess.TimeoutExpired:
                print("[프로필복사] robocopy 타임아웃(5분)")
                break
            except FileNotFoundError:
                print("[프로필복사] robocopy 없음, shutil로...")
                break
            except Exception as e:
                print(f"[프로필복사] robocopy 실패: {e}")
                break

    # shutil 폴백: Chrome이 잠근 파일(Cookies, Sessions 등) 제외 후 복사
    def ignore_lock(path, names):
        base_skip = {"SingletonLock", "SingletonSocket", "lockfile", "SingletonCookie"}
        skip = set(base_skip)
        p = str(path)
        if "Network" in p:
            skip |= {"Cookies", "Cookies-journal"}
        if "Safe Browsing Network" in p:
            skip |= {"Safe Browsing Cookies", "Safe Browsing Cookies-journal"}
        if "Sessions" in p:
            skip |= {n for n in names if n.startswith("Session_") or n.startswith("Tabs_")}
        return [n for n in names if n in skip or n.endswith(".lock")]
    try:
        if dst.exists():
            shutil.rmtree(dst, ignore_errors=True)
        dst.mkdir(parents=True, exist_ok=True)
        # dirs_exist_ok=True: 잠긴 파일로 rmtree가 완전삭제 못 해 잔존 폴더가 있어도
        # WinError 183 없이 덮어쓰기 (재실행 안정성)
        shutil.copytree(src, dst, ignore=ignore_lock, dirs_exist_ok=True)
        print("[프로필복사] 완료 (잠긴 파일 일부 제외, 쿠키/세션은 새로 생성됨)")
        return True
    except (PermissionError, OSError) as e:
        print(f"[프로필복사] shutil 실패: {e}")
        return False


# Phase 2 전용 Profile 67 복사본 (실제 Chrome User Data와 충돌 방지)
REAL_PHASE2_COPY_DIR = PROJECT_DIR / "chrome_phase2_profile67_copy"
# Phase 2 실행 전용 별도 복사본: Phase 1 이 쓰던 디렉터리와 분리하여 잠금 경합 방지
REAL_PHASE2_RUN_DIR = PROJECT_DIR / "chrome_phase2_run_copy"


def _chrome_persistent_launch_kw_phase2_identical(user_data_dir: str) -> dict:
    """Phase 2 블록(`launch_persistent_context`)과 동일 키·동일 args. user_data_dir 만 인자."""
    pd = (PHASE2_CHROME_PROFILE_DIR or "Profile 67").strip()
    kw = {
        "user_data_dir": str(user_data_dir),
        "headless": False,
        "channel": "chrome",
        "ignore_default_args": list(CHROME_PERSISTENT_IGNORE_ARGS_SHARED),
        "args": [
            "--start-maximized",
            "--disable-blink-features=AutomationControlled",
            f"--profile-directory={pd}",
        ],
        "locale": "ko-KR",
        "timeout": 90000,
    }
    if os.path.isfile(CHROME_EXECUTABLE):
        kw["executable_path"] = CHROME_EXECUTABLE
        kw.pop("channel", None)
    return kw


def _copy_phase2_profile67(dst_override=None, *, fresh: bool = False):
    """Phase 2용 Profile 67 복사. '다른 프로세스에서 사용 중' 충돌 방지.

    dst_override: 복사 대상 베이스 디렉터리 (기본 REAL_PHASE2_COPY_DIR).
                  Phase 2 실행 시엔 Phase 1 이 쓰던 디렉터리와 분리된 별도 경로를 넘긴다.
    fresh: True 면 mtime 캐시(24h)를 무시하고 항상 새로 복사.
    """
    profile_dir_name = PHASE2_CHROME_PROFILE_DIR or "Profile 67"
    if USE_REAL_PROFILE_COPY and REAL_PROFILE_COPY_DIR.exists():
        src_profile = REAL_PROFILE_COPY_DIR / profile_dir_name
    else:
        src_profile = REAL_CHROME_USER_DATA / profile_dir_name
    dst_base = Path(dst_override) if dst_override is not None else REAL_PHASE2_COPY_DIR
    dst_profile = dst_base / profile_dir_name
    if not src_profile.exists():
        print(f"[Phase2 프로필복사] 소스 없음: {src_profile}")
        return False
    if not fresh and dst_profile.exists():
        try:
            mtime = dst_profile.stat().st_mtime
            if time.time() - mtime < 24 * 3600:
                return True
        except OSError:
            pass

    def _ignore(path, names):
        return [n for n in names if n in {"SingletonLock", "SingletonSocket", "lockfile", "SingletonCookie"} or n.endswith(".lock")]

    print(f"[Phase2 프로필복사] {src_profile.name} → {dst_base}")
    for _try in range(1, 3):
        if dst_base.exists():
            shutil.rmtree(dst_base, ignore_errors=True)
        try:
            dst_base.mkdir(parents=True, exist_ok=True)
            shutil.copytree(src_profile, dst_profile, ignore=_ignore, dirs_exist_ok=True)
            return True
        except (PermissionError, OSError) as e:
            print(f"[Phase2 프로필복사] 시도 {_try}/2 실패: {e}")
            time.sleep(2)
    return False


def _wait_for_chrome_exit_using_path(user_data_path, timeout_sec: int = 45) -> bool:
    """user_data_dir을 사용 중인 Chrome 프로세스가 종료될 때까지 대기. Phase 1→2 전환용."""
    try:
        import psutil
    except ImportError:
        return True
    path_norm = str(Path(user_data_path).resolve()).replace("\\", "/").lower()
    for elapsed in range(timeout_sec):
        found = False
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                name = (proc.info.get("name") or "").lower()
                if "chrome" not in name and "chromium" not in name:
                    continue
                cmdline = proc.info.get("cmdline") or []
                cmd_str = " ".join(str(c or "") for c in cmdline).replace("\\", "/").lower()
                if path_norm in cmd_str:
                    found = True
                    if elapsed == 0:
                        print(f"[Phase 2 대기] Chrome 프로세스(pid={proc.info.get('pid')}) 종료 대기 중...")
                    break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        if not found:
            if elapsed > 0:
                print(f"[Phase 2 대기] Chrome 종료 확인 ({elapsed}초)")
            return True
        time.sleep(1)
    print(f"[Phase 2 대기] {timeout_sec}초 타임아웃. 계속 진행합니다.")
    return False


def _mybox_all_contexts(page):
    """메인 프레임 + child frame (중복 없음. Page+main_frame 이중 평가 방지)."""
    out = []
    try:
        mf = page.main_frame
        out.append(mf)
        for fr in page.frames:
            if fr != mf:
                out.append(fr)
    except Exception:
        out = [page.main_frame]
    return out


def _mybox_count_item_checkboxes(page) -> int:
    """스피드고 마이박스 상품 행 input[name='item[]'] 개수. (CSS [] 이스케이프 이슈 회피로 getElementsByName 사용)"""
    nmax = 0
    for ctx in _mybox_all_contexts(page):
        try:
            n = ctx.evaluate(
                """() => document.getElementsByName('item[]').length"""
            )
            nmax = max(nmax, int(n))
        except Exception:
            continue
    return nmax


def _mybox_wait_for_rows(page, max_seconds: int = 45) -> int:
    """행이 늦게 로드되면 대기 후 개수 반환."""
    interval = 0.5 if FAST_MODE else 1.0
    steps = max(1, int(max_seconds / interval))
    for _ in range(steps):
        n = _mybox_count_item_checkboxes(page)
        if n > 0:
            return n
        time.sleep(interval)
    return _mybox_count_item_checkboxes(page)


def _mybox_select_all_items(page) -> int:
    """
    excelMbDown() 전제: 체크된 item[] 필요.
    개수 제한 없음 (200/244/500 등 상관없이 getElementsByName으로 전체 선택).
    우선 Playwright로 #selectAll 클릭 → 실패 시 JS 직접 체크.
    """
    best_ctx = None
    best_n = 0
    for ctx in _mybox_all_contexts(page):
        try:
            cnt = ctx.evaluate("""() => document.getElementsByName('item[]').length""")
            cnt = int(cnt)
            if cnt > best_n:
                best_n = cnt
                best_ctx = ctx
        except Exception:
            continue
    if best_n == 0 or best_ctx is None:
        return 0

    def _verify_checked(ctx) -> int:
        try:
            return int(ctx.evaluate(
                """() => { const i = document.getElementsByName('item[]'); let n=0; for(let j=0;j<i.length;j++) if(i[j].checked)n++; return n; }"""
            ))
        except Exception:
            return 0

    # 1) Playwright 클릭: label[for="selectAll"] 또는 #selectAll (미체크 상태에서 클릭 → 체크 + jQuery 전체선택)
    for sel in ('label[for="selectAll"]', '#selectAll', 'input[name="selectAll"]'):
        try:
            loc = best_ctx.locator(sel).first
            if loc.count() == 0:
                continue
            loc.scroll_into_view_if_needed(timeout=3000)
            time.sleep(_S(0.15, 0.3))
            # selectAll이 이미 체크돼 있으면 클릭 시 토글로 해제됨 → 먼저 체크 여부 확인
            try:
                cb = best_ctx.locator("#selectAll").first
                if cb.count() > 0 and cb.is_checked():
                    cb.click(timeout=2000)  # 한 번 해제
                    time.sleep(_S(0.2, 0.4))
            except Exception:
                pass
            loc.click(timeout=3000)
            time.sleep(_S(0.4, 0.8))
            chk = _verify_checked(best_ctx)
            if chk > 0:
                return chk
        except Exception:
            continue

    # 2) JS 직접 체크 (클릭 실패 시)
    for _attempt in range(2):  # 재시도 1회 (DOM 안정화 대기)
        try:
            best_ctx.evaluate(
                """() => {
    const items = document.getElementsByName('item[]');
    for (let i = 0; i < items.length; i++) { items[i].checked = true; }
    const sa = document.getElementById('selectAll');
    if (sa) sa.checked = true;
    const el = document.getElementById('itemCnt');
    if (el) el.textContent = String(items.length);
    if (typeof jQuery !== 'undefined') { jQuery('#itemCnt').html(items.length); }
}"""
            )
            time.sleep(_S(0.4, 0.8) if _attempt == 0 else _S(0.6, 1.2))
            chk = _verify_checked(best_ctx)
            if chk > 0:
                return chk
            if _attempt == 0 and best_n > 50:
                time.sleep(0.5)  # 행 많을 때 DOM 안정화 대기 후 재시도
        except Exception:
            if _attempt == 1:
                return 0
            time.sleep(0.3)
    return 0


def _folder_has_mergeable_source_excel(folder: Path) -> bool:
    """
    STEP1·STEP1-1에 쓸 '원본' 스프레드시트가 있는지.
    마이박스 엑셀 미다운로드 시 파이프라인 산출물만 있으면 False.
    """
    if not folder.is_dir():
        return False
    for f in folder.iterdir():
        if not f.is_file():
            continue
        name = f.name
        low = name.lower()
        if low.endswith(".xls") and not low.endswith(".xlsx"):
            return True
        if not low.endswith(".xlsx"):
            continue
        if "통합상품명" in name:
            continue
        if "keywords" in low:
            continue
        if "domeme_links" in low:
            continue
        if "카테고리매핑" in name:
            continue
        if name.endswith("_최종.xlsx") or "회_최종" in name:
            continue
        return True
    return False


def _append_profile_compare_log(profile_type, excel_result, events):
    """프로필 비교 로그를 파일에 추가 (자동화 vs 실제 프로필 실행 결과 비교용)"""
    try:
        with open(PROFILE_COMPARE_LOG_FILE, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 80 + "\n")
            f.write(f"[{datetime.now().isoformat()}] 프로필: {profile_type} | 엑셀결과: {excel_result}\n")
            f.write("-" * 60 + "\n")
            for ev in events:
                f.write(f"  {ev['type']}: {ev['detail']}\n")
            f.write("=" * 80 + "\n")
        print(f"[프로필비교로그] 저장됨: {PROFILE_COMPARE_LOG_FILE}")
    except Exception as e:
        print(f"[프로필비교로그] 저장 실패: {e}")


def _log_chrome_process_args(user_data_dir: str | None = None):
    """Chrome 프로세스 cmdline 출력. user_data_dir가 있으면 그 경로가 들어 있는 프로세스를 먼저 표시.
    (일반 Chrome이 백그라운드에 떠 있으면 AppData\\...\\User Data 만 보여 오해하기 쉬움)"""
    SUSPECT = ("--incognito", "--guest", "--off-the-record", "--bwsi", "--inprivate")
    hint = None
    if user_data_dir:
        try:
            hint = str(Path(user_data_dir).resolve()).replace("/", "\\").lower()
        except Exception:
            hint = str(user_data_dir).replace("/", "\\").lower()
    print(f"[진단] 이번 실행 Playwright user_data_dir: {user_data_dir or '(알 수 없음)'}")
    if hint:
        print(
            "[진단] 아래 '프로필 일치'에 줄이 없으면, 출력은 대부분 "
            "평소 쓰시는 Chrome(예: …\\Google\\Chrome\\User Data)일 수 있습니다."
        )

    matched, others = [], []

    def _emit_block(pid, name, cmd_str: str, tag: str):
        found = [s for s in SUSPECT if s in cmd_str]
        print(f"[브라우저 프로세스][{tag}] pid={pid} name={name}")
        lim = 900
        print(f"  cmdline: {cmd_str[:lim]}..." if len(cmd_str) > lim else f"  cmdline: {cmd_str}")
        if found:
            print(f"  *** 의심 인자 발견: {found}")
        print("-" * 60)

    try:
        import psutil
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                name = (proc.info.get("name") or "").lower()
                cmdline = proc.info.get("cmdline") or []
                if "chrome" not in name and "chromium" not in name:
                    continue
                cmd_str = " ".join(c or "" for c in cmdline)
                pid = proc.info.get("pid")
                if hint and hint in cmd_str.lower():
                    matched.append((pid, name, cmd_str))
                else:
                    others.append((pid, name, cmd_str))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        if matched:
            print(f"[진단] 프로필 일치(user_data_dir 포함) 프로세스: {len(matched)}개")
            for pid, name, cmd_str in matched:
                _emit_block(pid, name, cmd_str, "프로필 일치")
        else:
            print("[진단] 프로필 일치(user_data_dir 문자열 포함) Chrome 프로세스: 없음")
            print("  → 자동화 창이 아직 뜨기 전이거나, cmdline에 경로가 잘린 경우일 수 있습니다.")
        show_others = 4
        if others:
            print(f"[진단] 기타 Chrome 프로세스(최대 {show_others}개, 참고용):")
            for pid, name, cmd_str in others[:show_others]:
                _emit_block(pid, name, cmd_str, "기타")
        if not matched and not others:
            print("[진단] 실행 중인 chrome.exe / chromium 프로세스를 찾지 못했습니다.")
    except ImportError:
        print("[진단] psutil 미설치 → wmic/PowerShell만 사용합니다. (pip install psutil 권장)")
        try:
            import subprocess
            r = subprocess.run(
                ["wmic", "process", "where", "name like '%chrome%'", "get", "ProcessId,CommandLine", "/format:list"],
                capture_output=True, text=True, timeout=10, creationflags=0x08000000 if sys.platform == "win32" else 0
            )
            if r.returncode == 0 and r.stdout and hint:
                blocks = [b.strip() for b in r.stdout.strip().split("\n\n") if b.strip()]
                hit_blocks = [b for b in blocks if hint in b.lower()]
                if hit_blocks:
                    print("[브라우저 프로세스] wmic (user_data_dir 포함 블록만):")
                    print("\n\n".join(hit_blocks)[:6000])
                else:
                    print("[브라우저 프로세스] wmic 출력(앞부분, user_data_dir 미포함 가능):")
                    print(r.stdout[:2000])
            elif r.returncode == 0 and r.stdout:
                print("[브라우저 프로세스] wmic 출력:")
                print(r.stdout[:2000])
        except Exception as e:
            print(f"[진단] 프로세스 인자 확인 실패: {e}")
    try:
        import subprocess as _sp
        if hint:
            safe = str(user_data_dir).replace("'", "''")
            ps_cmd = (
                f"$d = '{safe}'; "
                "Get-CimInstance Win32_Process -Filter \"name='chrome.exe'\" "
                "| Where-Object { $_.CommandLine -and ($_.CommandLine -like ('*' + $d + '*')) } "
                "| Select-Object ProcessId,CommandLine | Format-List | Out-String -Width 4096"
            )
        else:
            ps_cmd = (
                "Get-CimInstance Win32_Process -Filter \"name='chrome.exe'\" "
                "| Select-Object -First 10 ProcessId,CommandLine | Format-List"
            )
        r2 = _sp.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            capture_output=True,
            text=True,
            timeout=20,
            creationflags=getattr(_sp, "CREATE_NO_WINDOW", 0) if sys.platform == "win32" else 0,
        )
        if r2.returncode == 0 and (r2.stdout or "").strip():
            label = "PowerShell(user_data_dir 일치)" if hint else "PowerShell(일부)"
            print(f"[브라우저 프로세스] {label}:")
            print((r2.stdout or "")[:8000])
        elif hint and r2.stderr:
            print(f"[진단] PowerShell 필터 조회 stderr: {r2.stderr[:500]}")
    except Exception:
        pass


def _get_keyword_for_rank(target_month, target_week, rank):
    """N번 사업자용 N등 키워드 반환 (rank 1~6). 주차별 (target_month, target_week) 기준.
    월별 판매량/트렌드가 높은 세분류 키워드를 Naver API로 선정하여 도매매 검색에 사용."""
    candidates = WEEK_KEYWORD_CANDIDATES.get((target_month, target_week))
    if not candidates:
        fallback = MONTH_KEYWORD_FALLBACK.get(target_month, "인기상품")
        candidates = [fallback]
    idx = min(rank - 1, len(candidates) - 1)
    base_keyword = candidates[idx]

    # 세분류 후보가 있으면 Naver API로 월별 트렌드 가장 높은 키워드 선택
    subcategories = KEYWORD_SUBCATEGORIES.get(base_keyword)
    if subcategories and len(subcategories) >= 2:
        api_keyword = _get_keyword_from_naver_api(
            target_month, target_week, rank, candidates_override=subcategories
        )
        if api_keyword:
            return api_keyword
        # API 실패 시 세분류 첫 번째 사용
        return subcategories[0]

    return base_keyword


def _keyword_to_hashtag_safe(keyword: str) -> str:
    """해시태그에 쓸 수 있게 키워드 정리 (공백→_, #제거 등)."""
    if not keyword:
        return "키워드"
    s = str(keyword).strip().replace(" ", "_").replace("#", "")
    return s if s else "키워드"


def _get_last_week_ymw(target_year: int, target_month: int, target_week: int):
    """직전 주차 (년, 월, 주) 반환. 월별 주차 1~5 가정."""
    if target_week > 1:
        return target_year, target_month, target_week - 1
    if target_month > 1:
        return target_year, target_month - 1, 5
    return target_year - 1, 12, 5


def _month_season(target_month: int) -> str:
    """타깃월(현재+2 반영된 값) → 계절. 3~5 봄 / 6~8 여름 / 9~10 가을 / 11~2 겨울."""
    m = ((int(target_month) - 1) % 12) + 1
    if m in (3, 4, 5):
        return "봄"
    if m in (6, 7, 8):
        return "여름"
    if m in (9, 10):
        return "가을"
    return "겨울"


# ④ 하이브리드: 카테고리×계절 시즌 후보풀(필수·결정적). 네이버 트렌드 재정렬은
# 기존 _get_keyword_from_naver_api 호출부에서 이 시즌 후보 내에서 작동.
# 스포츠/레저는 기존 WEEK_KEYWORD_CANDIDATES(월·주 큐레이션) 유지. 제외 카테고리 키워드 없음.
SEASONAL_POOL = {
    "패션의류": {
        "봄": ["가디건", "트렌치코트", "블라우스", "린넨자켓", "야상", "후드집업", "면바지", "셔츠", "니트가디건", "봄원피스", "슬랙스", "바람막이"],
        "여름": ["반팔티", "린넨바지", "반바지", "시원한원피스", "냉감티", "래쉬가드", "민소매", "7부바지", "쿨링셔츠", "여름원피스", "린넨셔츠", "캐주얼반바지"],
        "가을": ["니트", "가디건", "트렌치코트", "맨투맨", "후드티", "야상", "청자켓", "슬랙스", "가을원피스", "블레이저", "코듀로이", "니트원피스"],
        "겨울": ["패딩", "롱패딩", "숏패딩", "코트", "기모바지", "니트", "목도리", "플리스", "무스탕", "패딩조끼", "기모레깅스", "터틀넥"],
    },
    "디지털/가전": {
        "봄": ["무선이어폰", "보조배터리", "블루투스스피커", "USB허브", "공기청정기", "가습기", "스마트워치", "충전기", "차량용거치대", "키보드", "마우스", "보호필름"],
        "여름": ["USB선풍기", "휴대용선풍기", "미니선풍기", "서큘레이터", "제습기", "무선선풍기", "넥밴드선풍기", "보조배터리", "차량용선풍기", "쿨링패드", "탁상선풍기", "이동식에어컨"],
        "가을": ["가습기", "무선이어폰", "스마트워치", "공기청정기", "블루투스스피커", "보조배터리", "충전기", "노트북파우치", "키보드", "USB허브", "거치대", "케이블"],
        "겨울": ["가습기", "전기요", "온풍기", "미니히터", "전기방석", "손난로", "발난로", "가열식가습기", "전기담요", "USB손난로", "무선이어폰", "보조배터리"],
    },
    "가구/인테리어": {
        "봄": ["커튼", "러그", "수납장", "행거", "화분", "무드등", "침구", "책장", "선반", "정리함", "거울", "옷걸이"],
        "여름": ["인견침구", "시원한이불", "대나무매트", "차렵이불", "제습용품", "수납장", "돗자리", "쿨매트", "발매트", "정리함", "행거", "선반"],
        "가을": ["차렵이불", "러그", "커튼", "무드등", "수납장", "침구", "카펫", "책장", "선반", "행거", "거울", "정리함"],
        "겨울": ["극세사이불", "전기요", "발열매트", "두꺼운커튼", "러그", "온수매트", "카펫", "수납장", "침구세트", "무드등", "발매트", "단열뽁뽁이"],
    },
    "여가/생활편의": {
        "봄": ["캠핑용품", "피크닉매트", "텀블러", "보냉백", "돗자리", "캠핑의자", "휴대용품", "여행용품", "정리수납", "파우치", "물병", "캠핑테이블"],
        "여름": ["보냉백", "아이스박스", "캠핑선풍기", "휴대용선풍기", "물놀이용품", "쿨토시", "보냉가방", "아이스팩", "텀블러", "캠핑의자", "워터저그", "비치용품"],
        "가을": ["캠핑용품", "캠핑의자", "캠핑테이블", "보온병", "텀블러", "랜턴", "피크닉", "등산용품", "보냉백", "파우치", "여행용품", "캐리어"],
        "겨울": ["보온병", "텀블러", "핫팩", "캠핑난로", "손난로", "무릎담요", "캠핑용품", "히터", "캠핑의자", "보온도시락", "방한용품", "캐리어"],
    },
    "생활/건강": {
        "봄": ["마스크", "물티슈", "세탁용품", "섬유유연제", "제습제", "청소용품", "방향제", "위생용품", "수건", "살균스프레이", "정리용품", "탈취제"],
        "여름": ["제습제", "모기퇴치", "해충퇴치", "살충제", "쿨토시", "물티슈", "땀패드", "제습용품", "탈취제", "방충망", "쿨링패치", "휴대용선풍기"],
        "가을": ["마스크", "물티슈", "가습기용품", "세탁용품", "섬유유연제", "방향제", "청소용품", "수건", "위생용품", "핫팩", "탈취제", "정리용품"],
        "겨울": ["핫팩", "손난로", "마스크", "가습기용품", "발열패치", "찜질팩", "온열패치", "물티슈", "핸드크림", "보온용품", "전기방석", "수건"],
    },
    "패션잡화": {
        "봄": ["에코백", "크로스백", "캡모자", "스카프", "선글라스", "양말", "벨트", "백팩", "토트백", "키링", "머플러", "파우치"],
        "여름": ["밀짚모자", "선글라스", "양산", "비치백", "썬캡", "라탄백", "비치슬리퍼", "쿨토시", "발토시", "메쉬백", "부채", "버킷햇"],
        "가을": ["베레모", "머플러", "크로스백", "백팩", "가죽벨트", "스카프", "캡모자", "토트백", "장갑", "양말", "버킷햇", "키링"],
        "겨울": ["목도리", "장갑", "비니", "머플러", "기모양말", "귀마개", "방한모자", "패딩가방", "백팩", "넥워머", "가죽장갑", "털슬리퍼"],
    },
}


def _get_6_keywords_for_category(target_month: int, target_week: int, category: str):
    """회차(카테고리)별 6개 키워드.
    ④ 하이브리드: 스포츠/레저=WEEK_KEYWORD_CANDIDATES(월·주). 그 외=카테고리×계절 SEASONAL_POOL
    에서 주차별 6개(겹침 방지). 시즌풀 없으면 기존 CATEGORY_DEFAULT_KEYWORDS 폴백."""
    if category == "스포츠/레저":
        base6 = WEEK_KEYWORD_CANDIDATES.get((target_month, target_week))
        if not base6:
            fallback = MONTH_KEYWORD_FALLBACK.get(target_month, "인기상품")
            base6 = [fallback] * 6
        return base6[:6]
    season = _month_season(target_month)
    spool = (SEASONAL_POOL.get(category) or {}).get(season)
    if spool and len(spool) >= 6:
        start = ((target_week - 1) * 6) % len(spool)
        return [spool[(start + i) % len(spool)] for i in range(6)]
    # 폴백: 기존 고정 풀(계절 무관) — 시즌풀 누락 시 안전망
    pool = CATEGORY_DEFAULT_KEYWORDS.get(category)
    if pool and len(pool) >= 30:
        week_index = (target_month - 1) * 5 + (target_week - 1)
        slot = week_index % 5
        start = slot * 6
        return [pool[(start + i) % len(pool)] for i in range(6)]
    if pool:
        return pool[:6]
    fallback = MONTH_KEYWORD_FALLBACK.get(target_month, "인기상품")
    return [fallback] * 6


def _build_keywords_30(target_month: int, target_week: int):
    """해당 주차의 1~42등 키워드 (6사업자×7회차). 회차별 카테고리로 6개씩, 겹치지 않음."""
    out = []
    for week_run in range(1, RUNS_PER_WEEK + 1):
        category = RUN_CATEGORY.get(week_run, "스포츠/레저")
        base6 = _get_6_keywords_for_category(target_month, target_week, category)
        for rank in range(1, 7):
            out.append(base6[rank - 1] if rank <= len(base6) else base6[-1])
    return out


def _week_key_key(y: int, m: int, w: int) -> str:
    return f"{y}_{m:02d}_{w}"


def _load_week_keywords(year_week_key: str):
    """저장된 해당 주차 키워드(42개) 로드. 없으면 None."""
    if not WEEK_KEYWORDS_FILE.exists():
        return None
    try:
        data = json.loads(WEEK_KEYWORDS_FILE.read_text(encoding="utf-8"))
        return data.get(year_week_key)
    except Exception:
        return None


def _save_week_keywords(year_week_key: str, keywords_30: list):
    """해당 주차 키워드(42개) 저장 (다음 주차에서 직전주 겹침 판단용)."""
    data = {}
    if WEEK_KEYWORDS_FILE.exists():
        try:
            data = json.loads(WEEK_KEYWORDS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    data[year_week_key] = keywords_30
    WEEK_KEYWORDS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=0), encoding="utf-8")


def _get_keyword_for_run(target_month, target_week, rank, week_run, this_week_30=None, last_week_30=None):
    """회차별 카테고리(패션의류·스포츠/레저·디지털/가전 등)로 6개 키워드 사용.
    - 직전 주차와 같은 슬롯 키워드 겹치면 → 2회차 해당 번(7~12위) 사용.
    - 아니면 해당 회차 1~6번 키워드 사용."""
    # 이번 주 42개가 있으면 슬롯으로 조회 (겹침 시 7~12위)
    if this_week_30 is not None and len(this_week_30) >= 12:
        slot = (week_run - 1) * 6 + (rank - 1)
        default_keyword = this_week_30[slot]
        if last_week_30 and len(last_week_30) > slot and last_week_30[slot] == default_keyword:
            idx_7_12 = 6 + (rank - 1)
            return this_week_30[idx_7_12]
        return default_keyword
    # 단일 키워드 계산: 회차→카테고리→6개 중 rank번
    category = RUN_CATEGORY.get(week_run, "스포츠/레저")
    base6 = _get_6_keywords_for_category(target_month, target_week, category)
    return base6[rank - 1] if rank <= len(base6) else base6[-1]


def build_speedgo_hashtag(kw_tag, target_year, target_month, target_week):
    """스피드고 해시태그 문자열 생성 (domeme 전용 포맷, 다른 스크립트에서도 동일하게 사용)."""
    return f"#{kw_tag}_{target_year}{target_month:02d}w{target_week}"


def get_target_ymw():
    """domeme과 동일: 현재 시점 기준 target_year, target_month, target_week (현재+2월, 오늘 주차)."""
    now = datetime.now()
    target_month = (now.month + 2 - 1) % 12 + 1
    target_year = now.year + (now.month + 2) // 13
    target_week = _get_week_of_month(now)
    return target_year, target_month, target_week


def get_upload_path_from_state():
    """Phase 1과 동일한 (ymw_str, week_run) 반환. 업로드만 할 때 .week_run_state 의 회차 필드 사용.

    본 스크립트 실행 시 `_write_week_run_state` 로 이번에 돈 회차가 기록된다.
    """
    now = datetime.now()
    ymw_str = f"{now.year % 100}년{now.month}월{_get_week_of_month(now)}주차"
    week_run = 1
    if WEEK_RUN_STATE_FILE.exists():
        try:
            raw = WEEK_RUN_STATE_FILE.read_text(encoding="utf-8").strip()
            parts = raw.split(",")
            if len(parts) >= 2 and parts[1].strip().isdigit():
                week_run = min(max(int(parts[1].strip()), 1), RUNS_PER_WEEK)
        except Exception:
            pass
    return ymw_str, week_run


def _speedgo_upload_final_excel(context, page, final_path, kw_tag, target_year, target_month, target_week, rank):
    """run_all_steps 후: 스피드고 접속 → 해시태그 검색 → 200개 보기 → 엑셀업로드 클릭 → 파일선택 → _최종.xlsx 업로드. 전송은 별도."""
    speedgo_hash = build_speedgo_hashtag(kw_tag, target_year, target_month, target_week)
    # 200개씩 보기 페이지 URL
    mb_save_list_url = (
        "https://speedgo.domeggook.com/mybox/mb_saveList.php?"
        f"pagenum=&hashTag={quote(speedgo_hash, safe='')}&sf=subject&sw=&itemNos=&mnp=&mxp="
        "&titleStatus=&editStatus=&useOption=&sender_date1=&sender_date2="
        "&sort1=&sort2=&sort3=&sort4=&sort5=&b2bStatus=0&pageLimit=200"
    )
    # 1) 스피드고 탭 찾기 또는 새 탭
    speedgo_page = None
    for p in context.pages:
        try:
            if "speedgo" in (p.url or "").lower():
                speedgo_page = p
                break
        except Exception:
            pass
    if speedgo_page is None:
        speedgo_page = context.new_page()
    speedgo_page.bring_to_front()
    # 2) 스피드고 마이박스 접속 후 해시태그 검색
    work_page = speedgo_page
    work_page.goto(SPEEDGO_URL, wait_until=_WAIT, timeout=30000)
    time.sleep(_S(0.6, 1.2))
    # 마이박스 메뉴 클릭
    try:
        mybox = speedgo_page.get_by_role("link", name="마이박스").first
        if mybox.count() == 0:
            mybox = speedgo_page.locator('a:has-text("마이박스")').first
        if mybox.count() > 0 and mybox.is_visible():
            mybox.click()
            time.sleep(_S(0.6, 1.2))
    except Exception:
        pass
    # 해시태그 입력란: "해시태그는 #태그명으로 입력하세요" placeholder
    try:
        inp = speedgo_page.locator('input[placeholder*="해시태그"], input[placeholder*="#태그명"]').first
        if inp.count() > 0 and inp.is_visible():
            inp.fill(speedgo_hash)
            print(f"[{rank}번] 해시태그 입력: {speedgo_hash}")
    except Exception:
        pass
    time.sleep(_S(0.3, 0.6))
    # 검색 버튼 클릭
    try:
        search_btn = speedgo_page.locator('button:has-text("검색"):not(:has-text("초기화"))').first
        if search_btn.count() == 0:
            search_btn = speedgo_page.get_by_role("button", name="검색").first
        if search_btn.count() > 0 and search_btn.is_visible():
            search_btn.click()
            speedgo_page.wait_for_load_state(_WAIT, timeout=15000)
            time.sleep(_S(0.6, 1.2))
    except Exception:
        pass
    # 3) 200개씩 보기 페이지로 이동
    work_page = speedgo_page
    work_page.goto(mb_save_list_url, wait_until=_WAIT, timeout=30000)
    time.sleep(_S(0.8, 1.5))
    try:
        speedgo_page.wait_for_selector('text=상품목록', state="visible", timeout=15000)
    except Exception:
        pass
    time.sleep(_S(0.4, 0.8))
    # 4) 1차 엑셀업로드 버튼 클릭 (상품목록 옆 → excelMbUpload() 호출, 조그만 창 뜸)
    try:
        excel_up_btn = speedgo_page.locator('#mbUploadBtn').first
        if excel_up_btn.count() == 0:
            excel_up_btn = speedgo_page.locator('button[onclick*="excelMbUpload"]').first
        if excel_up_btn.count() == 0:
            excel_up_btn = speedgo_page.get_by_text("엑셀업로드").first
        if excel_up_btn.count() > 0 and excel_up_btn.is_visible():
            excel_up_btn.click()
            print(f"[{rank}번] 1차 엑셀업로드 클릭 (#mbUploadBtn)")
        else:
            raise RuntimeError("엑셀업로드 버튼 없음")
    except Exception as e:
        print(f"[{rank}번] 엑셀업로드 버튼 클릭 실패: {e}")
        return
    time.sleep(_S(0.8, 1.5))
    # 5) 2차: 조그맣게 뜬 창에서 "파일 선택" 버튼 클릭 → 파일 선택기에서 _최종.xlsx 지정
    try:
        with speedgo_page.expect_file_chooser(timeout=12000) as fc:
            file_sel_btn = speedgo_page.get_by_text("파일 선택").first
            if file_sel_btn.count() == 0:
                file_sel_btn = speedgo_page.get_by_role("button", name="파일 선택").first
            if file_sel_btn.count() > 0 and file_sel_btn.is_visible():
                file_sel_btn.click()
            else:
                raise RuntimeError("파일 선택 버튼 없음")
        fc.value.set_files(str(final_path))
        print(f"[{rank}번] _최종.xlsx 파일 선택 완료: {final_path.name}")
    except Exception as e:
        try:
            file_input = speedgo_page.locator('input[type="file"]').first
            if file_input.count() > 0:
                file_input.set_input_files(str(final_path))
                print(f"[{rank}번] _최종.xlsx 설정 (input[file] 직접): {final_path.name}")
            else:
                raise e
        except Exception as e2:
            print(f"[{rank}번] 파일 선택/업로드 실패: {e2}")
            return
    time.sleep(_S(1, 2))
    # 6) 업로드 실행 (엑셀업로드 제출 또는 업로드 버튼)
    for btn_text in ("엑셀업로드", "업로드", "적용", "확인"):
        try:
            b = speedgo_page.get_by_role("button", name=btn_text).first
            if b.count() == 0:
                b = speedgo_page.locator(f'button:has-text("{btn_text}")').first
            if b.count() > 0 and b.is_visible():
                b.click()
                print(f"[{rank}번] 업로드 실행 클릭: {btn_text}")
                break
        except Exception:
            continue
    time.sleep(_S(2, 4))
    print(f"[{rank}번] 스피드고 _최종.xlsx 업로드 요청 완료 (전송은 별도)")


def _login_submit_contexts(page, form_ctx):
    """로그인 폼이 있는 프레임 우선, 이후 메인·나머지 iframe 순으로 검색."""
    out, seen = [], set()

    def _add(ctx):
        if ctx is None:
            return
        k = id(ctx)
        if k in seen:
            return
        seen.add(k)
        out.append(ctx)

    _add(form_ctx)
    _add(page)
    try:
        for fr in page.frames:
            _add(fr)
    except Exception:
        pass
    return out


def _click_domeggook_login_submit(page, form_ctx) -> bool:
    """도매꾹 통합 로그인 등: '로그인하기' 제출(텍스트 우선, generic submit은 마지막)."""
    for ctx in _login_submit_contexts(page, form_ctx):
        for name in ("로그인하기",):
            try:
                btn = ctx.get_by_role("button", name=name, exact=True).first
                if btn.count() > 0 and btn.is_visible():
                    btn.click(timeout=15000)
                    return True
            except Exception:
                try:
                    btn = ctx.get_by_role("button", name=name, exact=True).first
                    if btn.count() > 0 and btn.is_visible():
                        btn.click(force=True, timeout=15000)
                        return True
                except Exception:
                    pass
        # 문구만 '로그인' 인 녹색 버튼·공백 변형
        try:
            btn = ctx.get_by_role("button", name=re.compile(r"^\s*로그인\s*$", re.I)).first
            if btn.count() > 0 and btn.is_visible():
                btn.click(timeout=15000)
                return True
        except Exception:
            try:
                btn = ctx.get_by_role("button", name=re.compile(r"로그인\s*하기", re.I)).first
                if btn.count() > 0 and btn.is_visible():
                    btn.click(timeout=15000)
                    return True
            except Exception:
                pass
        for sel in (
            'button:has-text("로그인하기")',
            'input[type="submit"][value="로그인하기"]',
            'button[type="submit"]:has-text("로그인하기")',
            'a:has-text("로그인하기")',
        ):
            try:
                loc = ctx.locator(sel).first
                if loc.count() > 0 and loc.is_visible():
                    loc.click(timeout=15000)
                    return True
            except Exception:
                try:
                    loc = ctx.locator(sel).first
                    if loc.count() > 0 and loc.is_visible():
                        loc.click(force=True, timeout=15000)
                        return True
                except Exception:
                    continue
        # 구 도매매: '로그인'만 있는 제출 버튼 (다른 소셜 버튼과 구분)
        for sel in (
            'button[type="submit"]:has-text("로그인"):not(:has-text("하기"))',
            'button:has-text("로그인"):not(:has-text("하기")):not(:has-text("네이버")):not(:has-text("카카오"))',
            'input[type="submit"][value="로그인"]',
        ):
            try:
                loc = ctx.locator(sel).first
                if loc.count() > 0 and loc.is_visible():
                    loc.click(timeout=15000)
                    return True
            except Exception:
                continue
        try:
            loc = ctx.locator('button[type="submit"]').first
            if loc.count() > 0 and loc.is_visible():
                loc.click(timeout=15000)
                return True
        except Exception:
            pass
    return False


def _domeme_login_success_contexts(page):
    """메인 문서 + 모든 iframe(로그아웃 링크가 프레임 안에만 있을 수 있음)."""
    out, seen = [], set()
    for ctx in [page, *list(getattr(page, "frames", []) or ())]:
        if ctx is None:
            continue
        k = id(ctx)
        if k in seen:
            continue
        seen.add(k)
        out.append(ctx)
    return out


def _domeme_page_looks_logged_in(page) -> bool:
    """한 시점에서 로그인된 UI인지(프레임·본문 텍스트·domemedb URL 보조)."""
    for ctx in _domeme_login_success_contexts(page):
        try:
            lo = ctx.get_by_role("link", name="로그아웃").first
            if lo.count() > 0 and lo.is_visible():
                return True
        except Exception:
            pass
        try:
            if ctx.evaluate(
                """() => !!(
                    document.querySelector('a[href*="logout"], a[href*="Logout"], a[href*="log_out"]')
                    || (document.body && /로그\\s*아웃/.test(document.body.innerText || ''))
                )"""
            ):
                return True
        except Exception:
            pass
    try:
        u = (page.url or "").lower()
        if "domemedb.domeggook.com" in u and "mem_formlogin" not in u and "login=pc" not in u:
            if page.evaluate(
                """() => !!(document.body && (
                    document.querySelector('a[href*="logout"], a[href*="Logout"]')
                    || (document.body.innerText || '').includes('로그아웃')
                ))"""
            ):
                return True
    except Exception:
        pass
    return False


def _wait_logged_in_domeme(page, timeout_ms: int = 28000) -> bool:
    """도매매(또는 도매꾹 계열) 로그인 후 로그아웃 노출·URL 변화로 성공 판별(프레임 포함)."""
    deadline = time.time() + timeout_ms / 1000.0
    step = 0.35 if FAST_MODE else 0.45
    while time.time() < deadline:
        if _domeme_page_looks_logged_in(page):
            return True
        time.sleep(step)
    return False


def _domeme_wait_login_success_any_tab(context, page, timeout_ms: int = 28000):
    """제출 후 현재 탭 또는 새로 연 탭에서 로그인 성공을 기다림. 성공 시 해당 Page, 실패 시 None.

    ★중요: 도매매 로그인이므로 'domemedb.domeggook.com' 탭을 최우선으로 본다.
    이전 사업자의 stale 스피드고 마이박스 탭이 로그인된 것처럼 보여 그 탭을 success_page 로
    반환하면, 이후 검색이 스피드고 탭에 입력되어 '도매매 로그인을 해주세요' alert → 0건 →
    _최종.xlsx 미생성으로 이어졌다(2~6번 동시 실패 원인). 그래서 우선순위를 둔다:
      1순위: domemedb 탭이면서 logged-in
      2순위: 인자로 받은 현재 page 가 logged-in
      (speedgo·기타 탭은 domeme 로그인 성공 판정 근거로 쓰지 않는다)
    """
    deadline = time.time() + timeout_ms / 1000.0
    step = 0.35 if FAST_MODE else 0.45
    while time.time() < deadline:
        try:
            plist = [p for p in list(context.pages) if not p.is_closed()]
        except Exception:
            plist = [page]
        # 1순위: domemedb 탭
        for p in plist:
            try:
                u = (p.url or "").lower()
                if "domemedb.domeggook.com" in u and _domeme_page_looks_logged_in(p):
                    return p
            except Exception:
                pass
        # 2순위: 현재 작업 page 자체(도매매 도메인일 때만)
        try:
            pu = (page.url or "").lower()
            if "domeggook.com" in pu and "speedgo" not in pu and _domeme_page_looks_logged_in(page):
                return page
        except Exception:
            pass
        time.sleep(step)
    return None


def _domeme_logout_if_logged_in(page, max_rounds: int = 3) -> None:
    """프로필·새 탭에 남은 로그인 세션을 끊는다. 1번 사업자도 동일(기존엔 rank>1 만 로그아웃)."""
    for _ in range(max_rounds):
        clicked = False
        try:
            lo = page.get_by_role("link", name="로그아웃").first
            if lo.count() > 0 and lo.is_visible():
                lo.click(timeout=12000, force=True)
                clicked = True
        except Exception:
            pass
        if not clicked:
            try:
                if page.evaluate(
                    """() => {
                        const a = document.querySelector(
                          'a[href*="Logout"], a[href*="logout"], a[href*="log_out"], a[href*="LOGOUT"]'
                        );
                        if (!a) return false;
                        a.click();
                        return true;
                    }"""
                ):
                    clicked = True
            except Exception:
                pass
        if not clicked:
            break
        time.sleep(1.0)
        try:
            page.wait_for_load_state("domcontentloaded", timeout=20000)
        except Exception:
            pass
        _domeme_focus_window_for_navigation(page)


def _domeme_find_login_fields_speedgo_phase_style(page):
    """Phase 2 파일 `test_speedgo_upload_1번.find_login_form` 과 동일: 메인 Page + frames[1:] 만 검색."""
    id_selectors = [
        'input[name="userId"]',
        'input[name="user_id"]',
        'input[name="id"]',
        'input[name="loginId"]',
        'input[id*="user"]',
        'input[id*="id"]',
        'input[type="text"]',
        "input#userId",
        "input#user_id",
        'input[name="member_id"]',
        'input[name="mb_id"]',
        'input[type="email"]',
        'input[type="tel"]',
    ]
    pw_selectors = [
        'input[name="password"]',
        'input[name="passwd"]',
        'input[name="pw"]',
        'input[type="password"]',
    ]
    contexts = [page]
    try:
        frames = list(getattr(page, "frames", []) or [])
        if len(frames) > 1:
            contexts.extend(frames[1:])
    except Exception:
        pass
    for ctx in contexts:
        for sel in id_selectors:
            try:
                el = ctx.locator(sel).first
                if el.count() > 0 and el.is_visible():
                    for ps in pw_selectors:
                        try:
                            pe = ctx.locator(ps).first
                            if pe.count() > 0 and pe.is_visible():
                                return el, pe, ctx
                        except Exception:
                            continue
            except Exception:
                continue
    return None, None, None


def _domeme_login_speedgo_phase_style(page, user_id: str, password: str, log_prefix: str) -> bool:
    """
    Phase 2(`test_speedgo_upload_1번.main_upload_impl`)과 동일 사용자 동선.
    domemedb index → 헤더「로그인」→ id/pw fill → 제출 시 expect_navigation(domcontentloaded).

    Phase 1 이전 구현은 mem_formLogin 직접 진입(`_domeme_goto_login_pc`)이라 Referer·쿠키·폼 DOM 이
    달라질 수 있어, 잘 되는 Phase 2 와 불일치할 수 있음 → 본 경로를 먼저 시도.
    """
    _domeme_dismiss_blocking_layers(page)
    _domeme_focus_window_for_navigation(page)
    try:
        page.goto(_DOMEME_GOTO_URL, wait_until="commit", timeout=45000)
        time.sleep(_S(0.35, 1.0))
        page.wait_for_load_state("domcontentloaded", timeout=20000)
    except Exception as e:
        print(f"{log_prefix}[Phase2동선] domemedb index 실패: {e!r}", flush=True)
        return False
    time.sleep(_S(0.4, 1.0))
    try:
        login_link = page.get_by_role("link", name="로그인").first
        if login_link.count() == 0:
            login_link = page.locator('a:has-text("로그인"):not(:has-text("진행중"))').first
        if login_link.count() == 0:
            login_link = page.get_by_text("로그인", exact=True).first
        if login_link.count() > 0 and login_link.is_visible():
            login_link.click(timeout=5000)
            print(f"{log_prefix}[Phase2동선] 헤더「로그인」클릭", flush=True)
    except Exception as e:
        print(f"{log_prefix}[Phase2동선] 로그인 링크: {e!r}", flush=True)
    time.sleep(_S(0.8, 1.5))
    id_field, pw_field, form_ctx = _domeme_find_login_fields_speedgo_phase_style(page)
    if not id_field or not pw_field:
        print(f"{log_prefix}[Phase2동선] 폼 미검출", flush=True)
        return False
    try:
        id_field.fill(user_id, timeout=15000)
        time.sleep(_S(0.12, 0.35))
        pw_field.fill(password, timeout=15000)
    except Exception as e:
        print(f"{log_prefix}[Phase2동선] fill 실패: {e!r}", flush=True)
        return False
    time.sleep(_S(0.15, 0.4))
    ctx = form_ctx or page
    nav_until = "domcontentloaded"
    for sel in (
        'button[type="submit"]',
        'input[type="submit"]',
        'a:has-text("로그인")',
        'button:has-text("로그인")',
        ".btn_login",
        "#btn_login",
    ):
        try:
            btn = ctx.locator(sel).first
            if btn.count() > 0 and btn.is_visible():
                try:
                    with page.expect_navigation(wait_until=nav_until, timeout=20000):
                        btn.click(timeout=5000)
                except Exception:
                    try:
                        btn.click(timeout=5000)
                    except Exception:
                        continue
                print(f"{log_prefix}[Phase2동선] 로그인 제출 완료: {user_id}", flush=True)
                return True
        except Exception:
            continue
    try:
        pw_field.press("Enter")
        print(f"{log_prefix}[Phase2동선] Enter 로 제출 시도", flush=True)
        return True
    except Exception:
        return False


def _domeme_goto_login_pc(work_page, log_prefix: str) -> None:
    """헤더와 동일: domeme…/mem_formLogin.php?back=… (실제 로그인 폼). 실패 시 ?login=pc 폴백."""
    member_url = _domeme_member_login_url()
    if not _domeme_goto_load(work_page, member_url, log_prefix):
        print(f"{log_prefix}mem_formLogin load 실패 → ?login=pc load 시도")
        if not _domeme_goto_load(work_page, DOMEME_LOGIN_URL, log_prefix):
            print(f"{log_prefix}?login=pc 실패 → 다중 URL 폴백")
            _domeme_navigate_with_fallbacks(work_page, log_prefix)
    time.sleep(_S(0.35, 0.75))
    _domeme_focus_window_for_navigation(work_page)
    _domeme_dismiss_blocking_layers(work_page)
    time.sleep(_S(0.25, 0.55))
    try:
        work_page.wait_for_load_state("domcontentloaded", timeout=15000)
    except Exception:
        pass
    _domeme_wait_login_password_visible(work_page, log_prefix)
    _domeme_focus_window_for_navigation(work_page)


def _domeme_click_open_login_modal(page, log_prefix: str) -> None:
    """이미 로그인 폼이 없을 때만 헤더 등에서 '로그인' 클릭."""
    try:
        for ctx in _domeme_login_contexts(page):
            pw = ctx.locator("input[type='password']").first
            if pw.count() > 0 and pw.is_visible():
                return
    except Exception:
        pass
    openers = [
        lambda: page.locator('a[href*="mem_formLogin.php"]').first,
        lambda: page.get_by_role("link", name="로그인").first,
        lambda: page.locator('header a:has-text("로그인")').first,
        lambda: page.locator('a:has-text("로그인"):not(:has-text("진행중"))').first,
        lambda: page.get_by_text("로그인", exact=True).first,
    ]
    for mk in openers:
        try:
            loc = mk()
            if loc.count() > 0 and loc.is_visible():
                loc.click(timeout=15000)
                time.sleep(_S(0.5, 1.2))
                return
        except Exception as e:
            print(f"{log_prefix}로그인 링크/버튼 시도: {e}")
    print(f"{log_prefix}[안내] '로그인' 클릭 없이 진행 (?login=pc 또는 이미 폼 표시)")


def _domeme_login_contexts(page):
    """메인 문서 + 모든 iframe(중복 id 제거)."""
    seen, out = set(), []
    for ctx in [page, *list(getattr(page, "frames", []) or ())]:
        if ctx is None:
            continue
        k = id(ctx)
        if k in seen:
            continue
        seen.add(k)
        out.append(ctx)
    return out


def _domeme_dismiss_blocking_layers(page) -> None:
    """공지·쿠키 등 상단 레이어가 입력칸을 가리면 닫기 시도."""
    for txt in ("닫기", "오늘 하루 안보기", "하루동안 창 보이지않기", "창 닫기", "동의하고 계속"):
        try:
            loc = page.get_by_role("button", name=txt).first
            if loc.count() > 0 and loc.is_visible():
                loc.click(timeout=2500)
                time.sleep(0.2)
        except Exception:
            pass
        try:
            loc = page.locator(f'a:has-text("{txt}")').first
            if loc.count() > 0 and loc.is_visible():
                loc.click(timeout=2500)
                time.sleep(0.2)
        except Exception:
            pass


def _domeme_wait_login_password_visible(page, log_prefix: str, total_sec: float = 28.0) -> None:
    """SPA·iframe 지연 후 비밀번호 칸이 보일 때까지 대기."""
    t0 = time.time()
    while time.time() - t0 < total_sec:
        for ctx in _domeme_login_contexts(page):
            try:
                pw = ctx.locator("input[type='password']").first
                if pw.count() > 0 and pw.is_visible():
                    return
            except Exception:
                pass
        time.sleep(0.35)
    print(f"{log_prefix}[안내] 비밀번호 입력칸 대기 타임아웃({total_sec}s) — 그대로 폼 탐색 시도")


def _find_login_form_pair(page):
    """같은 form 안의 아이디·비번 쌍(검색창 등 다른 text input 오인식 방지)."""
    id_order = [
        'input[name="userId"]',
        'input[name="userid"]',
        'input[name="UserId"]',
        'input[name="user_id"]',
        'input[name="member_id"]',
        'input[name="mb_id"]',
        'input[name="strMemberId"]',
        'input[name="strMemberID"]',
        'input[name="loginId"]',
        'input[name="login_id"]',
        'input[name="id"]',
        'input[id="userId"]',
        'input[id="member_id"]',
        'input[id*="userId"]',
        'input[id*="member"]',
        'input[id*="login"]',
        'input[autocomplete="username"]',
        'input[placeholder*="아이디"]',
        'input[type="email"]',
        'input[type="tel"]',
        'input[type="text"]',
    ]
    for ctx in _domeme_login_contexts(page):
        try:
            form = ctx.locator("form:has(input[type='password'])").first
            if form.count() == 0:
                continue
            pw = form.locator("input[type='password']").first
            if pw.count() == 0:
                continue
            try:
                if not pw.is_visible():
                    continue
            except Exception:
                continue
            for sel in id_order:
                cand = form.locator(sel).first
                if cand.count() == 0:
                    continue
                try:
                    if not cand.is_visible():
                        continue
                except Exception:
                    continue
                hint = ""
                try:
                    hint = (cand.get_attribute("name") or "") + " " + (cand.get_attribute("id") or "")
                except Exception:
                    hint = ""
                if re.search(r"search|keyword|query|\bsw\b", hint, re.I):
                    continue
                return cand, pw, ctx
        except Exception:
            continue
    for ctx in _domeme_login_contexts(page):
        try:
            n = ctx.locator("input[type='password']").count()
            for i in range(min(n, 8)):
                pw = ctx.locator("input[type='password']").nth(i)
                try:
                    if not pw.is_visible():
                        continue
                except Exception:
                    continue
                anc = pw.locator("xpath=./ancestor::form[1]")
                if anc.count() == 0:
                    continue
                for sel in id_order:
                    cand = anc.locator(sel).first
                    if cand.count() == 0:
                        continue
                    try:
                        if not cand.is_visible():
                            continue
                    except Exception:
                        continue
                    hint = ""
                    try:
                        hint = (cand.get_attribute("name") or "") + " " + (cand.get_attribute("id") or "")
                    except Exception:
                        hint = ""
                    if re.search(r"search|keyword|query|\bsw\b", hint, re.I):
                        continue
                    return cand, pw, ctx
        except Exception:
            continue
    # 3) <form> 없이 div/section 안에만 비밀번호·아이디가 있는 경우(구조 변경 대비)
    for ctx in _domeme_login_contexts(page):
        try:
            roots = ctx.locator(
                "div:has(input[type='password']), section:has(input[type='password'])"
            )
            rc = min(roots.count(), 8)
            for ri in range(rc):
                box = roots.nth(ri)
                pw = box.locator("input[type='password']").first
                if pw.count() == 0:
                    continue
                try:
                    if not pw.is_visible():
                        continue
                except Exception:
                    continue
                pool = box.locator(
                    "input[type='text'], input[type='email'], input[type='tel'], input:not([type])"
                )
                pc = min(pool.count(), 10)
                for j in range(pc):
                    cand = pool.nth(j)
                    try:
                        if not cand.is_visible():
                            continue
                    except Exception:
                        continue
                    hint = ""
                    try:
                        hint = (cand.get_attribute("name") or "") + " " + (cand.get_attribute("id") or "")
                    except Exception:
                        hint = ""
                    if re.search(
                        r"search|keyword|query|\bsw\b|save|remember|auto|chk|captcha",
                        hint,
                        re.I,
                    ):
                        continue
                    return cand, pw, ctx
        except Exception:
            continue
    return None, None, None


def _domeme_fill_login_field(loc, value: str, log_prefix: str, field_label: str) -> bool:
    """클릭 후 fill(force)·press_sequentially·JS value 주입 순."""
    try:
        loc.scroll_into_view_if_needed(timeout=8000)
    except Exception:
        pass
    last_err = None
    for use_force in (False, True):
        try:
            loc.click(timeout=8000, force=use_force)
            time.sleep(0.1)
            try:
                loc.fill("", timeout=4000)
            except Exception:
                pass
            loc.fill(value, timeout=15000, force=use_force)
            return True
        except Exception as e:
            last_err = e
    print(f"{log_prefix}{field_label} fill 실패({last_err!r}) → press_sequentially")
    try:
        loc.click(timeout=8000, force=True)
        loc.press("Control+a")
        time.sleep(0.05)
        loc.press("Backspace")
        loc.press_sequentially(value, delay=30)
        return True
    except Exception as e2:
        print(f"{log_prefix}{field_label} press_sequentially 실패: {e2!r}")
    try:
        loc.evaluate(
            """(el, v) => {
                el.focus();
                el.value = v;
                el.dispatchEvent(new Event("input", { bubbles: true }));
                el.dispatchEvent(new Event("change", { bubbles: true }));
            }""",
            value,
        )
        return True
    except Exception as e3:
        print(f"{log_prefix}{field_label} JS value 주입 실패: {e3!r}")
    return False


def _domeme_click_mybox_add(page) -> bool:
    """도매매 검색결과 하단 '마이박스담기' 클릭. 가림(disabled/오버레이) 대비 스크롤·force·JS 폴백."""
    try:
        page.evaluate("window.scrollTo(0, Math.max(document.body.scrollHeight, document.documentElement.scrollHeight))")
    except Exception:
        pass
    time.sleep(_S(0.2, 0.45))

    def _contexts():
        out, seen = [], set()
        for c in (page, *(getattr(page, "frames", None) or ())):
            if c is None:
                continue
            k = id(c)
            if k in seen:
                continue
            seen.add(k)
            out.append(c)
        return out

    def _try_ctx(ctx) -> bool:
        candidates = [
            lambda: ctx.get_by_role("button", name="마이박스담기"),
            lambda: ctx.get_by_role("link", name="마이박스담기"),
            lambda: ctx.locator("button, a, [role='button']").filter(has_text="마이박스담기"),
            lambda: ctx.locator('button:has-text("마이박스담기")'),
            lambda: ctx.locator('a:has-text("마이박스담기")'),
            lambda: ctx.locator('.footer_position_btn1:has-text("마이박스담기")'),
            lambda: ctx.locator('[class*="footer"] button:has-text("마이박스담기")'),
            lambda: ctx.locator('button.footer_position_btn1'),
        ]
        for mk in candidates:
            try:
                loc = mk().first
                if loc.count() == 0:
                    continue
                loc.scroll_into_view_if_needed(timeout=8000)
                try:
                    loc.click(timeout=12000)
                except Exception:
                    loc.click(timeout=12000, force=True)
                return True
            except Exception:
                continue
        return False

    for ctx in _contexts():
        if _try_ctx(ctx):
            return True

    try:
        return bool(
            page.evaluate(
                """() => {
                    const hit = (root) => {
                        const sel = 'button, a, input[type="button"], input[type="submit"], [role="button"]';
                        for (const el of root.querySelectorAll(sel)) {
                            const raw = (el.textContent || el.innerText || el.value || '').replace(/\\s+/g, ' ').trim();
                            if (!raw) continue;
                            if (raw.includes('마이박스담기') || (raw.includes('마이박스') && raw.includes('담기'))) {
                                el.scrollIntoView({ block: 'center', inline: 'nearest' });
                                el.click();
                                el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
                                return true;
                            }
                        }
                        return false;
                    };
                    if (hit(document)) return true;
                    for (const fr of document.querySelectorAll('iframe')) {
                        try {
                            const d = fr.contentDocument;
                            if (d && hit(d)) return true;
                        } catch (e) {}
                    }
                    return false;
                }"""
            )
        )
    except Exception:
        return False


def _domeme_pick_page_for_work(context, fallback):
    """context.pages 에서 고르지 않는다. 호출부가 넘긴 work_page(fallback)만 사용."""
    return fallback


def _ensure_domeme_ready_for_login(work_page, rank: int) -> bool:
    """로그인 전까지 domemedb 진입. Phase 2(test_speedgo)와 동일하게 commit → domcontentloaded 우선."""
    tag = "[시작]" if rank == 0 else f"[{rank}번]"
    _goto_ms = 120_000
    for attempt in range(1, 4):
        _domeme_focus_window_for_navigation(work_page)
        try:
            u = work_page.url or ""
        except Exception:
            u = ""
        if _url_is_domeme_session(u):
            try:
                work_page.wait_for_load_state("domcontentloaded", timeout=35000)
            except Exception:
                pass
            return True
        print(f"{tag} 도매매 접속 시도 {attempt}/3… 현재 url={repr(u)}", flush=True)
        _domeme_debug_dump_pages(work_page.context, f"{tag}")
        err = None
        try:
            work_page.set_default_navigation_timeout(_goto_ms)
            work_page.set_default_timeout(_goto_ms)
        except Exception:
            pass
        try:
            work_page.goto(_DOMEME_GOTO_URL, wait_until="commit", timeout=45000)
            time.sleep(_S(0.25, 0.55))
            work_page.wait_for_load_state("domcontentloaded", timeout=20000)
        except Exception as e:
            err = e
            try:
                work_page.goto(
                    "https://domemedb.domeggook.com/index/",
                    wait_until="domcontentloaded",
                    timeout=_goto_ms,
                )
            except Exception as e2:
                err = e2
        try:
            u2 = (work_page.url or "").strip()
        except Exception as e:
            u2 = ""
            if err is None:
                err = e
        print(f"{tag} work_page.goto(Phase2식) 직후 work_page.url={repr(u2)}", flush=True)
        ok = _url_is_domeme_session(u2)
        _domeme_log_navigation_assert(ok, u2, f"{tag} ", f"ensure 시도{attempt}")
        if ok:
            return True
        if _domeme_url_is_blankish(u2):
            print(f"{tag} domemedb 미도착(blank 유지 가능) → domcontentloaded 로 1회 더", flush=True)
            try:
                work_page.goto(
                    "https://domemedb.domeggook.com/index/",
                    wait_until="domcontentloaded",
                    timeout=_goto_ms,
                )
            except Exception as e2:
                err = e2
            try:
                u3 = (work_page.url or "").strip()
            except Exception:
                u3 = ""
            print(f"{tag} 재 goto 직후 work_page.url={repr(u3)}", flush=True)
            if _url_is_domeme_session(u3):
                _domeme_log_navigation_assert(True, u3, f"{tag} ", "ensure 재goto")
                return True
        _domeme_print_page_nav_diagnostics(work_page, err, f"{tag}")
        time.sleep(1.5)
    try:
        tail = repr(work_page.url)
    except Exception:
        tail = "?"
    print(f"{tag} [오류] 도매매 URL로 전환되지 않았습니다. url={tail}")
    return False


# === Phase 1 전용 프로필 + 원격디버깅 CDP (Phase 3 와 동일 방식) ===
# 실제 Chrome 프로필 robocopy 복사를 폐기 → 개인 Chrome 켜져 있어도 '프로필복사실패' 없음.
# 도매매 로그인은 ID/PW 라 전용 프로필이면 충분(쿠키는 이 폴더에 누적 유지).
_P1_CDP_DIR = PROJECT_DIR / "chrome_phase1_cdp"
_P1_CDP_PORT = int(os.environ.get("PHASE1_CDP_PORT", "9223"))


def _p1_kill_debug_chrome() -> int:
    """이 전용 user-data-dir 를 쓰는 chrome 만 종료 (사용자 일반 Chrome 미접촉)."""
    marker = str(_P1_CDP_DIR).lower()
    try:
        import psutil
    except ImportError:
        return 0
    n = 0
    for p in psutil.process_iter(["pid", "name", "cmdline"]):
        try:
            if "chrome" not in (p.info.get("name") or "").lower():
                continue
            cmd = " ".join(str(c or "") for c in (p.info.get("cmdline") or [])).lower()
            if marker in cmd:
                p.kill()
                n += 1
        except Exception:
            continue
    return n


def _p1_wait_cdp(port: int, timeout: int = 40) -> bool:
    import urllib.request
    url = f"http://127.0.0.1:{port}/json/version"
    end = time.time() + timeout
    while time.time() < end:
        try:
            with urllib.request.urlopen(url, timeout=3) as r:
                if r.status == 200:
                    return True
        except Exception:
            time.sleep(1)
    return False


def _p1_launch_debug_chrome():
    import subprocess
    chrome = CHROME_EXECUTABLE if os.path.isfile(CHROME_EXECUTABLE) else "chrome"
    _P1_CDP_DIR.mkdir(parents=True, exist_ok=True)
    for lk in ("SingletonLock", "SingletonSocket", "SingletonCookie", "lockfile"):
        try:
            (_P1_CDP_DIR / lk).unlink(missing_ok=True)
        except Exception:
            pass
    pd = (PHASE2_CHROME_PROFILE_DIR or "Profile 67").strip()
    for sub in ("Default", pd):
        try:
            pj = _P1_CDP_DIR / sub / "Preferences"
            if pj.exists():
                t = pj.read_text(encoding="utf-8", errors="ignore").replace(
                    '"exit_type":"Crashed"', '"exit_type":"Normal"')
                pj.write_text(t, encoding="utf-8")
        except Exception:
            pass
    return subprocess.Popen([
        chrome, f"--remote-debugging-port={_P1_CDP_PORT}",
        f"--user-data-dir={_P1_CDP_DIR}", f"--profile-directory={pd}",
        "--start-maximized", "--no-first-run", "--no-default-browser-check",
        "about:blank",
    ])


def main():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("playwright가 설치되지 않았습니다.")
        print("실행: pip install playwright")
        print("그 다음: playwright install chromium")
        sys.exit(1)

    # Phase 3 와 동일 방식: 전용 프로필 + 원격디버깅 CDP. 실제 Chrome 프로필 robocopy 복사 폐기.
    # → 개인 Chrome 켜져 있어도 '프로필복사실패' 없음. 도매매 로그인은 ID/PW 라 전용 프로필이면 충분.
    user_data_dir = _P1_CDP_DIR
    print(
        f"[실행 설정] Phase1 = Phase3 방식: 전용 프로필({_P1_CDP_DIR.name}) + 원격디버깅 CDP(port {_P1_CDP_PORT}). "
        "실제 프로필 복사 안 함 → 개인 Chrome 켜져 있어도 무방.",
        flush=True,
    )

    now = datetime.now()
    target_month = (now.month + 2 - 1) % 12 + 1  # 현재+2 (3월→5월)
    target_year = now.year + (now.month + 2) // 13
    target_week = _get_week_of_month(now)
    ymw_str = f"{now.year % 100}년{now.month}월{_get_week_of_month(now)}주차"
    year_week_key = _week_key_key(target_year, target_month, target_week)
    year_week_token = _state_year_week_token(target_year, target_month, target_week)
    n_accounts = min(6, len(ACCOUNTS))
    forced_wrun = _parse_week_run_env()
    plan = _resolve_week_plan_from_folders(ymw_str, n_accounts, forced_wrun)
    if plan is None:
        if forced_wrun is not None:
            print(
                f"[스케줄] WEEK_RUN={forced_wrun} 지정: {ymw_str} 에 해당 회차 사업자 폴더가 모두 있습니다. "
                f"WEEK_RUN 없이 실행하면 {EXCEL_SAVE_BASE}\\{ymw_str} 아래 1~{RUNS_PER_WEEK}회차 중 "
                "비어 있는 가장 앞 회차를 자동 선택합니다."
            )
        else:
            print(
                f"[스케줄] {ymw_str} — 1~{RUNS_PER_WEEK}회차 전 사업자 폴더가 모두 있습니다. 추가 작업 없이 종료합니다."
            )
        return
    week_run, ranks_to_run = plan
    # 제어판 '선택 사업자만' : ONLY_RANKS=2,4 처럼 지정 시 해당 사업자로만 제한
    _only = os.environ.get("ONLY_RANKS", "").strip()
    if _only:
        _sel = {int(x) for x in _only.replace(" ", "").split(",") if x.strip().isdigit()}
        _filtered = [r for r in ranks_to_run if r in _sel]
        if _filtered:
            ranks_to_run = _filtered
            print(f"[스케줄] ONLY_RANKS={_only} → 선택 사업자만: {ranks_to_run}")
        else:
            print(f"[스케줄] ONLY_RANKS={_only} 이나 이번 회차 대상에 없음 → 기존 대상 유지")
    _write_week_run_state(year_week_token, week_run)
    wr_note = f" (환경변수 WEEK_RUN={forced_wrun} 고정)" if forced_wrun is not None else ""
    print(
        f"\n[스케줄] 폴더 기준{wr_note}: 이번 회차={week_run}회차, "
        f"실행할 사업자만 ({len(ranks_to_run)}명): {', '.join(f'{r}번' for r in ranks_to_run)}"
    )
    this_week_30 = _build_keywords_30(target_month, target_week)
    ly, lm, lw = _get_last_week_ymw(target_year, target_month, target_week)
    last_week_30 = _load_week_keywords(_week_key_key(ly, lm, lw))

    with sync_playwright() as pw:
        # Phase 3 와 동일: 전용 프로필 Chrome 을 원격디버깅으로 띄우고 CDP 로 연결 (프로필 복사 없음)
        launch_kw_saved = {"user_data_dir": str(_P1_CDP_DIR)}  # 후방 호환(잔존 참조용)
        _pre = _p1_kill_debug_chrome()
        if _pre:
            print(f"[시작] 이전 Phase1 디버그 Chrome {_pre}개 정리", flush=True)
            time.sleep(2)
        _chrome_proc = _p1_launch_debug_chrome()
        print(
            f"[시작] 전용 Chrome 기동 (pid={_chrome_proc.pid}, port={_P1_CDP_PORT}, "
            f"dir={_P1_CDP_DIR.name}) — 개인 Chrome 무관",
            flush=True,
        )
        if not _p1_wait_cdp(_P1_CDP_PORT, timeout=45):
            print(f"[시작] [오류] CDP 포트 {_P1_CDP_PORT} 응답 없음 → 종료", flush=True)
            try:
                _p1_kill_debug_chrome()
            except Exception:
                pass
            sys.exit(1)
        browser = pw.chromium.connect_over_cdp(f"http://127.0.0.1:{_P1_CDP_PORT}")
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        try:  # keep-alive 탭: 작업 탭이 닫혀도 Chrome 유지
            _ = context.pages[0] if context.pages else context.new_page()
        except Exception:
            pass
        print(
            "[시작] Chrome 연결 완료(CDP). new_page → domemedb(commit→domcontentloaded).",
            flush=True,
        )
        work_page = _domeme_startup_new_tab_goto_domemedb(context, log_prefix="[시작] ")
        page = work_page
        if not _ensure_domeme_ready_for_login(work_page, 0):
            print("[시작] 경고: 도매매 URL 진입 실패. 사업자 루프에서 재시도합니다.")

        time.sleep(_S(0.12, 0.35))
        if CHROME_PROCESS_DIAG:
            print("\n[진단] 실제 브라우저 프로세스 인자 (CHROME_PROCESS_DIAG=1):")
            _log_chrome_process_args(launch_kw_saved.get("user_data_dir"))
            print("\n")

        _phase2_deferred = False
        _phase2_upload_items = []
        try:
            # 주차·회차·사업자 목록은 Chrome 실행 전에 확정(마이박스 저장 폴더 미생성 기준).

            # === [1] 도매매 상품검색후 마이박스 추출: 이번에 비어 있던 사업자만 ===
            for rank in ranks_to_run:
                if rank > len(ACCOUNTS):
                    break
                user_id = ACCOUNTS[rank - 1]
                keyword = _get_keyword_for_run(
                    target_month, target_week, rank, week_run,
                    this_week_30=this_week_30, last_week_30=last_week_30,
                )
                kw_tag = _keyword_to_hashtag_safe(keyword)
                biz_hash_tag = f"{kw_tag}_{target_year}{target_month:02d}w{target_week}"
                print(f"\n{'='*60}")
                category = RUN_CATEGORY.get(week_run, "스포츠/레저")
                print(f"[1] 도매매 상품검색후 마이박스 추출 | {rank}번사업자 ({user_id}) | {week_run}회차 ({category})")
                print(f"    {target_month}월 {target_week}째주 | 키워드: {keyword}")
                print(f"    해시태그: {biz_hash_tag} (1주 7회 동일 해시로 누적)")
                print(f"{'='*60}")

                # [수정2] 사업자 전환 시 stale 탭 강제 정리. 도매매 검색결과·스피드고 마이박스·
                # 다운로드 팝업이 사업자별로 누적되면 login/검색 click 이 이전 탭으로 빨려들어가
                # wrong-account 동작·event listener 혼선이 발생.
                # ★핵심 버그: 직전 사업자 처리 끝에 work_page 가 '스피드고 마이박스 탭'으로
                #   바뀌어 있으면, 그 탭을 keep 해버려 검색이 도매매가 아닌 스피드고 탭에 입력됨
                #   → "도매매 로그인을 해주세요" alert → 마이박스담기 0건 → _최종.xlsx 미생성.
                #   따라서 work_page 는 'domemedb' 탭일 때만 보존하고, 그 외(speedgo 등)는 닫는다.
                _keep = []
                try:
                    if work_page is not None and not work_page.is_closed():
                        _wpu = (work_page.url or "").lower()
                        if "domemedb.domeggook.com" in _wpu:
                            _keep = [work_page]
                except Exception:
                    _keep = []
                try:
                    _p1_close_stale_tabs(context, keep_pages=_keep, label=f"[{rank}번] ")
                except Exception as _te:
                    print(f"[{rank}번] stale 탭 정리 예외(무시): {_te}", flush=True)
                # 보존 대상이 아니었으면 work_page 참조를 버려 _domeme_work_page 가 새 도매매 탭을 만들게 한다.
                if not _keep:
                    work_page = None

                # 새 탭을 만들지 않고 이전 작업 탭(또는 이미 열린 도매매 탭)에서 이어서 로그인·검색.
                work_page = _domeme_work_page(context, work_page, f"[{rank}번] ")
                page = work_page
                if page is None:
                    print(f"[{rank}번] 탭 확보 불가 — 이 사업자 건너뜀")
                    continue
                try:
                    page.bring_to_front()
                except Exception:
                    pass
                if not _ensure_domeme_ready_for_login(work_page, rank):
                    print(f"[{rank}번] 도매매 접속 실패로 이 사업자 단계를 건너뜁니다.")
                    continue
                time.sleep(_S(0.35, 0.8))
                # [수정] 사업자 전환 시 컨텍스트 쿠키 명시 초기화(Phase2/3 와 동일) →
                # 이전 사업자 세션이 chrome_phase1_cdp 프로필에 남아 '이미 로그인됨'으로
                # 폼이 표시되지 않아 "로그인 폼을 찾지 못했습니다. 다음 사업자로." 로 빠지던 문제 차단.
                try:
                    context.clear_cookies()
                    print(f"[{rank}번] 컨텍스트 쿠키 초기화 → 새 사업자 fresh 로그인 ({user_id})", flush=True)
                except Exception as _ce:
                    print(f"[{rank}번] 쿠키 초기화 실패(무시): {_ce}", flush=True)
                # 쿠키 초기화 직후 도매매 페이지를 다시 로드해야 '로그인' 폼이 노출된다.
                try:
                    page.goto("https://domemedb.domeggook.com/index/", wait_until="domcontentloaded", timeout=30000)
                except Exception as _ge:
                    print(f"[{rank}번] 쿠키 초기화 후 재진입 실패(무시): {_ge}", flush=True)
                time.sleep(_S(0.3, 0.7))
                print(f"[{rank}번] 기존 세션 정리(필요 시 로그아웃)…")
                _domeme_logout_if_logged_in(page)

                print(f"도매매 로그인 시도: {user_id}")
                lp = f"[{rank}번] "
                # Phase 2(test_speedgo_upload_1번)와 동일 동선을 먼저 시도 — 직접 mem_formLogin 과 DOM/세션이 달라질 수 있음
                phase2_style_ok = _domeme_login_speedgo_phase_style(page, user_id, PASSWORD, lp)
                if not phase2_style_ok:
                    print(f"{lp}Phase2 동선 실패 → mem_formLogin·강화 폼 경로로 재시도", flush=True)
                    _domeme_goto_login_pc(page, lp)
                    try:
                        print(f"{lp}로그인(폴백) 진입 후 url={(page.url or '')[:200]!r}", flush=True)
                    except Exception:
                        pass

                    id_field, pw_field, form_ctx = _find_login_form_pair(page)
                    if not id_field or not pw_field:
                        print(f"{lp}로그인 폼 미표시 → '로그인' 모달 시도…")
                        _domeme_click_open_login_modal(page, lp)
                        time.sleep(_S(0.8, 2))
                        _domeme_dismiss_blocking_layers(page)
                        _domeme_wait_login_password_visible(page, lp)
                        id_field, pw_field, form_ctx = _find_login_form_pair(page)

                    if not id_field or not pw_field:
                        print(f"{lp}로그인 폼을 찾지 못했습니다. 다음 사업자로.")
                        continue

                    if not _domeme_fill_login_field(id_field, user_id, lp, "아이디"):
                        print(f"{lp}아이디 입력 실패. 다음 사업자로.")
                        continue
                    time.sleep(_S(0.12, 0.35))
                    if not _domeme_fill_login_field(pw_field, PASSWORD, lp, "비밀번호"):
                        print(f"{lp}비밀번호 입력 실패. 다음 사업자로.")
                        continue
                    time.sleep(_S(0.15, 0.4))

                    if not _click_domeggook_login_submit(page, form_ctx):
                        try:
                            pw_field.press("Enter")
                        except Exception as e:
                            print(f"{lp}로그인 제출(Enter) 실패: {e}")
                time.sleep(_S(0.5, 1.2))
                _login_wait_ms = 28000 if not FAST_MODE else 12000
                success_page = _domeme_wait_login_success_any_tab(context, page, _login_wait_ms)
                if success_page is not None:
                    if success_page != page:
                        print(
                            f"[{rank}번] 로그인 후 새 탭에서 세션 확인 → 이후 작업 탭을 전환합니다.",
                            flush=True,
                        )
                    page = success_page
                    work_page = success_page
                    try:
                        page.bring_to_front()
                    except Exception:
                        pass
                    print("로그인 완료.")
                else:
                    try:
                        u_tail = (page.url or "")[:200]
                    except Exception:
                        u_tail = "?"
                    print(
                        f"[{rank}번] [경고] 로그인 성공을 확인하지 못했습니다. url={u_tail!r} "
                        "녹색 '로그인하기'를 수동으로 누르거나, 캡차·2단계·아이디/비밀번호를 확인하세요.",
                        flush=True,
                    )

                # === 검색 → 200개 → 마이박스담기 (해시태그로 7회 누적) ===
                print(f"\n--- 키워드: {keyword} ---")

                # === 검색창에 입력 후 검색 ===
                search_input_selectors = [
                    'input[name*="keyword"]',
                    'input[name*="search"]',
                    'input[id*="search"]',
                    'input[id*="keyword"]',
                    '#searchKeyword',
                    '.search_input input',
                    'input[placeholder*="검색"]',
                    'input[placeholder*="상품"]',
                    'input[type="text"]',
                ]
                search_input = None
                for sel in search_input_selectors:
                    try:
                        el = page.locator(sel).first
                        if el.count() > 0 and el.is_visible():
                            search_input = el
                            break
                    except Exception:
                        continue

                if search_input:
                    search_input.fill(keyword)
                    time.sleep(_S(0.2, 0.5))
                    searched = False
                    search_btn_selectors = [
                        'button[type="submit"]',
                        'input[type="submit"]',
                        '.btn_search',
                        '[class*="search"] button',
                        'button:has(svg)',
                        'button:has-text("검색")',
                        'a:has-text("검색")',
                    ]
                    for sel in search_btn_selectors:
                        try:
                            btn = page.locator(sel).first
                            if btn.count() > 0 and btn.is_visible():
                                btn.click()
                                searched = True
                                break
                        except Exception:
                            continue
                    if not searched:
                        search_input.press("Enter")
                    print(f"검색 실행: {keyword}")
                    time.sleep(_S(1, 3))

                try:
                    page.evaluate("window.scrollBy(0, 200)")
                    time.sleep(_S(0.2, 0.5))
                except Exception:
                    pass

                # 국내배송상품 클릭 (200개 보기 전)
                try:
                    domestic_btn = page.locator('button:has-text("국내배송상품")').first
                    if domestic_btn.count() > 0 and domestic_btn.is_visible():
                        domestic_btn.click()
                        print("국내배송상품 클릭")
                        time.sleep(_S(0.5, 1.5))
                except Exception as e:
                    print(f"국내배송상품 버튼 클릭 실패(무시): {e}")

                # === 검색 결과: 200개 보기 + 전체선택 ===
                try:
                    view_changed = False
                    for sel in ('select:has(option:has-text("50"))', 'select:has(option:has-text("개"))', 'select'):
                        if view_changed:
                            break
                        try:
                            view_select = page.locator(sel).first
                            if view_select.count() > 0 and view_select.is_visible():
                                for opt in ("200개씩 보기", "200개 보기", "200"):
                                    try:
                                        view_select.select_option(label=opt)
                                        view_changed = True
                                        print("200개 보기로 변경")
                                        break
                                    except Exception:
                                        pass
                                if not view_changed:
                                    try:
                                        view_select.select_option(value="200")
                                        view_changed = True
                                        print("200개 보기로 변경")
                                    except Exception:
                                        pass
                        except Exception:
                            continue

                    if not view_changed:
                        try:
                            dd = page.get_by_text("50개씩 보기").first
                            if dd.count() > 0 and dd.is_visible():
                                dd.click()
                                time.sleep(_S(0.2, 0.5))
                                page.get_by_text("200개", exact=False).first.click()
                                view_changed = True
                                print("200개 보기로 변경")
                        except Exception:
                            pass

                    time.sleep(_S(0.8, 2))

                    try:
                        page.evaluate("window.scrollBy(0, 150)")
                        time.sleep(_S(0.2, 0.5))
                    except Exception:
                        pass
                    all_checked = False

                    try:
                        loc = page.locator('span.txt8:has-text("전체선택")').first
                        if loc.count() > 0:
                            loc.scroll_into_view_if_needed()
                            loc.click(force=True, timeout=3000)
                            all_checked = True
                    except Exception:
                        pass

                    if not all_checked:
                        try:
                            clicked = page.evaluate("""
                                () => {
                                    const txt8 = Array.from(document.querySelectorAll('span.txt8')).find(s => s.textContent?.trim() === '전체선택');
                                    if (txt8) {
                                        txt8.click();
                                        return true;
                                    }
                                    const row = Array.from(document.querySelectorAll('*')).find(el =>
                                        el.querySelector('span.input_check3_span') && el.textContent?.includes('전체선택') && el.textContent?.includes('기본보관함')
                                    );
                                    if (row) {
                                        row.querySelector('span.input_check3_span').click();
                                        return true;
                                    }
                                    return false;
                                }
                            """)
                            if clicked:
                                all_checked = True
                        except Exception:
                            pass

                    if not all_checked:
                        for loc, use_force in [
                            (page.locator(':has-text("전체선택"):has-text("기본보관함") input[type="checkbox"]').first, True),
                            (page.get_by_role("checkbox", name="전체선택").first, True),
                            (page.locator('label:has-text("전체선택") input[type="checkbox"]').first, True),
                            (page.get_by_text("전체선택", exact=True).first, True),
                        ]:
                            try:
                                if loc.count() > 0:
                                    loc.scroll_into_view_if_needed()
                                    if use_force:
                                        loc.click(force=True, timeout=3000)
                                    else:
                                        loc.click(timeout=3000)
                                    all_checked = True
                                    break
                            except Exception:
                                continue
                    if all_checked:
                        print("전체선택 체크 완료")
                    else:
                        print("전체선택 체크 실패 - 수동으로 체크해 주세요.")

                    # === 해시태그 입력 + 마이박스담기 (동일 해시로 5세분화 누적) ===
                    try:
                        hash_tag_value = biz_hash_tag

                        try:
                            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                            time.sleep(_S(0.2, 0.5))
                        except Exception:
                            pass

                        hash_filled = False
                        try:
                            inp = page.get_by_placeholder("해시태그입력").first
                            if inp.count() > 0:
                                inp.scroll_into_view_if_needed()
                                inp.fill(hash_tag_value)
                                hash_filled = True
                        except Exception:
                            pass
                        if not hash_filled:
                            try:
                                inp = page.locator('input[placeholder*="해시태그"]').first
                                if inp.count() > 0:
                                    inp.scroll_into_view_if_needed()
                                    inp.fill(hash_tag_value)
                                    hash_filled = True
                            except Exception:
                                pass
                        if not hash_filled:
                            try:
                                filled = page.evaluate("""
                                    (val) => {
                                        const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent?.includes('마이박스담기'));
                                        if (!btn) return false;
                                        const container = btn.closest('div')?.parentElement || document.body;
                                        const inp = container.querySelector('input[type="text"]');
                                        if (inp) {
                                            inp.value = val;
                                            inp.dispatchEvent(new Event('input', { bubbles: true }));
                                            inp.dispatchEvent(new Event('change', { bubbles: true }));
                                            return true;
                                        }
                                        return false;
                                    }
                                """, hash_tag_value)
                                if filled:
                                    hash_filled = True
                            except Exception:
                                pass
                        if not hash_filled:
                            try:
                                inp = page.locator('.footer_position input, [class*="footer"] input[type="text"]').first
                                if inp.count() > 0:
                                    inp.scroll_into_view_if_needed()
                                    inp.fill(hash_tag_value)
                                    hash_filled = True
                            except Exception:
                                pass
                        if not hash_filled:
                            try:
                                filled = page.evaluate("""
                                    (val) => {
                                        const all = document.querySelectorAll('input[type="text"]');
                                        for (const inp of all) {
                                            const ph = (inp.placeholder || '').toLowerCase();
                                            if (ph.includes('해시') || ph.includes('hashtag')) {
                                                inp.value = val;
                                                inp.dispatchEvent(new Event('input', { bubbles: true }));
                                                return true;
                                            }
                                        }
                                        return false;
                                    }
                                """, hash_tag_value)
                                if filled:
                                    hash_filled = True
                            except Exception:
                                pass

                        if hash_filled:
                            print(f"해시태그 입력: {hash_tag_value}")
                            try:
                                page.keyboard.press("Tab")
                            except Exception:
                                pass
                        else:
                            print(f"해시태그 입력 실패 - 수동으로 '{hash_tag_value}' 입력해 주세요.")
                        time.sleep(_S(0.22, 0.45))

                        if _domeme_click_mybox_add(page):
                            print("마이박스담기 완료")
                            time.sleep(_S(0.8, 2))
                        else:
                            print("마이박스담기 버튼 클릭 실패 — 화면 하단 '마이박스담기'를 수동으로 눌러 주세요.")

                    except Exception as e:
                        print(f"해시태그/마이박스담기 실패: {e}")

                    # === 스피드고 → 엑셀 ===
                    try:
                        speedgo_hash = f"#{biz_hash_tag}"
                        # 스피드고 접속: 이미 열린 스피드고 탭이 있으면 사용, 없으면 새 탭에서 이동 (도매매 탭에서 goto 시 다른 네비게이션에 끊기지 않도록)
                        page.bring_to_front()
                        time.sleep(0.5)
                        speedgo_page = None
                        for p in context.pages:
                            try:
                                if "speedgo" in (p.url or "").lower():
                                    speedgo_page = p
                                    speedgo_page.bring_to_front()
                                    break
                            except Exception:
                                pass
                        if speedgo_page is None:
                            work_page = context.new_page()
                            work_page.goto(SPEEDGO_URL, wait_until=_WAIT, timeout=30000)
                            page = work_page
                        else:
                            work_page = speedgo_page
                            page = work_page
                            work_page.wait_for_load_state(_WAIT, timeout=15000)
                        print("스피드고전송기 접속 완료")
                        time.sleep(_S(0.6, 1.5))

                        try:
                            # 1) 마이박스 메뉴 클릭
                            mybox_menu = page.get_by_role("link", name="마이박스").first
                            if mybox_menu.count() == 0:
                                mybox_menu = page.locator('a:has-text("마이박스")').first
                            if mybox_menu.count() > 0 and mybox_menu.is_visible():
                                mybox_menu.click()
                                print("마이박스 메뉴 클릭")
                                time.sleep(_S(0.8, 2))

                            # 2) 해시태그 입력 후 검색
                            hash_filled = False
                            for sel in ('input[placeholder*="해시태그"]', 'input[placeholder*="#태그명"]', 'input[placeholder*="태그"]'):
                                try:
                                    inp = page.locator(sel).first
                                    if inp.count() > 0 and inp.is_visible():
                                        inp.fill(speedgo_hash)
                                        hash_filled = True
                                        print(f"해시태그 입력: {speedgo_hash}")
                                        break
                                except Exception:
                                    continue
                            if not hash_filled:
                                try:
                                    filled = page.evaluate("""
                                    (val) => {
                                        const all = document.querySelectorAll('input[type="text"]');
                                        for (const i of all) {
                                            if ((i.placeholder || '').includes('해시') || (i.placeholder || '').includes('태그')) {
                                                i.value = val; i.dispatchEvent(new Event('input', { bubbles: true }));
                                                return true;
                                            }
                                        }
                                        return false;
                                    }
                                    """, speedgo_hash)
                                    if filled:
                                        hash_filled = True
                                        print(f"해시태그 입력 (JS): {speedgo_hash}")
                                except Exception:
                                    pass
                            time.sleep(_S(0.2, 0.5))

                            # 3) 검색 버튼 클릭
                            search_btn = page.locator('button:has-text("검색"):not(:has-text("초기화"))').first
                            if search_btn.count() == 0:
                                search_btn = page.get_by_role("button", name="검색").first
                            if search_btn.count() > 0 and search_btn.is_visible():
                                search_btn.click()
                                print("검색 클릭")
                                page.wait_for_load_state(_WAIT, timeout=30000)
                                time.sleep(_S(0.6, 1.5))

                            # 4) mb_saveList.php URL로 직접 이동 (pageLimit=1000, 5세분화 누적)
                            mb_save_list_url = (
                                "https://speedgo.domeggook.com/mybox/mb_saveList.php?"
                                f"pagenum=&hashTag={quote(speedgo_hash, safe='')}&sf=subject&sw=&itemNos=&mnp=&mxp="
                                "&titleStatus=&editStatus=&useOption=&sender_date1=&sender_date2="
                                f"&sort1=&sort2=&sort3=&sort4=&sort5=&b2bStatus=0&pageLimit={TARGET_PER_WEEK}"
                            )
                            work_page = page
                            # [안정화] 상품목록 로딩: 30s + 3회 재시도(매번 재이동) — 일시 지연/ERR_ABORTED로
                            # 사업자가 영구 누락되던 최대 원인 완화
                            _mb_ok = False
                            for _mb_try in range(1, 4):
                                try:
                                    work_page.goto(mb_save_list_url, wait_until=_WAIT, timeout=30000)
                                    print(f"마이박스 {TARGET_PER_WEEK}개 보기 페이지 이동: {speedgo_hash} (시도 {_mb_try}/3)")
                                    time.sleep(_S(0.8, 2))
                                    page.wait_for_selector('text=상품목록', state="visible", timeout=30000)
                                    print("마이박스 상품목록 페이지 로딩 완료")
                                    _mb_ok = True
                                    break
                                except Exception as _mbe:
                                    print(f"[{rank}번] 마이박스 상품목록 로딩 실패(시도 {_mb_try}/3): {str(_mbe)[:120]}")
                                    time.sleep(_S(1.5, 3.5))
                            if not _mb_ok:
                                raise RuntimeError("마이박스 상품목록 로딩 3회 실패")
                            time.sleep(_S(0.4, 1))

                            # 상품 행 로딩 대기 (pageLimit=1000이므로 200/244/500 등 개수 무관)
                            # 스크롤 영역이 있으면 하단까지 스크롤하여 lazy-render된 행이 DOM에 붙도록 함
                            try:
                                page.evaluate("""
                                    () => {
                                        const el = document.querySelector('div[style*="overflow-y"]') || document.querySelector('.tab1')?.closest('div');
                                        if (el && el.scrollHeight > el.clientHeight) {
                                            el.scrollTop = el.scrollHeight;
                                            return true;
                                        }
                                        return false;
                                    }
                                """)
                                time.sleep(_S(0.3, 0.6))
                                page.evaluate("""
                                    () => {
                                        const el = document.querySelector('div[style*="overflow-y"]') || document.querySelector('.tab1')?.closest('div');
                                        if (el) el.scrollTop = 0;
                                    }
                                """)
                                time.sleep(_S(0.15, 0.3))
                            except Exception:
                                pass

                            # 전체선택 (#selectAll + jQuery / item[] 직접체크). 개수 제한 없음.
                            n_rows = _mybox_wait_for_rows(page, max_seconds=45)
                            if n_rows == 0:
                                print(
                                    f"[{rank}번 경고] 마이박스에 해당 해시태그 상품이 0건입니다. "
                                    "도매매에서 마이박스담기·해시태그가 맞는지 확인하세요. 엑셀 다운로드·STEP1~6 생략."
                                )
                            else:
                                print(f"마이박스 상품 행: {n_rows}건 (로딩 대기 후)")
                                chk = _mybox_select_all_items(page)
                                if chk > 0:
                                    print(f"전체선택 완료: 체크된 상품 {chk}건")
                                else:
                                    print(
                                        "[경고] 전체선택 후에도 체크된 상품이 0건입니다. "
                                        "엑셀 버튼 클릭 시 '선택해 주세요' 알림이 뜰 수 있습니다."
                                    )
                                time.sleep(_S(0.2, 0.5))

                            # 엑셀다운로드 버튼 클릭 (안정형 fallback 체인)
                            # ---- 프로필 비교 로그: popup, navigation, response, dialog ----
                            excel_events = []
                            profile_type = "실제_사용자_Chrome_프로필" if USE_REAL_CHROME_PROFILE else "자동화_전용_프로필"

                            def _log_excel_event(ev_type, detail):
                                entry = {"ts": time.time(), "type": ev_type, "detail": str(detail)[:500]}
                                excel_events.append(entry)
                                print(f"  [엑셀이벤트] {ev_type}: {str(detail)[:200]}")

                            if ENABLE_PROFILE_COMPARE_LOG:
                                try:
                                    _cid = id(context)
                                    if _cid not in _EXCEL_PROFILE_LOG_CONTEXT_IDS:
                                        _EXCEL_PROFILE_LOG_CONTEXT_IDS.add(_cid)
                                        page.on("popup", lambda p: _log_excel_event("popup", getattr(p, "url", str(p))))
                                        context.on("page", lambda p: _log_excel_event("new_page", getattr(p, "url", str(p))))
                                        page.on("framenavigated", lambda f: _log_excel_event("navigation", getattr(f, "url", str(f))))

                                        def _on_response(r):
                                            u = r.url.lower()
                                            if "excel" in u or "download" in u or "mb_save" in u or "xlsx" in u or ".php" in u:
                                                _log_excel_event(
                                                    "response",
                                                    {"url": r.url[:150], "status": r.status, "ct": r.headers.get("content-type", "")[:80]},
                                                )

                                        page.on("response", _on_response)
                                        # ★dialog 핸들러는 여기(컨텍스트당 1회)서 등록하지 않는다.
                                        #   page.on('dialog') 는 페이지 단위라, 첫 사업자 page 에만 붙어
                                        #   2번째 사업자부터 다운로드 confirm 이 자동 dismiss → 60s timeout.
                                        #   → 아래 다운로드 직전에 '사업자마다' 등록한다.
                                except Exception as le:
                                    print(f"[엑셀로그] 리스너 등록 실패: {le}")

                            def _find_excel_btn(ctx):
                                xpath = "//a[contains(., '엑셀다운로드')] | //button[contains(., '엑셀다운로드')] | //*[@onclick][contains(., '엑셀다운로드')]"
                                return ctx.locator(f"xpath={xpath}").first

                            def _poll_downloads_for_xlsx(before_files_set, target_path, poll_seconds=60):
                                """expect_download 실패 시, 기본 다운로드 폴더에서 새 파일 찾아 .xlsx로 이동.
                                서버가 UUID(확장자 없음)로 저장하는 경우도 처리."""
                                def _try_move(p, dest):
                                    dest = Path(dest)
                                    dest.parent.mkdir(parents=True, exist_ok=True)
                                    for _w in range(10):
                                        try:
                                            p.rename(str(dest))
                                            return True
                                        except PermissionError:
                                            time.sleep(_S(0.4, 1))
                                    shutil.copy2(p, dest)
                                    p.unlink(missing_ok=True)
                                    return True

                                downloads = Path.home() / "Downloads"
                                if not downloads.exists():
                                    return False
                                dest = Path(target_path)
                                if not dest.suffix.lower() == ".xlsx":
                                    dest = dest.parent / (dest.stem + ".xlsx")

                                for _ in range(max(1, poll_seconds // 2)):
                                    time.sleep(_S(0.8, 2))
                                    # 1) *.xlsx 파일
                                    for p in downloads.glob("*.xlsx"):
                                        try:
                                            p = p.resolve()
                                            if p in before_files_set:
                                                continue
                                            _try_move(p, dest)
                                            print(f"엑셀 저장 완료 (다운로드폴더 폴백): {dest}")
                                            return True
                                        except (OSError, PermissionError):
                                            continue
                                    # 2) 확장자 없는 파일 (서버가 UUID 등으로 저장한 경우, 20~50KB)
                                    for p in downloads.iterdir():
                                        try:
                                            if p.suffix or not p.is_file():
                                                continue
                                            p = p.resolve()
                                            if p in before_files_set:
                                                continue
                                            sz = p.stat().st_size
                                            if 15000 < sz < 100000:  # 엑셀 export 대략 20~50KB
                                                _try_move(p, dest)
                                                print(f"엑셀 저장 완료 (확장자없음→.xlsx 폴백): {dest}")
                                                return True
                                        except (OSError, PermissionError):
                                            continue
                                return False

                            iframes = page.frames  # 속성(리스트), 메서드 아님 - frames() 호출 시 'list' object is not callable
                            print(f"[엑셀] iframe 개수: {len(iframes)}")
                            excel_btn = None
                            for i, f in enumerate(iframes):
                                try:
                                    cand = _find_excel_btn(f)
                                    if cand.count() > 0 and cand.is_visible():
                                        excel_btn = cand
                                        print(f"[엑셀] 버튼 발견 (frame index={i})")
                                        break
                                except Exception:
                                    continue
                            if excel_btn is None:
                                excel_btn = _find_excel_btn(page)
                                if excel_btn.count() > 0:
                                    print("[엑셀] 버튼 발견 (main)")

                            if n_rows == 0:
                                print(f"[{rank}번] 엑셀다운로드 단계 생략 (마이박스 상품 0건)")
                            elif excel_btn and excel_btn.count() > 0:
                                excel_result = "미시도"
                                try:
                                    # ★[핵심 수정] "선택한 상품을 다운로드하겠습니까?" 네이티브 confirm 자동 수락 핸들러를
                                    #   '사업자마다(이 다운로드 page 에)' 등록. page.on('dialog') 는 페이지 단위이므로
                                    #   컨텍스트당 1회만 등록하면 2번째 사업자부터 confirm 이 자동 dismiss 되어
                                    #   download 이벤트가 영영 안 와 60s timeout → _최종.xlsx 미생성으로 이어졌다.
                                    def _dl_dialog(d):
                                        try:
                                            if ENABLE_PROFILE_COMPARE_LOG:
                                                _log_excel_event("dialog", f"type={d.type} message={(d.message or '')[:100]}")
                                        except Exception:
                                            pass
                                        try:
                                            d.accept()
                                        except Exception:
                                            pass
                                    try:
                                        page.on("dialog", _dl_dialog)
                                    except Exception:
                                        pass
                                    excel_btn.scroll_into_view_if_needed()
                                    time.sleep(_S(0.2, 0.5))
                                    for _ in range(2):
                                        try:
                                            page.wait_for_load_state("domcontentloaded")
                                            excel_btn.wait_for(state="visible", timeout=5000)
                                        except Exception:
                                            pass
                                        break
                                    download = None
                                    downloads_dir = Path.home() / "Downloads"
                                    before_files = set()
                                    if downloads_dir.exists():
                                        before_files = {p.resolve() for p in downloads_dir.iterdir() if p.is_file()}
                                    for method in ("normal", "force", "js"):
                                        try:
                                            with page.expect_download(timeout=60000) as dl_info:
                                                if method == "js":
                                                    excel_btn.evaluate("el => el.click()")
                                                else:
                                                    excel_btn.click(force=(method == "force"))
                                                # 커스텀 팝업 확인 버튼 (네이티브 confirm은 dialog handler에서 이미 수락)
                                                time.sleep(_S(0.3, 0.6))
                                                try:
                                                    page.get_by_role("button", name="확인").click(timeout=3000)
                                                except Exception:
                                                    try:
                                                        page.locator('button:has-text("확인")').first.click(timeout=2000)
                                                    except Exception:
                                                        pass  # 네이티브 confirm은 이미 수락됨
                                            download = dl_info.value
                                            print(f"[엑셀] 클릭 성공 (방식: {method})")
                                            break
                                        except Exception as e:
                                            print(f"[엑셀] {method} 클릭 실패: {e}")
                                            if method == "js":
                                                try:
                                                    with page.expect_download(timeout=60000) as dl_info:
                                                        page.evaluate("""
                                                        () => {
                                                            const btn = document.querySelector('button[onclick*="excelMbDown"]') ||
                                                                document.querySelector('*[onclick*="excelMbDown"]');
                                                            if (btn) { btn.click(); return; }
                                                            if (typeof excelMbDown === 'function') excelMbDown();
                                                        }
                                                        """)
                                                        time.sleep(_S(0.3, 0.6))
                                                        try:
                                                            page.get_by_role("button", name="확인").click(timeout=3000)
                                                        except Exception:
                                                            try:
                                                                page.locator('button:has-text("확인")').first.click(timeout=2000)
                                                            except Exception:
                                                                pass
                                                    download = dl_info.value
                                                    print("[엑셀] excelMbDown() JS 호출 성공")
                                                    break
                                                except Exception as e2:
                                                    print(f"[엑셀] JS fallback 실패: {e2}")
                                    if download:
                                        now_dt = datetime.now()
                                        ymw = f"{now_dt.year % 100}년{now_dt.month}월{_get_week_of_month(now_dt)}주차"
                                        biz_folder = EXCEL_SAVE_BASE / ymw / f"{week_run}회차" / f"{rank}번사업자"
                                        biz_folder.mkdir(parents=True, exist_ok=True)
                                        # 파일명: 사업자ID_도매매검색상품명_xx년xx월xx주차_xx회.xlsx
                                        excel_filename = f"{user_id}_{kw_tag}_{ymw}_{week_run}회.xlsx"
                                        target_path = biz_folder / excel_filename
                                        download.save_as(str(target_path))
                                        print(f"엑셀 저장 완료: {target_path}")
                                        excel_result = "성공(expect_download)"
                                    else:
                                        # expect_download 실패 시 다운로드폴더 폴링 폴백
                                        now_dt = datetime.now()
                                        ymw = f"{now_dt.year % 100}년{now_dt.month}월{_get_week_of_month(now_dt)}주차"
                                        biz_folder = EXCEL_SAVE_BASE / ymw / f"{week_run}회차" / f"{rank}번사업자"
                                        biz_folder.mkdir(parents=True, exist_ok=True)
                                        excel_filename = f"{user_id}_{kw_tag}_{ymw}_{week_run}회.xlsx"
                                        target_path = biz_folder / excel_filename
                                        if _poll_downloads_for_xlsx(before_files, target_path, poll_seconds=60):
                                            excel_result = "성공(다운로드폴더폴백)"
                                        else:
                                            excel_result = "실패"
                                            raise RuntimeError("엑셀 다운로드 트리거 실패 (expect_download 타임아웃, 다운로드폴더 폴링에도 파일 없음)")
                                    # 프로필 비교 로그 저장
                                    if ENABLE_PROFILE_COMPARE_LOG and excel_events:
                                        _append_profile_compare_log(profile_type, excel_result, excel_events)
                                except Exception as e:
                                    excel_result = f"예외: {str(e)[:100]}"
                                    if ENABLE_PROFILE_COMPARE_LOG and excel_events:
                                        _append_profile_compare_log(profile_type, excel_result, excel_events)
                                    try:
                                        (PROJECT_DIR / "excel_click_fail_page.html").write_text(page.content(), encoding="utf-8")
                                        print(f"[{rank}번 엑셀 실패] 페이지소스: excel_click_fail_page.html")
                                    except Exception:
                                        pass
                                    print(f"[{rank}번 엑셀 실패] {e}")
                            else:
                                print("엑셀다운로드 버튼을 찾지 못했습니다.")
                        except Exception as e:
                            print(f"[{rank}번] 스피드고·엑셀다운로드 구간 예외(상위로 전달): {e}")
                            raise
                    except Exception as e:
                        print(f"[{rank}번] 스피드고 마이박스/해시태그/검색 실패: {e}")

                except Exception as e:
                    print(f"[{rank}번] 검색결과·해시태그·마이박스·스피드고 전체 구간 실패: {e}")

                # === 해당 사업자 폴더에 대해 엑셀 합치기 ~ 최종 가공 (run_all_steps) ===
                biz_folder = EXCEL_SAVE_BASE / ymw_str / f"{week_run}회차" / f"{rank}번사업자"
                final_name = f"{user_id}_{kw_tag}_{ymw_str}_{week_run}회_최종.xlsx"
                if biz_folder.exists():
                    if not _folder_has_mergeable_source_excel(biz_folder):
                        print(
                            f"[{rank}번] 원본 엑셀(.xls/.xlsx) 없음 → STEP1~6 스킵. "
                            "마이박스에 상품이 있고 엑셀다운로드가 성공했는지 확인하세요."
                        )
                    else:
                        try:
                            from run_all_steps import run_all_steps_for_dir
                            run_all_steps_for_dir(biz_folder, f"{rank}번", final_output_name=final_name)
                        except Exception as e:
                            print(f"[{rank}번] run_all_steps 실패 (엑셀 합치기~최종): {e}")
                # 업로드는 Phase 2에서 6사업자 모두 완료 후 일괄 수행
                print(f"[{rank}번 사업자] 구간 종료 → 다음 사업자 루프", flush=True)

            # === _최종.xlsx 생성 후 X번_통합상품명_이미지 폴더 내 파일 전부 완전삭제 ===
            for _r in ranks_to_run:
                _biz = EXCEL_SAVE_BASE / ymw_str / f"{week_run}회차" / f"{_r}번사업자"
                _img_dir = _biz / f"{_r}번_통합상품명_이미지"
                if _img_dir.is_dir():
                    _cnt = 0
                    for _f in _img_dir.rglob("*"):
                        if _f.is_file():
                            try:
                                _f.unlink()
                                _cnt += 1
                            except OSError as e:
                                print(f"[이미지삭제] {_f.name} 실패: {e}")
                    if _cnt > 0:
                        print(f"[{_r}번사업자] {_img_dir.name} 내 {_cnt}개 파일 완전삭제 완료")

            # 이번 주 42개 키워드 저장 (다음 주에 직전주 겹침 시 7~12위 적용용)
            _save_week_keywords(year_week_key, this_week_30)

            # === Phase 1 종료. Phase 2(스피드고 업로드·전송)는 분리 실행 ===
            # Phase 2 를 Phase 1 에 붙이면(같은/분리 인스턴스 모두) cp949 크래시·프로필 잠금·
            # 라이브 프로필 재실행 거부로 실패한다. Phase 2 는 run_phase2.py 로 독립 실행한다
            # (전용 user-data-dir + 원격디버깅 연결, 검증된 동선).
            _phase2_upload_items = []
            print(f"\n{'='*60}")
            print("[Phase 1 완료] STEP1~9·_최종.xlsx 생성까지 끝났습니다.")
            print("[Phase 2] 스피드고 업로드·전송은 분리 실행하세요:")
            print(f"    python -u run_phase2.py --week-run {week_run}")
            print(f"{'='*60}")
            if not _phase2_deferred:
                print("브라우저를 열어두었습니다. 엔터를 누르면 스크립트를 종료합니다.")
                input()

        finally:
            try:
                _n = _p1_kill_debug_chrome()
                print(f"[종료] Phase1 전용 Chrome 정리 ({_n}개). 개인 Chrome 미접촉.", flush=True)
            except Exception:
                pass

    # === Phase 2 (별도 Playwright 블록): event loop/Profile 충돌 방지 ===
    if _phase2_deferred and _phase2_upload_items:
        print(f"\n{'='*60}\n[Phase 2] 스피드고 업로드 (ymw={ymw_str}, 회차={week_run})\n{'='*60}")

        # Phase 1 리소스 정리 (Event loop/httpx 등 충돌 방지)
        gc.collect()
        time.sleep(1)

        # Phase 1 Chrome 이 실제로 사용한 user_data_dir 을 쓰는 프로세스가 종료될 때까지 대기
        try:
            _phase1_ud = launch_kw_saved["user_data_dir"]
        except (NameError, KeyError, TypeError):
            _phase1_ud = str(_resolve_user_data_dir())
        print(f"[Phase 2] Phase 1 Chrome 종료 대기 중... (감시 경로: {_phase1_ud})")
        _wait_for_chrome_exit_using_path(_phase1_ud, timeout_sec=45)
        time.sleep(3)  # 추가 버퍼 (파일 잠금 해제 등)

        try:
            from test_speedgo_upload_1번 import run_speedgo_upload_phase
            # Phase 1 이 쓰던 디렉터리와 분리된 별도 경로로 새로 복사 (잠금 경합 방지)
            print(f"[Phase 2] 프로필 복사 중... → {REAL_PHASE2_RUN_DIR}")
            _copy_ok = _copy_phase2_profile67(REAL_PHASE2_RUN_DIR, fresh=True)
            _phase2_user_data = str(REAL_PHASE2_RUN_DIR) if _copy_ok else str(REAL_CHROME_USER_DATA)
            if not _copy_ok:
                print("[Phase 2] 프로필 복사 실패. 실제 User Data 사용 (Chrome 전체 종료 필요)")
            else:
                print(f"[Phase 2] 프로필 복사 완료 → {REAL_PHASE2_RUN_DIR}")

            _phase2_kw = _chrome_persistent_launch_kw_phase2_identical(str(_phase2_user_data))

            upload_context = None
            max_attempts = 3
            for attempt in range(1, max_attempts + 1):
                try:
                    _remove_chrome_profile_lock_files(Path(_phase2_kw["user_data_dir"]))
                    print(f"[Phase 2] Chrome 실행 시도 ({attempt}/{max_attempts})...")
                    with sync_playwright() as pw2:
                        try:
                            upload_context = pw2.chromium.launch_persistent_context(
                                **{**_phase2_kw, "no_viewport": True}
                            )
                        except TypeError:
                            _kw = {k: v for k, v in _phase2_kw.items() if k != "no_viewport"}
                            upload_context = pw2.chromium.launch_persistent_context(**_kw)
                        print("[Phase 2] Chrome 실행 완료")
                        upload_page = upload_context.new_page()
                        run_speedgo_upload_phase(upload_context, upload_page, _phase2_upload_items)
                        try:
                            upload_context.close()
                        except Exception:
                            pass
                    print("모든 작업 완료. 엔터를 누르면 스크립트를 종료합니다.")
                    break
                except Exception as e:
                    print(f"[Phase 2] Chrome 실행 실패 (시도 {attempt}/{max_attempts}): {e}")
                    if attempt < max_attempts:
                        wait_sec = 5 * attempt
                        print(f"[Phase 2] {wait_sec}초 후 재시도...")
                        time.sleep(wait_sec)
                        # 재시도 전 프로필 재복사 (손상된 복사본 방지)
                        if _copy_ok:
                            _copy_ok = _copy_phase2_profile67(REAL_PHASE2_RUN_DIR, fresh=True)
                            _phase2_kw["user_data_dir"] = str(REAL_PHASE2_RUN_DIR) if _copy_ok else str(REAL_CHROME_USER_DATA)
                    else:
                        import traceback
                        traceback.print_exc()
                        print("엔터를 누르면 스크립트를 종료합니다.")
        except Exception as e:
            print(f"[Phase 2] 스피드고 업로드 실패: {e}")
            import traceback
            traceback.print_exc()
            print("엔터를 누르면 스크립트를 종료합니다.")
        input()


if __name__ == "__main__":
    print(f"실행 경로: {Path(__file__).resolve()}")
    main()
