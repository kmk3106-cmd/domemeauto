# -*- coding: utf-8 -*-
"""
Phase 2 정식 분리 런처 (Phase 1 과 완전 독립).

검증된 동선: 전용 user-data-dir 로 원격디버깅 Chrome 자동 기동 → CDP 연결 →
            test_speedgo_upload_1번.main_upload_impl 로 사업자별 스피드고 업로드·전송.

Phase 1 의 cp949 크래시 / 프로필 복사 잠금 / 라이브 프로필 재실행 거부 문제를
구조적으로 회피한다 (전용 프로필 디렉터리 + 원격디버깅 연결, 사용자의 일반 Chrome 미접촉).

사용법:
  python -u run_phase2.py                 # .week_run_state 기준 (Phase 1 과 동일 회차)
  python -u run_phase2.py --week-run 1    # 회차 지정
  python -u run_phase2.py --week-run 1 --ymw 26년5월3주차
  python -u run_phase2.py --port 9333     # CDP 포트 지정(기본 9222)

종료 시: 이 런처가 띄운 디버그 Chrome 만 정리한다(사용자 일반 Chrome 은 건드리지 않음).
"""
import argparse
import os
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

# 콘솔 코드페이지(cp949 등)에서 print 크래시 방지
for _sn in ("stdout", "stderr"):
    try:
        getattr(sys, _sn).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

PROJECT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_DIR))

CDP_USER_DATA_DIR = PROJECT_DIR / "chrome_phase2_cdp"  # 이 런처 전용 (일반 프로필과 분리)


def _log(msg):
    print(msg, flush=True)


def _kill_our_debug_chrome():
    """이 런처 전용 user-data-dir 를 쓰는 chrome 프로세스만 종료 (사용자 일반 Chrome 미접촉)."""
    marker = str(CDP_USER_DATA_DIR).lower()
    killed = 0
    try:
        import psutil
    except ImportError:
        return killed
    for proc in psutil.process_iter(["pid", "name", "cmdline"]):
        try:
            name = (proc.info.get("name") or "").lower()
            if "chrome" not in name:
                continue
            cmd = " ".join(str(c or "") for c in (proc.info.get("cmdline") or [])).lower()
            if marker in cmd:
                proc.kill()
                killed += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return killed


def _wait_cdp(port, timeout=40):
    url = f"http://127.0.0.1:{port}/json/version"
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=3) as r:
                if r.status == 200:
                    return True
        except Exception:
            time.sleep(1)
    return False


def _build_items(ymw_str, week_run, only_ranks=None, skip_ranks=None):
    """EXCEL_SAVE_BASE/{ymw}/{week_run}회차/{n}번사업자/*_최종.xlsx → main_upload_impl 아이템.

    only_ranks: 이 사업자 번호만 (집합). None이면 전체.
    skip_ranks: 이 사업자 번호는 제외 (집합, 이미 전송 완료분 재전송 방지).
    """
    from domeme_auto_login_temp import (
        EXCEL_SAVE_BASE, ACCOUNTS, build_speedgo_hashtag, get_target_ymw,
    )
    from test_speedgo_upload_1번 import parse_final_filename

    target_year, target_month, target_week = get_target_ymw()
    base_run = Path(EXCEL_SAVE_BASE) / ymw_str / f"{week_run}회차"
    items = []
    for n in range(1, 7):
        if only_ranks and n not in only_ranks:
            continue
        if skip_ranks and n in skip_ranks:
            _log(f"[건너뜀] {n}번 (--skip 지정, 이미 전송 완료분)")
            continue
        biz_folder = base_run / f"{n}번사업자"
        if not biz_folder.is_dir():
            continue
        finals = list(biz_folder.glob("*_최종.xlsx"))
        if not finals:
            continue
        final_path = max(finals, key=lambda p: p.stat().st_mtime)
        kw_tag, _, _, _ = parse_final_filename(final_path)
        if kw_tag is None:
            _log(f"[건너뜀] {n}번 파일명 파싱 실패: {final_path.name}")
            continue
        user_id = ACCOUNTS[n - 1] if n <= len(ACCOUNTS) else ""
        if not user_id:
            _log(f"[건너뜀] {n}번 계정 없음")
            continue
        speedgo_hash = build_speedgo_hashtag(kw_tag, target_year, target_month, target_week)
        items.append((n, biz_folder, final_path, kw_tag, speedgo_hash, user_id))
    return items, base_run


def main():
    parser = argparse.ArgumentParser(description="Phase 2 분리 실행 (스피드고 업로드·전송)")
    parser.add_argument("--week-run", type=int, default=None, help="회차(1~7). 미지정 시 .week_run_state")
    parser.add_argument("--ymw", default=None, help='주차 문자열 예: "26년5월3주차". 미지정 시 현재 기준')
    parser.add_argument("--port", type=int, default=9222, help="CDP 원격디버깅 포트(기본 9222)")
    parser.add_argument("--ranks", default=None, help='이 사업자만 (쉼표) 예: "2,3,5,6"')
    parser.add_argument("--skip", default=None, help='이 사업자 제외 (쉼표) 예: "1" (이미 전송 완료분)')
    args = parser.parse_args()

    def _parse_ranks(s):
        if not s:
            return None
        return {int(x) for x in s.replace(" ", "").split(",") if x.strip().isdigit()}

    only_ranks = _parse_ranks(args.ranks)
    skip_ranks = _parse_ranks(args.skip)

    from domeme_auto_login_temp import get_upload_path_from_state, CHROME_EXECUTABLE, PASSWORD
    from test_speedgo_upload_1번 import main_upload_impl

    state_ymw, state_run = get_upload_path_from_state()
    ymw_str = args.ymw or state_ymw
    week_run = args.week_run if args.week_run is not None else state_run

    if not PASSWORD:
        _log("[오류] .env 의 DOMEME_PASSWORD 가 비어 있습니다.")
        sys.exit(1)

    items, base_run = _build_items(ymw_str, week_run, only_ranks=only_ranks, skip_ranks=skip_ranks)
    if not items:
        _log(f"[오류] 처리할 _최종.xlsx 없음: {base_run}")
        sys.exit(1)
    _log(f"[Phase 2] 대상 경로: {base_run}")
    _log(f"[Phase 2] 순차 전송 대상: {[x[0] for x in items]}번 사업자 ({len(items)}건)")

    chrome_exe = CHROME_EXECUTABLE if os.path.isfile(CHROME_EXECUTABLE) else "chrome"
    CDP_USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

    # 포트/디렉터리 선점한 이전 런처 Chrome 정리 (사용자 일반 Chrome 은 미접촉)
    pre_killed = _kill_our_debug_chrome()
    if pre_killed:
        _log(f"[Phase 2] 이전 디버그 Chrome {pre_killed}개 정리")
        time.sleep(2)

    chrome_proc = subprocess.Popen([
        chrome_exe,
        f"--remote-debugging-port={args.port}",
        f"--user-data-dir={CDP_USER_DATA_DIR}",
        "--profile-directory=Profile 67",
        "--start-maximized",
        "--no-first-run",
        "--no-default-browser-check",
        "about:blank",
    ])
    _log(f"[Phase 2] 디버그 Chrome 기동 (pid={chrome_proc.pid}, port={args.port}, dir={CDP_USER_DATA_DIR.name})")

    exit_code = 0
    try:
        if not _wait_cdp(args.port, timeout=40):
            _log(f"[오류] CDP 포트 {args.port} 응답 없음 (40s). 중단.")
            sys.exit(1)
        _log("[Phase 2] CDP 연결 준비됨")

        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            browser = pw.chromium.connect_over_cdp(f"http://127.0.0.1:{args.port}")
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            page = context.new_page()
            _log("[Phase 2] Chrome 연결 완료 → 전송 시작")
            main_upload_impl(page, items, PASSWORD)
            _log("[Phase 2] 모든 사업자 업로드·전송 루프 종료")
    except SystemExit:
        raise
    except Exception as e:
        exit_code = 1
        import traceback
        _log(f"[Phase 2] 예외: {e}")
        traceback.print_exc()
    finally:
        try:
            chrome_proc.terminate()
        except Exception:
            pass
        time.sleep(2)
        n = _kill_our_debug_chrome()
        _log(f"[Phase 2] 디버그 Chrome 정리 완료 ({n}개). 사용자 일반 Chrome 은 미접촉.")

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
