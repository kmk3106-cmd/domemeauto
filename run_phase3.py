# -*- coding: utf-8 -*-
"""
Phase 3 분리 런처: 스피드고 '공급사판매중지' 사업자별 1배치 삭제.

Phase 2 와 동일한 격리 방식(전용 user-data-dir + 원격디버깅 CDP 연결)으로
프로필 잠금/재실행 문제를 회피한다. 폴더/_최종.xlsx 의존 없음(사업자=ACCOUNTS).

사용법:
  python -u run_phase3.py                # 1~6번 전체
  python -u run_phase3.py --ranks 2,4    # 선택 사업자만
  python -u run_phase3.py --port 9222

⚠ 모든 활성 마켓에서 영구 삭제. 제어판 Phase3 버튼으로만 실행 권장.
종료 시 이 런처가 띄운 디버그 Chrome 만 정리(사용자 일반 Chrome 미접촉).
"""
import argparse
import os
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

for _sn in ("stdout", "stderr"):
    try:
        getattr(sys, _sn).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

PROJECT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_DIR))
CDP_DIR = PROJECT_DIR / "chrome_phase3_cdp"


def _log(m):
    print(m, flush=True)


def _kill_our_chrome():
    marker = str(CDP_DIR).lower()
    killed = 0
    try:
        import psutil
    except ImportError:
        return 0
    for p in psutil.process_iter(["pid", "name", "cmdline"]):
        try:
            if "chrome" not in (p.info.get("name") or "").lower():
                continue
            cmd = " ".join(str(c or "") for c in (p.info.get("cmdline") or [])).lower()
            if marker in cmd:
                p.kill()
                killed += 1
        except Exception:
            continue
    return killed


def _wait_cdp(port, timeout=40):
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


def main():
    ap = argparse.ArgumentParser(description="Phase 3 공급사판매중지 삭제")
    ap.add_argument("--ranks", default=None, help='이 사업자만 예: "2,4". 미지정=1~6 전체')
    ap.add_argument("--port", type=int, default=9222)
    args = ap.parse_args()

    from domeme_auto_login_temp import ACCOUNTS, CHROME_EXECUTABLE, PASSWORD
    from phase3_delete import main_delete_impl

    if not PASSWORD:
        _log("[오류] .env DOMEME_PASSWORD 비어있음")
        sys.exit(1)
    if not ACCOUNTS:
        _log("[오류] .env DOMEME_ACCOUNTS 비어있음")
        sys.exit(1)

    only = None
    if args.ranks:
        only = {int(x) for x in args.ranks.replace(" ", "").split(",")
                if x.isdigit() and 1 <= int(x) <= 6}
    items = []
    for rank in range(1, min(6, len(ACCOUNTS)) + 1):
        if only and rank not in only:
            continue
        items.append((rank, ACCOUNTS[rank - 1]))
    if not items:
        _log("[오류] 대상 사업자 없음")
        sys.exit(1)
    _log(f"[Phase 3] 대상 사업자: {[r for r, _ in items]}번 ({len(items)}건)")
    _log("[Phase 3] ⚠ 모든 활성 마켓에서 영구 삭제 — 공급사판매중지 1배치(500)")

    chrome = CHROME_EXECUTABLE if os.path.isfile(CHROME_EXECUTABLE) else "chrome"
    CDP_DIR.mkdir(parents=True, exist_ok=True)
    pre = _kill_our_chrome()
    if pre:
        _log(f"[Phase 3] 이전 디버그 Chrome {pre}개 정리")
        time.sleep(2)
    # 강제종료 후 프로필 잠금/비정상종료 플래그 정리 (Chrome 조기 사망 방지)
    for lk in ("SingletonLock", "SingletonSocket", "SingletonCookie", "lockfile"):
        try:
            (CDP_DIR / lk).unlink(missing_ok=True)
        except Exception:
            pass
    for sub in ("Default", "Profile 67"):
        try:
            pj = CDP_DIR / sub / "Preferences"
            if pj.exists():
                t = pj.read_text(encoding="utf-8", errors="ignore")
                t = t.replace('"exit_type":"Crashed"', '"exit_type":"Normal"')
                pj.write_text(t, encoding="utf-8")
        except Exception:
            pass

    proc = subprocess.Popen([
        chrome, f"--remote-debugging-port={args.port}",
        f"--user-data-dir={CDP_DIR}", "--profile-directory=Profile 67",
        "--start-maximized", "--no-first-run", "--no-default-browser-check",
        "about:blank",
    ])
    _log(f"[Phase 3] 디버그 Chrome 기동 (pid={proc.pid}, port={args.port})")

    code = 0
    try:
        if not _wait_cdp(args.port, 40):
            _log(f"[오류] CDP {args.port} 응답 없음")
            sys.exit(1)
        _log("[Phase 3] CDP 연결 준비됨")
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            br = pw.chromium.connect_over_cdp(f"http://127.0.0.1:{args.port}")
            ctx = br.contexts[0] if br.contexts else br.new_context()
            # keep-alive 탭: 작업 탭이 사이트에 의해 닫혀도 Chrome 자체가 종료되지 않게 유지
            try:
                keep = ctx.pages[0] if ctx.pages else ctx.new_page()
            except Exception:
                keep = ctx.new_page()
            _log("[Phase 3] Chrome 연결 완료 → 삭제 시작 (사업자별 새 탭)")
            main_delete_impl(ctx, items, PASSWORD)
            _log("[Phase 3] 삭제 루프 종료")
    except SystemExit:
        raise
    except Exception as e:
        code = 1
        import traceback
        _log(f"[Phase 3] 예외: {e}")
        traceback.print_exc()
    finally:
        try:
            proc.terminate()
        except Exception:
            pass
        time.sleep(2)
        n = _kill_our_chrome()
        _log(f"[Phase 3] 디버그 Chrome 정리 ({n}개). 사용자 일반 Chrome 미접촉.")
    sys.exit(code)


if __name__ == "__main__":
    main()
