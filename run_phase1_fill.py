# -*- coding: utf-8 -*-
"""Phase1 회차 자동 채우기.

domeme_auto_login_temp.py 를 '미완 사업자가 없을 때까지' 최대 FILL_RETRY(기본 3) 회 반복 실행.
domeme 의 폴더-기반 스케줄러가 매번 가장 앞 미완 회차의 빈 사업자만 처리하므로,
외부에서 N회 반복하면 자연히 그 회차가 6/6 으로 채워지고, 다음 회차로 진행한다.

domeme 내부 main 루프(약 1000줄)를 건드리지 않는 안전한 외부 래퍼.
"""
import os
import subprocess
import sys
from pathlib import Path

for _sn in ("stdout", "stderr"):
    try:
        getattr(sys, _sn).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

PROJECT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_DIR))


def _has_work_left():
    try:
        from domeme_auto_login_temp import (
            _resolve_week_plan_from_folders, ACCOUNTS, get_upload_path_from_state,
            _parse_week_run_env,
        )
    except Exception as e:
        print(f"[FILL] domeme 임포트 실패: {e}", flush=True)
        return None
    ymw_str, _ = get_upload_path_from_state()
    n = min(6, len(ACCOUNTS))
    forced = _parse_week_run_env()
    plan = _resolve_week_plan_from_folders(ymw_str, n, forced)
    if plan is None:
        return None
    wr, ranks = plan
    return (ymw_str, wr, ranks)


def main() -> int:
    max_attempts = int(os.environ.get("FILL_RETRY", "3"))
    print(f"[FILL] Phase1 회차 채우기 시작 (최대 {max_attempts}회)", flush=True)
    py = sys.executable
    for attempt in range(1, max_attempts + 1):
        plan = _has_work_left()
        if plan is None:
            print(f"[FILL] 더 채울 회차 없음 → 종료 (총 {attempt-1}회 실행)", flush=True)
            return 0
        ymw, wr, ranks = plan
        print(f"[FILL {attempt}/{max_attempts}] {ymw} {wr}회차 미완 {ranks} → Phase1 실행", flush=True)
        env = dict(os.environ)
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        # stdin은 panel 호출 시 이미 _panel_stdin.txt 가 공급. 단독 실행 시 NUL 로.
        r = subprocess.run(
            [py, "-u", "domeme_auto_login_temp.py"],
            cwd=str(PROJECT_DIR), env=env,
            stdin=subprocess.DEVNULL if sys.stdin.closed or not sys.stdin.isatty() else None,
        )
        if r.returncode != 0:
            print(f"[FILL] Phase1 종료코드 {r.returncode} → 채우기 중단", flush=True)
            return r.returncode
    print(f"[FILL] 최대 {max_attempts}회 도달 — 종료", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
