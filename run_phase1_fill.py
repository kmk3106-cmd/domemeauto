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


def _has_work_left(locked_week=None):
    """미완 사업자 식별. locked_week 가 주어지면 그 회차만 본다(다른 회차로 절대 안 넘어감).

    locked_week=None 이고 WEEK_RUN env 도 없으면 _resolve_week_plan_from_folders 가
    '1~7회차 중 가장 앞의 미완 회차'를 고른다 — 단 이 회차도 첫 attempt 후 main() 이
    잠가둔다. 즉 한 fill 호출 동안 절대 다른 회차로 안 넘어감.
    """
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
    forced = locked_week if locked_week is not None else _parse_week_run_env()
    plan = _resolve_week_plan_from_folders(ymw_str, n, forced)
    if plan is None:
        return None
    wr, ranks = plan
    # 선택실행(ONLY_RANKS) 존중: 선택한 사업자 중 '미완'인 것만 남긴다.
    only = os.environ.get("ONLY_RANKS", "").strip()
    if only:
        sel = {int(x) for x in only.replace(" ", "").split(",") if x.strip().isdigit()}
        ranks = [r for r in ranks if r in sel]
        if not ranks:
            return None
    return (ymw_str, wr, ranks)


def main() -> int:
    """한 fill 호출 = 한 회차 사이클 (절대 다음 회차로 안 넘어감).

    원칙: 사용자 멘탈 모델 '한 회차 = 한 P1~P3 사이클' 을 준수.
      - 첫 attempt 에서 처리할 회차를 결정(WEEK_RUN env > .week_run_state 의 첫 미완 회차).
      - 그 회차를 locked_week 로 잠그고 이후 attempt 는 그 회차만 본다.
      - 그 회차의 모든 사업자가 완료되면 fill 종료 (다른 미완 회차가 있어도 무시).
      - 그래야 'Phase 1~3 일괄' 이 1개 회차만 처리하고 P2/P3 로 넘어감.
    """
    max_attempts = int(os.environ.get("FILL_RETRY", "3"))
    print(f"[FILL] Phase1 회차 채우기 시작 (최대 {max_attempts}회, '한 회차 = 한 사이클' 원칙)", flush=True)
    py = sys.executable
    locked_week = None
    for attempt in range(1, max_attempts + 1):
        plan = _has_work_left(locked_week=locked_week)
        if plan is None:
            if locked_week is not None:
                print(f"[FILL] {locked_week}회차 모든 사업자 완료 → 종료 (총 {attempt-1}회 실행)", flush=True)
            else:
                print(f"[FILL] 더 채울 사업자 없음 → 종료 (총 {attempt-1}회 실행)", flush=True)
            return 0
        ymw, wr, ranks = plan
        if locked_week is None:
            locked_week = wr
            print(f"[FILL] 이번 사이클 회차 잠금: {locked_week}회차 (다른 회차는 건드리지 않음)", flush=True)
        print(f"[FILL {attempt}/{max_attempts}] {ymw} {wr}회차 미완 {ranks} → Phase1 실행", flush=True)
        env = dict(os.environ)
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        env["WEEK_RUN"] = str(locked_week)  # domeme 가 다른 회차로 안 가게 명시
        # stdin은 panel 호출 시 이미 _panel_stdin.txt 가 공급. 단독 실행 시 NUL 로.
        r = subprocess.run(
            [py, "-u", "domeme_auto_login_temp.py"],
            cwd=str(PROJECT_DIR), env=env,
            stdin=subprocess.DEVNULL if sys.stdin.closed or not sys.stdin.isatty() else None,
        )
        if r.returncode != 0:
            print(f"[FILL] Phase1 종료코드 {r.returncode} → 채우기 중단", flush=True)
            return r.returncode
    print(f"[FILL] 최대 {max_attempts}회 도달 — {locked_week}회차 미완 사업자 남았을 수 있음", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
