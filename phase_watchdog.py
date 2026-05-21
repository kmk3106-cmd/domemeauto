# -*- coding: utf-8 -*-
"""Phase 1/2/3 동작 감시 모듈 (서브에이전트의 입력 소스).

설계 목표
---------
- panel_*.log 한 파일을 파싱해 "지금 어디까지 진행됐고, 무엇이 꼬였는지" 를 구조화.
- 단순 키워드 매칭이 아닌 *사업자 단위 상태 머신* 으로 검증한다.
  · Phase 1: ranks_to_run 에서 받은 사업자별로 expected_steps 가 순서대로 도달해야 한다.
  · 같은 사업자 안에서 다른 사업자의 user_id/키워드/계정이 등장하면 → 탭 혼선 신호.
- 결과는 JSON 으로 stdout, 핵심 요약은 stderr (사람이 읽을 용도).

호출 방법
--------
    python -u phase_watchdog.py                  # 가장 최근 logs/panel_*.log
    python -u phase_watchdog.py <log_path>       # 특정 로그
    python -u phase_watchdog.py --json           # JSON 만
    python -u phase_watchdog.py --since 30       # 최근 30초 분량만 (스트리밍)

신호 레벨
--------
- INFO   🟢 정상 (참고용)
- WARN   🟡 의심 (계속 진행 가능)
- FAIL   🔴 실패/혼선 (개발자 개입 권장)
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

for _sn in ("stdout", "stderr"):
    try:
        getattr(sys, _sn).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

PROJECT_DIR = Path(__file__).resolve().parent
LOG_DIR = PROJECT_DIR / "logs"


# ===== 정규식 표 (한 곳에 모아둔다 — 새 패턴 추가가 잦으므로) =====
RE = {
    # 패널 메타
    "panel_start": re.compile(r"^===\s*(.+?)\s*시작\s+([\d\-:\s]+)\s*===\s*$"),
    "panel_step":  re.compile(r"^---\s*(.+?)\s*실행:\s*(.+?)\s*\|\s*env=(.+?)\s*---\s*$"),
    "panel_end":   re.compile(r"^---\s*(.+?)\s*종료코드\s*(\-?\d+)\s*---\s*$"),
    "panel_final": re.compile(r"^===\s*종료:\s*(\w+)\s+([\d\-:\s]+)\s*===\s*$"),

    # Phase 1 사업자 헤더
    "p1_biz_header": re.compile(r"\[1\]\s*도매매.*마이박스.*\|\s*(\d+)번사업자\s*\(([^)]+)\)\s*\|\s*(\d+)회차\s*\(([^)]+)\)"),
    "p1_keyword":    re.compile(r"^--- 키워드:\s*(.+?)\s*---\s*$"),
    "p1_hashtag":    re.compile(r"해시태그(?:\s*입력)?:\s*#?([^\s]+)"),

    # 사업자 prefix 라인
    "rank_prefix":   re.compile(r"^\[(\d+)번\]\s*(.+)$"),

    # 탭/페이지 상태
    "pages_count":   re.compile(r"pages 개수=(\d+)"),
    "page_url":      re.compile(r"page\[(\d+)\]\s*id=\d+\s*is_closed=\w+\s*url='([^']*)'"),
    "work_page_url": re.compile(r"work_page\s*id=\d+\s*is_closed=\w+\s*url='([^']*)'"),
    "stale_close":   re.compile(r"\[(\d+)번\]\s*stale 탭\s*(\d+)개 정리"),

    # 로그인 시퀀스
    "p1_login_try":   re.compile(r"^도매매 로그인 시도:\s*(\S+)"),
    "p1_submit_done": re.compile(r"로그인 제출 완료:\s*(\S+)"),
    "p1_login_ok":    re.compile(r"^로그인 완료\.$"),
    "p1_form_missed": re.compile(r"로그인 폼을 찾지 못했습니다"),
    "p1_cookie_clr":  re.compile(r"\[(\d+)번\]\s*컨텍스트 쿠키 초기화.*\(([^)]+)\)"),

    # 작업 진행
    "p1_search":      re.compile(r"^검색 실행:\s*(.+)$"),
    "p1_mybox_add":   re.compile(r"^마이박스담기 완료\s*$"),
    "p1_speedgo_ok":  re.compile(r"^스피드고전송기 접속 완료\s*$"),
    "p1_mybox_rows":  re.compile(r"마이박스 상품 행:\s*(\d+)건"),
    "p1_excel_btn":   re.compile(r"\[엑셀\]\s*버튼 발견"),
    "p1_excel_fail":  re.compile(r"\[엑셀\].*실패:.*Timeout"),
    "p1_dl_fail":     re.compile(r"엑셀 다운로드 트리거 실패"),
    "p1_step_skip":   re.compile(r"원본 엑셀\(\.xls/\.xlsx\) 없음 → STEP1~6 스킵"),
    "p1_seg_end":     re.compile(r"\[(\d+)번 사업자\] 구간 종료"),
    "p1_phase1_done": re.compile(r"\[Phase 1 완료\]"),

    # Phase 2
    "p2_chrome_up":   re.compile(r"\[Phase 2\] 디버그 Chrome 기동.*pid=(\d+).*port=(\d+)"),
    "p2_cdp_ready":   re.compile(r"\[Phase 2\] CDP 연결 준비됨"),
    "p2_targets":     re.compile(r"\[Phase 2\] 순차 전송 대상:\s*\[([^\]]+)\]번 사업자"),
    "p2_cookie":      re.compile(r"\[로그인\]\s*컨텍스트 쿠키 초기화.*\(([^)]+)\)"),
    "p2_send_ok":     re.compile(r"\[전송\].*전송 완료|전송 완료\s*$"),
    "p2_loop_end":    re.compile(r"\[Phase 2\] 모든 사업자 업로드.*종료"),

    # Phase 3
    "p3_chrome_up":   re.compile(r"\[Phase 3\] 디버그 Chrome 기동.*pid=(\d+)"),
    "p3_targets":     re.compile(r"\[Phase 3\] 대상 사업자:\s*\[([^\]]+)\]번"),
    "p3_loop_end":    re.compile(r"\[Phase 3\] 삭제 루프 종료"),

    # 공통 안전 신호
    "exception":      re.compile(r"예외|Traceback|Exception:"),
    "timeout":        re.compile(r"Timeout\s+(\d+)ms"),
}


# Phase 1 사업자 단위 expected 진행 단계 (순서 보존). 단계 누락 시 WARN.
P1_EXPECTED_STEPS = [
    ("stale_clean", "stale 탭 정리"),
    ("cookie_clear", "쿠키 초기화"),
    ("login_try", "로그인 시도"),
    ("login_ok", "로그인 완료"),
    ("search", "키워드 검색"),
    ("mybox_add", "마이박스담기 완료"),
    ("speedgo", "스피드고전송기 접속"),
    ("excel_btn", "엑셀 버튼 발견"),
    ("seg_end", "구간 종료"),
]


@dataclass
class Signal:
    level: str           # "INFO" | "WARN" | "FAIL"
    code: str            # 짧은 식별자
    msg: str
    line_no: int = 0
    rank: Optional[int] = None
    phase: Optional[str] = None

    def to_dict(self):
        return dict(level=self.level, code=self.code, msg=self.msg,
                    line=self.line_no, rank=self.rank, phase=self.phase)


@dataclass
class RankState:
    rank: int
    user_id: str = ""
    keyword: str = ""
    hashtag: str = ""
    steps_seen: Dict[str, int] = field(default_factory=dict)  # step_key -> line_no
    login_user_seen: Optional[str] = None
    login_submit_user: Optional[str] = None
    mybox_rows: Optional[int] = None
    excel_fail_count: int = 0
    seg_ended_line: Optional[int] = None
    started_line: int = 0


@dataclass
class WatchState:
    job_name: str = ""
    started_at: str = ""
    finished_at: str = ""
    final_result: str = ""
    current_step_name: str = ""
    current_rcode: Optional[int] = None
    phase: str = ""              # "p1" | "p2" | "p3"
    current_rank: Optional[int] = None
    pages_observed_max: int = 0
    ranks: Dict[int, RankState] = field(default_factory=dict)
    p2_targets: List[int] = field(default_factory=list)
    p3_targets: List[int] = field(default_factory=list)
    signals: List[Signal] = field(default_factory=list)

    def add_signal(self, sig: Signal):
        sig.phase = sig.phase or self.phase
        if sig.rank is None:
            sig.rank = self.current_rank
        self.signals.append(sig)


def _pick_latest_log() -> Optional[Path]:
    if not LOG_DIR.exists():
        return None
    logs = sorted(LOG_DIR.glob("panel_*.log"), key=lambda p: p.stat().st_mtime)
    return logs[-1] if logs else None


def parse_log(log_path: Path) -> WatchState:
    state = WatchState()
    if not log_path.exists():
        state.add_signal(Signal("FAIL", "no_log", f"로그 파일 없음: {log_path}"))
        return state

    text = log_path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()

    for i, raw in enumerate(lines, 1):
        line = raw.rstrip()
        if not line:
            continue
        _scan_line(state, line, i)

    _post_analyze(state, lines)
    return state


def _step_token_from_cmd(step_name: str, cmd: str) -> str:
    s = (step_name + " " + cmd).lower()
    if "phase1" in s.replace(" ", "") or "domeme_auto_login_temp" in s or "phase 1" in s:
        return "p1"
    if "run_phase2" in s or "phase 2" in s.lower() or "phase2" in s:
        return "p2"
    if "run_phase3" in s or "phase 3" in s.lower() or "phase3" in s:
        return "p3"
    if "phase 1 채우기" in s or "run_phase1_fill" in s:
        return "p1"
    return ""


def _scan_line(st: WatchState, line: str, i: int):
    m = RE["panel_start"].search(line)
    if m:
        st.job_name = m.group(1).strip()
        st.started_at = m.group(2).strip()
        return
    m = RE["panel_step"].search(line)
    if m:
        st.current_step_name = m.group(1).strip()
        st.phase = _step_token_from_cmd(m.group(1), m.group(2)) or st.phase
        st.current_rank = None
        return
    m = RE["panel_end"].search(line)
    if m:
        rc = int(m.group(2))
        st.current_rcode = rc
        if rc != 0:
            st.add_signal(Signal("FAIL", "step_nonzero",
                                 f"단계 종료코드 {rc}: {m.group(1).strip()}", i))
        return
    m = RE["panel_final"].search(line)
    if m:
        st.final_result = m.group(1).strip()
        st.finished_at = m.group(2).strip()
        return

    # Phase 1
    m = RE["p1_biz_header"].search(line)
    if m:
        rank = int(m.group(1))
        st.current_rank = rank
        rs = st.ranks.setdefault(rank, RankState(rank=rank, started_line=i))
        rs.user_id = m.group(2).strip()
        # 동일 사업자 헤더가 이미 등장한 적 있는데 또 등장 → 재시도(폴더 비어있어 다시 잡힘 OR 무한루프)
        if rs.started_line != i and rs.started_line > 0:
            st.add_signal(Signal("WARN", "rank_re_enter",
                                 f"{rank}번 사업자 헤더 2회 이상 등장 (재시도 가능성)", i, rank))
        rs.started_line = i
        return

    m = RE["p1_keyword"].search(line)
    if m and st.current_rank:
        rs = st.ranks[st.current_rank]
        new_kw = m.group(1).strip()
        if rs.keyword and rs.keyword != new_kw:
            st.add_signal(Signal("WARN", "keyword_changed",
                                 f"{st.current_rank}번 키워드 변경: {rs.keyword} → {new_kw}",
                                 i, st.current_rank))
        rs.keyword = new_kw
        return

    m = RE["stale_close"].search(line)
    if m:
        rank = int(m.group(1)); count = int(m.group(2))
        rs = st.ranks.setdefault(rank, RankState(rank=rank, started_line=i))
        rs.steps_seen["stale_clean"] = i
        if count >= 4:
            st.add_signal(Signal("WARN", "many_stale",
                                 f"{rank}번 진입 시 stale 탭 {count}개 정리 (이전 사업자 누적)",
                                 i, rank))
        return

    m = RE["p1_cookie_clr"].search(line)
    if m:
        rank = int(m.group(1)); uid = m.group(2).strip()
        rs = st.ranks.setdefault(rank, RankState(rank=rank, started_line=i))
        rs.steps_seen["cookie_clear"] = i
        if rs.user_id and rs.user_id != uid:
            st.add_signal(Signal("FAIL", "cookie_user_mismatch",
                                 f"{rank}번 쿠키초기화 user_id={uid} 가 헤더의 {rs.user_id} 와 불일치",
                                 i, rank))
        return

    m = RE["p1_login_try"].search(line)
    if m and st.current_rank:
        uid = m.group(1).strip()
        rs = st.ranks[st.current_rank]
        rs.login_user_seen = uid
        rs.steps_seen.setdefault("login_try", i)
        if rs.user_id and uid != rs.user_id:
            st.add_signal(Signal("FAIL", "login_user_mismatch",
                                 f"{st.current_rank}번 로그인 시도={uid} 가 헤더의 {rs.user_id} 와 불일치",
                                 i, st.current_rank))
        return

    m = RE["p1_submit_done"].search(line)
    if m and st.current_rank:
        uid = m.group(1).strip()
        rs = st.ranks[st.current_rank]
        rs.login_submit_user = uid
        if rs.user_id and uid != rs.user_id:
            st.add_signal(Signal("FAIL", "submit_user_mismatch",
                                 f"{st.current_rank}번 로그인 제출={uid} 가 헤더의 {rs.user_id} 와 불일치",
                                 i, st.current_rank))
        return

    if RE["p1_login_ok"].search(line) and st.current_rank:
        st.ranks[st.current_rank].steps_seen.setdefault("login_ok", i)
        return

    if RE["p1_form_missed"].search(line) and st.current_rank:
        st.add_signal(Signal("FAIL", "login_form_missed",
                             f"{st.current_rank}번 로그인 폼 미발견 → 사업자 스킵", i))
        return

    if RE["p1_search"].search(line) and st.current_rank:
        st.ranks[st.current_rank].steps_seen.setdefault("search", i)
        return

    if RE["p1_mybox_add"].search(line) and st.current_rank:
        st.ranks[st.current_rank].steps_seen.setdefault("mybox_add", i)
        return

    if RE["p1_speedgo_ok"].search(line) and st.current_rank:
        st.ranks[st.current_rank].steps_seen.setdefault("speedgo", i)
        return

    m = RE["p1_mybox_rows"].search(line)
    if m and st.current_rank:
        rows = int(m.group(1))
        st.ranks[st.current_rank].mybox_rows = rows
        if rows == 0:
            st.add_signal(Signal("WARN", "mybox_zero",
                                 f"{st.current_rank}번 마이박스 행수 0 — 검색 결과/해시태그 누락 가능", i))
        return

    if RE["p1_excel_btn"].search(line) and st.current_rank:
        st.ranks[st.current_rank].steps_seen.setdefault("excel_btn", i)
        return

    if RE["p1_excel_fail"].search(line) and st.current_rank:
        st.ranks[st.current_rank].excel_fail_count += 1
        return

    if RE["p1_dl_fail"].search(line) and st.current_rank:
        st.add_signal(Signal("FAIL", "excel_download_fail",
                             f"{st.current_rank}번 엑셀 다운로드 트리거 실패 (60s timeout)", i))
        return

    if RE["p1_step_skip"].search(line) and st.current_rank:
        st.add_signal(Signal("FAIL", "step1_6_skip",
                             f"{st.current_rank}번 원본 엑셀 없음 → STEP1~6 스킵 (_최종.xlsx 미생성)",
                             i))
        return

    m = RE["p1_seg_end"].search(line)
    if m:
        rank = int(m.group(1))
        rs = st.ranks.setdefault(rank, RankState(rank=rank, started_line=i))
        rs.seg_ended_line = i
        rs.steps_seen.setdefault("seg_end", i)
        return

    m = RE["pages_count"].search(line)
    if m:
        n = int(m.group(1))
        if n > st.pages_observed_max:
            st.pages_observed_max = n
        if n >= 5:
            st.add_signal(Signal("WARN", "pages_creep",
                                 f"탭 개수 {n} (정상 ≤ 3) — stale 탭 누적 의심", i))
        return

    # Phase 2
    m = RE["p2_chrome_up"].search(line)
    if m:
        st.phase = "p2"
        st.add_signal(Signal("INFO", "p2_up", f"Phase2 Chrome 기동 pid={m.group(1)} port={m.group(2)}", i))
        return

    m = RE["p2_targets"].search(line)
    if m:
        st.p2_targets = [int(x.strip()) for x in m.group(1).split(",") if x.strip().isdigit()]
        return

    m = RE["p2_cookie"].search(line)
    if m:
        # uid 가 등장하면 그 사용자가 작업 대상
        return

    # Phase 3
    m = RE["p3_chrome_up"].search(line)
    if m:
        st.phase = "p3"
        st.add_signal(Signal("INFO", "p3_up", f"Phase3 Chrome 기동 pid={m.group(1)}", i))
        return

    m = RE["p3_targets"].search(line)
    if m:
        st.p3_targets = [int(x.strip()) for x in m.group(1).split(",") if x.strip().isdigit()]
        return

    if RE["exception"].search(line):
        # 너무 시끄러우니 첫 줄만
        st.add_signal(Signal("WARN", "exception", line.strip()[:160], i))
        return


def _post_analyze(st: WatchState, lines: List[str]):
    """파싱 후 사업자별 누락 단계·이상 패턴 검사."""
    # P1 사업자별 누락 단계
    sorted_ranks = sorted(st.ranks)
    first_rank = sorted_ranks[0] if sorted_ranks else None
    if st.phase == "p1" or any(rs.started_line for rs in st.ranks.values()):
        for rank in sorted_ranks:
            rs = st.ranks[rank]
            # 종결되지 않은 사업자는 진행 중일 수 있으니 통과
            if rs.seg_ended_line is None:
                continue
            for k, label in P1_EXPECTED_STEPS:
                if k == "seg_end" or k in rs.steps_seen:
                    continue
                # 첫 사업자에는 stale 탭이 자연히 없으므로 누락 아님
                if k == "stale_clean" and rank == first_rank:
                    continue
                st.add_signal(Signal("FAIL", f"p1_missing_{k}",
                                     f"{rank}번 P1 누락 단계: {label}",
                                     rs.seg_ended_line, rank))

    # 동일 사업자가 헤더 2회 이상 등장 시 위에서 잡힘. 추가: 동일 user_id 가 다른 rank 에서 등장하면 혼선.
    uid_to_rank = {}
    for rank, rs in st.ranks.items():
        if not rs.user_id:
            continue
        if rs.user_id in uid_to_rank and uid_to_rank[rs.user_id] != rank:
            st.add_signal(Signal("FAIL", "uid_cross",
                                 f"같은 user_id={rs.user_id} 가 {uid_to_rank[rs.user_id]}번과 {rank}번에 동시 등장",
                                 rs.started_line, rank))
        else:
            uid_to_rank[rs.user_id] = rank


def render_human(st: WatchState, log_path: Path) -> str:
    out = []
    out.append(f"╔══ Phase Watchdog Report")
    out.append(f"║ log : {log_path}")
    out.append(f"║ job : {st.job_name or '(미정)'}")
    out.append(f"║ phase=  {st.phase or '?'}   step= {st.current_step_name or '?'}")
    out.append(f"║ result= {st.final_result or 'running'}   {st.started_at} → {st.finished_at or '진행 중'}")
    out.append(f"║ tabs(max observed)= {st.pages_observed_max}")
    out.append(f"╠══ 사업자별 진행")
    for rank in sorted(st.ranks):
        rs = st.ranks[rank]
        ticks = []
        for k, label in P1_EXPECTED_STEPS:
            ticks.append(("✓" if k in rs.steps_seen else "·") + label[:6])
        end = "DONE" if rs.seg_ended_line else "..."
        out.append(f"║ {rank}번 [{rs.user_id or '?':<10}] kw={rs.keyword or '?':<10}  rows={rs.mybox_rows if rs.mybox_rows is not None else '?'} {' '.join(ticks)} {end}")
    if st.p2_targets:
        out.append(f"╠══ Phase2 대상: {st.p2_targets}")
    if st.p3_targets:
        out.append(f"╠══ Phase3 대상: {st.p3_targets}")
    fails = [s for s in st.signals if s.level == "FAIL"]
    warns = [s for s in st.signals if s.level == "WARN"]
    infos = [s for s in st.signals if s.level == "INFO"]
    out.append(f"╠══ Signals: 🔴 FAIL={len(fails)}  🟡 WARN={len(warns)}  🟢 INFO={len(infos)}")
    for s in fails + warns:
        marker = "🔴" if s.level == "FAIL" else "🟡"
        rank = f" {s.rank}번" if s.rank else ""
        out.append(f"║   {marker} L{s.line_no:>5}{rank} [{s.code}] {s.msg}")
    out.append("╚══")
    return "\n".join(out)


def render_json(st: WatchState, log_path: Path) -> str:
    d = dict(
        log=str(log_path),
        job=st.job_name,
        started=st.started_at, finished=st.finished_at, result=st.final_result,
        phase=st.phase, step=st.current_step_name,
        current_rank=st.current_rank,
        pages_max=st.pages_observed_max,
        p2_targets=st.p2_targets, p3_targets=st.p3_targets,
        ranks={
            r: dict(user_id=rs.user_id, keyword=rs.keyword, mybox_rows=rs.mybox_rows,
                    steps=list(rs.steps_seen), seg_ended=bool(rs.seg_ended_line),
                    excel_fail=rs.excel_fail_count)
            for r, rs in st.ranks.items()
        },
        signals=[s.to_dict() for s in st.signals],
    )
    return json.dumps(d, ensure_ascii=False, indent=2)


def main():
    ap = argparse.ArgumentParser(description="Phase 1/2/3 watchdog (log → 구조화 분석)")
    ap.add_argument("log", nargs="?", help="logs/panel_*.log 경로. 미지정시 최신")
    ap.add_argument("--json", action="store_true", help="JSON 만 출력")
    ap.add_argument("--watch", type=int, default=0, help="N초마다 재분석 반복 (0=1회)")
    args = ap.parse_args()

    path = Path(args.log) if args.log else _pick_latest_log()
    if path is None:
        print("[watchdog] logs/panel_*.log 없음", file=sys.stderr)
        sys.exit(2)

    def _once():
        st = parse_log(path)
        if args.json:
            print(render_json(st, path))
        else:
            print(render_human(st, path))
        return st

    if args.watch <= 0:
        _once()
        return

    while True:
        _once()
        try:
            time.sleep(max(2, args.watch))
        except KeyboardInterrupt:
            break


if __name__ == "__main__":
    main()
