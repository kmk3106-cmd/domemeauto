#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
로컬 제어판 (버튼 실행 + 진척도 + 선택 실행).  실행: 패널실행.bat → http://localhost:8001/

기능:
  1) 버튼 실행: Phase 1만 / 2만 / 3만(미구현) / 1~2 일괄 / 1~3 일괄 / 중단
  2) 진척도: 사업자별·회차별로 어느 STEP/Phase 까지 됐는지 (파일 스캔 기반)
  3) 선택 실행: 회차 + 미진 사업자만 골라 Phase 1/2 따로 구동
한 번에 하나의 작업만 실행(겹침 방지). 로그 logs/, 실시간 표시. UTF-8·무버퍼.
"""
import os
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

for _sn in ("stdout", "stderr"):
    try:
        getattr(sys, _sn).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    from flask import Flask, jsonify, render_template_string, request
except ImportError:
    print("pip install flask 후 실행하세요.")
    raise

PROJECT_DIR = Path(__file__).resolve().parent
LOG_DIR = PROJECT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
sys.path.insert(0, str(PROJECT_DIR))

try:
    from domeme_auto_login_temp import EXCEL_SAVE_BASE, ACCOUNTS, get_upload_path_from_state
except Exception as _e:
    EXCEL_SAVE_BASE, ACCOUNTS = Path("."), []
    def get_upload_path_from_state():
        return "", 1
    print(f"[panel] domeme import 경고: {_e}")

_VENV_PY = Path(r"C:\Users\USER\PycharmProjects\PythonProject\.venv\Scripts\python.exe")
PY = str(_VENV_PY) if _VENV_PY.exists() else "python"

_STDIN_FEED = PROJECT_DIR / "_panel_stdin.txt"
if not _STDIN_FEED.exists():
    _STDIN_FEED.write_text("\n\n\n\n\n\n", encoding="ascii")

P1 = [PY, "-u", "domeme_auto_login_temp.py"]
P1_FILL = [PY, "-u", "run_phase1_fill.py"]  # 회차 채우기(자동 재시도) 래퍼
P2 = [PY, "-u", "run_phase2.py"]
P3 = [PY, "-u", "run_phase3.py"]
# 고정 버튼: (표시명, [(단계명, cmd|None, env추가dict), ...])
JOBS = {
    "p1":      ("Phase 1만", [("Phase 1", P1, {})]),
    "p1_fill": ("Phase 1 채우기 (회차 완성까지 자동 재시도)", [("Phase 1 채우기", P1_FILL, {})]),
    "p2":      ("Phase 2만", [("Phase 2", P2, {})]),
    "p3":      ("Phase 3만 (공급사판매중지 삭제)", [("Phase 3", P3, {})]),
    "p1_2":    ("Phase 1~2 일괄", [("Phase 1", P1, {}), ("Phase 2", P2, {})]),
    "p1_3":    ("Phase 1~3 일괄", [("Phase 1", P1, {}), ("Phase 2", P2, {}), ("Phase 3", P3, {})]),
    "pfill_2_3": ("Phase 채우기→2→3", [("Phase 1 채우기", P1_FILL, {}), ("Phase 2", P2, {}), ("Phase 3", P3, {})]),
}

app = Flask(__name__)
STATE = {"running": False, "job_name": "", "step": "", "started": "", "finished": "",
         "result": "idle", "log_path": None, "pid": None}
_LOCK = threading.Lock()
_CUR = {"p": None}


def _env(extra):
    e = dict(os.environ)
    e["PYTHONUTF8"] = "1"
    e["PYTHONIOENCODING"] = "utf-8"
    e.update({k: str(v) for k, v in (extra or {}).items()})
    return e


# === Phase 간 Chrome 완전 격리 ===
# 각 Phase 런처는 '자기 마커'(chrome_phase{1,2,3}_cdp)에 일치하는 Chrome 만 정리한다.
# 그러나 어떤 Phase 가 비정상 종료되거나(중단·예외), 직전 Phase Chrome 의 process tree 가
# OS 에서 완전히 사라지기 전에 다음 Phase 가 같은 포트로 띄우려 하면 충돌 가능.
# → 패널이 step 진입 전에 1·2·3 phase 마커 Chrome 전부를 강제 정리해 사용자 직관대로
#    "Phase 단계 사이에 Chrome 완전히 끄고 다시 시작"을 보장한다.
_PHASE_CHROME_MARKERS = ("chrome_phase1_cdp", "chrome_phase2_cdp", "chrome_phase3_cdp")


def _kill_all_phase_chromes(lf=None):
    try:
        import psutil
    except ImportError:
        if lf is not None:
            lf.write("[정리] psutil 미설치 → Chrome 사전 정리 스킵\n"); lf.flush()
        return 0
    killed = 0
    for p in psutil.process_iter(["pid", "name", "cmdline"]):
        try:
            if "chrome" not in (p.info.get("name") or "").lower():
                continue
            cmd = " ".join(str(c or "") for c in (p.info.get("cmdline") or [])).lower()
            if any(m in cmd for m in _PHASE_CHROME_MARKERS):
                p.kill()
                killed += 1
        except Exception:
            continue
    if lf is not None and killed:
        lf.write(f"[정리] 잔존 phase Chrome {killed}개 강제종료(다음 단계 시작 전)\n"); lf.flush()
    return killed


def _run_sequence(job_name, steps):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"panel_{ts}.log"
    with _LOCK:
        STATE.update(running=True, job_name=job_name, step="",
                     started=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                     finished="", result="running", log_path=str(log_path), pid=None)
    ok = True
    with open(log_path, "w", encoding="utf-8") as lf:
        lf.write(f"=== {job_name} 시작 {STATE['started']} ===\n"); lf.flush()
        for step_name, cmd, extra in steps:
            if STATE["result"] == "stopped":
                break
            STATE["step"] = step_name
            if cmd is None:
                lf.write(f"\n--- {step_name}: 미구현 — 스킵 ---\n"); lf.flush()
                continue
            lf.write(f"\n--- {step_name} 실행: {' '.join(cmd)} | env={extra} ---\n"); lf.flush()
            # 다음 단계 시작 전 phase chrome 잔존 정리(Phase1/2/3 모두) +
            # 포트가 OS 에서 해제될 시간을 1초 정도 확보. 같은 phase 의 첫 호출엔 영향 0.
            try:
                _killed_pre = _kill_all_phase_chromes(lf)
                if _killed_pre:
                    time.sleep(2)
            except Exception as _ke:
                lf.write(f"[정리] 사전 정리 예외(무시): {_ke}\n"); lf.flush()
            try:
                with open(_STDIN_FEED, "r", encoding="ascii") as fin:
                    proc = subprocess.Popen(cmd, cwd=str(PROJECT_DIR), env=_env(extra),
                                            stdin=fin, stdout=lf, stderr=subprocess.STDOUT)
                _CUR["p"] = proc; STATE["pid"] = proc.pid
                rc = proc.wait(); _CUR["p"] = None
                lf.write(f"\n--- {step_name} 종료코드 {rc} ---\n"); lf.flush()
                if rc != 0:
                    ok = False
                    lf.write(f"[중단] {step_name} 실패(rc={rc}) → 이후 단계 생략\n"); break
            except Exception as e:
                ok = False; lf.write(f"[예외] {step_name}: {e}\n"); break
        final = "stopped" if STATE["result"] == "stopped" else ("success" if ok else "failed")
        lf.write(f"\n=== 종료: {final} {datetime.now():%Y-%m-%d %H:%M:%S} ===\n")
    with _LOCK:
        STATE.update(running=False, step="", result=final,
                     finished=datetime.now().strftime("%Y-%m-%d %H:%M:%S"), pid=None)


def _start(job_name, steps):
    with _LOCK:
        if STATE["running"]:
            return False, f"이미 실행 중: {STATE['job_name']}"
    threading.Thread(target=_run_sequence, args=(job_name, steps), daemon=True).start()
    time.sleep(0.3)
    return True, f"{job_name} 시작"


@app.route("/run/<job_key>")
def run(job_key):
    if job_key not in JOBS:
        return jsonify({"ok": False, "msg": "알 수 없는 작업"}), 400
    name, steps = JOBS[job_key]
    ok, msg = _start(name, steps)
    return jsonify({"ok": ok, "msg": msg}), (200 if ok else 409)


@app.route("/run_sel")
def run_sel():
    phase = request.args.get("phase", "")
    try:
        wr = int(request.args.get("wr", "0"))
    except ValueError:
        wr = 0
    ranks = request.args.get("ranks", "").replace(" ", "")
    rset = sorted({int(x) for x in ranks.split(",") if x.isdigit() and 1 <= int(x) <= 6})
    if not rset:
        return jsonify({"ok": False, "msg": "사업자를 1명 이상 선택하세요"}), 400
    rcsv = ",".join(map(str, rset))
    needs_wr = phase in ("p1", "p2", "p1_2", "p1_3")  # Phase3 는 회차 무관
    if needs_wr and not (1 <= wr <= 7):
        return jsonify({"ok": False, "msg": "Phase1/2 는 회차(1~7)를 선택하세요"}), 400
    p1_env = {"WEEK_RUN": str(wr), "ONLY_RANKS": rcsv}
    p2_cmd = [PY, "-u", "run_phase2.py", "--week-run", str(wr), "--ranks", rcsv]
    p3_cmd = [PY, "-u", "run_phase3.py", "--ranks", rcsv]
    if phase == "p1":
        steps = [(f"Phase1 {wr}회차 {rcsv}번", P1, p1_env)]
    elif phase == "p2":
        steps = [(f"Phase2 {wr}회차 {rcsv}번", p2_cmd, {})]
    elif phase == "p3":
        steps = [(f"Phase3 {rcsv}번(공급사판매중지)", p3_cmd, {})]
    elif phase == "p1_2":
        steps = [(f"Phase1 {wr}회차 {rcsv}번", P1, p1_env),
                 (f"Phase2 {wr}회차 {rcsv}번", p2_cmd, {})]
    elif phase == "p1_3":
        steps = [(f"Phase1 {wr}회차 {rcsv}번", P1, p1_env),
                 (f"Phase2 {wr}회차 {rcsv}번", p2_cmd, {}),
                 (f"Phase3 {rcsv}번", p3_cmd, {})]
    else:
        return jsonify({"ok": False, "msg": "phase=p1|p2|p3|p1_2|p1_3"}), 400
    name = f"선택실행 {phase} · {('('+str(wr)+'회차) ') if needs_wr else ''}{rcsv}번"
    ok, msg = _start(name, steps)
    return jsonify({"ok": ok, "msg": msg}), (200 if ok else 409)


@app.route("/stop")
def stop():
    proc = _CUR.get("p")
    if not STATE["running"] or proc is None:
        return jsonify({"ok": False, "msg": "실행 중 작업 없음"})
    STATE["result"] = "stopped"
    try:
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
    except Exception as e:
        return jsonify({"ok": False, "msg": f"중단 실패: {e}"})
    return jsonify({"ok": True, "msg": "중단 요청됨 (프로세스 트리 종료)"})


# 파일명 {user_id}_{kw_tag}_{YY년M월W주차}_{N}회(_최종)?.xlsx 에서 kw_tag 추출
import re as _re
_KW_RE = _re.compile(r"^([^_]+)_(.+?)_\d{2}년\d{1,2}월\d+주차_\d+회(?:_최종)?\.xlsx$")
_INTER = ("통합상품명", "domeme_links", "카테고리매핑", "keywords")


def _extract_kw_from_folder(folder: Path) -> str:
    """사업자 폴더에서 키워드 추출. _최종.xlsx 우선, 없으면 원본 mybox 다운로드 xlsx 사용.
    중간 STEP 산출물(_통합상품명/_keywords/_domeme_links/_카테고리매핑)은 제외."""
    if not folder.is_dir():
        return ""
    finals = sorted(folder.glob("*_최종.xlsx"),
                    key=lambda p: p.stat().st_mtime, reverse=True)
    others = [p for p in folder.glob("*.xlsx")
              if "_최종" not in p.name and not any(k in p.name for k in _INTER)]
    for p in finals + others:
        m = _KW_RE.match(p.name)
        if m:
            return m.group(2)
    return ""


def _extract_p2_kw_from_marker(folder: Path) -> str:
    """.phase2_sent 마커의 파일명에서 kw_tag 추출 (Phase2가 실제 업로드에 사용한 _최종 파일)."""
    sent = folder / ".phase2_sent"
    if not sent.exists():
        return ""
    try:
        line = sent.read_text(encoding="utf-8").strip()
        parts = line.split("\t")
        fn = parts[1] if len(parts) >= 2 else ""
        m = _KW_RE.match(fn)
        return m.group(2) if m else ""
    except Exception:
        return ""


def _p1_steps(folder: Path, n: int):
    P = f"{n}번"
    inter = ("통합상품명", "_최종", "domeme_links", "카테고리매핑", "keywords")
    src = [p for p in folder.glob("*.xls*") if not any(k in p.name for k in inter)]
    finals = list(folder.glob("*_최종.xlsx"))
    steps = {
        "다운로드": bool(src),
        "STEP1합치기": (folder / f"{P}_통합상품명.xlsx").exists(),
        "STEP2이미지": (folder / f"{P}_통합상품명_이미지").exists(),
        "STEP3키워드": (folder / f"{P}_통합상품명_keywords.xlsx").exists(),
        "STEP4상품명": (folder / f"{P}_통합상품명_keywords_쭌쭌쌤.xlsx").exists(),
        "STEP5링크": (folder / f"{P}_domeme_links.xlsx").exists(),
        "STEP6카테고리": (folder / f"{P}_통합상품명_카테고리매핑_결과.xlsx").exists(),
        "최종": bool(finals),
    }
    return steps, bool(finals)


@app.route("/progress")
def progress():
    try:
        from domeme_auto_login_temp import RUNS_PER_WEEK
    except Exception:
        RUNS_PER_WEEK = 7
    ymw_str, _wr = get_upload_path_from_state()
    base = Path(EXCEL_SAVE_BASE) / ymw_str
    n_acc = max(6, len(ACCOUNTS))
    # Phase 3 결과 마커 로드 (회차 무관, 사업자 rank 키)
    p3_state = {}
    try:
        import json as _json
        p3_path = PROJECT_DIR / "phase3_state.json"
        if p3_path.exists():
            p3_state = _json.loads(p3_path.read_text(encoding="utf-8")) or {}
    except Exception:
        p3_state = {}
    runs = []
    for wr_no in range(1, RUNS_PER_WEEK + 1):
        wd = base / f"{wr_no}회차"
        wr_exists = wd.is_dir()
        rows = []
        for n in range(1, min(n_acc, 6) + 1):
            biz_id = ACCOUNTS[n - 1] if n - 1 < len(ACCOUNTS) else f"사업자{n}"
            bf = wd / f"{n}번사업자"
            p3 = p3_state.get(str(n)) or {}
            p3_payload = {"p3_result": p3.get("result", ""), "p3_at": p3.get("ts", ""),
                          "p3_before": p3.get("before", -1), "p3_after": p3.get("after", -1)}
            if wr_exists and bf.is_dir():
                steps, p1_done = _p1_steps(bf, n)
                sent = bf / ".phase2_sent"
                p2_at = ""
                if sent.exists():
                    try:
                        p2_at = sent.read_text(encoding="utf-8").strip().split("\t")[0]
                    except Exception:
                        p2_at = "기록됨"
                kw = _extract_kw_from_folder(bf)
                p2_kw = _extract_p2_kw_from_marker(bf)
                p1_incomplete = [k for k, ok in steps.items() if not ok] if not p1_done else []
                kw_match = None
                if kw and p2_kw:
                    kw_match = (kw == p2_kw)
                rows.append({"rank": n, "biz_id": biz_id, "exists": True,
                             "kw": kw, "p2_kw": p2_kw, "kw_match": kw_match,
                             "p1_done": p1_done, "p1_incomplete": p1_incomplete,
                             "p2_sent": bool(p2_at), "p2_at": p2_at, **p3_payload})
            else:
                rows.append({"rank": n, "biz_id": biz_id, "exists": False,
                             "kw": "", "p2_kw": "", "kw_match": None,
                             "p1_done": False, "p1_incomplete": [],
                             "p2_sent": False, "p2_at": "",
                             **p3_payload})
        runs.append({"week_run": wr_no, "wr_exists": wr_exists, "rows": rows})
    return jsonify({"ymw": ymw_str, "runs": runs,
                    "updated": datetime.now().strftime("%H:%M:%S")})


@app.route("/status")
def status():
    tail = ""
    lp = STATE.get("log_path")
    if lp and Path(lp).exists():
        try:
            tail = "\n".join(Path(lp).read_text(encoding="utf-8", errors="replace")
                             .splitlines()[-80:])
        except Exception as e:
            tail = f"(로그 읽기 실패: {e})"
    s = dict(STATE); s["log_tail"] = tail
    return jsonify(s)


@app.route("/watchdog")
def watchdog():
    """phase_watchdog.parse_log 결과를 JSON 으로 반환.
    sub-agent · 패널 UI · 외부 도구 모두 이 endpoint 를 호출해 일관된 형태로 신호를 받는다.
    """
    try:
        from phase_watchdog import parse_log, render_json, _pick_latest_log
    except Exception as e:
        return jsonify({"ok": False, "msg": f"watchdog import 실패: {e}"}), 500
    lp_q = request.args.get("log")
    path = Path(lp_q) if lp_q else (Path(STATE.get("log_path")) if STATE.get("log_path") else _pick_latest_log())
    if path is None or not Path(path).exists():
        return jsonify({"ok": False, "msg": f"로그 파일 없음: {path}"}), 404
    try:
        st_w = parse_log(Path(path))
        return app.response_class(render_json(st_w, Path(path)),
                                  mimetype="application/json; charset=utf-8")
    except Exception as e:
        return jsonify({"ok": False, "msg": f"watchdog 분석 실패: {e}"}), 500


PANEL_HTML = r"""
<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>실행 제어판</title>
<meta name="viewport" content="width=device-width, initial-scale=1"><style>
 body{font-family:'Malgun Gothic',sans-serif;margin:0;background:#0f1115;color:#e6e6e6}
 header{padding:14px 22px;background:#161a22;border-bottom:1px solid #2a2f3a}
 h1{font-size:17px;margin:0}.sub{color:#8b93a3;font-size:12px}
 .wrap{padding:18px;max-width:1180px;margin:0 auto}
 .box{background:#161a22;border:1px solid #2a2f3a;border-radius:10px;padding:14px;margin-bottom:14px}
 .btns{display:flex;flex-wrap:wrap;gap:10px}
 button{font-size:13px;font-weight:700;padding:11px 15px;border-radius:9px;border:1px solid #2f3644;
  background:#1d2330;color:#e6e6e6;cursor:pointer}button:hover{background:#252c3b}
 button:disabled{opacity:.4;cursor:not-allowed}
 .b1{border-color:#2f6df6}.b2{border-color:#39d98a}.b3{border-color:#7d8696}
 .bb{border-color:#f6a52f}.bs{border-color:#ff5b6e;color:#ff8a96}
 .pill{display:inline-block;padding:3px 11px;border-radius:13px;font-weight:700;font-size:12px}
 .p-run{background:#10324a;color:#4ab3ff}.p-ok{background:#15301f;color:#39d98a}
 .p-fail{background:#3a1620;color:#ff5b6e}.p-idle{background:#222732;color:#9aa3b2}.p-stop{background:#3a2a16;color:#f6a52f}
 table{width:100%;border-collapse:collapse;font-size:12px}
 th,td{padding:7px 8px;border-bottom:1px solid #232833;text-align:center}
 th{background:#1c212b;color:#9aa3b2}
 .y{color:#39d98a;font-weight:700}.x{color:#5b6472}
 pre{background:#0b0d12;border:1px solid #232833;border-radius:8px;padding:11px;height:240px;
  overflow:auto;font-size:12px;white-space:pre-wrap;color:#cdd3df}
 select,input{background:#1d2330;color:#e6e6e6;border:1px solid #2f3644;border-radius:7px;padding:7px}
 .rk{display:inline-flex;align-items:center;gap:4px;margin-right:10px}
 h3{margin:6px 0 10px;font-size:14px}
</style></head><body>
<header><h1>실행 제어판 <span class="sub">· Phase 1/2/3 · 진척도 · 선택실행</span></h1></header>
<div class="wrap">
 <div class="box"><h3>① 전체 실행</h3>
  <div class="btns">
   <button class="b1" onclick="run('p1')">▶ Phase 1만</button>
   <button class="b1" onclick="run('p1_fill')">▶ P1 채우기(회차완성)</button>
   <button class="b2" onclick="run('p2')">▶ Phase 2만</button>
   <button class="b3" onclick="if(confirm('Phase 3: 공급사판매중지 상품을 전 마켓에서 영구 삭제합니다 (6사업자). 진행?'))run('p3')">▶ Phase 3만 <span class="sub">(공급사판매중지 삭제)</span></button>
   <button class="bb" onclick="run('p1_2')">⏩ 1~2 일괄</button>
   <button class="bb" onclick="if(confirm('1~3 일괄: 3단계에서 공급사판매중지 상품을 전 마켓 영구 삭제합니다. 진행?'))run('p1_3')">⏩ 1~3 일괄</button>
   <button class="bb" onclick="if(confirm('채우기→2→3: P1 회차 채우기(자동 재시도) 후 2·3(영구삭제). 진행?'))run('pfill_2_3')">⏩ 채우기→2→3</button>
   <button class="bs" onclick="stop()">■ 중단</button>
   <button class="b3" onclick="openWatchdog()" title="phase_watchdog: 현재/최근 로그 면밀 점검">📊 점검</button>
  </div>
  <div style="margin-top:10px"><span id="pill" class="pill p-idle">대기</span>
   <span id="jn" class="sub"></span><div class="sub" id="meta"></div></div>
  <pre id="wd" style="display:none;margin-top:10px;background:#0b0e14;padding:10px;border-radius:8px;font-size:11px;line-height:1.4;color:#cdd3df;overflow:auto;max-height:480px"></pre>
 </div>

 <div class="box"><h3>② 선택 실행 (미진 사업자만)</h3>
  <div style="display:flex;flex-wrap:wrap;gap:14px;align-items:center">
   <label>회차 <select id="selWr"></select></label>
   <span id="selRanks"></span>
   <button class="b1" onclick="runSel('p1')">선택 Phase1만</button>
   <button class="b2" onclick="runSel('p2')">선택 Phase2만</button>
   <button class="b3" onclick="if(confirm('선택 사업자 Phase3: 공급사판매중지 상품을 전 마켓 영구삭제합니다. 진행?'))runSel('p3')">선택 Phase3만</button>
   <button class="bb" onclick="runSel('p1_2')">선택 1~2</button>
   <button class="bb" onclick="if(confirm('선택 사업자 1~3: 3단계에서 공급사판매중지 영구삭제. 진행?'))runSel('p1_3')">선택 1~3</button>
  </div>
  <div class="sub" style="margin-top:8px">체크 후 버튼 → 선택 사업자만 구동. Phase1=ONLY_RANKS·Phase2=--ranks (회차 사용) · <b>Phase3=--ranks (회차 무관, 공급사판매중지 전체)</b></div>
 </div>

 <div class="box"><h3>③ 진척도 <span class="sub" id="pgmeta"></span></h3>
  <div id="prog">불러오는 중…</div>
 </div>

 <div class="box"><h3>실행 로그</h3>
  <pre id="log">실행 버튼을 누르면 로그 표시</pre>
  <div class="sub">로그: <span id="lp">-</span> · 자동 새로고침 3초</div>
 </div>
</div>
<script>
async function run(k){const r=await fetch('/run/'+k);const d=await r.json();if(!d.ok)alert(d.msg);refresh();}
async function stop(){const r=await fetch('/stop');const d=await r.json();alert(d.msg);refresh();}
async function openWatchdog(){
 const el=document.getElementById('wd');
 el.style.display='block';el.textContent='watchdog 분석 중…';
 try{
  const r=await fetch('/watchdog');const j=await r.json();
  if(!r.ok || j.ok===false){el.textContent='[오류] '+(j.msg||r.status);return;}
  const fails=(j.signals||[]).filter(s=>s.level==='FAIL');
  const warns=(j.signals||[]).filter(s=>s.level==='WARN');
  let s=`[Phase Monitor] ${j.phase||'?'} / ${j.step||'?'} · ${j.result||'?'}\n`;
  s+=`tabs(max)=${j.pages_max}  ranks=`;
  const rs=j.ranks||{};
  s+=Object.keys(rs).map(k=>`${k}번:${rs[k].seg_ended?'✓':'…'}`).join(' ')+`\n`;
  s+=`\n🔴 FAIL=${fails.length}  🟡 WARN=${warns.length}\n`;
  fails.forEach(x=>{s+=`  🔴 L${x.line} ${x.rank?x.rank+'번ㆍ':''}[${x.code}] ${x.msg}\n`;});
  warns.forEach(x=>{s+=`  🟡 L${x.line} ${x.rank?x.rank+'번ㆍ':''}[${x.code}] ${x.msg}\n`;});
  s+=`\nlog: ${j.log}`;
  el.textContent=s;
 }catch(e){el.textContent='[오류] '+e;}
}
async function runSel(phase){
 const wr=document.getElementById('selWr').value;
 const ranks=[...document.querySelectorAll('.selrk:checked')].map(c=>c.value).join(',');
 if(!ranks){alert('사업자를 선택하세요');return;}
 const r=await fetch('/run_sel?phase='+phase+'&wr='+wr+'&ranks='+ranks);
 const d=await r.json();alert(d.msg);refresh();
}
function pill(res){const m={running:['p-run','실행중'],success:['p-ok','성공'],
 failed:['p-fail','실패'],stopped:['p-stop','중단됨'],idle:['p-idle','대기']};return m[res]||m.idle;}
function mark(b){return b?'<span class=y>✓</span>':'<span class=x>·</span>';}
async function refresh(){
 try{
  const s=await(await fetch('/status')).json();
  const p=pill(s.result);const el=document.getElementById('pill');
  el.className='pill '+p[0];el.textContent=p[1];
  document.getElementById('jn').textContent=s.job_name?('· '+s.job_name+(s.step?(' · '+s.step):'')):'';
  document.getElementById('meta').textContent=(s.started?('시작 '+s.started):'')+(s.finished?('  종료 '+s.finished):'')+(s.pid?('  pid '+s.pid):'');
  document.getElementById('log').textContent=s.log_tail||'(로그 없음)';
  document.getElementById('lp').textContent=s.log_path||'-';
  document.querySelectorAll('button').forEach(b=>{if(!b.classList.contains('bs'))b.disabled=s.running;});
 }catch(e){}
}
async function loadProg(){
 try{
  const d=await(await fetch('/progress')).json();
  document.getElementById('pgmeta').textContent='주차='+d.ymw+' · '+d.updated;
  // 회차 셀렉트
  const sw=document.getElementById('selWr');
  if(sw.options.length===0||sw.dataset.ymw!==d.ymw){
   sw.innerHTML='';sw.dataset.ymw=d.ymw;
   const wrs=d.runs.map(r=>r.week_run);
   const opts=wrs.length?wrs:[1,2,3,4,5,6,7];
   opts.forEach(w=>{const o=document.createElement('option');o.value=w;o.textContent=w+'회차';sw.appendChild(o);});
  }
  // 사업자 체크박스
  let rk='';for(let i=1;i<=6;i++)rk+='<label class=rk><input type=checkbox class=selrk value='+i+'>'+i+'번</label>';
  document.getElementById('selRanks').innerHTML=rk;
  // 진척표 (간소화: 키워드 · P1완료 · P2전송 · P3 삭제)
  let h='';
  d.runs.forEach(run=>{
   var tag = run.wr_exists ? '' : ' <span class=x style="font-weight:400">(폴더없음·미실행)</span>';
   h+='<h3>'+run.week_run+'회차'+tag+'</h3><table><thead><tr>';
   h+='<th>사업자</th><th>계정</th><th>키워드(P1)</th><th>P1 완료</th><th>P2 전송 (키워드)</th><th>P3 삭제</th>';
   h+='</tr></thead><tbody>';
   run.rows.forEach(r=>{
    h+='<tr><td>'+r.rank+'번</td><td>'+r.biz_id+'</td>';
    // 키워드
    h+='<td>'+(r.kw?('<b>'+r.kw+'</b>'):'<span class=x>—</span>')+'</td>';
    // P1 완료: 완료 / 미완 + 어느 STEP 빠졌는지
    if(r.p1_done){
      h+='<td><span class=y>완료</span></td>';
    } else if(r.exists){
      var inc = (r.p1_incomplete||[]).join(' · ') || '진행 전';
      h+='<td><span class=alert>미완</span> <span class=sub style="font-size:11px">('+inc+')</span></td>';
    } else {
      h+='<td><span class=x>—</span></td>';
    }
    // P2 전송 + 키워드 일치 검증
    if(r.p2_sent){
      var p2kw = r.p2_kw || '?';
      var match = r.kw_match;
      var ind = (match===true)?'<span class=y> ✓일치</span>':
                (match===false)?'<span class=alert> ⚠불일치('+r.kw+'≠'+p2kw+')</span>':'';
      h+='<td><span class=y>전송 '+(r.p2_at||'')+'</span> <span class=sub>#'+p2kw+'</span>'+ind+'</td>';
    } else {
      h+='<td><span class=x>미전송</span></td>';
    }
    // P3 삭제 결과 매핑 (성공/잠금추정/의심/대상없음/실패/없음)
    var p3html='<span class=x>·</span>';
    var pr=r.p3_result||'';
    var sx={alert:'성공',count_drop:'성공',revert_drop:'성공',page_closed:'성공',
            revert_noop:'완료(잠금)',timeout_suspect:'⚠의심',no_target:'대상없음',
            no_login:'실패(로그인)',no_open:'실패(진입)',no_popup:'실패(팝업)',error:'실패(예외)'};
    if(pr){ var t=sx[pr]||pr; var cls=(['alert','count_drop','revert_drop','page_closed'].indexOf(pr)>=0)?'y':
       (pr==='timeout_suspect'||pr.indexOf('실패')===0||sx[pr]&&sx[pr].indexOf('실패')===0?'alert':
       (pr==='no_target'?'sub':'sub'));
       p3html='<span class='+cls+'>'+t+'</span><br><span class=sub style="font-size:10px">'+(r.p3_at||'')+'</span>'; }
    h+='<td>'+p3html+'</td></tr>';
   });
   h+='</tbody></table>';
  });
  document.getElementById('prog').innerHTML=h;
 }catch(e){document.getElementById('prog').textContent='진척도 조회 실패: '+e;}
}
refresh();loadProg();setInterval(refresh,3000);setInterval(loadProg,15000);
</script>
</body></html>
"""


@app.route("/")
def index():
    return render_template_string(PANEL_HTML)


if __name__ == "__main__":
    print("실행 제어판: http://localhost:8001/", flush=True)
    app.run(host="127.0.0.1", port=8001, debug=False, threaded=True)
