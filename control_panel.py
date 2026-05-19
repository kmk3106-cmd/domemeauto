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
P2 = [PY, "-u", "run_phase2.py"]
P3 = [PY, "-u", "run_phase3.py"]
# 고정 버튼: (표시명, [(단계명, cmd|None, env추가dict), ...])
JOBS = {
    "p1":   ("Phase 1만", [("Phase 1", P1, {})]),
    "p2":   ("Phase 2만", [("Phase 2", P2, {})]),
    "p3":   ("Phase 3만 (공급사판매중지 삭제)", [("Phase 3", P3, {})]),
    "p1_2": ("Phase 1~2 일괄", [("Phase 1", P1, {}), ("Phase 2", P2, {})]),
    "p1_3": ("Phase 1~3 일괄", [("Phase 1", P1, {}), ("Phase 2", P2, {}), ("Phase 3", P3, {})]),
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
    runs = []
    # 1~7회차(RUNS_PER_WEEK) 전부 항상 표시. 폴더 없는 회차는 wr_exists=False(미실행)
    for wr_no in range(1, RUNS_PER_WEEK + 1):
        wd = base / f"{wr_no}회차"
        wr_exists = wd.is_dir()
        rows = []
        for n in range(1, min(n_acc, 6) + 1):
            biz_id = ACCOUNTS[n - 1] if n - 1 < len(ACCOUNTS) else f"사업자{n}"
            bf = wd / f"{n}번사업자"
            if wr_exists and bf.is_dir():
                steps, p1_done = _p1_steps(bf, n)
                sent = bf / ".phase2_sent"
                p2_at = ""
                if sent.exists():
                    try:
                        p2_at = sent.read_text(encoding="utf-8").strip().split("\t")[0]
                    except Exception:
                        p2_at = "기록됨"
                rows.append({"rank": n, "biz_id": biz_id, "exists": True,
                             "steps": steps, "p1_done": p1_done,
                             "p2_sent": bool(p2_at), "p2_at": p2_at})
            else:
                rows.append({"rank": n, "biz_id": biz_id, "exists": False,
                             "steps": {}, "p1_done": False, "p2_sent": False, "p2_at": ""})
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
   <button class="b2" onclick="run('p2')">▶ Phase 2만</button>
   <button class="b3" onclick="if(confirm('Phase 3: 공급사판매중지 상품을 전 마켓에서 영구 삭제합니다 (6사업자). 진행?'))run('p3')">▶ Phase 3만 <span class="sub">(공급사판매중지 삭제)</span></button>
   <button class="bb" onclick="run('p1_2')">⏩ 1~2 일괄</button>
   <button class="bb" onclick="if(confirm('1~3 일괄: 3단계에서 공급사판매중지 상품을 전 마켓 영구 삭제합니다. 진행?'))run('p1_3')">⏩ 1~3 일괄</button>
   <button class="bs" onclick="stop()">■ 중단</button>
  </div>
  <div style="margin-top:10px"><span id="pill" class="pill p-idle">대기</span>
   <span id="jn" class="sub"></span><div class="sub" id="meta"></div></div>
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
  // 진척표
  const steps=['다운로드','STEP1합치기','STEP2이미지','STEP3키워드','STEP4상품명','STEP5링크','STEP6카테고리','최종'];
  let h='';
  d.runs.forEach(run=>{
   var tag = run.wr_exists ? '' : ' <span class=x style="font-weight:400">(폴더없음·미실행)</span>';
   h+='<h3>'+run.week_run+'회차'+tag+'</h3><table><thead><tr><th>사업자</th><th>계정</th>';
   steps.forEach(s=>h+='<th>'+s+'</th>');
   h+='<th>P1완료</th><th>P2전송</th></tr></thead><tbody>';
   run.rows.forEach(r=>{
    h+='<tr><td>'+r.rank+'번</td><td>'+r.biz_id+'</td>';
    steps.forEach(s=>h+='<td>'+(r.exists?mark(r.steps[s]):'<span class=x>·</span>')+'</td>');
    h+='<td>'+(r.p1_done?'<span class=y>완료</span>':'<span class=x>미완</span>')+'</td>';
    h+='<td>'+(r.p2_sent?('<span class=y>'+(r.p2_at||'전송')+'</span>'):'<span class=x>미전송</span>')+'</td></tr>';
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
