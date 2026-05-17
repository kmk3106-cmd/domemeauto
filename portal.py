#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통합 포털 (MVP): 6사업자 × 마켓별 OpenAPI 집계.
  - 출고중지 미인지건 (발송 전 인지 → 배송비 손해 방지)
  - 미응대 고객문의 (응대 누락 → 판매자점수 하락 방지)

실행:
  pip install flask requests python-dotenv
  python portal.py            →  http://localhost:8002/portal

자격증명 미설정 시에도 절대 죽지 않음 ("API 미연동" 표시). .env 에 키를 채우면 자동 연동.
"""
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

for _sn in ("stdout", "stderr"):
    try:
        getattr(sys, _sn).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

PROJECT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_DIR))

try:
    from flask import Flask, jsonify, render_template_string
except ImportError:
    print("pip install flask 후 실행하세요.")
    raise

from portal_connectors import build_connectors

try:
    from domeme_auto_login_temp import ACCOUNTS
except Exception:
    ACCOUNTS = []

app = Flask(__name__)
CONNECTORS = build_connectors(ACCOUNTS or [f"사업자{i}" for i in range(1, 7)])


def _collect_one(rank: int):
    biz_id = ACCOUNTS[rank - 1] if rank - 1 < len(ACCOUNTS) else f"사업자{rank}"
    holds, inquiries, markets_state = [], [], []
    for conn in CONNECTORS.get(rank, []):
        markets_state.append({"market": conn.market, "configured": conn.configured})
        if not conn.configured:
            continue
        try:
            holds += conn.shipment_holds()
        except Exception as e:
            print(f"[portal] {rank}번 {conn.market} 출고중지 조회 실패: {e}")
        try:
            inquiries += conn.unanswered_inquiries()
        except Exception as e:
            print(f"[portal] {rank}번 {conn.market} 문의 조회 실패: {e}")
    any_configured = any(m["configured"] for m in markets_state)
    return {
        "rank": rank,
        "biz_id": biz_id,
        "configured": any_configured,
        "markets": markets_state,
        "shipment_hold_count": len(holds),
        "inquiry_count": len(inquiries),
        "shipment_holds": holds,
        "inquiries": inquiries,
    }


@app.route("/api/portal")
def api_portal():
    ranks = sorted(CONNECTORS.keys())
    with ThreadPoolExecutor(max_workers=min(8, max(1, len(ranks)))) as pool:
        rows = list(pool.map(_collect_one, ranks))
    total_holds = sum(r["shipment_hold_count"] for r in rows)
    total_inq = sum(r["inquiry_count"] for r in rows)
    return jsonify({
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "rows": rows,
        "total_shipment_holds": total_holds,
        "total_inquiries": total_inq,
        "any_configured": any(r["configured"] for r in rows),
    })


PORTAL_HTML = """
<!doctype html><html lang="ko"><head><meta charset="utf-8">
<title>통합 포털</title><meta name="viewport" content="width=device-width, initial-scale=1">
<style>
 body{font-family:'Malgun Gothic',sans-serif;margin:0;background:#0f1115;color:#e6e6e6}
 header{padding:16px 22px;background:#161a22;border-bottom:1px solid #2a2f3a;display:flex;justify-content:space-between;align-items:center}
 h1{font-size:18px;margin:0}
 .sub{color:#8b93a3;font-size:12px}
 .wrap{padding:20px;max-width:1100px;margin:0 auto}
 .kpi{display:flex;gap:14px;margin-bottom:18px}
 .card{flex:1;background:#161a22;border:1px solid #2a2f3a;border-radius:10px;padding:16px}
 .card .n{font-size:30px;font-weight:700}
 .alert{color:#ff5b6e}.ok{color:#39d98a}
 table{width:100%;border-collapse:collapse;background:#161a22;border-radius:10px;overflow:hidden}
 th,td{padding:11px 12px;text-align:center;border-bottom:1px solid #232833;font-size:13px}
 th{background:#1c212b;color:#9aa3b2}
 tr:hover{background:#1b2029}
 .badge{display:inline-block;min-width:26px;padding:3px 8px;border-radius:12px;font-weight:700}
 .b-red{background:#3a1620;color:#ff5b6e}.b-grey{background:#222732;color:#7d8696}.b-green{background:#15301f;color:#39d98a}
 .nc{color:#7d8696;font-style:italic}
 footer{color:#5f6776;font-size:11px;padding:14px 22px}
</style></head><body>
<header><div><h1>통합 포털 <span class="sub">· 출고중지 / 미응대 문의 경보</span></h1></div>
<div class="sub">자동 새로고침 60초 · <span id="upd">-</span></div></header>
<div class="wrap">
 <div class="kpi">
  <div class="card"><div class="sub">전체 출고중지 미인지</div><div class="n" id="th">-</div><div class="sub">발송 전 즉시 처리 대상</div></div>
  <div class="card"><div class="sub">전체 미응대 문의</div><div class="n" id="ti">-</div><div class="sub">응대 누락 = 점수 하락</div></div>
 </div>
 <table><thead><tr><th>사업자</th><th>계정</th><th>출고중지 미인지</th><th>미응대 문의</th><th>연동 상태</th></tr></thead>
 <tbody id="tb"><tr><td colspan="5" class="nc">불러오는 중…</td></tr></tbody></table>
 <footer id="ft"></footer>
</div>
<script>
function badge(n){ if(n>0) return '<span class="badge b-red">'+n+'</span>';
 return '<span class="badge b-grey">0</span>'; }
async function load(){
 try{
  const r = await fetch('/api/portal'); const d = await r.json();
  document.getElementById('upd').textContent = d.updated;
  document.getElementById('th').innerHTML = (d.total_shipment_holds>0?'<span class=alert>':'<span class=ok>')+d.total_shipment_holds+'</span>';
  document.getElementById('ti').innerHTML = (d.total_inquiries>0?'<span class=alert>':'<span class=ok>')+d.total_inquiries+'</span>';
  const tb=document.getElementById('tb'); tb.innerHTML='';
  d.rows.forEach(row=>{
   const mk = row.markets.map(m=>m.market+(m.configured?'✓':'✗')).join(' ');
   const st = row.configured? mk : '<span class="nc">API 미연동</span>';
   tb.innerHTML += '<tr><td>'+row.rank+'번</td><td>'+row.biz_id+'</td><td>'+badge(row.shipment_hold_count)+'</td><td>'+badge(row.inquiry_count)+'</td><td>'+st+'</td></tr>';
  });
  document.getElementById('ft').textContent = d.any_configured? '' : '※ 아직 마켓 API 키가 .env 에 없습니다. 키를 채우면 자동으로 집계됩니다.';
 }catch(e){ document.getElementById('tb').innerHTML='<tr><td colspan=5 class=nc>조회 실패: '+e+'</td></tr>'; }
}
load(); setInterval(load, 60000);
</script>
</body></html>
"""


@app.route("/portal")
@app.route("/portal/")
def portal():
    return render_template_string(PORTAL_HTML)


@app.route("/")
def index():
    return '<a href="/portal">통합 포털로 이동</a>'


if __name__ == "__main__":
    print("통합 포털: http://localhost:8002/portal", flush=True)
    app.run(host="127.0.0.1", port=8002, debug=False)
