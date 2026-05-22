# -*- coding: utf-8 -*-
"""
Phase 3: 스피드고 '전송관리 > 공급사판매중지' 상품을 사업자별로 삭제.

확정 요건(2026-05-17):
  - 사업자 1~6 루핑. 각 사업자: 로그인 → 공급사판매중지 → 목록 비면 스킵
  - 500개씩 보기 → 전체선택(현재 페이지, 최대 500) → 삭제
  - 팝업 '상품삭제'에서 "활성화 된 마켓에서 모두 삭제" 선택 → 삭제
  - S 마크 뜨며 행 사라짐 = 완료. **딱 1배치(500)만, 반복 없음**
  - 모든 활성 마켓에서 영구 삭제 (되돌리기 불가 — 의도된 정리 작업)

로그인/네비게이션은 검증된 test_speedgo_upload_1번 의 동선·셀렉터를 재사용.
"""
import json
import os
import time
from pathlib import Path

from test_speedgo_upload_1번 import _goto_with_retry, DOMEME_URL, SPEEDGO_URL, _WAIT

_PROJECT_DIR = Path(__file__).resolve().parent
_PHASE3_STATE_FILE = _PROJECT_DIR / "phase3_state.json"


def _current_week_run():
    """이번 P3 가 어느 회차 사이클에서 돌았는지 식별.
    '한 회차 = 한 P1~P3 사이클' 이므로 P3 결과는 그 회차 행에만 표시돼야 한다.
    .week_run_state(get_upload_path_from_state) 의 현재 회차를 기록해 둔다."""
    # 1) 환경변수 우선 (제어판 선택실행이 WEEK_RUN 을 넘기는 경우)
    ev = os.environ.get("WEEK_RUN", "").strip()
    if ev.isdigit():
        return int(ev)
    # 2) .week_run_state
    try:
        from domeme_auto_login_temp import get_upload_path_from_state
        _ymw, wr = get_upload_path_from_state()
        return int(wr)
    except Exception:
        return None


def _write_phase3_marker(rank: int, biz_id: str, result: str,
                          before_cnt: int = -1, after_cnt: int = -1) -> None:
    """phase3_state.json 갱신: rank → {ts, biz_id, result, before, after, week_run}.
    week_run: 이 P3 가 실행된 회차(대시보드는 이 회차 행에만 P3 결과를 표시).
    result 코드: alert/count_drop/revert_drop/page_closed (성공) |
                 revert_noop (잠금-only 추정) | timeout_suspect (의심) |
                 no_target (대상없음) | no_login/no_open/no_popup/error (실패)."""
    data = {}
    try:
        if _PHASE3_STATE_FILE.exists():
            data = json.loads(_PHASE3_STATE_FILE.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                data = {}
    except Exception:
        data = {}
    data[str(rank)] = {
        "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
        "biz_id": biz_id,
        "result": result,
        "before": before_cnt,
        "after": after_cnt,
        "week_run": _current_week_run(),
    }
    _PHASE3_STATE_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )

_ID_SEL = ['input[name="userId"]', 'input[name="user_id"]', 'input[name="id"]',
           'input[name="loginId"]', 'input[id*="user"]', 'input[id*="id"]',
           'input[type="text"]', 'input#userId', 'input#user_id']
_PW_SEL = ['input[name="password"]', 'input[name="passwd"]', 'input[name="pw"]',
           'input[type="password"]']


def _find_login_form(pg):
    ctxs = [pg]
    try:
        ctxs += pg.frames[1:]
    except Exception:
        pass
    for ctx in ctxs:
        for s in _ID_SEL:
            try:
                el = ctx.locator(s).first
                if el.count() > 0 and el.is_visible():
                    for ps in _PW_SEL:
                        pe = ctx.locator(ps).first
                        if pe.count() > 0 and pe.is_visible():
                            return el, pe, ctx
            except Exception:
                continue
    return None, None, None


def _verify_logged_in_as(page, user_id, timeout_s: float = 4.0) -> bool:
    """로그인 후 페이지 텍스트에 user_id 가 보이면 본인 계정으로 확인. 베스트에포트."""
    import time as _t
    end = _t.time() + timeout_s
    while _t.time() < end:
        try:
            txt = page.evaluate(
                "() => (document.body ? (document.body.innerText||'') : '') + ' ' + (document.title||'')")
            if user_id and user_id in (txt or ""):
                return True
        except Exception:
            pass
        _t.sleep(0.4)
    return False


def _login(page, user_id, password, do_logout):
    """[수정] 사업자 전환 시 컨텍스트 쿠키를 명시적으로 삭제해 '이전 사업자 세션 잔류 → 잘못된 계정으로 진행' 위험 차단.
    이후 도매매로 fresh 진입 → 로그인 링크 → 폼 채움 → 제출 → user_id 매칭 확인.
    do_logout 은 항상 True 로 동작(과거 click-logout 의 비결정성 회피)."""
    try:
        page.context.clear_cookies()
        print(f"[로그인] 컨텍스트 쿠키 초기화 → 새 사업자 fresh 로그인 ({user_id})")
    except Exception as e:
        print(f"[로그인] 쿠키 초기화 실패(무시): {e}")
    if not _goto_with_retry(page, DOMEME_URL, "commit", 45000):
        return False
    try:
        page.wait_for_load_state("domcontentloaded", timeout=20000)
    except Exception:
        pass
    time.sleep(1)
    try:
        ll = page.get_by_role("link", name="로그인").first
        if ll.count() == 0:
            ll = page.locator('a:has-text("로그인"):not(:has-text("진행중"))').first
        if ll.count() == 0:
            ll = page.get_by_text("로그인", exact=True).first
        if ll.count() > 0 and ll.is_visible():
            ll.click(timeout=5000)
            print("로그인 링크 클릭")
    except Exception as e:
        print(f"로그인 링크: {e}")
    time.sleep(1.2)
    idf, pwf, ctx = _find_login_form(page)
    if not (idf and pwf):
        # 쿠키 초기화 후라면 정상 흐름에선 폼이 떠야 함. 그래도 폼이 없으면 한번 더 새로고침 후 재시도.
        print("[로그인] 폼 못 찾음 — 새로고침 후 재시도")
        try:
            page.reload(wait_until="domcontentloaded", timeout=20000)
            time.sleep(1.2)
        except Exception:
            pass
        idf, pwf, ctx = _find_login_form(page)
    if not (idf and pwf):
        print(f"[로그인] 폼 못 찾음 — 실패로 처리({user_id})")
        return False
    idf.fill(user_id)
    time.sleep(0.2)
    pwf.fill(password)
    time.sleep(0.2)
    c = ctx or page
    submitted = False
    for s in ['button[type="submit"]', 'input[type="submit"]', 'a:has-text("로그인")',
              'button:has-text("로그인")', '.btn_login', '#btn_login']:
        try:
            b = c.locator(s).first
            if b.count() > 0 and b.is_visible():
                with page.expect_navigation(wait_until=_WAIT, timeout=15000):
                    b.click()
                print(f"도매매 로그인 제출: {user_id}")
                submitted = True
                break
        except Exception:
            continue
    if not submitted:
        try:
            pwf.press("Enter")
            time.sleep(1.5)
            submitted = True
        except Exception:
            pass
    # 본인 계정 확인 (베스트에포트)
    if _verify_logged_in_as(page, user_id):
        print(f"[로그인] {user_id} 계정 확인됨")
        return True
    print(f"[로그인] [경고] 본인 계정({user_id}) 확인 실패 — 진행은 하되 결과 점검 필요")
    return True


def _open_supplier_stop(page) -> bool:
    """좌측 전송관리 > 공급사판매중지 진입."""
    if not _goto_with_retry(page, SPEEDGO_URL, _WAIT, 30000):
        return False
    time.sleep(1.5)
    for _try in range(2):
        # 1) '공급사판매중지' 링크 바로 시도
        for loc in (page.get_by_role("link", name="공급사판매중지").first,
                    page.locator('a:has-text("공급사판매중지")').first,
                    page.get_by_text("공급사판매중지", exact=True).first):
            try:
                if loc.count() > 0 and loc.is_visible():
                    loc.click(timeout=5000)
                    page.wait_for_load_state(_WAIT, timeout=20000)
                    time.sleep(2)
                    return True
            except Exception:
                continue
        # 2) 안 보이면 '전송관리' 펼치고 재시도
        try:
            tm = page.locator('a:has-text("전송관리"), :text("전송관리")').first
            if tm.count() > 0 and tm.is_visible():
                tm.click(timeout=4000)
                time.sleep(1.2)
        except Exception:
            pass
    print("[경고] 공급사판매중지 메뉴 진입 실패")
    return False


def _count_total(page) -> int:
    """'상품목록 (총 N건...)' 에서 N 추출. 못 찾으면 -1."""
    try:
        n = page.evaluate(r"""() => {
            const t = document.body ? (document.body.innerText||'') : '';
            const m = t.match(/총\s*([0-9,]+)\s*건/);
            return m ? parseInt(m[1].replace(/,/g,''),10) : -1;
        }""")
        return int(n)
    except Exception:
        return -1


def _set_500(page) -> bool:
    for sel in ('select:has(option:has-text("500"))',
                'select:has(option:has-text("개"))', 'select'):
        try:
            sl = page.locator(sel).first
            if sl.count() > 0 and sl.is_visible():
                for opt in ("500개씩 보기", "500개 보기", "500"):
                    try:
                        sl.select_option(label=opt)
                        time.sleep(2)
                        page.wait_for_load_state(_WAIT, timeout=15000)
                        print("500개씩 보기 적용")
                        return True
                    except Exception:
                        pass
                try:
                    sl.select_option(value="500")
                    time.sleep(2)
                    print("500개씩 보기 적용(value)")
                    return True
                except Exception:
                    pass
        except Exception:
            continue
    print("[경고] 500개 보기 변경 실패 (기본 개수로 진행)")
    return False


def _wait_list_loaded(page, timeout: int = 120) -> bool:
    """개수 변경(500) 후 목록이 실제로 다 렌더될 때까지 대기.
    sleep 의존 금지 — 행 수가 '연속 안정'될 때까지 폴링. 안정 후에만 전체선택."""
    try:
        page.wait_for_load_state(_WAIT, timeout=20000)
    except Exception:
        pass
    last, stable = -1, 0
    steps = max(1, int(timeout / 1.5))
    for i in range(steps):
        try:
            c = page.evaluate("""() => {
                const tb = document.querySelectorAll('table tbody tr');
                if (tb && tb.length) return tb.length;
                return document.querySelectorAll('table input[type="checkbox"]').length;
            }""")
        except Exception:
            c = -1
        if c and c > 0 and c == last:
            stable += 1
            if stable >= 3:           # 3회 연속 동일 = 로딩 완료로 판정
                print(f"목록 로딩 완료 (행 {c}개 안정)")
                time.sleep(1)         # 렌더 직후 미세 안정
                return True
        else:
            stable = 0
        last = c
        time.sleep(1.5)
    print(f"[경고] 목록 로딩 안정 대기 타임아웃 (마지막 행 {last})")
    return last and last > 0


def _select_all(page) -> bool:
    """목록 크기와 무관하게 모든 행 체크박스를 실제로 체크 + 검증.
    (작은 목록에서 헤더 selectAll 만으론 안 먹던 문제 → 행별 강제 체크로 해결)"""
    # 1) 헤더 selectAll 도 눌러줌(사이트 모델 연동용)
    for sel in ('input#selectAll', 'table thead input[type="checkbox"]',
                'thead input[type="checkbox"]'):
        try:
            cb = page.locator(sel).first
            if cb.count() > 0:
                cb.scroll_into_view_if_needed()
                if not cb.is_checked():
                    cb.click(force=True, timeout=4000)
                break
        except Exception:
            continue
    time.sleep(0.5)
    # 2) 모든 행 체크박스 직접 체크 + 이벤트 발생 후 '체크 수' 검증 (2회 재시도)
    for attempt in range(1, 4):
        try:
            res = page.evaluate("""() => {
                let boxes = [...document.querySelectorAll('table tbody tr input[type="checkbox"]')];
                if (boxes.length === 0)
                    boxes = [...document.querySelectorAll('table input[type="checkbox"]')]
                            .filter(b => !/selectAll/i.test(b.id || ''));
                let checked = 0;
                boxes.forEach(b => {
                    if (!b.checked) {
                        b.checked = true;
                        b.dispatchEvent(new Event('click', {bubbles:true}));
                        b.dispatchEvent(new Event('change', {bubbles:true}));
                    }
                    if (b.checked) checked++;
                });
                const sa = document.getElementById('selectAll');
                if (sa && !sa.checked) {
                    sa.checked = true;
                    sa.dispatchEvent(new Event('change', {bubbles:true}));
                }
                return {total: boxes.length, checked: checked};
            }""")
            total = int(res.get("total", 0) or 0)
            checked = int(res.get("checked", 0) or 0)
            if total > 0 and checked >= total:
                print(f"전체선택 완료 (체크 {checked}/{total})")
                return True
            print(f"전체선택 시도 {attempt}/3: {checked}/{total} — 재시도")
            time.sleep(1.2)
        except Exception as e:
            print(f"전체선택 시도 {attempt}/3 예외: {e}")
            time.sleep(1.0)
    # 마지막 상태로 판정
    try:
        res = page.evaluate("""() => {
            const bs=[...document.querySelectorAll('table tbody tr input[type=checkbox]')];
            return {t:bs.length, c:bs.filter(b=>b.checked).length}; }""")
        if (res.get("c", 0) or 0) > 0:
            print(f"전체선택 부분 완료 (체크 {res['c']}/{res['t']}) — 진행")
            return True
    except Exception:
        pass
    print("[경고] 전체선택 실패(체크 0) — 그래도 삭제 시도(품절삭제라 무방)")
    return False


def _click_delete_open_popup(page) -> bool:
    """목록 상단 '삭제' 클릭 → '상품삭제' 팝업."""
    for loc in (page.get_by_role("button", name="삭제", exact=True).first,
                page.locator('button:text-is("삭제")').first,
                page.locator('button:has-text("삭제"):not(:has-text("품절"))').first):
        try:
            if loc.count() > 0 and loc.is_visible():
                loc.click(timeout=5000)
                time.sleep(1.5)
                return True
        except Exception:
            continue
    print("[경고] 삭제 버튼 클릭 실패")
    return False


def _popup_all_markets_delete(page) -> bool:
    """팝업 '상품삭제': '활성화 된 마켓에서 모두 삭제' 선택 → 팝업 내 확정 '삭제' 클릭.
    page + 모든 frame 을 훑고, 안 잡히면 JS 로 보이는 모달의 정확히 '삭제' 버튼 클릭."""
    try:
        page.wait_for_selector('text=상품삭제', timeout=8000)
    except Exception:
        pass
    time.sleep(0.7)
    ctxs = [page]
    try:
        ctxs += list(page.frames)
    except Exception:
        pass


    # 1) 옵션: '활성화 된 마켓에서 모두 삭제' = a#delItem2 (실 DOM 확인된 정확 셀렉터)
    picked = False
    for ctx in ctxs:
        try:
            o = ctx.locator('#delItem2').first
            if o.count() > 0 and o.is_visible():
                o.click(timeout=4000)
                picked = True
                print("옵션 '활성화 된 마켓에서 모두 삭제'(#delItem2) 선택")
                break
        except Exception:
            continue
    if not picked:
        try:
            picked = bool(page.evaluate("""() => {
                const el=document.getElementById('delItem2');
                if(el){el.click();return true;}
                const a=[...document.querySelectorAll('a,button,div,span,li')]
                  .find(e=>/활성화\\s*된?\\s*마켓에서\\s*모두\\s*삭제/.test((e.innerText||e.textContent||'').trim()));
                if(a){a.click();return true;} return false; }"""))
        except Exception:
            pass
        print("옵션 선택(JS #delItem2)" if picked else "[경고] '모두 삭제' 옵션(#delItem2) 못 찾음")
    time.sleep(0.7)

    # [P3-A] 옵션 클릭만으로 삭제가 트리거되는 사이트 동작이 있음(현 코드 실측: 알림/원복으로 완료).
    # 그런 경우엔 .pup_del 시도 자체를 건너뛰어 '못 찾음' 노이즈 제거 + 진짜 실패와 구분.
    try:
        if not _popup_open(page):
            print("옵션 클릭 직후 팝업 종료 감지 → 확정 클릭 생략(삭제 트리거됨)")
            return True
    except Exception:
        pass

    # 2) 팝업이 여전히 열려있을 때만 확정 '삭제' 시도 = button.button2.pup_del
    for ctx in ctxs:
        for sel in ('button.pup_del', '.pup_del', 'button.button2.pup_del'):
            try:
                b = ctx.locator(sel).first
                if b.count() > 0 and b.is_visible():
                    b.click(timeout=5000)
                    print(f"팝업 확정 '삭제' 클릭 ({sel})")
                    return True
            except Exception:
                continue
    try:
        ok = page.evaluate("""() => {
            const b=document.querySelector('.pup_del')||document.querySelector('button.button2.pup_del');
            if(b){b.click();return true;} return false; }""")
        if ok:
            print("팝업 확정 '삭제' 클릭(JS .pup_del)")
            return True
    except Exception:
        pass
    print("[경고] 팝업 확정 '삭제'(.pup_del) 못 찾음 — 완료 대기 후 진행")
    return False


def _popup_open(page) -> bool:
    """'상품삭제' 팝업이 화면에 떠 있는가."""
    try:
        loc = page.locator('div.layer_pup_tit1:has-text("상품삭제")').first
        if loc.count() > 0 and loc.is_visible():
            return True
    except Exception:
        pass
    try:
        return bool(page.evaluate("""() => {
            const e=[...document.querySelectorAll('div.layer_pup_tit1')]
              .find(x=>/상품삭제/.test(x.innerText||x.textContent||''));
            if(!e) return false;
            const r=e.getBoundingClientRect();
            return r.width>0&&r.height>0&&getComputedStyle(e).visibility!=='hidden'; }"""))
    except Exception:
        return False


def _wait_delete_done(page, before_cnt, state, max_sec: int = 12 * 60) -> bool:
    """완료 규칙(사용자 확정):
       ① '삭제되었습니다' 알림 발생  OR
       ② 상품삭제 팝업이 닫혀 화면이 공급사판매중지 목록으로 원복  OR
       ③ 페이지/창 종료  OR  총건수 감소
       (잠금상품만 남아 즉시 원복되는 경우도 ②로 완료 간주). 완료팝업/처리중 스피너 안 봄."""
    poll = 4
    time.sleep(3)
    gone_hits = 0
    state["after_cnt"] = -1
    for elapsed in range(0, max_sec, poll):
        try:
            if page.is_closed():
                print("페이지/창 종료 → 삭제 완료 간주")
                state["result"] = "page_closed"
                return True
        except Exception:
            pass
        if state.get("done"):
            print("삭제 완료 — ‘삭제되었습니다’ 알림 감지")
            state["result"] = "alert"
            time.sleep(2)
            return True
        popup = _popup_open(page)
        cur = _count_total(page)
        state["after_cnt"] = cur
        if elapsed % 40 == 0:
            print(f"[삭제 대기] {elapsed}s · 총건수 {cur}(시작 {before_cnt}) 팝업열림={popup} 알림={state.get('done')}")
        if not popup:
            gone_hits += 1
            if gone_hits >= 2:
                if cur != -1 and before_cnt > 0 and cur >= before_cnt:
                    print(f"[판정] 화면 원복하나 총건수 무변({before_cnt}) — 잠금상품-only로 추정·완료 간주"
                          " (실제 미삭제면 해당 사업자 수동 확인 필요)")
                    state["result"] = "revert_noop"
                else:
                    print("화면 원복(상품삭제 팝업 닫힘) → 삭제 완료")
                    state["result"] = "revert_drop"
                return True
        else:
            gone_hits = 0
        if cur != -1 and before_cnt > 0 and cur < before_cnt:
            print(f"총건수 감소({before_cnt}→{cur}) → 삭제 완료")
            state["result"] = "count_drop"
            return True
        time.sleep(poll)
    print(f"[판정] ⚠ 미삭제 의심 — {max_sec}s 내 알림·화면원복·건수변화 전무 "
          "(확정 '삭제' 클릭 누락 가능). 이 사업자 수동 확인 필요. 다음 사업자로 진행")
    state["result"] = "timeout_suspect"
    return False


def _make_dialog_handler(state):
    """삭제 확정 후 네이티브 confirm/alert 를 '확인' 수락.
    '삭제되었습니다' 류 성공 알림이면 state['done']=True (완료 신호).
    (핸들러 없으면 Playwright 가 자동 dismiss=취소 → 삭제가 안 되던 원인)"""
    def _h(d):
        msg = (d.message or "")
        try:
            if any(k in msg for k in ("삭제되었", "삭제 되었", "완료", "되었습니다")):
                state["done"] = True
            print(f"[다이얼로그] {d.type}: {msg[:80]} → 수락 (done={state.get('done')})")
            d.accept()
        except Exception:
            try:
                d.dismiss()
            except Exception:
                pass
    return _h


def _is_logged_in(page) -> bool:
    try:
        for s in ('a:has-text("로그아웃")', 'text=로그아웃'):
            if page.locator(s).first.count() > 0:
                return True
    except Exception:
        pass
    return False


def main_delete_impl(context, items, password):
    """context 기반: 사업자별로 '새 탭'을 열어 처리.
    한 사업자 탭이 사이트에 의해 닫혀도(삭제 제출 시 흔함) Chrome/다음 사업자에 영향 없음.
    items = [(rank, user_id), ...]. 각 사업자 결과는 phase3_state.json 에 마커로 기록(대시보드용)."""
    for idx, (rank, user_id) in enumerate(items):
        print(f"\n{'='*54}\n[{rank}번] 공급사판매중지 삭제 | {user_id}\n{'='*54}")
        page = None
        deleted_attempted = False
        state = {"done": False, "after_cnt": -1}
        outcome = "error"  # 기본값: 어디서 죽었는지 모를 때
        before = -1
        try:
            page = context.new_page()
            try:
                page.on("dialog", _make_dialog_handler(state))
            except Exception:
                pass
            if not _login(page, user_id, password, do_logout=(idx > 0)):
                print(f"[{rank}번] 로그인 진입 실패 → 다음 사업자")
                outcome = "no_login"
                continue
            if not _is_logged_in(page):
                print(f"[{rank}번] [경고] 로그인 상태 미확인(로그아웃 링크 없음) — 그래도 진입 시도")
            else:
                print(f"[{rank}번] 로그인 상태 확인됨")
            if not _open_supplier_stop(page):
                print(f"[{rank}번] 공급사판매중지 진입 실패 → 다음 사업자")
                outcome = "no_open"
                continue
            before = _count_total(page)
            print(f"[{rank}번] 공급사판매중지 목록: 총 {before}건")
            if before == 0:
                print(f"[{rank}번] 삭제할 상품 없음 → 스킵")
                outcome = "no_target"
                continue
            _set_500(page)
            _wait_list_loaded(page)
            before = _count_total(page)
            if not _select_all(page):
                print(f"[{rank}번] 전체선택 일부 실패 — 그대로 진행")
            if not _click_delete_open_popup(page):
                print(f"[{rank}번] 삭제 버튼/팝업 미표시 → 이 사업자만 스킵")
                outcome = "no_popup"
                continue
            deleted_attempted = True
            _popup_all_markets_delete(page)
            _wait_delete_done(page, before, state)
            outcome = state.get("result", "unknown")
            print(f"[{rank}번] 1배치 삭제 종료 (result={outcome})")
            time.sleep(2)
        except Exception as e:
            msg = str(e)
            if deleted_attempted and "closed" in msg.lower():
                print(f"[{rank}번] 삭제 제출 후 페이지/창 종료 — 정상 처리로 간주, 다음 사업자")
                outcome = "page_closed"
            else:
                print(f"[{rank}번] 예외 → 다음 사업자: {msg[:160]}")
                outcome = "error"
        finally:
            try:
                _write_phase3_marker(rank, user_id, outcome, before, state.get("after_cnt", -1))
            except Exception as _me:
                print(f"[{rank}번] 마커 기록 실패(무시): {_me}")
            try:
                if page is not None and not page.is_closed():
                    page.close()
            except Exception:
                pass
    print("Phase 3: 모든 사업자 처리 완료.")
