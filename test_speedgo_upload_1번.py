# -*- coding: utf-8 -*-
r"""
스피드고 엑셀업로드 1~6번 사업자 순차 실행.
- 1~6번사업자 폴더에서 각각 *_최종.xlsx 찾아서
- 사업자별: 도매매 로그인 → 스피드고 접속 → 해시태그 검색 → 200개 보기 → 엑셀업로드 → 전송 → 상품전송 완료 → X

사용법:
  python test_speedgo_upload_1번.py
  python test_speedgo_upload_1번.py "C:\Users\USER\Documents\국내위탁\마이박스\26년3월3주차\6회차\1번사업자"

[스크립트로 띄운 Chrome이 불안정/구글 로그아웃/설정 초기화처럼 보일 때]
  스크립트가 새 Chrome을 "자동 제어용"으로 띄우기 때문에, 같은 Profile 67이라도
  일반에 켠 창과 다르게 동작할 수 있습니다 (구글 로그아웃, 확장/설정 제한 등).
  - 완전히 동일한 창을 쓰려면: Chrome을 직접 실행한 뒤 스크립트가 그 창에 "연결"하는
    방식을 쓰세요. (아래 USE_EXISTING_CHROME 사용법 참고)

  USE_EXISTING_CHROME=1 로 실행하면, 이미 켜 둔 Chrome에 연결합니다.
  - Chrome 띄우기: 프로젝트 폴더의 chrome_Profile67_원격디버깅.bat 더블클릭
  - 그 다음: set USE_EXISTING_CHROME=1 && python test_speedgo_upload_1번.py
  (수동 실행: "C:\...\chrome.exe" --remote-debugging-port=9222 --profile-directory="Profile 67")
"""
import re
import sys
import time
from pathlib import Path
from urllib.parse import quote

# cp949 등 콘솔에서 한글·치환문자 출력 시 print() 크래시 방지
for _sn in ("stdout", "stderr"):
    try:
        getattr(sys, _sn).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

PROJECT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_DIR))

EXCEL_SAVE_BASE = Path(r"C:\Users\USER\Documents\국내위탁\마이박스")
DOMEME_URL = "https://domemedb.domeggook.com/index/"
SPEEDGO_URL = "https://speedgo.domeggook.com/"
_WAIT = "domcontentloaded"
# 사용할 Chrome 서브 프로필 (User Data 아래 폴더명). None이면 Default 프로필 사용
CHROME_PROFILE_DIR = "Profile 67"


def _S(lo, hi):
    return lo


def build_upload_items(ymw_str: str, week_run: int, target_year: int, target_month: int, target_week: int):
    """domeme Phase1 저장 경로와 동일: EXCEL_SAVE_BASE/ymw_str/{week_run}회차/{n}번사업자"""
    from domeme_auto_login_temp import ACCOUNTS, build_speedgo_hashtag
    base_run = EXCEL_SAVE_BASE / ymw_str / f"{week_run}회차"
    items = []
    for n in range(1, 7):
        biz_folder = base_run / f"{n}번사업자"
        if not biz_folder.is_dir():
            continue
        finals = list(biz_folder.glob("*_최종.xlsx"))
        if not finals:
            continue
        final_path = max(finals, key=lambda p: p.stat().st_mtime)
        kw_tag, _, _, _ = parse_final_filename(final_path)
        if kw_tag is None:
            continue
        user_id = ACCOUNTS[n - 1] if n <= len(ACCOUNTS) else ""
        if not user_id:
            continue
        speedgo_hash = build_speedgo_hashtag(kw_tag, target_year, target_month, target_week)
        items.append((n, biz_folder, final_path, kw_tag, speedgo_hash, user_id))
    return items


def run_speedgo_upload_phase(context, page, items):
    """사업자별 스피드고 업로드 (도매매 로그인→스피드고→업로드→전송→상품일괄전송 X). domeme Phase2에서 호출."""
    from domeme_auto_login_temp import PASSWORD
    main_upload_impl(page, items, PASSWORD)


def parse_final_filename(path: Path):
    """파일명 {user_id}_{kw_tag}_{26년N월N주차}_{N}회_최종.xlsx 에서 해시태그용 kw_tag, year, month, week 추출."""
    name = path.stem  # 예: r1_니트_26년3월3주차_6회_최종
    if not name.endswith("_최종"):
        return None, None, None, None
    parts = name.replace("_최종", "").split("_")  # ['r1', '니트', '26년3월3주차', '6회']
    ymw = None
    for p in parts:
        if "년" in p and "월" in p and "주차" in p:
            ymw = p
            break
    if not ymw:
        return None, None, None, None
    m = re.search(r"(\d{2})년(\d{1,2})월(\d+)주차", ymw)
    if not m:
        return None, None, None, None
    y, mo, w = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        idx = parts.index(ymw)
        kw_tag = parts[idx - 1] if idx > 0 else (parts[0] if parts else "키워드")
    except Exception:
        kw_tag = parts[0] if parts else "키워드"
    return kw_tag, y, mo, w


def _goto_with_retry(target, url, wait_until, timeout, tries=4, settle=2.0):
    """transient 'interrupted by another navigation'/timeout 시 재시도.
    성공 시 True, 끝까지 실패 시 False (예외를 상위로 전파하지 않음 → 호출부에서 사업자 격리)."""
    last = None
    for i in range(1, tries + 1):
        try:
            target.goto(url, wait_until=wait_until, timeout=timeout)
            return True
        except Exception as e:
            last = e
            msg = str(e)
            transient = (
                "interrupted by another navigation" in msg
                or "Navigation" in msg
                or "Timeout" in msg
                or "net::ERR" in msg
            )
            print(f"[goto 재시도 {i}/{tries}] {url} : {msg[:150]}")
            if not transient or i == tries:
                break
            time.sleep(settle)
    print(f"[goto 실패] {url} ({tries}회 시도): {str(last)[:150]}")
    return False


def main_upload_impl(page, items, password):
    """업로드 루프 (도매매 로그인→스피드고→엑셀업로드→전체선택→전송→상품일괄전송 X)"""
    def make_mb_save_list_url(speedgo_hash):
        return (
            "https://speedgo.domeggook.com/mybox/mb_saveList.php?"
            f"pagenum=&hashTag={quote(speedgo_hash, safe='')}&sf=subject&sw=&itemNos=&mnp=&mxp="
            "&titleStatus=&editStatus=&useOption=&sender_date1=&sender_date2="
            "&sort1=&sort2=&sort3=&sort4=&sort5=&b2bStatus=0&pageLimit=200"
        )
    id_selectors = [
        'input[name="userId"]', 'input[name="user_id"]', 'input[name="id"]', 'input[name="loginId"]',
        'input[id*="user"]', 'input[id*="id"]', 'input[type="text"]', 'input#userId', 'input#user_id',
    ]
    pw_selectors = ['input[name="password"]', 'input[name="passwd"]', 'input[name="pw"]', 'input[type="password"]']

    def find_login_form(pg):
        contexts = [pg]
        try:
            contexts.extend(pg.frames[1:])
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

    def _accept_upload_confirm(dialog):
        try:
            dialog.accept()
            print("업로드 확인 대화상자: 확인 클릭")
        except Exception:
            pass
    page.on("dialog", _accept_upload_confirm)

    for idx, (n, biz_folder, final_path, kw_tag, speedgo_hash, user_id) in enumerate(items):
        mb_save_list_url = make_mb_save_list_url(speedgo_hash)
        print(f"\n{'='*50}\n{n}번 사업자 업로드: {final_path.name} | {user_id}\n{'='*50}")
        # [수정] 사업자 전환 시 컨텍스트 쿠키 명시 초기화 → 이전 사업자 세션 잔류로
        # "로그인 폼 못 찾음(이미 로그인됨)" 경로로 빠져 잘못된 계정으로 마이박스 업로드하던 위험 차단.
        # (click-logout 은 비결정적이라 폐기)
        try:
            page.context.clear_cookies()
            print(f"[로그인] 컨텍스트 쿠키 초기화 → 새 사업자 fresh 로그인 ({user_id})")
        except Exception as _ce:
            print(f"[로그인] 쿠키 초기화 실패(무시): {_ce}")
        print("도매매 URL 로 이동 중...")
        if not _goto_with_retry(page, DOMEME_URL, "commit", 45000):
            print(f"[{n}번] 도매매 진입 실패 → 이 사업자 건너뛰고 다음으로")
            continue
        try:
            page.wait_for_load_state("domcontentloaded", timeout=20000)
        except Exception:
            pass
        time.sleep(1)
        try:
            login_link = page.get_by_role("link", name="로그인").first
            if login_link.count() == 0:
                login_link = page.locator('a:has-text("로그인"):not(:has-text("진행중"))').first
            if login_link.count() == 0:
                login_link = page.get_by_text("로그인", exact=True).first
            if login_link.count() > 0 and login_link.is_visible():
                login_link.click(timeout=5000)
                print("도매매 로그인 링크 클릭")
        except Exception as e:
            print(f"로그인 링크 클릭: {e}")
        time.sleep(1.2)
        id_field, pw_field, form_ctx = find_login_form(page)
        if not (id_field and pw_field):
            # 쿠키 초기화 후라면 정상적으로 폼이 떠야 함. 한 번 새로고침 후 재탐색.
            print(f"[{n}번] 로그인 폼 못 찾음 — 새로고침 후 재시도")
            try:
                page.reload(wait_until="domcontentloaded", timeout=20000)
                time.sleep(1.2)
            except Exception:
                pass
            id_field, pw_field, form_ctx = find_login_form(page)
        if id_field and pw_field:
            id_field.fill(user_id)
            time.sleep(0.2)
            pw_field.fill(password)
            time.sleep(0.2)
            ctx = form_ctx or page
            for sel in ['button[type="submit"]', 'input[type="submit"]', 'a:has-text("로그인")', 'button:has-text("로그인")', '.btn_login', '#btn_login']:
                try:
                    btn = ctx.locator(sel).first
                    if btn.count() > 0 and btn.is_visible():
                        with page.expect_navigation(wait_until=_WAIT, timeout=15000):
                            btn.click()
                        print(f"도매매 로그인 제출: {user_id}")
                        break
                except Exception:
                    continue
        else:
            print(f"[{n}번] 로그인 폼 없음 — 이 사업자 건너뜀(잘못된 계정으로 진행 차단)")
            continue
        # 본인 계정 확인 (page text 에 user_id 포함 여부)
        _ok_who = False
        for _ in range(8):
            try:
                _t = page.evaluate("() => (document.body?document.body.innerText:'')+' '+(document.title||'')")
                if user_id and user_id in (_t or ""):
                    _ok_who = True
                    break
            except Exception:
                pass
            time.sleep(0.4)
        if _ok_who:
            print(f"[{n}번] 로그인 확인됨: {user_id}")
        else:
            print(f"[{n}번] [경고] 본인 계정({user_id}) 확인 실패 — 진행은 하되 결과 점검 필요")
        time.sleep(1)
        try:
            page.wait_for_load_state(_WAIT, timeout=15000)
        except Exception:
            pass
        if not _goto_with_retry(page, SPEEDGO_URL, _WAIT, 30000):
            print(f"[{n}번] 스피드고 진입 실패 → 이 사업자 건너뛰고 다음으로")
            continue
        time.sleep(1)
        try:
            mybox = page.get_by_role("link", name="마이박스").first
            if mybox.count() > 0 and mybox.is_visible():
                mybox.click()
                time.sleep(1)
        except Exception:
            pass
        try:
            inp = page.locator('input[placeholder*="해시태그"], input[placeholder*="#태그명"]').first
            inp.wait_for(state="visible", timeout=10000)
            inp.fill(speedgo_hash)
            print("해시태그 입력 완료")
        except Exception as e:
            print(f"해시태그 입력 실패: {e}")
        time.sleep(0.4)
        try:
            inp = page.locator('input[placeholder*="해시태그"], input[placeholder*="#태그명"]').first
            inp.press("Enter")
            print("해시태그 입력 후 Enter로 검색 제출")
            page.wait_for_load_state(_WAIT, timeout=15000)
            time.sleep(2)
        except Exception as e1:
            try:
                page.evaluate("""() => {
                    const list = document.querySelectorAll('button[type="submit"], input[type="submit"], button');
                    for (const b of list) {
                        const t = (b.value || b.innerText || b.textContent || '').trim();
                        if (t === '검색') { b.click(); return true; }
                    }
                    return false;
                }""")
                print("검색 버튼 클릭 (JS, 텍스트=검색만)")
                page.wait_for_load_state(_WAIT, timeout=15000)
                time.sleep(2)
            except Exception:
                print(f"검색 제출 실패: {e1}")
        try:
            page.wait_for_selector('text=/총\\s*\\d+건/', timeout=15000)
            print("검색 결과 목록 로드됨")
        except Exception:
            pass
        time.sleep(1)
        done_200 = False
        try:
            page_limit_dropdown = page.locator('select:has(option[value="200"])').first
            if page_limit_dropdown.count() > 0 and page_limit_dropdown.is_visible():
                page_limit_dropdown.select_option(value="200")
                time.sleep(1.5)
                page.wait_for_load_state(_WAIT, timeout=10000)
                print("200개씩 보기 선택")
                done_200 = True
        except Exception:
            pass
        if not done_200:
            try:
                page.goto(mb_save_list_url, wait_until=_WAIT, timeout=30000)
                time.sleep(1.5)
                print("200개씩 보기 (URL 이동)")
                done_200 = True
            except Exception:
                pass
        time.sleep(1)
        try:
            page.wait_for_selector('text=/총\\s*\\d+건/', timeout=8000)
        except Exception:
            pass
        time.sleep(0.5)
        try:
            btn = page.locator("#mbUploadBtn").first
            if btn.count() == 0:
                btn = page.locator('button[onclick*="excelMbUpload"]').first
            if btn.count() > 0 and btn.is_visible():
                btn.click()
                print("1차 엑셀업로드 클릭 (#mbUploadBtn)")
            else:
                raise RuntimeError("엑셀업로드 버튼 없음")
        except Exception as e:
            print(f"1차 엑셀업로드 실패: {e}")
            continue
        time.sleep(1.2)
        try:
            with page.expect_file_chooser(timeout=12000) as fc:
                file_sel = page.get_by_text("파일 선택").first
                if file_sel.count() == 0:
                    file_sel = page.get_by_role("button", name="파일 선택").first
                if file_sel.count() > 0 and file_sel.is_visible():
                    file_sel.click()
                else:
                    raise RuntimeError("파일 선택 버튼 없음")
            fc.value.set_files(str(final_path))
            print(f"2차 파일 선택 완료: {final_path.name}")
        except Exception as e:
            try:
                page.locator('input[type="file"]').first.set_input_files(str(final_path))
                print(f"파일 설정 (input 직접): {final_path.name}")
            except Exception:
                print(f"파일 선택 실패: {e}")
                continue
        time.sleep(1.5)
        for btn_text in ("업로드", "적용", "확인"):
            try:
                b = page.get_by_role("button", name=btn_text).first
                if b.count() == 0:
                    b = page.locator(f'button:has-text("{btn_text}")').first
                if b.count() == 0:
                    b = page.locator(f'input[type="submit"][value*="{btn_text}"]').first
                if b.count() == 0:
                    b = page.locator(f'a:has-text("{btn_text}")').first
                if b.count() > 0 and b.is_visible():
                    b.click()
                    print(f"업로드 실행 버튼 클릭: {btn_text}")
                    break
            except Exception:
                continue
        else:
            print("업로드 실행 버튼을 찾지 못함. 수동으로 업로드 버튼을 눌러주세요.")
        time.sleep(1)
        try:
            confirm_btn = page.get_by_role("button", name="확인").first
            if confirm_btn.count() == 0:
                confirm_btn = page.locator('button:has-text("확인")').first
            if confirm_btn.count() > 0:
                confirm_btn.wait_for(state="visible", timeout=5000)
                confirm_btn.click(timeout=3000)
                print("업로드 확인 대화상자: 확인 클릭")
        except Exception:
            pass
        time.sleep(2)
        # 전체선택 ~ 스피드고전송 ~ 준비창 ~ 상품일괄전송 X (test_speedgo 기존 로직)
        _do_select_and_send_flow(page)
        # 진척도 표시용 마커(전송 동작/재전송엔 영향 없음. 제어판 진척도에서 사용)
        try:
            from pathlib import Path as _P
            _P(biz_folder, ".phase2_sent").write_text(
                f"{time.strftime('%Y-%m-%d %H:%M:%S')}\t{final_path.name}", encoding="utf-8")
        except Exception:
            pass
        # [P2-A] 상품일괄전송 결과 그리드 원본 텍스트 덤프 (마켓별 전송성공/실패+사유 파악용)
        try:
            from pathlib import Path as _P2
            _LD = _P2(PROJECT_DIR) / "logs"
            _LD.mkdir(exist_ok=True)
            _txt = page.evaluate(
                "() => document.body ? (document.body.innerText||'') : ''"
            ) or ""
            _ts = time.strftime("%Y%m%d_%H%M%S")
            (_LD / f"phase2_results_{user_id}_{_ts}.txt").write_text(
                f"=== {final_path.name} ({_ts}) ===\n{_txt[:200000]}", encoding="utf-8")
            # 빠른 요약: 전송실패/전송성공 카운트
            try:
                _ok = _txt.count("전송성공")
                _ng = _txt.count("전송실패")
                _na = _txt.count("전송안함")
                print(f"[P2 결과 요약] {user_id}: 전송성공={_ok} 전송실패={_ng} 전송안함={_na}",
                      flush=True)
            except Exception:
                pass
        except Exception as _de:
            print(f"[P2 결과 덤프] 실패(무시): {_de}", flush=True)
        time.sleep(2)
    print("모든 사업자 업로드 완료.")


def _do_select_and_send_flow(page):
    """전체선택 → 스피드고전송 → 준비창(iframe) 최종 전송 → 상품일괄전송 X"""
    time.sleep(1.5)
    try:
        page.wait_for_selector('text=상품목록', state="visible", timeout=10000)
    except Exception:
        pass
    try:
        page.wait_for_selector('input#selectAll, table thead input[type="checkbox"]', state="visible", timeout=8000)
    except Exception:
        pass
    time.sleep(0.5)
    all_checked = False
    try:
        cb = page.locator('input#selectAll').first
        if cb.count() > 0:
            cb.scroll_into_view_if_needed()
            time.sleep(0.2)
            if not cb.is_checked():
                cb.click(force=True)
                all_checked = True
                print("전체선택 체크 완료 (#selectAll)")
    except Exception:
        pass
    if not all_checked:
        try:
            done = page.evaluate("""() => {
                const el = document.getElementById('selectAll');
                if (el && !el.checked) { el.checked = true; el.dispatchEvent(new Event('change', { bubbles: true })); el.click(); return true; }
                return false;
            }""")
            if done:
                all_checked = True
                print("전체선택 체크 완료 (#selectAll JS)")
        except Exception:
            pass
    if not all_checked:
        for sel in ('table thead input[type="checkbox"]', 'thead th:first-child input[type="checkbox"]', 'table thead th input[type="checkbox"]', 'table input[type="checkbox"]'):
            try:
                cb = page.locator(sel).first
                if cb.count() > 0:
                    cb.scroll_into_view_if_needed()
                    time.sleep(0.2)
                    if not cb.is_checked():
                        cb.click(force=True)
                        all_checked = True
                        print("전체선택 체크 완료")
                        break
            except Exception:
                continue
    if not all_checked:
        try:
            clicked = page.evaluate("""() => {
                const thead = document.querySelector('table thead');
                if (thead) { const cb = thead.querySelector('input[type="checkbox"]');
                if (cb && !cb.checked) { cb.click(); return true; } }
                const first = document.querySelector('table input[type="checkbox"]');
                if (first && !first.checked) { first.click(); return true; }
                return false;
            }""")
            if clicked:
                all_checked = True
                print("전체선택 체크 완료 (JS thead)")
        except Exception:
            pass
    if not all_checked:
        try:
            el = page.evaluate("""() => {
                const el = document.getElementById('selectAll');
                if (!el) return false;
                if (el.checked) return true;
                el.checked = true;
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
                return true;
            }""")
            if el:
                all_checked = True
                print("전체선택 체크 완료 (#selectAll JS)")
        except Exception:
            pass
    if not all_checked:
        print("전체선택 실패 - 수동으로 체크해 주세요.")
    time.sleep(0.5)
    try:
        speedgo_clicked = False
        for sel in ('button[onclick*="speedGoSend"]', 'button.button2:has-text("스피드고전송")', 'a:has-text("스피드고전송")'):
            btn = page.locator(sel).first
            if btn.count() > 0 and btn.is_visible():
                btn.scroll_into_view_if_needed()
                time.sleep(0.2)
                btn.click(force=True)
                print("스피드고전송 버튼 클릭 (준비창)")
                speedgo_clicked = True
                break
        if not speedgo_clicked:
            clicked = page.evaluate("""() => {
                const btns = document.querySelectorAll('button, a');
                for (const b of btns) {
                    const onclick = b.getAttribute('onclick') || (b.onclick ? b.onclick.toString() : '');
                    const text = (b.innerText || b.textContent || '').trim();
                    if (onclick.indexOf('speedGoSend') >= 0 || (text === '스피드고전송' && b.closest && b.closest('.fl'))) {
                        b.click(); return true;
                    }
                }
                return false;
            }""")
            if clicked:
                print("스피드고전송 버튼 클릭 (JS)")
    except Exception as e:
        print(f"스피드고전송 버튼: {e}")
    time.sleep(1.5)
    try:
        page.wait_for_selector('.layui-layer', state="visible", timeout=12000)
    except Exception:
        pass
    time.sleep(0.5)
    try:
        iframe_layer = page.locator('div.layui-layer.layui-layer-iframe')
        iframe_count_main = iframe_layer.count()
        use_frame = iframe_count_main > 0
        frame = None
        frame_loc = None
        if use_frame:
            try:
                page.wait_for_selector('div.layui-layer.layui-layer-iframe iframe', state="attached", timeout=5000)
            except Exception:
                pass
            iframe_handle = page.query_selector('div.layui-layer.layui-layer-iframe iframe')
            if iframe_handle:
                frame = iframe_handle.content_frame()
            frame_loc = page.frame_locator('div.layui-layer.layui-layer-iframe >> iframe')
            if frame is None:
                for f in page.frames():
                    if f != page.main_frame and f.url:
                        frame = f
                        break
        target_page = frame if (use_frame and frame) else page
        target_locator = frame_loc if (use_frame and frame_loc) else page
        can_use_frame = use_frame and frame is not None
        if can_use_frame or not use_frame:
            def scroll_inside():
                if target_page == page:
                    return page.evaluate("""() => {
                        const all = document.body.querySelectorAll('*');
                        const scrollables = [];
                        for (const el of all) {
                            if (el.scrollHeight <= el.clientHeight) continue;
                            const oy = getComputedStyle(el).overflowY || getComputedStyle(el).overflow;
                            if (oy !== 'auto' && oy !== 'scroll' && oy !== 'overlay') continue;
                            scrollables.push(el);
                        }
                        let done = true;
                        for (const el of scrollables) {
                            el.scrollTop = el.scrollHeight;
                            if (el.scrollTop + el.clientHeight < el.scrollHeight - 2) done = false;
                        }
                        return { done, count: scrollables.length };
                    }""")
                return target_page.evaluate("""() => {
                    const all = document.body.querySelectorAll('*');
                    const scrollables = [];
                    for (const el of all) {
                        if (el.scrollHeight <= el.clientHeight) continue;
                        const oy = getComputedStyle(el).overflowY || getComputedStyle(el).overflow;
                        if (oy !== 'auto' && oy !== 'scroll' && oy !== 'overlay') continue;
                        scrollables.push(el);
                    }
                    let done = true;
                    for (const el of scrollables) {
                        el.scrollTop = el.scrollHeight;
                        if (el.scrollTop + el.clientHeight < el.scrollHeight - 2) done = false;
                    }
                    return { done, count: scrollables.length };
                }""")
            for _ in range(20):
                r = scroll_inside()
                if r.get("done") or r.get("count", 0) == 0:
                    break
                time.sleep(0.3)
        time.sleep(0.3)
        BUTTON_SELECTORS = ['button.cont_btn1[onclick*="goProduct"]', 'button.cont_btn1', 'button[onclick*="goProduct"]']
        if can_use_frame or not use_frame:
            # 대기 타임아웃은 비치명적으로 변경(20s) — 이전엔 8s timeout 이 try 밖이라
            # 한 번 늦으면 Phase 2 전체가 rc=1 로 죽었음. 이제 못 찾아도 다음 단계로 진행.
            try:
                if target_page == page:
                    page.wait_for_selector('button.cont_btn1[onclick*="goProduct"]', state="attached", timeout=20000)
                else:
                    target_page.wait_for_selector('button.cont_btn1[onclick*="goProduct"]', state="attached", timeout=20000)
            except Exception as _we:
                print(f"[전송버튼 대기] attached 타임아웃(무시): {str(_we)[:120]}")
            try:
                if target_page == page:
                    page.wait_for_function("() => { const b = document.querySelector('button.cont_btn1[onclick*=\"goProduct\"]'); return b && !b.disabled && getComputedStyle(b).pointerEvents !== 'none' && b.getBoundingClientRect().height > 0; }", timeout=10000)
                else:
                    target_page.wait_for_function("() => { const b = document.querySelector('button.cont_btn1[onclick*=\"goProduct\"]') || Array.from(document.querySelectorAll('button.cont_btn1')).find(x => (x.innerText||'').trim().includes('스피드고전송')); return b && !b.disabled && getComputedStyle(b).pointerEvents !== 'none' && b.getBoundingClientRect().height > 0; }", timeout=10000)
            except Exception:
                pass
        elif use_frame and frame_loc:
            try:
                frame_loc.locator('button.cont_btn1[onclick*="goProduct"]').first.wait_for(state="visible", timeout=10000)
            except Exception:
                pass
        btn_ok = False
        if can_use_frame or not use_frame:
            if target_page == page:
                btn_ok = page.evaluate("""() => {
                    const btn = document.querySelector('button.cont_btn1[onclick*="goProduct"]') || Array.from(document.querySelectorAll('button.cont_btn1')).find(b => (b.innerText||'').trim().includes('스피드고전송'));
                    if (!btn) return false;
                    const s = getComputedStyle(btn);
                    if (s.display==='none'||s.visibility==='hidden'||s.pointerEvents==='none') return false;
                    if (btn.disabled) return false;
                    const r = btn.getBoundingClientRect();
                    return r.height > 0 && r.width > 0;
                }""")
            else:
                btn_ok = target_page.evaluate("""() => {
                    const btn = document.querySelector('button.cont_btn1[onclick*="goProduct"]') || Array.from(document.querySelectorAll('button.cont_btn1')).find(b => (b.innerText||'').trim().includes('스피드고전송'));
                    if (!btn) return false;
                    const s = getComputedStyle(btn);
                    if (s.display==='none'||s.visibility==='hidden'||s.pointerEvents==='none') return false;
                    if (btn.disabled) return false;
                    const r = btn.getBoundingClientRect();
                    return r.height > 0 && r.width > 0;
                }""")
        else:
            btn_ok = True
        if not btn_ok:
            print("[디버그] 전송 버튼이 visible/enabled/clickable 아님. 클릭 생략.")
        final_clicked = False
        if btn_ok:
            for sel in BUTTON_SELECTORS:
                try:
                    loc = target_locator.locator(sel).first
                    if loc.count() > 0:
                        loc.scroll_into_view_if_needed(timeout=5000)
                        time.sleep(0.15)
                        loc.click(force=True, timeout=3000)
                        final_clicked = True
                        print("최종 스피드고전송 버튼 클릭:", sel)
                        break
                except Exception:
                    continue
            if not final_clicked and target_page != page:
                clicked = target_page.evaluate("""() => {
                    const btn = document.querySelector('button.cont_btn1[onclick*="goProduct"]') || Array.from(document.querySelectorAll('button.cont_btn1')).find(b => (b.innerText||'').trim().includes('스피드고전송'));
                    if (!btn || btn.disabled) return false;
                    btn.scrollIntoView({ block: 'end', behavior: 'instant' });
                    btn.click();
                    return true;
                }""")
                if clicked:
                    final_clicked = True
                    print("최종 스피드고전송 버튼 클릭 (JS, iframe 내부)")
            if not final_clicked and target_page == page:
                clicked = page.evaluate("() => { const b = document.querySelector('button.cont_btn1[onclick*=\"goProduct\"]'); if (!b || b.disabled) return false; b.scrollIntoView({ block: 'end' }); b.click(); return true; }")
                if clicked:
                    final_clicked = True
                    print("최종 스피드고전송 버튼 클릭 (JS)")
        if not final_clicked:
            print("[디버그] 최종 스피드고전송 버튼 클릭 실패")
    except Exception as e:
        print(f"최종 스피드고전송: {e}")
    try:
        # 상품전송 완료 메시지 감지 (페이지/iframe/텍스트 변형 대응, 폴링으로 감지)
        _max_wait_sec = 45 * 60  # 최대 45분
        _poll_interval = 8       # 8초마다 확인
        _complete_patterns = [
            "상품전송이 완료되었습니다",
            "상품 전송이 완료되었습니다",
            "상품전송 완료되었습니다",
            "상품 전송 완료되었습니다",
            "상품전송이 완료되었습니다.",
        ]
        found = False
        for elapsed in range(0, _max_wait_sec, _poll_interval):
            for ctx in [page] + list(page.frames):
                try:
                    has_text = ctx.evaluate("""(patterns) => {
                        const html = document.body ? document.body.innerText || document.body.textContent || '' : '';
                        return patterns.some(p => html.includes(p));
                    }""", _complete_patterns)
                    if has_text:
                        found = True
                        break
                except Exception:
                    continue
            if found:
                print(f"상품전송 완료 메시지 감지됨 (대기 {elapsed}초)")
                break
            if elapsed > 0 and elapsed % 60 == 0:
                print(f"[상품전송 대기] {elapsed // 60}분 경과...")
            time.sleep(_poll_interval)
        if not found:
            print("[상품전송] 완료 메시지 미감지. 수동 확인 후 창을 닫아 주세요.")
        else:
            print("상품전송 완료 메시지 확인. 10초 대기 후 다음 사업자로 진행.")
            time.sleep(10)
        closed = False
        for close_sel in ['.layui-layer-setwin .layui-layer-close', '.layui-layer-close', 'a.layui-layer-close']:
            try:
                btn_close = page.locator(close_sel).first
                if btn_close.count() > 0:
                    btn_close.click(timeout=3000)
                    closed = True
                    print("상품일괄전송 창 닫기 (X) 완료.")
                    break
            except Exception:
                continue
        if not closed:
            print("상품일괄전송 창 X 버튼 클릭 실패. 수동으로 닫아주세요.")
    except Exception as e:
        print(f"상품전송 완료 대기 또는 창 닫기: {e}")


def main():
    from domeme_auto_login_temp import build_speedgo_hashtag, get_target_ymw, get_upload_path_from_state, ACCOUNTS, PASSWORD
    target_year, target_month, target_week = get_target_ymw()

    # 주차/회차 기준 경로 (인자로 1번사업자 폴더 주면 그 부모에서 1~6번 전체, 없으면 .week_run_state 기준으로 Phase 1과 동일 경로)
    if len(sys.argv) >= 2:
        first_biz = Path(sys.argv[1])
        base_run = first_biz.parent  # .../1회차
        base_ymw = base_run.parent   # .../26년3월3주차
    else:
        ymw_str, week_run = get_upload_path_from_state()
        base_ymw = EXCEL_SAVE_BASE / ymw_str
        base_run = base_ymw / f"{week_run}회차"
        print(f"[업로드 경로] {ymw_str} {week_run}회차 (Phase 1과 동일)")

    # 1~6번 사업자별 (폴더, 파일, kw_tag, user_id) 목록 생성
    items = []
    for n in range(1, 7):
        biz_folder = base_run / f"{n}번사업자"
        if not biz_folder.is_dir():
            continue
        finals = list(biz_folder.glob("*_최종.xlsx"))
        if not finals:
            continue
        final_path = max(finals, key=lambda p: p.stat().st_mtime)
        kw_tag, _, _, _ = parse_final_filename(final_path)
        if kw_tag is None:
            print(f"[건너뜀] {n}번사업자 파일명 파싱 실패: {final_path.name}")
            continue
        user_id = ACCOUNTS[n - 1] if n <= len(ACCOUNTS) else ""
        if not user_id:
            print(f"[건너뜀] {n}번사업자 계정 없음")
            continue
        speedgo_hash = build_speedgo_hashtag(kw_tag, target_year, target_month, target_week)
        items.append((n, biz_folder, final_path, kw_tag, speedgo_hash, user_id))

    if not items:
        print("[오류] 처리할 사업자 폴더(_최종.xlsx + 계정)가 없습니다.")
        sys.exit(1)
    print(f"순차 업로드 대상: {[x[0] for x in items]}번 사업자")

    # 엑셀 업로드 시 파일 선택 등이 정상 동작하도록 실제 크롬 프로필 사용 (복사본 X)
    import os
    from domeme_auto_login_temp import (
        REAL_CHROME_USER_DATA,
        REAL_PROFILE_COPY_DIR,
        CHROME_EXECUTABLE,
        _copy_real_chrome_profile,
    )
    try:
        from playwright.sync_api import sync_playwright
        from playwright._impl._errors import TargetClosedError
    except Exception:
        from playwright.sync_api import sync_playwright
        TargetClosedError = Exception  # fallback

    user_data_dir = REAL_CHROME_USER_DATA
    if not user_data_dir.exists():
        print(f"[오류] 실제 Chrome 프로필 경로 없음: {user_data_dir}")
        sys.exit(1)
    if not os.environ.get("USE_EXISTING_CHROME", "").strip() == "1":
        print("[안내] 스크립트가 새 Chrome을 띄우면 구글 로그아웃/설정이 달라 보일 수 있습니다.")
        print("       그대로 쓰려면 USE_EXISTING_CHROME=1 로 '이미 켠 Chrome에 연결'하세요. (상단 docstring 참고)")
    profile_dir_name = os.environ.get("CHROME_PROFILE_DIR", CHROME_PROFILE_DIR)
    # 자동 제어로 인한 제한 완화 (사이트가 '자동화'로 보는 것 일부 완화)
    launch_args = ["--start-maximized", "--disable-blink-features=AutomationControlled"]
    if profile_dir_name:
        launch_args.append(f"--profile-directory={profile_dir_name}")
        if (user_data_dir / profile_dir_name).exists():
            print(f"[실행 설정] Chrome 프로필: {user_data_dir}\\{profile_dir_name}")
        else:
            print(f"[실행 설정] Chrome 프로필: {profile_dir_name} (폴더 없으면 Default로 열릴 수 있음)")
    print("[실행 설정] Chrome 창을 모두 닫은 뒤 실행하세요.")

    IGNORE_ARGS = [
        "--incognito", "--guest", "--off-the-record", "--bwsi", "--inprivate",
        "--enable-automation", "--no-sandbox",
        "--disable-extensions", "--disable-sync", "--disable-default-apps",
        "--disable-component-extensions-with-background-pages",
        "--disable-background-networking", "--disable-background-timer-throttling",
        "--disable-backgrounding-occluded-windows", "--disable-renderer-backgrounding",
        "--disable-client-side-phishing-detection",
    ]
    launch_kw = {
        "user_data_dir": str(user_data_dir),
        "headless": False,
        "channel": "chrome",
        "ignore_default_args": IGNORE_ARGS,
        "args": launch_args,
        "locale": "ko-KR",
    }
    if os.path.isfile(CHROME_EXECUTABLE):
        launch_kw["executable_path"] = CHROME_EXECUTABLE
        launch_kw.pop("channel", None)
    if not PASSWORD:
        print("[오류] domeme_auto_login_temp.py 에 PASSWORD 가 설정되어 있어야 합니다.")
        sys.exit(1)

    def _launch(context_creator):
        try:
            return context_creator(no_viewport=True)
        except TypeError:
            return context_creator()

    use_existing_chrome = os.environ.get("USE_EXISTING_CHROME", "").strip() == "1"

    with sync_playwright() as p:
        context = None
        if use_existing_chrome:
            port = os.environ.get("REMOTE_DEBUGGING_PORT", "9222").strip()
            try:
                browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
                if browser.contexts:
                    context = browser.contexts[0]
                    print("[연결] 이미 켜 둔 Chrome에 연결했습니다. (구글/설정 그대로 사용)")
                else:
                    print("[연결] 컨텍스트 없음. Chrome을 --remote-debugging-port=9222 로 켜 주세요.")
            except Exception as e:
                print(f"[연결 실패] 포트 {port}에 Chrome이 없습니다: {e}")
                print("  Chrome을 한 번만 이렇게 실행하세요: chrome.exe --remote-debugging-port=9222")
        if context is None and not use_existing_chrome:
            try:
                ctx_kw = {**launch_kw}
                context = _launch(lambda no_viewport=False: p.chromium.launch_persistent_context(**ctx_kw, **({"no_viewport": True} if no_viewport else {})))
            except TypeError:
                try:
                    context = p.chromium.launch_persistent_context(**launch_kw)
                except (TargetClosedError, Exception):
                    context = None
            except (TargetClosedError, Exception) as e:
                if type(e).__name__ == "TargetClosedError" or "closed" in str(e).lower():
                    print("\n[안내] Chrome이 이미 실행 중입니다. 기존 창에 연결합니다 (새 창 안 띔).\n")
                context = None
        if context is None and not use_existing_chrome:
            # 1) 이미 켜 둔 Chrome(원격디버깅 9222)에 연결 시도
            port = os.environ.get("REMOTE_DEBUGGING_PORT", "9222").strip()
            try:
                browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
                if browser.contexts:
                    context = browser.contexts[0]
                    use_existing_chrome = True
                    print(f"[연결] 기존 Chrome에 연결했습니다 (포트 {port}).")
            except Exception:
                pass
            # 2) 연결 실패 시 프로필 복사본으로 새 창 띄워서 진행 (막히지 않도록)
            if context is None:
                print("[안내] 원격디버깅 포트에 연결할 수 없습니다. 프로필 복사본으로 새 창을 띄워 진행합니다.")
                if not REAL_PROFILE_COPY_DIR.exists() or not (REAL_PROFILE_COPY_DIR / "Default").exists():
                    if not _copy_real_chrome_profile():
                        print("[오류] 프로필 복사 실패. Chrome을 모두 닫고 다시 실행하거나, chrome_Profile67_원격디버깅.bat 으로 Chrome을 켠 뒤 실행하세요.")
                        sys.exit(1)
                launch_kw["user_data_dir"] = str(REAL_PROFILE_COPY_DIR)
                try:
                    context = _launch(lambda no_viewport=False: p.chromium.launch_persistent_context(**{**launch_kw, **({"no_viewport": True} if no_viewport else {})}))
                except TypeError:
                    context = p.chromium.launch_persistent_context(**launch_kw)
        if context is None:
            print("[오류] 브라우저 실행 실패.")
            sys.exit(1)
        page = context.new_page()
        main_upload_impl(page, items, PASSWORD)
        print("모든 사업자 업로드 완료.")
        input("엔터를 누르면 종료합니다...")


if __name__ == "__main__":
    main()
