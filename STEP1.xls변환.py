import os
import time
import pythoncom
import win32com.client as win32

RPC_E_CALL_REJECTED = -2147418111  # "피호출자가 호출을 거부했습니다."

# 1이면 Excel 창을 보이게 해서(다른 인스턴스·대화상자 때문에 멈출 때)원인 확인
STEP1_EXCEL_VISIBLE = os.environ.get("STEP1_EXCEL_VISIBLE", "").lower() in ("1", "true", "yes")


def excel_call_with_retry(func, *args, retries=5, delay=0.5, backoff=1.5, **kwargs):
    attempt = 0
    while True:
        try:
            return func(*args, **kwargs)
        except pythoncom.com_error as e:
            hr = getattr(e, "hresult", e.args[0] if e.args else None)
            if hr == RPC_E_CALL_REJECTED and attempt < retries:
                time.sleep(delay)
                delay *= backoff
                attempt += 1
                continue
            raise


def _collect_xls_jobs(folder):
    """변환 대상 (.xls만, 이미 .xlsx 있으면 dst 경로는 수집 시 제외는 루프에서)."""
    jobs = []
    for root, _, files in os.walk(folder):
        for f in files:
            name = f.lower()
            if name.endswith(".xls") and not name.endswith(".xlsx"):
                src = os.path.join(root, f)
                dst = os.path.splitext(src)[0] + ".xlsx"
                jobs.append((src, dst))
    return jobs


def convert_xls_to_xlsx(folder):
    print(f"[STEP1] 작업 폴더: {folder}", flush=True)
    xls_jobs = _collect_xls_jobs(folder)
    if not xls_jobs:
        print(
            "[STEP1] 변환할 .xls 파일이 없습니다. "
            "(다운로드한 파일이 .xlsx이거나 폴더가 비었을 수 있습니다. STEP1-1로 진행합니다.)",
            flush=True,
        )
        return

    need_excel = []
    for src, dst in xls_jobs:
        if os.path.exists(dst):
            print(f"[SKIP] {os.path.relpath(dst, folder)} already exists", flush=True)
        else:
            need_excel.append((src, dst))

    if not need_excel:
        print("[STEP1] 모두 이미 .xlsx로 존재합니다.", flush=True)
        return

    print(
        f"[STEP1] Microsoft Excel 시작… 대상 {len(need_excel)}개 "
        f"(처음이면 1~2분 걸릴 수 있습니다. 멈추면 작업 관리자에서 EXCEL.EXE 확인)",
        flush=True,
    )
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        vis = 1 if STEP1_EXCEL_VISIBLE else 0
        for prop, val in (("Visible", vis), ("DisplayAlerts", False)):
            try:
                setattr(excel, prop, val)
            except Exception:
                pass

        for src, dst in need_excel:
            print(f"[OPEN] {os.path.relpath(src, folder)}", flush=True)
            # UpdateLinks=0, ReadOnly=True → 다른 Excel이 연 파일·잠금 시 대기 완화
            # UpdateLinks=0, ReadOnly=True (다른 Excel 점유·잠금 시 무응답 완화)
            wb = excel_call_with_retry(excel.Workbooks.Open, src, 0, True)
            try:
                print(f"[SAVE] -> {os.path.relpath(dst, folder)}", flush=True)
                excel_call_with_retry(wb.SaveAs, dst, FileFormat=51)
            finally:
                excel_call_with_retry(wb.Close, False)
        print("[DONE] All .xls converted to .xlsx", flush=True)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    target_folder = r"C:\Users\USER\Documents\국내위탁\마이박스\26년1월\1번사업자"
    convert_xls_to_xlsx(target_folder)
