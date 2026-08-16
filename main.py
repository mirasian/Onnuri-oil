import io
import json
import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel():
    """새 팝업창(새 탭)에서 발생하는 다운로드 이벤트까지 모두 감지하여 엑셀 수신"""
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            accept_downloads=True
        )
        
        # 다운로드된 파일 저장용 변수
        download_container = []

        # 컨텍스트 내의 모든 탭/팝업에서 다운로드 이벤트 수신 대기
        def on_download(download):
            temp_path = download.path()
            with open(temp_path, "rb") as f:
                download_container.append(f.read())
            print(f"다운로드 파일 수신 성공: {download.suggested_filename}")

        # 새 페이지(팝업)가 열릴 때도 다운로드 리스너 연결
        def on_page(new_page):
            new_page.on("download", on_download)

        context.on("page", on_page)

        page = context.new_page()
        page.on("download", on_download)

        print("1. 오피넷 유가내려받기 페이지 접속 중...")
        page.goto("https://www.opinet.co.kr/user/opdown/opDownload.do", wait_until="networkidle", timeout=60000)

        # 엑셀저장 버튼 대기
        btn = page.wait_for_selector('a[href*="fn_Download(2)"]', timeout=30000)
        time.sleep(2)

        print("2. 엑셀저장 버튼 클릭 및 다운로드 대기...")
        btn.click()

        # 파일 다운로드가 완료될 때까지 최대 30초 대기
        for _ in range(30):
            if download_container:
                break
            time.sleep(1)

        # 클릭으로 감지가 안 되었을 경우 JS 강제 실행 후 재대기
        if not download_container:
            print("자바스크립트 fn_Download(2) 직접 호출...")
            page.evaluate("fn_Download(2);")
            for _ in range(30):
                if download_container:
                    break
                time.sleep(1)

        browser.close()

        if not download_container:
            raise Exception("엑셀 파일 다운로드에 실패했습니다. (다운로드 이벤트 감지 안 됨)")

        return download_container[0]

def parse_and_update_sheet(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드된 유가 데이터를 파싱하여 구글 스프레드시트에 기록"""
    df = None
    
    # 1. HTML Table 포맷 XLS 파싱 시도
    for enc in ['cp949', 'euc-kr', 'utf-8']:
        try:
            text = excel_bytes.decode(enc)
            dfs = pd.read_html(io.StringIO(text))
            if dfs and len(dfs[0]) > 0:
                df = dfs[0]
                break
        except Exception:
            continue

    # 2. 일반 바이너리 엑셀(XLS/XLSX) 파싱 시도
    if df is None or len(df) == 0:
        for engine in ['xlrd', 'openpyxl', None]:
            try:
                df = pd.read_excel(io.BytesIO(excel_bytes), engine=engine)
                if df is not None and len(df) > 0:
                    break
            except Exception:
                continue

    # 3. CSV 파싱 시도
    if df is None or len(df) == 0:
        for enc in ['cp949', 'euc-kr', 'utf-8']:
            try:
                df = pd.read_csv(io.BytesIO(excel_bytes), encoding=enc)
                if df is not None and len(df) > 0:
                    break
            except Exception:
                continue

    if df is None or len(df) == 0:
        raise Exception(f"유가 데이터 파싱 실패 (수신 바이트 크기: {len(excel_bytes)})")

    # NaN 정리 및 2차원 리스트 구성
    df = df.fillna("")
    rows = [df.columns.tolist()] + df.astype(str).values.tolist()
    
    # 구글 스프레드시트 API 연동
    service_account_info = json.loads(sa_json_str)
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)
    
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_name = sheet_metadata['sheets'][0]['properties']['title']
    
    # 데이터 초기화 및 갱신
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A:Z"
    ).execute()
    
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A1",
        valueInputOption="USER_ENTERED",
        body={'values': rows}
    ).execute()
    
    print(f"구글 스프레드시트 갱신 성공: 총 {len(rows)}개 행 입력 완료")

if __name__ == "__main__":
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    sa_json = os.environ.get("GCP_SA_KEY")
    
    excel_bytes = download_opinet_excel()
    parse_and_update_sheet(excel_bytes, spreadsheet_id, sa_json)
    print("모든 작업이 성공적으로 완료되었습니다.")
