import io
import json
import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel():
    """헤드리스 브라우저 환경에서 탭 로딩을 완료하고 폼 submit을 통해 정식 엑셀 다운로드"""
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            accept_downloads=True
        )
        page = context.new_page()

        print("1. 오피넷 유가내려받기 페이지 접속 중...")
        page.goto("https://www.opinet.co.kr/user/opdown/opDownload.do", wait_until="networkidle", timeout=60000)

        # 탭 내부의 '사업자별 현재 판매가격' 엑셀저장 버튼이 렌더링될 때까지 대기
        btn_selector = 'a[href*="fn_Download(2)"]'
        page.wait_for_selector(btn_selector, timeout=30000)
        time.sleep(2)

        print("2. 엑셀 파일 다운로드 실행 중...")
        # 엑셀저장 버튼 클릭 시 발생하는 다운로드 파일 직접 수신
        with page.expect_download(timeout=60000) as download_info:
            page.evaluate("""
                () => {
                    if (typeof fn_Download === 'function') {
                        fn_Download(2);
                    } else {
                        document.querySelector('a[href*="fn_Download(2)"]').click();
                    }
                }
            """)
            
        download = download_info.value
        temp_path = download.path()
        
        with open(temp_path, "rb") as f:
            excel_bytes = f.read()

        browser.close()
        return excel_bytes

def parse_and_update_sheet(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드된 오피넷 주유소 유가 데이터를 구글 스프레드시트에 기록"""
    df = None
    
    # 1. HTML Table 형식 XLS 파싱 (오피넷 엑셀 파일 표준 형식)
    for enc in ['cp949', 'euc-kr', 'utf-8']:
        try:
            text = excel_bytes.decode(enc)
            dfs = pd.read_html(io.StringIO(text))
            if dfs and len(dfs[0]) > 0:
                df = dfs[0]
                break
        except Exception:
            continue

    # 2. 일반 바이너리 엑셀 파싱 시도
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

    # NaN 빈칸 처리 및 2차원 리스트 구성
    df = df.fillna("")
    rows = [df.columns.tolist()] + df.astype(str).values.tolist()
    
    # 구글 스프레드시트 API 연동
    service_account_info = json.loads(sa_json_str)
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)
    
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_name = sheet_metadata['sheets'][0]['properties']['title']
    
    # 기존 데이터 초기화 및 갱신
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
    
    print(f"구글 스프레드시트 갱신 완료: 총 {len(rows)}개 행(주유소 유가 데이터) 입력 완료")

if __name__ == "__main__":
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    sa_json = os.environ.get("GCP_SA_KEY")
    
    excel_bytes = download_opinet_excel()
    parse_and_update_sheet(excel_bytes, spreadsheet_id, sa_json)
    print("모든 작업이 성공적으로 완료되었습니다.")
