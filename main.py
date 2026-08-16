import io
import json
import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel():
    """브라우저 환경에서 fn_Download(2) 실행 후 서버 응답 본문을 직접 캡처"""
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

        # 엑셀저장 버튼 요소 대기
        page.wait_for_selector('a[href*="fn_Download(2)"]', timeout=30000)
        time.sleep(2)

        print("2. 엑셀저장(fn_Download(2)) 실행 및 다운로드 응답 대기...")
        
        # 폼 전송 시 서버가 돌려주는 실제 HTTP 응답(URL에 opDownload 또는 download가 포함된 응답) 대기
        try:
            with page.expect_response(lambda r: "opDownload" in r.url or "download" in r.url.lower(), timeout=30000) as response_info:
                page.evaluate("fn_Download(2);")
            
            response = response_info.value
            excel_bytes = response.body()
            print(f"응답 수신 성공 (상태 코드: {response.status}, 크기: {len(excel_bytes)} 바이트)")
        except Exception as e:
            print(f"expect_response 대기 실패, download 이벤트 재시도: {e}")
            with page.expect_download(timeout=30000) as download_info:
                page.locator('a[href*="fn_Download(2)"]').click()
            download = download_info.value
            with open(download.path(), "rb") as f:
                excel_bytes = f.read()

        browser.close()
        return excel_bytes

def parse_and_update_sheet(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드된 데이터를 분석하여 구글 스프레드시트에 기록"""
    df = None
    
    # 1. 엑셀 파싱 시도
    for engine in ['xlrd', 'openpyxl', None]:
        try:
            df = pd.read_excel(io.BytesIO(excel_bytes), engine=engine)
            if df is not None and len(df) > 0:
                break
        except Exception:
            continue
            
    # 2. HTML Table 형식 XLS 읽기 시도
    if df is None or len(df) == 0:
        for enc in ['cp949', 'euc-kr', 'utf-8']:
            try:
                text = excel_bytes.decode(enc)
                dfs = pd.read_html(io.StringIO(text))
                if dfs and len(dfs[0]) > 0:
                    df = dfs[0]
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
        preview = excel_bytes[:300].decode('cp949', errors='ignore')
        raise Exception(f"유효한 유가 데이터 테이블을 파싱하지 못했습니다. 서버 반환 내용: {preview}")

    # NaN 값 빈칸 처리 및 2차원 배열 변환
    df = df.fillna("")
    rows = [df.columns.tolist()] + df.astype(str).values.tolist()
    
    # 구글 스프레드시트 API 연동
    service_account_info = json.loads(sa_json_str)
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)
    
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_name = sheet_metadata['sheets'][0]['properties']['title']
    
    # 기존 데이터 비우고 갱신
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
    
    print(f"구글 스프레드시트 갱신 완료: 총 {len(rows)}개 행 입력 완료")

if __name__ == "__main__":
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    sa_json = os.environ.get("GCP_SA_KEY")
    
    excel_bytes = download_opinet_excel()
    parse_and_update_sheet(excel_bytes, spreadsheet_id, sa_json)
    print("모든 작업이 성공적으로 완료되었습니다.")
