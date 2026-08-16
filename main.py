import io
import json
import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel():
    """브라우저에서 fn_Download(2) 클릭 시 발생하는 네트워크 응답 바이너리를 직접 캡처"""
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

        # 캡처된 엑셀 데이터를 저장할 버퍼
        captured_data = []

        # 오피넷 다운로드 응답 가로채기 핸들러
        def handle_response(response):
            try:
                # 다운로드 요청 또는 첨부파일(Content-Disposition) 응답 캡처
                content_disposition = response.headers.get("content-disposition", "")
                content_type = response.headers.get("content-type", "")
                
                if "attachment" in content_disposition or "excel" in content_type or "vnd.ms-excel" in content_type:
                    body = response.body()
                    if len(body) > 1000:
                        captured_data.append(body)
            except Exception:
                pass

        page.on("response", handle_response)

        print("1. 오피넷 페이지 접속 중...")
        page.goto("https://www.opinet.co.kr/user/opdown/opDownload.do", wait_until="networkidle", timeout=60000)

        # 엑셀저장 버튼 요소 대기
        btn = page.wait_for_selector('a[href*="fn_Download(2)"]', timeout=30000)
        time.sleep(2)

        print("2. 엑셀저장(fn_Download(2)) 실행 및 데이터 가로채기...")
        # 1) 다운로드 이벤트 트리거 또는 클릭
        try:
            with page.expect_download(timeout=10000) as download_info:
                btn.click()
            download = download_info.value
            with open(download.path(), "rb") as f:
                captured_data.append(f.read())
        except Exception:
            # 브라우저 다운로드 이벤트가 안 뜨는 경우 evaluate 실행
            page.evaluate("fn_Download(2);")
            time.sleep(5)

        browser.close()

        if not captured_data:
            raise Exception("엑셀 응답 데이터를 캡처하지 못했습니다.")

        return captured_data[-1]

def parse_and_update_sheet(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드된 유가 바이너리 데이터를 분석하여 구글 스프레드시트에 기록"""
    df = None
    
    # 1. 엑셀 엔진으로 읽기 시도
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
        raise Exception(f"유효한 유가 데이터 파싱 실패 (수신 바이트 크기: {len(excel_bytes)})")

    # NaN 처리 및 구글 시트 입력용 2차원 리스트 구성
    df = df.fillna("")
    rows = [df.columns.tolist()] + df.astype(str).values.tolist()
    
    # 구글 시트 API 연동
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
