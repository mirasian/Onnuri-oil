import io
import json
import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel():
    """브라우저 환경에서 실제 fn_Download(2) 엑셀 다운로드를 가로채서 데이터 수신"""
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
        
        print("1. 오피넷 유가내려받기 페이지 로딩 중...")
        page.goto("https://www.opinet.co.kr/user/opdown/opDownload.do", wait_until="networkidle", timeout=60000)
        
        # 엑셀 다운로드 버튼이 렌더링될 때까지 대기
        page.wait_for_selector('a[href*="fn_Download(2)"]', timeout=30000)
        time.sleep(2)
        
        print("2. 엑셀 다운로드 요청 실행 중...")
        with page.expect_download(timeout=60000) as download_info:
            # 브라우저 내부에서 fn_Download(2)를 직접 호출
            page.evaluate("fn_Download(2);")
            
        download = download_info.value
        temp_file_path = download.path()
        
        print(f"다운로드된 파일 확인: {download.suggested_filename}")
        with open(temp_file_path, "rb") as f:
            excel_bytes = f.read()
            
        browser.close()
        return excel_bytes

def parse_and_update_sheet(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드된 엑셀/HTML 데이터를 분석하여 구글 스프레드시트에 입력"""
    df = None
    
    # 1. 일반 엑셀(XLS/XLSX) 파싱 시도
    for engine in ['xlrd', 'openpyxl', None]:
        try:
            df = pd.read_excel(io.BytesIO(excel_bytes), engine=engine)
            break
        except Exception:
            continue
            
    # 2. HTML Table 형태의 XLS 파일 파싱 시도
    if df is None:
        try:
            for enc in ['cp949', 'euc-kr', 'utf-8']:
                try:
                    text = excel_bytes.decode(enc)
                    dfs = pd.read_html(io.StringIO(text))
                    if dfs and len(dfs[0]) > 0:
                        df = dfs[0]
                        break
                except Exception:
                    continue
        except Exception as e:
            print("HTML 파싱 오류:", e)
            
    # 3. CSV 텍스트 형태 파싱 시도
    if df is None:
        for enc in ['cp949', 'euc-kr', 'utf-8']:
            try:
                df = pd.read_csv(io.BytesIO(excel_bytes), encoding=enc)
                break
            except Exception:
                continue

    if df is None or len(df) == 0:
        # 반환된 데이터가 올바르지 않을 때 앞부분 내용 출력 확인
        preview = excel_bytes[:300]
        raise Exception(f"유효한 유가 데이터 테이블을 찾을 수 없습니다. (응답 앞부분: {preview})")

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
    
    # 기존 데이터 초기화 후 갱신
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
