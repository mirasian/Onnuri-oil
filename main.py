import io
import json
import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel_via_browser():
    """오피넷 유가 엑셀 다운로드 (버튼 로딩 대기 및 직접 클릭)"""
    with sync_playwright() as p:
        # 헤드리스 크롬 실행
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
        page.goto("https://www.opinet.co.kr/user/opdown/opDownload.do", timeout=60000)
        
        print("2. 엑셀저장 버튼 로딩 대기 중...")
        # '사업자별 현재 판매가격'의 엑셀저장 링크(fn_Download(2))가 화면에 렌더링될 때까지 대기
        save_btn = page.wait_for_selector('a[href*="fn_Download(2)"]', timeout=30000)
        time.sleep(2) # 안정성을 위한 짧은 대기
        
        print("3. 엑셀저장 버튼 클릭 및 다운로드 대기...")
        with page.expect_download(timeout=60000) as download_info:
            save_btn.click()
            
        download = download_info.value
        download_path = download.path()
        
        with open(download_path, "rb") as f:
            excel_bytes = f.read()
            
        browser.close()
        return excel_bytes

def update_google_sheet(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드한 엑셀 데이터를 파싱하여 구글 시트에 업데이트"""
    try:
        df = pd.read_excel(io.BytesIO(excel_bytes))
    except Exception:
        dfs = pd.read_html(io.BytesIO(excel_bytes))
        df = dfs[0]
        
    df = df.fillna("")
    rows = [df.columns.tolist()] + df.astype(str).values.tolist()
    
    service_account_info = json.loads(sa_json_str)
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)
    
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_name = sheet_metadata['sheets'][0]['properties']['title']
    
    # 기존 내용 비우기
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A:Z"
    ).execute()
    
    # 새 유가 데이터 작성
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A1",
        valueInputOption="USER_ENTERED",
        body={'values': rows}
    ).execute()
    
    print(f"구글 스프레드시트 갱신 성공: 총 {len(rows)}개 주유소 데이터 입력 완료")

if __name__ == "__main__":
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    sa_json = os.environ.get("GCP_SA_KEY")
    
    excel_bytes = download_opinet_excel_via_browser()
    update_google_sheet(excel_bytes, spreadsheet_id, sa_json)
    print("모든 작업이 성공적으로 완료되었습니다.")
