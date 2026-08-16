import io
import json
import os
import pandas as pd
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel_via_browser():
    """실제 브라우저를 실행하여 오피넷 화면에서 엑셀저장 버튼을 직접 클릭 및 다운로드"""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            accept_downloads=True
        )
        page = context.new_page()
        
        print("오피넷 유가내려받기 페이지 접속 중...")
        page.goto("https://www.opinet.co.kr/user/opdown/opDownload.do", wait_until="networkidle", timeout=60000)
        
        # '사업자별 현재 판매가격' 영역의 엑셀저장 버튼 (javascript:fn_Download(2))
        excel_btn = page.locator("a[href*='javascript:fn_Download(2)']")
        excel_btn.wait_for(state="visible", timeout=15000)
        
        print("엑셀저장 버튼 클릭 및 다운로드 대기...")
        with page.expect_download(timeout=60000) as download_info:
            excel_btn.click()
            
        download = download_info.value
        temp_path = download.path()
        
        with open(temp_path, "rb") as f:
            excel_bytes = f.read()
            
        browser.close()
        return excel_bytes

def update_google_sheet(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드한 엑셀 데이터를 구글 스프레드시트에 저장"""
    # 오피넷 엑셀 파일 파싱 (HTML Table 형식 또는 Binary XLS 형식 대응)
    try:
        df = pd.read_excel(io.BytesIO(excel_bytes))
    except Exception:
        dfs = pd.read_html(io.BytesIO(excel_bytes))
        df = dfs[0]
        
    df = df.fillna("")
    rows = [df.columns.tolist()] + df.astype(str).values.tolist()
    
    # 구글 시트 API 연동
    service_account_info = json.loads(sa_json_str)
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)
    
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_name = sheet_metadata['sheets'][0]['properties']['title']
    
    # 기존 시트 초기화 후 새 유가 데이터 반영
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
    
    print(f"구글 스프레드시트 갱신 성공: 총 {len(rows)}행의 유가 데이터 반영 완료")

if __name__ == "__main__":
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    sa_json = os.environ.get("GCP_SA_KEY")
    
    print("1. 브라우저 구동 및 유가 엑셀 다운로드 시작...")
    excel_bytes = download_opinet_excel_via_browser()
    
    print("2. 엑셀 데이터 구글 스프레드시트로 동기화 시작...")
    update_google_sheet(excel_bytes, spreadsheet_id, sa_json)
    print("모든 작업이 완료되었습니다.")
