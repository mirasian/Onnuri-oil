import io
import json
import os
import time
import pandas as pd
import requests
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel():
    """브라우저로 오피넷 유효 세션을 생성한 후 엑셀 데이터를 직접 수신"""
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        
        print("1. 오피넷 유가내려받기 페이지 접속 및 세션 획득 중...")
        page.goto("https://www.opinet.co.kr/user/opdown/opDownload.do", timeout=60000)
        
        # 다운로드 버튼이 렌더링될 때까지 대기
        page.wait_for_selector('a[href*="fn_Download(2)"]', timeout=30000)
        time.sleep(2)
        
        # 브라우저에 저장된 오피넷 쿠키 추출
        cookies = context.cookies()
        session_cookies = {c['name']: c['value'] for c in cookies}
        browser.close()

    print("2. 엑셀 데이터 파일 수신 요청 중...")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.opinet.co.kr/user/opdown/opDownload.do",
        "Origin": "https://www.opinet.co.kr",
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
    # 엑셀저장(fn_Download(2)) 전송 데이터
    payload = {
        "BTN_DIV": "2",
        "FILE_DIV": "1",
        "SAVE_DIV": "XLS",
        "API_GDN": "A",
        "SIDO_NM": "",
        "SIGUN_NM": "",
        "netfunnel_key": ""
    }
    
    res = requests.post(
        "https://www.opinet.co.kr/user/opdown/opDownload.do",
        data=payload,
        headers=headers,
        cookies=session_cookies
    )
    
    if res.status_code != 200 or len(res.content) < 1000:
        raise Exception(f"엑셀 파일 다운로드 실패 (상태 코드: {res.status_code})")
        
    return res.content

def update_google_sheet(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드받은 엑셀 데이터를 구글 스프레드시트에 입력"""
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
    
    # 기존 시트 비우고 새 유가 데이터 작성
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
    update_google_sheet(excel_bytes, spreadsheet_id, sa_json)
    print("모든 작업이 완료되었습니다.")
