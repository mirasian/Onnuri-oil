import io
import json
import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel():
    """브라우저 내부 fetch API를 이용해 세션/보안을 완벽히 유지한 채 엑셀 바이너리를 직접 수신"""
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        
        print("1. 오피넷 유가내려받기 페이지 접속 중...")
        page.goto("https://www.opinet.co.kr/user/opdown/opDownload.do", wait_until="networkidle", timeout=60000)
        
        # 버튼이 렌더링될 때까지 대기
        page.wait_for_selector('a[href*="fn_Download(2)"]', timeout=30000)
        time.sleep(2)
        
        print("2. 브라우저 컨텍스트 내에서 엑셀 데이터 직접 다운로드 요청 중...")
        # 브라우저의 활성 세션을 사용하여 폼 데이터를 fetch로 직접 전송하고 바이너리 배열을 반환받음
        js_code = """
        async () => {
            const formData = new URLSearchParams();
            formData.append("BTN_DIV", "2");
            formData.append("FILE_DIV", "1");
            formData.append("SAVE_DIV", "XLS");
            formData.append("API_GDN", "A");
            formData.append("SIDO_NM", "");
            formData.append("SIGUN_NM", "");
            formData.append("netfunnel_key", "");

            const response = await fetch("/user/opdown/opDownload.do", {
                method: "POST",
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded"
                },
                body: formData.toString()
            });

            const buffer = await response.arrayBuffer();
            return Array.from(new Uint8Array(buffer));
        }
        """
        byte_array = page.evaluate(js_code)
        excel_bytes = bytes(byte_array)
        
        browser.close()
        return excel_bytes

def parse_and_update_sheet(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드된 유가 데이터를 파싱하여 구글 스프레드시트에 기록"""
    df = None
    
    # 1. 엑셀 파싱 시도
    for engine in ['xlrd', 'openpyxl', None]:
        try:
            df = pd.read_excel(io.BytesIO(excel_bytes), engine=engine)
            break
        except Exception:
            continue
            
    # 2. HTML Table 형식 XLS 파싱 시도
    if df is None:
        for enc in ['cp949', 'euc-kr', 'utf-8']:
            try:
                text = excel_bytes.decode(enc)
                dfs = pd.read_html(io.StringIO(text))
                if dfs and len(dfs[0]) > 0:
                    df = dfs[0]
                    break
            except Exception:
                continue

    # 3. CSV 포맷 파싱 시도
    if df is None:
        for enc in ['cp949', 'euc-kr', 'utf-8']:
            try:
                df = pd.read_csv(io.BytesIO(excel_bytes), encoding=enc)
                break
            except Exception:
                continue

    if df is None or len(df) == 0:
        preview = excel_bytes[:200]
        raise Exception(f"유효한 유가 데이터 테이블을 찾지 못했습니다. (데이터 앞부분: {preview})")

    # NaN 정리 및 2차원 리스트 변환
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
    
    print(f"구글 스프레드시트 업데이트 완료: 총 {len(rows)}개 주유소 데이터 반영")

if __name__ == "__main__":
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    sa_json = os.environ.get("GCP_SA_KEY")
    
    excel_bytes = download_opinet_excel()
    parse_and_update_sheet(excel_bytes, spreadsheet_id, sa_json)
    print("모든 작업이 성공적으로 완료되었습니다.")
