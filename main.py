import io
import json
import os
import pandas as pd
import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel():
    """requests 세션을 이용해 오피넷 엑셀 파일 바이너리 직접 다운로드"""
    session = requests.Session()
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.opinet.co.kr/user/opdown/opDownload.do",
        "Origin": "https://www.opinet.co.kr",
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
    # 1. 먼저 메인 다운로드 페이지에 접속하여 쿠키 및 세션 초기화
    session.get("https://www.opinet.co.kr/user/opdown/opDownload.do", headers=headers)
    
    # 2. 오피넷 서버로 사업자별 현재 판매가격(주유소) 엑셀 다운로드 POST 요청
    download_url = "https://www.opinet.co.kr/user/opdown/opDownload.do"
    payload = {
        "BTN_DIV": "2",
        "FILE_DIV": "1",
        "SAVE_DIV": "XLS",
        "API_GDN": "A"
    }
    
    response = session.post(download_url, data=payload, headers=headers)
    
    if response.status_code != 200 or len(response.content) < 1000:
        raise Exception(f"엑셀 다운로드 실패 (상태 코드: {response.status_code}, 크기: {len(response.content)})")
        
    return response.content

def update_google_sheet(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드된 유가 데이터를 파싱하여 구글 스프레드시트에 기록"""
    df = None
    
    # 1. 엑셀 파일 형식(XLS/XLSX) 파싱 시도
    for engine in ['xlrd', 'openpyxl', None]:
        try:
            df = pd.read_excel(io.BytesIO(excel_bytes), engine=engine)
            if df is not None and len(df) > 0:
                break
        except Exception:
            continue
            
    # 2. HTML Table 형식 XLS 파싱 시도
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

    # NaN 처리 및 구글 시트 입력용 2차원 리스트 구성
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
    update_google_sheet(excel_bytes, spreadsheet_id, sa_json)
    print("모든 작업이 성공적으로 완료되었습니다.")
