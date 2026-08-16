import io
import json
import os
import requests
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet_excel():
    """fn_Download(2)의 요청을 재현하여 실제 엑셀 파일 바이너리를 다운로드"""
    url = "https://www.opinet.co.kr/user/opdown/opDownload.do"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.opinet.co.kr/user/opdown/opDownload.do",
        "Origin": "https://www.opinet.co.kr",
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
    # 폼(priceInfoVO)에서 fn_Download(2) 클릭 시 전송되는 폼 필드
    payload = {
        "BTN_DIV": "2",          # 2: 사업자별 현재 판매가격
        "FILE_DIV": "1",         # 1: 주유소
        "SAVE_DIV": "XLS",       # XLS 엑셀 파일
        "API_GDN": "A",
        "SIDO_NM": "",
        "SIGUN_NM": "",
        "netfunnel_key": ""
    }
    
    session = requests.Session()
    # 1. 초기 세션 쿠키 획득
    session.get(url, headers=headers)
    
    # 2. 엑셀 다운로드 POST 요청
    response = session.post(url, data=payload, headers=headers)
    
    # HTML 페이지가 반환되었는지 체크 (파일이 아니면 에러)
    if b"<!DOCTYPE html>" in response.content[:100]:
        raise Exception("오피넷에서 파일이 아닌 HTML 페이지가 반환되었습니다. 세션 또는 파라미터를 확인하세요.")
        
    return response.content

def update_google_sheet_with_excel(excel_bytes, spreadsheet_id, sa_json_str):
    """다운로드받은 엑셀 데이터를 파싱하여 구글 시트에 업로드"""
    # 엑셀 파일 읽기 (오피넷 엑셀 파일은 HTML-table 기반 또는 XLS 형식)
    try:
        df = pd.read_excel(io.BytesIO(excel_bytes))
    except Exception:
        # 오피넷 일부 XLS 파일은 HTML Table 포맷인 경우 지원
        dfs = pd.read_html(io.BytesIO(excel_bytes))
        df = dfs[0]

    # NaN(빈값) 처리 및 문자열 변환
    df = df.fillna("")
    rows = [df.columns.tolist()] + df.astype(str).values.tolist()

    # 구글 시트 API 연결
    service_account_info = json.loads(sa_json_str)
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)
    
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_name = sheet_metadata['sheets'][0]['properties']['title']
    
    # 기존 데이터 비우고 실제 유가 데이터만 채우기
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
    
    print(f"구글 스프레드시트 갱신 성공: 총 {len(rows)}행의 유가 데이터가 입력되었습니다.")

if __name__ == "__main__":
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    sa_json = os.environ.get("GCP_SA_KEY")
    
    print("1. 오피넷 엑셀 유가 데이터 다운로드 시도 중...")
    excel_bytes = download_opinet_excel()
    
    print("2. 엑셀 데이터 파싱 및 구글 시트 반영 중...")
    update_google_sheet_with_excel(excel_bytes, spreadsheet_id, sa_json)
    print("모든 작업이 정상 완료되었습니다.")
