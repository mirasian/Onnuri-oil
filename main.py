import csv
import io
import json
import os
import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def download_opinet():
    url = "https://www.opinet.co.kr/user/opdown/opDownload.do"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.opinet.co.kr/user/opdown/opDownload.do",
    }
    payload = {
        "BTN_DIV": "2",
        "FILE_DIV": "1",
        "SAVE_DIV": "CSV"
    }
    session = requests.Session()
    session.get(url, headers=headers)
    response = session.post(url, data=payload, headers=headers)
    
    if response.status_code == 200 and len(response.content) > 1000:
        try:
            return response.content.decode('cp949')
        except UnicodeDecodeError:
            return response.content.decode('utf-8')
    else:
        raise Exception(f"오피넷 다운로드 실패 (코드: {response.status_code})")

def update_google_sheet(csv_text, spreadsheet_id, sa_json_str):
    service_account_info = json.loads(sa_json_str)
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)
    
    # CSV 파싱
    csv_reader = csv.reader(io.StringIO(csv_text))
    rows = list(csv_reader)
    
    # 첫 번째 시트 이름 가져오기
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_name = sheet_metadata['sheets'][0]['properties']['title']
    
    # 기존 내용 지우고 새 데이터 입력
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
    
    print(f"구글 스프레드시트 갱신 성공: 총 {len(rows)}행")

if __name__ == "__main__":
    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    sa_json = os.environ.get("GCP_SA_KEY")
    
    print("1. 오피넷 유가 다운로드 시작...")
    csv_text = download_opinet()
    
    print("2. 구글 스프레드시트 업데이트 시작...")
    update_google_sheet(csv_text, spreadsheet_id, sa_json)
    print("모든 작업 완료!")
