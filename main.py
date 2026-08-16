import io
import json
import os
from datetime import datetime
import pytz
import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

def download_opinet_excel():
    """fn_Download(2)와 동일하게 사업자별 현재 판매가격 주유소 엑셀(XLS) 다운로드"""
    url = "https://www.opinet.co.kr/user/opdown/opDownload.do"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.opinet.co.kr/user/opdown/opDownload.do",
        "Origin": "https://www.opinet.co.kr"
    }
    
    # fn_Download(2) 엑셀 다운로드 요청 값
    payload = {
        "BTN_DIV": "2",
        "FILE_DIV": "1",
        "SAVE_DIV": "XLS"
    }
    
    session = requests.Session()
    session.get(url, headers=headers)
    
    response = session.post(url, data=payload, headers=headers)
    if response.status_code == 200 and len(response.content) > 1000:
        return response.content
    else:
        raise Exception(f"엑셀 다운로드 실패 (응답 코드: {response.status_code})")

def upload_excel_to_drive(file_content, filename, folder_id, sa_json_str):
    """구글 드라이브 폴더에 엑셀 파일 업로드"""
    service_account_info = json.loads(sa_json_str)
    scopes = ['https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    service = build('drive', 'v3', credentials=creds)
    
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    
    media = MediaIoBaseUpload(
        io.BytesIO(file_content),
        mimetype='application/vnd.ms-excel',
        resumable=True
    )
    
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()
    
    print(f"구글 드라이브 업로드 완료: {filename} (ID: {file.get('id')})")

if __name__ == "__main__":
    kst = pytz.timezone('Asia/Seoul')
    today_str = datetime.now(kst).strftime('%Y%m%d')
    filename = f"현재_판매가격(주유소)_{today_str}.xls"
    
    folder_id = os.environ.get("DRIVE_FOLDER_ID")
    sa_json = os.environ.get("GCP_SA_KEY")
    
    print("1. 오피넷 엑셀 파일 다운로드 중 (fn_Download(2))...")
    excel_content = download_opinet_excel()
    
    print(f"2. 구글 드라이브 '오늘의가격' 폴더로 업로드 중... (파일명: {filename})")
    upload_excel_to_drive(excel_content, filename, folder_id, sa_json)
    print("모든 작업이 완료되었습니다.")
