import io
import json
import os
from datetime import datetime
import pytz
import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

def download_opinet():
    """오피넷에서 사업자별 현재 판매가격(주유소) CSV 파일 다운로드"""
    url = "https://www.opinet.co.kr/user/opdown/opDownload.do"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.opinet.co.kr/user/opdown/opDownload.do",
    }
    
    payload = {
        "BTN_DIV": "2",     # 사업자별 현재 판매가격
        "FILE_DIV": "1",    # 1: 주유소
        "SAVE_DIV": "CSV"   # CSV 형식
    }
    
    session = requests.Session()
    session.get(url, headers=headers)
    
    response = session.post(url, data=payload, headers=headers)
    if response.status_code == 200 and len(response.content) > 1000:
        return response.content
    else:
        raise Exception(f"오피넷 다운로드 실패 (상태 코드: {response.status_code})")

def upload_to_drive(file_content, filename, folder_id, sa_json_str):
    """구글 드라이브 지정 폴더로 파일 업로드"""
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
        mimetype='text/csv',
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
    filename = f"현재_판매가격(주유소)_{today_str}.csv"
    
    folder_id = os.environ.get("DRIVE_FOLDER_ID")
    sa_json = os.environ.get("GCP_SA_KEY")
    
    print("1. 오피넷 유가 데이터 수집 시작...")
    content = download_opinet()
    
    print("2. 구글 드라이브 업로드 시작...")
    upload_to_drive(content, filename, folder_id, sa_json)
    print("모든 작업이 성공적으로 완료되었습니다.")
