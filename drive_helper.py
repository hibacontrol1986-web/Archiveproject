import os
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "service_account.json")

def get_drive_service():
    try:
        # أوامر التشخيص
        print(f"\n--- DEBUG: Looking for file at: {SERVICE_ACCOUNT_FILE}")
        print(f"--- DEBUG: Does the file exist? {os.path.exists(SERVICE_ACCOUNT_FILE)}")
        
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        print(f"Error loading service account: {e}")
        return None
def find_folder_id(service, folder_name, parent_id=None):
    safe_name = folder_name.replace("'", "\\'")
    if parent_id:
        query = f"name='{safe_name}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed=false"
    else:
        # إذا لم يتم تحديد مجلد أب، نبحث في كل مكان (بما فيها المجلدات المشاركة)
        query = f"name='{safe_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        
    results = service.files().list(q=query, fields="files(id, name)", includeItemsFromAllDrives=True, supportsAllDrives=True).execute()
    files = results.get('files', [])
    return files[0]['id'] if files else None

def list_files_in_folder(service, folder_id):
    query = f"'{folder_id}' in parents and trashed=false"
    results = service.files().list(
        q=query, pageSize=1000, 
        fields="files(id, name, mimeType)"
    ).execute()
    return results.get('files', [])

def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def find_file_by_name(service, file_name):
    safe_name = file_name.replace("'", "\\'")
    query = f"name='{safe_name}' and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    files = results.get('files', [])
    return files[0]['id'] if files else None

def find_file_in_tree(service, file_name, folder_id):
    safe_name = file_name.replace("'", "\\'")
    query = f"name='{safe_name}' and '{folder_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    files = results.get('files', [])
    if files: return files[0]['id']
    
    query_folders = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    folders_res = service.files().list(q=query_folders, fields="files(id)").execute()
    for folder in folders_res.get('files', []):
        res = find_file_in_tree(service, file_name, folder['id'])
        if res: return res
    return None