#!/usr/bin/env python3
"""fetch_from_drive.py - بينزّل ملف الإكسل من Google Drive.
الاستخدام:  python fetch_from_drive.py book.xlsx
Secrets المطلوبة:
  GDRIVE_KEY      = محتوى ملف JSON الخاص بالـ service account (كله)
  GDRIVE_FILE_ID  = رقم ملف Customers.xlsx على Drive
الملف لازم يكون متشارك (Viewer) مع إيميل الـ service account."""
import io, json, os, sys
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
out = sys.argv[1] if len(sys.argv) > 1 else 'book.xlsx'
key = os.environ.get('GDRIVE_KEY', '').strip()
fid = os.environ.get('GDRIVE_FILE_ID', '').strip()
if not key or not fid:
    sys.exit('GDRIVE_KEY أو GDRIVE_FILE_ID ناقصين في الـ Secrets.')
try:
    info = json.loads(key)
except Exception:
    sys.exit('GDRIVE_KEY مش JSON صحيح. انسخ محتوى ملف المفتاح كله.')

creds = service_account.Credentials.from_service_account_info(
    info, scopes=['https://www.googleapis.com/auth/drive.readonly'])
svc = build('drive', 'v3', credentials=creds, cache_discovery=False)
try:
    meta = svc.files().get(fileId=fid, fields='id,name,mimeType,modifiedTime', supportsAllDrives=True).execute()
    print('Drive file:', meta.get('name'), '| modified:', meta.get('modifiedTime'))
    if meta.get('mimeType') == 'application/vnd.google-apps.spreadsheet':
        req = svc.files().export_media(fileId=fid, mimeType=XLSX)       # Google Sheet -> xlsx
    else:
        req = svc.files().get_media(fileId=fid, supportsAllDrives=True)  # ملف xlsx عادي
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req, chunksize=8 * 1024 * 1024)
    done = False
    while not done:
        _, done = dl.next_chunk()
except HttpError as e:
    sys.exit('فشل التحميل من Drive (%s). اتأكد إن GDRIVE_FILE_ID صح وإن الملف متشارك مع: %s'
             % (getattr(e.resp, 'status', '?'), info.get('client_email', '?')))

data = buf.getvalue()
if len(data) < 1000 or data[:2] != b'PK':
    sys.exit('الملف اللي اتحمّل مش xlsx صحيح.')
with open(out, 'wb') as f:
    f.write(data)
print('saved', out, len(data) // 1024, 'KB')
