"""Read publicly approved Google spreadsheets or Drive XLSX files only."""
import re
from io import BytesIO
from zipfile import ZipFile, is_zipfile
from urllib.parse import urlparse, parse_qs
import requests
import streamlit as st

MAX_BYTES = 25 * 1024 * 1024

def export_url(link):
    link = link.strip()
    u = urlparse(link)
    if u.scheme != 'https' or u.hostname not in ('docs.google.com','drive.google.com'):
        raise ValueError('Gunakan link sumber Google Sheets atau Google Drive (bukan link Google Sites).')
    if u.hostname == 'docs.google.com':
        m = re.match(r'/spreadsheets/d/([A-Za-z0-9_-]+)(?:/|$)', u.path)
        if m and m.group(1) != 'e':
            return 'https://docs.google.com/spreadsheets/d/'+m.group(1)+'/export?format=xlsx'
    else:
        m = re.match(r'/file/d/([A-Za-z0-9_-]+)(?:/|$)',u.path)
        ident = m.group(1) if m else parse_qs(u.query).get('id',[''])[0]
        if re.fullmatch(r'[A-Za-z0-9_-]+',ident):
            return 'https://drive.google.com/uc?export=download&id='+ident
    raise ValueError('Link tidak dikenali. Salin link Bagikan dari file sumber.')

@st.cache_data(ttl=300, show_spinner=False, max_entries=24)
def download_excel(url):
    with requests.get(url, timeout=(10,60), stream=True) as response:
        response.raise_for_status()
        buf=bytearray()
        for chunk in response.iter_content(65536):
            buf.extend(chunk)
            if len(buf)>MAX_BYTES: raise ValueError('File melebihi batas 25 MB.')
    blob=bytes(buf)
    if not is_zipfile(BytesIO(blob)):
        raise ValueError('Respons bukan XLSX. Periksa izin Viewer, format file, atau halaman konfirmasi unduhan Google.')
    with ZipFile(BytesIO(blob)) as z:
        if 'xl/workbook.xml' not in z.namelist(): raise ValueError('File bukan workbook Excel.')
        if sum(i.file_size for i in z.infolist())>150*1024*1024: raise ValueError('Ukuran workbook setelah ekstraksi terlalu besar.')
    return blob

class WorkbookSource:
    def __init__(self,name,blob): self.name=name; self.blob=blob
    def getvalue(self): return self.blob
