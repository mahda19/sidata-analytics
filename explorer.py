import re
import math
from io import BytesIO
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

COLS=['Tahun','Bab','Sheet','Judul','Kolom','Indikator','Wilayah','Sel','Nilai','Status']
def extract(blob,bab):
    wb=load_workbook(BytesIO(blob),data_only=True); rows=[]; issues=[]
    try:
        for ws in wb.worksheets:
            if ws.title.upper()=='REKAP': continue
            period=re.search(r'\b20\d{2}\b',str(ws['A3'].value or ''))
            regions=[(c.row,str(c.value).strip()) for (c,) in ws.iter_rows(min_col=2,max_col=2) if isinstance(c.value,str) and re.match(r'^\s*(Kab\.\s*|Kab\s|Kabupaten\s|Kota\s)',c.value,re.I)]
            if not period or len(regions)!=35 or len({v.casefold() for _,v in regions})!=35:
                issues.append((bab,ws.title,'Tahun/wilayah belum dipetakan')); continue
            regions=[(r,re.sub(r"^(?:Kabupaten|Kab\.?)\s*", "Kab. ", v, flags=re.I)) for r,v in regions]
            if len({v.casefold() for _,v in regions})!=35:
                issues.append((bab,ws.title,"Wilayah duplikat setelah normalisasi")); continue
            first=regions[0][0]; last=regions[-1][0]
            totals=[r for r in range(last+1,ws.max_row+1) if any(re.fullmatch(r'\s*(total|jumlah|jumlah\s*/\s*total|jawa tengah|provinsi jawa tengah)\s*',str(ws.cell(r,c).value or ''),re.I) for c in (1,2))]
            if not totals:
                issues.append((bab,ws.title,'Baris total belum dipetakan')); continue
            tr=totals[0]; title=str(ws['A2'].value or ws.title)
            merged={}
            for rg in ws.merged_cells.ranges:
                if rg.min_row>=4 and rg.min_row<first:
                    for r in range(rg.min_row,min(rg.max_row,first-1)+1):
                        for c in range(rg.min_col,rg.max_col+1): merged[r,c]=ws.cell(rg.min_row,rg.min_col).value
            count=0
            for c in range(3,ws.max_column+1):
                total=ws.cell(tr,c); vals=[ws.cell(r,c) for r,_ in regions]
                if not isinstance(total.value,(int,float)) and total.data_type!='e': continue
                labels=list(dict.fromkeys(str(v).strip() for r in range(5,first) if (v:=merged.get((r,c),ws.cell(r,c).value)) is not None and str(v).strip()))
                label=' / '.join(labels) or 'Kolom '+get_column_letter(c)+' (konfirmasi judul)'
                if str(bab)=='3' and ws.title in [f'{i}_Penma' for i in range(1,8)] and c in (3,4,5,6): label={3:'Jumlah lembaga',4:'Jumlah guru',5:'Jumlah siswa',6:'Jumlah rombel'}[c]
                parsed=[]
                for cell in vals:
                    v=cell.value; status='Numerik'
                    if cell.data_type=='e': v=None; status='Error Excel'
                    elif isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v): pass
                    elif isinstance(v,str) and re.fullmatch(r'[+-]?\d+(?:\.\d+)?',v.strip()): v=float(v); status='Teks numerik dikonversi'
                    else: status='Kosong' if v is None or v=='' else 'Nonnumerik'; v=None
                    parsed.append((v,status))
                if all(v is None for v,_ in parsed): continue
                count+=1
                for (r,region),cell,(v,status) in zip(regions,vals,parsed): rows.append([int(period.group()),str(bab),ws.title,title,get_column_letter(c),label,region,cell.coordinate,v,status])
            if not count: issues.append((bab,ws.title,'Indikator belum dipetakan'))
    finally: wb.close()
    if len({r[0] for r in rows})>1:
        for y in sorted({r[0] for r in rows}):
            for sheet in sorted({r[2] for r in rows if r[0]==y}):
                if sum(1 for r in rows if r[0]==y)<len(rows)/2: issues.append((bab,sheet,f"Tahun {y} berbeda dari mayoritas tabel dalam file; konfirmasi periode"))
    return pd.DataFrame(rows,columns=COLS),pd.DataFrame(issues,columns=['Bab','Sheet','Catatan'])
