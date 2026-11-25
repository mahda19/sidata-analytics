import zipfile,xml.etree.ElementTree as E,re,math,collections,csv,os
N='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; Q=lambda s:'{'+N+'}'+s

def read(path):
 with zipfile.ZipFile(path) as z:
  ss=[ ''.join(t.text or '' for t in x.iter(Q('t'))) for x in E.fromstring(z.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in z.namelist() else []
  rel={x.get('Id'):x.get('Target') for x in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
  out={}
  for sh in E.fromstring(z.read('xl/workbook.xml')).find(Q('sheets')):
   target=rel[sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]; target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
   cells={}; shared={}
   for c in E.fromstring(z.read(target)).iter(Q('c')):
    v=c.findtext(Q('v')); t=c.get('t'); f=c.find(Q('f')); a=c.get('r')
    if t=='s' and v is not None:v=ss[int(v)]
    elif t=='inlineStr':v=''.join(x.text or '' for x in c.iter(Q('t')))
    elif t not in ('str','e','b') and v is not None:
     try:v=float(v)
     except ValueError:pass
    ft=f.text if f is not None else None
    if f is not None and f.get('t')=='shared':
     si=f.get('si')
     if ft:shared[si]=(a,ft)
     elif si in shared:
      anchor,base=shared[si]
      def coord(a):
       m=re.fullmatch(r'([A-Z]+)(\d+)',a); col=0
       for ch in m[1]:col=col*26+ord(ch)-64
       return col,int(m[2])
      def letters(c):
       out=''
       while c>0:c,r=divmod(c-1,26);out=chr(65+r)+out
       return out
      ac,ar=coord(anchor);bc,br=coord(a)
      def shift(m):
       c,r=coord(m[2]+m[4]);return m[1]+letters(c+(0 if m[1] else bc-ac))+m[3]+str(r+(0 if m[3] else br-ar))
      if re.fullmatch(r'SUM\(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+\)',base,re.I):ft=re.sub(r'(\$?)([A-Z]+)(\$?)(\d+)',shift,base)
    if v is not None or f is not None:cells[a]=(v,ft,t, f is not None)
   out[sh.get('name')]=cells
 return out


import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
import numpy as np
import shutil

def num(x): return isinstance(x,(int,float)) and math.isfinite(x)
def val(d,a): return d.get(a,(None,None,None,False))[0]
def colnum(c):
    x=0
    for k in c: x=x*26+ord(k)-64
    return x

def analyse(paths,out='Hasil_Analisis_Kemenag'):
    out=Path(out); out.mkdir(exist_ok=True); (out/'grafik').mkdir(exist_ok=True)
    stats=[]; long=[]; coverage=[]; errors=[]; strategic=[]; ratios=[]; conversions=[]
    for path in paths:
        bab=Path(path).name[0]; assert bab in '123', 'Nama file harus diawali 1, 2 atau 3'
        for sheet,d in read(path).items():
            if sheet.upper()=='REKAP': continue
            title=val(d,'A2'); period=val(d,'A3'); table=val(d,'A1')
            for a,(v,f,t,h) in d.items():
                if t=='e' or (f and '#REF!' in f): errors.append([bab,sheet,a,v,f])
            regions=sorted((int(a[1:]),v.strip()) for a,(v,*_) in d.items() if re.fullmatch(r'B\d+',a) and isinstance(v,str) and re.match(r'\s*(Kab\.\s*|Kab\s|Kabupaten\s|Kota\s)',v,re.I))
            if len(regions)!=35 or len(set(v.lower() for _,v in regions))!=35:
                coverage.append([bab,sheet,title,len(regions),0,'Perlu pemetaan wilayah']); continue
            ids=[r for r,_ in regions]; first=min(ids)
            # Only confirmed count columns in the seven education tables.
            if bab=='3' and sheet in [f'{i}_Penma' for i in range(1,8)]:
                for r,region in regions:
                    for c in 'CDEF':
                        a=c+str(r); raw,f,t,h=d.get(a,(None,None,None,False))
                        if isinstance(raw,str) and not f and re.fullmatch(r'[0-9]+',raw.strip()):
                            converted=int(raw.strip())
                            conversions.append([bab,sheet,a,region,raw,converted,'Teks angka dikonversi untuk analisis'])
                            d[a]=(converted,f,None,h)

            totals=sorted({int(a[1:]) for a,(v,*_) in d.items() if re.fullmatch(r'[AB]\d+',a) and isinstance(v,str) and re.fullmatch(r'\s*(total|jumlah|jumlah\s*/\s*total|jawa tengah|provinsi jawa tengah)\s*',v,re.I) and int(a[1:])>max(ids)})
            tr=totals[0] if totals else None
            if not tr:
                coverage.append([bab,sheet,title,35,0,'Baris total belum dipetakan']); continue
            # Restrict to columns with headers and numeric/formula total; no auxiliary notes or duplicate reference blocks unless explicitly mapped.
            candidates=sorted({m[1] for a,(v,f,t,h) in d.items() if (m:=re.fullmatch(r'([A-Z]+)(\d+)',a)) and int(m[2])==tr and colnum(m[1])>2 and (num(v) or h or t=='e')},key=colnum)
            count=0; extra=[int(a[1:]) for a,(v,*_) in d.items() if re.fullmatch(r'B\d+',a) and isinstance(v,str) and v.strip().lower()=='kanwil kemenag']
            for c in candidates:
                label=' / '.join(dict.fromkeys(str(val(d,c+str(r))).strip() for r in range(5,first) if isinstance(val(d,c+str(r)),str) and str(val(d,c+str(r))).strip())) or 'Kolom '+c+' (judul gabungan perlu konfirmasi)'
                values=[val(d,c+str(r)) for r in ids]; s=pd.Series([x if num(x) else np.nan for x in values],index=[v for _,v in regions],dtype=float); good=s.dropna()
                if good.empty: continue
                a=c+str(tr); v,f,t,h=d.get(a,(None,None,None,False)); included=sorted(ids+extra)
                additive=bool(f and f.replace('$','').upper()==f'SUM({c}{min(included)}:{c}{max(included)})' and included==list(range(min(included),max(included)+1)))
                # Do not interpret even a SUM as a meaningful total for explicitly non-additive indicators.
                nonadd=bool(re.search(r'persentase|percentage|rasio|ratio|rata.rata|%|indeks',label,re.I))
                eligible=additive and not nonadd and len(good)==35
                status='Lengkap numerik' if len(good)==35 else 'Input tidak lengkap/nonangka'
                stats.append(dict(Bab=bab,Sheet=sheet,Tabel=table,Judul=title,Periode=period,Kolom=c,Indikator=label,N=len(good),Tidak_numerik=35-len(good),Total_35_wilayah=good.sum() if eligible else np.nan,Total_Excel=v,Rumus_total=f,Metode='SUM eksplisit' if additive else 'Konfirmasi metode',Rata_rata=good.mean(),Median=good.median(),Minimum=good.min(),Maksimum=good.max(),SD_populasi=good.std(ddof=0),Wilayah_min='; '.join(good[good==good.min()].index),Wilayah_max='; '.join(good[good==good.max()].index),Status=status))
                for (r,region),x in zip(regions,values): long.append([bab,sheet,title,c,label,region,c+str(r),x,'Numerik' if num(x) else 'Periksa'])
                count+=1
            coverage.append([bab,sheet,title,35,count,'Dianalisis' if count else 'Perlu pemetaan indikator'])
            # Hand-mapped strategic indicators based on the actual 2026 table structure.
            mapping={('1','2_UP'): [('C','Kankemenag'),('D','MAN'),('E','MTsN'),('F','MIN')],('1','4_UP'):[('C','PPPK laki-laki'),('D','PPPK perempuan')],('2','2_BIMAS'):[('C','Islam'),('D','Gereja Kristen'),('E','Gereja Katolik'),('F','Pura'),('G','Vihara'),('H','Kelenteng')]}
            if bab=='3' and sheet in [f'{i}_Penma' for i in range(1,8)]: mapping[(bab,sheet)]=[('C','Lembaga'),('D','Guru'),('E','Siswa'),('F','Rombel')]
            if (bab,sheet) not in mapping: continue
            totals2={}; cols=mapping[(bab,sheet)]
            for c,label in cols:
                rows=ids+extra if (bab,sheet)==('1','4_UP') else ids
                vv=[val(d,c+str(r)) for r in rows]; total=sum(vv) if all(num(x) for x in vv) else np.nan; totals2[label]=total
                strategic.append([bab,sheet,title,label,total,len(rows),'Termasuk Kanwil' if len(rows)>35 else '35 kabupaten/kota'])
            if bab=='3':
                for label,a,b in [('Siswa per guru','Siswa','Guru'),('Siswa per rombel','Siswa','Rombel'),('Guru per lembaga','Guru','Lembaga'),('Siswa per lembaga','Siswa','Lembaga')]:
                    den=totals2[b]; ratios.append([sheet,title,label,totals2[a],den,totals2[a]/den if num(den) and den>0 else np.nan])
            else:
                ss=pd.Series(totals2).dropna()
                if len(ss):
                    ax=ss.sort_values().plot.barh(figsize=(9,4),title=str(title)); ax.set_xlabel('Jumlah'); ax.bar_label(ax.containers[0],fmt='{:,.0f}'); ax.figure.tight_layout(); ax.figure.savefig(out/'grafik'/f'{bab}_{sheet}.png',dpi=150); plt.close(ax.figure)
    frames={'Statistik_deskriptif':pd.DataFrame(stats),'Data_wilayah':pd.DataFrame(long,columns=['Bab','Sheet','Judul','Kolom','Indikator','Wilayah','Sel','Nilai','Status']),'Cakupan':pd.DataFrame(coverage,columns=['Bab','Sheet','Judul','Wilayah','Indikator','Status']),'Error':pd.DataFrame(errors,columns=['Bab','Sheet','Sel','Nilai','Rumus']),'Indikator_strategis':pd.DataFrame(strategic,columns=['Bab','Sheet','Judul','Indikator','Total','Unit','Cakupan']),'Rasio_pendidikan':pd.DataFrame(ratios,columns=['Sheet','Judul','Indikator','Pembilang','Penyebut','Rasio'])}
    frames['Log_cleansing']=pd.DataFrame(conversions,columns=['Bab','Sheet','Sel','Wilayah','Nilai_asli','Nilai_numerik','Tindakan'])
    st=frames['Indikator_strategis']; st['Persentase_komposisi']=np.nan
    for key in [('1','2_UP'),('1','4_UP'),('2','2_BIMAS')]:
        mask=(st.Bab==key[0])&(st.Sheet==key[1]); x=st.loc[mask,'Total'];
        if len(x) and x.notna().all() and x.sum()>0: st.loc[mask,'Persentase_komposisi']=100*x/x.sum()
    rr=frames['Rasio_pendidikan']; subset=rr[rr.Indikator.isin(['Siswa per guru','Siswa per rombel'])]
    if len(subset):
        labels={f'{i}_Penma':v for i,v in enumerate(['RA','MIN','MIS','MTsN','MTsS','MAN','MAS'],1)}
        plotdata=subset.pivot(index='Sheet',columns='Indikator',values='Rasio').reindex(list(labels)).rename(index=labels)
        ax=plotdata.plot.bar(figsize=(11,6),rot=0)
        ax.set_xlabel('Jenis satuan pendidikan'); ax.set_ylabel('Jumlah siswa per guru / rombel')
        ax.set_title('Rasio pendidikan madrasah – Jawa Tengah 2026')
        for bars in ax.containers:
            ax.bar_label(bars,labels=[f'{v:.2f}' if np.isfinite(v) else '' for v in bars.datavalues],padding=3,fontsize=9)
        for i,(_,row) in enumerate(plotdata.iterrows()):
            if row.isna().any(): ax.text(i,0,'Data belum lengkap',rotation=90,ha='center',va='bottom',fontsize=8)
        ax.margins(y=0.18); ax.figure.tight_layout(); ax.figure.savefig(out/'grafik'/'Rasio_pendidikan.png',dpi=170); plt.close(ax.figure)
    for name,df in frames.items(): df.to_csv(out/(name+'.csv'),index=False,encoding='utf-8-sig')
    return frames


def grafik_tambahan(OUT):
    # Grafik tambahan: data wilayah saja (tidak mencampurkan unit Kanwil).
    import pandas as pd
    import matplotlib.pyplot as plt
    from pathlib import Path

    G = OUT / 'grafik_tambahan'
    G.mkdir(exist_ok=True)
    s = pd.read_csv(OUT / 'Statistik_deskriptif.csv', dtype={'Kolom':str})
    w = pd.read_csv(OUT / 'Data_wilayah.csv', dtype={'Kolom':str})
    w['angka'] = pd.to_numeric(w['Nilai'], errors='coerce')
    ringkasan = []
    def grafik(series, judul, unit, nama, horizontal=True):
        series = series.dropna()
        if series.empty:
            print('Tidak ada data:', judul); return
        fig, ax = plt.subplots(figsize=(11, max(4, len(series)*.42+1.5) if horizontal else 5))
        bars = ax.barh(series.index, series.values) if horizontal else ax.bar(series.index, series.values)
        ax.set_title(judul, pad=15); ax.set_xlabel(unit) if horizontal else ax.set_ylabel(unit)
        if horizontal: ax.invert_yaxis(); ax.set_xlim(0,max(series.max()*1.22,1))
        else: ax.set_ylim(0,max(series.max()*1.22,1))
        ax.bar_label(bars, labels=[f'{v:,.0f}'.replace(',','.') for v in series], padding=4, fontsize=9)
        ax.grid(axis='x' if horizontal else 'y',alpha=.2); ax.set_axisbelow(True)
        fig.tight_layout(); fig.savefig(G / (nama+'.png'),dpi=170,bbox_inches='tight'); plt.close(fig)
        ringkasan.append({'Grafik':nama,'Indikator':judul,'Kategori_tertinggi':series.idxmax(),'Nilai_tertinggi':series.max(),'Kategori_terendah_ditampilkan':series.idxmin(),'Nilai_terendah_ditampilkan':series.min()})
    def wilayah(sheet,col):
        z=w[(w.Sheet==sheet)&(w.Kolom==col)]
        return z.set_index('Wilayah')['angka']
    def top(sheet,col,judul,unit,nama):
        z=wilayah(sheet,col)
        if len(z)!=35 or z.isna().any(): print('Cakupan tidak lengkap:',sheet,col); return
        grafik(z.nlargest(10),judul+' – 10 wilayah tertinggi',unit,nama)
        print('Minimum seluruh 35 wilayah:',z.idxmin(),z.min())
    top('2_UP','G','Jumlah satuan kerja','Satuan kerja','01_satker_wilayah')
    top('4_UP','E','Jumlah PPPK (tanpa unit Kanwil)','Orang','02_pppk_wilayah')
    z=s[(s.Sheet=='4_UP') & s.Kolom.isin(list('FGHIJKLMNOPQR')) & (s.N==35)]
    grafik(z.set_index('Indikator')['Total_35_wilayah'].sort_values(ascending=False),'Distribusi PPPK menurut golongan (35 wilayah)','Orang','03_pppk_golongan')
    top('2_BIMAS','I','Jumlah rumah ibadah','Rumah ibadah','04_rumah_ibadah_wilayah')
    # Gunakan kolom jumlah yang benar-benar berlabel jumlah, bukan menjumlahkan subtotal.
    penyuluh={}
    for sheet,label in [('4_BIMAS','PNS'),('5_BIMAS','PPPK'),('6_BIMAS','Non ASN')]:
        z=s[(s.Sheet==sheet)&(s.Indikator.str.strip().str.lower()=='jumlah')&(s.N==35)]
        if len(z)==1: penyuluh[label]=z.iloc[0]['Total_35_wilayah']
        else: print('Periksa kolom jumlah:',sheet)
    if len(penyuluh)==3: grafik(pd.Series(penyuluh),'Jumlah penyuluh agama menurut status','Orang','05_penyuluh_status',False)
    jenjang=dict(zip([f'{i}_Penma' for i in range(1,8)],['RA','MIN','MIS','MTsN','MTsS','MAN','MAS']))
    for col,label,unit,no in [('C','Lembaga pendidikan','Lembaga','06'),('D','Guru','Orang','07'),('E','Siswa','Orang','08'),('F','Rombongan belajar','Rombel','09')]:
        z=s[(s.Sheet.isin(jenjang))&(s.Kolom==col)&(s.N==35)].copy()
        z['Jenjang']=z.Sheet.map(jenjang)
        grafik(z.set_index('Jenjang')['Total_35_wilayah'].reindex(jenjang.values()),'Jumlah '+label.lower()+' menurut jenis pendidikan',unit,no+'_'+label.split()[0].lower(),False)
    for sheet,label in [('1_Penma','RA'),('3_Penma','MIS'),('5_Penma','MTsS'),('7_Penma','MAS')]:
        top(sheet,'E','Jumlah siswa '+label,'Orang','10_siswa_'+label)
    pd.DataFrame(ringkasan).to_csv(OUT/'Ringkasan_grafik_tambahan.csv',index=False,encoding='utf-8-sig')
    return pd.DataFrame(ringkasan)
