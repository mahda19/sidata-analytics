import hashlib, zipfile
from io import BytesIO
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px
from explorer import extract

st.set_page_config(page_title='SIDATA Analytics | Kemenag Jawa Tengah',page_icon='📊',layout='wide')
st.markdown('''<style>
.block-container{padding-top:2rem;max-width:1400px} [data-testid="stSidebar"]{background:#e9f1ed} [data-testid="stMetric"]{border:1px solid #d6e3dc;border-radius:12px;padding:15px;background:white} [data-testid="stMetricValue"]{color:#12634e} .hero{background:linear-gradient(110deg,#104c3e,#23785f);padding:25px 30px;border-radius:16px;color:white;margin-bottom:20px}.hero h1{color:white;font-size:32px;margin:0;padding:0}.hero p{margin:8px 0 0;color:#e0f2e9} </style>''',unsafe_allow_html=True)
st.markdown('<div class="hero"><h1>SIDATA Analytics</h1><p>Statistik Sektoral • Kanwil Kementerian Agama Provinsi Jawa Tengah</p></div>',unsafe_allow_html=True)
NAMES={'1':'Tata Kelola dan Dukungan Manajemen','2':'Pelayanan Keagamaan','3':'Pendidikan Agama dan Pendidikan Keagamaan'}
st.sidebar.header('Sumber data')
import json
from google_data import export_url, download_excel, WorkbookSource
st.sidebar.caption('Data publik yang telah disetujui saja. Tahun dibaca dari judul tabel.')
config=json.loads(Path(__file__).with_name('sources.json').read_text(encoding='utf-8'))
mode=st.sidebar.radio('Pilih sumber data',['Google Sheets / Drive otomatis','Upload Excel manual'])
if mode=='Google Sheets / Drive otomatis':
    uploads={b:[] for b in NAMES}
    configured=any(config.get(b) for b in NAMES)
    if st.sidebar.button('Perbarui data Google'):
        download_excel.clear()
    for b,n in NAMES.items():
        links=config.get(b,[])
        if not configured:
            raw=st.sidebar.text_area('Link Bab '+b,placeholder='https://docs.google.com/spreadsheets/d/...',key='google'+b,help='Satu link per baris. Gunakan file sumber yang telah disetujui untuk publikasi.')
            links=raw.splitlines()
        for i,link in enumerate(links,1):
            if not link.strip(): continue
            try:
                with st.spinner('Mengambil data Bab '+b+' dari Google…'):
                    blob=download_excel(export_url(link))
                uploads[b].append(WorkbookSource(f'{b}_Google_{i}.xlsx',blob))
            except Exception as e:
                st.error(f'Bab {b}, sumber {i}: {e}')
        if uploads[b]: st.sidebar.success(f'Bab {b}: {len(uploads[b])} file tersedia')
    if not configured: st.sidebar.info('Isi link sumber untuk mencoba. Untuk memuat otomatis bagi semua pengunjung, isi sources.json di GitHub.')
else:
    uploads={b:st.sidebar.file_uploader(n,type='xlsx',accept_multiple_files=True,key='upload'+b) for b,n in NAMES.items()}
menu=st.sidebar.radio('Menu utama',['Dashboard','Analisis statistik','Pemeriksaan & validasi','Agregasi','Grafik tambahan','Clustering 4 model'])
st.sidebar.caption('Data publik saja • Jangan unggah data internal ke aplikasi publik.')
# No global upload cache for public deployment: results live in this session.
sig=tuple((b,f.name,hashlib.sha256(f.getvalue()).hexdigest()) for b,fs in uploads.items() for f in fs)
if st.session_state.get('signature')!=sig:
    frames=[]; problems=[]; sources={}; duplicates=[]
    with st.spinner('Membaca tabel Excel…'):
        for b,fs in uploads.items():
            for f in fs:
                try:
                    df,issues=extract(f.getvalue(),b)
                    problems.append(issues)
                    for y,part in df.groupby('Tahun'):
                        key=(int(y),b)
                        if key in sources: duplicates.append(f'{y} / Bab {b}: {f.name} diabaikan (satu versi per tahun dan bab)'); continue
                        sources[key]=f.getvalue(); frames.append(part)
                except Exception as e: st.error(f'{f.name}: {e}')
    st.session_state.update(signature=sig, dataset=pd.concat(frames,ignore_index=True) if frames else pd.DataFrame(),issues=pd.concat(problems,ignore_index=True) if problems else pd.DataFrame(),sources=sources,duplicates=duplicates,results={})
df=st.session_state.dataset
for msg in st.session_state.duplicates: st.warning(msg)
if not st.session_state.issues.empty:
    mixed=st.session_state.issues[st.session_state.issues.Catatan.str.contains("berbeda dari mayoritas",na=False)]
    for _,r in mixed.iterrows(): st.warning(f"Bab {r.Bab} / {r.Sheet}: {r.Catatan}")
if df.empty:
    st.info('Isi link Google atau unggah file Excel pada panel kiri. Filter dan grafik akan muncul setelah data terbaca.')
    st.markdown('### Filter Analisis Statistik')
    a,b,c,d=st.columns(4)
    a.selectbox('Tahun data',['Belum ada data'],disabled=True); b.selectbox('Kelompok data',list(NAMES.values()),disabled=True); c.selectbox('Kabupaten/kota',['Seluruh wilayah'],disabled=True); d.selectbox('Indikator',['Unggah data dahulu'],disabled=True)
    st.radio('Jenis visualisasi',['Bar chart','Pie chart','Boxplot','Scatter plot'],horizontal=True,disabled=True)
    st.stop()
years=sorted(df.Tahun.unique(),reverse=True)
if menu=='Dashboard':
    year=st.selectbox('Tahun data',years); d=df[df.Tahun==year]
    a,b,c,e=st.columns(4); a.metric('Kabupaten/kota',d.Wilayah.nunique()); b.metric('Tabel terbaca',len(d[['Bab','Sheet']].drop_duplicates())); c.metric('Indikator',len(d[['Bab','Sheet','Kolom']].drop_duplicates())); e.metric('Nilai numerik',f'{d.Nilai.notna().mean():.1%}')
    st.plotly_chart(px.bar(d[['Bab','Sheet']].drop_duplicates().groupby('Bab').size().rename('Jumlah tabel').reset_index().assign(Kelompok=lambda x:x.Bab.map(NAMES)),x='Kelompok',y='Jumlah tabel',color='Kelompok',title='Cakupan tabel yang terbaca'),use_container_width=True)
    st.caption('Cakupan tabel terbaca bukan ukuran akurasi data. Hasil hanya mencakup 35 wilayah yang dikenali.')
    if not st.session_state.issues.empty:
        with st.expander('Tabel yang memerlukan pemetaan'): st.dataframe(st.session_state.issues,use_container_width=True)
    st.info('Pilih Analisis statistik pada menu kiri untuk menggunakan seluruh filter dan grafik interaktif.')
elif menu=='Analisis statistik':
    st.subheader('Filter Analisis Statistik')
    with st.container(border=True):
        a,b=st.columns(2); year=a.selectbox('Tahun data',years); dy=df[df.Tahun==year]; bab=b.selectbox('Kelompok data',sorted(dy.Bab.unique()),format_func=NAMES.get); db=dy[dy.Bab==bab]
        a,b=st.columns(2); region=a.multiselect('Kabupaten/kota',sorted(db.Wilayah.unique()),placeholder='Seluruh wilayah'); table=db[['Sheet','Judul']].drop_duplicates(); sheet=b.selectbox('Tabel / Sheet',table.Sheet.tolist(),format_func=lambda x:f'{x} — '+str(table.loc[table.Sheet==x,'Judul'].iloc[0])); ds=db[db.Sheet==sheet]; labels=ds[['Kolom','Indikator']].drop_duplicates().set_index('Kolom').Indikator.to_dict()
        a,b=st.columns(2); col=a.selectbox('Indikator',list(labels),format_func=lambda x:f'{labels[x]} [{x}]'); status=b.selectbox('Status data',['Semua','Numerik','Teks numerik dikonversi','Kosong','Nonnumerik','Error Excel']); kind=st.radio('Jenis visualisasi',['Bar chart','Pie chart','Boxplot','Scatter plot'],horizontal=True)
        if kind=='Scatter plot': xcol=st.selectbox('Indikator sumbu X',[x for x in labels if x!=col] or [col],format_func=lambda x:f'{labels[x]} [{x}]')
    view=ds[ds.Kolom==col].copy()
    if region: view=view[view.Wilayah.isin(region)]
    if status!='Semua': view=view[view.Status==status]
    good=view.dropna(subset=['Nilai']).copy()
    if not good.empty:
        lo,hi=float(good.Nilai.min()),float(good.Nilai.max())
        if lo<hi and st.checkbox('Aktifkan filter rentang nilai'):
            bounds=st.slider('Rentang nilai',lo,hi,(lo,hi)); view=view[view.Nilai.between(*bounds)]; good=view.dropna(subset=['Nilai'])
    order=st.selectbox('Urutkan wilayah',['Tertinggi','Terendah','Nama wilayah']); good=good.sort_values('Wilayah' if order=='Nama wilayah' else 'Nilai',ascending=order!='Tertinggi')
    a,b,c,d=st.columns(4); a.metric('Wilayah terpilih',len(view)); b.metric('Nilai tersedia',len(good)); c.metric('Rata-rata wilayah',f'{good.Nilai.mean():,.2f}' if len(good) else '—'); d.metric('Median',f'{good.Nilai.median():,.2f}' if len(good) else '—')
    st.subheader(labels[col]); st.caption(f'{year} • {NAMES[bab]} • {sheet} | {len(view)} wilayah')
    fig=None
    if good.empty: st.warning('Tidak ada nilai numerik sesuai filter.')
    elif kind=='Bar chart': fig=px.bar(good,x='Wilayah',y='Nilai',hover_data=['Sel','Status'],color_discrete_sequence=['#24785e'])
    elif kind=='Pie chart':
        if (good.Nilai<0).any() or good.Nilai.sum()<=0: st.warning('Pie chart memerlukan nilai nonnegatif dengan total positif.')
        elif not st.checkbox('Saya telah memastikan indikator berupa jumlah yang dapat dijumlahkan antarwilayah'): st.info('Konfirmasi sifat aditif indikator sebelum menampilkan komposisi. Persentase dan rasio tidak boleh dijumlahkan.')
        else: fig=px.pie(good,names='Wilayah',values='Nilai',hole=.42)
    elif kind=='Boxplot':
        if len(good)<2: st.warning('Pilih minimal dua wilayah untuk melihat distribusi.')
        else: fig=px.box(good,y='Nilai',points='all',hover_data=['Wilayah'])
    else:
        pair=good[['Wilayah','Nilai']].merge(ds[ds.Kolom==xcol][['Wilayah','Nilai']].rename(columns={'Nilai':'X'}),on='Wilayah').dropna()
        if len(pair)<2 or xcol==col: st.warning('Pilih dua indikator berbeda dan minimal dua pasangan nilai.')
        else: fig=px.scatter(pair,x='X',y='Nilai',hover_name='Wilayah',labels={'X':labels[xcol],'Nilai':labels[col]})
    if fig is not None:
        fig.update_layout(template='plotly_white',height=480,margin=dict(l=20,r=20,t=30,b=20)); st.plotly_chart(fig,use_container_width=True); st.download_button('Unduh grafik interaktif (HTML)',fig.to_html(include_plotlyjs=True).encode(),'grafik.html','text/html')
    st.dataframe(view,use_container_width=True,hide_index=True); st.download_button('Unduh data sesuai filter (CSV)',view.to_csv(index=False).encode('utf-8-sig'),'data_filter.csv','text/csv')
    if len(good): st.dataframe(good.Nilai.describe().rename('Nilai').to_frame(),use_container_width=True)
    st.caption('Rata-rata dan median merupakan ringkasan antarwilayah, bukan nilai agregat provinsi. Nilai kosong/error tidak diganti nol. Tahun dan kelompok mengikuti file sumber.')
else:
    st.subheader(menu); task={'Pemeriksaan & validasi':'validasi','Agregasi':'agregasi','Grafik tambahan':'analisis','Clustering 4 model':'clustering'}[menu]
    year=st.selectbox('Tahun data',years)
    if year!=2026 or any((int(year),b) not in st.session_state.sources for b in NAMES): st.info('Modul lama dipetakan untuk tiga file tahun 2026. Sediakan ketiga bab 2026 untuk menjalankannya.'); st.stop()
    if st.button('Jalankan / ulangi analisis',type='primary'):
        try:
            from engine import execute
            with st.spinner('Menjalankan modul analisis…'): st.session_state.results[task]=execute(task,[st.session_state.sources[(int(year),b)] for b in NAMES])
        except Exception as e: st.exception(e)
    data=st.session_state.results.get(task)
    if not data: st.info('Klik tombol untuk menjalankan modul.'); st.stop()
    keys=sorted(data)
    if task=='clustering':
        model=st.selectbox('Model',['Semua','01_Tata_Kelola','02_Pelayanan','03_Pendidikan','04_Gabungan']); keys=[k for k in keys if model=='Semua' or k.startswith(model+'/')]; st.caption('Cluster eksploratif, bukan peringkat kinerja.')
    if task=='validasi': st.caption('Simulasi Python bukan uji penolakan input Excel langsung. Rumus yang tidak didukung memerlukan pemeriksaan manual.')
    kind=st.radio('Jenis hasil',['Tabel','Grafik'],horizontal=True); keys=[k for k in keys if k.endswith('.csv' if kind=='Tabel' else '.png')]
    if keys:
        key=st.selectbox('Pilih hasil',keys)
        if kind=='Grafik': st.image(data[key]); st.download_button('Unduh PNG',data[key],Path(key).name,'image/png')
        else:
            try: st.dataframe(pd.read_csv(BytesIO(data[key])),use_container_width=True)
            except pd.errors.EmptyDataError: st.info('Tidak ada baris hasil.')
            st.download_button('Unduh CSV',data[key],Path(key).name,'text/csv')
    else: st.info('Tidak tersedia hasil jenis ini.')
    buf=BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
        for k,v in data.items(): z.writestr(k,v)
    st.download_button('Unduh semua hasil ZIP',buf.getvalue(),f'SIDATA_{task}.zip','application/zip')
