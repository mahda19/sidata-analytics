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

from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans, AgglomerativeClustering
from sklearn.metrics import silhouette_score, adjusted_rand_score, davies_bouldin_score
from sklearn.decomposition import PCA
from scipy.cluster.hierarchy import linkage, dendrogram
from threadpoolctl import threadpool_limits

import re, shutil

def display(*args): pass

def run(files, OUT):
    OUT=Path(OUT); OUT.mkdir(parents=True,exist_ok=True)
    books = [read(f) for f in files]
    log=[]
    def extract(bab, sheet, col):
        d=books[bab-1][sheet]; result={}
        for a, rec in d.items():
            v=rec[0]
            if not re.fullmatch(r'B\d+',a) or not isinstance(v,str) or not re.match(r'^\s*(Kab(?:\.\s*|\s+)|Kota\s+)',v,re.I): continue
            region=re.sub(r'^Kab\.?\s*','Kab. ',v.strip(),flags=re.I)
            addr=col+a[1:]; raw=d.get(addr,(None,None,None,False))[0]
            value=pd.to_numeric(raw,errors='coerce')
            if isinstance(raw,str) and pd.notna(value): log.append([bab,sheet,addr,raw,value])
            if region in result: raise ValueError(f'Duplikat {sheet}: {region}')
            result[region]=value
        s=pd.Series(result,dtype=float)
        if len(s)!=35: raise ValueError(f'{sheet}: {len(s)} wilayah, bukan 35')
        return s
    base=pd.DataFrame({'Kecamatan':extract(1,'1_UP','C'),'Satker':extract(1,'2_UP','G'),'PNS':extract(1,'3_UP','E'),'PPPK':extract(1,'4_UP','E'),'PNS_S2_S3':extract(1,'7_UP','H')+extract(1,'7_UP','I'),'Penduduk':extract(2,'1_BIMAS','J'),'Rumah_ibadah':extract(2,'2_BIMAS','I'),'Penyuluh_PNS':extract(2,'4_BIMAS','I'),'Penyuluh_PPPK':extract(2,'5_BIMAS','I')})
    for label,col in [('Lembaga','C'),('Guru','D'),('Siswa','E'),('Rombel','F')]:
        parts=pd.concat([extract(3,f'{i}_Penma',col) for i in range(1,8)],axis=1)
        base[label]=parts.sum(axis=1,min_count=7)
        if label=='Lembaga': base['Negeri']=parts.iloc[:,[1,3,5]].sum(axis=1,min_count=3)
    assert len(base)==35
    if base.isna().any().any() or (base<0).any().any():
        print(base[base.isna().any(axis=1)|(base<0).any(axis=1)])
        raise ValueError('Data kosong/error/negatif: periksa sumber, tanpa imputasi.')
    assert (base[['Kecamatan','Penduduk','PNS','Satker','Guru','Rombel','Lembaga']]>0).all().all()
    a=base
    asn=a.Penyuluh_PNS+a.Penyuluh_PPPK
    models_X={
     '01_Tata_Kelola':pd.DataFrame({'Satker/kecamatan':a.Satker/a.Kecamatan,'PNS/satker':a.PNS/a.Satker,'PPPK/satker':a.PPPK/a.Satker,'PNS S2-S3 (%)':100*a.PNS_S2_S3/a.PNS}),
     '02_Pelayanan':pd.DataFrame({'Rumah ibadah/10000 penduduk':1e4*a.Rumah_ibadah/a.Penduduk,'Penyuluh ASN/10000 penduduk':1e4*asn/a.Penduduk,'Penyuluh PPPK (%)':100*a.Penyuluh_PPPK/asn.replace(0,np.nan),'Rumah ibadah/kecamatan':a.Rumah_ibadah/a.Kecamatan}),
     '03_Pendidikan':pd.DataFrame({'Siswa/guru':a.Siswa/a.Guru,'Siswa/rombel':a.Siswa/a.Rombel,'Guru/lembaga':a.Guru/a.Lembaga,'Lembaga negeri (%)':100*a.Negeri/a.Lembaga}),
     '04_Gabungan':pd.DataFrame({'Satker/kecamatan':a.Satker/a.Kecamatan,'PPPK/kecamatan':a.PPPK/a.Kecamatan,'Rumah ibadah/kecamatan':a.Rumah_ibadah/a.Kecamatan,'Penyuluh ASN/kecamatan':asn/a.Kecamatan,'Siswa/guru':a.Siswa/a.Guru,'Siswa/rombel':a.Siswa/a.Rombel})}
    base.to_csv(OUT/'Data_dasar.csv'); pd.DataFrame(log,columns=['Bab','Sheet','Sel','Teks','Angka']).to_csv(OUT/'Log_konversi.csv',index=False)

    summary=[]
    ROOT=OUT
    for model_name,X in models_X.items():
        print('\n',model_name)
        OUT=ROOT/model_name; OUT.mkdir(exist_ok=True)
        if not np.isfinite(X.to_numpy()).all(): raise ValueError(f'{model_name}: indikator tidak valid')
        X.to_csv(OUT/'Indikator_awal.csv')
        corr=X.corr(method='spearman'); corr.to_csv(OUT/'Korelasi_spearman.csv')
        kept=[]; removed=[]
        for col in X.columns:
            reason='Konstan' if X[col].nunique()<2 else next((f'Korelasi dengan {j}' for j in kept if abs(corr.loc[col,j])>=.85 and model_name!='04_Gabungan'),None)
            if reason: removed.append([col,reason])
            else: kept.append(col)
        pd.DataFrame(removed,columns=['Indikator','Alasan']).to_csv(OUT/'Seleksi_fitur.csv',index=False)
        X=X[kept]; X.to_csv(OUT/'Indikator_clustering.csv')
        assert X.shape[1]>=2
        Z=StandardScaler().fit_transform(X)
        results=[]; models={}
        with threadpool_limits(limits=1):
            for k in range(2,7):
                model=KMeans(n_clusters=k,n_init=50,random_state=42).fit(Z)
                y=model.labels_; sizes=np.bincount(y)
                results.append([k,silhouette_score(Z,y),davies_bouldin_score(Z,y),model.inertia_,sizes.min()]); models[k]=model
        scores=pd.DataFrame(results,columns=['k','Silhouette','Davies_Bouldin','Inertia','Min_anggota'])
        eligible=scores[scores.Min_anggota>=2]
        if eligible.empty: raise ValueError('Tidak ada kandidat tanpa cluster tunggal')
        k=int(eligible.sort_values(['Silhouette','k'],ascending=[False,True]).iloc[0].k)
        y=models[k].labels_+1
        ward=AgglomerativeClustering(n_clusters=k,linkage='ward').fit_predict(Z)+1
        ari=adjusted_rand_score(y,ward)
        # Uji sensitivitas: keluarkan satu fitur pada setiap iterasi, k tetap.
        sens=[]
        with threadpool_limits(limits=1):
            for j,name in enumerate(X.columns):
                alt=KMeans(n_clusters=k,n_init=50,random_state=42).fit_predict(np.delete(Z,j,axis=1))
                sens.append([name,adjusted_rand_score(y,alt)])
        sensitivity=pd.DataFrame(sens,columns=['Fitur_dikeluarkan','ARI'])
        assign=X.assign(Cluster=y,Ward=ward); profile=assign.groupby('Cluster')[X.columns].mean(); counts=assign.Cluster.value_counts().sort_index()
        profile.insert(0,'Jumlah wilayah',counts)
        standard=pd.DataFrame(Z,index=X.index,columns=X.columns).assign(Cluster=y).groupby('Cluster').mean()
        for name,df in [('Evaluasi_k',scores),('Anggota_cluster',assign),('Profil_cluster',profile),('Profil_zscore',standard),('Sensitivitas',sensitivity)]: df.to_csv(OUT/(name+'.csv'),index=name!='Evaluasi_k' and name!='Sensitivitas')
        print('k terpilih:',k,'; ARI K-Means vs Ward:',round(ari,3)); display(scores.round(3)); display(profile.round(2)); display(sensitivity.round(3))
        for group in sorted(set(y)): print('Cluster',group,':',', '.join(X.index[y==group]))
        def finish(fig,name):
            fig.tight_layout(); fig.savefig(OUT/(name+'.png'),dpi=180,bbox_inches='tight'); plt.close(fig)
        fig,ax=plt.subplots(figsize=(7,4)); ax.plot(scores.k,scores.Silhouette,'o-'); ax.axvline(k,linestyle='--'); ax.set(xlabel='Jumlah cluster',ylabel='Silhouette score',title='Pemilihan jumlah cluster'); finish(fig,'01_silhouette')
        pca=PCA(n_components=2); pc=pca.fit_transform(Z)
        fig,ax=plt.subplots(figsize=(12,8))
        for group in sorted(set(y)):
            mask=y==group; ax.scatter(pc[mask,0],pc[mask,1],label=f'Cluster {group}',s=50)
        for i,region in enumerate(X.index): ax.annotate(region,(pc[i,0],pc[i,1]),xytext=(3,4),textcoords='offset points',fontsize=7)
        ax.set(xlabel=f'PC1 ({pca.explained_variance_ratio_[0]:.1%})',ylabel=f'PC2 ({pca.explained_variance_ratio_[1]:.1%})',title='PCA – pengelompokan kabupaten/kota'); ax.legend(); finish(fig,'02_pca')
        fig,ax=plt.subplots(figsize=(10,4)); im=ax.imshow(standard.values,aspect='auto',cmap='RdBu_r',vmin=-max(1,abs(standard.values).max()),vmax=max(1,abs(standard.values).max())); ax.set_xticks(range(len(X.columns)),X.columns,rotation=25,ha='right'); ax.set_yticks(range(k),[f'Cluster {i}' for i in standard.index]);
        for i in range(k):
            for j in range(len(X.columns)): ax.text(j,i,f'{standard.iloc[i,j]:.2f}',ha='center',va='center',fontsize=9)
        fig.colorbar(im,ax=ax,label='Rata-rata z-score'); finish(fig,'03_profil')
        fig,ax=plt.subplots(figsize=(10,10)); dendrogram(linkage(Z,method='ward'),labels=X.index.tolist(),orientation='right',ax=ax,leaf_font_size=8); ax.set_title('Hierarchical clustering – Ward'); finish(fig,'04_dendrogram')
        fig,ax=plt.subplots(figsize=(7,4)); bars=ax.bar(counts.index.astype(str),counts.values); ax.bar_label(bars); ax.set(xlabel='Cluster',ylabel='Jumlah wilayah',title='Jumlah anggota cluster'); finish(fig,'05_anggota')
        print('PCA menjelaskan',round(sum(pca.explained_variance_ratio_)*100,2),'persen variasi.')
        fig,ax=plt.subplots(figsize=(7,4)); ax.plot(scores.k,scores.Inertia,'o-'); ax.axvline(k,linestyle='--'); ax.set(xlabel='Jumlah cluster',ylabel='Inertia',title='Elbow method'); finish(fig,'06_elbow')
        summary.append(dict(Model=model_name,Wilayah=len(X),Fitur=len(kept),K=k,Silhouette=float(scores.loc[scores.k==k,'Silhouette'].iloc[0]),DBI=float(scores.loc[scores.k==k,'Davies_Bouldin'].iloc[0]),ARI_Ward=ari,PCA_2=sum(pca.explained_variance_ratio_),Min_ARI=sensitivity.ARI.min(),Anggota=' / '.join(map(str,counts.values))))
    OUT=ROOT
    summary=pd.DataFrame(summary); summary.to_csv(OUT/'Ringkasan_model.csv',index=False); display(summary.round(3))
    return summary
