from pathlib import Path
from tempfile import TemporaryDirectory
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import agregasi, analisis, clustering, validasi

def execute(task, files):
    with TemporaryDirectory() as tmp:
        root=Path(tmp); paths=[]
        for i,f in enumerate(files,1):
            p=root/f'{i}_Bab_{i}.xlsx'; p.write_bytes(f); paths.append(p)
        out=root/'output'; out.mkdir()
        if task=='validasi':
            frames=validasi.run(paths)
            for key,df in frames.items():
                df=df.copy()
                if 'File' in df: df['File']=df['File'].map(lambda x:Path(x).name)
                df.to_csv(out/f'{key}.csv',index=False,encoding='utf-8-sig')
        elif task=='agregasi':
            r,e,c,x=agregasi.audit(paths)
            headers={'Agregasi':['Bab','Sheet','Judul','Sel total','Metode','Excel','Python SUM','Selisih','Input nonnumerik','Status'],'Error':['Bab','Sheet','Sel','Nilai','Rumus'],'Cakupan':['Bab','Sheet','Jumlah wilayah','Baris total','Jumlah kandidat','Status'],'Analisis':['Bab','Sheet','Judul','Sel','Total','Minimum','Maksimum','Wilayah maksimum pertama']}
            for (name,cols),rows in zip(headers.items(),[r,e,c,x]):pd.DataFrame(rows,columns=cols).to_csv(out/f'{name}.csv',index=False,encoding='utf-8-sig')
        elif task=='analisis':
            analisis.analyse(paths,out); analisis.grafik_tambahan(out)
        elif task=='clustering': clustering.run(paths,out)
        return {str(p.relative_to(out)):p.read_bytes() for p in out.rglob('*') if p.is_file()}

