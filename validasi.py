import openpyxl as xl
import pandas as pd
import re, math, operator
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.formula.translate import Translator

OPS = {'=':operator.eq,'<>':operator.ne,'<':operator.lt,'>':operator.gt,'<=':operator.le,'>=':operator.ge}
def num(v): return isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v)
def blank(v): return v is None or v == ''
def coords(r):
    a,b,c,d=range_boundaries(str(r))
    if (c-a+1)*(d-b+1)>100000: raise ValueError('Rentang rumus terlalu luas untuk evaluasi otomatis')
    return [f'{get_column_letter(j)}{i}' for i in range(b,d+1) for j in range(a,c+1)]
def get(w,s,c,overlay): return overlay.get((s,c),w[s][c].value)
def decimal(d,v):
    if blank(v): return None
    if not num(v) or (d.type=='whole' and v%1): return False
    a=float(d.formula1); op=d.operator or 'between'
    if op in ('between','notBetween'):
        ok=a<=v<=float(d.formula2); return ok if op=='between' else not ok
    return {'equal':operator.eq,'notEqual':operator.ne,'lessThan':operator.lt,'lessThanOrEqual':operator.le,'greaterThan':operator.gt,'greaterThanOrEqual':operator.ge}[op](v,a)
REF=r"(?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?(\$?[A-Z]+\$?\d+)"
def ref(s,t):
    m=re.fullmatch(REF,t,re.I)
    if not m: raise ValueError('Referensi tidak didukung')
    return (m[1] or m[2] or s,m[3].replace('$','').upper())
def parse_custom(s,f):
    m=re.fullmatch(r'(.+?)(<=|>=|<>|=|<|>)(.+)',f.strip().lstrip('='),re.I)
    if not m: raise ValueError('Custom tidak didukung')
    left,op,right=m.groups(); right=ref(s,right)
    sm=re.fullmatch(r'SUM\(([^)]+)\)',left,re.I)
    if sm: keys=[(s,c) for c in coords(sm[1].replace('$',''))]; kind='sum'
    else: keys=[ref(s,left)]; kind='ref'
    return keys,op,right,kind
def custom(w,s,f,o):
    keys,op,right,kind=parse_custom(s,f)
    vals=[get(w,*k,o) for k in keys]; b=get(w,*right,o)
    if not num(b) or any(not num(v) for v in vals): return None
    return OPS[op](sum(vals) if kind=='sum' else vals[0],b)
def status(vals):
    k=sum(blank(v) for v in vals)
    return 'Belum mengisi' if k==len(vals) else 'Terisi penuh' if k==0 else 'Terisi sebagian'
def rekap_range(f):
    m=re.search(r"COUNTBLANK\('([^']+)'!([A-Z]+\d+:[A-Z]+\d+)\)",f or '',re.I)
    return (m[1],coords(m[2])) if m else None
def error(f,c): return c.data_type=='e' or (isinstance(f,str) and f.startswith('=') and '#REF!' in f)

def run(paths):
    temuan=[]; aturan=[]; uji=[]; ringkas=[]
    def record(p,s,c,jenis,detail): temuan.append(dict(File=p,Sheet=s,Sel=c,Jenis=jenis,Detail=str(detail)))
    def test(p,id,s,c,expected,actual,detail):
        uji.append(dict(File=p,Skenario=id,Sheet=s,Sel=c,Harapan=expected,Hasil=actual,Status='LULUS' if actual==expected else 'GAGAL',Detail=detail))
    def skip(p,id,detail): uji.append(dict(File=p,Skenario=id,Status='TIDAK DIUJI',Detail=detail))
    for p in paths:
        f=xl.load_workbook(p,data_only=False); w=xl.load_workbook(p,data_only=True)
        candidates={}; rekap=[]; checked=matched=0
        for s in f:
            for row in (tuple(s._cells.values()),):
                for c in row:
                    if c.value is None: continue
                    v=w[s.title][c.coordinate]
                    if error(c.value,v): record(p,s.title,c.coordinate,'Error rumus',c.value)
                    if s.title=='REKAP' and c.data_type=='f':
                        r=rekap_range(c.value)
                        if r:
                            sh,cs=r
                            if sh not in w.sheetnames:
                                record(p,s.title,c.coordinate,'Referensi sheet tidak tersedia',sh); continue
                            expected=status([w[sh][x].value for x in cs]); checked+=1
                            if expected==v.value: matched+=1
                            else: record(p,s.title,c.coordinate,'Status REKAP berbeda',f'{v.value} vs {expected}')
                            rekap.append((c.coordinate,sh,cs))
            for d in s.data_validations.dataValidation:
                rule=dict(File=p,Sheet=s.title,Jenis=d.type,Rentang=str(d.sqref),Rumus=d.formula1,Status='Didukung')
                try:
                    # Anchor paling atas-kiri sesuai urutan koordinat; aturan multi-range perlu konfirmasi di Excel.
                    anchor=min((r.min_row,r.min_col) for r in d.sqref.ranges); anchor=f'{get_column_letter(anchor[1])}{anchor[0]}'
                    if d.type in ('decimal','whole'):
                        decimal(d,0); candidates.setdefault('numeric',(s.title,anchor,d))
                    elif d.type=='custom':
                        keys,op,right,kind=parse_custom(s.title,d.formula1)
                        if op not in ('=','<=','>=','<','>','<>'): raise ValueError()
                        if right not in keys and op in ('=','<='):
                            candidates.setdefault('cross' if right[0]!=s.title else 'local',(s.title,anchor,d))
                    else: raise ValueError()
                    for rng in d.sqref.ranges:
                        for cell in tuple(w[s.title]._cells.values()):
                            if not (rng.min_row<=cell.row<=rng.max_row and rng.min_col<=cell.column<=rng.max_col): continue
                            addr=cell.coordinate
                            val=cell.value
                            if blank(val): continue
                            if d.type in ('decimal','whole'): ok=decimal(d,val)
                            else:
                                expr=Translator('='+d.formula1.lstrip('='),origin=anchor).translate_formula(addr)
                                ok=custom(w,s.title,expr,{})
                            if ok is False: record(p,s.title,addr,'Pelanggaran validasi',f'{val} | {d.formula1}')
                            elif ok is None: record(p,s.title,addr,'Validasi belum dapat dievaluasi',d.formula1)
                except Exception as e: rule['Status']='Perlu pemeriksaan manual: '+str(e)
                aturan.append(rule)
            for cf in s.conditional_formatting:
                for rr in s.conditional_formatting[cf]:
                    aturan.append(dict(File=p,Sheet=s.title,Jenis='Conditional Formatting',Rentang=str(cf.sqref),Rumus=str(rr.formula),Status='Inventarisasi saja'))
        ringkas.append(dict(File=p,Status_REKAP_diperiksa=checked,Status_REKAP_sesuai=matched))
        # Uji 1,2,7,9: aturan numerik yang benar-benar ditemukan
        if 'numeric' in candidates:
            s,c,d=candidates['numeric']; a=float(d.formula1)
            if (d.operator or 'between')=='between' and a>=0:
                b=float(d.formula2)
                for id,val,exp in [('U1 negatif',a-1,True),('U2 teks','ABC',True),('U7 di atas batas',b+1,True),('U9 kontrol valid',a,False)]:
                    test(p,id,s,c,exp,decimal(d,val) is False,f'Nilai simulasi {val}; batas {a}–{b}')
            else:
                for id in ['U1 negatif','U2 teks','U7 di atas batas','U9 kontrol valid']: skip(p,id,'Aturan rentang nonnegatif tidak ditemukan pada kandidat pertama')
        else:
            for id in ['U1 negatif','U2 teks','U7 di atas batas','U9 kontrol valid']: skip(p,id,'Tidak ada aturan numerik konstan')
        # U3: rumus rusak dibuat pada objek memori, bukan file asli
        s=next(x for x in f if x.title!='REKAP'); c=s['C7']; original=c.value
        c.value='=#REF!+1'; test(p,'U3 referensi rusak',s.title,'C7',True,error(c.value,c),'Simulasi formula literal #REF!'); c.value=original
        # U5/U6: uji pelanggaran aturan Custom yang didukung, dengan nilai kontrol eksplisit
        for key,id in [('local','U5 konsistensi lokal'),('cross','U6 konsistensi antartabel')]:
            if key not in candidates: skip(p,id,'Tidak ditemukan aturan Custom langsung yang didukung'); continue
            s,c,d=candidates[key]; keys,op,right,kind=parse_custom(s,d.formula1)
            o={k:0 for k in keys}; o[right]=0; o[keys[0]]=10
            test(p,id,s,c,True,custom(w,s,d.formula1,o) is False,f'{d.formula1}; sisi kiri 10, pembanding 0')
            o[keys[0]]=0
            test(p,'U9 kontrol '+key,s,c,False,custom(w,s,d.formula1,o) is False,'Kedua sisi 0')
        # U4 dan U8: simulasi isi kosong serta status REKAP yang tidak sesuai
        if rekap:
            c,s,cs=rekap[0]
            for id,vals,exp in [('U4 kosong total',[None]*len(cs),'Belum mengisi'),('U4 kosong sebagian',[None]+[1]*(len(cs)-1),'Terisi sebagian' if len(cs)>1 else 'Belum mengisi'),('U9 kontrol REKAP',[1]*len(cs),'Terisi penuh')]:
                test(p,id,s,','.join(cs),exp,status(vals),'Simulasi rentang COUNTBLANK aktual')
            test(p,'U8 REKAP tidak sesuai','REKAP',c,True,status([None]*len(cs))!='Terisi penuh','Status simulasi tersimpan: Terisi penuh; seluruh sumber kosong')
        else:
            skip(p,'U4 kosong','Tidak ada pola REKAP yang didukung'); skip(p,'U8 REKAP','Tidak ada pola REKAP yang didukung')

    du=pd.DataFrame(uji); dt=pd.DataFrame(temuan,columns=['File','Sheet','Sel','Jenis','Detail']); da=pd.DataFrame(aturan); dr=pd.DataFrame(ringkas)
    rows=[]
    for p,g in du.groupby('File'):
        for label,h in [('Deteksi kesalahan',g[~g.Skenario.str.startswith('U9')]),('Kontrol valid',g[g.Skenario.str.startswith('U9')])]:
            run=h[h.Status!='TIDAK DIUJI']; passed=(run.Status=='LULUS').sum()
            rows.append({'File':p,'Indikator':label,'Diuji':len(run),'Lulus':passed,'Tidak diuji':len(h)-len(run),'Persentase':100*passed/len(run) if len(run) else None})
    df=pd.DataFrame(rows); 

    return {"Efektivitas":df,"Skenario":du,"Temuan":dt,"Aturan":da,"REKAP":dr}