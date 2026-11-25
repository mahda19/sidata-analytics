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

def audit(paths):
 rows=[]; errors=[]; coverage=[]; examples=[]
 for path in paths:
  bab=os.path.basename(path)[0]; wb=read(path); get=lambda d,a:d.get(a,(None,None,None,False))[0]
  for name,d in wb.items():
   for a,(v,f,t,h) in d.items():
    if t=='e' or (f and '#REF!' in f):errors.append([bab,name,a,str(v),f or ''])
   if name=='REKAP':continue
   title=str(get(d,'A2') or ''); regions=[]
   for a,(v,f,t,h) in d.items():
    if re.fullmatch(r'B\d+',a) and isinstance(v,str) and re.match(r'\s*(Kab\.\s*|Kab\s|Kabupaten\s|Kota\s)',v,re.I):regions.append((int(a[1:]),v.strip()))
   regions.sort(); ids=[r for r,v in regions]; total=[]
   for a,(v,f,t,h) in d.items():
    if re.fullmatch(r'[AB]\d+',a) and isinstance(v,str) and re.fullmatch(r'\s*(total|jumlah|jumlah\s*/\s*total|jawa tengah|provinsi jawa tengah)\s*',v,re.I) and ids and int(re.search(r'\d+',a).group())>max(ids):total.append(int(re.search(r'\d+',a).group()))
   tr=min(total) if total else None; count=0
   extra=[int(a[1:]) for a,(v,*_) in d.items() if re.fullmatch(r'B\d+',a) and isinstance(v,str) and v.strip().lower()=='kanwil kemenag']
   
   if tr and len(ids)==35 and len(set(v.lower() for _,v in regions))==35:
    for a,(v,f,t,h) in d.items():
     m=re.fullmatch(r'([A-Z]+)(\d+)',a)
     if not m or int(m[2])!=tr or m[1] in ('A','B') or not (h or isinstance(v,(int,float)) or t=='e'):continue
     col=m[1]; included=sorted(ids+extra); vals=[get(d,col+str(r)) for r in included]; nums=[x for x in vals if isinstance(x,(int,float))]; missing=len(vals)-len(nums); calc=sum(nums) if not missing else None
     # only sum if explicit SUM or shared formula matching neighbouring anchor; other totals flagged
     expected=f'SUM({col}{min(included)}:{col}{max(included)})'; direct=bool(f and f.replace('$','').upper()==expected and included==list(range(min(included),max(included)+1)))
     shared=False
     if False:
      for aa,(_,ff,_,_) in d.items():
       mm=re.fullmatch(r'([A-Z]+)'+str(tr),aa)
       if mm and ff and ff.replace('$','').upper()==f'SUM({mm[1]}{min(ids)}:{mm[1]}{max(ids)})':shared=True;break
     # shared formula inference not proven; compare numerical sums but mark separate
     method='SUM eksplisit' if direct else 'Kandidat SUM (formula bersama)' if shared else 'Metode perlu konfirmasi'
     status='Perlu konfirmasi metode' if not (direct or shared) else 'Input nonnumerik/kosong' if missing else 'Total error/kosong' if not isinstance(v,(int,float)) else 'Sesuai' if math.isclose(v,calc,rel_tol=1e-10,abs_tol=1e-7) else 'Selisih'
     rows.append([bab,name,title,a,method,v,calc, v-calc if isinstance(v,(int,float)) and calc is not None else None,missing,status]);count+=1
     if status=='Sesuai' and direct:examples.append([bab,name,title,a,calc,min(nums),max(nums),str(get(d,'B'+str(included[nums.index(max(nums))]))).strip()])
   coverage.append([bab,name,len(ids),tr,count,'Diperiksa' if count else 'Perlu pemetaan manual'])
 return rows,errors,coverage,examples
