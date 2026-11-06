import openpyxl,sys,json,datetime,re
D=sys.argv[1]; OUT=sys.argv[2]
cache={}
def wb(f):
    if f not in cache: cache[f]=openpyxl.load_workbook(f"{D}/NEW - {f}.xlsx",data_only=True,read_only=True)
    return cache[f]
def cv(v):
    if isinstance(v,datetime.datetime): return v.strftime('%Y-%m-%d')
    if isinstance(v,float):
        return int(round(v)) if abs(v-round(v))<1e-6 or abs(v)>=1000 else round(v,4)
    if isinstance(v,str):
        v=v.strip().replace('\xa0',' ')
        return v if v not in ('','-') else None
    return v
def table(f,s,col,val,cols,minlen=1):
    ws=wb(f)[s]; hdr=False; out=[]
    for r in ws.iter_rows(values_only=True):
        if not hdr:
            if len(r)>col and isinstance(r[col],str) and r[col].strip()==val: hdr=True
            continue
        r=list(r)+[None]*60
        row=[cv(r[i]) for i in cols]
        if sum(x is not None for x in row)<minlen: continue
        out.append(row)
    return out
def num(x):
    try: return float(x)
    except: return 0
P0='2026-05'
inwin=lambda d,s=P0: isinstance(d,str) and d[:7]>=s
data={}
data['units']=[r for r in table("01. Database Invoice","M.Unit Aktif",1,"S_Unit",[2,3,4,5,6,7,8,9,10,11,12,13,1],3) if r[0]]
data['tenants']=[r for r in table("01. Database Invoice","M. Tenant",0,"Un_Br_Tn",[5,4,6,7,8,9,10,11,12,13,14],3) if r[0]]
data['sold']=[r for r in table("01. Database Invoice","M.Sold",0,"M.Unit",[9,8,10,7,5,6,2,3],3) if r[0]]
data['rates']=[r for r in table("01. Database Invoice","M. Rate",0,"No",[1,2,3,4,5,6,7],3) if r[3]]
# PPN
ws=wb("01. Database Invoice")["M.PER&SPD"]; ppn=[]; periods=[]
for r in ws.iter_rows(values_only=True):
    r=list(r)+[None]*10
    if isinstance(r[4],(int,float)) and isinstance(r[5],datetime.datetime): ppn.append([r[4],cv(r[5]),cv(r[6])])
    if isinstance(r[1],datetime.datetime): periods.append([cv(r[1]),r[2]])
data['ppn']=ppn
data['meter']=[r for r in table("03. Invoice IPL (AIR)","D.Meteran Air",1,"No Unit",[1,2,4,3,6,5,7,8,15],4) if inwin(r[3])]
data['inet']=table("03. Invoice IPL (INTERNET)","Lg Internet",1,"S_Unit",[6,7,1,2,3,4,5],3)
data['pctps']=table("02. Invoice Profit Sharing","D. % PS",1,"Unit_BrandTenant",[1,2,3,4,5,6,7,8],3)
data['pctowner']=table("02. Invoice Profit Sharing","D. % Pemilik",1,"Unit_BrandTenant",[1,2,3,4,5,6],3)
# invoices
scpl=table("03. Invoice IPL (SC,PL)","DB_SCPL",0,"MP_INV",[3,1,2,6,9,16,7,8,18,17,19,20,12,21,22,23,13,24,11,14,25,26,15],5)
data['inv_scpl']=[r for r in scpl if inwin(r[3])]
air=table("03. Invoice IPL (AIR)","DB_AIR",0,"MP_INV",[3,1,2,6,7,15,16,17,18,19,20,21,24,25,9,10,11,26,27,14,13],5)
data['inv_air']=[r for r in air if inwin(r[3])]
it=table("03. Invoice IPL (INTERNET)","DB_IT",0,"MP_INV",[3,1,2,6,9,16,7,17,18,19,12,20,21,22,11,23,24,15,13],5)
data['inv_it']=[r for r in it if inwin(r[3])]
ps=table("02. Invoice Profit Sharing","DB_PS",0,"MP_INV",[4,2,3,7,10,11,24,8,15,17,18,19,14,20,21,22,23,25,13,16],5)
data['inv_ps']=[r for r in ps if inwin(r[3],'2025-11')]
sw=table("05. Invoice Sewa","DB.INV SEWA",0,"Unit",[0,1,3,4,5,7,8,9,10,11,12,13,15],5)
data['inv_sewa']=[r for r in sw if inwin(r[3],'2026-01')]
data['inv_fee']=table("04. Invoice Fee Sewa","DB.INV FEE SW",1,"Unit",[1,2,10,5,6,7,8,9,4],4)
data['vendor']=table("05. Invoice Sewa","D. Vendor",0,"MP_INV",[2,3,4,6,7,8,9,10,11,12,13],3)
# receipts
rc=table("00. Penerimaan & Pembayaran Rev","Received",0,"U_Mapping",[3,4,5,6,7,8,9,10,11,12,13,14,15,24,29,34,39,1],4)
data['received']=[r for r in rc if inwin(r[4])]
data['paidps']=table("00. Penerimaan & Pembayaran Rev","Paid PS To Owner",1,"No Unit",[1,2,3,4,5,6,7,8,9,10,11,12],4)
data['renov']=table("00. Penerimaan & Pembayaran Rev","FC_Renov",1,"No Unit",[1,2,3,4,5,6,7,8,9,11,12,13,14,15,16,17,21],4)
data['bast']=table("01. Datasbase BAST_STO_BAST PP","D BASTBAST PP",1,"Unit",[1,2,3,4,5,6,7],2)
# aging: outstanding rows
ASOF=datetime.date(2026,11,6)
def aging(sheet,cols):
    rows=table("00. Report IPL Marchand",sheet,1,"Unit",cols,4)
    by={}
    for r in rows:
        out=num(r[7])
        if out<=0.5 or not isinstance(r[4],str): continue
        days=(ASOF-datetime.date.fromisoformat(r[4])).days
        b=0 if days<=30 else 1 if days<=60 else 2 if days<=90 else 3 if days<=120 else 4
        u=by.setdefault(r[0],{'u':r[0],'t':r[1],'b':r[2],'k':[0,0,0,0,0],'inv':[]})
        u['k'][b]+=round(out); u['t']=r[1] or u['t']; u['b']=r[2] or u['b']
        u['inv'].append([str(r[3]),r[4],r[5],r[6],round(num(r[13])),round(out)])
    res=[]
    for u in by.values():
        u['inv'].sort(key=lambda x:x[1])
        res.append([u['u'],u['t'],u['b'],u['k'],u['inv']])
    res.sort(key=lambda x:-sum(x[3]))
    return res
data['ag_scpl']=aging("DA SC, PL",[1,2,3,4,5,6,8,18,19,20,21,22,23,16,17])
data['ag_air']=aging("DA AIR",[1,2,3,4,5,6,6,17,18,19,20,21,22,15,16])
data['ag_it']=aging("DA IT",[1,2,3,4,5,6,8,18,19,20,21,22,23,16,17])
data['ag_ps']=aging("DA PS",[1,2,3,4,5,6,7,16,17,18,19,20,21,14,15])
# aging totals (all invoices)
def agtot(sheet,ci,co):
    rows=table("00. Report IPL Marchand",sheet,1,"Unit",[ci,co],2)
    return [sum(num(r[0]) for r in rows), sum(num(r[1]) for r in rows)]
data['ag_tot']={k:agtot(s,a,b) for k,s,a,b in [("scpl","DA SC, PL",16,17),("air","DA AIR",15,16),("it","DA IT",16,17),("ps","DA PS",14,15)]}
# report
ws=wb("00. Report IPL Marchand")["Report"]; rows=list(ws.iter_rows(values_only=True))
lab=next(i for i,r in enumerate(rows) if len(r)>1 and r[1]=='KETERANGAN'); hdr=rows[lab-1]; months=[]
for i,v in enumerate(hdr):
    if isinstance(v,datetime.datetime) and i>=5 and rows[lab][i]=='Forecast':
        months.append((i,v.strftime('%Y-%m'),))
# dedupe: total columns have same date as prior month (year total). keep first occurrence of each month where next col is Actual and label not total
seen=set(); mcols=[]
for i,m in months:
    if m in seen: continue
    seen.add(m); mcols.append((i,m))
want={6:'SCPL',281:'Air',556:'Internet',831:'Profit Sharing',1106:'Sewa/Pameran',1107:'Fee Sewa',1116:'Biaya Estate Management',1117:'Biaya Personnel',1133:'Transportasi & Akomodasi',1149:'Perbaikan & Perawatan',1164:'Utilities (Air & Listrik)',1184:'Biaya Operasional',1411:'Surplus (Defisit)'}
rep={'months':[m for _,m in mcols],'rows':[]}
for ri,name in want.items():
    r=rows[ri]
    rep['rows'].append([name,[[cv(r[i]) or 0,cv(r[i+1]) or 0] for i,_ in mcols]])
data['report']=rep
json.dump(data,open(OUT,'w'),ensure_ascii=False,separators=(',',':'))
for k,v in data.items(): print(k, len(v) if isinstance(v,list) else '-')
