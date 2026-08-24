# ==============================================================================
# ⚡ SPT DESIGNER - app.py MAESTRO FINAL (8 pestañas, listo Render/GitHub)
# ==============================================================================
import gradio as gr, numpy as np, math, json, os, tempfile
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
try: from scipy import optimize as opt; HAS_SCIPY=True
except Exception: HAS_SCIPY=False

OUT=tempfile.gettempdir()
def ruta(nombre): return os.path.join(OUT,nombre)

def es_finito(v):
    try: return bool(np.isfinite(float(v)))
    except (TypeError,ValueError): return False
def sanitizar(v,d=0.0): return float(v) if es_finito(v) else float(d)
def div_segura(n,d,default=0.0):
    n=float(n); d=float(d)
    if not es_finito(n) or not es_finito(d) or d==0: return float(default)
    return n/d
def sqrt_segura(x,default=0.0):
    x=float(x); return math.sqrt(x) if es_finito(x) and x>0 else float(default)
def fmt(v,dec=1,suf=""): return "N/A" if not es_finito(v) else f"{float(v):.{dec}f}{suf}"
class ErrorValidacion(Exception): pass
def validar_positivo(v,n):
    if not es_finito(v): raise ErrorValidacion(f"{n}: no numérico")
    if float(v)<=0: raise ErrorValidacion(f"{n}: >0")
    return float(v)

PROYECTO={'nombre':'Subestación 115/34.5 kV','cliente':'Corp. Eléctrica del Caribe','ubicacion':'Cartagena','ingeniero':'Ing. C. Martínez'}
SUELO={'rho1':150.0,'rho2':40.0,'h':3.0}
PARAMETROS={'I_falla':8000.0,'t_falla':0.3,'peso':50.0,'rho_grava':3000.0,'h_grava':0.10,'prof_malla':0.6,
 'long_jabalina':2.4,'radio_conductor':0.0053,'k_material':202.8,'material_conductor':'Cobre blando','calibre':'2/0 AWG'}
MATERIALES={'Grava húmeda':{'rho_s':3000,'hs':0.10,'fuente':'IEEE 80-2013','nota':'Estándar.'},
 'Grava seca':{'rho_s':10000,'hs':0.10,'fuente':'IEEE 80-2013','nota':'Seco.'},
 'Concreto seco':{'rho_s':2000,'hs':0.10,'fuente':'Referencial','nota':'Loseta seca.'},
 'Concreto húmedo':{'rho_s':150,'hs':0.10,'fuente':'Referencial','nota':'Loseta saturada (Cs bajo).'},
 'Suelo con pasto':{'rho_s':None,'hs':0.0,'fuente':'Campo','nota':'Sin capa.'},
 'Tapete dieléctrico':{'rho_s':1000000,'hs':0.01,'fuente':'ASTM F605','nota':'Puntos de operación.'}}
MATCOND={'Cobre blando':{'k':202.8},'Cobre duro':{'k':197.0},'Copperweld':{'k':125.0}}
CALIBRES={'2 AWG':6.54,'1/0 AWG':8.25,'2/0 AWG':9.27,'4/0 AWG':11.68,'250 kcmil':14.22,'300 kcmil':15.54,'500 kcmil':20.07}
DIRECCIONES=['N-S','E-W','NE-SW','NW-SE','D1','D2','D3','D4']
estado={'conductores':[],'jabalinas':[],'suelo_fit':None,'suelo_aplicado':False,'alerts':[],'capa':None,'suelo_rows':[],'ultimo_res':None,'dirty':False}

def stepper():
    c2='✅' if estado.get('suelo_fit') else '⬜'; c3='✅' if estado['conductores'] else '⬜'; c4='✅' if estado.get('ultimo_res') else '⬜'
    nxt='3 📐 Geometría' if not estado['conductores'] else ('4 ⚡ Análisis' if not estado.get('ultimo_res') else '7 📄 Informe')
    dirty=' · ⚠️ cambios sin guardar (💾 Exportar en 7)' if estado.get('dirty') else ''
    return f"**Ruta:** 1📋✅ → 2🌍{c2} → 3📐{c3} → 4⚡{c4} → 5📈 → 6🎯 → 7📄 → 8🔬 |  siguiente: **{nxt}**{dirty}"

def aplicar_datos(nom,cli,ubi,ing, I,t, peso,pm,lj, mat_cond, calibre):
    diam=CALIBRES[calibre]; radio_m=diam/2000.0; k=MATCOND[mat_cond]['k']
    PROYECTO.update({'nombre':nom,'cliente':cli,'ubicacion':ubi,'ingeniero':ing})
    PARAMETROS.update({'I_falla':float(I),'t_falla':float(t),'peso':float(peso),'prof_malla':float(pm),
        'long_jabalina':float(lj),'radio_conductor':radio_m,'k_material':k,'material_conductor':mat_cond,'calibre':calibre})
    estado['dirty']=True
    return f"✅ Guardado: {nom} · I={float(I):.0f} A · t={float(t):.2f} s · {mat_cond} {calibre} (Ø={diam:.2f} mm)"
def aplicar_suelo_medido():
    f=estado.get('suelo_fit')
    if not f: return "⚠️ Primero pulsa '📊 Analizar' con ≥4 mediciones.", stepper()
    SUELO['rho1']=f['r1']; SUELO['rho2']=f['r2']; SUELO['h']=f['h']
    estado['suelo_aplicado']=True; estado['dirty']=True
    return f"📌 Suelo medido aplicado: ρ1={f['r1']:.1f}, ρ2={f['r2']:.1f}, h={f['h']:.2f} m. El BEM ya lo usa.", stepper()
def aplicar_capa(material,hs_sel,hs_custom):
    m=MATERIALES[material]; rho_s=m['rho_s'] or SUELO['rho1']
    hs=float(hs_custom) if (hs_custom is not None and float(hs_custom)>0) else float(hs_sel)
    PARAMETROS['rho_grava']=rho_s; PARAMETROS['h_grava']=hs
    estado['capa']={'material':material,'rho_s':rho_s,'hs':hs,'fuente':m['fuente'],'nota':m['nota']}
    estado['dirty']=True
    return f"✅ Capa: {material}, ρs={rho_s:.0f} Ω·m, hs={hs:.2f} m · Fuente: {m['fuente']}"
def exportar_proyecto():
    data={'PROYECTO':PROYECTO,'PARAMETROS':PARAMETROS,'SUELO':SUELO,'conductores':estado['conductores'],'jabalinas':estado['jabalinas'],'capa':estado['capa']}
    open(ruta('proyecto_spt.json'),'w').write(json.dumps(data)); estado['dirty']=False
    return "✅ Proyecto exportado (proyecto_spt.json)", ruta('proyecto_spt.json')
def importar_proyecto(f):
    if f is None: return "⚠️ sube el JSON del proyecto", None
    d=json.load(open(f.name))
    PROYECTO.update(d.get('PROYECTO',{})); PARAMETROS.update(d.get('PARAMETROS',{})); SUELO.update(d.get('SUELO',{}))
    estado['conductores']=[tuple(c) for c in d.get('conductores',[])]; estado['jabalinas']=[tuple(j) for j in d.get('jabalinas',[])]
    estado['capa']=d.get('capa')
    return f"✅ Proyecto cargado: {len(estado['conductores'])} cond, {len(estado['jabalinas'])} jab", None

def base_rect(wx,wy,sp):
    cond=[]; nx=max(1,int(round(wx/sp))); ny=max(1,int(round(wy/sp))); sx=wx/nx; sy=wy/ny
    for j in range(ny+1):
        for i in range(nx): cond.append((i*sx,j*sy,(i+1)*sx,j*sy))
    for i in range(nx+1):
        for j in range(ny): cond.append((i*sx,j*sy,i*sx,(j+1)*sy))
    return cond
def crear_malla_rect(wx,wy,sp,patron='perimetro'):
    wx=validar_positivo(wx,"X"); wy=validar_positivo(wy,"Y"); sp=validar_positivo(sp,"esp")
    sp=min(sp,min(wx,wy)); sp=max(sp,0.5); cond=base_rect(wx,wy,sp); jab=[]
    if patron=='perimetro':
        for i in range(int(round(wx/sp))+1): jab.append((i*sp,0)); jab.append((i*sp,wy))
        for j in range(1,int(round(wy/sp))): jab.append((0,j*sp)); jab.append((wx,j*sp))
    elif patron=='todos':
        for i in range(int(round(wx/sp))+1):
            for j in range(int(round(wy/sp))+1): jab.append((i*sp,j*sp))
    else: jab=[(0,0),(wx,0),(0,wy),(wx,wy)]
    return cond,jab
def snap_coords(cond,tol=0.05):
    pts=[]
    for c in cond: pts+=[(c[0],c[1]),(c[2],c[3])]
    cl=[]
    def find(p):
        for c in cl:
            if math.hypot(p[0]-c[0][0],p[1]-c[0][1])<=tol: return c
        return None
    for p in pts:
        c=find(p)
        if c is None: cl.append([p,[p]])
        else: c[1].append(p)
    rep={}
    for p in pts:
        c=find(p); xs=[q[0] for q in c[1]]; ys=[q[1] for q in c[1]]; rep[p]=(sum(xs)/len(xs),sum(ys)/len(ys))
    return [(rep[(a,b)][0],rep[(a,b)][1],rep[(c,d)][0],rep[(c,d)][1]) for (a,b,c,d) in cond]
def limpiar_dup_cero(cond,tol=0.05):
    seen=set(); out=[]; cero=0; dup=0
    for (x1,y1,x2,y2) in cond:
        if math.hypot(x2-x1,y2-y1)<tol: cero+=1; continue
        p1=(round(x1,3),round(y1,3)); p2=(round(x2,3),round(y2,3)); k=(min(p1,p2),max(p1,p2))
        if k in seen: dup+=1; continue
        seen.add(k); out.append((x1,y1,x2,y2))
    return out,cero,dup
def _cross(a,b):
    (x1,y1,x2,y2)=a;(x3,y3,x4,y4)=b
    def ccw(A,B,C): return (C[1]-A[1])*(B[0]-A[0])>(B[1]-A[1])*(C[0]-A[0])
    A=(x1,y1);B=(x2,y2);C=(x3,y3);D=(x4,y4)
    return ccw(A,C,D)!=ccw(B,C,D) and ccw(A,B,C)!=ccw(A,B,D)
def detectar_flotantes(cond,tol=0.05):
    from collections import defaultdict
    n=len(cond); parent=list(range(n))
    def find(a):
        while parent[a]!=a: parent[a]=parent[parent[a]]; a=parent[a]
        return a
    def key(p): return (round(p[0]/tol),round(p[1]/tol))
    nod={}
    for i,c in enumerate(cond):
        for p in [(c[0],c[1]),(c[2],c[3])]:
            k=key(p)
            if k in nod: parent[find(i)]=find(nod[k])
            else: nod[k]=i
    for i in range(n):
        for j in range(i+1,n):
            if _cross(cond[i],cond[j]): parent[find(i)]=find(j)
    comp=defaultdict(list)
    for i in range(n): comp[find(i)].append(i)
    if not comp: return []
    main=max(comp,key=lambda k:len(comp[k]))
    return sorted(i for k,idx in comp.items() if k!=main for i in idx)

def wenner_2capas(a,rho1,rho2,h,terms=150):
    a=np.asarray(a,float); K=(rho2-rho1)/(rho2+rho1); s=np.zeros_like(a)
    for n in range(1,terms+1): s+=(K**n)*(1/np.sqrt(1+(2*n*h/a)**2)-1/np.sqrt(4+(2*n*h/a)**2))
    return rho1*(1+4*s)
def detectar_anomalias(rows):
    from collections import Counter
    alerts=[]; by_a={}
    for a,rho,d in rows: by_a.setdefault(a,[]).append((rho,d))
    dirs=sorted(set(d for _,_,d in rows)); out=[]
    for a,vals in sorted(by_a.items()):
        if len(vals)>=2:
            arr=np.array([v for v,_ in vals]); med=np.median(arr)
            for rho,d in vals:
                dev=(rho-med)/med*100 if med>0 else 0
                if abs(dev)>40: out.append((d,a,rho,dev))
    cnt=Counter(d for d,_,_,_ in out)
    for d,a,rho,dev in out:
        alerts.append({'tipo':'hallazgo' if cnt[d]>=2 else 'error_medida','dir':d,'a':a,'rho':rho,'desvio':dev})
    dm={d:np.mean([rho for _,rho,dd in rows if dd==d]) for d in dirs}; ov=np.mean(list(dm.values())) if dm else 0
    for d,m in dm.items():
        dev=(m-ov)/ov*100 if ov>0 else 0
        if abs(dev)>25: alerts.append({'tipo':'anisotropia','dir':d,'a':None,'rho':m,'desvio':dev})
    return alerts
def _analizar_rows(rows):
    try:
        if len(rows)<4: return "⚠️ Agrega ≥4 filas (elige dirección y carga a,ρ)",[],None
        A=np.array([r[0] for r in rows]); R=np.array([r[1] for r in rows]); uniq=sorted(set(A))
        stats=[[a,round(R[A==a].mean(),1),int((A==a).sum())] for a in uniq]
        amean=np.array([s[0] for s in stats]); rmean=np.array([s[1] for s in stats])
        rho1c=float(np.exp(np.log(R).mean()))
        def obj(x):
            m=wenner_2capas(amean,np.exp(x[0]),np.exp(x[1]),np.exp(x[2]))
            if np.any(~np.isfinite(m)) or np.any(m<=0): return 1e12
            return np.sum((np.log(m)-np.log(rmean))**2)
        r1,r2,h=(np.exp(opt.minimize(obj,[np.log(max(rmean[0],1)),np.log(max(rmean[-1],1)),np.log(max(np.median(amean)/2,0.3))],method='Nelder-Mead',options={'maxiter':2000}).x) if HAS_SCIPY else (rmean[0],rmean[-1],1.0))
        estado['suelo_fit']={'rho1c':rho1c,'r1':r1,'r2':r2,'h':h}; estado['alerts']=detectar_anomalias(rows)
        fig,ax=plt.subplots(figsize=(9,6))
        for d in sorted(set(r[2] for r in rows)):
            ax.scatter([A[i] for i in range(len(A)) if rows[i][2]==d],[R[i] for i in range(len(A)) if rows[i][2]==d],label=d,alpha=0.6)
        ax.plot(amean,rmean,'ko-',lw=2); aa=np.logspace(np.log10(min(A)*0.8),np.log10(max(A)*1.2),60)
        ax.plot(aa,wenner_2capas(aa,r1,r2,h),'r-'); ax.set_xscale('log'); ax.set_yscale('log'); ax.grid(alpha=0.3,which='both'); ax.legend(); plt.close(fig)
        return f"**2 capas:** ρ1={r1:.1f}, ρ2={r2:.1f}, h={h:.2f} m · **Alertas:** {len(estado['alerts'])}. Pulsa **📌 Usar suelo medido**.",stats,fig
    except Exception as e: return f"❌ {e}",[],None
def agregar_direccion(dsel, txt):
    add=[]
    for l in (txt or "").split('\n'):
        p=l.replace(',',' ').split()
        if len(p)>=2:
            try:
                a=float(p[0]); rho=float(p[1])
                if a>0 and rho>0: add.append((a,rho,dsel))
            except Exception: continue
    estado['suelo_rows']+=add
    return [[f"{x[0]:g}",f"{x[1]:g}",x[2]] for x in estado['suelo_rows']], f"➕ {len(add)} filas de **{dsel}** · total {len(estado['suelo_rows'])}", ""
def limpiar_conjunto():
    estado['suelo_rows']=[]; return [], "🧹 Conjunto vacío"

def bem_analizar(cond,rho,I,d=0.6,radio=0.0053,res_sup=1.0):
    N=len(cond); A=np.array([[c[0],c[1],-d] for c in cond],float); B=np.array([[c[2],c[3],-d] for c in cond],float)
    Lseg=np.linalg.norm(B-A,axis=1); mid=(A+B)/2; Ai=A.copy(); Ai[:,2]=d; Bi=B.copy(); Bi[:,2]=d
    coeff=rho/(4*np.pi); R=np.zeros((N,N))
    for j in range(N):
        Lj=Lseg[j]
        r1=np.linalg.norm(mid-A[j],axis=1); r2=np.linalg.norm(mid-B[j],axis=1)
        real=np.log((r1+r2+Lj)/np.maximum(r1+r2-Lj,1e-9))
        r1i=np.linalg.norm(mid-Ai[j],axis=1); r2i=np.linalg.norm(mid-Bi[j],axis=1)
        imag=np.log((r1i+r2i+Lj)/np.maximum(r1i+r2i-Lj,1e-9))
        col=coeff/Lj*(real+imag)
        rs=np.sqrt((Lj/2)**2+radio**2); sr=np.log((2*rs+Lj)/max(2*rs-Lj,1e-9))
        rsi=np.sqrt((Lj/2)**2+(2*d)**2); si=np.log((2*rsi+Lj)/max(2*rsi-Lj,1e-9))
        col[j]=coeff/Lj*(sr+si); R[:,j]=col
    try: x=np.linalg.solve(R,np.ones(N))
    except np.linalg.LinAlgError: x=np.linalg.pinv(R)@np.ones(N)
    S=x.sum()
    if not es_finito(S) or S<=0: return None
    Rg=1.0/S; GPR=I*Rg; Iseg=GPR*x
    xs=[c[0] for c in cond]+[c[2] for c in cond]; ys=[c[1] for c in cond]+[c[3] for c in cond]
    gx=np.arange(min(xs)-5,max(xs)+5,res_sup); gy=np.arange(min(ys)-5,max(ys)+5,res_sup)
    GX,GY=np.meshgrid(gx,gy)
    P=np.concatenate([GX.ravel().reshape(-1,1),GY.ravel().reshape(-1,1),np.zeros((GX.size,1))],axis=1)
    V=np.zeros(P.shape[0])
    for j in range(N):
        Lj=Lseg[j]; r1=np.linalg.norm(P-A[j],axis=1); r2=np.linalg.norm(P-B[j],axis=1)
        V+=(coeff/Lj)*Iseg[j]*2*np.log((r1+r2+Lj)/np.maximum(r1+r2-Lj,1e-9))
    Vs=V.reshape(GX.shape)
    Es=float(max(np.abs(np.diff(Vs,axis=1)).max(),np.abs(np.diff(Vs,axis=0)).max())); Et=0.0
    for c in cond:
        mx=(c[0]+c[2])/2; my=(c[1]+c[3])/2; dx=c[2]-c[0]; dy=c[3]-c[1]; Lc=math.hypot(dx,dy)
        if Lc<1e-9: continue
        Et=max(Et,GPR-_interp(GX,GY,Vs,mx-dy/Lc,my+dx/Lc))
    return {'R_grid':Rg,'GPR':GPR,'Es':Es,'Et':float(Et),'GX':GX,'GY':GY,'Vsurf':Vs}
def _interp(GX,GY,V,px,py):
    xs=GX[0,:]; ys=GY[:,0]
    if px<xs.min() or px>xs.max() or py<ys.min() or py>ys.max(): return 0.0
    i=int(np.clip(np.searchsorted(xs,px)-1,0,len(xs)-2)); j=int(np.clip(np.searchsorted(ys,py)-1,0,len(ys)-2))
    tx=(px-xs[i])/max(xs[i+1]-xs[i],1e-9); ty=(py-ys[j])/max(ys[j+1]-ys[j],1e-9)
    return (V[j,i]*(1-tx)+V[j,i+1]*tx)*(1-ty)+(V[j+1,i]*(1-tx)+V[j+1,i+1]*tx)*ty
def limites_ieee(hs,rho_s,rho1,t,peso=50.0):
    Cs=max(0.0,min(1.0,1.0-0.09*(1.0-rho1/rho_s)/(2.0*hs+0.09)))
    k=0.116 if peso<=50 else 0.157
    return Cs,(1000+6*Cs*rho_s)*k/math.sqrt(t),(1000+1.5*Cs*rho_s)*k/math.sqrt(t)
def calcular(cond,jab,I=None,t=None):
    if not cond: return None
    L=sum(math.hypot(c[2]-c[0],c[3]-c[1]) for c in cond)
    xs=[c[0] for c in cond]+[c[2] for c in cond]; ys=[c[1] for c in cond]+[c[3] for c in cond]
    area=(max(xs)-min(xs))*(max(ys)-min(ys))
    I=sanitizar(I or PARAMETROS['I_falla'],1000); t=min(max(sanitizar(t or PARAMETROS['t_falla'],0.3),0.01),10)
    bem=bem_analizar(cond,SUELO['rho1'],I,PARAMETROS['prof_malla'],PARAMETROS['radio_conductor']) if len(cond)<=600 else None
    if bem is None:
        R=div_segura(SUELO['rho1'],4*sqrt_segura(area/math.pi)); GPR=I*R; Es=GPR*0.15; Et=GPR*0.08; met='Simplificado'
    else: R,GPR,Es,Et=bem['R_grid'],bem['GPR'],bem['Es'],bem['Et']; met='BEM'
    Cs,Ep,Tp=limites_ieee(PARAMETROS['h_grava'],PARAMETROS['rho_grava'],SUELO['rho1'],t,PARAMETROS['peso'])
    k=PARAMETROS.get('k_material',202.8)
    return {'n_cond':len(cond),'n_jab':len(jab),'L':L,'area':area,'R':sanitizar(R),'GPR':sanitizar(GPR),'Cs':Cs,
            'Es':sanitizar(Es),'Es_perm':Ep,'Et':sanitizar(Et),'Et_perm':Tp,'A_min':div_segura(I*sqrt_segura(t),k),
            'cumple':bool(Es<=Ep and Et<=Tp),'metodo':met,'bem':bem}

# ================= FASE A: PRECISIÓN =================
def factor_decremento(XoR,t,f=60.0):
    if not (XoR>0 and t>0): return 1.0
    Ta=XoR/(2*math.pi*f)
    return math.sqrt(1.0+(Ta/t)*(1.0-math.exp(-2.0*t/Ta)))
def segmentar(cond,max_len=5.0):
    out=[]
    for (x1,y1,x2,y2) in cond:
        L=math.hypot(x2-x1,y2-y1); n=max(1,int(math.ceil(L/max_len)))
        for k in range(n):
            a=k/n; b=(k+1)/n
            out.append((x1+(x2-x1)*a,y1+(y2-y1)*a,x1+(x2-x1)*b,y1+(y2-y1)*b))
    return out
def bem_preciso(cond,rho,I,d=0.6,radio=0.0053):
    N=len(cond); A=np.array([[c[0],c[1],-d] for c in cond]); B=np.array([[c[2],c[3],-d] for c in cond])
    Lseg=np.linalg.norm(B-A,axis=1); mid=(A+B)/2; Ai=A.copy(); Ai[:,2]=d; Bi=B.copy(); Bi[:,2]=d
    cf=rho/(4*np.pi); R=np.zeros((N,N))
    for j in range(N):
        Lj=Lseg[j]
        r1=np.linalg.norm(mid-A[j],axis=1); r2=np.linalg.norm(mid-B[j],axis=1)
        re=np.log((r1+r2+Lj)/np.maximum(r1+r2-Lj,1e-9))
        r1i=np.linalg.norm(mid-Ai[j],axis=1); r2i=np.linalg.norm(mid-Bi[j],axis=1)
        im=np.log((r1i+r2i+Lj)/np.maximum(r1i+r2i-Lj,1e-9))
        col=cf/Lj*(re+im)
        rs=np.sqrt((Lj/2)**2+radio**2); sr=np.log((2*rs+Lj)/max(2*rs-Lj,1e-9))
        rsi=np.sqrt((Lj/2)**2+(2*d)**2); si=np.log((2*rsi+Lj)/max(2*rsi-Lj,1e-9))
        col[j]=cf/Lj*(sr+si); R[:,j]=col
    x=np.linalg.solve(R,np.ones(N)); S=x.sum(); Rg=1.0/S; GPR=I*Rg; Iseg=GPR*x
    def pot(pts):
        pts=np.asarray(pts,float); V=np.zeros(len(pts))
        for j in range(N):
            Lj=Lseg[j]
            r1=np.linalg.norm(pts[:,None,:]-A[j],axis=2)[:,0]; r2=np.linalg.norm(pts[:,None,:]-B[j],axis=2)[:,0]
            r1i=np.linalg.norm(pts[:,None,:]-Ai[j],axis=2)[:,0]; r2i=np.linalg.norm(pts[:,None,:]-Bi[j],axis=2)[:,0]
            V+=(cf/Lj)*Iseg[j]*(np.log((r1+r2+Lj)/np.maximum(r1+r2-Lj,1e-9))+np.log((r1i+r2i+Lj)/np.maximum(r1i+r2i-Lj,1e-9)))
        return V
    return {'R':Rg,'GPR':GPR,'Iseg':Iseg,'pot':pot}
def calcular_preciso(cond,I,t,rho1,XoR=10.0,max_len=5.0,puntos=None,d=0.6,radio=0.0053):
    Df=factor_decremento(XoR,t); Ieff=I*Df; cond2=segmentar(cond,max_len)
    b=bem_preciso(cond2,rho1,Ieff,d,radio)
    out={'Df':Df,'Ieff':Ieff,'n_seg':len(cond2),'R':b['R'],'GPR':b['GPR']}
    if puntos: out['trans']=[(p,float(v)) for p,v in zip(puntos,b['pot'](np.array([[p[0],p[1],0.0] for p in puntos])))]
    return out

def perfiles(I,t,tipo,px1,py1,px2,py2):
    cond=estado['conductores']; res=calcular(cond,estado['jabalinas'],I,t)
    if res is None or res['bem'] is None: return "⚠️",None,None,None
    b=res['bem']; GPR=b['GPR']; GX,GY,Vs=b['GX'],b['GY'],b['Vsurf']
    Cs,Ep,Tp=limites_ieee(PARAMETROS['h_grava'],PARAMETROS['rho_grava'],SUELO['rho1'],t,PARAMETROS['peso'])
    xs=[c[0] for c in cond]+[c[2] for c in cond]; ys=[c[1] for c in cond]+[c[3] for c in cond]
    x0,x1,y0,y1=min(xs),max(xs),min(ys),max(ys)
    P1,P2={'diagonal':((x0,y0),(x1,y1)),'horizontal':((x0,(y0+y1)/2),(x1,(y0+y1)/2)),'vertical':(((x0+x1)/2,y0),((x0+x1)/2,y1))}.get(tipo,((px1,py1),(px2,py2)))
    Ll=math.hypot(P2[0]-P1[0],P2[1]-P1[1]); n=max(2,int(Ll/0.25)); ss=np.linspace(0,Ll,n)
    Vp=np.array([_interp(GX,GY,Vs,P1[0]+(P2[0]-P1[0])*s/Ll,P1[1]+(P2[1]-P1[1])*s/Ll) for s in ss])
    idx=max(1,int(0.5/(Ll/(n-1)))); paso=np.array([abs(Vp[min(i+idx,n-1)]-Vp[max(i-idx,0)]) for i in range(n)]); cont=GPR-Vp
    figs=[]
    for dat,col,tit,lim,lv in [(Vp,'b','Potencial',None,GPR),(paso,'g','Paso','Es',Ep),(cont,'m','Contacto','Et',Tp)]:
        fig,ax=plt.subplots(figsize=(9,5)); ax.plot(ss,dat,col+'-',lw=2)
        ax.axhline(lv,color='r',ls='--',lw=2,label=f'{lim or "GPR"}={lv:.0f}V')
        if lim: ax.fill_between(ss,0,lv,alpha=0.1,color='green')
        ax.legend(); ax.grid(alpha=0.3); ax.set_title(tit); figs.append(fig)
    return f"**Paso máx: {paso.max():.0f} V de {Ep:.0f} V · Contacto máx: {cont.max():.0f} V de {Tp:.0f} V**",figs[0],figs[1],figs[2]
def optimizar(wx,wy,I,t,rho1):
    mejor=None
    for s in [5,4,3]:
        for pat in ['perimetro','todos']:
            cond,jab=crear_malla_rect(wx,wy,s,pat)
            if len(cond)>600: continue
            bem=bem_analizar(cond,rho1,I,res_sup=2.0)
            if bem is None: continue
            for hs in [0.1,0.2,0.3]:
                for rs in [3000,5000]:
                    Cs,Ep,Tp=limites_ieee(hs,rs,rho1,t)
                    I_max=I*min(div_segura(Ep,bem['Es'],0),div_segura(Tp,bem['Et'],0)) if bem['Es']>0 and bem['Et']>0 else I
                    if mejor is None or I_max>mejor[0]: mejor=(I_max,s,pat,hs,rs,bem)
    if mejor is None: return "⚠️"
    I_max,s,pat,hs,rs,bem=mejor
    return f"**Mejor configuración:** esp {s} m, jab {pat}, hs={hs} m, ρs={rs} Ω·m\n- I máx que cumple: {I_max:.0f} A\n## {'✅ CUMPLE' if I_max>=I else f'❌ NO con {I:.0f} A'}"

def parse_dxf(text):
    tok=text.split(); n=len(tok); i=0; segs=[]; pts=[]; cur=None; d={}; verts=[]; closed=False
    def flush():
        nonlocal d,verts,closed
        if cur=='LINE':
            if all(k in d for k in (10,20,11,21)): segs.append((d[10],d[20],d[11],d[21]))
        elif cur=='LWPOLYLINE':
            v=[(a,b) for a,b in verts if a is not None and b is not None]
            for k in range(len(v)-1): segs.append((v[k][0],v[k][1],v[k+1][0],v[k+1][1]))
            if closed and len(v)>2: segs.append((v[-1][0],v[-1][1],v[0][0],v[0][1]))
        elif cur in ('POINT','CIRCLE'):
            if 10 in d and 20 in d: pts.append((d[10],d[20]))
        d={}; verts=[]; closed=False
    while i<n-1:
        code=tok[i]; val=tok[i+1]; i+=2
        if code=='0': flush(); cur=val
        else:
            try: num=float(val)
            except ValueError: num=None
            if cur=='LWPOLYLINE':
                if code=='10' and num is not None: verts.append([num,None])
                elif code=='20' and num is not None and verts: verts[-1][1]=num
                elif code=='70' and num is not None: closed=bool(int(num)&1)
            elif num is not None:
                try: d[int(code)]=num
                except ValueError: pass
    flush(); return segs,pts
def importar_dxf(texto,modo):
    segs,pts=parse_dxf(texto)
    if not segs: return "⚠️ sin segmentos",[],[]
    mx=max([abs(v) for s in segs for v in s]); sc=0.001 if (modo=='mm' or (modo=='auto' and mx>1000)) else 1.0
    cond=[(a*sc,b*sc,c*sc,d*sc) for a,b,c,d in segs]
    cond,cero,dup=limpiar_dup_cero(cond); cond=snap_coords(cond); flot=detectar_flotantes(cond)
    return f"📥 ×{sc}: {len(segs)} seg, {len(pts)} jab, {len(flot)} flot",cond,(pts or [(0,0)])
def write_dxf(cond,jab,path):
    L=['0','SECTION','2','ENTITIES']
    for (x1,y1,x2,y2) in cond: L+=['0','LINE','8','MALLA','10',f"{x1:.3f}",'20',f"{y1:.3f}",'30','0','11',f"{x2:.3f}",'21',f"{y2:.3f}",'31','0']
    for (x,y) in jab: L+=['0','CIRCLE','8','JABALINAS','10',f"{x:.3f}",'20',f"{y:.3f}",'30','0','40','0.1']
    L+=['0','ENDSEC','0','EOF']; open(path,'w').write("\n".join(L)); return path
def tok(*p): return [str(x) for x in p]
def dxf_line(x1,y1,x2,y2): return tok(0,'LINE',8,'MALLA',10,x1,20,y1,30,0,11,x2,21,y2,31,0)
def dxf_circle(x,y): return tok(0,'CIRCLE',8,'JABALINAS',10,x,20,y,30,0,40,0.1)
def dxf_body(e): return "\n".join(tok(0,'SECTION',2,'ENTITIES')+[t for ent in e for t in ent]+tok(0,'ENDSEC',0,'EOF'))
def ejemplo_sucio():
    e=[]
    for i in range(0,30000,5000): e.append(dxf_line(i,0,i,20000))
    for j in range(0,20000,5000): e.append(dxf_line(0,j,30000,j))
    e.append(dxf_line(0,25000,10000,25000)); e.append(dxf_line(10020,25000,20000,25000))
    e.append(dxf_line(0,0,5000,0)); e.append(dxf_line(0,0,5000,0)); e.append(dxf_line(5000,5000,5000,5000))
    e.append(dxf_line(100000,100000,110000,110000)); e+=[dxf_circle(0,0),dxf_circle(30000,20000)]
    return dxf_body(e)

def figura_2d(cond,jab,t='Malla'):
    fig,ax=plt.subplots(figsize=(8,6))
    for (x1,y1,x2,y2) in cond: ax.plot([x1,x2],[y1,y2],'b-',lw=1.5)
    for (x,y) in jab: ax.plot(x,y,'ro',ms=6)
    ax.set_aspect('equal'); ax.grid(alpha=0.3); ax.set_title(t,fontweight='bold'); return fig
def figura_3d(cond,jab):
    fig=plt.figure(figsize=(10,7)); ax=fig.add_subplot(111,projection='3d'); z=-PARAMETROS['prof_malla']
    for (x1,y1,x2,y2) in cond: ax.plot([x1,x2],[y1,y2],[z,z],'b-',lw=1.5)
    for (x,y) in jab: ax.plot([x,x],[y,y],[z,z-PARAMETROS['long_jabalina']],'r-',lw=2.5)
    ax.set_title('3D',fontweight='bold'); return fig
def geo_render(msg):
    c=estado['conductores']; j=estado['jabalinas']
    L=sum(math.hypot(x2-x1,y2-y1) for (x1,y1,x2,y2) in c)
    dfc=[[i+1,round(a,1),round(b,1),round(cc,1),round(dd,1)] for i,(a,b,cc,dd) in enumerate(c)]
    dfj=[[i+1,round(x,1),round(y,1)] for i,(x,y) in enumerate(j)]
    info=f"{msg}\n**Conductores:** {len(c)} | **Jabalinas:** {len(j)} | **Longitud:** {L:.1f} m"
    return info,dfc,dfj,figura_2d(c,j),figura_3d(c,j)
def editor_add_cond(x1,y1,x2,y2):
    if not all(es_finito(v) for v in (x1,y1,x2,y2)): return geo_render("⚠️ coordenadas inválidas")
    estado['conductores'].append((float(x1),float(y1),float(x2),float(y2))); estado['dirty']=True
    return geo_render("➕ Conductor agregado")
def editor_add_jab(x,y):
    if not all(es_finito(v) for v in (x,y)): return geo_render("⚠️ coordenadas inválidas")
    estado['jabalinas'].append((float(x),float(y))); estado['dirty']=True
    return geo_render("➕ Jabalina agregada")
def editor_del_cond(i):
    c=estado['conductores']
    if c and es_finito(i) and 1<=int(i)<=len(c): c.pop(int(i)-1); estado['dirty']=True; return geo_render("➖ Conductor eliminado")
    return geo_render("⚠️ id de conductor fuera de rango (mira la tabla)")
def editor_del_jab(i):
    j=estado['jabalinas']
    if j and es_finito(i) and 1<=int(i)<=len(j): j.pop(int(i)-1); estado['dirty']=True; return geo_render("➖ Jabalina eliminada")
    return geo_render("⚠️ id de jabalina fuera de rango (mira la tabla)")
def guardar_graficas(cond,jab,res):
    fig=figura_2d(cond,jab); fig.savefig(ruta('vista_planta.png'),dpi=150,bbox_inches='tight'); plt.close(fig)
    fig=figura_3d(cond,jab); fig.savefig(ruta('vista_3d.png'),dpi=150,bbox_inches='tight'); plt.close(fig)
    if res['bem'] is not None: GX,GY,V=res['bem']['GX'],res['bem']['GY'],res['bem']['Vsurf']
    else: GX,GY=np.meshgrid(np.arange(0,50),np.arange(0,50)); V=res['GPR']*np.exp(-np.hypot(GX-25,GY-25)/15)
    fig,ax=plt.subplots(figsize=(8,6)); im=ax.pcolormesh(GX,GY,V,cmap='hot',shading='auto')
    plt.colorbar(im,ax=ax); fig.savefig(ruta('mapa_calor.png'),dpi=150,bbox_inches='tight'); plt.close(fig)
def crear_informe(cond,jab,res):
    doc=Document()
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run("SPT DESIGNER"); r.font.size=Pt(28); r.bold=True
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run(PROYECTO['nombre']); r.font.size=Pt(14)
    doc.add_paragraph(f"{PROYECTO['cliente']} · {PROYECTO['ubicacion']} · {PROYECTO['ingeniero']}")
    doc.add_heading("TEORÍA IEEE 80",level=1)
    doc.add_paragraph("Paso/contacto vs límites Es_perm=(1000+6Csρs)0.116/√t, Et_perm=(1000+1.5Csρs)0.116/√t.")
    doc.add_heading(f"RESULTADOS ({res['metodo']})",level=1)
    t=doc.add_table(rows=8,cols=2); t.style='Table Grid'
    for i,(k,v) in enumerate([("R",fmt(res['R'],3," Ω")),("GPR",fmt(res['GPR'],1," V")),("Es",f"{fmt(res['Es'],1)}/{fmt(res['Es_perm'],1)} V"),("Et",f"{fmt(res['Et'],1)}/{fmt(res['Et_perm'],1)} V"),("Cs",fmt(res['Cs'],3)),("Conductor",f"{PARAMETROS['material_conductor']} {PARAMETROS['calibre']}"),("Área cond.",fmt(res['A_min'],1," mm²")),("Veredicto","CUMPLE" if res['cumple'] else "NO CUMPLE")]):
        t.rows[i].cells[0].text=k; t.rows[i].cells[1].text=v
    if estado.get('capa'):
        doc.add_heading("CAPA SUPERFICIAL",level=1); doc.add_paragraph(f"{estado['capa']['material']} ρs={estado['capa']['rho_s']:.0f} hs={estado['capa']['hs']:.2f} Fuente:{estado['capa']['fuente']}")
    if estado['alerts']:
        doc.add_heading("HALLAZGOS",level=1)
        for a in estado['alerts']: doc.add_paragraph(f"[{a['tipo']}] {a['dir']} ρ={a['rho']:.0f}",style='List Bullet')
    try:
        doc.add_picture(ruta('vista_planta.png'),width=Inches(6)); doc.add_picture(ruta('mapa_calor.png'),width=Inches(6))
    except Exception as e: doc.add_paragraph(str(e))
    doc.save(ruta('Informe_SPT.docx')); return ruta('Informe_SPT.docx')

def _card(t,v,s=""): return f'<div style="flex:1;min-width:100px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:10px;text-align:center"><div style="font-size:12px;color:#64748b">{t}</div><div style="font-size:20px;font-weight:700">{v}</div><div style="font-size:11px;color:#94a3b8">{s}</div></div>'
def _barra(t,c,p):
    r=c/p if p>0 else 0; pct=min(r,1.0)*100; col='#22c55e' if r<=1 else '#ef4444'
    et="holgado" if r<0.7 else ("ajustado" if r<=1 else "EXCEDIDO")
    return f'<div style="margin:10px 0"><div style="display:flex;justify-content:space-between;font-size:13px"><b>{t}</b><span>{c:.0f}/{p:.0f}V · {r*100:.0f}%</span></div><div style="background:#e5e7eb;border-radius:6px;height:14px;overflow:hidden"><div style="width:{pct:.0f}%;height:100%;background:{col}"></div></div><div style="font-size:12px;color:#475569">{t}: {c:.0f} V de {p:.0f} V permitidos → {et}</div></div>'
def tablero_html(res):
    if res is None: return "<p>⚠️ Sin datos</p>"
    bc='#16a34a' if res['cumple'] else '#dc2626'; bd='🟢 CUMPLE IEEE 80-2013' if res['cumple'] else '🔴 NO CUMPLE IEEE 80-2013'
    return f'<div style="font-family:system-ui"><div style="background:{bc};color:#fff;border-radius:10px;padding:10px 14px;font-weight:700;margin-bottom:10px">{bd}<span style="float:right;font-weight:400;font-size:13px">{res["metodo"]}</span></div><div style="display:flex;gap:10px;flex-wrap:wrap">{_card("R malla",f"{res['R']:.3f}","Ω")}{_card("GPR (potencial de malla)",f"{res['GPR']:.0f}","V")}{_card("Cond.",res["n_cond"])}{_card("Jab.",res["n_jab"])}</div>{_barra("Tensión de PASO (Es)",res["Es"],res["Es_perm"])}{_barra("Tensión de CONTACTO (Et)",res["Et"],res["Et_perm"])}</div>'
def recomendaciones(res):
    if res is None: return ""
    if res['cumple']: return "✅ **Conforme.** Puedes optimizar costo (mayor espaciado) y re-verificar."
    r=[]
    if res['Et']>res['Et_perm']: r+=["🔴 **Contacto excede:** densifica perímetro, jabalinas en esquinas, o grava (ρs↑)."]
    if res['Es']>res['Es_perm']: r+=["🔴 **Paso excede:** grava o menor t."]
    if res['GPR']>5000: r+=["⚠️ **GPR alto:** más jabalinas/electrodos profundos o mayor área."]
    r+=["💡 Usa **Optimizar** para buscar configuración que cumple."]; return "\n".join(r)

with gr.Blocks(theme=gr.themes.Soft(),title="SPT Designer") as app:
    gr.Markdown("# ⚡ SPT Designer — Diseño de Mallas de Puesta a Tierra")
    guia=gr.Markdown(stepper())
    with gr.Tabs():
        with gr.Tab("1 📋 Proyecto"):
            gr.Markdown("### Datos del proyecto")
            with gr.Row():
                p_nom=gr.Textbox(label="Nombre",value=PROYECTO['nombre']); p_cli=gr.Textbox(label="Cliente",value=PROYECTO['cliente'])
            with gr.Row():
                p_ubi=gr.Textbox(label="Ubicación",value=PROYECTO['ubicacion']); p_ing=gr.Textbox(label="Ingeniero",value=PROYECTO['ingeniero'])
            gr.Markdown("### Parámetros eléctricos")
            with gr.Row():
                q_I=gr.Number(value=PARAMETROS['I_falla'],label="Corriente de falla monofásica I (A)",info="Corriente que drena a tierra en la falla")
                q_t=gr.Number(value=PARAMETROS['t_falla'],label="Tiempo de despeje t (s)",info="Relé + interruptor; típico 0.2-0.5 s")
                q_peso=gr.Dropdown(['50','70'],value='50',label="Peso del trabajador (kg)",info="IEEE 80: 50 kg (k=0.116) o 70 kg (k=0.157)")
            with gr.Row():
                q_pm=gr.Dropdown(['0.3','0.5','0.6','0.8'],value='0.6',label="Profundidad de malla (m)")
                q_lj=gr.Dropdown(['2.4','3.0'],value='2.4',label="Longitud de jabalina (m)")
                q_mc=gr.Dropdown(list(MATCOND),value=PARAMETROS['material_conductor'],label="Material conductor")
                q_cal=gr.Dropdown(list(CALIBRES),value=PARAMETROS['calibre'],label="Calibre (Ø)")
            gr.Markdown("*La capa protectora (grava) se define en 2 🌍 Suelo — fuente única.*")
            b_datos=gr.Button("💾 Guardar datos",variant="primary"); datos_out=gr.Markdown()
        with gr.Tab("2 🌍 Suelo"):
            gr.Markdown("Elige la **dirección de la medición** y pega tus pares `a rho` (uno por línea).")
            dir_sel=gr.Dropdown(DIRECCIONES,value='N-S',label="Dirección de la medición",info="Orientación del tendido Wenner")
            df_dir=gr.Textbox(label="Pares 'a rho' por línea (m, Ω·m)",lines=5,placeholder="5 120\n10 90\n20 70\n40 55")
            with gr.Row(): b_add=gr.Button("➕ Agregar dirección"); b_clr=gr.Button("🧹 Limpiar")
            dir_msg=gr.Markdown(); df_all=gr.Dataframe(headers=["a (m)","ρ (Ω·m)","dirección"],label="Conjunto acumulado")
            with gr.Row():
                b_suelo=gr.Button("📊 Analizar",variant="primary"); b_usars=gr.Button("📌 Usar suelo medido",variant="primary")
            suelo_out=gr.Markdown(); suelo_tbl=gr.Dataframe(); suelo_plot=gr.Plot()
            with gr.Row():
                mat=gr.Dropdown(list(MATERIALES),value='Grava húmeda',label="Capa protectora")
                hs_sel=gr.Dropdown(['0.10','0.15','0.20'],value='0.10',label="Espesor hs (m)")
                hs_custom=gr.Number(label="hs personalizado (m, opcional)",info="Si >0 reemplaza al desplegable")
            b_capa=gr.Button("✅ Aplicar capa"); capa_out=gr.Markdown()
        with gr.Tab("3 📐 Geometría"):
            with gr.Accordion("Forma predefinida",open=True):
                with gr.Row(): gx=gr.Number(value=40,label="X (m)"); gy=gr.Number(value=40,label="Y (m)"); gsp=gr.Number(value=5,label="Espaciado (m)")
                b_gen=gr.Button("🎨 Generar",variant="primary")
            with gr.Accordion("🖊️ Editor por coordenadas",open=False):
                gr.Markdown("Agregar conductor (x1,y1 → x2,y2)")
                with gr.Row(): ex1=gr.Number(label="x1"); ey1=gr.Number(label="y1"); ex2=gr.Number(label="x2"); ey2=gr.Number(label="y2")
                b_addc=gr.Button("➕ Conductor")
                with gr.Row(): ejx=gr.Number(label="x"); ejy=gr.Number(label="y")
                b_addj=gr.Button("➕ Jabalina")
                gr.Markdown("Eliminar por **id** (mira la tabla)")
                with gr.Row(): eic=gr.Number(label="id conductor"); eij=gr.Number(label="id jabalina")
                with gr.Row(): b_delc=gr.Button("➖ Conductor"); b_delj=gr.Button("➖ Jabalina")
                tc=gr.Dataframe(label="Conductores [id,x1,y1,x2,y2]"); tj=gr.Dataframe(label="Jabalinas [id,x,y]")
            with gr.Accordion("Importar CAD/REVIT (DXF)",open=False):
                with gr.Row(): cad_file=gr.File(); cad_mode=gr.Radio(['auto','m','mm'],value='auto',label="Unidades")
                with gr.Row(): b_cad=gr.Button("📥 Importar"); b_cadj=gr.Button("🧪 Cargar ejemplo de prueba")
            geo_out=gr.Markdown()
            with gr.Row(): geo2=gr.Plot(); geo3=gr.Plot()
        with gr.Tab("4 ⚡ Análisis"):
            with gr.Row():
                aI=gr.Number(value=8000,label="Corriente de falla I (A)"); aT=gr.Number(value=0.3,label="Tiempo de despeje t (s)")
            b_calc=gr.Button("🚀 Ejecutar BEM",variant="primary")
            dash=gr.HTML()
            with gr.Accordion("Tabla numérica",open=False): tabla_num=gr.Markdown()
            heat=gr.Image(); recs=gr.Markdown()
        with gr.Tab("5 📈 Perfiles"):
            with gr.Row(): pI=gr.Number(value=8000,label="I (A)"); pT=gr.Number(value=0.3,label="t (s)")
            pcut=gr.Radio(['diagonal','horizontal','vertical','custom'],value='diagonal',label="Corte")
            with gr.Row(): qx1=gr.Number(value=0,label="x1"); qy1=gr.Number(value=0,label="y1"); qx2=gr.Number(value=40,label="x2"); qy2=gr.Number(value=40,label="y2")
            b_perf=gr.Button("📈 Trazar",variant="primary"); perf_out=gr.Markdown()
            pf1=gr.Plot(); pf2=gr.Plot(); pf3=gr.Plot()
        with gr.Tab("6 🎯 Optimizar"):
            with gr.Row(): oI=gr.Number(value=8000,label="I (A)"); oT=gr.Number(value=0.3,label="t (s)")
            b_opt=gr.Button("🎯 Optimizar",variant="primary"); opt_out=gr.Markdown()
        with gr.Tab("7 📄 Informe"):
            with gr.Row():
                b_inf=gr.Button("📥 Generar informe Word",variant="primary"); b_exp=gr.Button("💾 Exportar proyecto (JSON)")
            with gr.Row():
                b_imp=gr.Button("📂 Importar proyecto"); imp_file=gr.File()
            pers_out=gr.Markdown(); json_out=gr.File(); inf_out=gr.File()
        with gr.Tab("8 🔬 Precisión"):
            gr.Markdown("Corriente efectiva (X/R), segmentación fina y tensión transferida a puntos externos.")
            with gr.Row():
                x_xr=gr.Dropdown(['5','10','15'],value='10',label="Relación X/R del sistema",info="Transmisión≈15, distribución≈10, baja≈5")
                x_ml=gr.Number(value=5,label="Long. máx de tramo (m)",info="El BEM subdivide tramos mayores")
            x_pts=gr.Textbox(label="Puntos externos x,y (uno por línea)",placeholder="50,0\n60,10",info="Cercas/postes: tensión transferida")
            b_prec=gr.Button("🔬 Calcular preciso",variant="primary"); prec_out=gr.Markdown()

    b_datos.click(aplicar_datos,inputs=[p_nom,p_cli,p_ubi,p_ing,q_I,q_t,q_peso,q_pm,q_lj,q_mc,q_cal],outputs=datos_out)
    b_add.click(agregar_direccion,inputs=[dir_sel,df_dir],outputs=[df_all,dir_msg,df_dir])
    b_clr.click(limpiar_conjunto,outputs=[df_all,dir_msg])
    b_suelo.click(lambda: _analizar_rows(estado['suelo_rows']),outputs=[suelo_out,suelo_tbl,suelo_plot])
    b_usars.click(aplicar_suelo_medido,outputs=[suelo_out,guia])
    b_capa.click(aplicar_capa,inputs=[mat,hs_sel,hs_custom],outputs=capa_out)
    def on_gen(x,y,s):
        try: c,j=crear_malla_rect(x,y,s)
        except ErrorValidacion as e: return geo_render(f"⚠️ {e}")+(stepper(),)
        estado['conductores']=c; estado['jabalinas']=j; estado['dirty']=True
        return geo_render(f"✅ Malla {x}×{y} generada")+(stepper(),)
    b_gen.click(on_gen,inputs=[gx,gy,gsp],outputs=[geo_out,tc,tj,geo2,geo3,guia])
    b_addc.click(lambda a,b,c,d: editor_add_cond(a,b,c,d)+(stepper(),),inputs=[ex1,ey1,ex2,ey2],outputs=[geo_out,tc,tj,geo2,geo3,guia])
    b_addj.click(lambda x,y: editor_add_jab(x,y)+(stepper(),),inputs=[ejx,ejy],outputs=[geo_out,tc,tj,geo2,geo3,guia])
    b_delc.click(lambda i: editor_del_cond(i)+(stepper(),),inputs=[eic],outputs=[geo_out,tc,tj,geo2,geo3,guia])
    b_delj.click(lambda i: editor_del_jab(i)+(stepper(),),inputs=[eij],outputs=[geo_out,tc,tj,geo2,geo3,guia])
    def on_cad(f,m):
        if f is None: return geo_render("⚠️ sube DXF")+(stepper(),)
        msg,c,j=importar_dxf(open(f.name).read(),m); estado['conductores']=c; estado['jabalinas']=j; estado['dirty']=True
        return geo_render(msg)+(stepper(),)
    b_cad.click(on_cad,inputs=[cad_file,cad_mode],outputs=[geo_out,tc,tj,geo2,geo3,guia])
    def on_cadj():
        msg,c,j=importar_dxf(ejemplo_sucio(),'auto'); estado['conductores']=c; estado['jabalinas']=j; estado['dirty']=True
        return geo_render(msg)+(stepper(),)
    b_cadj.click(on_cadj,outputs=[geo_out,tc,tj,geo2,geo3,guia])
    def on_calc(I,t):
        try:
            res=calcular(estado['conductores'],estado['jabalinas'],I,t)
            if res is None: return "<p>⚠️ genera malla en 3 📐</p>","",None,"",stepper()
            estado['ultimo_res']=res
            guardar_graficas(estado['conductores'],estado['jabalinas'],res)
            num=f"| Par | Valor |\n|---|---|\n| R | {res['R']:.3f} Ω |\n| GPR | {res['GPR']:.0f} V |\n| Es | {res['Es']:.0f}/{res['Es_perm']:.0f} V |\n| Et | {res['Et']:.0f}/{res['Et_perm']:.0f} V |\n| Cs | {res['Cs']:.3f} |"
            return tablero_html(res),num,ruta('mapa_calor.png'),recomendaciones(res),stepper()
        except Exception as e:
            import traceback; tb=traceback.format_exc(); print(tb)
            return f"<p>🔴 ERROR: {type(e).__name__}: {e}</p>","",None,f"```\n{tb[-600:]}\n```",stepper()
    b_calc.click(on_calc,inputs=[aI,aT],outputs=[dash,tabla_num,heat,recs,guia])
    b_perf.click(perfiles,inputs=[pI,pT,pcut,qx1,qy1,qx2,qy2],outputs=[perf_out,pf1,pf2,pf3])
    def on_opt(I,t):
        if not estado['conductores']: return "⚠️ genera malla en 3 📐"
        xs=[c[0] for c in estado['conductores']]+[c[2] for c in estado['conductores']]
        ys=[c[1] for c in estado['conductores']]+[c[3] for c in estado['conductores']]
        return optimizar(max(xs)-min(xs),max(ys)-min(ys),I,t,SUELO['rho1'])
    b_opt.click(on_opt,inputs=[oI,oT],outputs=opt_out)
    def on_prec(XoR,ml,pts_txt):
        pts=[]
        for l in pts_txt.split('\n'):
            p=l.replace(',',' ').split()
            if len(p)>=2:
                try: pts.append((float(p[0]),float(p[1])))
                except Exception: pass
        r=calcular_preciso(estado['conductores'],PARAMETROS['I_falla'],PARAMETROS['t_falla'],SUELO['rho1'],float(XoR),float(ml),pts or None)
        txt=f"**Df={r['Df']:.3f} → I efectiva={r['Ieff']:.0f} A** · tramos={r['n_seg']} · R={r['R']:.3f} Ω · GPR={r['GPR']:.0f} V"
        if r.get('trans'): txt+="<br>**Transferida:** "+"; ".join(f"({p[0]:.0f},{p[1]:.0f})={v:.0f} V" for p,v in r['trans'])
        return txt
    b_prec.click(on_prec,inputs=[x_xr,x_ml,x_pts],outputs=prec_out)
    def on_inf():
        if not estado['conductores']: return None
        res=calcular(estado['conductores'],estado['jabalinas'])
        guardar_graficas(estado['conductores'],estado['jabalinas'],res)
        return crear_informe(estado['conductores'],estado['jabalinas'],res)
    b_inf.click(on_inf,outputs=inf_out)
    b_exp.click(exportar_proyecto,outputs=[pers_out,json_out])
    b_imp.click(importar_proyecto,inputs=imp_file,outputs=[pers_out,json_out])
    ci,cj=crear_malla_rect(40,40,5); estado['conductores']=ci; estado['jabalinas']=cj

# ---- arranque (Render / Hugging Face / local) ----
app.launch(server_name="0.0.0.0", server_port=int(os.environ.get("PORT", os.environ.get("GRADIO_SERVER_PORT","7860"))))
