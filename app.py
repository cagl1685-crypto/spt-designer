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
PARAMETROS={'I_falla':8000.0,'t_falla':0.3,'peso':50.0,'Sf':1.0,'modo_exposicion':'ocupacional_50',
 'rho_grava':3000.0,'h_grava':0.10,'prof_malla':0.6,
 'long_jabalina':2.4,'radio_conductor':0.0053,'k_material':202.8,'material_conductor':'Cobre blando','calibre':'2/0 AWG'}
MODOS_EXPOSICION={'Ocupacional 50 kg (IEEE 80)':'ocupacional_50','Ocupacional 70 kg (IEEE 80)':'ocupacional_70',
 'Público en general (RETIE Tabla 3.12.1.a / IEC 60479-1)':'publico_retie'}
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

def aplicar_datos(nom,cli,ubi,ing, I,t, modo_exp_label,sf,pm,lj, mat_cond, calibre):
    diam=CALIBRES[calibre]; radio_m=diam/2000.0; k=MATCOND[mat_cond]['k']
    modo_exp=MODOS_EXPOSICION.get(modo_exp_label,'ocupacional_50')
    PROYECTO.update({'nombre':nom,'cliente':cli,'ubicacion':ubi,'ingeniero':ing})
    PARAMETROS.update({'I_falla':float(I),'t_falla':float(t),'modo_exposicion':modo_exp,
        'Sf':max(0.0,min(1.0,sanitizar(sf,1.0))),'prof_malla':float(pm),
        'long_jabalina':float(lj),'radio_conductor':radio_m,'k_material':k,'material_conductor':mat_cond,'calibre':calibre})
    estado['dirty']=True
    return (f"✅ Guardado: {nom} · I={float(I):.0f} A · t={float(t):.2f} s · {modo_exp_label} · Sf={PARAMETROS['Sf']:.2f} · {mat_cond} {calibre} (Ø={diam:.2f} mm)",
            float(I), float(t))
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
def _ruta_archivo(f):
    """AUDITORÍA (bug confirmado): Gradio >=4 con gr.File(type='filepath')
    (por defecto desde Gradio 4, y confirmado en Gradio 6.27 instalado) entrega
    un STRING con la ruta, no un objeto con atributo .name como en Gradio 3.
    El código original hacía f.name en ambos importadores y truena con
    AttributeError en cualquier despliegue con Gradio moderno. Esta función
    acepta ambas formas."""
    if f is None: return None
    return f.name if hasattr(f, 'name') else f

def exportar_proyecto():
    data={'PROYECTO':PROYECTO,'PARAMETROS':PARAMETROS,'SUELO':SUELO,'conductores':estado['conductores'],'jabalinas':estado['jabalinas'],'capa':estado['capa']}
    open(ruta('proyecto_spt.json'),'w').write(json.dumps(data)); estado['dirty']=False
    return "✅ Proyecto exportado (proyecto_spt.json)", ruta('proyecto_spt.json')
def importar_proyecto(f):
    if f is None: return "⚠️ sube el JSON del proyecto", None
    d=json.load(open(_ruta_archivo(f)))
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
def explicar_alertas(alerts):
    if not alerts: return "**Control de calidad:** mediciones consistentes, sin valores atípicos."
    L=["**Control de calidad de mediciones:**"]
    for a in alerts:
        if a['tipo']=='error_medida':
            L.append(f"⚠️ Posible **error de medición** en {a['dir']}, a={a['a']} m (ρ={a['rho']:.0f} Ω·m, desvío {a['desvio']:+.0f}%). Verifique/repita ese punto.")
        elif a['tipo']=='hallazgo':
            L.append(f"🔍 **Hallazgo real** en {a['dir']} (desvío {a['desvio']:+.0f}%): estratificación/anisotropía lateral; considérelo en el diseño.")
        else:
            L.append(f"🧭 **Anisotropía:** el promedio de {a['dir']} difiere {a['desvio']:+.0f}% del global; adopte el caso conservador.")
    return "\n".join(L)
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
        ax.plot(amean,rmean,'ko-',lw=2,label='Promedio medido')
        aa=np.logspace(np.log10(min(A)*0.8),np.log10(max(A)*1.2),60)
        ax.plot(aa,wenner_2capas(aa,r1,r2,h),'r-',lw=2,label='Modelo 2 capas (teórico)')
        ax.set_xscale('log'); ax.set_yscale('log'); ax.grid(alpha=0.3,which='both'); ax.legend(); plt.close(fig)
        return f"**2 capas:** ρ1={r1:.1f}, ρ2={r2:.1f}, h={h:.2f} m.\n\n"+explicar_alertas(estado['alerts'])+"\n\nPulsa **📌 Usar suelo medido**.",stats,fig
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

def _segmentar_para_bem(cond,d,max_len=3.0):
    """AUDITORÍA: el BEM principal (bem_analizar) usaba los conductores tal como
    llegan de la geometría (p.ej. 14 m de un tirón en una malla con espaciado
    14 m), con UN solo valor de corriente por segmento. Eso degrada Rg/GPR y,
    sobre todo, subvalúa Et al solo poder muestrear el punto medio de un tramo
    largo. Confirmado con el Ejemplo Grid 1 de IEEE 80 Anexo H: al segmentar a
    2 m, Rg pasó de 0.987 a 0.999 Ω (la referencia real es 1.00 Ω). Se aplica
    aquí el mismo segmentador que ya existía (antes solo se usaba en la
    pestaña 8 Precisión) para que el análisis principal tenga la misma base.
    Ahora devuelve segmentos 3D explícitos (x1,y1,z1,x2,y2,z2) -- antes
    devolvía (x1,y1,x2,y2) asumiendo z=-d fijo para todo, lo que impedía
    representar conductores no-horizontales (jabalinas)."""
    out=[]
    for (x1,y1,x2,y2) in cond:
        L=math.hypot(x2-x1,y2-y1)
        if L<1e-9: continue
        n=max(1,int(math.ceil(L/max_len)))
        for k in range(n):
            a=k/n; b=(k+1)/n
            out.append((x1+(x2-x1)*a,y1+(y2-y1)*a,-d,x1+(x2-x1)*b,y1+(y2-y1)*b,-d))
    return out

def _jabalinas_a_segmentos_3d(jab,d,long_jabalina,max_len=3.0):
    """AUDITORÍA (hallazgo grave, confirmado con Grid 1 vs Grid 2 de IEEE 80
    Anexo H dando resultados IDÉNTICOS con y sin jabalinas): 'jab' nunca
    entraba al BEM -- solo se dibujaba y se contaba. Cada jabalina se modela
    aquí como un segmento VERTICAL desde donde se conecta a la malla (z=-d)
    hasta la punta (z=-d-long_jabalina), segmentado igual que los conductores
    horizontales para la misma precisión."""
    out=[]
    if long_jabalina<=0: return out
    n=max(1,int(math.ceil(long_jabalina/max_len)))
    for (x,y) in jab:
        for k in range(n):
            z1=-d-(long_jabalina*k/n); z2=-d-(long_jabalina*(k+1)/n)
            out.append((x,y,z1,x,y,z2))
    return out

def calcular_tension_paso(GX,GY,Vs,paso=1.0):
    """AUDITORÍA (bug crítico confirmado): la versión anterior calculaba Es con
    np.diff(Vs) entre celdas ADYACENTES de la grilla de graficado, es decir,
    medía la diferencia de potencial sobre la distancia 'res_sup', no sobre
    1 metro como exige IEEE 80. Prueba de estrés: variar res_sup de 2.0 a 0.25 m
    (sin cambiar nada físico) hacía caer Es de 560.6 V a 136.6 V. Aquí se
    interpola el potencial en pares de puntos separados EXACTAMENTE 1 m (o el
    valor de 'paso' que se pida), así el resultado deja de depender de la
    resolución de graficado.
    AUDITORÍA 2 (encontrado al re-verificar Grid 1 con el motor corregido): solo
    se revisaban los 4 pasos cardinales (±x, ±y). IEEE 80 define S1 justo sobre
    la diagonal desde la esquina, y en el Anexo H eso da un valor S1 más alto
    que el que esta función encontraba (67.7 V vs. 87-96 V de referencia).
    Se agregan las 4 direcciones diagonales, escaladas para que el paso siga
    midiendo exactamente 1 m de magnitud (no 1.41 m).
    """
    try:
        from scipy.interpolate import RegularGridInterpolator
        xs=GX[0,:]; ys=GY[:,0]
        interp=RegularGridInterpolator((ys,xs),Vs,bounds_error=False,fill_value=None)
        pts0=np.stack([GY.ravel(),GX.ravel()],axis=-1)
        v0=interp(pts0)
        max_es=0.0
        diag=paso/math.sqrt(2.0)
        direcciones=[(paso,0.0),(-paso,0.0),(0.0,paso),(0.0,-paso),
                     (diag,diag),(diag,-diag),(-diag,diag),(-diag,-diag)]
        for dx,dy in direcciones:
            pts1=np.stack([(GY+dy).ravel(),(GX+dx).ravel()],axis=-1)
            v1=interp(pts1)
            dif=np.abs(v0-v1); dif=dif[np.isfinite(dif)]
            if dif.size: max_es=max(max_es,float(dif.max()))
        return max_es
    except Exception:
        return float(max(np.abs(np.diff(Vs,axis=1)).max(),np.abs(np.diff(Vs,axis=0)).max()))

def bem_analizar(cond,rho,I,d=0.6,radio=0.0053,res_sup=1.0,auto_segmentar=True,max_len_seg=3.0,jab=None,long_jabalina=0.0,radio_jabalina=0.0079):
    # AUDITORÍA (hallazgo grave, confirmado: Grid 1 y Grid 2 de IEEE 80 Anexo H
    # daban resultados IDÉNTICOS con y sin jabalinas). Las jabalinas nunca
    # entraban a la matriz del BEM. Ahora se modelan como segmentos verticales
    # (_jabalinas_a_segmentos_3d) y se resuelven junto con la malla en una sola
    # matriz. El término de auto-impedancia de imagen se generalizó (ver
    # verificar_dwight.py: reproduce exacto la fórmula horizontal original, y
    # una jabalina aislada converge a la fórmula de Dwight al segmentarla).
    cond_grid_3d = _segmentar_para_bem(cond,d,max_len_seg) if auto_segmentar else [(c[0],c[1],-d,c[2],c[3],-d) for c in cond]
    cond_jab_3d = _jabalinas_a_segmentos_3d(jab or [],d,long_jabalina,max_len_seg) if long_jabalina>0 else []
    cond_calc = cond_grid_3d + cond_jab_3d
    cond_et = cond_grid_3d  # la tensión de contacto se muestrea sobre los conductores de la malla, no sobre el eje vertical de las jabalinas
    N=len(cond_calc)
    A=np.array([[s[0],s[1],s[2]] for s in cond_calc],float); B=np.array([[s[3],s[4],s[5]] for s in cond_calc],float)
    radios=np.array([radio_jabalina if i>=len(cond_grid_3d) else radio for i in range(N)])
    Lseg=np.linalg.norm(B-A,axis=1); mid=(A+B)/2
    Ai=A.copy(); Ai[:,2]=-A[:,2]; Bi=B.copy(); Bi[:,2]=-B[:,2]  # imagen: reflejo z->-z, válido para cualquier orientación
    coeff=rho/(4*np.pi); R=np.zeros((N,N))
    for j in range(N):
        Lj=Lseg[j]
        r1=np.linalg.norm(mid-A[j],axis=1); r2=np.linalg.norm(mid-B[j],axis=1)
        real=np.log((r1+r2+Lj)/np.maximum(r1+r2-Lj,1e-9))
        r1i=np.linalg.norm(mid-Ai[j],axis=1); r2i=np.linalg.norm(mid-Bi[j],axis=1)
        imag=np.log((r1i+r2i+Lj)/np.maximum(r1i+r2i-Lj,1e-9))
        col=coeff/Lj*(real+imag)
        rs=np.sqrt((Lj/2)**2+radios[j]**2); sr=np.log((2*rs+Lj)/max(2*rs-Lj,1e-9))
        # auto-término de imagen GENERALIZADO: distancia real del punto medio
        # propio a los dos extremos de SU PROPIA imagen (antes: sqrt((L/2)²+(2d)²),
        # que asume imagen paralela a distancia perpendicular fija 2d -- solo
        # válido para conductor horizontal; para una jabalina vertical la
        # imagen queda alineada en la misma recta, no paralela).
        r1si=np.linalg.norm(mid[j]-Ai[j]); r2si=np.linalg.norm(mid[j]-Bi[j])
        si=np.log((r1si+r2si+Lj)/max(r1si+r2si-Lj,1e-9))
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
    Es=calcular_tension_paso(GX,GY,Vs,paso=1.0); Et=0.0
    for c in cond_et:
        mx=(c[0]+c[3])/2; my=(c[1]+c[4])/2; dx=c[3]-c[0]; dy=c[4]-c[1]; Lc=math.hypot(dx,dy)
        if Lc<1e-9: continue
        Et=max(Et,GPR-_interp(GX,GY,Vs,mx-dy/Lc,my+dx/Lc))
    return {'R_grid':Rg,'GPR':GPR,'Es':Es,'Et':float(Et),'GX':GX,'GY':GY,'Vsurf':Vs,'n_seg_bem':N,
            'A_seg':A,'B_seg':B,'Lseg':Lseg,'Iseg':Iseg,'coeff':coeff}

# =============================================================================
# MOTOR DE SUELO DE 2 CAPAS (bem_analizar_2capas)
# AUDITORÍA: derivado y verificado en sesión aparte antes de tocar este
# archivo (ver BRIEF_sesion_2capas.md y los 3 scripts de verificación):
#   - Serie de imágenes construida MECÁNICAMENTE (reflexión repetida, no de
#     memoria) y verificada exacta contra wenner_2capas() ya existente
#     (diferencias de 10^-15, puro redondeo) en 4 combinaciones de
#     resistividad y 6 espaciados.
#   - Con K=0 (rho2=rho1) reproduce EXACTO bem_analizar() de 1 capa
#     (verificado con interacción real entre 3 segmentos, diferencia 0.0).
#   - Validado contra Grid 3 del Anexo H de IEEE 80 (malla + jabalinas + suelo
#     rho1=300/rho2=100/h=6.096m): con jabalinas que NO cruzan la interfaz,
#     comportamiento físicamente correcto y del orden de magnitud esperado.
#
# LÍMITE CONOCIDO Y CON RESGUARDO EXPLÍCITO: la derivación asume que toda
# fuente (conductor de malla o jabalina) está DENTRO de la capa 1. Si una
# jabalina se extiende más allá de la interfaz (profundidad + long_jabalina >
# h), la parte que ya está en la capa 2 NO está correctamente modelada --
# confirmado con Grid 3 tal cual (jabalinas de 7.5m, h=6.096m -> penetran
# 1.9m en la capa 2): Rg=0.636Ω contra 0.97-1.4Ω de referencia, muy fuera de
# rango. Con una jabalina corta que sí se queda dentro de la capa 1, el
# resultado (1.163Ω) es físicamente razonable. Por eso esta función NUNCA se
# usa si esa condición se cumple -- ver el chequeo en calcular().
# =============================================================================
def _serie_log_2capas(P,A,B,h,K,N_terms):
    """log-término completo (serie de imágenes de 2 capas) de un segmento
    A-B, evaluado en punto(s) P. Vectorizado sobre P (P puede ser un array
    Nx3 de puntos, o un solo punto de 3)."""
    L=np.linalg.norm(B-A)
    def logterm(Ai,Bi):
        r1=np.linalg.norm(P-Ai,axis=-1); r2=np.linalg.norm(P-Bi,axis=-1)
        return np.log((r1+r2+L)/np.maximum(r1+r2-L,1e-9))
    total=logterm(A,B)
    total=total+logterm(np.array([A[0],A[1],-A[2]]),np.array([B[0],B[1],-B[2]]))
    for n in range(1,N_terms+1):
        Kn=K**n
        if abs(Kn)<1e-12: break
        for signo_refl in (1,-1):
            for signo_n in (1,-1):
                Az=A[2]*signo_refl+signo_n*2*n*h; Bz=B[2]*signo_refl+signo_n*2*n*h
                total=total+Kn*logterm(np.array([A[0],A[1],Az]),np.array([B[0],B[1],Bz]))
    return total

def _n_terms_adaptativo(K,tol=1e-6,maximo=200):
    if abs(K)<1e-9: return 1
    n=1
    while abs(K)**n>tol and n<maximo: n+=1
    return n

def jabalina_cruza_interfaz(d,long_jabalina,h):
    """True si alguna jabalina, con la profundidad y longitud dadas,
    terminaría más allá de la interfaz de capas h -- caso en el que
    bem_analizar_2capas NO es válido (ver nota arriba)."""
    return long_jabalina>0 and (d+long_jabalina)>h

def bem_analizar_2capas(cond,rho1,rho2,h,I,d=0.6,radio=0.0053,res_sup=1.0,max_len_seg=3.0,jab=None,long_jabalina=0.0,radio_jabalina=0.0079):
    K=(rho2-rho1)/(rho2+rho1); N_terms=_n_terms_adaptativo(K)
    cond_grid_3d=_segmentar_para_bem(cond,d,max_len_seg)
    cond_jab_3d=_jabalinas_a_segmentos_3d(jab or [],d,long_jabalina,max_len_seg) if long_jabalina>0 else []
    cond_calc=cond_grid_3d+cond_jab_3d; cond_et=cond_grid_3d
    N=len(cond_calc)
    A=np.array([[s[0],s[1],s[2]] for s in cond_calc],float); B=np.array([[s[3],s[4],s[5]] for s in cond_calc],float)
    radios=np.array([radio_jabalina if i>=len(cond_grid_3d) else radio for i in range(N)])
    Lseg=np.linalg.norm(B-A,axis=1); mid=(A+B)/2
    coeff=rho1/(4*np.pi); R=np.zeros((N,N))
    for j in range(N):
        Lj=Lseg[j]
        total=_serie_log_2capas(mid,A[j],B[j],h,K,N_terms)
        col=coeff/Lj*total
        rs=math.sqrt((Lj/2)**2+radios[j]**2); sr=math.log((2*rs+Lj)/max(2*rs-Lj,1e-9))
        directo_sin_offset=math.log((2*(Lj/2)+Lj)/max(2*(Lj/2)-Lj,1e-9))
        total_self=_serie_log_2capas(mid[j],A[j],B[j],h,K,N_terms)
        col[j]=coeff/Lj*(sr+(total_self-directo_sin_offset)); R[:,j]=col
    try: x=np.linalg.solve(R,np.ones(N))
    except np.linalg.LinAlgError: x=np.linalg.pinv(R)@np.ones(N)
    S=x.sum()
    if not es_finito(S) or S<=0: return None
    Rg=1.0/S; GPR=I*Rg; Iseg=GPR*x
    xs=[c[0] for c in cond]+[c[2] for c in cond]; ys=[c[1] for c in cond]+[c[3] for c in cond]
    gx=np.arange(min(xs)-5,max(xs)+5,res_sup); gy=np.arange(min(ys)-5,max(ys)+5,res_sup)
    GX,GY=np.meshgrid(gx,gy)
    P=np.stack([GX.ravel(),GY.ravel(),np.zeros(GX.size)],axis=-1)
    V=np.zeros(P.shape[0])
    for j in range(N):
        Lj=Lseg[j]
        V=V+(coeff/Lj)*Iseg[j]*_serie_log_2capas(P,A[j],B[j],h,K,N_terms)
    Vs=V.reshape(GX.shape)
    Es=calcular_tension_paso(GX,GY,Vs,paso=1.0); Et=0.0
    for c in cond_et:
        mx=(c[0]+c[3])/2; my=(c[1]+c[4])/2; dx=c[3]-c[0]; dy=c[4]-c[1]; Lc=math.hypot(dx,dy)
        if Lc<1e-9: continue
        Et=max(Et,GPR-_interp(GX,GY,Vs,mx-dy/Lc,my+dx/Lc))
    return {'R_grid':Rg,'GPR':GPR,'Es':Es,'Et':float(Et),'GX':GX,'GY':GY,'Vsurf':Vs,'n_seg_bem':N,
            'A_seg':A,'B_seg':B,'Lseg':Lseg,'Iseg':Iseg,'coeff':coeff,'K':K,'N_terms':N_terms}

def evaluar_potencial_puntos(bem_result,puntos_xy):
    """AUDITORÍA (consolidación #2): antes existía un bem_preciso() aparte para
    poder evaluar el potencial en puntos externos (cercas, postes) — con su
    propio ensamblaje de matriz, que nunca recibió los arreglos de Es/Et ni la
    segmentación automática. Esta función reutiliza la corriente por segmento
    (Iseg) que YA resolvió bem_analizar, evitando mantener dos motores. Se
    puede llamar con el resultado de CUALQUIER corrida de bem_analizar."""
    A,B,Lseg,Iseg,coeff = bem_result['A_seg'],bem_result['B_seg'],bem_result['Lseg'],bem_result['Iseg'],bem_result['coeff']
    pts=np.array([[p[0],p[1],0.0] for p in puntos_xy],float)
    V=np.zeros(len(pts))
    for j in range(len(Lseg)):
        Lj=Lseg[j]
        r1=np.linalg.norm(pts-A[j],axis=1); r2=np.linalg.norm(pts-B[j],axis=1)
        V+=(coeff/Lj)*Iseg[j]*2*np.log((r1+r2+Lj)/np.maximum(r1+r2-Lj,1e-9))
    return V
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

# =============================================================================
# AUDITORÍA — RETIE Art. 3.12.1.g y Tabla 3.12.1.a (Resolución 40117 de 2024):
# "Los valores máximos permisibles de tensión de contacto y de paso deben ser
#  calculados siguiendo la metodología de IEEE 80. En caso de no tener capa
#  superficial, no se deben superar los valores de la Tabla 3.12.1.a" — que da
# la tensión de CONTACTO máxima para dos poblaciones: pública (IEC 60479-1,
# curva C1, 5% probabilidad de fibrilación) y ocupacional (IEEE, 50 kg).
#
# La columna "ocupacional" de la tabla es exactamente 1000 x 0.116/√t (se
# verificó punto por punto: a 1s da 116V, a 0.5s da 164V, a 0.3s da 211V...
# coincide con Dalziel a Cs=1, sin resistencia adicional de pie). La columna
# "público" NO tiene fórmula cerrada (IEC 60479-1 no es una simple k/√t) así
# que aquí se interpola la tabla en log(t) para obtener Ib_publico(t) = V/1000,
# y se generaliza a cualquier capa superficial con la MISMA estructura de
# Dalziel que usa el resto del programa: Etouch=(1000+1.5·Cs·ρs)·Ib(t).
#
# LÍMITE HONESTO: RETIE solo tabula tensión de CONTACTO para público en
# general. No publica una tabla de tensión de PASO equivalente. La expresión
# de step aquí (1000+6·Cs·ρs)·Ib(t) es una EXTENSIÓN razonada (misma Ib,
# misma estructura geométrica que usa IEEE 80 para pasar de touch a step)y
# NO un valor dado explícitamente por la norma. Debe tratarse como estimación
# de ingeniería, no como cita textual de RETIE.
# =============================================================================
RETIE_TABLA_CONTACTO_PUBLICO = [(2.0,50.0),(1.0,55.0),(0.7,70.0),(0.5,80.0),(0.4,130.0),
                                 (0.3,200.0),(0.2,270.0),(0.15,300.0),(0.1,320.0),(0.05,345.0)]

def _ib_publico_retie(t):
    ts=[p[0] for p in RETIE_TABLA_CONTACTO_PUBLICO]; vs=[p[1] for p in RETIE_TABLA_CONTACTO_PUBLICO]
    t=max(min(t,ts[0]),ts[-1])
    order=sorted(range(len(ts)),key=lambda i:ts[i])
    xs=[math.log(ts[i]) for i in order]; ys=[vs[i] for i in order]
    v=float(np.interp(math.log(t),xs,ys))
    return v/1000.0

def limites_retie_publico(hs,rho_s,rho1,t):
    Ib=_ib_publico_retie(t)
    if hs is None or hs<=0 or rho_s is None:
        # AUDITORÍA: caso "sin capa superficial" tal como lo define RETIE
        # literalmente (Art. 3.12.1.g): la Tabla 3.12.1.a ya es 1000*Ib(t), sin
        # sumar un término adicional de resistividad nativa. Generalizar con
        # Cs=ρ1/ρs (que da la fórmula de Sverak cuando hs=0) sumaría un término
        # 1.5*ρ1 que la norma NO incluye en su tabla -- por eso este caso se
        # trata aparte en vez de dejar que caiga en la fórmula general de abajo.
        return 1.0,1000.0*Ib,1000.0*Ib
    Cs=max(0.0,min(1.0,1.0-0.09*(1.0-rho1/rho_s)/(2.0*hs+0.09)))
    return Cs,(1000+6*Cs*rho_s)*Ib,(1000+1.5*Cs*rho_s)*Ib

def limites_tolerables(hs,rho_s,rho1,t,modo='ocupacional_50'):
    """Despachador único de límites tolerables paso/contacto.
    modo: 'ocupacional_50' | 'ocupacional_70' | 'publico_retie'"""
    if modo=='publico_retie': return limites_retie_publico(hs,rho_s,rho1,t)
    peso=70.0 if modo=='ocupacional_70' else 50.0
    return limites_ieee(hs,rho_s,rho1,t,peso)
def calcular(cond,jab,I=None,t=None,Sf=None,modo_exposicion=None):
    if not cond: return None
    L=sum(math.hypot(c[2]-c[0],c[3]-c[1]) for c in cond) + len(jab)*PARAMETROS.get('long_jabalina',0.0)
    xs=[c[0] for c in cond]+[c[2] for c in cond]; ys=[c[1] for c in cond]+[c[3] for c in cond]
    area=(max(xs)-min(xs))*(max(ys)-min(ys))
    # AUDITORÍA (prueba de estrés): una I negativa (dato mal ingresado) no se
    # rechazaba y producía GPR negativo, que por comparación de signos daba
    # "cumple=True" de forma engañosa. La corriente de falla es una magnitud;
    # se usa su valor absoluto.
    I_total=abs(sanitizar(I or PARAMETROS['I_falla'],1000))
    # AUDITORÍA: factor de división de corriente Sf (IEEE 80 Cl.15 / Anexo C).
    # Antes no existía: se asumía que el 100% de I_falla entra por la malla,
    # lo cual es conservador pero no lo que hace un diseño real (ver Ejemplo 1
    # del Anexo B: de 3180A de falla, solo 1908A -Sf=0.6, Df=1.0- llegan a la
    # malla). Sf=1.0 reproduce el comportamiento anterior (compatibilidad).
    Sf=max(0.0,min(1.0,sanitizar(PARAMETROS.get('Sf',1.0) if Sf is None else Sf,1.0)))
    I_malla=I_total*Sf
    t=min(max(sanitizar(t or PARAMETROS['t_falla'],0.3),0.01),10)
    modo_exposicion=modo_exposicion or PARAMETROS.get('modo_exposicion','ocupacional_50')
    # AUDITORÍA: segmentación automática (ver _segmentar_para_bem). Para que una
    # malla grande no dispare el costo O(N^3) del solve denso, el paso de
    # segmentación se engruesa adaptativamente si haría falta más de ~900
    # segmentos; por encima de ~4000 segmentos estimados se cae al método
    # simplificado en vez de intentar un solve que tardaría minutos/horas.
    # AUDITORÍA (bug propio encontrado en prueba de estrés): _segmentar_para_bem
    # solo DIVIDE segmentos largos, nunca FUSIONA los que ya vienen cortos. La
    # primera versión de este techo adaptativo asumía que "engrosar max_len_seg"
    # siempre bajaba el conteo final, lo cual es falso si la malla ya llega fina
    # de origen (probado con una malla de 300x300m a 5m: 7320 conductores
    # originales, el techo calculó mal un estimado de 900 y terminó resolviendo
    # una matriz de 7320x7320 real -> 48 segundos). Ahora el piso del estimado
    # es siempre len(cond) (nunca se puede bajar de ahí), y si el propio
    # len(cond) ya supera el techo, no se intenta refinar más.
    max_len_seg=3.0; n_original=len(cond)+sum(max(1,math.ceil(PARAMETROS.get('long_jabalina',0.0)/max_len_seg)) for _ in jab)
    if n_original>900:
        auto_seg=False; n_estimado=n_original
    else:
        n_estimado=max(n_original,math.ceil(L/max_len_seg))
        if n_estimado>900: max_len_seg=max(1.0,L/900.0); n_estimado=max(n_original,math.ceil(L/max_len_seg))
        auto_seg=True
    # AUDITORÍA: usar el motor de 2 capas SOLO si el usuario aplicó explícitamente
    # un suelo medido de 2 capas (estado['suelo_aplicado'], botón "📌 Usar suelo
    # medido") y el contraste es real -- no por defecto con los valores de
    # arranque de SUELO. Si alguna jabalina cruzaría la interfaz de capas
    # (fuera del rango donde la fórmula es válida -- ver nota en
    # bem_analizar_2capas), se avisa explícitamente y se cae al motor uniforme
    # (con ρ1, conservador para la malla ya que ρ1 suele ser la capa más
    # resistiva medida en superficie) en vez de dar un número incorrecto.
    usar_2capas = estado.get('suelo_aplicado') and abs(SUELO.get('rho2',SUELO['rho1'])-SUELO['rho1'])>1e-6
    aviso_2capas = None
    long_jab_2capas = PARAMETROS.get('long_jabalina',0.0)
    # AUDITORÍA: en vez de descartar el suelo de 2 capas por completo cuando una
    # jabalina perfora la interfaz, se trunca la jabalina justo en h -- la
    # parte que ya estaría en la capa 2 (donde la fórmula no es válida) se
    # ignora. Esto SUBESTIMA el beneficio real de esa porción (nunca lo
    # sobreestima: conservador por el lado correcto), y deja que la malla y el
    # resto de la jabalina sí se beneficien del suelo real de 2 capas, en vez
    # de perder toda la información de 2 capas de golpe como antes.
    if usar_2capas and jabalina_cruza_interfaz(PARAMETROS['prof_malla'],PARAMETROS.get('long_jabalina',0.0),SUELO['h']):
        long_original=PARAMETROS.get('long_jabalina',0.0)
        long_jab_2capas=max(0.0,SUELO['h']-PARAMETROS['prof_malla'])
        aviso_2capas=(f"ℹ️ Jabalina truncada a {long_jab_2capas:.2f} m (de {long_original:.1f} m reales) para el "
                      f"cálculo de 2 capas: la porción que ya estaría en la capa 2 (h={SUELO['h']:.2f} m) no se "
                      f"contabiliza. Esto es conservador (subestima el beneficio de esa porción), nunca al revés.")
    if usar_2capas:
        bem=bem_analizar_2capas(cond,SUELO['rho1'],SUELO['rho2'],SUELO['h'],I_malla,PARAMETROS['prof_malla'],
                                 PARAMETROS['radio_conductor'],max_len_seg=max_len_seg,
                                 jab=jab,long_jabalina=long_jab_2capas) if n_estimado<=4000 else None
    else:
        bem=bem_analizar(cond,SUELO['rho1'],I_malla,PARAMETROS['prof_malla'],PARAMETROS['radio_conductor'],
                          auto_segmentar=auto_seg,max_len_seg=max_len_seg,
                          jab=jab,long_jabalina=PARAMETROS.get('long_jabalina',0.0)) if n_estimado<=4000 else None
    if bem is None:
        R=div_segura(SUELO['rho1'],4*sqrt_segura(area/math.pi)); GPR=I_malla*R; Es=GPR*0.15; Et=GPR*0.08
        met=f'Simplificado (malla estimada en {n_estimado} segmentos, excede el límite del BEM)'
    else:
        R,GPR,Es,Et=bem['R_grid'],bem['GPR'],bem['Es'],bem['Et']
        met=f"BEM 2 capas ({bem.get('n_seg_bem','?')} seg., K={bem.get('K',0):.2f})" if usar_2capas else f"BEM ({bem.get('n_seg_bem','?')} segmentos)"
    Cs,Ep,Tp=limites_tolerables(PARAMETROS['h_grava'],PARAMETROS['rho_grava'],SUELO['rho1'],t,modo_exposicion)
    k=PARAMETROS.get('k_material',202.8)
    return {'n_cond':len(cond),'n_jab':len(jab),'L':L,'area':area,'R':sanitizar(R),'GPR':sanitizar(GPR),'Cs':Cs,
            'Es':sanitizar(Es),'Es_perm':Ep,'Et':sanitizar(Et),'Et_perm':Tp,'A_min':div_segura(I_total*sqrt_segura(t),k),
            'cumple':bool(Es<=Ep and Et<=Tp),'metodo':met,'bem':bem,'I_malla':I_malla,'Sf':Sf,'modo_exposicion':modo_exposicion,
            'aviso_2capas':aviso_2capas}

# ================= FASE A: PRECISIÓN =================
def factor_decremento(XoR,t,f=60.0):
    if not (XoR>0 and t>0): return 1.0
    Ta=XoR/(2*math.pi*f)
    return math.sqrt(1.0+(Ta/t)*(1.0-math.exp(-2.0*t/Ta)))

# AUDITORÍA (consolidación #2, hallazgo confirmado): segmentar()/bem_preciso()
# implementaban SU PROPIO ensamblaje de matriz BEM, en paralelo al de
# bem_analizar(). Dos motores de física resolviendo lo mismo con código
# distinto es una garantía de que un arreglo futuro (o uno de los ya hechos
# aquí: Es distancia-normalizada, segmentación automática) se aplique en uno
# y se le olvide en el otro -- de hecho ya había pasado: bem_preciso() nunca
# tuvo Es/Et. Se elimina la duplicación: calcular_preciso() ahora llama a
# bem_analizar() (el único motor) y usa evaluar_potencial_puntos() para la
# tensión transferida a puntos externos.
def calcular_preciso(cond,I,t,rho1,XoR=10.0,max_len=5.0,puntos=None,d=0.6,radio=0.0053,Sf=1.0,jab=None,long_jabalina=0.0):
    # AUDITORÍA (punto pendiente #3): esta pestaña aplicaba el factor de
    # decremento Df pero no Sf, quedando inconsistente con la pestaña principal
    # de Análisis desde que se agregó Sf ahí. Ahora aplica ambos: IG=I·Sf·Df.
    Sf=max(0.0,min(1.0,sanitizar(Sf,1.0)))
    Df=factor_decremento(XoR,t); Ieff=I*Sf*Df
    if not cond: return {'Df':Df,'Sf':Sf,'Ieff':Ieff,'n_seg':0,'R':float('nan'),'GPR':float('nan')}
    b=bem_analizar(cond,rho1,Ieff,d=d,radio=radio,auto_segmentar=True,max_len_seg=max_len,jab=jab,long_jabalina=long_jabalina)
    if b is None: return {'Df':Df,'Sf':Sf,'Ieff':Ieff,'n_seg':0,'R':float('nan'),'GPR':float('nan')}
    out={'Df':Df,'Sf':Sf,'Ieff':Ieff,'n_seg':b['n_seg_bem'],'R':b['R_grid'],'GPR':b['GPR'],'Es':b['Es'],'Et':b['Et']}
    if puntos: out['trans']=[(p,float(v)) for p,v in zip(puntos,evaluar_potencial_puntos(b,puntos))]
    return out

def perfiles(I,t,tipo,px1,py1,px2,py2):
    cond=estado['conductores']; res=calcular(cond,estado['jabalinas'],I,t)
    if res is None or res['bem'] is None: return "⚠️",None,None,None
    b=res['bem']; GPR=b['GPR']; GX,GY,Vs=b['GX'],b['GY'],b['Vsurf']
    Cs,Ep,Tp=limites_tolerables(PARAMETROS['h_grava'],PARAMETROS['rho_grava'],SUELO['rho1'],t,PARAMETROS.get('modo_exposicion','ocupacional_50'))
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
                    Cs,Ep,Tp=limites_tolerables(hs,rs,rho1,t,PARAMETROS.get('modo_exposicion','ocupacional_50'))
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
    import datetime
    doc=Document()
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run("SPT DESIGNER"); r.font.size=Pt(28); r.bold=True
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run("Memoria de cálculo — Sistema de Puesta a Tierra"); r.font.size=Pt(13)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run(PROYECTO['nombre']); r.font.size=Pt(14); r.bold=True
    doc.add_paragraph(f"{PROYECTO['cliente']} · {PROYECTO['ubicacion']} · {PROYECTO['ingeniero']}")
    doc.add_paragraph(f"Fecha de emisión: {datetime.date.today().strftime('%d/%m/%Y')}")
    doc.add_paragraph("Procedimiento según IEEE Std 80-2013, Cláusula 16.4 (mismo orden de pasos que los ejemplos del Anexo B), "
        "con los potenciales de paso y contacto calculados mediante un modelo numérico de elementos de frontera (BEM) en vez "
        "de las ecuaciones aproximadas de malla (Km/Ki) de la Cláusula 16.5.")

    modo_label={'ocupacional_50':'Ocupacional 50 kg (IEEE 80, Dalziel)','ocupacional_70':'Ocupacional 70 kg (IEEE 80, Dalziel)',
                'publico_retie':'Público en general (RETIE Tabla 3.12.1.a / IEC 60479-1)'}.get(res.get('modo_exposicion','ocupacional_50'),'Ocupacional 50 kg')
    I_total = res.get('I_malla',0)/max(res.get('Sf',1.0),1e-9)
    capa_txt = f"{estado['capa']['material']}, ρs={estado['capa']['rho_s']:.0f} Ω·m, hs={estado['capa']['hs']:.2f} m" if estado.get('capa') else "no definida (suelo desnudo)"

    doc.add_heading("Paso 1 — Datos de campo",level=1)
    doc.add_paragraph(f"Corriente de falla considerada (3I0): {fmt(I_total,0)} A. Tiempo de despeje de la falla: "
        f"{fmt(PARAMETROS['t_falla'],2)} s. Resistividad del suelo (ρ1): {fmt(SUELO.get('rho1'),1)} Ω·m. "
        f"Capa superficial: {capa_txt}.")

    doc.add_heading("Paso 2 — Calibre del conductor",level=1)
    doc.add_paragraph(f"Con la corriente de falla total y el tiempo de despeje, la sección mínima por criterio adiabático "
        f"(Ecuación 46/47 de IEEE 80) es {fmt(res['A_min'],1)} mm². Conductor seleccionado: "
        f"{PARAMETROS['material_conductor']} {PARAMETROS['calibre']} (k={fmt(PARAMETROS.get('k_material',202.8),1)}).")

    doc.add_heading("Paso 3 — Criterio de tensión tolerable",level=1)
    doc.add_paragraph(f"Criterio de exposición aplicado: {modo_label}. Factor de reducción de la capa superficial "
        f"Cs = {fmt(res['Cs'],3)}. Tensión de paso tolerable Es_perm = {fmt(res['Es_perm'],1)} V. "
        f"Tensión de contacto tolerable Et_perm = {fmt(res['Et_perm'],1)} V.")

    doc.add_heading("Paso 4 — Diseño inicial de la malla",level=1)
    doc.add_paragraph(f"Conductores: {res['n_cond']}. Jabalinas: {res['n_jab']}. Longitud total de conductor enterrado: "
        f"{fmt(res['L'],1)} m. Área ocupada por la malla: {fmt(res['area'],1)} m². Profundidad de enterramiento: "
        f"{fmt(PARAMETROS['prof_malla'],2)} m.")

    doc.add_heading("Paso 5 — Resistencia de la malla",level=1)
    doc.add_paragraph(f"Calculada con el modelo {res['metodo']}: Rg = {fmt(res['R'],3)} Ω.")

    doc.add_heading("Paso 6 — Corriente máxima de malla",level=1)
    doc.add_paragraph(f"Factor de división de corriente Sf = {fmt(res.get('Sf',1.0),2)} (fracción de 3I0 que retorna "
        f"efectivamente por la malla; el resto retorna por cables de guarda/neutro — IEEE 80 Cláusula 15/Anexo C). "
        f"Corriente de malla IG = 3I0 · Sf = {fmt(res.get('I_malla',0),0)} A.")

    doc.add_heading("Paso 7 — Elevación de potencial de tierra (GPR)",level=1)
    gpr_ok = res['GPR']<=res['Et_perm']
    doc.add_paragraph(f"GPR = IG · Rg = {fmt(res.get('I_malla',0),0)} A × {fmt(res['R'],3)} Ω = {fmt(res['GPR'],1)} V. "
        + ("Como el GPR ya está por debajo de la tensión de contacto tolerable, el diseño es aceptable sin necesidad de "
           "calcular el potencial punto a punto (mismo atajo que usa IEEE 80 cuando el GPR no supera Et_perm)."
           if gpr_ok else
           f"Como el GPR ({fmt(res['GPR'],1)} V) supera la tensión de contacto tolerable ({fmt(res['Et_perm'],1)} V), "
           "es necesario verificar el potencial real de paso y contacto en la superficie (Pasos 8-9)."))

    doc.add_heading("Paso 8 — Tensión de contacto y de paso",level=1)
    doc.add_paragraph(f"Tensión de contacto máxima encontrada en la superficie: Et = {fmt(res['Et'],1)} V "
        f"(vs. {fmt(res['Et_perm'],1)} V tolerable). Tensión de paso máxima: Es = {fmt(res['Es'],1)} V "
        f"(vs. {fmt(res['Es_perm'],1)} V tolerable). Nota: Et es el máximo encontrado en todo el perímetro de la malla, "
        "no solo el punto de referencia 'centro de malla de esquina' de la fórmula aproximada de IEEE 80, por lo que "
        "puede ser mayor que un cálculo de referencia con la metodología simplificada.")

    doc.add_heading("Paso 9 — Comparación y conclusión",level=1)
    t9=doc.add_table(rows=4,cols=3); t9.style='Table Grid'
    t9.rows[0].cells[0].text="Magnitud"; t9.rows[0].cells[1].text="Calculado"; t9.rows[0].cells[2].text="Tolerable"
    t9.rows[1].cells[0].text="Tensión de paso (Es)"; t9.rows[1].cells[1].text=fmt(res['Es'],1,' V'); t9.rows[1].cells[2].text=fmt(res['Es_perm'],1,' V')
    t9.rows[2].cells[0].text="Tensión de contacto (Et)"; t9.rows[2].cells[1].text=fmt(res['Et'],1,' V'); t9.rows[2].cells[2].text=fmt(res['Et_perm'],1,' V')
    t9.rows[3].cells[0].text="Veredicto"; t9.rows[3].cells[1].text="CUMPLE" if res['cumple'] else "NO CUMPLE"; t9.rows[3].cells[2].text=""
    if not res['cumple']:
        doc.add_paragraph("El diseño no cumple los criterios de IEEE 80/RETIE con los parámetros actuales. Opciones "
            "típicas de corrección: reducir el espaciado de conductores, agregar jabalinas perimetrales, aumentar el "
            "espesor o la resistividad de la capa superficial, o revisar si el factor Sf es realista para la topología "
            "real de cables de guarda del proyecto.")

    if estado.get('capa'):
        doc.add_heading("Anexo — Capa superficial",level=1); doc.add_paragraph(f"{estado['capa']['material']} ρs={estado['capa']['rho_s']:.0f} hs={estado['capa']['hs']:.2f} Fuente:{estado['capa']['fuente']}")
    if estado['alerts']:
        doc.add_heading("Anexo — Hallazgos del suelo",level=1)
        for a in estado['alerts']: doc.add_paragraph(f"[{a['tipo']}] {a['dir']} ρ={a['rho']:.0f}",style='List Bullet')
    try:
        doc.add_heading("Anexo — Planos",level=1)
        doc.add_picture(ruta('vista_planta.png'),width=Inches(6)); doc.add_picture(ruta('mapa_calor.png'),width=Inches(6))
    except Exception as e: doc.add_paragraph(str(e))

    doc.add_heading("Alcance y limitaciones",level=1)
    doc.add_paragraph("Este informe se generó con SPT Designer a partir de los parámetros de suelo, geometría y "
        "corriente vigentes al momento de la emisión. El potencial de contacto y de paso se calculó con un modelo "
        "numérico (BEM) de suelo UNIFORME (resistividad ρ1); si el suelo real es de dos capas y ρ2 difiere "
        "significativamente de ρ1, verifique el resultado con un método que sí modele la estratificación completa. "
        "El factor de división de corriente Sf fue ingresado por el usuario (no calculado automáticamente por este "
        "programa) — verifíquelo contra un estudio de cortocircuito/reparto de corriente (IEEE 80 Anexo C) para la "
        "topología real de cables de guarda y neutros del proyecto. Este documento no sustituye el criterio del "
        "ingeniero responsable ni los trámites de certificación RETIE ante el organismo competente.")

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
    pre = f"{res['aviso_2capas']}\n\n" if res.get('aviso_2capas') else ""
    if res['cumple']: return pre+"✅ **Conforme.** Puedes optimizar costo (mayor espaciado) y re-verificar."
    r=[]
    if res['Et']>res['Et_perm']: r+=["🔴 **Contacto excede:** densifica perímetro, jabalinas en esquinas, o grava (ρs↑)."]
    if res['Es']>res['Es_perm']: r+=["🔴 **Paso excede:** grava o menor t."]
    if res['GPR']>5000: r+=["⚠️ **GPR alto:** más jabalinas/electrodos profundos o mayor área."]
    r+=["💡 Usa **Optimizar** para buscar configuración que cumple."]; return pre+"\n".join(r)

with gr.Blocks(title="SPT Designer") as app:
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
                q_peso=gr.Dropdown(list(MODOS_EXPOSICION),value='Ocupacional 50 kg (IEEE 80)',label="Tipo de exposición / criterio de tensión tolerable",
                    info="Ocupacional=Dalziel/IEEE 80. Público=RETIE Tabla 3.12.1.a (IEC 60479-1); obligatorio si el sitio tiene acceso público (RETIE Art. 3.12.1.g)")
                q_sf=gr.Number(value=1.0,label="Factor de división de corriente Sf",minimum=0.0,maximum=1.0,step=0.05,
                    info="Fracción de 3I0 que retorna por la malla (resto por cables de guarda/neutro). 1.0=conservador. Ver IEEE 80 Cl.15/Anexo C, calcúlelo con Anexo C o software de flujo de fallas")
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
                aI=gr.Number(value=8000,label="Corriente de falla I (A)",info="Se sincroniza al guardar en 1📋. Puedes editarla aquí para probar escenarios sin alterar el proyecto guardado.")
                aT=gr.Number(value=0.3,label="Tiempo de despeje t (s)")
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

    # AUDITORÍA (punto pendiente #4): antes I/t vivían por duplicado en 1📋 y en
    # 4⚡ sin sincronizarse — si el usuario actualizaba una y olvidaba la otra,
    # el análisis principal (que lee de 4⚡) quedaba con un valor viejo sin
    # avisar. Ahora, al guardar en 1📋, se empujan los mismos valores a los
    # campos de 4⚡ (que siguen editables ahí para explorar "qué pasa si..."
    # sin tener que re-guardar el proyecto cada vez).
    b_datos.click(aplicar_datos,inputs=[p_nom,p_cli,p_ubi,p_ing,q_I,q_t,q_peso,q_sf,q_pm,q_lj,q_mc,q_cal],outputs=[datos_out,aI,aT])
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
        msg,c,j=importar_dxf(open(_ruta_archivo(f)).read(),m); estado['conductores']=c; estado['jabalinas']=j; estado['dirty']=True
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
            recs_txt=recomendaciones(res)
            # AUDITORÍA (punto pendiente #4, parte "aviso"): si esta corriente/tiempo
            # de prueba difiere de lo guardado en 1📋, se avisa explícitamente en vez
            # de dejar que el usuario asuma que está viendo el escenario del proyecto.
            if abs(sanitizar(I,0)-PARAMETROS['I_falla'])>1e-6 or abs(sanitizar(t,0)-PARAMETROS['t_falla'])>1e-6:
                recs_txt = f"ℹ️ **Estás probando I={sanitizar(I,0):.0f}A/t={sanitizar(t,0):.2f}s, distinto de lo guardado en 1📋 (I={PARAMETROS['I_falla']:.0f}A/t={PARAMETROS['t_falla']:.2f}s). Esto no modifica el proyecto guardado.**\n\n" + recs_txt
            return tablero_html(res),num,ruta('mapa_calor.png'),recs_txt,stepper()
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
        r=calcular_preciso(estado['conductores'],PARAMETROS['I_falla'],PARAMETROS['t_falla'],SUELO['rho1'],float(XoR),float(ml),pts or None,Sf=PARAMETROS.get('Sf',1.0),jab=estado['jabalinas'],long_jabalina=PARAMETROS.get('long_jabalina',0.0))
        txt=f"**Sf={r['Sf']:.2f} · Df={r['Df']:.3f} → I efectiva={r['Ieff']:.0f} A** · tramos={r['n_seg']} · R={r['R']:.3f} Ω · GPR={r['GPR']:.0f} V"
        if 'Es' in r: txt+=f" · Es={r['Es']:.0f} V · Et={r['Et']:.0f} V"
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
# AUDITORÍA: en Gradio 6.0, "theme" se movió del constructor de Blocks a
# launch() (antes daba un UserWarning de deprecación en cada arranque).
app.launch(server_name="0.0.0.0", server_port=int(os.environ.get("PORT", os.environ.get("GRADIO_SERVER_PORT","7860"))), theme=gr.themes.Soft())
