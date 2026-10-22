#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""FSTT: análisis por sujeto y landmark. Nunca modifica la entrada.
Instalar: py -m pip install numpy pandas scipy openpyxl matplotlib
Ejecutar: py estadisticas_tesis_HC3.py "Excel_addon_tesis.xlsx"

Modelo factorial saturado de cuatro medias, covarianza HC3, inferencia t(n-4).
Efectos principales: contrastes equiponderados; interacción: diferencia de diferencias.
BH en cada familia de 26 efectos; Holm en los 104 contrastes planificados.
CCU principal: hoja Resultados, por 10 mm, ajustada por edad continua y sexo, HC3.
CCU secundaria: clasificación de cribado por puntos de corte sexo×edad (Fuentes et al., 2018), 7–18 años.
Asimetría: media pareada con IC bootstrap; t pareada y BH; Wilcoxon/signos secundarios.
ICC(A,1), TEM, rTEM y Bland–Altman con remuestreo de pares completos.
Las advertencias se exportan para revisión; no se eliminan outliers automáticamente.
"""
from pathlib import Path
import argparse, hashlib, platform, sys, warnings
from datetime import datetime
import numpy as np
import pandas as pd
from scipy import stats
import scipy

LANDMARKS = ['g','n','mn','rhi','sn','mp','ls','li','mls','pg','gn','m',
 'mso-izq','mso-der','mio-izq','mio-der','acp-izq','acp-der','go-izq','go-der',
 'zy-izq','zy-der','mr-izq','mr-der','mmb-izq','mmb-der']
PAIRS = ['mso','mio','acp','go','zy','mr','mmb']
GROUPS = ['F_6-11','F_12-18','M_6-11','M_12-18']
EFFECTS = {'Sexo_M-F':[-.5,-.5,.5,.5], 'Edad_mayores-menores':[-.5,.5,-.5,.5],
 'Interaccion':[1,-1,-1,1]}
CONTRASTS = {'M-F_6-11':[-1,0,1,0], 'M-F_12-18':[0,-1,0,1],
 'Mayores-menores_F':[-1,1,0,0], 'Mayores-menores_M':[0,0,-1,1]}

# Puntos de corte CCU (cm) para cribado de malnutrición por exceso en población chilena.
# Fuentes Fuentes J, Hidalgo Fernández A, Durán Agüero S, Silva Ocampo P.
# Rev Esp Nutr Comunitaria. 2018;24(4):130-134. DOI:10.14642/RENC.2018.24.4.5259
# La referencia publica cortes para 7–18 años, separados por sexo y edad.
CCU_CUTOFF_CM = {
    'F': {7:26.8, 8:27.1, 9:28.2, 10:28.5, 11:29.2, 12:31.2,
          13:31.4, 14:32.0, 15:32.2, 16:33.5, 17:34.3, 18:34.7},
    'M': {7:27.1, 8:28.0, 9:30.0, 10:30.2, 11:31.0, 12:32.0,
          13:32.4, 14:36.0, 15:37.0, 16:37.5, 17:38.0, 18:38.5},
}



def adjust(p, method='BH'):
    """BH/Holm; NaN mantiene un lugar en la familia prevista (p=1 al ajustar)."""
    p=np.asarray(p,float); valid=np.isfinite(p); filled=np.where(valid,p,1.)
    idx=np.argsort(filled); v=filled[idx]; n=len(v)
    if not n: return p
    if method=='BH': a=np.minimum.accumulate((v*n/np.arange(1,n+1))[::-1])[::-1]
    else: a=np.maximum.accumulate(v*np.arange(n,0,-1))
    result=np.empty(n); result[idx]=np.minimum(a,1);result[~valid]=np.nan
    return result


def flag(value):
    if pd.isna(value): return False
    return str(value).strip().upper() in {'1','1.0','SI','SÍ','TRUE','YES'}


def sex(value):
    v=str(value).strip().upper()
    if v in {'F','FEMENINO','FEMALE','MUJER'}: return 'F'
    if v in {'M','MASCULINO','MALE','HOMBRE'}: return 'M'
    return None


def ccu_cutoff_cm(sexo, edad):
    """Corte por sexo y edad cumplida. Devuelve NaN fuera de 7–18 años."""
    if sexo not in CCU_CUTOFF_CM or not np.isfinite(edad): return np.nan
    age=int(np.floor(float(edad)))
    return CCU_CUTOFF_CM[sexo].get(age, np.nan)


def ccu_cutoff_table():
    rows=[]
    for sx,label in [('F','Femenino'),('M','Masculino')]:
        for age,cut in CCU_CUTOFF_CM[sx].items():
            rows.append({'Sexo':sx,'Sexo_etiqueta':label,'Edad_anios':age,'Punto_corte_CCU_cm':cut,
                         'Regla':'Cribado positivo si CCU >= punto de corte'})
    return pd.DataFrame(rows)


def thesis_descriptive_tables(desc):
    """Tablas de tesis: una fila por landmark/lado y bloques sexo×edad en columnas.

    Evita repetir cuatro filas (n, media, mediana, rango) por landmark. Los puntos
    bilaterales quedan emparejados anatómicamente: izquierda y derecha en filas
    consecutivas. La tabla larga `Descriptivos` conserva todos los números exactos.
    """
    d=pd.DataFrame(desc)
    if d.empty:return pd.DataFrame(),pd.DataFrame()
    d=d[d.Grupo.ne('TOTAL')].copy().set_index(['Landmark','Grupo'])

    def values_for(lm, group):
        try:r=d.loc[(lm,group)]
        except KeyError:return {'n':'','Media ± DE':'','Mediana [Q1–Q3]':'','Min–Max':''}
        return {
            'n':int(r.n),
            'Media ± DE':f"{r.Media:.2f} ± {r.DE:.2f}",
            'Mediana [Q1–Q3]':f"{r.Mediana:.2f} [{r.Q1:.2f}–{r.Q3:.2f}]",
            'Min–Max':f"{r.Min:.2f}–{r.Max:.2f}",
        }

    labels={'F_6-11':'Mujer 6–11','F_12-18':'Mujer 12–18',
            'M_6-11':'Hombre 6–11','M_12-18':'Hombre 12–18'}
    def row_for(lm, landmark_label=None, side=''):
        row={'Landmark':landmark_label or lm.upper(),'Lado':side}
        for group in GROUPS:
            vals=values_for(lm,group);prefix=labels[group]
            row[prefix+' | n']=vals['n']
            row[prefix+' | Media ± DE']=vals['Media ± DE']
            row[prefix+' | Mediana [Q1–Q3]']=vals['Mediana [Q1–Q3]']
            row[prefix+' | Min–Max']=vals['Min–Max']
        return row

    unilateral=[lm for lm in LANDMARKS if not lm.endswith('-izq') and not lm.endswith('-der')]
    uni=pd.DataFrame([row_for(lm) for lm in unilateral])
    bilateral=[]
    for pair in PAIRS:
        bilateral.append(row_for(pair+'-izq',pair.upper(),'Izquierda'))
        bilateral.append(row_for(pair+'-der','', 'Derecha'))
    bi=pd.DataFrame(bilateral)
    return uni,bi



def bilateral_table_with_asymmetry(desc, asym_group):
    """Tabla bilateral final: Izquierda, Derecha y Asimetría I-D como FILAS.

    Cada grupo sexo×edad ocupa un bloque de columnas. Para las filas izquierda y
    derecha, ``n`` corresponde a sujetos con ese lado válido. Para la fila de
    asimetría, ``n`` corresponde exclusivamente a sujetos con AMBOS lados del par.
    La asimetría se resume como diferencia izquierda-derecha y su p ajustado BH
    se calcula dentro de cada grupo sobre los siete pares.
    """
    d=pd.DataFrame(desc)
    if d.empty:return pd.DataFrame()
    d=d[d.Grupo.ne('TOTAL')].copy().set_index(['Landmark','Grupo'])
    ag=pd.DataFrame(asym_group).copy()
    if not ag.empty:ag=ag.set_index(['Par','Grupo'])
    labels={'F_6-11':'Mujer 6–11','F_12-18':'Mujer 12–18',
            'M_6-11':'Hombre 6–11','M_12-18':'Hombre 12–18'}
    rows=[]
    for pair in PAIRS:
        for side,lm in [('Izquierda',pair+'-izq'),('Derecha',pair+'-der')]:
            row={'Landmark':pair.upper(),'Fila':side}
            for group in GROUPS:
                pref=labels[group]
                try:r=d.loc[(lm,group)]
                except KeyError:r=None
                row[pref+' | n']=int(r.n) if r is not None else ''
                row[pref+' | Media ± DE']=(f"{r.Media:.2f} ± {r.DE:.2f}" if r is not None else '')
                row[pref+' | Mediana [Q1–Q3]']=(f"{r.Mediana:.2f} [{r.Q1:.2f}–{r.Q3:.2f}]" if r is not None else '')
                row[pref+' | Min–Max']=(f"{r.Min:.2f}–{r.Max:.2f}" if r is not None else '')
                row[pref+' | p asim. BH']=''
            rows.append(row)
        row={'Landmark':pair.upper(),'Fila':'Asimetría I−D'}
        for group in GROUPS:
            pref=labels[group]
            try:r=ag.loc[(pair,group)] if not ag.empty else None
            except KeyError:r=None
            if r is None:
                row[pref+' | n']='';row[pref+' | Media ± DE']='';row[pref+' | Mediana [Q1–Q3]']='';row[pref+' | Min–Max']='';row[pref+' | p asim. BH']=''
            else:
                row[pref+' | n']=int(r.n_pares)
                row[pref+' | Media ± DE']=(f"{r['Media_I-D']:.2f} ± {r['DE_I-D']:.2f}" if np.isfinite(r.get('DE_I-D',np.nan)) else '')
                row[pref+' | Mediana [Q1–Q3]']=(f"{r['Mediana_I-D']:.2f} [{r['Q1_I-D']:.2f}–{r['Q3_I-D']:.2f}]" if np.isfinite(r.get('Mediana_I-D',np.nan)) else '')
                row[pref+' | Min–Max']=(f"{r['Min_I-D']:.2f}–{r['Max_I-D']:.2f}" if np.isfinite(r.get('Min_I-D',np.nan)) else '')
                pval=r.get('p_BH_t_pareada',np.nan)
                row[pref+' | p asim. BH']=(float(pval) if np.isfinite(pval) else '')
        rows.append(row)
    return pd.DataFrame(rows)


def interaction_contrast_table(effects, posts):
    """Una fila por landmark con interacción y efectos simples por sexo/edad."""
    e=pd.DataFrame(effects);p=pd.DataFrame(posts)
    inter=e[e.Efecto.eq('Interaccion')].set_index('Landmark') if not e.empty else pd.DataFrame()
    rows=[]
    labels={
        'M-F_6-11':'M−F 6–11',
        'M-F_12-18':'M−F 12–18',
        'Mayores-menores_F':'12–18 − 6–11 Mujeres',
        'Mayores-menores_M':'12–18 − 6–11 Hombres',
    }
    for lm in LANDMARKS:
        row={'Landmark':lm}
        if not inter.empty and lm in inter.index:
            r=inter.loc[lm]
            row.update({'Interacción_estimación_mm':r.get('Estimacion_mm',np.nan),
                        'Interacción_IC95_inf':r.get('IC95_inf',np.nan),
                        'Interacción_IC95_sup':r.get('IC95_sup',np.nan),
                        'Interacción_p_BH':r.get('p_BH',np.nan),
                        'Interacción_significativa':bool(np.isfinite(r.get('p_BH',np.nan)) and r.get('p_BH',1)<0.05)})
        for code,label in labels.items():
            rr=p[(p.Landmark.eq(lm))&(p.Contraste.eq(code))]
            if rr.empty:continue
            rr=rr.iloc[0]
            row[label+' | Estimación mm']=rr.get('Estimacion_mm',np.nan)
            row[label+' | IC95 inf']=rr.get('IC95_inf',np.nan)
            row[label+' | IC95 sup']=rr.get('IC95_sup',np.nan)
            row[label+' | p Holm(26)']=rr.get('p_Holm_familia',np.nan)
        rows.append(row)
    return pd.DataFrame(rows)

def bootstrap_ci(x, rng, b, stat='mean'):
    x=np.asarray(x,float); x=x[np.isfinite(x)]
    if len(x)<2:return np.nan,np.nan
    sample=x[rng.integers(len(x),size=(b,len(x)))]
    est=sample.mean(axis=1) if stat=='mean' else np.median(sample,axis=1)
    return tuple(np.quantile(est,[.025,.975]))


def fit_hc3(X,y):
    """OLS y sandwich HC3: (X'X)^-1 X'diag(e²/(1-h)²)X (X'X)^-1.
    SVD/pseudoinversa evita invertir directamente matrices mal condicionadas.
    Rechaza diseños singulares: no inventa coeficientes identificables.
    """
    X=np.asarray(X,float); y=np.asarray(y,float);n,p=X.shape
    if n<=p or np.linalg.matrix_rank(X)<p:raise ValueError('Diseño singular o n insuficiente')
    pinv=np.linalg.pinv(X); beta=pinv@y; residual=y-X@beta
    h=np.sum(X*pinv.T,axis=1)
    if np.any(h>=1-1e-10):raise ValueError('Apalancamiento unitario: HC3 no estimable')
    covariance=(pinv*(residual/(1-h))**2)@pinv.T
    sse=residual@residual;sst=np.sum((y-y.mean())**2);mse=sse/(n-p)
    cook=np.divide(residual**2*h,p*mse*(1-h)**2,out=np.zeros(n),where=mse>0)
    return {'beta':beta,'cov':covariance,'df':n-p,'n':n,'p':p,'resid':residual,
            'h':h,'cook':cook,'fitted':X@beta,'R2':1-sse/sst if sst>0 else np.nan}


def contrast(fit,L):
    L=np.asarray(L,float);estimate=float(L@fit['beta']);se=np.sqrt(max(0,float(L@fit['cov']@L)))
    if se>0:t=estimate/se;p=2*stats.t.sf(abs(t),fit['df'])
    else:t=np.nan;p=np.nan
    critical=stats.t.ppf(.975,fit['df'])
    return {'Estimacion_mm':estimate,'EE_HC3':se,'IC95_inf':estimate-critical*se,
            'IC95_sup':estimate+critical*se,'t':t,'gl':fit['df'],'p':p,'n':fit['n']}


def icc_a1(a):
    """ICC de acuerdo absoluto, dos sesiones, medidas individuales."""
    n,k=a.shape; grand=a.mean(); ssr=k*np.sum((a.mean(axis=1)-grand)**2)
    ssc=n*np.sum((a.mean(axis=0)-grand)**2)
    sse=np.sum((a-grand)**2)-ssr-ssc
    msr=ssr/(n-1);msc=ssc/(k-1);mse=sse/((n-1)*(k-1))
    den=msr+(k-1)*mse+k*(msc-mse)/n
    return (msr-mse)/den if abs(den)>1e-15 else np.nan


def load_data(path, observer, first, second, replica):
    w=pd.read_excel(path,sheet_name=None,engine='openpyxl');qc=[];discard=[]
    def alert(kind,case='',detail=''):qc.append({'Tipo':kind,'CaseID':case,'Detalle':detail})
    for name in ['Casos','Mediciones_LM','Resultados']:
        if name not in w:raise ValueError('Falta hoja '+name)
    c=w['Casos'].copy();m=w['Mediciones_LM'].copy();r=w['Resultados'].copy()
    c['CaseID']=c.get('CaseID',c.Paciente).fillna(c.Paciente).astype(str).str.strip()
    for col in ['¿Incluir?','¿Es la mejor carpeta del paciente? (1 = Sí)']:
        if col in c:c=c[c[col].map(flag)].copy()
        else:alert('Filtro_ausente','',col+'; se conservan casos y se informa ausencia')
    if c.CaseID.duplicated().any() or c.Paciente.duplicated().any():
        raise ValueError('Casos incluidos duplicados; resolver antes de analizar')
    c['Sexo']=c.Sexo.map(sex);c['Edad']=pd.to_numeric(c.Edad,errors='coerce')
    bad=c.Sexo.isna()|~c.Edad.between(6,18)
    if bad.any():raise ValueError('Edad/sexo inválidos: '+', '.join(c.loc[bad,'CaseID']))
    c['GrupoEdad']=np.where(c.Edad<12,'6-11','12-18');c['Grupo']=c.Sexo+'_'+c.GrupoEdad
    c['TieneCuello']=c['¿Tiene cuello? (1 = Sí, 0 = No)'].map(flag)
    c=c[['CaseID','Paciente','Sexo','Edad','GrupoEdad','Grupo','TieneCuello']].copy()
    for col in ['Observador','Sesion','Landmark','Status']:
        m[col]=m[col].astype(str).str.strip()
    m['Observador']=m.Observador.str.upper();m['Sesion']=m.Sesion.str.upper()
    m['Landmark']=m.Landmark.str.lower();m['Status']=m.Status.str.upper()
    m['CaseID']=m.CaseID.fillna(m.Paciente).astype(str).str.strip()
    m['Replica']=pd.to_numeric(m.Replica,errors='coerce')
    m['Thickness_mm']=pd.to_numeric(m.Thickness_mm,errors='coerce')
    m['Modo']=m.get('ModoMedicion',pd.Series('',index=m.index)).fillna('').astype(str).str.upper().str.strip()
    m['Fila_Excel']=np.arange(len(m))+2
    unknown=set(m.CaseID)-set(c.CaseID)
    if unknown:alert('Mediciones_fuera_muestra','',str(sorted(unknown)))
    m=m[m.CaseID.isin(c.CaseID)&m.Observador.eq(observer)&m.Replica.eq(replica)].copy()
    def session_data(session,desired):
        # La medición se identifica por contexto explícito; Timestamp NO participa.
        s=m[m.Sesion.eq(session)].copy()
        if desired=='NORMAL':
            good=s.Modo.eq('NORMAL')|s.Modo.eq('')  # legado S1 sin modo
        else:
            good=s.Modo.eq(desired)                 # ERROR debe estar explícitamente marcado
        for row in s[~good].itertuples():
            discard.append({'CaseID':row.CaseID,'Landmark':row.Landmark,'Sesion':session,
                            'Fila_Excel':row.Fila_Excel,'Motivo':'Contexto no seleccionado: modo '+row.Modo})
        s=s[good].copy()
        legacy=int(s.Modo.eq('').sum())
        if legacy:alert('Modo_historico_vacio','',f'{session}: {legacy} filas históricas sin modo; se interpretan como NORMAL principal')
        key=['CaseID','Landmark'];dups=s[s.duplicated(key,False)]
        if len(dups):
            examples=dups[key].drop_duplicates().head(12).astype(str).agg(' / '.join,axis=1).tolist()
            raise ValueError('Más de un registro para la misma medición lógica; no se resolverá por fecha. Revisar: '+', '.join(examples))
        return s
    s1=session_data(first,'NORMAL');s2=session_data(second,'ERROR')
    if discard:alert('Filas_no_seleccionadas','',f'{len(discard)}; ver hoja Seleccion_registros')
    grid=pd.MultiIndex.from_product([c.CaseID,LANDMARKS],names=['CaseID','Landmark']).to_frame(index=False)
    base=grid.merge(s1[['CaseID','Landmark','Status','Thickness_mm','Modo','Fila_Excel']],
                    how='left',on=['CaseID','Landmark'],validate='one_to_one').merge(c,on='CaseID',validate='many_to_one')
    base['Status']=base.Status.fillna('SIN_REGISTRO')
    base['Valido']=base.Status.eq('SET')&np.isfinite(base.Thickness_mm)&base.Thickness_mm.gt(0)
    for row in base[base.Status.eq('SET')&~base.Valido].itertuples():alert('Espesor_facial_invalido',row.CaseID,row.Landmark)
    # Resultados es la fuente explícita del registro cervical actual, nunca S2 de la tabla larga.
    r['Paciente']=r.Paciente.astype(str).str.strip()
    r=r[r.Paciente.isin(c.Paciente)].copy()
    if r.Paciente.duplicated().any():raise ValueError('Resultados tiene pacientes duplicados')
    neck_cols=['Paciente','CCU_mm']
    if 'Estado_Nutricional' in r.columns:neck_cols.append('Estado_Nutricional')
    neck=c.merge(r[neck_cols],on='Paciente',how='left',validate='one_to_one')
    neck['CCU_mm']=pd.to_numeric(neck.CCU_mm,errors='coerce');neck['Fuente']='Resultados.CCU_mm'
    neck['CCU_valida']=neck.TieneCuello&np.isfinite(neck.CCU_mm)&neck.CCU_mm.gt(0)
    neck['CCU_cm']=neck.CCU_mm/10
    neck['CCU_corte_cm']=[ccu_cutoff_cm(sx,age) for sx,age in zip(neck.Sexo,neck.Edad)]
    # 200–500 mm es SOLO una regla de revisión de magnitud ya usada por el script; no define nutrición.
    neck['CCU_clasificable_Fuentes']=neck.CCU_valida&neck.CCU_mm.between(200,500)&np.isfinite(neck.CCU_corte_cm)
    neck['CCU_cribado_Fuentes']='No clasificable'
    neck.loc[neck.CCU_valida&~np.isfinite(neck.CCU_corte_cm),'CCU_cribado_Fuentes']='Sin referencia Fuentes (edad fuera de 7–18)'
    neck.loc[neck.CCU_valida&~neck.CCU_mm.between(200,500),'CCU_cribado_Fuentes']='Revisar magnitud antes de clasificar'
    neg=neck.CCU_clasificable_Fuentes&(neck.CCU_cm<neck.CCU_corte_cm)
    pos=neck.CCU_clasificable_Fuentes&(neck.CCU_cm>=neck.CCU_corte_cm)
    neck.loc[neg,'CCU_cribado_Fuentes']='Cribado negativo para malnutrición por exceso'
    neck.loc[pos,'CCU_cribado_Fuentes']='Cribado positivo para malnutrición por exceso'
    neck['CCU_cribado_positivo']=np.where(neck.CCU_clasificable_Fuentes,pos.astype(int),np.nan)
    neck['Revision_fuente']=False
    for row in neck[neck.CCU_valida&((neck.CCU_mm<200)|(neck.CCU_mm>500))].itertuples():
        alert('CCU_revisar_magnitud',row.CaseID,str(row.CCU_mm)+' mm; no excluida automáticamente')
    # Cotejo de resumen contra la sesión principal: una divergencia no reemplaza S1.
    for lm in LANDMARKS:
        col=lm+'_mm'
        if col not in r:continue
        d=base[(base.Landmark==lm)&base.Valido].merge(r[['Paciente',col]],on='Paciente')
        mismatch=pd.to_numeric(d[col],errors='coerce').isna()|((d.Thickness_mm-pd.to_numeric(d[col],errors='coerce')).abs()>1e-3)
        for row in d[mismatch].itertuples():
            alert('Resumen_distinto_S1',row.CaseID,lm+'; se mantiene S1, revisar procedencia del resumen/CCU')
            neck.loc[neck.CaseID.eq(row.CaseID),'Revision_fuente']=True
    return c,base,s1,s2,neck,qc,discard


def analyze(c,base,s1,s2,neck,args,qc):
    rng=np.random.default_rng(args.seed);b=args.bootstrap;tabs={}
    main=base[base.Valido].copy(); tabs['Muestra']=c;tabs['Base_principal']=base;tabs['CCU_fuente']=neck
    samples=c.groupby('Grupo',sort=False).Edad.agg(n='size',Media='mean',DE='std',Min='min',Max='max').reset_index()
    tabs['Resumen_muestra']=samples
    availability=[];desc=[];normal=[];effects=[];posts=[];fits=[];resids=[]
    for lm in LANDMARKS:
        full=base[base.Landmark.eq(lm)];d=main[main.Landmark.eq(lm)]
        for group in ['TOTAL']+GROUPS:
            f=full if group=='TOTAL' else full[full.Grupo.eq(group)]
            x=f.loc[f.Valido,'Thickness_mm'].to_numpy(float)
            availability.append({'Landmark':lm,'Grupo':group,'Esperado':len(f),'Valido':len(x),
                 'No_valido':len(f)-len(x),'MISSING':int(f.Status.eq('MISSING').sum()),
                 'Pendiente_OMIT':int(f.Status.eq('OMIT').sum()),'Sin_registro':int(f.Status.eq('SIN_REGISTRO').sum()),
                 'Cobertura_pct':100*len(x)/len(f) if len(f) else np.nan})
            if len(x)<2:continue
            q1,med,q3=np.quantile(x,[.25,.5,.75]);ci=bootstrap_ci(x,rng,b);mc=bootstrap_ci(x,rng,b,'median')
            desc.append({'Landmark':lm,'Grupo':group,'n':len(x),'Media':x.mean(),'DE':x.std(ddof=1),
                         'Mediana':med,'Q1':q1,'Q3':q3,'RIC':q3-q1,'Min':x.min(),'Max':x.max(),
                         'IC95_Media_inf':ci[0],'IC95_Media_sup':ci[1],
                         'IC95_Mediana_inf':mc[0],'IC95_Mediana_sup':mc[1]})
            sh=stats.shapiro(x) if len(x)>=3 and np.ptp(x)>0 else (np.nan,np.nan)
            normal.append({'Landmark':lm,'Grupo':group,'n':len(x),'W':sh[0],'p_diagnostico':sh[1],
                           'Asimetria':stats.skew(x,bias=False),'Curtosis_exceso':stats.kurtosis(x,bias=False)})
        X=np.column_stack([d.Grupo.eq(g).astype(float) for g in GROUPS])
        try:
            fit=fit_hc3(X,d.Thickness_mm)
            for term,L in EFFECTS.items():effects.append({'Landmark':lm,'Efecto':term,**contrast(fit,L)})
            for term,L in CONTRASTS.items():posts.append({'Landmark':lm,'Contraste':term,**contrast(fit,L)})
            fits.append({'Landmark':lm,'n':fit['n'],'R2':fit['R2'],'Max_Cook':fit['cook'].max(),
                         'Alertas_Cook_4_n':int((fit['cook']>4/fit['n']).sum())})
            for j,row in enumerate(d.itertuples()):resids.append({'CaseID':row.CaseID,'Landmark':lm,
                'Ajustado':fit['fitted'][j],'Residuo':fit['resid'][j],'Cook':fit['cook'][j],
                'Alerta_Cook':fit['cook'][j]>4/fit['n']})
        except ValueError as ex:
            qc.append({'Tipo':'Modelo_no_estimable','CaseID':'','Detalle':lm+': '+str(ex)})
            for term in EFFECTS:effects.append({'Landmark':lm,'Efecto':term,'p':np.nan})
            for term in CONTRASTS:posts.append({'Landmark':lm,'Contraste':term,'p':np.nan})
    e=pd.DataFrame(effects);e['p_BH']=e.groupby('Efecto',sort=False).p.transform(lambda p:adjust(p))
    p=pd.DataFrame(posts);p['p_Holm_familia']=p.groupby('Contraste',sort=False).p.transform(lambda x:adjust(x,'Holm'));p['p_Holm_104']=adjust(p.p,'Holm')
    tabs.update({'Disponibilidad':pd.DataFrame(availability),'Descriptivos':pd.DataFrame(desc),
        'Normalidad':pd.DataFrame(normal),
        'Factorial_HC3':e,'Contrastes_Holm':p,'Ajuste_modelos':pd.DataFrame(fits),'Diagnosticos':pd.DataFrame(resids)})
    tabs['Interaccion_y_contrastes']=interaction_contrast_table(e,p)
    uni_desc,bi_desc=thesis_descriptive_tables(desc)
    tabs['Descriptivos_unilaterales']=uni_desc
    tabs['Descriptivos_bilaterales']=bi_desc
    n_valid_per_case=main.groupby('CaseID').Landmark.nunique().reindex(c.CaseID,fill_value=0)
    flow=[
        {'Indicador':'Sujetos incluidos en Casos','n':len(c),'Nota':'Denominador de la muestra; no exige 26 landmarks completos'},
        {'Indicador':'Sujetos con 26/26 landmarks válidos','n':int((n_valid_per_case==len(LANDMARKS)).sum()),'Nota':'Solo informativo; NO es el n obligatorio para los modelos por landmark'},
        {'Indicador':'Mediciones FSTT válidas','n':int(base.Valido.sum()),'Nota':f'De {len(c)*len(LANDMARKS)} combinaciones sujeto×landmark posibles'},
        {'Indicador':'CCU válidas continuas','n':int(neck.CCU_valida.sum()),'Nota':'Usables en el modelo continuo, sujeto a revisión QC'},
        {'Indicador':'CCU clasificables Fuentes 7–18','n':int(neck.CCU_clasificable_Fuentes.sum()),'Nota':'Requiere edad 7–18 y magnitud CCU 200–500 mm'},
    ]
    tabs['Flujo_analitico']=pd.DataFrame(flow)
    tabs['CCU_puntos_corte']=ccu_cutoff_table()
    # Tabla de referencia: media(DE), mediana[Q1;Q3], n; los números exactos quedan arriba.
    ref=pd.DataFrame(desc);ref=ref[ref.Grupo.ne('TOTAL')].copy()
    ref['Referencia']=ref.apply(lambda r:f"{r.Media:.2f} ({r.DE:.2f}); {r.Mediana:.2f} [{r.Q1:.2f}; {r.Q3:.2f}]; n={int(r.n)}",axis=1)
    tabs['Tabla_referencia']=ref.pivot(index='Landmark',columns='Grupo',values='Referencia').reindex(index=LANDMARKS,columns=GROUPS).reset_index()
    # Asimetría: el sujeto, con sus dos lados, es la unidad del remuestreo.
    wide=main.pivot(index='CaseID',columns='Landmark',values='Thickness_mm').reindex(columns=LANDMARKS)
    asym=[]
    for pair in PAIRS:
        a=wide[[pair+'-izq',pair+'-der']].dropna();diff=(a.iloc[:,0]-a.iloc[:,1]).to_numpy();n=len(diff)
        if n<3:continue
        ci=bootstrap_ci(diff,rng,b);medci=bootstrap_ci(diff,rng,b,'median');sd=diff.std(ddof=1)
        wil=stats.wilcoxon(np.round(diff,args.decimales_diferencia)).pvalue if np.any(np.round(diff,args.decimales_diferencia)!=0) else 1.
        positive=int((diff>0).sum());negative=int((diff<0).sum());nz=positive+negative
        asym.append({'Par':pair,'n':n,'Media_I':a.iloc[:,0].mean(),'Media_D':a.iloc[:,1].mean(),
            'Media_I-D':diff.mean(),'IC95_media_inf':ci[0],'IC95_media_sup':ci[1],
            'Mediana_I-D':np.median(diff),'IC95_mediana_inf':medci[0],'IC95_mediana_sup':medci[1],
            'Mediana_absoluta':np.median(abs(diff)),'p_t_pareada':stats.ttest_1samp(diff,0).pvalue if sd>0 else np.nan,
            'p_Wilcoxon':wil,'p_signos':stats.binomtest(positive,nz,.5).pvalue if nz else 1.,
            'Positivas':positive,'Negativas':negative,'Ceros':n-nz})
    a=pd.DataFrame(asym)
    for name in ['t_pareada','Wilcoxon','signos']:a['p_BH_'+name]=adjust(a['p_'+name])
    tabs['Asimetria']=a
    # Asimetría por grupo: cada par usa EXCLUSIVAMENTE sujetos con ambos lados válidos.
    # No exige que el sujeto tenga los otros pares ni los 26 landmarks completos.
    wide_group=wide.join(c.set_index('CaseID')[['Grupo']],how='left')
    asym_group=[]
    for group in GROUPS:
        wg=wide_group[wide_group.Grupo.eq(group)]
        for pair in PAIRS:
            a2=wg[[pair+'-izq',pair+'-der']].dropna()
            diff=(a2.iloc[:,0]-a2.iloc[:,1]).to_numpy(float);n2=len(diff)
            if n2<3:
                asym_group.append({'Grupo':group,'Par':pair,'n_pares':n2})
                continue
            ci=bootstrap_ci(diff,rng,b);sd=diff.std(ddof=1);q1d,medd,q3d=np.quantile(diff,[.25,.5,.75])
            asym_group.append({'Grupo':group,'Par':pair,'n_pares':n2,
                'Media_I':a2.iloc[:,0].mean(),'Media_D':a2.iloc[:,1].mean(),
                'Media_I-D':diff.mean(),'DE_I-D':sd,'IC95_media_inf':ci[0],'IC95_media_sup':ci[1],
                'Mediana_I-D':medd,'Q1_I-D':q1d,'Q3_I-D':q3d,'Min_I-D':diff.min(),'Max_I-D':diff.max(),
                'Mediana_absoluta':np.median(abs(diff)),
                'p_t_pareada':stats.ttest_1samp(diff,0).pvalue if sd>0 else np.nan})
    ag=pd.DataFrame(asym_group)
    if not ag.empty:
        ag['p_BH_t_pareada']=ag.groupby('Grupo',sort=False).p_t_pareada.transform(lambda p:adjust(p))
    tabs['Asimetria_por_grupo']=ag
    if not ag.empty:
        tabs['Asimetria_n_pares']=ag.pivot(index='Par',columns='Grupo',values='n_pares').reindex(index=PAIRS,columns=GROUPS).reset_index()
    tabs['Bilaterales_con_asimetria']=bilateral_table_with_asymmetry(desc,ag.reset_index() if not ag.empty else ag)
    # Error: selección de modo ERROR en S2; nunca sustituir S1 por Resultados.
    def valid(s):return s[s.Landmark.isin(LANDMARKS)&s.Status.eq('SET')&np.isfinite(s.Thickness_mm)&s.Thickness_mm.gt(0)]
    paired=valid(s1)[['CaseID','Landmark','Thickness_mm']].merge(valid(s2)[['CaseID','Landmark','Thickness_mm']],on=['CaseID','Landmark'],suffixes=('_S1','_S2'),validate='one_to_one')
    tabs['Pares_error']=paired;err=[]
    for lm,g in paired.groupby('Landmark',sort=False):
        z=g[['Thickness_mm_S1','Thickness_mm_S2']].to_numpy();n=len(z)
        if n<3:continue
        diff=z[:,1]-z[:,0];tem=np.sqrt(np.mean(diff**2)/2);mean=z.mean();sd=diff.std(ddof=1)
        boot=np.array([icc_a1(z[rng.integers(n,size=n)]) for _ in range(b)]);boot=boot[np.isfinite(boot)]
        ci=np.quantile(boot,[.025,.975]) if len(boot) else [np.nan,np.nan]
        err.append({'Landmark':lm,'n_pares':n,'Sesgo_S2-S1':diff.mean(),'TEM_mm':tem,'rTEM_pct':100*tem/mean,
            'ICC_A1':icc_a1(z),'ICC_IC95_inf':ci[0],'ICC_IC95_sup':ci[1],'Bootstrap_valido':len(boot),
            'LoA_inf':diff.mean()-1.96*sd,'LoA_sup':diff.mean()+1.96*sd})
    tabs['Error_intra']=pd.DataFrame(err)
    # CCU: modelo de medias con edad centrada y CCU por 10 mm, no nutrición ni causalidad.
    nr=neck[neck.CCU_valida].copy();ccu=[];ccdiag=[]
    for lm in LANDMARKS:
        d=main[main.Landmark.eq(lm)][['CaseID','Thickness_mm']].merge(nr,on='CaseID')
        if len(d)<20:
            ccu.append({'Landmark':lm,'n':len(d),'p':np.nan});continue
        X=np.column_stack([np.ones(len(d)),d.CCU_mm/10,d.Edad-d.Edad.mean(),d.Sexo.eq('M').astype(float)])
        try:
            fit=fit_hc3(X,d.Thickness_mm)
            row={'Landmark':lm,**contrast(fit,[0,1,0,0]),'R2':fit['R2'],'Unidad':'mm espesor / 10 mm CCU',
                 'n_fuente_revisar':int(d.Revision_fuente.sum())};ccu.append(row)
            for j,r in enumerate(d.itertuples()):ccdiag.append({'Landmark':lm,'CaseID':r.CaseID,'CCU_mm':r.CCU_mm,
                'Residuo':fit['resid'][j],'Ajustado':fit['fitted'][j],'Cook':fit['cook'][j],
                'Alerta_Cook':fit['cook'][j]>4/fit['n']})
        except ValueError as ex:
            ccu.append({'Landmark':lm,'n':len(d),'p':np.nan});qc.append({'Tipo':'CCU_modelo','CaseID':'','Detalle':lm+str(ex)})
    cc=pd.DataFrame(ccu);cc['p_BH']=adjust(cc.p);tabs['CCU_HC3']=cc;tabs['CCU_diagnosticos']=pd.DataFrame(ccdiag)
    # Secundario: usa el punto de corte específico sexo×edad como variable binaria.
    # No reemplaza CCU continua: cuantifica la diferencia ajustada entre cribado positivo y negativo.
    nrc=neck[neck.CCU_clasificable_Fuentes].copy();ccucat=[]
    for lm in LANDMARKS:
        d=main[main.Landmark.eq(lm)][['CaseID','Thickness_mm']].merge(nrc,on='CaseID')
        if len(d)<20 or d.CCU_cribado_positivo.nunique()<2:
            ccucat.append({'Landmark':lm,'n':len(d),'p':np.nan});continue
        X=np.column_stack([np.ones(len(d)),d.CCU_cribado_positivo.astype(float),
                           d.Edad-d.Edad.mean(),d.Sexo.eq('M').astype(float)])
        try:
            fit=fit_hc3(X,d.Thickness_mm)
            row={'Landmark':lm,**contrast(fit,[0,1,0,0]),'R2':fit['R2'],
                 'Comparacion':'Cribado positivo − negativo (Fuentes 2018)',
                 'Unidad':'diferencia media ajustada de FSTT (mm)'};ccucat.append(row)
        except ValueError as ex:
            ccucat.append({'Landmark':lm,'n':len(d),'p':np.nan});qc.append({'Tipo':'CCU_categoria_modelo','CaseID':'','Detalle':lm+': '+str(ex)})
    ccat=pd.DataFrame(ccucat);ccat['p_BH']=adjust(ccat.p);tabs['CCU_Categoria_HC3']=ccat
    return tabs,main


def figures(tabs,main,out):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    plt.rcParams.update({'font.family':'serif','font.serif':['Times New Roman','DejaVu Serif'],
        'font.size':11,'axes.spines.top':False,'axes.spines.right':False,'axes.titleweight':'normal'})
    out.mkdir(exist_ok=True);colors=['#B87387','#793D64','#73AFC1','#286482']
    def save(fig,name):
        fig.savefig(out/(name+'.png'),dpi=300,bbox_inches='tight');plt.close(fig)
    for lm in LANDMARKS:
        d=main[main.Landmark.eq(lm)];fig,ax=plt.subplots(figsize=(7,4.8));xs=[d.loc[d.Grupo.eq(g),'Thickness_mm'] for g in GROUPS]
        bp=ax.boxplot(xs,patch_artist=True,showmeans=True,meanprops={'marker':'D','markerfacecolor':'white','markeredgecolor':'black','markersize':4},medianprops={'color':'black','linewidth':1.5})
        for box,color in zip(bp['boxes'],colors):box.set_facecolor(color);box.set_alpha(.7)
        ax.set_xticks(range(1,5),[g.replace('_',' ')+'\nn='+str(len(x)) for g,x in zip(GROUPS,xs)])
        ax.set(ylabel='Espesor tisular (mm)',title=lm.upper());fig.text(.12,.01,'Línea: mediana · Rombo: media · Caja: Q1–Q3 · Bigotes: 1,5 RIC',fontsize=9)
        fig.tight_layout(rect=[0,.05,1,1]);save(fig,'Espesor_'+lm)
    def forest(df,label,estimate,low,high,name,xlabel):
        if df.empty or estimate not in df:return
        d=df.dropna(subset=[estimate,low,high]);fig,ax=plt.subplots(figsize=(7,max(4,len(d)*.28)))
        y=np.arange(len(d));ax.hlines(y,d[low],d[high],color='#286482');ax.scatter(d[estimate],y,color='#286482',s=20)
        ax.set_yticks(y,d[label].str.upper());ax.invert_yaxis();ax.axvline(0,color='.6',lw=.8);ax.set_xlabel(xlabel);fig.tight_layout();save(fig,name)
    for effect,g in tabs['Factorial_HC3'].groupby('Efecto'):
        forest(g,'Landmark','Estimacion_mm','IC95_inf','IC95_sup','Factorial_'+effect,'Diferencia media (mm); IC95 % individual')
    forest(tabs['CCU_HC3'],'Landmark','Estimacion_mm','IC95_inf','IC95_sup','CCU','Cambio medio por 10 mm de CCU; IC95 % individual')
    forest(tabs['Error_intra'],'Landmark','ICC_A1','ICC_IC95_inf','ICC_IC95_sup','ICC','ICC(A,1); IC95 % bootstrap')
    forest(tabs['Asimetria'],'Par','Media_I-D','IC95_media_inf','IC95_media_sup','Asimetria','Media izquierda − derecha (mm); IC95 % bootstrap')
    # Diagnósticos gráficos necesarios para evaluar influencia, forma y linealidad de CCU.
    for lm in LANDMARKS:
        d=tabs['Diagnosticos'];d=d[d.Landmark.eq(lm)]
        cd=tabs['CCU_diagnosticos'];cd=cd[cd.Landmark.eq(lm)] if not cd.empty else cd
        fig,ax=plt.subplots(1,3,figsize=(12,3.5))
        if len(d):
            stats.probplot(d.Residuo,plot=ax[0]);ax[0].set_title('Residuos factorial: Q–Q')
            ax[1].scatter(d.Ajustado,d.Residuo,s=10,alpha=.5);ax[1].axhline(0,color='gray');ax[1].set(xlabel='Ajustado',ylabel='Residuo',title='Modelo factorial')
        if len(cd):
            ax[2].scatter(cd.CCU_mm,cd.Residuo,s=10,alpha=.5);ax[2].axhline(0,color='gray');ax[2].set(xlabel='CCU (mm)',ylabel='Residuo',title='Modelo CCU')
        fig.suptitle(lm.upper());fig.tight_layout();save(fig,'Diagnostico_'+lm)
    for lm,d in tabs['Pares_error'].groupby('Landmark'):
        x=(d.Thickness_mm_S1+d.Thickness_mm_S2)/2;y=d.Thickness_mm_S2-d.Thickness_mm_S1
        fig,ax=plt.subplots(figsize=(6,4));ax.scatter(x,y,s=20,alpha=.7)
        for h in [y.mean(),y.mean()-1.96*y.std(ddof=1),y.mean()+1.96*y.std(ddof=1)]:ax.axhline(h,color='#793D64',ls='--')
        ax.set(xlabel='Promedio de las dos sesiones (mm)',ylabel='S2 − S1 (mm)',title='Bland–Altman: '+lm.upper());fig.tight_layout();save(fig,'BlandAltman_'+lm)


def main():
    parser=argparse.ArgumentParser(description=__doc__,formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('excel',type=Path);parser.add_argument('--salida',type=Path)
    parser.add_argument('--bootstrap',type=int,default=2000);parser.add_argument('--seed',type=int,default=27082026)
    parser.add_argument('--observador',default='OBS1');parser.add_argument('--sesion-principal',default='S1')
    parser.add_argument('--sesion-error',default='S2');parser.add_argument('--replica',type=int,default=1)
    parser.add_argument('--sin-figuras',action='store_true');parser.add_argument('--decimales-diferencia',type=int,default=4,
        help='Redondeo de diferencias exclusivamente para rangos de Wilcoxon; documentar resolución')
    args=parser.parse_args()
    if args.bootstrap<100:parser.error('--bootstrap debe ser >=100; para resultados finales usar 2000 o más')
    path=args.excel.resolve()
    if not path.exists():parser.error('No existe el Excel: '+str(path))
    stamp=datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    out=args.salida.resolve() if args.salida else path.parent/('ANALISIS_HC3_'+stamp)
    if out.exists() and any(out.iterdir()):parser.error('La carpeta de salida no está vacía; usar otra para no sobrescribir')
    out.mkdir(parents=True,exist_ok=True)
    print('Leyendo base y comprobando fuentes...',flush=True)
    c,base,s1,s2,neck,qc,discard=load_data(path,args.observador.upper(),args.sesion_principal.upper(),args.sesion_error.upper(),args.replica)
    print(f'{len(c)} sujetos; {int(base.Valido.sum())} espesores; {int(neck.CCU_valida.sum())} CCU.',flush=True)
    print('Calculando estadísticos, HC3 y bootstrap...',flush=True)
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter('always');tabs,main_data=analyze(c,base,s1,s2,neck,args,qc)
    for message in caught:qc.append({'Tipo':'Advertencia_numerica','CaseID':'','Detalle':str(message.message)})
    tabs['QC']=pd.DataFrame(qc,columns=['Tipo','CaseID','Detalle'])
    tabs['Seleccion_registros']=pd.DataFrame(discard,columns=['CaseID','Landmark','Sesion','Fila_Excel','Motivo'])
    metadata={'Entrada':str(path),'SHA256':hashlib.sha256(path.read_bytes()).hexdigest(),'Fecha':stamp,
      'Python':platform.python_version(),'numpy':np.__version__,'pandas':pd.__version__,'scipy':scipy.__version__,
      'Observador':args.observador,'Sesion_principal':args.sesion_principal,'Sesion_error':args.sesion_error,'Replica':args.replica,
      'Bootstrap':args.bootstrap,'Semilla':args.seed,'HC3':'t con gl=n-p; cuatro medias sin intercepto',
      'Efectos':'BH separado: sexo, edad, interacción; 26 por familia',
      'Contrastes':'Efectos simples planificados; Holm principal dentro de cada familia de 26 landmarks; Holm 104 solo suplementario',
      'CCU':'Principal continua: Resultados; edad continua + sexo; beta por 10 mm; BH 26',
      'CCU_categoria':'Secundaria: corte sexo×edad Fuentes 2018 (7–18); positivo vs negativo; ajusta edad y sexo; BH 26',
      'Asimetria':'Pares completos por cada par, nunca exige 26 LM completos; global y sexo×edad. Bilaterales se exportan como filas Izquierda/Derecha/Asimetría I-D',
      'Seleccion_registros':'Por modo/sesión/observador/réplica; Timestamp es solo bitácora y nunca elige registros',
      'Outliers':'No hay exclusión automática ni análisis sin outliers; Cook/residuos son diagnósticos de revisión',
      'IC':'Individuales (no simultáneos); bootstrap de individuos, sin imputación',
      'Wilcoxon_decimales_diferencias':args.decimales_diferencia}
    ordered={'LEEME':pd.DataFrame(metadata.items(),columns=['Parametro','Valor']),**tabs}
    print('Exportando Excel...',flush=True)
    from openpyxl.styles import Font,PatternFill,Alignment
    with pd.ExcelWriter(out/'RESULTADOS_HC3.xlsx',engine='openpyxl') as writer:
        for name,df in ordered.items():
            df.to_excel(writer,sheet_name=name,index=False)
            ws=writer.sheets[name];ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:cell.font=Font(name='Calibri',bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='286482');cell.alignment=Alignment(wrap_text=True)
            ws.row_dimensions[1].height=32
            for col in ws.columns:
                letter=col[0].column_letter;ws.column_dimensions[letter].width=min(55,max(14,len(str(col[0].value))+2))
                for cell in col[1:]:
                    if isinstance(cell.value,float):
                        cell.number_format='0.000E+00' if str(col[0].value).startswith('p') and 0 < abs(cell.value) < .0001 else '0.0000'
    if not args.sin_figuras:
        print('Generando figuras principales...',flush=True)
        figures(tabs,main_data,out/'FIGURAS')
    note=f'''ANÁLISIS FSTT HC3
Sujetos: {len(c)}. Mediciones válidas: {int(base.Valido.sum())}. CCU válidas: {int(neck.CCU_valida.sum())}.
Consultar LEEME, QC, Seleccion_registros y CCU_fuente antes de interpretar.
La medición principal se identifica como NORMAL/S1; la reproducibilidad como ERROR/S2.
Timestamp se conserva solo como bitácora y nunca se usa para escoger entre registros.
No se eliminan outliers automáticamente; Cook y residuos son diagnósticos para revisión.
La interacción sexo×edad se interpreta junto a cuatro efectos simples, con Holm dentro de cada familia de 26 landmarks.
La asimetría usa únicamente pares izquierda-derecha completos para cada par bilateral.
CCU se analiza como covariable secundaria; no es requisito para aplicar la tabla de referencia forense.
Las cifras no sustituyen la revisión anatómica o de trazabilidad de casos señalados en QC.
'''
    (out/'LEEME.txt').write_text(note,encoding='utf-8')
    print(note);print('Salida:',out)


if __name__=='__main__':
    main()
