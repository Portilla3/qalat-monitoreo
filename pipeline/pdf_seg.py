"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   SCRIPT_TOP_Universal_PDF_Seguimiento.py                                  ║
║   Informe PDF comparativo TOP1 (Ingreso) vs TOP2 (Seguimiento)             ║
║   11 gráficos · 6 páginas · Compatible con cualquier país TOP              ║
║   Versión Universal 1.0                                                    ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                             ║
║  CÓMO USAR LA PRÓXIMA VEZ:                                                 ║
║  1. Abre un chat nuevo con Claude                                           ║
║  2. Sube DOS archivos:                                                      ║
║       • Este script: SCRIPT_TOP_Universal_PDF_Seguimiento.py               ║
║       • La base en formato Wide (generada por SCRIPT_TOP_Universal_Wide)   ║
║  3. Escribe exactamente:                                                    ║
║     "Ejecuta el script universal PDF Seguimiento con esta base Wide"       ║
║  4. Claude ajustará NOMBRE_SERVICIO y PERIODO según corresponda            ║
║                                                                             ║
║  ESTRUCTURA DEL PDF (6 páginas):                                           ║
║    Pág 1 – Portada                                                         ║
║    Pág 2 – Presentación + KPIs + G1 Sexo                                  ║
║    Pág 3 – G2 Edad + G_torta sust. principal + G3 barras sust. comp.      ║
║    Pág 4 – G4 Días consumo + G5 Cambio en consumo (apilado)               ║
║    Pág 5 – G6 % Consumidores + G7 Días por sust. + G8 Transgresión        ║
║    Pág 6 – G9 Tipos transgresión + G10 Salud + G11 Vivienda               ║
║                                                                             ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""


import glob, os, unicodedata


# ── Detección automática de país ─────────────────────────────────────────────
_PAISES = {
    'republica_dominicana':'República Dominicana', 'repdomini':'República Dominicana',
    'dominicana':'República Dominicana', 'honduras':'Honduras',
    'panama':'Panamá', 'panam':'Panamá', 'el_salvador':'El Salvador',
    'salvador':'El Salvador', 'mexico':'México', 'mexic':'México',
    'ecuador':'Ecuador', 'peru':'Perú', 'argentina':'Argentina',
    'colombia':'Colombia', 'chile':'Chile', 'bolivia':'Bolivia',
    'paraguay':'Paraguay', 'uruguay':'Uruguay', 'venezuela':'Venezuela',
    'guatemala':'Guatemala', 'costa_rica':'Costa Rica',
    'costarica':'Costa Rica', 'nicaragua':'Nicaragua',
}
def _extraer_pais(filename):
    fn = _norm(str(filename).replace('.','_'))
    for key, nombre in _PAISES.items():
        if key in fn: return nombre
    return None

def _detectar_pais(wide_file):
    import pandas as _pd
    try:
        rs = _pd.read_excel(wide_file, sheet_name='Resumen', header=None)
        for _, row in rs.iterrows():
            for v in row.tolist():
                p = _extraer_pais(str(v))
                if p: return p
    except: pass
    return _extraer_pais(os.path.basename(wide_file))

def _norm(s):
    return unicodedata.normalize('NFD', str(s).lower()).encode('ascii','ignore').decode()

def auto_archivo_wide():
    """Encuentra automáticamente la base Wide subida al chat"""
    candidatos = (
        glob.glob('/mnt/user-data/uploads/*Wide*.xlsx') +
        glob.glob('/mnt/user-data/uploads/*wide*.xlsx') +
        glob.glob('/mnt/user-data/uploads/TOP_Base*.xlsx') +
        glob.glob('/home/claude/TOP_Base_Wide.xlsx'))
    if not candidatos:
        raise FileNotFoundError(
            "\n\u26a0  No se encontró la base Wide.\n"
            "   Sube el archivo TOP_Base_Wide.xlsx junto con este script.")
    print(f"  \u2192 Base Wide detectada: {os.path.basename(candidatos[0])}")
    return candidatos[0]

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN — Claude ajusta NOMBRE_SERVICIO y PERIODO según corresponda
# ══════════════════════════════════════════════════════════════════════════════

INPUT_FILE      = auto_archivo_wide()   # ← detecta automáticamente
SHEET_NAME      = 'Base Wide'
OUTPUT_FILE     = '/home/claude/TOP_Informe_Seguimiento.pdf'
# ── FILTRO OPCIONAL POR CENTRO ────────────────────────────────────────────────
# Dejar en None para procesar TODOS los centros.
# Poner el código exacto del centro para filtrar solo ese centro.
# Ejemplos:
#   FILTRO_CENTRO = None         ← todos los centros
#   FILTRO_CENTRO = "HCHN01"     ← solo ese centro
FILTRO_CENTRO = None
# ─────────────────────────────────────────────────────────────────────────────

# ── Auto-detección de país, período y nombre de servicio ────────────────────
_pais_detectado = _detectar_pais(INPUT_FILE)

# Si hay filtro de centro activo:
if 'FILTRO_CENTRO' in dir() and FILTRO_CENTRO:
    NOMBRE_SERVICIO = (f'{_pais_detectado}  —  Centro {FILTRO_CENTRO}'
                       if _pais_detectado else f'Centro {FILTRO_CENTRO}')
else:
    NOMBRE_SERVICIO = _pais_detectado if _pais_detectado else 'Servicio de Tratamiento'

# Período desde el archivo Wide (hoja Resumen)
_periodo_auto = None
try:
    import pandas as _pd2
    _rs = _pd2.read_excel(INPUT_FILE, sheet_name='Resumen', header=None)
    for _, _row in _rs.iterrows():
        for _v in _row.tolist():
            if 'Período' in str(_v) or 'periodo' in str(_v).lower():
                continue
            if '–' in str(_v) or (' ' in str(_v) and any(
                    m in str(_v) for m in ['Enero','Feb','Mar','Abr','May','Jun',
                                           'Jul','Ago','Sep','Oct','Nov','Dic','2025','2026'])):
                _periodo_auto = str(_v).strip()
                break
        if _periodo_auto: break
except: pass

PERIODO = _periodo_auto if _periodo_auto else '2025 – 2026'   # ← fallback
# ─────────────────────────────────────────────────────────────────────────────

# ══════════════════════════════════════════════════════════════════════════════
import pandas as pd, numpy as np, io, warnings
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
warnings.filterwarnings('ignore')

def _es_positivo(valor):
    s = str(valor).strip().lower()
    if s in ('sí', 'si'): return True
    if s in ('no', 'no aplica', 'nunca', 'nan', ''): return False
    n = pd.to_numeric(valor, errors='coerce')
    return not pd.isna(n) and n > 0

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, white
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image,
                                 Table, TableStyle, PageBreak, HRFlowable)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY

W, H = A4
C_DARK  = HexColor('#1F3864'); C_MID  = HexColor('#2E75B6')
C_LIGHT = HexColor('#BDD7EE'); C_GRAY = HexColor('#595959')
MC_MID = '#2E75B6'; MC_LIGHT = '#BDD7EE'; MC_ACCENT = '#00B0F0'
C_T1 = '#2E75B6'; C_T2 = '#00B0F0'                     # Ingreso vs Seguimiento
C_ABS2 = '#1F3864'; C_DIS2 = '#2E75B6'; C_SC2 = '#9DC3E6'; C_EMP2 = '#BDD7EE'
PIE_COLS = ['#2E75B6','#1F3864','#4472C4','#9DC3E6','#00B0F0','#538135','#BFBFBF','#C00000','#ED7D31']

# ══════════════════════════════════════════════════════════════════════════════
# DETECCIÓN DINÁMICA DE COLUMNAS  (_TOP1 y _TOP2)
# ══════════════════════════════════════════════════════════════════════════════
def detectar_columnas(cols):
    col_set = set(cols)

    def par(c1):
        c2 = c1.replace('_TOP1', '_TOP2')
        return (c1, c2 if c2 in col_set else None)

    # Sustancias días: "1) ... >> Nombre (unidad) >> Total (0-28)_TOP1"
    sust_cols = []
    for c in cols:
        if c.endswith('_TOP1') and 'Total (0-28)' in c:
            base = c.replace('_TOP1', '')
            if base.startswith('1)'):
                partes = base.split('>>')
                if len(partes) >= 3:
                    nombre = partes[-2].strip().split('(')[0].strip()
                    c1, c2 = par(c)
                    sust_cols.append((nombre, c1, c2))
    print(f'  Sustancias: {[s[0] for s in sust_cols]}')

    # Transgresión Sí/No
    tr_sn = []
    for c in cols:
        if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('3)') and '>>' in c:
            nombre = c.replace('_TOP1','').split('>>')[-1].strip()
            c1, c2 = par(c)
            tr_sn.append((nombre, c1, c2))
    print(f'  Transgresión: {[t[0] for t in tr_sn]}')

    vif     = par(next((c for c in cols if c.endswith('_TOP1') and '4)' in c
                        and 'Violencia Intrafamiliar' in c and 'Total (0-28)' in c), None) or '')
    sal_psi = par(next((c for c in cols if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('6)')), None) or '')
    sal_fis = par(next((c for c in cols if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('8)')), None) or '')
    cal_vid = par(next((c for c in cols if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('10)')), None) or '')
    viv1    = par(next((c for c in cols if c.endswith('_TOP1') and '9)' in c and 'estable' in c.lower()), None) or '')
    viv2    = par(next((c for c in cols if c.endswith('_TOP1') and '9)' in c and 'básicas' in c.lower()), None) or '')
    sust_pp = par(next((c for c in cols if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('2)')
                        and 'sustancia principal' in c.lower()), None) or '')
    sexo    = next((c for c in cols if c.endswith('_TOP1') and 'sexo' in c.lower()), None)
    fn_col  = next((c for c in cols if c.endswith('_TOP1') and 'nacimiento' in c.lower()), None)
    fecha   = next((c for c in cols if c.endswith('_TOP1') and 'fecha entrevista' in c.lower()), None)

    # limpiar pares donde c1 es cadena vacía (cuando next devuelve None)
    def safe(p): return p if (p[0] and p[0] in set(cols)) else (None, None)

    return dict(sust_cols=sust_cols, tr_sn=tr_sn,
                vif=safe(vif), sal_psi=safe(sal_psi), sal_fis=safe(sal_fis), cal_vid=safe(cal_vid),
                viv1=safe(viv1), viv2=safe(viv2), sust_pp=safe(sust_pp),
                sexo=sexo, fn_col=fn_col, fecha=fecha)

# ══════════════════════════════════════════════════════════════════════════════
# NORMALIZACIÓN SUSTANCIA PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def norm_sust(s):
    if pd.isna(s) or str(s).strip() == '0': return None
    s = str(s).strip().lower()
    if any(x in s for x in ['alcohol','cerveza','licor','aguard','alchol','bebida']): return 'Alcohol'
    if any(x in s for x in ['marihu','marjhu','marhuana','cannabis','cannbis']): return 'Cannabis/\nMarihuana'
    if any(x in s for x in ['pasta base','pasta','papelillo']): return 'Pasta\nBase'
    if any(x in s for x in ['crack','cristal','piedra','paco']): return 'Crack/\nCristal'
    if any(x in s for x in ['cocain','perico','coca ']): return 'Cocaína'
    if any(x in s for x in ['tabaco','cigarr','nicot']): return 'Tabaco'
    if any(x in s for x in ['sedant','benzod','tranqui']): return 'Sedantes'
    if any(x in s for x in ['opiod','heroina','morfin','fentanil']): return 'Opiáceos'
    if any(x in s for x in ['metanfet','anfetam']): return 'Metanfetamina'
    return 'Otras'

# ══════════════════════════════════════════════════════════════════════════════
# CARGA Y CÁLCULO DE DATOS
# ══════════════════════════════════════════════════════════════════════════════
def cargar_datos():
    print(f'  Leyendo: {INPUT_FILE}')
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME, header=1)

    # Aplicar filtro de centro si corresponde
    _col_centro = next((c for c in df.columns if any(x in _norm(c) for x in
                        ['codigo del centro', 'servicio de tratamiento',
                         'centro/ servicio', 'codigo centro'])), None)
    if FILTRO_CENTRO and _col_centro:
        n_antes = len(df)
        df = df[df[_col_centro].astype(str).str.strip() == FILTRO_CENTRO].copy()
        df = df.reset_index(drop=True)
        print(f'  ⚑ Filtro activo: Centro = "{FILTRO_CENTRO}"')
        print(f'    {n_antes} pacientes totales → {len(df)} del centro seleccionado')
    if FILTRO_CENTRO:
        global OUTPUT_FILE, NOMBRE_SERVICIO
        OUTPUT_FILE = f'/home/claude/TOP_Informe_Seguimiento_{FILTRO_CENTRO}.pdf'
        _pais_local = _detectar_pais(INPUT_FILE)
        NOMBRE_SERVICIO = (f'{_pais_local}  —  Centro {FILTRO_CENTRO}'
                           if _pais_local else f'Centro {FILTRO_CENTRO}')
    N_total = len(df)
    seg = df[df['Tiene_TOP2'] == 'Sí'].copy().reset_index(drop=True)
    N_seg = len(seg)
    print(f'  Total pacientes:        {N_total}')
    print(f'  Con seguimiento (TOP2): {N_seg}  ({round(N_seg/N_total*100,1)}%)')

    # ── Tiempo de seguimiento ────────────────────────────────────────────────
    _fc1 = next((c for c in seg.columns if 'fecha entrevista' in c.lower() and c.endswith('_TOP1')), None)
    _fc2 = next((c for c in seg.columns if 'fecha entrevista' in c.lower() and c.endswith('_TOP2')), None)
    seg_tiempo = {'mediana': None, 'min': None, 'max': None, 'n': 0, 'n_total': N_seg}
    if _fc1 and _fc2:
        _d1 = pd.to_datetime(seg[_fc1], errors='coerce')
        _d2 = pd.to_datetime(seg[_fc2], errors='coerce')
        _dias = (_d2 - _d1).dt.days
        _dias_ok = _dias[(_dias >= 0) & (_dias <= 730)].dropna()
        if len(_dias_ok) > 0:
            _m = _dias_ok / 30.44
            seg_tiempo = {
                'mediana': round(float(_m.median()), 1),
                'min':     round(float(_m.min()), 1),
                'max':     round(float(_m.max()), 1),
                'n':       len(_dias_ok),
                'n_total': int(_dias.notna().sum())
            }

    DC = detectar_columnas(seg.columns.tolist())
    R  = {'N_total': N_total, 'N_seg': N_seg, 'DC': DC, 'seg_tiempo': seg_tiempo}

    # Sexo
    if DC['sexo']:
        sc = seg[DC['sexo']].astype(str).str.strip().str.upper()
        nv = int(sc.isin(['H','M']).sum())
        R['n_hombre']   = int((sc=='H').sum()); R['n_mujer']  = int((sc=='M').sum())
        R['nv_sex']     = nv
        R['pct_hombre'] = round(R['n_hombre']/nv*100,1) if nv>0 else 0
        R['pct_mujer']  = round(R['n_mujer'] /nv*100,1) if nv>0 else 0
    else:
        R['n_hombre']=R['n_mujer']=R['nv_sex']=0; R['pct_hombre']=R['pct_mujer']=0

    # Edad
    if DC['fn_col'] and DC['fecha']:
        fn   = pd.to_datetime(seg[DC['fn_col']], errors='coerce')
        ref  = pd.to_datetime(seg[DC['fecha']], errors='coerce').fillna(pd.Timestamp.now())
        edad = ((ref-fn).dt.days/365.25).round(1); edad=edad[(edad>=10)&(edad<=100)]
        R['edad_media']=round(float(edad.mean()),1); R['edad_sd']=round(float(edad.std()),1)
        R['nv_edad']=int(edad.notna().sum()); R['edad_min']=int(edad.min()) if R['nv_edad']>0 else 0; R['edad_max']=int(edad.max()) if R['nv_edad']>0 else 0
        bins=[0,17,30,40,50,60,200]; labs=['Menos de 18','18 a 30','31 a 40','41 a 50','51 a 60','61 o más']
        ec=pd.cut(edad,bins=bins,labels=labs); R['edad_dist']={l:int((ec==l).sum()) for l in labs}
    else:
        R['edad_media']=R['edad_sd']=0; R['edad_min']=R['edad_max']=0
        R['nv_edad']=0; R['edad_dist']={'Sin datos': 0}

    # Sustancia principal TOP1 vs TOP2
    c1_sp, c2_sp = DC['sust_pp']
    if c1_sp:
        sr1 = seg[c1_sp].apply(norm_sust)
        sr2 = seg[c2_sp].apply(norm_sust) if c2_sp else pd.Series([None]*N_seg)
        R['nv_sust1'] = int(sr1.notna().sum()); R['nv_sust2'] = int(sr2.notna().sum())
        cats = ['Alcohol','Cannabis/\nMarihuana','Pasta\nBase','Cocaína','Crack/\nCristal',
                'Tabaco','Sedantes','Opiáceos','Metanfetamina','Otras']
        sust_comp = []
        for cat in cats:
            n1=int((sr1==cat).sum()); n2=int((sr2==cat).sum())
            if n1>0 or n2>0:
                sust_comp.append({'label': cat.replace('\n',' '), 'n1': n1, 'n2': n2,
                    'p1': round(n1/R['nv_sust1']*100,1) if R['nv_sust1']>0 else 0,
                    'p2': round(n2/R['nv_sust2']*100,1) if R['nv_sust2']>0 else 0})
        R['sust_comp'] = sust_comp
        top1 = max(sust_comp, key=lambda x: x['n1']) if sust_comp else {'label':'—','p1':0}
        R['sust_top1'] = top1['label']; R['sust_top1_pct'] = top1['p1']
    else:
        R['nv_sust1']=R['nv_sust2']=0; R['sust_comp']=[]; R['sust_top1']='—'; R['sust_top1_pct']=0

    # Días de consumo TOP1 vs TOP2 (G4, G7)
    dias_comp = []
    for lbl, c1, c2 in DC['sust_cols']:
        v1 = pd.to_numeric(seg[c1], errors='coerce')
        v2 = pd.to_numeric(seg[c2], errors='coerce') if c2 else pd.Series([np.nan]*N_seg)
        m1 = round(float(v1.mean()),1) if v1.notna().sum()>0 else 0
        m2 = round(float(v2.mean()),1) if (c2 and v2.notna().sum()>0) else 0
        if m1>0 or m2>0:
            dias_comp.append({'label':lbl,'m1':m1,'m2':m2,
                              'nv1':int(v1.notna().sum()), 'nv2':int(v2.notna().sum()) if c2 else 0})
    R['dias_comp'] = dias_comp

    # % consumidores por sustancia TOP1 vs TOP2 (G6)
    cons_pct = []
    for lbl, c1, c2 in DC['sust_cols']:
        v1 = pd.to_numeric(seg[c1], errors='coerce')
        v2 = pd.to_numeric(seg[c2], errors='coerce') if c2 else pd.Series([0]*N_seg)
        n1 = int((v1>0).sum()); n2 = int((v2>0).sum()) if c2 else 0
        if n1>0 or n2>0:
            cons_pct.append({'label':lbl,'n1':n1,'n2':n2,
                'p1':round(n1/N_seg*100,1), 'p2':round(n2/N_seg*100,1) if c2 else 0})
    R['cons_pct'] = cons_pct

    # Cambio en consumo (G5)
    cambio = []
    for lbl, c1, c2 in DC['sust_cols']:
        if not c2: continue
        v1 = pd.to_numeric(seg[c1], errors='coerce').fillna(0)
        v2 = pd.to_numeric(seg[c2], errors='coerce').fillna(0)
        mask = v1>0; n_cons = int(mask.sum())
        if n_cons < 2: continue
        s1=v1[mask]; s2=v2[mask]
        n_abs=int((s2==0).sum()); n_dis=int(((s2>0)&(s2<s1)).sum())
        n_sc=int((s2==s1).sum()); n_emp=int((s2>s1).sum())
        pct=lambda n: round(n/n_cons*100,1) if n_cons>0 else 0
        cambio.append({'label':lbl,'n_cons':n_cons,
            'pct_abs':pct(n_abs),'pct_dis':pct(n_dis),
            'pct_sc':pct(n_sc),'pct_emp':pct(n_emp)})
    R['cambio'] = cambio

    # Salud TOP1 vs TOP2 (G10)
    salud = []
    for lbl, (c1, c2) in [('Salud Psicológica', DC['sal_psi']),
                           ('Salud Física',      DC['sal_fis']),
                           ('Calidad de Vida',   DC['cal_vid'])]:
        if not c1: continue
        v1 = pd.to_numeric(seg[c1], errors='coerce')
        v2 = pd.to_numeric(seg[c2], errors='coerce') if c2 else pd.Series([np.nan]*N_seg)
        salud.append({'label':lbl,
            'm1':round(float(v1.mean()),1), 'm2':round(float(v2.mean()),1) if c2 else 0,
            'nv1':int(v1.notna().sum()), 'nv2':int(v2.notna().sum()) if c2 else 0})
    R['salud'] = salud

    # Vivienda TOP1 vs TOP2 (G11)
    def viv(col, df_):
        if not col or col not in df_.columns: return (0, 0, N_seg)
        nv = int(df_[col].isin(['Sí','No']).sum()) or N_seg
        n  = int((df_[col]=='Sí').sum())
        return n, round(n/nv*100,1), nv
    c_viv1_1, c_viv1_2 = DC['viv1']; c_viv2_1, c_viv2_2 = DC['viv2']
    R['viv1_t1'] = viv(c_viv1_1, seg); R['viv1_t2'] = viv(c_viv1_2, seg)
    R['viv2_t1'] = viv(c_viv2_1, seg); R['viv2_t2'] = viv(c_viv2_2, seg)

    # Transgresión TOP1 vs TOP2 (G8, G9)
    tr_cols1 = [c1 for _,c1,_ in DC['tr_sn']]; tr_cols2 = [c2 for _,_,c2 in DC['tr_sn']]
    vif_c1, vif_c2 = DC['vif']
    def has_tr(row, sn_cols, vif_col):
        for c in sn_cols:
            if c and _es_positivo(row.get(c,'')): return True
        if vif_col:
            v = pd.to_numeric(row.get(vif_col, np.nan), errors='coerce')
            return not np.isnan(v) and v>0
        return False
    tr1 = seg.apply(lambda r: int(has_tr(r, tr_cols1, vif_c1)), axis=1)
    tr2 = seg.apply(lambda r: int(has_tr(r, tr_cols2, vif_c2)), axis=1)
    R['n_tr1']=int(tr1.sum()); R['n_tr2']=int(tr2.sum())
    R['pct_tr1']=round(R['n_tr1']/N_seg*100,1); R['pct_tr2']=round(R['n_tr2']/N_seg*100,1)
    tipos = []
    for lbl, c1, c2 in DC['tr_sn']:
        n1=int(seg[c1].apply(_es_positivo).sum()) if c1 else 0
        n2=int(seg[c2].apply(_es_positivo).sum()) if c2 else 0
        tipos.append({'label':lbl,'n1':n1,'n2':n2,
            'p1':round(n1/N_seg*100,1),'p2':round(n2/N_seg*100,1)})
    if vif_c1:
        vif1_v=pd.to_numeric(seg[vif_c1],errors='coerce'); n_v1=int((vif1_v>0).sum())
        vif2_v=pd.to_numeric(seg[vif_c2],errors='coerce') if vif_c2 else pd.Series([np.nan]*N_seg)
        n_v2=int((vif2_v>0).sum())
        tipos.append({'label':'VIF','n1':n_v1,'n2':n_v2,
            'p1':round(n_v1/N_seg*100,1),'p2':round(n_v2/N_seg*100,1)})
    R['transgtipos'] = tipos
    return R

# ══════════════════════════════════════════════════════════════════════════════
# GRÁFICOS
# ══════════════════════════════════════════════════════════════════════════════
def to_rl(fig, w_cm, h_cm):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=180, bbox_inches='tight', facecolor='white', edgecolor='none')
    buf.seek(0); plt.close(fig)
    return Image(buf, width=w_cm*cm, height=h_cm*cm)

def ax_style(ax, horiz=False):
    (ax.xaxis if horiz else ax.yaxis).grid(True, color='#E2E8F0', linewidth=0.6, zorder=0)
    ax.set_axisbelow(True)
    for sp in ['top','right']: ax.spines[sp].set_visible(False)
    ax.spines['left'].set_color('#D0D0D0'); ax.spines['bottom'].set_color('#D0D0D0')
    ax.set_facecolor('white')

def leg_t1t2(ax):
    ax.legend([mpatches.Patch(color=C_T1), mpatches.Patch(color=C_T2)],
              ['Ingreso (TOP 1)','Seguimiento (TOP 2)'],
              fontsize=7.5, frameon=False, loc='upper right')

# G1 – Sexo
def g_sexo(R):
    fig, ax = plt.subplots(figsize=(3.8, 3.0))
    vals = [R['n_hombre'], R['n_mujer']]
    bars = ax.bar(['Hombre','Mujer'], vals, color=[MC_MID, MC_ACCENT], width=0.5, zorder=3)
    for bar, val in zip(bars, vals):
        pct = round(val/R['nv_sex']*100,1) if R['nv_sex']>0 else 0
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.2,
                f'{val}\n({pct}%)', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#333')
    ax.set_ylim(0, max(vals)*1.3 if max(vals)>0 else 1)
    ax.set_ylabel('N personas', fontsize=8, color='#595959')
    ax.tick_params(labelsize=9); ax_style(ax); fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.2, 5.0)

# G2 – Edad
def g_edad(R):
    fig, ax = plt.subplots(figsize=(4.2, 3.0))
    labs = list(R['edad_dist'].keys()); vals = list(R['edad_dist'].values())
    cols = [MC_MID if v==max(vals) else MC_LIGHT for v in vals]
    bars = ax.barh(labs, vals, color=cols, zorder=3)
    for bar, val in zip(bars, vals):
        if val == 0: continue
        pct = round(val/R['nv_edad']*100,1) if R['nv_edad']>0 else 0
        ax.text(bar.get_width()+0.05, bar.get_y()+bar.get_height()/2,
                f'{val} ({pct}%)', va='center', fontsize=8, color='#333')
    ax.set_xlim(0, max(vals)*1.5 if max(vals)>0 else 1)
    ax.tick_params(labelsize=8); ax_style(ax, horiz=True)
    fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.2, 5.0)

# G_torta – Sustancia principal TOP1
def g_torta_sust(R):
    sc = R['sust_comp']
    if not sc:
        fig, ax = plt.subplots(figsize=(5.5,5.5)); ax.text(0.5,0.5,'Sin datos',ha='center')
        return to_rl(fig, 8.5, 8.5)
    labels = [d['label'] for d in sc]; vals = [d['n1'] for d in sc]
    fig = plt.figure(figsize=(5.5, 5.5))
    ax  = fig.add_axes([0.05, 0.18, 0.90, 0.78])
    wedges, _, autotexts = ax.pie(vals, labels=None, colors=PIE_COLS[:len(vals)],
        autopct=lambda p: f'{p:.1f}%' if p>3 else '', startangle=140, pctdistance=0.70,
        wedgeprops={'edgecolor':'white','linewidth':2.0})
    for at in autotexts: at.set_fontsize(10); at.set_color('white'); at.set_fontweight('bold')
    ax.legend(wedges, [f'{l} (n={v})' for l,v in zip(labels,vals)],
              loc='upper center', bbox_to_anchor=(0.5,-0.06), ncol=2, fontsize=8, frameon=False)
    ax.set_aspect('equal'); ax.set_facecolor('white'); fig.patch.set_facecolor('white')
    return to_rl(fig, 8.5, 8.5)

# G3 – Sustancia principal barras agrupadas TOP1 vs TOP2
def g_sust_comp(R):
    datos = R['sust_comp']
    if not datos: return None
    labs = [d['label'] for d in datos]; p1 = [d['p1'] for d in datos]; p2 = [d['p2'] for d in datos]
    x = np.arange(len(labs)); ww = 0.35
    fig, ax = plt.subplots(figsize=(max(5.0, len(labs)*0.85), 3.8))
    b1 = ax.bar(x-ww/2, p1, ww, color=C_T1, label='Ingreso (TOP 1)', zorder=3)
    b2 = ax.bar(x+ww/2, p2, ww, color=C_T2, label='Seguimiento (TOP 2)', zorder=3)
    for bar, v in zip(list(b1)+list(b2), p1+p2):
        if v>0:
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.4,
                    f'{v}%', ha='center', va='bottom', fontsize=7.5, fontweight='bold', color='#333')
    ax.set_xticks(x); ax.set_xticklabels(labs, fontsize=8)
    ax.set_ylabel('% de personas', fontsize=8, color='#595959')
    ax.set_ylim(0, max(p1+p2)*1.35 if p1+p2 else 1)
    ax.legend(fontsize=8, frameon=False); ax_style(ax)
    fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 6.0)

# G4 – Días consumo TOP1 vs TOP2 (barras agrupadas)
def g_dias_comp(R):
    datos = R['dias_comp']
    if not datos: return None
    labs = [d['label'] for d in datos]; m1 = [d['m1'] for d in datos]; m2 = [d['m2'] for d in datos]
    x = np.arange(len(labs)); ww = 0.35
    fig, ax = plt.subplots(figsize=(max(4.8, len(labs)*0.85), 3.2))
    b1 = ax.bar(x-ww/2, m1, ww, color=C_T1, zorder=3)
    b2 = ax.bar(x+ww/2, m2, ww, color=C_T2, zorder=3)
    for bar, v in zip(list(b1)+list(b2), m1+m2):
        if v>0:
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.1,
                    f'{v}d', ha='center', va='bottom', fontsize=8, fontweight='bold', color='#333')
    ax.set_xticks(x); ax.set_xticklabels(labs, fontsize=8.5)
    ax.set_ylabel('Promedio días (0–28)', fontsize=8, color='#595959')
    ax.set_ylim(0, max(m1+m2)*1.32 if m1+m2 else 1)
    leg_t1t2(ax); ax_style(ax)
    fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.5)

# G5 – Cambio en consumo (barras apiladas)
def g_cambio(R):
    datos = R['cambio']
    if not datos: return None
    labs = [d['label'] for d in datos]
    abs_ = [d['pct_abs'] for d in datos]; dis = [d['pct_dis'] for d in datos]
    sc_  = [d['pct_sc']  for d in datos]; emp = [d['pct_emp'] for d in datos]
    x = np.arange(len(labs))
    fig, ax = plt.subplots(figsize=(max(5.0, len(labs)*0.85), 3.5))
    ax.bar(x, abs_, color=C_ABS2, label='Abstinencia', zorder=3)
    ax.bar(x, dis,  bottom=abs_, color=C_DIS2, label='Disminuyó', zorder=3)
    ax.bar(x, sc_,  bottom=[a+d for a,d in zip(abs_,dis)], color=C_SC2, label='Sin cambio', zorder=3)
    ax.bar(x, emp,  bottom=[a+d+s for a,d,s in zip(abs_,dis,sc_)], color=C_EMP2, label='Empeoró', zorder=3)
    for i, (a,d,s,e) in enumerate(zip(abs_,dis,sc_,emp)):
        y_pos = 0
        for val, col in [(a,C_ABS2),(d,C_DIS2),(s,C_SC2),(e,C_EMP2)]:
            if val > 9:
                ax.text(i, y_pos+val/2, f'{val:.0f}%', ha='center', va='center',
                        fontsize=7.5, color='white', fontweight='bold')
            y_pos += val
    ax.set_xticks(x); ax.set_xticklabels(labs, fontsize=9)
    ax.set_ylabel('% de consumidores al ingreso', fontsize=8, color='#595959')
    ax.set_ylim(0, 115)
    ax.legend(loc='upper right', fontsize=7.5, frameon=False, ncol=2)
    ax_style(ax); fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 6.0)

# G6 – % consumidores TOP1 vs TOP2
def g_cons_pct(R):
    datos = R['cons_pct']
    if not datos: return None
    labs = [d['label'] for d in datos]; p1 = [d['p1'] for d in datos]; p2 = [d['p2'] for d in datos]
    x = np.arange(len(labs)); ww = 0.35
    fig, ax = plt.subplots(figsize=(max(4.8, len(labs)*0.85), 3.2))
    b1 = ax.bar(x-ww/2, p1, ww, color=C_T1, zorder=3)
    b2 = ax.bar(x+ww/2, p2, ww, color=C_T2, zorder=3)
    for bar, v in zip(list(b1)+list(b2), p1+p2):
        if v>0:
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.3,
                    f'{v}%', ha='center', va='bottom', fontsize=7.5, fontweight='bold', color='#333')
    ax.set_xticks(x); ax.set_xticklabels(labs, fontsize=8.5)
    ax.set_ylabel('% de personas', fontsize=8, color='#595959')
    ax.set_ylim(0, max(p1+p2)*1.35 if p1+p2 else 1)
    leg_t1t2(ax); ax_style(ax)
    fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.5)

# G7 – Días por sustancia TOP1 vs TOP2
def g_dias_sust(R):
    return g_dias_comp(R)   # misma estructura, mismos datos

# G8 – Transgresión general: barras verticales
def g_transgresion(R):
    N = R['N_seg']
    cats = ['Ingreso (TOP 1)','Seguimiento (TOP 2)']
    pcts = [R['pct_tr1'], R['pct_tr2']]; ns = [R['n_tr1'], R['n_tr2']]
    fig, ax = plt.subplots(figsize=(3.8, 3.2))
    bars = ax.bar(cats, pcts, color=[C_T1, C_T2], width=0.5, zorder=3)
    for bar, pct, n in zip(bars, pcts, ns):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.8,
                f'{pct}%\n(n={n})', ha='center', va='bottom', fontsize=10, fontweight='bold', color='#333')
    ax.set_ylim(0, max(pcts)*1.4 if pcts else 1)
    ax.set_ylabel(f'% sobre N={N}', fontsize=8, color='#595959')
    ax.tick_params(labelsize=9.5); ax_style(ax)
    fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.5)

# G9 – Tipos de transgresión TOP1 vs TOP2
def g_tipos_tr(R):
    datos = R['transgtipos']
    if not datos: return None
    labs = [d['label'] for d in datos]; p1 = [d['p1'] for d in datos]; p2 = [d['p2'] for d in datos]
    x = np.arange(len(labs)); ww = 0.35
    fig, ax = plt.subplots(figsize=(max(5.0, len(labs)*0.9), 3.2))
    b1 = ax.bar(x-ww/2, p1, ww, color=C_T1, zorder=3)
    b2 = ax.bar(x+ww/2, p2, ww, color=C_T2, zorder=3)
    for bar, v in zip(list(b1)+list(b2), p1+p2):
        if v>0:
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.3,
                    f'{v}%', ha='center', va='bottom', fontsize=8, fontweight='bold', color='#333')
    ax.set_xticks(x); ax.set_xticklabels(labs, fontsize=8.5)
    ax.set_ylim(0, max(p1+p2)*1.38 if p1+p2 else 1)
    ax.set_ylabel(f'% sobre N={R["N_seg"]}', fontsize=8, color='#595959')
    leg_t1t2(ax); ax_style(ax)
    fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.5)

# G10 – Salud TOP1 vs TOP2
def g_salud(R):
    datos = R['salud']
    if not datos: return None
    labs = [d['label'] for d in datos]; m1 = [d['m1'] for d in datos]; m2 = [d['m2'] for d in datos]
    y = np.arange(len(labs)); ww = 0.35
    fig, ax = plt.subplots(figsize=(5.2, 3.0))
    fig.subplots_adjust(top=0.82)
    b1 = ax.barh(y-ww/2, m1, ww, color=C_T1, zorder=3)
    b2 = ax.barh(y+ww/2, m2, ww, color=C_T2, zorder=3)
    for bar, v in zip(list(b1)+list(b2), m1+m2):
        ax.text(bar.get_width()+0.2, bar.get_y()+bar.get_height()/2,
                f'{v}', va='center', fontsize=9, fontweight='bold', color='#333')
    ax.set_yticks(y); ax.set_yticklabels(labs, fontsize=9)
    ax.set_xlim(0, 26); ax.axvline(x=10, color='#BFBFBF', linestyle='--', linewidth=0.8)
    ax.set_xlabel('Promedio (0–20)', fontsize=8, color='#595959')
    ax_style(ax, horiz=True)
    fig.legend([mpatches.Patch(color=C_T1), mpatches.Patch(color=C_T2)],
               ['Ingreso (TOP 1)','Seguimiento (TOP 2)'],
               fontsize=8, frameon=False, loc='upper center',
               bbox_to_anchor=(0.5, 0.98), ncol=2)
    fig.patch.set_facecolor('white')
    return to_rl(fig, 8.5, 5.5)

# G11 – Vivienda TOP1 vs TOP2
def g_vivienda(R):
    cats = ['Lugar\nestable','Condiciones\nbásicas']
    p1 = [R['viv1_t1'][1], R['viv2_t1'][1]]; p2 = [R['viv1_t2'][1], R['viv2_t2'][1]]
    x = np.arange(len(cats)); ww = 0.35
    fig, ax = plt.subplots(figsize=(4.2, 3.0))
    fig.subplots_adjust(top=0.82)
    b1 = ax.bar(x-ww/2, p1, ww, color=C_T1, zorder=3)
    b2 = ax.bar(x+ww/2, p2, ww, color=C_T2, zorder=3)
    for bar, v in zip(list(b1)+list(b2), p1+p2):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+1.0,
                f'{v}%', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#333')
    ax.set_xticks(x); ax.set_xticklabels(cats, fontsize=10)
    ax.set_ylim(0, 118); ax.set_ylabel('% personas con condición', fontsize=8, color='#595959')
    ax_style(ax)
    fig.legend([mpatches.Patch(color=C_T1), mpatches.Patch(color=C_T2)],
               ['Ingreso (TOP 1)','Seguimiento (TOP 2)'],
               fontsize=8, frameon=False, loc='upper center',
               bbox_to_anchor=(0.5, 0.98), ncol=2)
    fig.patch.set_facecolor('white')
    return to_rl(fig, 8.5, 5.5)

# ══════════════════════════════════════════════════════════════════════════════
# ESTILOS Y HELPERS PDF
# ══════════════════════════════════════════════════════════════════════════════
TW = W - 3.4*cm

def make_styles():
    ST = {}
    ST['sec']  = ParagraphStyle('sec',  fontName='Helvetica-Bold',   fontSize=11, textColor=white,
                   backColor=C_MID, leading=16, spaceBefore=6, spaceAfter=4, leftIndent=6, borderPad=5)
    ST['sub']  = ParagraphStyle('sub',  fontName='Helvetica-Bold',   fontSize=10, textColor=C_DARK,
                   leading=13, spaceBefore=4, spaceAfter=3)
    ST['body'] = ParagraphStyle('body', fontName='Helvetica',        fontSize=9.5,
                   textColor=HexColor('#2D2D2D'), leading=14, spaceBefore=2, spaceAfter=3, alignment=TA_JUSTIFY)
    ST['nota'] = ParagraphStyle('nota', fontName='Helvetica-Oblique',fontSize=8,
                   textColor=C_GRAY, leading=11, spaceBefore=1, spaceAfter=3)
    ST['ptit'] = ParagraphStyle('ptit', fontName='Helvetica-Bold',   fontSize=24, textColor=white,
                   leading=30, alignment=TA_CENTER)
    ST['psub'] = ParagraphStyle('psub', fontName='Helvetica',        fontSize=13, textColor=C_LIGHT,
                   leading=18, alignment=TA_CENTER)
    ST['pserv']= ParagraphStyle('pserv',fontName='Helvetica-Bold',   fontSize=19, textColor=white,
                   leading=24, alignment=TA_CENTER)
    ST['kval'] = ParagraphStyle('kval', fontName='Helvetica-Bold',   fontSize=24, textColor=C_MID,
                   leading=30, alignment=TA_CENTER)
    ST['klab'] = ParagraphStyle('klab', fontName='Helvetica',        fontSize=9,  textColor=C_GRAY,
                   leading=12, alignment=TA_CENTER)
    return ST

def bloque(titulo, img, parrafos, ST, cw=0.53):
    CW = TW*cw; TXW = TW*(1-cw)
    row = Table([[img, parrafos]], colWidths=[CW, TXW])
    row.setStyle(TableStyle([
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('LEFTPADDING',(0,0),(-1,-1),0), ('RIGHTPADDING',(0,0),(-1,-1),5),
        ('TOPPADDING',(0,0),(-1,-1),2),  ('BOTTOMPADDING',(0,0),(-1,-1),2),
    ]))
    return [Paragraph(titulo, ST['sub']), row]

def sec_t(n, title, ST):
    txt = f'  {n}. {title.upper()}' if n else f'  {title.upper()}'
    return Paragraph(txt, ST['sec'])

def flecha(v1, v2, mejor_si_sube=True):
    if v1 == v2: return '→ Sin cambio'
    mejoro = (v2 > v1) == mejor_si_sube
    return f'↑ Mejoró ({v1}→{v2})' if mejoro else f'↓ Empeoró ({v1}→{v2})'

def hr(): return HRFlowable(width='100%', thickness=0.5, color=C_LIGHT, spaceAfter=4, spaceBefore=4)
def sp(h=0.2): return Spacer(1, h*cm)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTRUCCIÓN DEL PDF — 6 páginas
# ══════════════════════════════════════════════════════════════════════════════
def build_pdf(R):
    ST  = make_styles()
    doc = SimpleDocTemplate(OUTPUT_FILE, pagesize=A4,
          leftMargin=1.7*cm, rightMargin=1.7*cm, topMargin=1.4*cm, bottomMargin=1.5*cm)
    S = []
    N = R['N_seg']; pct_seg = round(N/R['N_total']*100,1)

    # ── PÁG 1: PORTADA ────────────────────────────────────────────────────────
    cover = Table([
        [Paragraph('INFORME DE SEGUIMIENTO', ST['ptit'])],
        [sp(0.5)],
        [Paragraph('Monitoreo de Resultados de Tratamiento<br/>Instrumento TOP', ST['psub'])],
        [sp(2.0)],
        [Paragraph(NOMBRE_SERVICIO.upper(), ST['pserv'])],
        [sp(0.4)],
        [Paragraph(PERIODO, ParagraphStyle('pp', fontName='Helvetica', fontSize=12,
            textColor=C_LIGHT, alignment=TA_CENTER))],
    ], colWidths=[TW])
    cover.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1), C_DARK), ('BACKGROUND',(0,4),(-1,6), C_MID),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),      ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(0,0),110),         ('BOTTOMPADDING',(0,-1),(-1,-1),110),
        ('LEFTPADDING',(0,0),(-1,-1),20),       ('RIGHTPADDING',(0,0),(-1,-1),20),
    ]))
    S += [cover, PageBreak()]

    # ── PÁG 2: PRESENTACIÓN + KPIs + G1 SEXO ─────────────────────────────────
    S.append(sec_t('','Presentación',ST)); S.append(sp(0.2))

    # Texto de tiempo de seguimiento
    _st = R.get('seg_tiempo', {})
    if _st.get('mediana') is not None:
        _txt_tiempo = (f'El tiempo transcurrido entre la evaluación de ingreso (TOP 1) y el seguimiento (TOP 2) '
                       f'fue en mediana de <b>{_st["mediana"]} meses</b> '
                       f'(rango: {_st["min"]}–{_st["max"]} meses; N válido: {_st["n"]}'
                       + (f' de {_st["n_total"]}' if _st["n_total"] != _st["n"] else '') + ').')
    else:
        _txt_tiempo = None

    S.append(Paragraph(
        f'El presente informe describe la evolución de las personas en tratamiento en '\
        f'<b>{NOMBRE_SERVICIO}</b> durante el período <b>{PERIODO}</b>, '\
        f'comparando resultados al <b>ingreso (TOP 1)</b> y al <b>seguimiento (TOP 2)</b>. '\
        f'De los <b>{R["N_total"]} pacientes</b> que ingresaron, '\
        f'<b>{N} ({pct_seg}%)</b> cuentan con ambas evaluaciones. '\
        f'La sustancia de mayor problema al ingreso fue el '\
        f'<b>{R["sust_top1"]} ({R["sust_top1_pct"]}%)</b>.',
        ST['body']))
    if _txt_tiempo:
        S.append(sp(0.1))
        S.append(Paragraph(_txt_tiempo, ST['body']))
    S.append(sp(0.2))

    _st = R.get('seg_tiempo', {})
    _tiempo_val   = f'{_st["mediana"]} m' if _st.get('mediana') is not None else '—'
    _tiempo_rango = f'{_st["min"]}–{_st["max"]} m' if _st.get('min') is not None else ''

    kpi = Table([[
        Paragraph(str(N),                ST['kval']),
        Paragraph(f'{R["pct_hombre"]}%', ST['kval']),
        Paragraph(str(R['edad_media']),  ST['kval']),
        Paragraph(_tiempo_val,           ST['kval']),
    ],[
        Paragraph('Personas con<br/>seguimiento', ST['klab']),
        Paragraph('Son<br/>hombres',              ST['klab']),
        Paragraph('Edad<br/>promedio',             ST['klab']),
        Paragraph('Mediana<br/>seguimiento' + (f'<br/><font size="7">{_tiempo_rango}</font>' if _tiempo_rango else ''), ST['klab']),
    ]], colWidths=[TW/4]*4)
    kpi.setStyle(TableStyle([
        ('ALIGN',(0,0),(-1,-1),'CENTER'),   ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('BACKGROUND',(0,0),(-1,-1), HexColor('#EEF4FB')),
        ('BOX',(0,0),(-1,-1),0.5,C_LIGHT), ('INNERGRID',(0,0),(-1,-1),0.5,C_LIGHT),
        ('TOPPADDING',(0,0),(-1,-1),8),     ('BOTTOMPADDING',(0,0),(-1,-1),8),
    ]))
    S += [kpi, sp(0.25), hr()]

    S.append(sec_t('1','Antecedentes Generales',ST)); S.append(sp(0.1))
    S += bloque('1.1. Distribución de Personas según Sexo', g_sexo(R), [
        Paragraph(f'De las <b>{N} personas</b> con seguimiento completo, '
                  f'<b>{R["n_hombre"]} ({R["pct_hombre"]}%) son hombres</b> y '
                  f'<b>{R["n_mujer"]} ({R["pct_mujer"]}%) son mujeres</b>.', ST['body']),
        sp(0.15),
        Paragraph(f'N válido: {R["nv_sex"]} personas.', ST['nota']),
    ], ST)
    S.append(PageBreak())

    # ── PÁG 3: G2 EDAD + G_torta + G3 SUST COMP ──────────────────────────────
    rm = max(R['edad_dist'], key=R['edad_dist'].get) if R['edad_dist'] else 'Sin datos'
    nm = R['edad_dist'].get(rm, 0)
    pm = round(nm/R['nv_edad']*100,1) if R['nv_edad']>0 else 0
    S += bloque('1.2. Distribución de Personas según Edad', g_edad(R), [
        Paragraph(f'El <b>promedio de edad</b> es de <b>{R["edad_media"]} años</b> '
                  f'(DE={R["edad_sd"]}; rango {R["edad_min"]}–{R["edad_max"]} años). '
                  f'El rango más frecuente es <b>{rm}</b> ({nm} personas, {pm}%).', ST['body']),
        sp(0.15),
        Paragraph(f'N válido: {R["nv_edad"]} personas.', ST['nota']),
    ], ST)
    S += [sp(0.25), hr()]

    S.append(sec_t('2','Consumo de Sustancias',ST)); S.append(sp(0.1))
    sc = R['sust_comp']
    top1_d = max(sc, key=lambda x: x['n1']) if sc else {'label':'—','p1':0}
    seg2b  = [d for d in sc if d['n1'] < top1_d['n1']]
    seg2b_top = max(seg2b, key=lambda x: x['n1']) if seg2b else None
    S += bloque('2.1. Consumo de Sustancia Principal al Ingreso', g_torta_sust(R), [
        Paragraph('Distribución de las personas según la '
                  '<b>sustancia que consideran que les genera más problemas</b> '
                  'al ingresar a tratamiento.', ST['body']),
        sp(0.15),
        Paragraph(f'La sustancia más frecuente es el <b>{top1_d["label"]} ({top1_d["p1"]}%)</b>'
                  + (f', seguida por <b>{seg2b_top["label"]} ({seg2b_top["p1"]}%)</b>' if seg2b_top else '')
                  + '.', ST['body']),
        sp(0.15),
        Paragraph(f'N válido: TOP1={R["nv_sust1"]}.', ST['nota']),
    ], ST)
    S += [sp(0.2), hr()]

    img_g3 = g_sust_comp(R)
    if img_g3:
        top2_d = max(sc, key=lambda x: x['n2']) if sc else {'label':'—','p2':0}
        S += bloque('2.2. Consumo Sustancia Principal — Ingreso vs Seguimiento', img_g3, [
            Paragraph('Comparación de la distribución de la sustancia principal '
                      'al <b>ingreso (TOP 1)</b> y al <b>seguimiento (TOP 2)</b>.', ST['body']),
            sp(0.15),
            Paragraph(f'Al ingreso: <b>{R["sust_top1"]} ({R["sust_top1_pct"]}%)</b>. '
                      f'Al seguimiento: <b>{top2_d["label"]} ({top2_d["p2"]}%)</b>.', ST['body']),
            sp(0.15),
            Paragraph(f'N válido: TOP1={R["nv_sust1"]}, TOP2={R["nv_sust2"]}.', ST['nota']),
        ], ST)
    S.append(PageBreak())

    # ── PÁG 4: G4 DÍAS CONSUMO + G5 CAMBIO ───────────────────────────────────
    img_g4 = g_dias_comp(R)
    if img_g4:
        S += bloque('2.3. Promedio de Días de Consumo — Ingreso vs Seguimiento', img_g4, [
            Paragraph('Promedio de días de consumo en las <b>últimas 4 semanas</b> '
                      'al ingreso y al seguimiento (sobre N total, incluyendo 0).', ST['body']),
            sp(0.15),
            Paragraph('Se espera una reducción en los promedios entre TOP 1 y TOP 2.', ST['body']),
            sp(0.15),
            Paragraph(f'N: {N} pacientes con ambas evaluaciones.', ST['nota']),
        ], ST)
        S += [sp(0.2), hr()]

    img_g5 = g_cambio(R)
    if img_g5:
        c = R['cambio']
        pct_abst = round(sum(d['pct_abs'] for d in c)/len(c),1) if c else 0
        S += bloque('2.4. Cambio en el Consumo por Sustancia', img_g5, [
            Paragraph('Evolución del consumo entre TOP 1 y TOP 2, clasificando a cada persona '
                      'según si logró <b>abstinencia</b>, <b>disminuyó</b>, '
                      'se mantuvo igual o empeoró.', ST['body']),
            sp(0.15),
            Paragraph(f'En promedio, el <b>{pct_abst}%</b> de los consumidores de cada '
                      f'sustancia logró abstinencia al seguimiento.', ST['body']),
            sp(0.15),
            Paragraph('% calculado sobre consumidores al ingreso (días > 0 en TOP 1).', ST['nota']),
        ], ST)
    S.append(PageBreak())

    # ── PÁG 5: G6 % CONSUMIDORES + G7 DÍAS POR SUST + G8 TRANSGRESIÓN ────────
    img_g6 = g_cons_pct(R)
    if img_g6:
        S += bloque('2.5. Consumo de Sustancias — % de Personas', img_g6, [
            Paragraph('Porcentaje de personas que consume cada sustancia al ingreso y al '
                      'seguimiento. Los % pueden sumar más de 100% (una persona puede consumir varias).', ST['body']),
            sp(0.15),
            Paragraph(f'N total: {N} pacientes.', ST['nota']),
        ], ST)
        S += [sp(0.2), hr()]

    img_g7 = g_dias_sust(R)
    if img_g7:
        S += bloque('2.6. Promedio de Días de Consumo por Sustancia', img_g7, [
            Paragraph(f'Días promedio de consumo por sustancia (sobre N={N}), '
                      f'incluyendo a quienes no consumen (días=0).', ST['body']),
            sp(0.15),
            Paragraph('La comparación refleja la reducción general del consumo '
                      'entre ingreso y seguimiento.', ST['body']),
        ], ST)
        S += [sp(0.2), hr()]

    S.append(sec_t('3','Transgresión a la Norma Social',ST)); S.append(sp(0.1))
    reduc = round(R['pct_tr1'] - R['pct_tr2'], 1)
    S += bloque('3.1. Transgresión a la Norma Social — Ingreso vs Seguimiento', g_transgresion(R), [
        Paragraph(f'Al ingreso, <b>{R["n_tr1"]} personas ({R["pct_tr1"]}%)</b> '
                  f'cometieron alguna transgresión. '
                  f'Al seguimiento ese número bajó a '
                  f'<b>{R["n_tr2"]} personas ({R["pct_tr2"]}%)</b>.', ST['body']),
        sp(0.15),
        Paragraph(f'Reducción de <b>{reduc} puntos porcentuales</b> entre TOP 1 y TOP 2.', ST['body']),
        sp(0.15),
        Paragraph(f'N total: {N} pacientes.', ST['nota']),
    ], ST)
    S.append(PageBreak())

    # ── PÁG 6: G9 TIPOS TRANSGRESIÓN + G10 SALUD + G11 VIVIENDA ─────────────
    img_g9 = g_tipos_tr(R)
    if img_g9:
        S += bloque('3.2. Distribución por Tipo de Transgresión', img_g9, [
            Paragraph('Reducción en cada tipo de transgresión entre ingreso y seguimiento.', ST['body']),
            sp(0.15),
            Paragraph(f'Los % no suman 100% (una persona puede cometer más de un tipo). '
                      f'N base: N={N} pacientes.', ST['nota']),
        ], ST)
        S += [sp(0.2), hr()]

    S.append(sec_t('4','Salud, Calidad de Vida y Vivienda',ST)); S.append(sp(0.1))
    if R['salud']:
        mejor = max(R['salud'], key=lambda s: s['m2'])
        img_g10 = g_salud(R)
        if img_g10:
            S += bloque('4.1. Autopercepción del Estado de Salud (0–20)', img_g10, [
                Paragraph('Las dimensiones de salud al ingreso y al seguimiento. '
                          f'La mayor mejora se observa en <b>{mejor["label"]}</b> '
                          f'({flecha(mejor["m1"],mejor["m2"],True)}).', ST['body']),
                sp(0.15),
                Paragraph(f'Escala 0 (muy mal) a 20 (excelente). '
                          f'Línea punteada = punto medio (10). '
                          f'N válido: {R["salud"][0]["nv1"]} TOP1 / {R["salud"][0]["nv2"]} TOP2.', ST['nota']),
            ], ST)
            S.append(sp(0.15))

    n1_1,p1_1,_ = R['viv1_t1']; n1_2,p1_2,_ = R['viv1_t2']
    n2_1,p2_1,_ = R['viv2_t1']; n2_2,p2_2,_ = R['viv2_t2']
    S += bloque('4.2. Condiciones de Vivienda — Ingreso y Seguimiento', g_vivienda(R), [
        Paragraph(f'<b>Lugar estable:</b> {p1_1}% al ingreso → {p1_2}% al seguimiento '
                  f'({flecha(p1_1,p1_2,True)}).', ST['body']),
        sp(0.1),
        Paragraph(f'<b>Condiciones básicas:</b> {p2_1}% al ingreso → {p2_2}% al seguimiento '
                  f'({flecha(p2_1,p2_2,True)}).', ST['body']),
        sp(0.15),
        Paragraph(f'N válido: {N} pacientes.', ST['nota']),
    ], ST)

    S += [sp(0.3), hr(),
          Paragraph(f'Informe generado automáticamente · TOP · {NOMBRE_SERVICIO} · {PERIODO}',
                    ParagraphStyle('pie', fontName='Helvetica-Oblique', fontSize=7.5,
                                   textColor=C_GRAY, alignment=TA_CENTER))]
    doc.build(S)
    print(f'  ✓ PDF generado: {OUTPUT_FILE}')

# ══════════════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    print('=' * 60)
    print('  SCRIPT_TOP_Universal_PDF_Seguimiento  —  Iniciando...')
    print('=' * 60)
    print('\n→ Calculando datos...')
    R = cargar_datos()
    print(f'  N_total={R["N_total"]} | N_seg={R["N_seg"]} | {R["sust_top1"]} {R["sust_top1_pct"]}%')
    print('\n→ Generando PDF...')
    build_pdf(R)
    print(f'\n{"=" * 60}')
    print(f'  ✅  LISTO  →  {OUTPUT_FILE}')
    print(f'{"=" * 60}')
