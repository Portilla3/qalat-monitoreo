"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   SCRIPT_TOP_Universal_PDF_Caracterizacion.py                              ║
║   Genera informe PDF de caracterización al ingreso (TOP1)                  ║
║   10 gráficos · 5 páginas · Compatible con cualquier país TOP              ║
║   Versión Universal 1.0                                                    ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                             ║
║  CÓMO USAR LA PRÓXIMA VEZ:                                                 ║
║  1. Abre un chat nuevo con Claude                                           ║
║  2. Sube DOS archivos:                                                      ║
║       • Este script: SCRIPT_TOP_Universal_PDF_Caracterizacion.py           ║
║       • La base en formato Wide (generada por SCRIPT_TOP_Universal_Wide)   ║
║  3. Escribe exactamente:                                                    ║
║     "Ejecuta el script universal PDF Caracterización con esta base Wide"   ║
║  4. Claude ajustará NOMBRE_SERVICIO y PERIODO según corresponda            ║
║                                                                             ║
║  ESTRUCTURA DEL PDF (5 páginas):                                           ║
║    Pág 1 – Portada                                                         ║
║    Pág 2 – Presentación + KPIs + G1 Sexo                                  ║
║    Pág 3 – G2 Edad + G3 Torta sustancia principal                         ║
║    Pág 4 – G4 Días sust.principal + G5 Consumo % + G6 Días por sustancia  ║
║    Pág 5 – G7 Donut transgresión + G8 Tipos + G9 Salud + G10 Vivienda     ║
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
OUTPUT_FILE     = '/home/claude/TOP_Informe_Caracterizacion.pdf'
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
MC_MID = '#2E75B6'; MC_LIGHT = '#BDD7EE'; MC_GRAY = '#BFBFBF'; MC_ACCENT = '#00B0F0'
PIE_COLS = ['#2E75B6','#1F3864','#4472C4','#9DC3E6','#00B0F0','#538135','#BFBFBF','#C00000','#ED7D31']

# ══════════════════════════════════════════════════════════════════════════════
# DETECCIÓN DINÁMICA DE COLUMNAS (_TOP1)
# ══════════════════════════════════════════════════════════════════════════════
def detectar_columnas(cols):
    col_set = set(cols)

    # Sustancias: "1) Registrar... >> Nombre (unidad) >> Total (0-28)_TOP1"
    sust_cols = []
    for c in cols:
        if c.endswith('_TOP1') and 'Total (0-28)' in c:
            base = c.replace('_TOP1', '')
            if base.startswith('1)'):
                partes = base.split('>>')
                if len(partes) >= 3:
                    nombre = partes[-2].strip().split('(')[0].strip()
                    sust_cols.append((nombre, c))
    print(f'  Sustancias: {[s[0] for s in sust_cols]}')

    # Transgresión Sí/No
    tr_sn = []
    for c in cols:
        if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('3)') and '>>' in c:
            nombre = c.replace('_TOP1','').split('>>')[-1].strip()
            tr_sn.append((nombre, c))
    print(f'  Transgresión: {[t[0] for t in tr_sn]}')

    vif     = next((c for c in cols if c.endswith('_TOP1') and '4)' in c
                    and 'Violencia Intrafamiliar' in c and 'Total (0-28)' in c), None)
    sal_psi = next((c for c in cols if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('6)')), None)
    sal_fis = next((c for c in cols if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('8)')), None)
    cal_vid = next((c for c in cols if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('10)')), None)
    viv1    = next((c for c in cols if c.endswith('_TOP1') and '9)' in c and 'estable' in c.lower()), None)
    viv2    = next((c for c in cols if c.endswith('_TOP1') and '9)' in c and 'básicas' in c.lower()), None)
    sust_pp = next((c for c in cols if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('2)')
                    and 'sustancia principal' in c.lower()), None)
    sexo    = next((c for c in cols if c.endswith('_TOP1') and 'sexo' in c.lower()), None)
    fn_col  = next((c for c in cols if c.endswith('_TOP1') and 'nacimiento' in c.lower()), None)
    fecha   = next((c for c in cols if c.endswith('_TOP1') and 'fecha entrevista' in c.lower()), None)

    return dict(sust_cols=sust_cols, tr_sn=tr_sn, vif=vif,
                sal_psi=sal_psi, sal_fis=sal_fis, cal_vid=cal_vid,
                viv1=viv1, viv2=viv2, sust_pp=sust_pp,
                sexo=sexo, fn_col=fn_col, fecha=fecha)

# ══════════════════════════════════════════════════════════════════════════════
# NORMALIZACIÓN SUSTANCIA PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def norm_sust(s):
    if pd.isna(s) or str(s).strip() == '0': return None
    s = str(s).strip().lower()
    if any(x in s for x in ['alcohol','cerveza','licor','aguard','alchol','bebida']): return 'Alcohol'
    if any(x in s for x in ['marihu','marjhu','marhuana','cannabis','cannbis','cannabin']): return 'Cannabis/\nMarihuana'
    if any(x in s for x in ['pasta base','pasta','papelillo']): return 'Pasta\nBase'
    if any(x in s for x in ['crack','cristal','piedra','paco']): return 'Crack/\nCristal'
    if any(x in s for x in ['cocain','perico','coca ']): return 'Cocaína'
    if any(x in s for x in ['tabaco','cigarr','nicot']): return 'Tabaco/\nNicotina'
    if any(x in s for x in ['inhalant','thiner','activo','resistol']): return 'Inhalantes'
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
        OUTPUT_FILE = f'/home/claude/TOP_Informe_Caracterizacion_{FILTRO_CENTRO}.pdf'
        _pais_local = _detectar_pais(INPUT_FILE)
        NOMBRE_SERVICIO = (f'{_pais_local}  —  Centro {FILTRO_CENTRO}'
                           if _pais_local else f'Centro {FILTRO_CENTRO}')
    N  = len(df)
    DC = detectar_columnas(df.columns.tolist())
    R  = {'N': N, 'd': df, 'DC': DC}

    # Sexo
    if DC['sexo']:
        sc = df[DC['sexo']].astype(str).str.strip().str.upper()
        nv = int(sc.isin(['H','M']).sum())
        R['n_hombre']   = int((sc=='H').sum())
        R['n_mujer']    = int((sc=='M').sum())
        R['nv_sex']     = nv
        R['pct_hombre'] = round(R['n_hombre']/nv*100,1) if nv>0 else 0
        R['pct_mujer']  = round(R['n_mujer'] /nv*100,1) if nv>0 else 0
    else:
        R['n_hombre']=R['n_mujer']=R['nv_sex']=0
        R['pct_hombre']=R['pct_mujer']=0

    # Edad
    if DC['fn_col'] and DC['fecha']:
        fn  = pd.to_datetime(df[DC['fn_col']], errors='coerce')
        ref = pd.to_datetime(df[DC['fecha']], errors='coerce').fillna(pd.Timestamp.now())
        edad = ((ref-fn).dt.days/365.25).round(1)
        edad = edad[(edad>=10)&(edad<=100)]
        R['nv_edad']    = int(edad.notna().sum())
        R['edad_media'] = round(float(edad.mean()),1) if R['nv_edad']>0 else 0
        R['edad_sd']    = round(float(edad.std()),1)  if R['nv_edad']>0 else 0
        R['edad_min']   = int(edad.min()) if R['nv_edad']>0 else 0
        R['edad_max']   = int(edad.max()) if R['nv_edad']>0 else 0
        bins = [0,17,30,40,50,60,200]
        labs = ['Menos de 18','18 a 30','31 a 40','41 a 50','51 a 60','61 o más']
        ec   = pd.cut(edad, bins=bins, labels=labs)
        R['edad_dist'] = {l: int((ec==l).sum()) for l in labs}
    else:
        R['edad_media']=R['edad_sd']=0; R['edad_min']=R['edad_max']=0
        R['nv_edad']=0; R['edad_dist']={'Sin datos':0}

    # Sustancia principal
    if DC['sust_pp']:
        sr = df[DC['sust_pp']].apply(norm_sust).dropna()
        R['nv_sust']       = len(sr)
        vc                 = sr.value_counts()
        R['sust_vc']       = vc
        R['sust_top1']     = vc.index[0].replace('\n',' ') if len(vc)>0 else '—'
        R['sust_top1_pct'] = round(vc.iloc[0]/len(sr)*100,1) if len(sr)>0 else 0
    else:
        R['nv_sust']=0; R['sust_vc']=pd.Series(dtype=int)
        R['sust_top1']='—'; R['sust_top1_pct']=0

    # Días consumo – sustancia principal (G4)
    sust_norm = df[DC['sust_pp']].apply(norm_sust) if DC['sust_pp'] else pd.Series([None]*N)
    dias_princ = {}
    for lbl, col in DC['sust_cols']:
        v = pd.to_numeric(df[col], errors='coerce')
        mask = sust_norm.apply(lambda s: isinstance(s,str) and lbl.lower() in s.lower().replace('\n',' '))
        sub  = v[mask & (v>0)]
        if len(sub) >= 1:
            dias_princ[lbl] = {'prom': round(float(sub.mean()),1), 'n': int(len(sub))}
    R['dias_princ'] = dias_princ

    # % consumidores por sustancia (G5)
    consumo_pct = {}
    for lbl, col in DC['sust_cols']:
        v   = pd.to_numeric(df[col], errors='coerce'); n_c = int((v>0).sum())
        if n_c > 0:
            consumo_pct[lbl] = {'pct': round(n_c/N*100,1), 'n': n_c}
    R['consumo_pct'] = consumo_pct

    # Promedio días por sustancia – todos los consumidores (G6)
    dias_sust = {}
    for lbl, col in DC['sust_cols']:
        v   = pd.to_numeric(df[col], errors='coerce'); sub = v[v>0]
        if len(sub) >= 1:
            dias_sust[lbl] = {'prom': round(float(sub.mean()),1), 'n': int(len(sub))}
    R['dias_sust'] = dias_sust

    # Salud (G9)
    salud = []
    for lbl, col in [('Salud Psicológica', DC['sal_psi']),
                     ('Salud Física',      DC['sal_fis']),
                     ('Calidad de Vida',   DC['cal_vid'])]:
        if col:
            v = pd.to_numeric(df[col], errors='coerce')
            salud.append({'label': lbl, 'prom': round(float(v.mean()),1),
                          'nv': int(v.notna().sum())})
    R['salud'] = salud

    # Vivienda (G10)
    def viv(col):
        if not col: return (0, 0, 0)
        nv_ = int(df[col].isin(['Sí','No']).sum()) or N
        n_  = int((df[col]=='Sí').sum())
        return n_, round(n_/nv_*100,1), nv_
    R['viv1'] = viv(DC['viv1'])
    R['viv2'] = viv(DC['viv2'])

    # Transgresión (G7, G8)
    tr_cols = [c for _,c in DC['tr_sn']]
    def has_tr(row):
        for c in tr_cols:
            if _es_positivo(row.get(c,'')): return True
        if DC['vif']:
            v = pd.to_numeric(row.get(DC['vif'], np.nan), errors='coerce')
            return not np.isnan(v) and v>0
        return False
    t = df.apply(lambda r: int(has_tr(r)), axis=1)
    R['n_transgresores']  = int(t.sum())
    R['pct_transgresores'] = round(R['n_transgresores']/N*100,1)
    tipos = []
    for lbl, col in DC['tr_sn']:
        n = int(df[col].apply(_es_positivo).sum())
        tipos.append({'label': lbl, 'n': n,
                      'pct': round(n/R['n_transgresores']*100,1) if R['n_transgresores']>0 else 0})
    if DC['vif']:
        vif_v = pd.to_numeric(df[DC['vif']], errors='coerce'); n_vif = int((vif_v>0).sum())
        tipos.append({'label': 'VIF', 'n': n_vif,
                      'pct': round(n_vif/R['n_transgresores']*100,1) if R['n_transgresores']>0 else 0})
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

def g_sexo(R):
    fig, ax = plt.subplots(figsize=(4.5, 3.8))
    vals = [R['n_hombre'], R['n_mujer']]
    bars = ax.bar(['Hombre','Mujer'], vals, color=[MC_MID, MC_ACCENT], width=0.5, zorder=3)
    for bar, val in zip(bars, vals):
        pct = round(val/R['nv_sex']*100,1) if R['nv_sex']>0 else 0
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.5,
                f'{val}\n({pct}%)', ha='center', va='bottom', fontsize=10, fontweight='bold', color='#333')
    ax.set_ylim(0, max(vals)*1.28 if max(vals)>0 else 1)
    ax.set_ylabel('N personas', fontsize=9, color='#595959')
    ax.tick_params(labelsize=10); ax_style(ax); fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.8)

def g_edad(R):
    fig, ax = plt.subplots(figsize=(5.0, 3.8))
    labs = list(R['edad_dist'].keys()); vals = list(R['edad_dist'].values())
    cols = [MC_MID if v==max(vals) else MC_LIGHT for v in vals]
    bars = ax.barh(labs, vals, color=cols, zorder=3)
    for bar, val in zip(bars, vals):
        pct = round(val/R['nv_edad']*100,1) if R['nv_edad']>0 else 0
        ax.text(bar.get_width()+0.2, bar.get_y()+bar.get_height()/2,
                f'{val} ({pct}%)', va='center', fontsize=9, color='#333')
    ax.set_xlim(0, max(vals)*1.45 if max(vals)>0 else 1)
    ax.tick_params(labelsize=9); ax_style(ax, horiz=True)
    fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.8)

def g_torta_sust(R):
    vc = R['sust_vc']
    if len(vc) == 0:
        fig, ax = plt.subplots(figsize=(5.5,5.5)); ax.text(0.5,0.5,'Sin datos',ha='center')
        return to_rl(fig, 9.0, 8.5)
    labels = [l.replace('\n',' ') for l in vc.index]; vals = list(vc.values)
    fig, ax = plt.subplots(figsize=(5.5, 5.5))
    wedges, _, autotexts = ax.pie(vals, labels=None, colors=PIE_COLS[:len(vals)],
        autopct=lambda p: f'{p:.1f}%' if p>3 else '', startangle=140, pctdistance=0.72,
        wedgeprops={'edgecolor':'white','linewidth':2.0})
    for at in autotexts: at.set_fontsize(9.5); at.set_color('white'); at.set_fontweight('bold')
    ax.legend(wedges, [f'{l} (n={v})' for l,v in zip(labels,vals)],
              loc='lower center', bbox_to_anchor=(0.5,-0.22), ncol=2, fontsize=8, frameon=False)
    ax.set_aspect('equal'); ax.set_facecolor('white'); fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 9.0, 8.5)

def g_dias_principal(R):
    datos = R['dias_princ']
    if not datos: return None
    labs  = list(datos.keys()); proms = [datos[l]['prom'] for l in labs]; ns = [datos[l]['n'] for l in labs]
    fig, ax = plt.subplots(figsize=(max(4.5, len(labs)*0.9), 3.5))
    cols = [MC_MID if p==max(proms) else MC_LIGHT for p in proms]
    bars = ax.bar(labs, proms, color=cols, width=0.55, zorder=3)
    for bar, p, n in zip(bars, proms, ns):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.15,
                f'{p}d\n(n={n})', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#333')
    ax.set_ylim(0, max(proms)*1.38); ax.set_ylabel('Promedio días (0–28)', fontsize=9, color='#595959')
    ax.tick_params(labelsize=9); ax_style(ax); fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.5)

def g_consumo_pct(R):
    datos = R['consumo_pct']
    if not datos: return None
    labs  = list(datos.keys()); vals = [datos[l]['pct'] for l in labs]; ns = [datos[l]['n'] for l in labs]
    fig, ax = plt.subplots(figsize=(max(4.5, len(labs)*0.9), 3.5))
    cols = [MC_MID if v==max(vals) else MC_LIGHT for v in vals]
    bars = ax.bar(labs, vals, color=cols, width=0.55, zorder=3)
    for bar, v, n in zip(bars, vals, ns):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.4,
                f'{v}%\n(n={n})', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#333')
    ax.set_ylim(0, max(vals)*1.4); ax.set_ylabel('% de personas', fontsize=9, color='#595959')
    ax.tick_params(labelsize=9); ax_style(ax); fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.5)

def g_dias_sust(R):
    datos = R['dias_sust']
    if not datos: return None
    labs  = list(datos.keys()); proms = [datos[l]['prom'] for l in labs]; ns = [datos[l]['n'] for l in labs]
    fig, ax = plt.subplots(figsize=(max(4.5, len(labs)*0.9), 3.5))
    cols = [MC_MID if p==max(proms) else MC_LIGHT for p in proms]
    bars = ax.bar(labs, proms, color=cols, width=0.55, zorder=3)
    for bar, p, n in zip(bars, proms, ns):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.15,
                f'{p}d\n(n={n})', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#333')
    ax.set_ylim(0, max(proms)*1.38); ax.set_ylabel('Promedio días (0–28)', fontsize=9, color='#595959')
    ax.tick_params(labelsize=9); ax_style(ax); fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.5)

def g_donut(R):
    n_si = R['n_transgresores']; n_no = R['N'] - n_si
    fig, ax = plt.subplots(figsize=(5.2, 5.2))
    wedges, _, autotexts = ax.pie([n_si,n_no], labels=None, colors=[MC_MID,MC_LIGHT],
        autopct='%1.1f%%', startangle=90, pctdistance=0.75,
        wedgeprops={'edgecolor':'white','linewidth':2.5,'width':0.52}, counterclock=False)
    for at in autotexts: at.set_fontsize(13); at.set_fontweight('bold'); at.set_color('white')
    ax.legend(wedges, [f'Cometió transgresión (n={n_si})',f'Sin transgresión (n={n_no})'],
              loc='lower center', bbox_to_anchor=(0.5,-0.08), fontsize=9, frameon=False)
    ax.set_aspect('equal'); ax.set_facecolor('white'); fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 9.0, 7.5)

def g_tipos(R):
    tipos = R['transgtipos']
    labs  = [t['label'] for t in tipos]; vals = [t['pct'] for t in tipos]; ns = [t['n'] for t in tipos]
    fig, ax = plt.subplots(figsize=(4.5, max(3.0, len(labs)*0.55)))
    cols = [MC_MID if v==max(vals) else MC_LIGHT for v in vals]
    bars = ax.barh(labs, vals, color=cols, zorder=3)
    for bar, v, n in zip(bars, vals, ns):
        ax.text(bar.get_width()+0.5, bar.get_y()+bar.get_height()/2,
                f'{v}% (n={n})', va='center', fontsize=9, color='#333')
    ax.set_xlim(0, max(vals)*1.6 if vals else 1)
    ax.set_xlabel(f'% sobre {R["n_transgresores"]} transgresores', fontsize=8.5, color='#595959')
    ax.tick_params(labelsize=9.5); ax_style(ax, horiz=True)
    fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.5)

def g_salud(R):
    labels = [s['label'] for s in R['salud']]; proms = [s['prom'] for s in R['salud']]
    fig, ax = plt.subplots(figsize=(4.5, 3.2))
    bars = ax.barh(labels, proms, color=[MC_MID,MC_LIGHT,MC_ACCENT], zorder=3)
    for bar, p in zip(bars, proms):
        ax.text(bar.get_width()+0.1, bar.get_y()+bar.get_height()/2,
                f'{p}/20', va='center', fontsize=10, fontweight='bold', color='#333')
    ax.set_xlim(0, 24); ax.axvline(x=10, color='#BFBFBF', linestyle='--', linewidth=1.0)
    ax.tick_params(labelsize=10); ax_style(ax, horiz=True)
    fig.patch.set_facecolor('white'); fig.tight_layout()
    return to_rl(fig, 8.5, 5.5)

def g_vivienda(R):
    fig, ax = plt.subplots(figsize=(4.2, 3.2))
    cats  = ['Lugar\nestable','Condiciones\nbásicas']
    n_si  = [R['viv1'][0], R['viv2'][0]]
    n_no  = [R['viv1'][2]-R['viv1'][0], R['viv2'][2]-R['viv2'][0]]
    nv    = [R['viv1'][2], R['viv2'][2]]
    x, w2 = np.arange(len(cats)), 0.35
    b1 = ax.bar(x-w2/2, n_si, w2, label='Sí', color=MC_MID,  zorder=3)
    b2 = ax.bar(x+w2/2, n_no, w2, label='No', color=MC_GRAY, zorder=3)
    for bar, n, nv_ in zip(list(b1)+list(b2), n_si+n_no, nv+nv):
        pct = round(n/nv_*100,1) if nv_>0 else 0
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.3,
                f'{n}\n({pct}%)', ha='center', va='bottom', fontsize=9, color='#333')
    ax.set_xticks(x); ax.set_xticklabels(cats, fontsize=10)
    ax.set_ylim(0, max(n_si+n_no)*1.45 if max(n_si+n_no)>0 else 1)
    ax.set_ylabel('N personas', fontsize=9, color='#595959')
    ax.legend(fontsize=9, frameon=False); ax_style(ax)
    fig.patch.set_facecolor('white'); fig.tight_layout()
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
    ST['ptit'] = ParagraphStyle('ptit', fontName='Helvetica-Bold',   fontSize=26, textColor=white,
                   leading=32, alignment=TA_CENTER)
    ST['psub'] = ParagraphStyle('psub', fontName='Helvetica',        fontSize=14, textColor=C_LIGHT,
                   leading=20, alignment=TA_CENTER)
    ST['pserv']= ParagraphStyle('pserv',fontName='Helvetica-Bold',   fontSize=20, textColor=white,
                   leading=26, alignment=TA_CENTER)
    ST['kval'] = ParagraphStyle('kval', fontName='Helvetica-Bold',   fontSize=26, textColor=C_MID,
                   leading=32, alignment=TA_CENTER)
    ST['klab'] = ParagraphStyle('klab', fontName='Helvetica',        fontSize=9,  textColor=C_GRAY,
                   leading=12, alignment=TA_CENTER)
    return ST

def bloque(titulo, img, parrafos, ST):
    CW  = TW * 0.53; TXW = TW * 0.47
    row = Table([[img, parrafos]], colWidths=[CW, TXW])
    row.setStyle(TableStyle([
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('LEFTPADDING',(0,0),(-1,-1),0), ('RIGHTPADDING',(0,0),(-1,-1),5),
        ('TOPPADDING',(0,0),(-1,-1),2),  ('BOTTOMPADDING',(0,0),(-1,-1),2),
    ]))
    return [Paragraph(titulo, ST['sub']), row]

def hr(): return HRFlowable(width='100%', thickness=0.5, color=C_LIGHT, spaceAfter=4, spaceBefore=4)
def sp(h=0.2): return Spacer(1, h*cm)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTRUCCIÓN DEL PDF
# ══════════════════════════════════════════════════════════════════════════════
def build_pdf(R):
    ST  = make_styles()
    doc = SimpleDocTemplate(OUTPUT_FILE, pagesize=A4,
          leftMargin=1.7*cm, rightMargin=1.7*cm, topMargin=1.4*cm, bottomMargin=1.5*cm)
    S = []

    # ── PÁG 1: PORTADA ────────────────────────────────────────────────────────
    cover = Table([
        [Paragraph('INFORME DE CARACTERIZACIÓN', ST['ptit'])],
        [sp(0.5)],
        [Paragraph('Monitoreo de Resultados de Tratamiento<br/>Instrumento TOP', ST['psub'])],
        [sp(2.2)],
        [Paragraph(NOMBRE_SERVICIO.upper(), ST['pserv'])],
        [sp(0.4)],
        [Paragraph(PERIODO, ParagraphStyle('pp', fontName='Helvetica', fontSize=12,
            textColor=C_LIGHT, alignment=TA_CENTER))],
    ], colWidths=[TW])
    cover.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1), C_DARK),
        ('BACKGROUND',(0,4),(-1,6), C_MID),
        ('ALIGN',(0,0),(-1,-1),'CENTER'), ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(0,0),115), ('BOTTOMPADDING',(0,-1),(-1,-1),115),
        ('LEFTPADDING',(0,0),(-1,-1),20), ('RIGHTPADDING',(0,0),(-1,-1),20),
    ]))
    S += [cover, PageBreak()]

    # ── PÁG 2: PRESENTACIÓN + KPIs + G1 SEXO ─────────────────────────────────
    S.append(Paragraph('  PRESENTACIÓN', ST['sec'])); S.append(sp(0.2))
    S.append(Paragraph(
        f'El presente informe describe el perfil de las personas que ingresan a tratamiento por '
        f'consumo de sustancias en <b>{NOMBRE_SERVICIO}</b>, durante el período <b>{PERIODO}</b>, '
        f'a través del instrumento TOP (Treatment Outcomes Profile). '
        f'Durante este período ingresaron <b>{R["N"]} personas</b>; '
        f'la sustancia de mayor problema al ingreso fue el <b>{R["sust_top1"]} ({R["sust_top1_pct"]}%)</b>. '
        f'El {R["pct_hombre"]}% de las personas son hombres y el {R["pct_mujer"]}% mujeres.',
        ST['body']))
    S.append(sp(0.2))

    kpi = Table([[
        Paragraph(str(R['N']),             ST['kval']),
        Paragraph(f'{R["pct_hombre"]}%',   ST['kval']),
        Paragraph(str(R['edad_media']),     ST['kval']),
    ],[
        Paragraph('Personas<br/>ingresaron', ST['klab']),
        Paragraph('Son<br/>hombres',         ST['klab']),
        Paragraph('Edad<br/>promedio',        ST['klab']),
    ]], colWidths=[TW/3]*3)
    kpi.setStyle(TableStyle([
        ('ALIGN',(0,0),(-1,-1),'CENTER'), ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('BACKGROUND',(0,0),(-1,-1), HexColor('#EEF4FB')),
        ('BOX',(0,0),(-1,-1),0.5,C_LIGHT), ('INNERGRID',(0,0),(-1,-1),0.5,C_LIGHT),
        ('TOPPADDING',(0,0),(-1,-1),8), ('BOTTOMPADDING',(0,0),(-1,-1),8),
    ]))
    S += [kpi, sp(0.25), hr()]

    S.append(Paragraph('  1. ANTECEDENTES GENERALES', ST['sec'])); S.append(sp(0.1))
    S += bloque('1.1. Distribución de Personas según Sexo', g_sexo(R), [
        Paragraph(f'Del total de <b>{R["N"]} personas</b> que ingresaron a tratamiento, '
                  f'<b>{R["n_hombre"]} ({R["pct_hombre"]}%) son hombres</b> y '
                  f'<b>{R["n_mujer"]} ({R["pct_mujer"]}%) son mujeres</b>.', ST['body']),
        sp(0.15),
        Paragraph('Esta distribución es consistente con el perfil observado en servicios de '
                  'tratamiento de la región, donde los hombres representan la gran mayoría de los usuarios.', ST['body']),
        sp(0.15),
        Paragraph(f'N válido: {R["nv_sex"]} personas.', ST['nota']),
    ], ST)
    S.append(PageBreak())

    # ── PÁG 3: G2 EDAD + G3 TORTA SUSTANCIA PRINCIPAL ────────────────────────
    rm = max(R['edad_dist'], key=R['edad_dist'].get) if R['edad_dist'] else 'Sin datos'
    nm = R['edad_dist'].get(rm, 0)
    pm = round(nm/R['nv_edad']*100,1) if R['nv_edad']>0 else 0
    S += bloque('1.2. Distribución de Personas según Edad', g_edad(R), [
        Paragraph(f'El <b>promedio de edad</b> es de <b>{R["edad_media"]} años</b> '
                  f'(DE={R["edad_sd"]}). La persona más joven tiene {R["edad_min"]} años '
                  f'y la mayor {R["edad_max"]} años.', ST['body']),
        sp(0.15),
        Paragraph(f'El rango etario más frecuente es <b>{rm}</b>, con '
                  f'<b>{nm} personas ({pm}%)</b> del total.', ST['body']),
        sp(0.15),
        Paragraph(f'N válido: {R["nv_edad"]} personas.', ST['nota']),
    ], ST)
    S += [sp(0.2), hr()]

    S.append(Paragraph('  2. CONSUMO DE SUSTANCIAS', ST['sec'])); S.append(sp(0.1))
    vc = R['sust_vc']
    seg2    = vc.index[1].replace('\n',' ') if len(vc)>1 else '—'
    pct_s2  = round(vc.iloc[1]/R['nv_sust']*100,1) if len(vc)>1 and R['nv_sust']>0 else 0
    S += bloque('2.1. Consumo Sustancia Principal', g_torta_sust(R), [
        Paragraph('El gráfico muestra la distribución de las personas según la '
                  '<b>sustancia que consideran que les genera más problemas</b> '
                  'al ingresar a tratamiento.', ST['body']),
        sp(0.15),
        Paragraph(f'La sustancia más frecuente es el <b>{R["sust_top1"]} ({R["sust_top1_pct"]}%)</b>, '
                  f'seguida por <b>{seg2} ({pct_s2}%)</b>.', ST['body']),
        sp(0.15),
        Paragraph(f'N válido: {R["nv_sust"]} personas.', ST['nota']),
    ], ST)
    S.append(PageBreak())

    # ── PÁG 4: G4 DÍAS SUST. PRINCIPAL + G5 CONSUMO % + G6 DÍAS SUSTANCIA ────
    dias_p = R['dias_princ']
    if dias_p:
        dp_max = max(dias_p, key=lambda k: dias_p[k]['prom'])
        img_g4 = g_dias_principal(R)
        if img_g4:
            S += bloque('2.2. Promedio de Días de Consumo por Sustancia Principal', img_g4, [
                Paragraph('El gráfico muestra el promedio de días de consumo en las '
                          '<b>últimas 4 semanas</b>, calculado entre quienes '
                          'declararon cada sustancia como su principal problema.', ST['body']),
                sp(0.15),
                Paragraph(f'<b>{dp_max}</b> presenta el mayor promedio: '
                          f'<b>{dias_p[dp_max]["prom"]} días</b> (n={dias_p[dp_max]["n"]}).', ST['body']),
                sp(0.15),
                Paragraph('Promedio calculado solo sobre quienes declararon esa sustancia como principal.', ST['nota']),
            ], ST)
            S += [sp(0.2), hr()]

    if R['consumo_pct']:
        cp = R['consumo_pct']; sk = max(cp, key=lambda k: cp[k]['pct'])
        S += bloque('2.3. Consumo de Sustancias', g_consumo_pct(R), [
            Paragraph('El gráfico muestra el porcentaje de personas que consume cada sustancia '
                      'al ingreso. <b>Los porcentajes pueden sumar más de 100%</b> ya que una '
                      'misma persona puede consumir más de una sustancia.', ST['body']),
            sp(0.15),
            Paragraph(f'<b>{sk}</b> es la más prevalente: '
                      f'<b>{cp[sk]["pct"]}%</b> ({cp[sk]["n"]} personas).', ST['body']),
            sp(0.15),
            Paragraph(f'N total: {R["N"]} personas.', ST['nota']),
        ], ST)
        S += [sp(0.2), hr()]

    if R['dias_sust']:
        ds = R['dias_sust']; dk = max(ds, key=lambda k: ds[k]['prom'])
        S += bloque('2.4. Promedio de Días de Consumo por Sustancia', g_dias_sust(R), [
            Paragraph('El gráfico muestra el promedio de días de consumo por sustancia en las '
                      '<b>últimas 4 semanas</b>, calculado sobre todos los consumidores '
                      'de cada sustancia.', ST['body']),
            sp(0.15),
            Paragraph(f'<b>{dk}</b> tiene el mayor promedio: '
                      f'<b>{ds[dk]["prom"]} días</b> (n={ds[dk]["n"]}).', ST['body']),
            sp(0.15),
            Paragraph('Promedio calculado solo entre consumidores (días > 0).', ST['nota']),
        ], ST)
    S.append(PageBreak())

    # ── PÁG 5: G7 DONUT + G8 TIPOS + G9 SALUD + G10 VIVIENDA ────────────────
    S.append(Paragraph('  3. TRANSGRESIÓN A LA NORMA SOCIAL', ST['sec'])); S.append(sp(0.1))
    n_no_tr = R['N'] - R['n_transgresores']
    pct_no  = round(n_no_tr/R['N']*100,1) if R['N']>0 else 0
    S += bloque('3.1. Transgresión a la Norma Social', g_donut(R), [
        Paragraph(f'<b>{R["n_transgresores"]} personas ({R["pct_transgresores"]}%)</b> '
                  f'declararon haber cometido algún tipo de transgresión durante el mes '
                  f'previo al ingreso a tratamiento.', ST['body']),
        sp(0.15),
        Paragraph(f'Las <b>{n_no_tr} personas restantes ({pct_no}%)</b> '
                  f'no reportaron ningún incidente.', ST['body']),
        sp(0.15),
        Paragraph(f'N total: {R["N"]} personas.', ST['nota']),
    ], ST)
    S += [sp(0.2), hr()]

    if R['transgtipos']:
        tm = max(R['transgtipos'], key=lambda t: t['n'])
        S += bloque('3.2. Distribución por Tipo de Transgresión', g_tipos(R), [
            Paragraph('El gráfico ilustra el porcentaje de personas que declaró haber cometido '
                      'cada tipo de transgresión en el mes previo al ingreso.', ST['body']),
            sp(0.15),
            Paragraph(f'El tipo más frecuente es <b>{tm["label"]} '
                      f'({tm["pct"]}%, n={tm["n"]})</b>. '
                      f'Los porcentajes no suman 100% porque una misma persona puede '
                      f'haber cometido más de un tipo.', ST['body']),
            sp(0.15),
            Paragraph(f'N base: {R["n_transgresores"]} personas con al menos una transgresión.', ST['nota']),
        ], ST)
        S += [sp(0.2), hr()]

    S.append(Paragraph('  4. SALUD, CALIDAD DE VIDA Y VIVIENDA', ST['sec'])); S.append(sp(0.1))
    if R['salud']:
        mejor = max(R['salud'], key=lambda s: s['prom'])
        S += bloque('4.1. Autopercepción del Estado de Salud', g_salud(R), [
            Paragraph('El gráfico muestra los puntajes promedio de autopercepción de salud y '
                      'calidad de vida al ingreso, en una escala de <b>0 (muy mal) a 20 (excelente)</b>. '
                      'La línea punteada indica el punto medio de la escala (10).', ST['body']),
            sp(0.15),
            Paragraph(f'La dimensión mejor evaluada es <b>{mejor["label"]} ({mejor["prom"]}/20)</b>. '
                      f'Puntajes bajo 10 indican percepción deficiente.', ST['body']),
            sp(0.15),
            Paragraph(f'N válido: {R["salud"][0]["nv"]} personas por dimensión.', ST['nota']),
        ], ST)
        S.append(sp(0.15))

    n1,p1,nv1 = R['viv1']; n2,p2,nv2 = R['viv2']
    S += bloque('4.2. Condiciones de Vivienda al Ingreso', g_vivienda(R), [
        Paragraph(f'El <b>{p1}%</b> de las personas ({n1} de {nv1}) declaran tener '
                  f'un <b>lugar estable donde vivir</b>, y el <b>{p2}%</b> ({n2} de {nv2}) '
                  f'habita en una vivienda que <b>cumple condiciones básicas</b>.', ST['body']),
        sp(0.15),
        Paragraph('Estos indicadores permiten contextualizar las condiciones de vulnerabilidad '
                  'habitacional de las personas al inicio del tratamiento.', ST['body']),
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
    print('  SCRIPT_TOP_Universal_PDF_Caracterizacion  —  Iniciando...')
    print('=' * 60)
    print('\n→ Calculando datos...')
    R = cargar_datos()
    print(f'  N={R["N"]} | {R["sust_top1"]} {R["sust_top1_pct"]}% | Transgr.: {R["n_transgresores"]}')
    print('\n→ Generando PDF...')
    build_pdf(R)
    print(f'\n{"=" * 60}')
    print(f'  ✅  LISTO  →  {OUTPUT_FILE}')
    print(f'{"=" * 60}')
