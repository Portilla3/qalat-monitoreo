#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   SCRIPT_TOP_Universal_PPTX_Seguimiento.py                                 ║
║   Genera presentación PowerPoint de seguimiento TOP1 vs TOP2               ║
║   6 slides · Compatible con cualquier país TOP                             ║
║   Versión Universal 1.0                                                    ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                             ║
║  CÓMO USAR LA PRÓXIMA VEZ:                                                 ║
║  1. Abre un chat nuevo con Claude                                           ║
║  2. Sube DOS archivos:                                                      ║
║       • Este script: SCRIPT_TOP_Universal_PPTX_Seguimiento.py              ║
║       • La base en formato Wide (generada por SCRIPT_TOP_Universal_Wide)   ║
║  3. Escribe exactamente:                                                    ║
║     "Ejecuta el script universal PPTX Seguimiento con esta base Wide"      ║
║  4. Claude ajustará NOMBRE_SERVICIO y PERIODO según corresponda            ║
║                                                                             ║
║  SLIDES:                                                                    ║
║    1. Portada                                                               ║
║    2. Consumo sustancia principal (torta)                                  ║
║    3. Promedio días de consumo TOP1 vs TOP2                                ║
║    4. Cambio en consumo (barras apiladas) + tabla resumen                  ║
║    5. Transgresión a la norma social                                       ║
║    6. Salud, Calidad de Vida y Vivienda                                    ║
║                                                                             ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""


import glob, os, unicodedata

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
OUTPUT_FILE     = '/home/claude/TOP_Presentacion_Seguimiento.pptx'
# ── FILTRO OPCIONAL POR CENTRO ────────────────────────────────────────────────
# Dejar en None para procesar TODOS los centros.
# Poner el código exacto del centro para filtrar solo ese centro.
# Ejemplos:
#   FILTRO_CENTRO = None         ← todos los centros
#   FILTRO_CENTRO = "HCHN01"     ← solo ese centro
FILTRO_CENTRO = None
# ─────────────────────────────────────────────────────────────────────────────

NOMBRE_SERVICIO = 'Perú'                        # ← Claude ajusta
PERIODO         = '2025 – 2026'                 # ← Claude ajusta

# ══════════════════════════════════════════════════════════════════════════════
import pandas as pd, numpy as np, json, subprocess, os, sys, warnings
warnings.filterwarnings('ignore')

def _es_positivo(valor):
    s = str(valor).strip().lower()
    if s in ('sí', 'si'): return True
    if s in ('no', 'no aplica', 'nunca', 'nan', ''): return False
    n = pd.to_numeric(valor, errors='coerce')
    return not pd.isna(n) and n > 0

print('=' * 60)
print('  SCRIPT_TOP_Universal_PPTX_Seguimiento  —  Iniciando...')
print('=' * 60)

# ── Carga ──────────────────────────────────────────────────────────────────
print(f'\n→ Leyendo: {INPUT_FILE}')
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
    OUTPUT_FILE = f'/home/claude/TOP_Presentacion_Seguimiento_{FILTRO_CENTRO}.pptx'
N_total = len(df)
seg = df[df['Tiene_TOP2'] == 'Sí'].copy().reset_index(drop=True)
N = len(seg)
print(f'  Total pacientes:        {N_total}')
print(f'  Con seguimiento (TOP2): {N}  ({round(N/N_total*100,1)}%)')

# ── Tiempo de seguimiento (días entre TOP1 y TOP2) ──────────────────────────
_fc1 = next((c for c in seg.columns if 'fecha entrevista' in c.lower() and c.endswith('_TOP1')), None)
_fc2 = next((c for c in seg.columns if 'fecha entrevista' in c.lower() and c.endswith('_TOP2')), None)
_seg_tiempo = {'mediana': None, 'media': None, 'min': None, 'max': None, 'n': 0}
if _fc1 and _fc2:
    _d1 = pd.to_datetime(seg[_fc1], errors='coerce')
    _d2 = pd.to_datetime(seg[_fc2], errors='coerce')
    _dias = (_d2 - _d1).dt.days
    # Excluir valores atípicos (> 24 meses = 730 días) para rango y mediana
    _dias_ok = _dias[(_dias >= 0) & (_dias <= 730)].dropna()
    if len(_dias_ok) > 0:
        _m = _dias_ok / 30.44
        _seg_tiempo = {
            'mediana': round(float(_m.median()), 1),
            'media':   round(float(_m.mean()), 1),
            'min':     round(float(_m.min()), 1),
            'max':     round(float(_m.max()), 1),
            'n':       len(_dias_ok),
            'n_total': int(_dias.notna().sum())
        }
_med = _seg_tiempo['mediana']; _mn = _seg_tiempo['min']; _mx = _seg_tiempo['max']
print(f'  Tiempo seguimiento: mediana={_med} meses  rango={_mn}–{_mx} meses  (N válido={_seg_tiempo["n"]})')

# ── Detección dinámica de columnas ─────────────────────────────────────────
def detectar_columnas(cols):
    col_set = set(cols)
    def par(c1):
        if not c1: return (None, None)
        c2 = c1.replace('_TOP1', '_TOP2')
        return (c1, c2 if c2 in col_set else None)

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

    tr_sn = []
    for c in cols:
        if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('3)') and '>>' in c:
            nombre = c.replace('_TOP1','').split('>>')[-1].strip()
            c1, c2 = par(c)
            tr_sn.append((nombre, c1, c2))

    def find(conds):
        for c in cols:
            if not c.endswith('_TOP1'): continue
            base = c.replace('_TOP1','')
            if all(cond(base, c) for cond in conds):
                return par(c)
        return (None, None)

    vif     = find([lambda b,c: b.startswith('4)'), lambda b,c: 'Violencia Intrafamiliar' in c, lambda b,c: 'Total (0-28)' in c])
    sal_psi = find([lambda b,c: b.startswith('6)')])
    sal_fis = find([lambda b,c: b.startswith('8)')])
    cal_vid = find([lambda b,c: b.startswith('10)')])
    viv1    = find([lambda b,c: '9)' in b, lambda b,c: 'estable' in c.lower()])
    viv2    = find([lambda b,c: '9)' in b, lambda b,c: 'básicas' in c.lower()])
    sust_pp = find([lambda b,c: b.startswith('2)'), lambda b,c: 'sustancia principal' in c.lower()])

    print(f'  Sustancias: {[s[0] for s in sust_cols]}')
    print(f'  Transgresión: {[t[0] for t in tr_sn]}')
    return dict(sust_cols=sust_cols, tr_sn=tr_sn,
                vif=vif, sal_psi=sal_psi, sal_fis=sal_fis, cal_vid=cal_vid,
                viv1=viv1, viv2=viv2, sust_pp=sust_pp)

DC = detectar_columnas(seg.columns.tolist())

# ── Helpers ────────────────────────────────────────────────────────────────
def pct(n, d): return round(n/d*100, 1) if d > 0 else 0
def smean(col):
    if not col or col not in seg.columns: return 0
    v = pd.to_numeric(seg[col], errors='coerce')
    return round(float(v.mean()), 1) if v.notna().sum() > 0 else 0
def viv_pct(col):
    if not col or col not in seg.columns: return 0
    nv = int(seg[col].isin(['Sí','No']).sum()) or N
    return pct(int((seg[col]=='Sí').sum()), nv)

# ── Normalización sustancia principal ──────────────────────────────────────
def norm_sust(s):
    if pd.isna(s) or str(s).strip() == '0': return None
    s = str(s).strip().lower()
    if any(x in s for x in ['alcohol','cerveza','licor','aguard','alchol','bebida']): return 'Alcohol'
    if any(x in s for x in ['marihu','marjhu','marhuana','cannabis','cannbis']): return 'Cannabis/Marihuana'
    if any(x in s for x in ['pasta base','pasta','papelillo']): return 'Pasta Base'
    if any(x in s for x in ['crack','cristal','piedra','paco']): return 'Crack/Cristal'
    if any(x in s for x in ['cocain','perico','coca ']): return 'Cocaína'
    if any(x in s for x in ['tabaco','cigarr','nicot']): return 'Tabaco'
    if any(x in s for x in ['sedant','benzod','tranqui']): return 'Sedantes'
    if any(x in s for x in ['opiod','heroina','morfin','fentanil']): return 'Opiáceos'
    if any(x in s for x in ['metanfet','anfetam']): return 'Metanfetamina'
    return 'Otras'

print('\n→ Calculando indicadores...')
data = {}

# Slide 2: Sustancia principal (torta)
c1_sp, _ = DC['sust_pp']
if c1_sp:
    sr1 = seg[c1_sp].apply(norm_sust).dropna()
    nv = len(sr1); vc = sr1.value_counts()
    data['sust'] = [{'label': k, 'pct': round(v/nv*100, 1)} for k,v in vc.items()]
    data['sust_top'] = data['sust'][0]['label'] if data['sust'] else '—'
else:
    data['sust'] = []; data['sust_top'] = '—'

# Slide 3: Días de consumo TOP1 vs TOP2
dias = []
for lbl, c1, c2 in DC['sust_cols']:
    v1 = pd.to_numeric(seg[c1], errors='coerce')
    v2 = pd.to_numeric(seg[c2], errors='coerce') if c2 else pd.Series([np.nan]*N)
    m1 = round(float(v1.mean()), 1) if v1.notna().sum() > 0 else 0
    m2 = round(float(v2.mean()), 1) if (c2 and v2.notna().sum() > 0) else 0
    if m1 > 0 or m2 > 0:
        dias.append({'label': lbl, 'top1': m1, 'top2': m2})
data['dias'] = dias

# Slide 4: Cambio en consumo
cambio = []
for lbl, c1, c2 in DC['sust_cols']:
    if not c2: continue
    v1 = pd.to_numeric(seg[c1], errors='coerce').fillna(0)
    v2 = pd.to_numeric(seg[c2], errors='coerce').fillna(0)
    mask = v1 > 0; nc = int(mask.sum())
    if nc < 2: continue
    s1 = v1[mask]; s2 = v2[mask]
    n_abs = int((s2==0).sum()); n_dis = int(((s2>0)&(s2<s1)).sum())
    n_sc  = int((s2==s1).sum()); n_emp = int((s2>s1).sum())
    p = lambda n: round(n/nc*100, 1)
    cambio.append({'label': lbl, 'n_cons': nc,
        'abs': p(n_abs), 'dis': p(n_dis), 'sin': p(n_sc), 'emp': p(n_emp),
        'combo': round((n_abs+n_dis)/nc*100, 1)})
data['cambio'] = cambio

# Slide 5: Transgresión
tr_cols1 = [c1 for _,c1,_ in DC['tr_sn']]
tr_cols2 = [c2 for _,_,c2 in DC['tr_sn']]
vif_c1, vif_c2 = DC['vif']

def has_tr(row, sn_cols, vif_col):
    for c in sn_cols:
        if c and _es_positivo(row.get(c, '')): return True
    if vif_col:
        v = pd.to_numeric(row.get(vif_col, np.nan), errors='coerce')
        return not np.isnan(v) and v > 0
    return False

tr1 = seg.apply(lambda r: int(has_tr(r, tr_cols1, vif_c1)), axis=1)
tr2 = seg.apply(lambda r: int(has_tr(r, tr_cols2, vif_c2)), axis=1)
data['transgTotal'] = {'top1': pct(int(tr1.sum()), N), 'top2': pct(int(tr2.sum()), N)}

tipos = []
for lbl, c1, c2 in DC['tr_sn']:
    n1 = int(seg[c1].apply(_es_positivo).sum()) if c1 else 0
    n2 = int(seg[c2].apply(_es_positivo).sum()) if c2 else 0
    tipos.append({'label': lbl, 'top1': pct(n1,N), 'top2': pct(n2,N)})
if vif_c1:
    vif1_v = pd.to_numeric(seg[vif_c1], errors='coerce')
    vif2_v = pd.to_numeric(seg[vif_c2], errors='coerce') if vif_c2 else pd.Series([np.nan]*N)
    tipos.append({'label': 'VIF', 'top1': pct(int((vif1_v>0).sum()),N),
                                   'top2': pct(int((vif2_v>0).sum()),N)})
data['transgtipos'] = tipos

# Slide 6: Salud y Vivienda
sal_psi_c1, sal_psi_c2 = DC['sal_psi']
sal_fis_c1, sal_fis_c2 = DC['sal_fis']
cal_vid_c1, cal_vid_c2 = DC['cal_vid']
viv1_c1, viv1_c2 = DC['viv1']
viv2_c1, viv2_c2 = DC['viv2']

data['salud'] = [
    {'label': 'Salud Psicológica', 'top1': smean(sal_psi_c1), 'top2': smean(sal_psi_c2)},
    {'label': 'Salud Física',      'top1': smean(sal_fis_c1), 'top2': smean(sal_fis_c2)},
    {'label': 'Calidad de Vida',   'top1': smean(cal_vid_c1), 'top2': smean(cal_vid_c2)},
]
data['vivienda'] = [
    {'label': 'Lugar estable',       'top1': viv_pct(viv1_c1), 'top2': viv_pct(viv1_c2)},
    {'label': 'Condiciones básicas', 'top1': viv_pct(viv2_c1), 'top2': viv_pct(viv2_c2)},
]
data['meta'] = {
    'N': N, 'total': N_total, 'servicio': NOMBRE_SERVICIO,
    'periodo': PERIODO, 'pct_seg': round(N/N_total*100, 1),
    'seg_mediana': _seg_tiempo['mediana'],
    'seg_min':     _seg_tiempo['min'],
    'seg_max':     _seg_tiempo['max'],
    'seg_n':       _seg_tiempo['n'],
    'seg_n_total': _seg_tiempo.get('n_total', N)
}

print(f'  Sust. principal: {data["sust_top"]}')
print(f'  Transgresión: TOP1={data["transgTotal"]["top1"]}% → TOP2={data["transgTotal"]["top2"]}%')
if data['salud']:
    print(f'  Salud psic:    TOP1={data["salud"][0]["top1"]} → TOP2={data["salud"][0]["top2"]}')

# ── JSON intermedio ────────────────────────────────────────────────────────
json_path = '/home/claude/_top_data.json'
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

# ══════════════════════════════════════════════════════════════════════════════
# NODE.JS — construye el PowerPoint con pptxgenjs
# ══════════════════════════════════════════════════════════════════════════════
JS_CODE = r"""
const fs      = require('fs');
const pptxgen = require('pptxgenjs');

const data   = JSON.parse(fs.readFileSync('/home/claude/_top_data.json', 'utf8'));
const OUTPUT = '""" + OUTPUT_FILE + r"""';

const C_DARK  = '1F3864', C_MID = '2E75B6', C_LIGHT = 'BDD7EE';
const C_TOP1  = '1F3864', C_TOP2 = '2E75B6';
const C_TITLE = '0070C0', C_GRAY = '595959', C_WHITE = 'FFFFFF';
// Paleta cambio en consumo (tonos azules)
const C_ABS = '1F3864', C_DIS = '2E75B6', C_SC = '9DC3E6', C_EMP = 'BDD7EE';
const PIE_COLORS = ['2E75B6','1F3864','4472C4','9DC3E6','00B0F0','538135','D9D9D9','C00000','ED7D31'];

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

// Header estándar para slides de contenido
function hdr(sl, txt) {
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:10,h:0.72,fill:{color:C_DARK},line:{color:C_DARK}});
  sl.addShape(pres.shapes.RECTANGLE, {x:5.5,y:0,w:4.5,h:0.72,
    fill:{color:C_MID,transparency:40},line:{color:C_MID,transparency:40}});
  sl.addText(txt, {x:0.25,y:0,w:9.5,h:0.72,
    fontSize:22,bold:true,color:C_WHITE,fontFace:'Calibri',valign:'middle'});
}
const TITULO = `Resultados: Ingreso y Seguimiento · ${data.meta.servicio}`;

// ── SLIDE 1: PORTADA ──────────────────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  // Panel izquierdo oscuro
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:3.8,h:5.625,
    fill:{color:C_DARK},line:{color:C_DARK}});
  sl.addShape(pres.shapes.RECTANGLE, {x:3.0,y:0,w:1.5,h:5.625,
    fill:{color:C_MID,transparency:60},line:{color:C_MID,transparency:60}});
  // Texto panel izquierdo
  sl.addText('Resultados', {x:0.25,y:1.8,w:3.0,h:0.7,
    fontSize:24,bold:true,color:C_WHITE,fontFace:'Calibri'});
  sl.addText('Monitoreo TOP', {x:0.25,y:2.55,w:3.0,h:0.55,
    fontSize:14,color:C_LIGHT,fontFace:'Calibri'});
  // Texto panel derecho
  sl.addText([
    {text:'TOP 1 - TOP 2', options:{breakLine:true}},
    {text:'Ingreso y Seguimiento'}
  ], {x:4.3,y:1.8,w:5.4,h:1.4,
    fontSize:32,bold:true,color:C_GRAY,fontFace:'Calibri',align:'center',valign:'middle'});
  sl.addText(data.meta.servicio.toUpperCase(), {x:4.3,y:3.3,w:5.4,h:0.45,
    fontSize:18,bold:true,color:C_MID,fontFace:'Calibri',align:'center'});
  sl.addText(data.meta.periodo, {x:4.3,y:3.8,w:5.4,h:0.35,
    fontSize:13,color:C_MID,fontFace:'Calibri',align:'center',bold:true});
  sl.addText(
    `N = ${data.meta.N} pacientes con seguimiento (${data.meta.pct_seg}% del total)`,
    {x:4.3,y:4.25,w:5.4,h:0.35,fontSize:11,color:'888888',fontFace:'Calibri',align:'center'});
  if (data.meta.seg_mediana !== null) {
    const rangoTxt = data.meta.seg_min !== null
      ? `Mediana: ${data.meta.seg_mediana} meses  ·  Rango: ${data.meta.seg_min}–${data.meta.seg_max} meses`
      : `Mediana tiempo de seguimiento: ${data.meta.seg_mediana} meses`;
    sl.addText(rangoTxt,
      {x:4.3,y:4.62,w:5.4,h:0.3,fontSize:9.5,color:'AAAAAA',fontFace:'Calibri',
       align:'center',italic:true});
  }
}

// ── SLIDE 2: SUSTANCIA PRINCIPAL (torta) ──────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addText('CONSUMO SUSTANCIA PRINCIPAL AL INGRESO',
    {x:1.5,y:0.82,w:7,h:0.38,
     fontSize:14,bold:true,color:C_TITLE,fontFace:'Calibri',align:'center'});
  if (data.sust.length > 0) {
    sl.addChart(pres.charts.PIE, [{
      name:'Sustancia',
      labels: data.sust.map(s => s.label),
      values: data.sust.map(s => s.pct),
    }], {
      x:1.3,y:1.28,w:7.4,h:4.1,
      showPercent:true,showLabel:false,showLegend:true,legendPos:'b',legendFontSize:10,
      dataLabelFontSize:11,
      chartColors: PIE_COLORS.slice(0, data.sust.length),
      chartArea:{fill:{color:'FFFFFF'}},
      dataLabelColor:'FFFFFF',dataLabelPosition:'bestFit',
    });
  }
}

// ── SLIDE 3: DÍAS DE CONSUMO TOP1 vs TOP2 ────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addText('PROMEDIO DE DÍAS DE CONSUMO EN LAS ÚLTIMAS 4 SEMANAS\nTOP 1 (Ingreso) vs TOP 2 (Seguimiento)',
    {x:1.0,y:0.80,w:8,h:0.65,
     fontSize:13,bold:true,color:C_TITLE,fontFace:'Calibri',align:'center'});
  if (data.dias.length > 0) {
    const labels = data.dias.map(d => d.label);
    sl.addChart(pres.charts.BAR, [
      {name:'Ingreso (TOP 1)',    labels, values:data.dias.map(d=>d.top1)},
      {name:'Seguimiento (TOP 2)',labels, values:data.dias.map(d=>d.top2)},
    ], {
      x:0.8,y:1.5,w:8.4,h:3.9,barDir:'col',barGrouping:'clustered',
      chartColors:[C_TOP1,C_TOP2],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:12,
      valAxisLabelColor:'595959',valAxisLabelFontSize:10,
      valAxisMaxVal:28,valAxisMinVal:0,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
      showLegend:true,legendPos:'b',legendFontSize:11,
    });
  }
}

// ── SLIDE 4: CAMBIO EN CONSUMO (barras apiladas + tabla) ─────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addText('CAMBIO EN EL CONSUMO POR SUSTANCIA  ·  Ingreso → Seguimiento',
    {x:0.25,y:0.82,w:9.5,h:0.38,
     fontSize:13,bold:true,color:C_TITLE,fontFace:'Calibri',align:'center'});
  if (data.cambio.length > 0) {
    const labels = data.cambio.map(d => d.label);
    sl.addChart(pres.charts.BAR, [
      {name:'Abstinencia', labels, values:data.cambio.map(d=>d.abs)},
      {name:'Disminuyó',   labels, values:data.cambio.map(d=>d.dis)},
      {name:'Sin cambios', labels, values:data.cambio.map(d=>d.sin)},
      {name:'Empeoró',     labels, values:data.cambio.map(d=>d.emp)},
    ], {
      x:0.3,y:0.88,w:6.2,h:4.52,barDir:'col',barGrouping:'percentStacked',
      chartColors:[C_ABS,C_DIS,C_SC,C_EMP],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:10,dataLabelColor:C_WHITE,
      catAxisLabelColor:'363636',catAxisLabelFontSize:12,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
      showLegend:true,legendPos:'b',legendFontSize:10,
    });
    // Panel derecho: tabla resumen % abstinencia + disminuyó
    sl.addText('Abstinencia o\nreducción al\nseguimiento',
      {x:6.65,y:0.95,w:3.1,h:0.85,
       fontSize:12,color:C_GRAY,fontFace:'Calibri',align:'center',valign:'top'});
    const colW = 3.5 / data.cambio.length;
    sl.addTable([
      data.cambio.map(d => ({text:d.label,       options:{bold:true,fontSize:10,color:'363636',align:'center'}})),
      data.cambio.map(d => ({text:`${d.combo}%`, options:{bold:true,fontSize:14,color:C_DARK,  align:'center'}})),
    ], {
      x:6.25,y:1.95,w:3.5,h:0.9,
      border:{pt:0.5,color:'BDD7EE'},fill:{color:'EEF4FB'},
      rowH:0.42, colW: data.cambio.map(() => colW),
    });
  }
}

// ── SLIDE 5: TRANSGRESIÓN ────────────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  // Línea divisoria vertical
  sl.addShape(pres.shapes.LINE, {x:4.95,y:0.78,w:0,h:4.85,
    line:{color:'D9D9D9',width:1}});
  sl.addText('Personas que cometieron alguna\ntransgresión a la norma social',
    {x:0.25,y:0.82,w:4.5,h:0.65,
     fontSize:13,bold:true,color:C_GRAY,fontFace:'Calibri',align:'left'});
  sl.addText('Distribución por tipo de transgresión',
    {x:5.1,y:0.82,w:4.7,h:0.65,
     fontSize:13,bold:true,color:C_GRAY,fontFace:'Calibri',align:'center'});
  const T = data.transgTotal;
  // Gráfico general (2 barras: TOP1 y TOP2)
  sl.addChart(pres.charts.BAR, [
    {name:'TOP 1', labels:['Ingreso\n(TOP 1)','Seguimiento\n(TOP 2)'], values:[T.top1, null]},
    {name:'TOP 2', labels:['Ingreso\n(TOP 1)','Seguimiento\n(TOP 2)'], values:[null,    T.top2]},
  ], {
    x:0.2,y:1.5,w:4.5,h:3.8,barDir:'col',barGrouping:'clustered',
    chartColors:[C_TOP1,C_TOP2],chartArea:{fill:{color:'FFFFFF'}},
    showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:14,
    dataLabelColor:C_WHITE,dataLabelPosition:'inEnd',
    catAxisLabelColor:'363636',catAxisLabelFontSize:12,
    valAxisLabelColor:'595959',valAxisLabelFontSize:9,
    valAxisMaxVal:100,valAxisNumFmt:'0"%"',
    valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
  });
  // Gráfico tipos (filtrar los que tengan al menos un valor > 0)
  const tiposFilt = data.transgtipos.filter(d => d.top1 > 0 || d.top2 > 0);
  const tiposUsar = tiposFilt.length > 0 ? tiposFilt : data.transgtipos;
  sl.addChart(pres.charts.BAR, [
    {name:'Ingreso (TOP 1)',    labels:tiposUsar.map(d=>d.label), values:tiposUsar.map(d=>d.top1)},
    {name:'Seguimiento (TOP 2)',labels:tiposUsar.map(d=>d.label), values:tiposUsar.map(d=>d.top2)},
  ], {
    x:5.1,y:1.5,w:4.7,h:3.8,barDir:'col',barGrouping:'clustered',
    chartColors:[C_TOP1,C_TOP2],chartArea:{fill:{color:'FFFFFF'}},
    showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:11,dataLabelColor:'363636',
    catAxisLabelColor:'363636',catAxisLabelFontSize:9,
    valAxisLabelColor:'595959',valAxisLabelFontSize:9,
    valAxisMaxVal:80,valAxisNumFmt:'0"%"',
    valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
    showLegend:true,legendPos:'b',legendFontSize:10,
  });
}

// ── SLIDE 6: SALUD Y VIVIENDA ─────────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addShape(pres.shapes.LINE, {x:5.05,y:0.78,w:0,h:4.85,
    line:{color:'D9D9D9',width:1}});
  sl.addText('AUTOPERCEPCIÓN DEL ESTADO DE SALUD\nY CALIDAD DE VIDA (escala 0–20)',
    {x:0.25,y:0.82,w:4.7,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  sl.addText('CONDICIONES DE VIVIENDA\n(% con condición al Sí)',
    {x:5.3,y:0.82,w:4.5,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  // Salud (barras horizontales agrupadas)
  sl.addChart(pres.charts.BAR, [
    {name:'Ingreso (TOP 1)',    labels:data.salud.map(d=>d.label), values:data.salud.map(d=>d.top1)},
    {name:'Seguimiento (TOP 2)',labels:data.salud.map(d=>d.label), values:data.salud.map(d=>d.top2)},
  ], {
    x:0.2,y:1.5,w:4.6,h:3.8,barDir:'bar',barGrouping:'clustered',
    chartColors:[C_TOP1,C_TOP2],chartArea:{fill:{color:'FFFFFF'}},
    showValue:true,dataLabelFontSize:11,dataLabelColor:'363636',
    catAxisLabelColor:'363636',catAxisLabelFontSize:11,
    valAxisLabelColor:'595959',valAxisLabelFontSize:9,valAxisMaxVal:20,
    valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
    showLegend:true,legendPos:'b',legendFontSize:10,
  });
  // Vivienda (barras horizontales agrupadas)
  sl.addChart(pres.charts.BAR, [
    {name:'Ingreso (TOP 1)',    labels:data.vivienda.map(d=>d.label), values:data.vivienda.map(d=>d.top1)},
    {name:'Seguimiento (TOP 2)',labels:data.vivienda.map(d=>d.label), values:data.vivienda.map(d=>d.top2)},
  ], {
    x:5.15,y:1.5,w:4.6,h:3.8,barDir:'bar',barGrouping:'clustered',
    chartColors:[C_TOP1,C_TOP2],chartArea:{fill:{color:'FFFFFF'}},
    showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:12,dataLabelColor:'363636',
    catAxisLabelColor:'363636',catAxisLabelFontSize:11,
    valAxisLabelColor:'595959',valAxisLabelFontSize:9,valAxisMaxVal:100,
    valAxisNumFmt:'0"%"',
    valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
    showLegend:true,legendPos:'b',legendFontSize:10,
  });
}

pres.writeFile({fileName: OUTPUT})
  .then(() => { console.log('✅  PowerPoint guardado: ' + OUTPUT); })
  .catch(e => { console.error('Error JS:', e); process.exit(1); });
"""

js_path = '/home/claude/_top_builder.js'
with open(js_path, 'w', encoding='utf-8') as f:
    f.write(JS_CODE)

print('\n→ Construyendo PowerPoint con Node.js + pptxgenjs...')
result = subprocess.run(['node', js_path], capture_output=True, text=True)
if result.returncode != 0:
    print('ERROR en Node.js:')
    print(result.stderr)
    sys.exit(1)
print(result.stdout.strip())

# Limpieza de archivos temporales
os.remove(json_path)
os.remove(js_path)

print('\n' + '='*60)
print(f'  ✅  LISTO  →  {OUTPUT_FILE}')
print('='*60)
