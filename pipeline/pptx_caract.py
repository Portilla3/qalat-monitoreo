#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   SCRIPT_TOP_Universal_PPTX_Caracterizacion.py  —  v1.0                   ║
║   Genera presentación PowerPoint de caracterización al ingreso (TOP1)     ║
║   6 slides · Compatible con cualquier país TOP                            ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  CÓMO USAR:                                                                 ║
║  1. Sube este script + la base Wide TOP                                    ║
║  2. Escribe: "Ejecuta el PPTX Caracterización TOP"                        ║
║                                                                             ║
║  SLIDES:                                                                    ║
║    1. Portada                                                               ║
║    2. Antecedentes generales (sexo + edad + tabla KPIs)                   ║
║    3. Sustancia principal (torta)                                          ║
║    4. Días consumo sustancia principal                                     ║
║    5. % Consumidores + Días promedio por sustancia                        ║
║    6. Transgresión a la norma social (total + por tipo)                   ║
║    7. Salud, Calidad de Vida y Vivienda                                   ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import glob, os, unicodedata

def _norm(s):
    return unicodedata.normalize('NFD', str(s).lower()).encode('ascii','ignore').decode()

# ── Detección de país ─────────────────────────────────────────────────────────
_PAISES = {
    'republica_dominicana':'República Dominicana','repdomini':'República Dominicana',
    'dominicana':'República Dominicana','honduras':'Honduras',
    'panama':'Panamá','panam':'Panamá','el_salvador':'El Salvador',
    'salvador':'El Salvador','mexico':'México','mexic':'México',
    'ecuador':'Ecuador','peru':'Perú','argentina':'Argentina',
    'colombia':'Colombia','chile':'Chile','bolivia':'Bolivia',
    'paraguay':'Paraguay','uruguay':'Uruguay','venezuela':'Venezuela',
    'guatemala':'Guatemala','costa_rica':'Costa Rica',
    'costarica':'Costa Rica','nicaragua':'Nicaragua',
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

def auto_archivo_wide():
    candidatos = (
        glob.glob('/mnt/user-data/uploads/*Wide*.xlsx') +
        glob.glob('/mnt/user-data/uploads/*wide*.xlsx') +
        glob.glob('/mnt/user-data/uploads/TOP_Base*.xlsx') +
        glob.glob('/mnt/user-data/outputs/TOP_Base_Wide*.xlsx') +
        glob.glob('/home/claude/TOP_Base_Wide.xlsx'))
    if not candidatos:
        raise FileNotFoundError('⚠  No se encontró la base Wide TOP.')
    uploads = [f for f in candidatos if 'uploads' in f]
    elegido = uploads[0] if uploads else max(candidatos, key=os.path.getsize)
    print(f'  → Base Wide: {os.path.basename(elegido)}')
    return elegido

# ══════════════════════════════════════════════════════════════════════════════
print('=' * 60)
print('  SCRIPT_TOP_Universal_PPTX_Caracterizacion  v1.0')
print('=' * 60)

INPUT_FILE  = auto_archivo_wide()
SHEET_NAME  = 'Base Wide'
OUTPUT_FILE = '/home/claude/TOP_Presentacion_Caracterizacion.pptx'

# ── FILTRO OPCIONAL POR CENTRO ────────────────────────────────────────────────
FILTRO_CENTRO = None
# ─────────────────────────────────────────────────────────────────────────────

import pandas as pd, numpy as np, json, subprocess, sys, warnings
warnings.filterwarnings('ignore')

# ── Detección universal de transgresión (Sí/No o numérico) ───────────────────
def _es_positivo(valor):
    """True si el valor indica que ocurrió la transgresión.
       Soporta formato Sí/No Y formato numérico (valor > 0)."""
    s = str(valor).strip().lower()
    if s in ('sí', 'si'): return True
    if s in ('no', 'no aplica', 'nunca', 'nan', ''): return False
    n = pd.to_numeric(valor, errors='coerce')
    return not pd.isna(n) and n > 0

print(f'\n→ Leyendo: {INPUT_FILE}')
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME, header=1)
df.columns = [str(c) for c in df.columns]
cols = df.columns.tolist()

# Filtro de centro
_col_centro = next((c for c in cols if any(x in _norm(c) for x in
                    ['codigo del centro', 'servicio de tratamiento',
                     'centro/ servicio', 'codigo centro'])), None)
if FILTRO_CENTRO and _col_centro:
    n_antes = len(df)
    df = df[df[_col_centro].astype(str).str.strip() == FILTRO_CENTRO].copy()
    df = df.reset_index(drop=True)
    print(f'  ⚑ Filtro: "{FILTRO_CENTRO}" ({n_antes} → {len(df)} pacientes)')
    OUTPUT_FILE = f'/home/claude/TOP_Presentacion_Caracterizacion_{FILTRO_CENTRO}.pptx'

# País y servicio
_pais = _detectar_pais(INPUT_FILE)
if FILTRO_CENTRO:
    NOMBRE_SERVICIO = f'{_pais}  —  Centro {FILTRO_CENTRO}' if _pais else f'Centro {FILTRO_CENTRO}'
else:
    NOMBRE_SERVICIO = _pais if _pais else 'Servicio de Tratamiento'

# ── Detección dinámica de columnas ───────────────────────────────────────────
def detectar_columnas(cols):
    sust_cols = []
    for c in cols:
        if c.endswith('_TOP1') and 'Total (0-28)' in c:
            base = c.replace('_TOP1','')
            if base.startswith('1)'):
                partes = base.split('>>')
                if len(partes) >= 3:
                    nombre = partes[-2].strip().split('(')[0].strip()
                    sust_cols.append((nombre, c))
    tr_sn = []
    for c in cols:
        if c.endswith('_TOP1') and c.replace('_TOP1','').startswith('3)') and '>>' in c:
            nombre = c.replace('_TOP1','').split('>>')[-1].strip()
            tr_sn.append((nombre, c))
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

def norm_sust(s):
    if pd.isna(s) or str(s).strip() == '0': return None
    s = str(s).strip().lower()
    if any(x in s for x in ['alcohol','cerveza','licor','aguard']): return 'Alcohol'
    if any(x in s for x in ['marihu','cannabis','marij']):          return 'Marihuana'
    if any(x in s for x in ['pasta base','pasta','papelillo']):     return 'Pasta Base'
    if any(x in s for x in ['crack','cristal','piedra','paco']):    return 'Crack'
    if any(x in s for x in ['cocain','perico']):                    return 'Cocaína'
    if any(x in s for x in ['tabaco','cigarr','nicot']):            return 'Tabaco'
    if any(x in s for x in ['inhalant','thiner','activo']):         return 'Inhalantes'
    if any(x in s for x in ['sedant','benzod','tranqui']):          return 'Sedantes'
    if any(x in s for x in ['opiod','heroina','morfin','fentanil']): return 'Opiáceos'
    if any(x in s for x in ['metanfet','anfetam']):                 return 'Metanfetamina'
    return 'Otras'

# ── Cálculo de datos ─────────────────────────────────────────────────────────
N  = len(df)
DC = detectar_columnas(cols)
print(f'  N = {N}')
print(f'  Sustancias: {[s[0] for s in DC["sust_cols"]]}')

hoy = pd.Timestamp.now()

# Período
PERIODO = '2025'
if DC['fecha']:
    fch = pd.to_datetime(df[DC['fecha']], errors='coerce').dropna()
    fch = fch[(fch.dt.year >= 2010) & (fch.dt.year <= hoy.year+1)]
    if len(fch):
        MESES = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
                 7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'}
        f0, f1 = fch.min(), fch.max()
        PERIODO = (f'{MESES[f0.month]} {f0.year}'
                   if f0.year==f1.year and f0.month==f1.month
                   else f'{MESES[f0.month]}–{MESES[f1.month]} {f0.year}'
                   if f0.year==f1.year
                   else f'{MESES[f0.month]} {f0.year} – {MESES[f1.month]} {f1.year}')

# Sexo
n_h = n_m = nv_s = 0; pct_h = pct_m = 0.0
if DC['sexo']:
    sc  = df[DC['sexo']].astype(str).str.strip().str.upper()
    nv_s = int(sc.isin(['H','M']).sum())
    n_h = int((sc=='H').sum()); n_m = int((sc=='M').sum())
    pct_h = round(n_h/nv_s*100,1) if nv_s else 0
    pct_m = round(n_m/nv_s*100,1) if nv_s else 0

# Edad
edad_media = 0; edad_grupos = []
if DC['fn_col'] and DC['fecha']:
    fn  = pd.to_datetime(df[DC['fn_col']], errors='coerce')
    ref = pd.to_datetime(df[DC['fecha']], errors='coerce').fillna(hoy)
    edad = ((ref-fn).dt.days/365.25).round(1)
    edad = edad[(edad>=10)&(edad<=100)]
    if len(edad):
        edad_media = round(float(edad.mean()),1)
        bins = [0,17,30,40,50,60,200]
        labs = ['<18','18–30','31–40','41–50','51–60','61+']
        ec = pd.cut(edad,bins=bins,labels=labs)
        total_e = len(edad)
        edad_grupos = [{'label':l,'n':int((ec==l).sum()),
                        'pct':round(int((ec==l).sum())/total_e*100,1)} for l in labs
                       if int((ec==l).sum())>0]

# Sustancia principal
sust_ppal = []
if DC['sust_pp']:
    sr = df[DC['sust_pp']].apply(norm_sust).dropna()
    vc = sr.value_counts()
    total_sp = len(sr)
    sust_ppal = [{'label':k,'pct':round(v/total_sp*100,1),'n':int(v)}
                 for k,v in vc.items()]
    sust_top1 = vc.index[0] if len(vc) else '—'
    sust_top1_pct = round(vc.iloc[0]/total_sp*100,1) if len(vc) else 0
else:
    sust_top1 = '—'; sust_top1_pct = 0

# Días consumo por sustancia PRINCIPAL (solo quienes la declaran como principal)
sust_norm = df[DC['sust_pp']].apply(norm_sust) if DC['sust_pp'] else pd.Series([None]*N)
dias_princ = []
for lbl, col in DC['sust_cols']:
    v = pd.to_numeric(df[col], errors='coerce')
    mask = sust_norm.apply(lambda s: isinstance(s,str) and lbl.lower() in s.lower())
    sub  = v[mask & (v>0)].dropna()
    if len(sub):
        dias_princ.append({'label':lbl,'prom':round(float(sub.mean()),1),'n':int(len(sub))})
dias_princ.sort(key=lambda x:-x['prom'])

# % Consumidores por sustancia
consumo_pct = []
for lbl, col in DC['sust_cols']:
    v = pd.to_numeric(df[col], errors='coerce').fillna(0)
    n_c = int((v>0).sum())
    if n_c > 0:
        consumo_pct.append({'label':lbl,'pct':round(n_c/N*100,1),'n':n_c})
consumo_pct.sort(key=lambda x:-x['pct'])

# Promedio días por sustancia
dias_sust = []
for lbl, col in DC['sust_cols']:
    v = pd.to_numeric(df[col], errors='coerce'); sub = v[v>0].dropna()
    if len(sub):
        dias_sust.append({'label':lbl,'prom':round(float(sub.mean()),1),'n':int(len(sub))})
dias_sust.sort(key=lambda x:-x['prom'])

# Transgresión
def has_tr(row):
    for _, c in DC['tr_sn']:
        if _es_positivo(row.get(c, '')): return True
    if DC['vif']:
        v = pd.to_numeric(row.get(DC['vif'], np.nan), errors='coerce')
        return not np.isnan(v) and v > 0
    return False
t = df.apply(lambda r: int(has_tr(r)), axis=1)
n_tr  = int(t.sum()); pct_tr = round(n_tr/N*100,1)
tipos_tr = []
for lbl, col in DC['tr_sn']:
    n = int(df[col].apply(_es_positivo).sum())
    tipos_tr.append({'label':lbl,'n':n,'pct':round(n/N*100,1)})
if DC['vif']:
    vif_v = pd.to_numeric(df[DC['vif']], errors='coerce')
    n_vif = int((vif_v>0).sum())
    tipos_tr.append({'label':'VIF','n':n_vif,'pct':round(n_vif/N*100,1)})
tipos_tr = [t for t in tipos_tr if t['pct']>0]

# Salud
salud = []
for lbl, col in [('Salud Psicológica', DC['sal_psi']),
                 ('Salud Física',      DC['sal_fis']),
                 ('Calidad de Vida',   DC['cal_vid'])]:
    if col:
        v = pd.to_numeric(df[col], errors='coerce')
        salud.append({'label':lbl,'prom':round(float(v.mean()),1),
                      'nv':int(v.notna().sum())})

# Vivienda
def viv(col):
    if not col: return {'label':'—','pct':0,'n':0}
    nv_ = int(df[col].isin(['Sí','No']).sum()) or N
    n_  = int((df[col]=='Sí').sum())
    return {'n':n_,'pct':round(n_/nv_*100,1)}
viv1 = viv(DC['viv1']); viv2 = viv(DC['viv2'])

print(f'  Sust. principal top: {sust_top1} ({sust_top1_pct}%)')
print(f'  Transgresión: {pct_tr}%  |  Salud: {len(salud)} indicadores')

# ── JSON ─────────────────────────────────────────────────────────────────────
data = {
    'meta': {
        'servicio':   NOMBRE_SERVICIO,
        'periodo':    PERIODO,
        'N':          N,
        'pct_h':      pct_h,
        'pct_m':      pct_m,
        'n_h':        n_h,
        'n_m':        n_m,
        'edad_media': edad_media,
        'sust_top1':  sust_top1,
        'sust_top1_pct': sust_top1_pct,
    },
    'sexo':       [{'label':'Hombre','pct':pct_h,'n':n_h},
                   {'label':'Mujer', 'pct':pct_m,'n':n_m}],
    'edad':       edad_grupos,
    'sust':       sust_ppal,
    'dias_princ': dias_princ,
    'consumo':    consumo_pct,
    'dias':       dias_sust,
    'transTotal': {'pct':pct_tr,'n':n_tr},
    'transTipos': tipos_tr,
    'salud':      salud,
    'viv1':       viv1,
    'viv2':       viv2,
}

json_path = '/home/claude/_top_car_data.json'
with open(json_path,'w',encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

# ══════════════════════════════════════════════════════════════════════════════
# NODE.JS — pptxgenjs
# ══════════════════════════════════════════════════════════════════════════════
JS_CODE = r"""
const fs      = require('fs');
const pptxgen = require('pptxgenjs');

const data   = JSON.parse(fs.readFileSync('/home/claude/_top_car_data.json', 'utf8'));
const OUTPUT = '""" + OUTPUT_FILE + r"""';

const C_DARK  = '1F3864', C_MID = '2E75B6', C_LIGHT = 'BDD7EE';
const C_TOP1  = '1F3864';
const C_TITLE = '0070C0', C_GRAY = '595959', C_WHITE = 'FFFFFF';
const PIE_COLORS = ['2E75B6','1F3864','4472C4','9DC3E6','00B0F0','538135','D9D9D9','C00000','ED7D31'];

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

// Header estándar
function hdr(sl, txt) {
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:10,h:0.72,
    fill:{color:C_DARK},line:{color:C_DARK}});
  sl.addShape(pres.shapes.RECTANGLE, {x:5.5,y:0,w:4.5,h:0.72,
    fill:{color:C_MID,transparency:40},line:{color:C_MID,transparency:40}});
  sl.addText(txt, {x:0.25,y:0,w:9.5,h:0.72,
    fontSize:22,bold:true,color:C_WHITE,fontFace:'Calibri',valign:'middle'});
}

// Divisor vertical
function divV(sl, x) {
  sl.addShape(pres.shapes.LINE, {x,y:0.78,w:0,h:4.85,
    line:{color:'D9D9D9',width:1}});
}

const TITULO = `Caracterización al Ingreso · ${data.meta.servicio}`;

// ── SLIDE 1: PORTADA ──────────────────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:3.8,h:5.625,
    fill:{color:C_DARK},line:{color:C_DARK}});
  sl.addShape(pres.shapes.RECTANGLE, {x:3.0,y:0,w:1.5,h:5.625,
    fill:{color:C_MID,transparency:60},line:{color:C_MID,transparency:60}});
  sl.addText('Caracterización', {x:0.25,y:1.6,w:3.2,h:0.7,
    fontSize:22,bold:true,color:C_WHITE,fontFace:'Calibri'});
  sl.addText('Ingreso a Tratamiento · TOP', {x:0.25,y:2.35,w:3.2,h:0.55,
    fontSize:12,color:C_LIGHT,fontFace:'Calibri'});
  sl.addText([
    {text:'TOP 1', options:{breakLine:true}},
    {text:'Ingreso a Tratamiento'}
  ], {x:4.3,y:1.6,w:5.4,h:1.4,
    fontSize:30,bold:true,color:C_GRAY,fontFace:'Calibri',align:'center',valign:'middle'});
  sl.addText(data.meta.servicio.toUpperCase(), {x:4.3,y:3.1,w:5.4,h:0.45,
    fontSize:18,bold:true,color:C_MID,fontFace:'Calibri',align:'center'});
  sl.addText(data.meta.periodo, {x:4.3,y:3.62,w:5.4,h:0.35,
    fontSize:13,bold:true,color:C_MID,fontFace:'Calibri',align:'center'});
  sl.addText(`N = ${data.meta.N} personas al ingreso a tratamiento`,
    {x:4.3,y:4.1,w:5.4,h:0.35,fontSize:11,color:'888888',fontFace:'Calibri',align:'center'});
}

// ── SLIDE 2: ANTECEDENTES GENERALES ───────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  divV(sl, 4.95);

  // KPIs: N · % hombres · edad promedio
  const kpiY = 0.85;
  const kpis = [
    {val: String(data.meta.N),         lab:'Personas\ningresaron'},
    {val: `${data.meta.pct_h}%`,        lab:'Son\nhombres'},
    {val: String(data.meta.edad_media), lab:'Edad\npromedio'},
  ];
  kpis.forEach((k,i) => {
    const x = 0.18 + i * 1.55;
    sl.addShape(pres.shapes.RECTANGLE, {x, y:kpiY, w:1.42, h:0.88,
      fill:{color:'EEF4FB'}, line:{color:'BDD7EE',width:0.5}});
    sl.addText(k.val, {x, y:kpiY+0.04, w:1.42, h:0.48,
      fontSize:20, bold:true, color:C_DARK, fontFace:'Calibri',
      align:'center', valign:'middle'});
    sl.addText(k.lab, {x, y:kpiY+0.52, w:1.42, h:0.34,
      fontSize:9, color:C_GRAY, fontFace:'Calibri',
      align:'center', valign:'top'});
  });

  // Izquierda: torta sexo
  sl.addText('Distribución por Sexo',
    {x:0.25,y:1.85,w:4.5,h:0.35,
     fontSize:12,bold:true,color:C_GRAY,fontFace:'Calibri'});
  if (data.sexo.filter(s=>s.n>0).length > 0) {
    const sexoFilt = data.sexo.filter(s=>s.n>0);
    sl.addChart(pres.charts.PIE, [{
      name:'Sexo',
      labels: sexoFilt.map(s=>s.label),
      values: sexoFilt.map(s=>s.n),
    }], {
      x:0.4,y:2.22,w:4.2,h:3.15,
      showPercent:true,showLabel:false,showLegend:true,legendPos:'b',legendFontSize:11,
      dataLabelFontSize:13,chartColors:['2E75B6','9DC3E6'],
      chartArea:{fill:{color:'FFFFFF'}},dataLabelColor:C_WHITE,
    });
  }

  // Derecha: distribución etaria
  sl.addText('Distribución por Rango de Edad',
    {x:5.15,y:0.85,w:4.6,h:0.35,
     fontSize:12,bold:true,color:C_GRAY,fontFace:'Calibri'});
  if (data.meta.edad_media > 0) {
    sl.addText(`Promedio: ${data.meta.edad_media} años`,
      {x:5.15,y:1.22,w:4.6,h:0.28,
       fontSize:10,color:C_GRAY,fontFace:'Calibri',italic:true});
  }
  if (data.edad.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'Personas',
      labels: data.edad.map(e=>e.label),
      values: data.edad.map(e=>e.pct),
    }], {
      x:5.1,y:1.52,w:4.65,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_TOP1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:10,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:11,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }
}

// ── SLIDE 3: SUSTANCIA PRINCIPAL ─────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addText('CONSUMO SUSTANCIA PRINCIPAL AL INGRESO',
    {x:1.5,y:0.82,w:7,h:0.38,
     fontSize:14,bold:true,color:C_TITLE,fontFace:'Calibri',align:'center'});
  if (data.sust.length > 0) {
    sl.addChart(pres.charts.PIE, [{
      name:'Sustancia',
      labels: data.sust.map(s=>s.label),
      values: data.sust.map(s=>s.pct),
    }], {
      x:1.3,y:1.28,w:7.4,h:4.1,
      showPercent:true,showLabel:false,showLegend:true,legendPos:'b',legendFontSize:10,
      dataLabelFontSize:11,
      chartColors: PIE_COLORS.slice(0, data.sust.length),
      chartArea:{fill:{color:'FFFFFF'}},
      dataLabelColor:C_WHITE,dataLabelPosition:'bestFit',
    });
    // Nota sust. más frecuente
    sl.addText(
      `Sustancia más frecuente: ${data.meta.sust_top1} (${data.meta.sust_top1_pct}%)  ·  N = ${data.meta.N}`,
      {x:1.0,y:5.3,w:8,h:0.25,
       fontSize:9,color:C_GRAY,fontFace:'Calibri',align:'center',italic:true});
  }
}

// ── SLIDE 4: DÍAS CONSUMO SUSTANCIA PRINCIPAL ────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addText('PROMEDIO DE DÍAS DE CONSUMO DE LA SUSTANCIA PRINCIPAL\nÚltimas 4 semanas · solo personas cuya sust. principal corresponde a cada categoría',
    {x:0.5,y:0.82,w:9.0,h:0.65,
     fontSize:12,bold:true,color:C_TITLE,fontFace:'Calibri',align:'center'});
  if (data.dias_princ.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'Días promedio',
      labels: data.dias_princ.map(d=>d.label),
      values: data.dias_princ.map(d=>d.prom),
    }], {
      x:1.2,y:1.55,w:7.6,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_TOP1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:11,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:28,valAxisMinVal:0,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }
  sl.addText(`N = ${data.meta.N}  ·  Escala: días en últimas 4 semanas (0–28)`,
    {x:0.25,y:5.35,w:9.5,h:0.25,
     fontSize:8.5,color:'AAAAAA',fontFace:'Calibri',align:'center',italic:true});
}

// ── SLIDE 5: % CONSUMIDORES + DÍAS POR SUSTANCIA (todos) ─────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  divV(sl, 4.95);

  // Izquierda: % consumidores
  sl.addText('% DE PERSONAS QUE CONSUME\nCada sustancia al ingreso',
    {x:0.25,y:0.82,w:4.5,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.consumo.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'% Consumidores',
      labels: data.consumo.map(d=>d.label),
      values: data.consumo.map(d=>d.pct),
    }], {
      x:0.2,y:1.52,w:4.5,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_TOP1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:10,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:9,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:100,valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }

  // Derecha: días promedio todos consumidores
  sl.addText('PROMEDIO DE DÍAS DE CONSUMO\nPor sustancia (solo consumidores)',
    {x:5.15,y:0.82,w:4.6,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.dias.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'Días promedio',
      labels: data.dias.map(d=>d.label),
      values: data.dias.map(d=>d.prom),
    }], {
      x:5.15,y:1.52,w:4.6,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_TOP1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:10,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:9,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:28,valAxisMinVal:0,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }
  sl.addText(`N = ${data.meta.N}  ·  Una persona puede consumir más de una sustancia`,
    {x:0.25,y:5.35,w:9.5,h:0.25,
     fontSize:8.5,color:'AAAAAA',fontFace:'Calibri',align:'center',italic:true});
}

// ── SLIDE 5: TRANSGRESIÓN ─────────────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  divV(sl, 4.95);
  sl.addText('Personas que cometieron alguna\ntransgresión a la norma social',
    {x:0.25,y:0.82,w:4.5,h:0.65,
     fontSize:13,bold:true,color:C_GRAY,fontFace:'Calibri',align:'left'});
  sl.addText('Distribución por tipo de transgresión',
    {x:5.15,y:0.82,w:4.6,h:0.65,
     fontSize:13,bold:true,color:C_GRAY,fontFace:'Calibri',align:'center'});

  const T = data.transTotal;
  // Izquierda: barra total
  sl.addChart(pres.charts.BAR, [
    {name:'Con transgresión',    labels:['Con\ntransgresión','Sin\ntransgresión'],
     values:[T.pct, null]},
    {name:'Sin transgresión',    labels:['Con\ntransgresión','Sin\ntransgresión'],
     values:[null, parseFloat((100-T.pct).toFixed(1))]},
  ], {
    x:0.3,y:1.52,w:4.4,h:3.85,barDir:'col',barGrouping:'clustered',
    chartColors:[C_DARK,'BDD7EE'],chartArea:{fill:{color:'FFFFFF'}},
    showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:14,
    dataLabelColor:C_WHITE,dataLabelPosition:'inEnd',
    catAxisLabelColor:'363636',catAxisLabelFontSize:11,
    valAxisLabelColor:'595959',valAxisLabelFontSize:9,
    valAxisMaxVal:100,valAxisNumFmt:'0"%"',
    valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
  });
  // N y % encima del bloque
  sl.addText(`${T.n} personas (${T.pct}%)`,
    {x:0.3,y:1.22,w:4.4,h:0.28,
     fontSize:12,bold:true,color:C_DARK,fontFace:'Calibri',align:'center'});

  // Derecha: tipos
  const tiposFilt = data.transTipos.filter(d=>d.pct>0);
  if (tiposFilt.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'% personas',
      labels: tiposFilt.map(d=>d.label),
      values: tiposFilt.map(d=>d.pct),
    }], {
      x:5.1,y:1.52,w:4.65,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_TOP1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:9,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }
  sl.addText(`N = ${data.meta.N}`,
    {x:0.25,y:5.35,w:9.5,h:0.25,
     fontSize:8.5,color:'AAAAAA',fontFace:'Calibri',align:'center',italic:true});
}

// ── SLIDE 6: SALUD Y VIVIENDA ─────────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  divV(sl, 5.05);
  sl.addText('AUTOPERCEPCIÓN DEL ESTADO DE SALUD\nY CALIDAD DE VIDA (escala 0–20)',
    {x:0.25,y:0.82,w:4.7,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  sl.addText('CONDICIONES DE VIVIENDA\n(% con condición Sí)',
    {x:5.3,y:0.82,w:4.5,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});

  if (data.salud.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'Promedio',
      labels: data.salud.map(d=>d.label),
      values: data.salud.map(d=>d.prom),
    }], {
      x:0.2,y:1.52,w:4.6,h:3.85,barDir:'bar',barGrouping:'clustered',
      chartColors:[C_TOP1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:11,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,valAxisMaxVal:20,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }

  // Vivienda
  const vDatos = [
    {label:'Lugar\nestable',     pct: data.viv1.pct || 0},
    {label:'Condiciones\nbásicas', pct: data.viv2.pct || 0},
  ];
  sl.addChart(pres.charts.BAR, [{
    name:'% Sí',
    labels: vDatos.map(d=>d.label),
    values: vDatos.map(d=>d.pct),
  }], {
    x:5.15,y:1.52,w:4.6,h:3.85,barDir:'bar',barGrouping:'clustered',
    chartColors:[C_TOP1],chartArea:{fill:{color:'FFFFFF'}},
    showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:12,dataLabelColor:'363636',
    catAxisLabelColor:'363636',catAxisLabelFontSize:11,
    valAxisLabelColor:'595959',valAxisLabelFontSize:9,
    valAxisMaxVal:100,valAxisNumFmt:'0"%"',
    valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
  });
  sl.addText(`N = ${data.meta.N}  ·  ${data.meta.servicio}  ·  ${data.meta.periodo}`,
    {x:0.25,y:5.35,w:9.5,h:0.25,
     fontSize:8.5,color:'AAAAAA',fontFace:'Calibri',align:'center',italic:true});
}

pres.writeFile({fileName: OUTPUT})
  .then(() => { console.log('✅  PowerPoint guardado: ' + OUTPUT); })
  .catch(e  => { console.error('Error JS:', e); process.exit(1); });
"""

js_path = '/home/claude/_top_car_builder.js'
with open(js_path, 'w', encoding='utf-8') as f:
    f.write(JS_CODE)

print('\n→ Construyendo PowerPoint con Node.js + pptxgenjs...')
result = subprocess.run(['node', js_path], capture_output=True, text=True)
if result.returncode != 0:
    print('ERROR en Node.js:')
    print(result.stderr)
    sys.exit(1)
print(result.stdout.strip())

os.remove(json_path)
os.remove(js_path)

print('\n' + '='*60)
print(f'  ✅  LISTO  →  {OUTPUT_FILE}')
print(f'      {N} pacientes TOP1  ·  {PERIODO}')
print('='*60)
