"""
app.py — QALAT · Sistema de Monitoreo de Resultados de Tratamiento
v5.0 — login por país · Perú / Ecuador / México / UNODC
       + pestaña Corrección de registros (editar / eliminar en Supabase)
"""
import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import tempfile, os, sys
from io import BytesIO
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pipeline.wide_top import procesar_wide
from pipeline.runner   import run_script, run_paquetes_centros

NAVY='#1F3864'; MID='#2E75B6'; ACCENT='#00B0F0'
ORANGE='#C8590A'; RED='#C00000'; GREEN='#538135'; WHITE='#FFFFFF'

st.set_page_config(
    page_title='QALAT · TOP · Sistema de Monitoreo de Resultados de Tratamiento',
    page_icon='📊', layout='wide', initial_sidebar_state='collapsed'
)

st.markdown(f"""<style>
html,body,[class*="css"]{{font-family:'Calibri',sans-serif;}}
.main{{background:#F8FAFD;}}
.qalat-hdr{{background:{NAVY};color:white;padding:1.2rem 2rem;border-radius:8px;margin-bottom:1.5rem;border-left:8px solid {MID};}}
.qalat-hdr h1{{color:white;font-size:1.6rem;margin:0;}}
.qalat-hdr h1 .instrumento{{font-size:2.2rem;font-weight:900;color:#9DC3E6;margin-left:.2rem;}}
.qalat-hdr p{{color:#BDD7EE;font-size:.9rem;margin:.3rem 0 0 0;}}
.kpi{{background:white;border-radius:8px;padding:1rem 1.2rem;border-left:4px solid {MID};
      box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:.5rem;}}
.kpi.red{{border-left-color:{RED};}}.kpi.orange{{border-left-color:{ORANGE};}}.kpi.green{{border-left-color:{GREEN};}}
.kpi-lbl{{font-size:.78rem;color:#666;margin-bottom:.2rem;}}
.kpi-val{{font-size:1.8rem;font-weight:700;color:{NAVY};}}
.kpi-sub{{font-size:.75rem;color:#888;}}
.sec{{background:{MID};color:white;padding:.5rem 1rem;border-radius:6px;
      font-weight:600;font-size:1rem;margin:1.2rem 0 .8rem 0;}}
.filter-box{{background:white;border:1px solid #D0DFF0;border-radius:8px;padding:1rem 1.2rem;margin-bottom:1rem;}}
.filter-box h4{{color:{NAVY};margin:0 0 .6rem 0;font-size:.95rem;}}
.outcard{{background:white;border-radius:8px;padding:1rem;border:1px solid #D0DFF0;margin-bottom:.5rem;}}
.outcard h4{{color:{NAVY};margin:0 0 .3rem 0;font-size:.95rem;}}
.outcard p{{color:#666;font-size:.8rem;margin:0;}}
.badge{{display:inline-block;padding:3px 10px;border-radius:12px;font-size:.78rem;font-weight:600;margin-right:4px;}}
.badge-centro{{background:#E8F0FE;color:{NAVY};}}
.badge-periodo{{background:#E8F5E9;color:#1B5E20;}}
.login-box{{max-width:420px;margin:3rem auto;background:white;border-radius:12px;
            padding:2rem 2.5rem;box-shadow:0 4px 20px rgba(31,56,100,.12);
            border-top:5px solid {MID};}}
div.stButton>button{{background:#1E7E34;color:white;border:none;
    padding:.6rem 2rem;border-radius:6px;font-size:1rem;font-weight:600;width:100%;
    box-shadow:0 2px 6px rgba(30,126,52,.35);letter-spacing:.3px;}}
div.stButton>button:hover{{background:#145222;box-shadow:0 3px 10px rgba(30,126,52,.5);}}
#MainMenu,footer,header{{visibility:hidden;}}
</style>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE PAÍSES
# Para agregar México u otro país: solo añadir entrada aquí y PASSWORD_X en Secrets
# ══════════════════════════════════════════════════════════════════════════════
PAISES_CONFIG = {
    'Perú':    {'flag': '🇵🇪', 'color': MID},
    'Ecuador': {'flag': '🇪🇨', 'color': '#007A5E'},
    'México':  {'flag': '🇲🇽', 'color': '#006847'},
    'UNODC':   {'flag': '🌐', 'color': NAVY},
}
PAISES_ACTIVOS = ['Perú', 'Ecuador', 'México']   # ← agregar aquí para sumar países

SECRET_KEY_MAP = {
    'Perú':    'PASSWORD_PERU',
    'Ecuador': 'PASSWORD_ECUADOR',
    'México':  'PASSWORD_MEXICO',
    'UNODC':   'PASSWORD_UNODC',
}

LABELS = {
    'caract_excel': ('📋 Tablas caracterización', 'Excel',      '11 tablas al ingreso: sexo, edad, sustancias, transgresión'),
    'seg_excel':    ('📋 Tablas seguimiento',      'Excel',      'Comparativo TOP1 vs TOP2'),
    'pdf_caract':   ('📄 Word caracterización',    'Word',       '4 secciones · gráficos · tablas'),
    'pdf_seg':      ('📄 Word seguimiento',        'Word',       'Comparativo ingreso vs seguimiento'),
    'pptx_caract':  ('📑 PPT caracterización',     'PowerPoint', '6 slides · perfil al ingreso'),
    'pptx_seg':     ('📑 PPT seguimiento',         'PowerPoint', '6 slides · ingreso vs seguimiento'),
}

RENAME_MAP = {
    'codigo_paciente':     'Código de identificación del paciente',
    'fecha_entrevista':    'Fecha de entrevista TOP',
    'fecha_nacimiento':    'Fecha de nacimiento',
    'centro':              'Código del centro de tratamiento',
    'etapa':               'Etapa',
    'sexo':                'Sexo',
    'nombre_entrevistador':'Nombre entrevistador',
    'sustancia_principal': '¿Cuál considera que es la sustancia principal que genera más problemas?',
    'alcohol_s4':          'Alcohol Última Semana (0-7)',
    'alcohol_s3':          'Alcohol Semana 3 (0-7)',
    'alcohol_s2':          'Alcohol Semana 2 (0-7)',
    'alcohol_s1':          'Alcohol Semana 1 (0-7)',
    'alcohol_total':       'Alcohol Total (0-28)',
    'alcohol_prom':        'Alcohol Promedio/día',
    'marihuana_s4':        'Marihuana Última Semana (0-7)',
    'marihuana_s3':        'Marihuana Semana 3 (0-7)',
    'marihuana_s2':        'Marihuana Semana 2 (0-7)',
    'marihuana_s1':        'Marihuana Semana 1 (0-7)',
    'marihuana_total':     'Marihuana Total (0-28)',
    'marihuana_prom':      'Marihuana Promedio/día',
    'pastabase_s4':        'Pasta Base Última Semana (0-7)',
    'pastabase_s3':        'Pasta Base Semana 3 (0-7)',
    'pastabase_s2':        'Pasta Base Semana 2 (0-7)',
    'pastabase_s1':        'Pasta Base Semana 1 (0-7)',
    'pastabase_total':     'Pasta Base Total (0-28)',
    'pastabase_prom':      'Pasta Base Promedio/día',
    'cocaina_s4':          'Cocaína Última Semana (0-7)',
    'cocaina_s3':          'Cocaína Semana 3 (0-7)',
    'cocaina_s2':          'Cocaína Semana 2 (0-7)',
    'cocaina_s1':          'Cocaína Semana 1 (0-7)',
    'cocaina_total':       'Cocaína Total (0-28)',
    'cocaina_prom':        'Cocaína Promedio/día',
    'sedantes_s4':         'Sedantes Última Semana (0-7)',
    'sedantes_s3':         'Sedantes Semana 3 (0-7)',
    'sedantes_s2':         'Sedantes Semana 2 (0-7)',
    'sedantes_s1':         'Sedantes Semana 1 (0-7)',
    'sedantes_total':      'Sedantes Total (0-28)',
    'sedantes_prom':       'Sedantes Promedio/día',
    'hurto':               'Hurto',
    'robo':                'Robo',
    'venta_droga':         'Venta de droga',
    'rina_pelea':          'Riña/Pelea',
    'vif_s4':              'VIF Última Semana (0-7)',
    'vif_s3':              'VIF Semana 3 (0-7)',
    'vif_s2':              'VIF Semana 2 (0-7)',
    'vif_s1':              'VIF Semana 1 (0-7)',
    'vif_total':           'VIF Total (0-28)',
    'salud_psicologica':   'Salud Psicológica (0-20)',
    'salud_fisica':        'Salud Física (0-20)',
    'calidad_vida':        'Calidad de Vida (0-20)',
    'dias_trabajo_s4':     'Trabajo Última Semana (0-7)',
    'dias_trabajo_s3':     'Trabajo Semana 3 (0-7)',
    'dias_trabajo_s2':     'Trabajo Semana 2 (0-7)',
    'dias_trabajo_s1':     'Trabajo Semana 1 (0-7)',
    'dias_trabajo_total':  'Trabajo Total (0-28)',
    'dias_educacion_s4':   'Educación Última Semana (0-7)',
    'dias_educacion_s3':   'Educación Semana 3 (0-7)',
    'dias_educacion_s2':   'Educación Semana 2 (0-7)',
    'dias_educacion_s1':   'Educación Semana 1 (0-7)',
    'dias_educacion_total':'Educación Total (0-28)',
    'vivienda_estable':    'Vivienda estable',
    'vivienda_basica':     'Vivienda básica',
}


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS SUPABASE
# ══════════════════════════════════════════════════════════════════════════════
def _sb_headers():
    return {
        'apikey':        st.secrets['SUPABASE_KEY'],
        'Authorization': f"Bearer {st.secrets['SUPABASE_KEY']}",
        'Content-Type':  'application/json',
        'Prefer':        'return=representation',
    }

def _sb_url(tabla='top_registros'):
    return f"{st.secrets['SUPABASE_URL']}/rest/v1/{tabla}"

def _cargar_supabase(pais=None):
    import urllib.request, urllib.parse, json
    url = _sb_url() + '?select=*&order=fecha_entrevista.asc'
    if pais and pais != 'Todos':
        url += f"&pais=eq.{urllib.parse.quote(pais)}"
    req = urllib.request.Request(url, headers=_sb_headers())
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode('utf-8'))

def _actualizar_registro(registro_id, campos):
    import urllib.request, json
    url  = _sb_url() + f'?id=eq.{registro_id}'
    data = json.dumps(campos).encode('utf-8')
    req  = urllib.request.Request(url, data=data, method='PATCH', headers=_sb_headers())
    with urllib.request.urlopen(req) as r:
        return r.status

def _eliminar_registro(registro_id):
    import urllib.request
    url = _sb_url() + f'?id=eq.{registro_id}'
    req = urllib.request.Request(url, method='DELETE', headers=_sb_headers())
    with urllib.request.urlopen(req) as r:
        return r.status


# ══════════════════════════════════════════════════════════════════════════════
# LOGIN
# ══════════════════════════════════════════════════════════════════════════════
def _verificar_login(pais_sel, clave):
    secret_key = SECRET_KEY_MAP.get(pais_sel)
    if not secret_key:
        return False
    try:
        return clave == st.secrets[secret_key]
    except Exception:
        return False

def _mostrar_login():
    st.markdown("""
    <div style="text-align:center;margin-top:2rem;">
      <div style="font-size:2.8rem;">📊</div>
      <div style="font-size:1.8rem;font-weight:900;color:#1F3864;margin:.3rem 0;">QALAT · TOP</div>
      <div style="font-size:1rem;color:#2E75B6;margin-bottom:2rem;">
        Sistema de Monitoreo de Resultados de Tratamiento
      </div>
    </div>
    """, unsafe_allow_html=True)
    col_l, col_c, col_r = st.columns([1, 1.4, 1])
    with col_c:
        st.markdown('<div class="login-box">', unsafe_allow_html=True)
        st.markdown(f'<p style="color:{NAVY};font-weight:700;font-size:1.1rem;margin-bottom:1.2rem;">🔐 Acceso al sistema</p>', unsafe_allow_html=True)
        pais_sel = st.selectbox(
            'Selecciona tu país / institución',
            PAISES_ACTIVOS + ['UNODC'],
            format_func=lambda p: f"{PAISES_CONFIG[p]['flag']}  {p}",
            key='login_pais'
        )
        clave = st.text_input('Contraseña', type='password', key='login_clave',
                              placeholder='Ingresa tu contraseña')
        if st.button('Ingresar →', use_container_width=True, key='btn_login'):
            if _verificar_login(pais_sel, clave):
                st.session_state['autenticado'] = True
                st.session_state['rol_pais']    = pais_sel
                st.rerun()
            else:
                st.error('❌ Contraseña incorrecta. Intenta nuevamente.')
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown(
        '<p style="text-align:center;color:#aaa;font-size:.78rem;margin-top:2rem;">'
        '© Rodrigo Portilla · UNODC Chile · Proyecto QALAT</p>',
        unsafe_allow_html=True
    )

if not st.session_state.get('autenticado', False):
    _mostrar_login()
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# USUARIO AUTENTICADO
# ══════════════════════════════════════════════════════════════════════════════
rol       = st.session_state['rol_pais']
es_unodc  = (rol == 'UNODC')
pais_fijo = None if es_unodc else rol
flag      = PAISES_CONFIG[rol]['flag']
rol_lbl   = f'{flag} {rol}'

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""<div class="qalat-hdr">
  <h1>📊 QALAT · Monitoreo de Resultados de Tratamiento — Instrumento <span class="instrumento">TOP</span></h1>
  <p>Procesamiento automático TOP · Sube tu Excel, aplica filtros y descarga todos los reportes</p>
  <p style="margin-top:.4rem;font-size:.8rem;color:#9DC3E6;">Sesión activa: <b>{rol_lbl}</b></p>
  <p style="margin-top:.2rem;font-size:.75rem;color:#7fa8cc;">© Rodrigo Portilla · UNODC</p>
</div>""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('### 📋 Pasos')
    st.markdown('1. Selecciona fuente de datos\n2. Aplica filtros (opcional)\n3. Elige reportes\n4. Clic en **Procesar**\n5. Descarga')
    st.markdown('---')
    st.caption(f'QALAT v5.0 · {datetime.now().strftime("%d/%m/%Y")}')
    st.markdown(f'**Sesión:** {rol_lbl}')
    st.markdown('---')
    if st.button('🚪 Cerrar sesión', use_container_width=True, key='btn_logout'):
        for k in ['autenticado','rol_pais','supabase_path','supabase_df',
                  'filename','result','outputs','seleccion','wide_path','work_dir','raw_path',
                  'corr_registros','corr_editando','corr_confirm_del']:
            st.session_state.pop(k, None)
        st.rerun()
    st.markdown('---')
    st.markdown(
        '<div style="font-size:.75rem;color:#999;line-height:1.6;">'
        '© Rodrigo Portilla<br><span style="color:#bbb;">UNODC Chile · Proyecto QALAT</span>'
        '</div>', unsafe_allow_html=True
    )

# ══════════════════════════════════════════════════════════════════════════════
# PESTAÑAS
# ══════════════════════════════════════════════════════════════════════════════
tab_reportes, tab_correccion = st.tabs(['📊 Reportes', '✏️ Corrección de registros'])


# ──────────────────────────────────────────────────────────────────────────────
# TAB 1: REPORTES
# ──────────────────────────────────────────────────────────────────────────────
with tab_reportes:

    st.markdown('<div class="sec">📁 Cargar base de datos</div>', unsafe_allow_html=True)

    fuente = st.radio(
        'Fuente de datos',
        ['📁 Subir Excel (JotForm)', '📡 Conectar con Supabase (Piloto)'],
        horizontal=True,
        help='Elige si subes un Excel exportado de JotForm o conectas directo a la base del piloto'
    )

    uploaded      = None
    supabase_data = None

    # ── Fuente: Supabase ──────────────────────────────────────────────────────
    if fuente == '📡 Conectar con Supabase (Piloto)':
        if es_unodc:
            st.markdown(
                '<div style="background:#EEF4FB;border-left:4px solid #2E75B6;'
                'padding:.8rem 1.2rem;border-radius:6px;margin-bottom:1rem;">'
                '<b>🌐 Vista UNODC</b> — Puedes ver datos de todos los países o filtrar por uno.'
                '</div>', unsafe_allow_html=True
            )
            col_pais, col_btn = st.columns([2, 1])
            with col_pais:
                pais_filtro = st.selectbox('Ver datos de', ['Todos'] + PAISES_ACTIVOS, key='pais_sb')
            with col_btn:
                st.markdown('<div style="margin-top:28px"></div>', unsafe_allow_html=True)
                cargar_sb = st.button('📥 Descargar datos', use_container_width=True, key='btn_sb')
        else:
            pais_filtro = pais_fijo
            st.markdown(
                f'<div style="background:#EEF4FB;border-left:4px solid #2E75B6;'
                f'padding:.8rem 1.2rem;border-radius:6px;margin-bottom:1rem;">'
                f'<b>📡 Conexión directa a Supabase</b><br>'
                f'Descarga los registros de <b>{flag} {pais_fijo}</b> capturados en el formulario web del piloto.'
                f'</div>', unsafe_allow_html=True
            )
            cargar_sb = st.button('📥 Descargar datos', use_container_width=True, key='btn_sb')

        if cargar_sb:
            try:
                registros = _cargar_supabase(pais_filtro)
                if not registros:
                    st.warning('⚠ No hay registros en Supabase para ese filtro.')
                else:
                    df_sb = pd.DataFrame(registros)
                    pais_label = pais_filtro if pais_filtro else 'Todos'
                    st.success(f'✓ {len(df_sb)} registros descargados de Supabase ({pais_label})')
                    df_sb = df_sb.rename(columns={k: v for k, v in RENAME_MAP.items() if k in df_sb.columns})
                    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
                    df_sb.to_excel(tmp.name, index=False)
                    tmp.close()
                    st.session_state['supabase_path'] = tmp.name
                    st.session_state['supabase_df']   = df_sb
                    st.session_state['filename']      = f'Supabase_{pais_label}'
                    supabase_data = df_sb
            except KeyError:
                st.error('⚠ Las credenciales de Supabase no están configuradas en Secrets.')
            except Exception as e:
                st.error(f'Error al conectar con Supabase: {e}')

        elif 'supabase_path' in st.session_state:
            supabase_data = st.session_state.get('supabase_df')

    # ── Fuente: Excel ─────────────────────────────────────────────────────────
    else:
        uploaded = st.file_uploader(
            'Arrastra tu Excel aquí o haz clic para buscar',
            type=['xlsx','xls'],
            help='Archivo bruto exportado de Jotform — instrumento TOP'
        )

    # ── Filtros y procesamiento ───────────────────────────────────────────────
    filtro_centro_val = None
    fecha_desde_val   = None
    fecha_hasta_val   = None
    centros_disponibles = []

    if uploaded:
        @st.cache_data(show_spinner=False)
        def _leer_preview(file_bytes):
            import pandas as _pd, io, unicodedata
            def _n(s): return unicodedata.normalize('NFD',str(s).lower()).encode('ascii','ignore').decode()
            df = _pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=0)
            df.columns = [str(c) for c in df.columns]
            col_c = next((c for c in df.columns if any(k in _n(c) for k in
                          ['codigo del centro','centro de tratamiento','servicio de tratamiento'])
                          and 'trabajo' not in _n(c) and 'estudio' not in _n(c)), None)
            col_f = next((c for c in df.columns if any(k in _n(c) for k in
                          ['fecha entrevista','fecha_entrevista','fecha de entrevista'])), None)
            centros = sorted(df[col_c].dropna().astype(str).str.strip().unique().tolist()) if col_c else []
            fechas  = _pd.to_datetime(df[col_f], errors='coerce').dropna() if col_f else _pd.Series([], dtype='datetime64[ns]')
            return centros, fechas

        file_bytes = uploaded.getvalue()
        centros_disponibles, fechas_serie = _leer_preview(file_bytes)

        st.markdown('<div class="sec">🔍 Filtros (opcional — por defecto procesa todo)</div>', unsafe_allow_html=True)
        fc1, fc2, fc3 = st.columns([1.5, 1.5, 1])

        with fc1:
            st.markdown('<div class="filter-box"><h4>🏥 Filtrar por centro</h4>', unsafe_allow_html=True)
            opciones_centro = ['Todos los centros'] + centros_disponibles
            sel_centro = st.selectbox('Centro / Servicio', opciones_centro, label_visibility='collapsed')
            if sel_centro != 'Todos los centros':
                filtro_centro_val = sel_centro
            st.markdown('</div>', unsafe_allow_html=True)

        with fc2:
            st.markdown('<div class="filter-box"><h4>📅 Filtrar por período</h4>', unsafe_allow_html=True)
            MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
            anio_actual = datetime.now().year
            if len(fechas_serie):
                anio_min = max(fechas_serie.dt.year.min(), anio_actual - 10)
                anio_max = min(fechas_serie.dt.year.max(), anio_actual + 1)
            else:
                anio_min, anio_max = anio_actual - 3, anio_actual
            anios = list(range(int(anio_min), int(anio_max)+1))
            p1, p2 = st.columns(2)
            with p1:
                st.caption('Desde')
                mes_d  = st.selectbox('Mes inicio', MESES, index=0,  key='mes_d',  label_visibility='collapsed')
                anio_d = st.selectbox('Año inicio', anios, index=0,  key='anio_d', label_visibility='collapsed')
            with p2:
                st.caption('Hasta')
                mes_h  = st.selectbox('Mes fin',  MESES, index=11,              key='mes_h',  label_visibility='collapsed')
                anio_h = st.selectbox('Año fin',  anios, index=len(anios)-1,    key='anio_h', label_visibility='collapsed')
            usar_periodo = st.checkbox('Aplicar filtro de período', value=False)
            if usar_periodo:
                fecha_desde_val = f'{anio_d}-{MESES.index(mes_d)+1:02d}'
                fecha_hasta_val = f'{anio_h}-{MESES.index(mes_h)+1:02d}'
            st.markdown('</div>', unsafe_allow_html=True)

        with fc3:
            st.markdown('<div class="filter-box"><h4>📄 Reportes a generar</h4>', unsafe_allow_html=True)
            cb_ce  = st.checkbox('Tablas caracterización', value=False, key='cb_ce')
            cb_se  = st.checkbox('Tablas seguimiento',     value=False, key='cb_se')
            cb_pc  = st.checkbox('Word caracterización',   value=False, key='cb_pc')
            cb_ps  = st.checkbox('Word seguimiento',       value=False, key='cb_ps')
            cb_ppc = st.checkbox('PPT caracterización',    value=False, key='cb_ppc')
            cb_pps = st.checkbox('PPT seguimiento',        value=False, key='cb_pps')
            st.markdown('</div>', unsafe_allow_html=True)

        SELECCION = {
            'caract_excel': cb_ce, 'seg_excel': cb_se,
            'pdf_caract':   cb_pc, 'pdf_seg':   cb_ps,
            'pptx_caract':  cb_ppc,'pptx_seg':  cb_pps,
        }

        badges = ''
        if filtro_centro_val:
            badges += f'<span class="badge badge-centro">🏥 Centro: {filtro_centro_val}</span>'
        if fecha_desde_val:
            badges += f'<span class="badge badge-periodo">📅 {fecha_desde_val} → {fecha_hasta_val}</span>'
        if not badges:
            badges = '<span style="color:#888;font-size:.85rem">Sin filtros — procesa toda la base</span>'
        st.markdown(f'**Archivo:** `{uploaded.name}` &nbsp;|&nbsp; {badges}', unsafe_allow_html=True)

        if st.button('⚡ Procesar y generar reportes', use_container_width=True):
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                tmp.write(uploaded.read()); tmp_raw = tmp.name
            work_dir = tempfile.mkdtemp(prefix='qalat_')
            try:
                with st.spinner('Paso 1/7 — Procesando base Wide...'):
                    result = procesar_wide(tmp_raw, filtro_centro=filtro_centro_val,
                                           fecha_desde=fecha_desde_val, fecha_hasta=fecha_hasta_val)
                    st.session_state['result']    = result
                    st.session_state['filename']  = uploaded.name
                    st.session_state['seleccion'] = SELECCION
                    wide_path = os.path.join(work_dir, 'TOP_Base_Wide.xlsx')
                    with open(wide_path,'wb') as f:
                        f.write(result['excel_bytes'].getvalue())
                    st.session_state['wide_path'] = wide_path
                    st.session_state['work_dir']  = work_dir
                st.success(f"✅ Base Wide — {result['stats']['N_total']} pacientes · {result['periodo']}")
                outputs  = {}
                keys_sel = [k for k,v in SELECCION.items() if v]
                prog = st.progress(0, text='Generando reportes...')
                for i, key in enumerate(keys_sel):
                    prog.progress(i/len(keys_sel), text=f"Generando {LABELS[key][0]}...")
                    try:
                        buf, fname, mime = run_script(key, wide_path, filtro_centro=filtro_centro_val)
                        outputs[key] = {'ok':True,'buf':buf,'fname':fname,'mime':mime}
                    except Exception as e:
                        outputs[key] = {'ok':False,'error':str(e)}
                prog.progress(1.0, text='✅ Listo')
                st.session_state['outputs'] = outputs
            except Exception as e:
                st.error(f'❌ Error: {e}')
            finally:
                st.session_state['raw_path'] = tmp_raw

    elif supabase_data is not None:
        col_centro_sb = 'Código del centro de tratamiento'
        centros_disponibles = []
        if col_centro_sb in supabase_data.columns:
            centros_disponibles = sorted(supabase_data[col_centro_sb].dropna().astype(str).str.strip().unique().tolist())

        st.markdown('<div class="sec">🔍 Filtros (opcional)</div>', unsafe_allow_html=True)
        fc1, fc2, fc3 = st.columns([1.5, 1.5, 1])

        with fc1:
            st.markdown('<div class="filter-box"><h4>🏥 Filtrar por centro</h4>', unsafe_allow_html=True)
            opciones_centro = ['Todos los centros'] + centros_disponibles
            sel_centro = st.selectbox('Centro / Servicio', opciones_centro, label_visibility='collapsed', key='sb_centro')
            if sel_centro != 'Todos los centros':
                filtro_centro_val = sel_centro
            st.markdown('</div>', unsafe_allow_html=True)

        with fc2:
            st.markdown('<div class="filter-box"><h4>📅 Filtrar por período</h4>', unsafe_allow_html=True)
            MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
            anio_actual = datetime.now().year
            anios = list(range(anio_actual - 3, anio_actual + 1))
            p1, p2 = st.columns(2)
            with p1:
                st.caption('Desde')
                mes_d  = st.selectbox('Mes inicio', MESES, index=0,  key='sb_mes_d',  label_visibility='collapsed')
                anio_d = st.selectbox('Año inicio', anios, index=0,  key='sb_anio_d', label_visibility='collapsed')
            with p2:
                st.caption('Hasta')
                mes_h  = st.selectbox('Mes fin',  MESES, index=11,           key='sb_mes_h',  label_visibility='collapsed')
                anio_h = st.selectbox('Año fin',  anios, index=len(anios)-1, key='sb_anio_h', label_visibility='collapsed')
            usar_periodo_sb = st.checkbox('Aplicar filtro de período', value=False, key='sb_periodo')
            if usar_periodo_sb:
                fecha_desde_val = f'{anio_d}-{MESES.index(mes_d)+1:02d}'
                fecha_hasta_val = f'{anio_h}-{MESES.index(mes_h)+1:02d}'
            st.markdown('</div>', unsafe_allow_html=True)

        with fc3:
            st.markdown('<div class="filter-box"><h4>📄 Reportes a generar</h4>', unsafe_allow_html=True)
            cb_ce  = st.checkbox('Tablas caracterización', value=False, key='sb_cb_ce')
            cb_se  = st.checkbox('Tablas seguimiento',     value=False, key='sb_cb_se')
            cb_pc  = st.checkbox('Word caracterización',   value=False, key='sb_cb_pc')
            cb_ps  = st.checkbox('Word seguimiento',       value=False, key='sb_cb_ps')
            cb_ppc = st.checkbox('PPT caracterización',    value=False, key='sb_cb_ppc')
            cb_pps = st.checkbox('PPT seguimiento',        value=False, key='sb_cb_pps')
            st.markdown('</div>', unsafe_allow_html=True)

        SELECCION = {
            'caract_excel': cb_ce, 'seg_excel': cb_se,
            'pdf_caract':   cb_pc, 'pdf_seg':   cb_ps,
            'pptx_caract':  cb_ppc,'pptx_seg':  cb_pps,
        }

        st.markdown(f'**Fuente:** Supabase · `{len(supabase_data)}` registros descargados', unsafe_allow_html=True)

        if st.button('⚡ Procesar y generar reportes', use_container_width=True, key='btn_proc_sb'):
            tmp_raw  = st.session_state.get('supabase_path')
            work_dir = tempfile.mkdtemp(prefix='qalat_')
            try:
                with st.spinner('Paso 1/7 — Procesando base Wide desde Supabase...'):
                    result = procesar_wide(tmp_raw, filtro_centro=filtro_centro_val,
                                           fecha_desde=fecha_desde_val, fecha_hasta=fecha_hasta_val)
                    st.session_state['result']    = result
                    st.session_state['seleccion'] = SELECCION
                    wide_path = os.path.join(work_dir, 'TOP_Base_Wide.xlsx')
                    with open(wide_path,'wb') as f:
                        f.write(result['excel_bytes'].getvalue())
                    st.session_state['wide_path'] = wide_path
                    st.session_state['work_dir']  = work_dir
                st.success(f"✅ Base Wide — {result['stats']['N_total']} pacientes · {result['periodo']}")
                outputs  = {}
                keys_sel = [k for k,v in SELECCION.items() if v]
                prog = st.progress(0, text='Generando reportes...')
                for i, key in enumerate(keys_sel):
                    prog.progress(i/len(keys_sel), text=f"Generando {LABELS[key][0]}...")
                    try:
                        buf, fname, mime = run_script(key, wide_path, filtro_centro=filtro_centro_val)
                        outputs[key] = {'ok':True,'buf':buf,'fname':fname,'mime':mime}
                    except Exception as e:
                        outputs[key] = {'ok':False,'error':str(e)}
                prog.progress(1.0, text='✅ Listo')
                st.session_state['outputs'] = outputs
            except Exception as e:
                st.error(f'❌ Error al procesar datos Supabase: {e}')

    else:
        st.markdown("""<div style="text-align:center;padding:3rem;color:#888;">
            <div style="font-size:3rem;">📤</div>
            <div style="font-size:1.1rem;margin-top:1rem;">Sube tu Excel o conecta con Supabase para comenzar</div>
            <div style="font-size:.85rem;margin-top:.5rem;color:#aaa;">Base bruta exportada de Jotform · instrumento TOP</div>
        </div>""", unsafe_allow_html=True)

    # ── Resultados ────────────────────────────────────────────────────────────
    if 'result' in st.session_state:
        R    = st.session_state['result']
        s    = R['stats']; wide = R['wide']
        fc   = R.get('filtro_centro'); fd = R.get('fecha_desde'); fh = R.get('fecha_hasta')
        filtro_str = (f' · Centro: {fc}' if fc else '') + (f' · {fd} → {fh}' if fd else '')

        st.markdown('---')
        st.markdown(f'<div class="sec">📊 Resultados — {R["periodo"]}{filtro_str}</div>', unsafe_allow_html=True)

        k1,k2,k3,k4,k5,k6 = st.columns(6)
        for col,lbl,val,sub,cls in [
            (k1,'Pacientes únicos',       s['N_total'],   '',                           ''),
            (k2,'Con seguimiento TOP2',   s['N_top2'],    f"{s['pct_top2']}% del total", ''),
            (k3,'Solo TOP1 (pendientes)', s['N_solo1'],   '',                           ''),
            (k4,'Valores corregidos',     s['N_alertas'], '', 'red' if s['N_alertas'] else 'green'),
            (k5,'🔴 Urgentes (90+ días)', s['n_rojo'],    '',                           'red'),
            (k6,'🟠 Próximos (60–89d)',   s['n_naranja'], '',                           'orange'),
        ]:
            with col:
                st.markdown(f'<div class="kpi {cls}"><div class="kpi-lbl">{lbl}</div>'
                            f'<div class="kpi-val">{val}</div>'
                            f'{"<div class=kpi-sub>"+sub+"</div>" if sub else ""}</div>',
                            unsafe_allow_html=True)

        centros = R.get('centros', [])
        if centros and not fc:
            st.markdown('<div class="sec">🏥 Resumen por Centro / Servicio de Tratamiento</div>', unsafe_allow_html=True)
            df_c = pd.DataFrame(centros)
            df_c.columns = ['Centro','Aplicaciones','Pacientes únicos','Con TOP2','Sin TOP2 (pendientes)','Valores corregidos']
            rows_html = ''
            for i, row in df_c.iterrows():
                is_total = str(row.iloc[0]) == 'TOTAL'
                bg = f'background:{NAVY};color:white;font-weight:700;' if is_total else \
                     ('background:#EEF4FB;' if i%2==0 else 'background:white;')
                cells = ''
                for j, val in enumerate(row):
                    align  = 'left' if j==0 else 'center'
                    corr   = (j==5 and not is_total and int(val)>0)
                    color  = 'white' if is_total else (RED if corr else '#333')
                    weight = 'font-weight:700;' if is_total or corr else ''
                    cells += f'<td style="padding:7px 12px;text-align:{align};color:{color};{weight}">{val}</td>'
                rows_html += f'<tr style="{bg}">{cells}</tr>'
            hdrs = ''.join(f'<th style="padding:9px 12px;text-align:{"left" if i==0 else "center"};'
                           f'background:{NAVY};color:white;font-size:.85rem;">{c}</th>'
                           for i,c in enumerate(df_c.columns))
            st.markdown(f'<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;'
                        f'font-family:Calibri,sans-serif;font-size:.9rem;">'
                        f'<thead><tr>{hdrs}</tr></thead><tbody>{rows_html}</tbody></table></div>',
                        unsafe_allow_html=True)

        st.markdown('<div class="sec">📈 Análisis visual</div>', unsafe_allow_html=True)
        gc1,gc2,gc3 = st.columns(3)
        sv=[s['n_verde'],s['n_naranja'],s['n_rojo'],s['N_top2']]
        sl=['<60d','60-89d','90+d','Completados']; sc=[GREEN,ORANGE,RED,MID]
        sv_f=[v for v in sv if v>0]; sl_f=[l for l,v in zip(sl,sv) if v>0]; sc_f=[c for c,v in zip(sc,sv) if v>0]
        sust=s.get('sust_dist',{})
        sd=pd.DataFrame(list(sust.items()),columns=['S','n']).sort_values('n') if sust else pd.DataFrame()
        colors_s=[MID if i%2==0 else ACCENT for i in range(len(sd))]

        with gc1:
            fig,ax=plt.subplots(figsize=(4.5,3.2))
            bars=ax.bar(['Con TOP2','Solo TOP1'],[s['N_top2'],s['N_solo1']],color=[MID,'#CCC'],width=.5)
            for b,v in zip(bars,[s['N_top2'],s['N_solo1']]):
                ax.text(b.get_x()+b.get_width()/2.,b.get_height()+.5,str(v),
                        ha='center',va='bottom',fontsize=11,fontweight='bold',color=NAVY)
            ax.set_title('Estado de seguimiento',fontsize=11,color=NAVY,fontweight='bold',pad=8)
            ax.set_facecolor('#F8FAFD'); fig.patch.set_facecolor('#F8FAFD')
            ax.spines[['top','right','left']].set_visible(False); ax.yaxis.set_visible(False)
            plt.tight_layout(); st.pyplot(fig); plt.close()

        with gc2:
            fig,ax=plt.subplots(figsize=(4.5,3.2))
            if sv_f:
                w,_,at=ax.pie(sv_f,colors=sc_f,autopct='%1.0f%%',startangle=90,
                    wedgeprops={'edgecolor':'white','linewidth':1.5},textprops={'fontsize':9})
                for a in at: a.set_color('white'); a.set_fontweight('bold')
                ax.legend(w,[f'{l} ({v})' for l,v in zip(sl_f,sv_f)],
                    loc='lower center',bbox_to_anchor=(.5,-.3),fontsize=7.5,ncol=2,frameon=False)
            ax.set_title('Semáforo de seguimiento',fontsize=11,color=NAVY,fontweight='bold',pad=8)
            fig.patch.set_facecolor('#F8FAFD'); plt.tight_layout(); st.pyplot(fig); plt.close()

        with gc3:
            fig,ax=plt.subplots(figsize=(4.5,3.2))
            if not sd.empty:
                ax.barh(sd['S'],sd['n'],color=colors_s,height=.6)
                tot=sd['n'].sum()
                for b,v in zip(ax.patches,sd['n']):
                    ax.text(b.get_width()+.3,b.get_y()+b.get_height()/2,
                            f'{v} ({round(v/tot*100,1) if tot else 0}%)',va='center',fontsize=8,color=NAVY)
                ax.spines[['top','right','bottom']].set_visible(False); ax.xaxis.set_visible(False)
            else:
                ax.text(.5,.5,'Sustancia no detectada',ha='center',va='center',
                        transform=ax.transAxes,fontsize=10,color='#888')
            ax.set_title('Sustancia principal (TOP1)',fontsize=11,color=NAVY,fontweight='bold',pad=8)
            ax.set_facecolor('#F8FAFD'); fig.patch.set_facecolor('#F8FAFD')
            plt.tight_layout(); st.pyplot(fig); plt.close()

        pend = wide[wide['Alerta_TOP2'].isin(['🟠 60-89 dias','🔴 90+ dias'])].copy()
        if len(pend):
            st.markdown('<div class="sec">🚨 Pendientes urgentes</div>', unsafe_allow_html=True)
            pend = pend.loc[:,~pend.columns.duplicated()]
            id_c  = wide.columns[0]; cs=[id_c]
            col_c = next((c for c in pend.columns if 'centro' in c.lower() and '_TOP1' in c), None)
            col_f = next((c for c in pend.columns if 'fecha entrevista' in c.lower() and '_TOP1' in c), None)
            if col_c: cs.append(col_c)
            if col_f: cs.append(col_f)
            cs += ['Dias_desde_TOP1','Alerta_TOP2']
            cs  = list(dict.fromkeys(c for c in cs if c in pend.columns))
            tab = pend[cs].copy()
            tab['_o'] = tab['Alerta_TOP2'].apply(lambda x: 0 if '90' in str(x) else 1)
            tab = tab.sort_values(['_o','Dias_desde_TOP1'],ascending=[True,False]).drop(columns='_o')
            st.dataframe(tab.head(30), use_container_width=True, height=280)

        with st.expander('📋 Log de procesamiento'):
            for log in R['logs']: st.text(log)

        st.markdown('---')
        st.markdown('<div class="sec">⬇️ Descargar reportes</div>', unsafe_allow_html=True)

        fname_base = os.path.splitext(st.session_state.get('filename','base'))[0]
        if fc:  fname_base += f'_{fc}'
        if fd:  fname_base += f'_{fd}_{fh}'
        today_str = datetime.now().strftime('%Y-%m-%d')
        outputs   = st.session_state.get('outputs',{})
        sel       = st.session_state.get('seleccion',{})

        d1,d2,d3 = st.columns(3)
        with d1:
            st.markdown('<div class="outcard"><h4>📊 Base Wide completa</h4>'
                        '<p>6 hojas: Wide · Resumen · Alertas · Calidad · Por Centro · Pendientes</p></div>',
                        unsafe_allow_html=True)
            st.download_button('⬇️ Base Wide (.xlsx)',
                data=R['excel_bytes'].getvalue(),
                file_name=f'TOP_Base_Wide_{fname_base}_{today_str}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True, key='dl_wide')

        for key,col,dlkey in [('caract_excel',d2,'dl_ce'),('seg_excel',d3,'dl_se')]:
            o=outputs.get(key,{}); lbl,fmt,desc=LABELS[key]
            with col:
                st.markdown(f'<div class="outcard"><h4>{lbl}</h4><p>{desc}</p></div>', unsafe_allow_html=True)
                if not sel.get(key,False):    st.caption('No seleccionado')
                elif o.get('ok'):             st.download_button(f'⬇️ {fmt}',data=o['buf'].getvalue(),
                                                  file_name=o['fname'],mime=o['mime'],use_container_width=True,key=dlkey)
                else:                         st.warning(f"⚠️ {o.get('error','Error')[:100]}")

        st.markdown('---')
        d4,d5 = st.columns(2)
        for key,col,dlkey in [('pdf_caract',d4,'dl_pc'),('pdf_seg',d5,'dl_ps')]:
            o=outputs.get(key,{}); lbl,fmt,desc=LABELS[key]
            with col:
                st.markdown(f'<div class="outcard"><h4>{lbl}</h4><p>{desc}</p></div>', unsafe_allow_html=True)
                if not sel.get(key,False):    st.caption('No seleccionado')
                elif o.get('ok'):             st.download_button(f'⬇️ {fmt}',data=o['buf'].getvalue(),
                                                  file_name=o['fname'],mime=o['mime'],use_container_width=True,key=dlkey)
                else:                         st.warning(f"⚠️ {o.get('error','Error')[:100]}")

        st.markdown('---')
        d6,d7 = st.columns(2)
        for key,col,dlkey in [('pptx_caract',d6,'dl_ppc'),('pptx_seg',d7,'dl_pps')]:
            o=outputs.get(key,{}); lbl,fmt,desc=LABELS[key]
            with col:
                st.markdown(f'<div class="outcard"><h4>{lbl}</h4><p>{desc}</p></div>', unsafe_allow_html=True)
                if not sel.get(key,False):    st.caption('No seleccionado')
                elif o.get('ok'):             st.download_button(f'⬇️ {fmt}',data=o['buf'].getvalue(),
                                                  file_name=o['fname'],mime=o['mime'],use_container_width=True,key=dlkey)
                else:                         st.warning(f"⚠️ {o.get('error','Error')[:100]}")

        # ── Distribución por centros ──────────────────────────────────────────
        if 'wide_path' in st.session_state and not filtro_centro_val:
            st.markdown('---')
            st.markdown('<div class="sec">📦 Distribución por centros</div>', unsafe_allow_html=True)
            st.markdown(
                '<div style="background:#EEF4FB;border-left:4px solid #2E75B6;'
                'padding:.8rem 1.2rem;border-radius:6px;margin-bottom:1rem;">'
                '<b>¿Qué genera este botón?</b><br>'
                'Un archivo <b>.zip</b> con una carpeta por cada centro. '
                'Cada carpeta incluye la base Wide filtrada + los reportes seleccionados.'
                '</div>', unsafe_allow_html=True
            )
            st.markdown('**Selecciona qué incluir en cada paquete:**')
            dc1, dc2, dc3 = st.columns(3)
            with dc1:
                d_ce = st.checkbox('📋 Excel caracterización', value=True,  key='d_ce')
                d_se = st.checkbox('📋 Excel seguimiento',     value=True,  key='d_se')
            with dc2:
                d_pc = st.checkbox('📄 Word caracterización',  value=True,  key='d_pc')
                d_ps = st.checkbox('📄 Word seguimiento',      value=True,  key='d_ps')
            with dc3:
                d_ppc = st.checkbox('📑 PPT caracterización',  value=False, key='d_ppc')
                d_pps = st.checkbox('📑 PPT seguimiento',      value=False, key='d_pps')

            keys_dist = [k for k,v in {
                'caract_excel':d_ce,'seg_excel':d_se,
                'pdf_caract':d_pc,'pdf_seg':d_ps,
                'pptx_caract':d_ppc,'pptx_seg':d_pps,
            }.items() if v]

            n_centros = len(centros_disponibles)
            st.caption(f'Se generarán **{n_centros} carpetas** — una por cada centro detectado')

            if st.button('📦 Generar paquetes por centro', use_container_width=True, key='btn_dist'):
                wide_path_dist = st.session_state['wide_path']
                status_box = st.empty()
                prog_dist  = st.progress(0, text='Iniciando...')
                def _cb(i, total, centro):
                    pct = i/total if total else 1
                    txt = f'Procesando centro {i+1}/{total}: {centro}' if centro != 'listo' else '✅ ZIP generado'
                    prog_dist.progress(pct, text=txt); status_box.info(txt)
                try:
                    with st.spinner('Generando paquetes — esto puede tomar varios minutos...'):
                        zip_buf = run_paquetes_centros(
                            wide_path_dist, keys_sel=keys_dist,
                            progress_cb=_cb, raw_input_path=st.session_state.get('raw_path')
                        )
                    today_str = datetime.now().strftime('%Y-%m-%d')
                    prog_dist.progress(1.0, text='✅ Listo')
                    status_box.success(f'✅ ZIP generado con {n_centros} carpetas · {len(keys_dist)} reportes por centro')
                    st.download_button(
                        label=f'⬇️ Descargar ZIP ({n_centros} centros)',
                        data=zip_buf.getvalue(),
                        file_name=f'QALAT_Paquetes_Centros_{today_str}.zip',
                        mime='application/zip',
                        use_container_width=True, key='dl_dist'
                    )
                except Exception as e:
                    st.error(f'❌ Error generando paquetes: {e}')



# ──────────────────────────────────────────────────────────────────────────────
# TAB 2: CORRECCIÓN DE REGISTROS
# Carga el formulario HTML original e inyecta credenciales desde Secrets
# ──────────────────────────────────────────────────────────────────────────────
import streamlit.components.v1 as _components
import os as _os

_CORRECCION_HTML = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'correccion_top_peru.html')

with tab_correccion:
    # Solo Perú y UNODC acceden a corrección por ahora
    if rol not in ('Perú', 'UNODC'):
        st.info(f'El formulario de corrección para {flag} {rol} estará disponible próximamente.')
    else:
        # Selector de país para UNODC
        if es_unodc:
            pais_corr = st.selectbox(
                'Corregir registros de:',
                ['Perú'],   # agregar 'Ecuador', 'México' cuando existan sus formularios
                key='corr_pais_sel'
            )
        else:
            pais_corr = rol

        st.markdown(
            f'''<div style="background:#FFF8E1;border-left:4px solid #F9A825;
            padding:.7rem 1.2rem;border-radius:6px;margin-bottom:.5rem;font-size:.85rem;">
            <b>⚠ Módulo de corrección — {flag} {pais_corr}.</b>
            Los cambios se aplican directamente en la base de datos QALAT.
            </div>''', unsafe_allow_html=True
        )

        try:
            sb_url = st.secrets['SUPABASE_URL']
            sb_key = st.secrets['SUPABASE_KEY']

            with open(_CORRECCION_HTML, encoding='utf-8') as _f:
                _html = _f.read()

            # Inyectar credenciales desde Secrets (reemplaza placeholders)
            _html = _html.replace('%%SUPABASE_URL%%', sb_url)
            _html = _html.replace('%%SUPABASE_KEY%%',  sb_key)

            # Renderizar el formulario completo dentro de la pestaña
            _components.html(_html, height=4200, scrolling=True)

        except FileNotFoundError:
            st.error(
                '❌ No se encontró el archivo `correccion_top_peru.html` en la raíz del repositorio. '
                'Asegúrate de que el archivo esté en el repo junto a `app.py`.'
            )
        except KeyError:
            st.error('⚠ Las credenciales de Supabase no están configuradas en Secrets.')
        except Exception as _e:
            st.error(f'❌ Error al cargar el formulario de corrección: {_e}')
