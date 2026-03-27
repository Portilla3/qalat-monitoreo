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
# TAB 2: CORRECCIÓN DE REGISTROS — HTML embebido directamente
# ──────────────────────────────────────────────────────────────────────────────
import streamlit.components.v1 as _components

_CORRECCION_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>QALAT — Corrección TOP · Perú</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root {
    --azul:       #1a3a5c;
    --azul-med:   #2563a8;
    --azul-claro: #dbeafe;
    --verde:      #16a34a;
    --naranja:    #d97706;
    --rojo:       #dc2626;
    --gris-bg:    #e8eef6;
    --gris-borde: #b0c4de;
    --blanco:     #ffffff;
    --texto:      #1e293b;
    --texto-suave:#4a6080;
    --radio:      6px;
    --sombra:     0 2px 8px rgba(26,58,92,.13);
  }
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'IBM Plex Sans', sans-serif;
    background: linear-gradient(160deg, #1a3a5c 0%, #2563a8 40%, #3b82c4 70%, #dbeafe 100%);
    background-attachment: fixed;
    color: var(--texto);
    min-height: 100vh;
    padding-bottom: 60px;
  }

  header {
    background: rgba(10,25,47,0.93);
    backdrop-filter: blur(6px);
    color: #fff;
    padding: 14px 28px;
    display: flex;
    align-items: center;
    gap: 18px;
    position: sticky;
    top: 0;
    z-index: 100;
    box-shadow: 0 2px 12px rgba(0,0,0,.3);
    border-bottom: 2px solid #d97706;
  }
  header .header-texto { display: flex; flex-direction: column; gap: 2px; }
  header .titulo-form { font-size: .95rem; font-weight: 600; color: #fcd34d; letter-spacing: .5px; }
  header .subtitulo   { font-size: .78rem; opacity: .65; color: #cbd5e1; }
  header .pais-badge {
    margin-left: auto;
    background: #dc2626;
    color: #fff;
    font-size: .75rem;
    font-weight: 700;
    padding: 5px 14px;
    border-radius: 20px;
    letter-spacing: 1px;
    text-transform: uppercase;
  }
  header .modo-badge {
    background: #d97706;
    color: #fff;
    font-size: .72rem;
    font-weight: 700;
    padding: 4px 12px;
    border-radius: 20px;
    letter-spacing: 1px;
    text-transform: uppercase;
  }

  .contenedor { max-width: 860px; margin: 28px auto; padding: 0 16px; }

  /* ── BÚSQUEDA ── */
  .busqueda-card {
    background: rgba(255,255,255,.97);
    border-radius: 10px;
    box-shadow: var(--sombra);
    margin-bottom: 20px;
    overflow: hidden;
    border: 2px solid #d97706;
  }
  .busqueda-header {
    background: #d97706;
    color: #fff;
    padding: 12px 20px;
    font-size: .85rem;
    font-weight: 700;
    letter-spacing: 1px;
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .busqueda-body { padding: 20px; }
  .busqueda-grid { display: grid; grid-template-columns: 1fr 1fr auto; gap: 14px; align-items: end; }
  .fecha-busq-grid { display: grid; grid-template-columns: 1fr 1fr 1.3fr; gap: 8px; }

  @media (max-width: 600px) {
    .busqueda-grid { grid-template-columns: 1fr; }
  }

  .campo { display: flex; flex-direction: column; gap: 5px; }
  .campo label {
    font-size: .75rem; font-weight: 700;
    color: var(--texto-suave);
    text-transform: uppercase; letter-spacing: .6px;
  }
  .campo input, .campo select {
    border: 1.5px solid var(--gris-borde);
    border-radius: var(--radio);
    padding: 9px 12px;
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: .9rem;
    color: var(--texto);
    background: var(--blanco);
    width: 100%;
  }
  .campo input:focus, .campo select:focus {
    outline: none;
    border-color: var(--azul-med);
    box-shadow: 0 0 0 3px rgba(37,99,168,.12);
  }
  .campo input.error { border-color: var(--rojo); }

  .btn-buscar {
    background: var(--azul-med);
    color: #fff;
    border: none;
    border-radius: var(--radio);
    padding: 10px 24px;
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: .9rem;
    font-weight: 600;
    cursor: pointer;
    white-space: nowrap;
    transition: background .15s;
    height: 42px;
  }
  .btn-buscar:hover { background: #1a3a5c; }
  .btn-buscar:disabled { background: #94a3b8; cursor: not-allowed; }

  /* ── REGISTRO ENCONTRADO ── */
  .seccion {
    background: rgba(255,255,255,.97);
    border-radius: 10px;
    box-shadow: var(--sombra);
    margin-bottom: 20px;
    overflow: hidden;
    border: 1px solid rgba(37,99,168,.15);
  }
  .seccion-header {
    background: var(--azul);
    color: #fff;
    padding: 12px 20px;
    font-size: .78rem;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .seccion-header .num {
    background: rgba(255,255,255,.22);
    width: 24px; height: 24px;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: .75rem; font-weight: 700;
  }
  .seccion-body { padding: 20px; }

  .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  .span-2 { grid-column: span 2; }
  @media (max-width: 600px) {
    .grid-2 { grid-template-columns: 1fr; }
    .span-2 { grid-column: span 1; }
  }

  /* Banner registro encontrado */
  .registro-encontrado {
    background: #dcfce7;
    border: 1.5px solid #16a34a;
    border-radius: 8px;
    padding: 12px 18px;
    margin-bottom: 18px;
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: .88rem;
    color: #14532d;
    font-weight: 600;
  }

  /* Banner advertencia edición */
  .advertencia-edicion {
    background: #fef3c7;
    border: 1.5px solid #d97706;
    border-radius: 8px;
    padding: 12px 18px;
    margin-bottom: 18px;
    font-size: .82rem;
    color: #78350f;
  }
  .advertencia-edicion strong { color: #92400e; }

  /* ── TABLA SUSTANCIAS ── */
  .tabla-wrap { overflow-x: auto; }
  table.sustancias {
    width: 100%; border-collapse: collapse; font-size: .82rem;
  }
  table.sustancias th {
    background: var(--azul); color: #fff;
    padding: 8px 10px; text-align: center;
    font-weight: 700; font-size: .73rem;
    letter-spacing: .5px; white-space: nowrap;
  }
  table.sustancias th:first-child { text-align: left; min-width: 140px; }
  table.sustancias td {
    padding: 6px 6px;
    border-bottom: 1px solid #e2eaf4;
    vertical-align: middle;
  }
  table.sustancias tr:last-child td { border-bottom: none; }
  table.sustancias tr:hover td { background: #f5f8fc; }
  table.sustancias td:first-child { font-weight: 600; padding-left: 10px; color: var(--texto); }
  table.sustancias td.total-cell {
    background: var(--azul-claro);
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 700; text-align: center;
    color: var(--azul); font-size: .85rem;
  }

  input.num-inp {
    width: 62px; border: 1.5px solid var(--gris-borde);
    border-radius: 4px; padding: 6px 4px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: .88rem; text-align: center;
    color: var(--texto); display: block; margin: 0 auto;
    -moz-appearance: textfield;
  }
  input.num-inp::-webkit-inner-spin-button,
  input.num-inp::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
  input.num-inp:focus {
    outline: none; border-color: var(--azul-med);
    box-shadow: 0 0 0 2px rgba(37,99,168,.14);
  }
  input.num-inp.inp-error { background: #fee2e2; border-color: var(--rojo); }

  input.prom-inp {
    width: 72px; border: 1.5px solid var(--gris-borde);
    border-radius: 4px; padding: 6px 4px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: .85rem; text-align: center;
    color: var(--texto-suave); display: block; margin: 0 auto;
    -moz-appearance: textfield;
  }
  input.prom-inp::-webkit-inner-spin-button,
  input.prom-inp::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }

  td.total-cell { min-width: 52px; }
  .sep { border: none; border-top: 1px solid #e2eaf4; margin: 16px 0; }

  /* ── TRANSGRESIÓN ── */
  .transgresion-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(175px, 1fr));
    gap: 12px; margin-bottom: 16px;
  }
  .toggle-campo { display: flex; flex-direction: column; gap: 5px; }
  .toggle-campo label {
    font-size: .75rem; font-weight: 700;
    color: var(--texto-suave);
    text-transform: uppercase; letter-spacing: .5px;
  }
  .toggle-btn-group { display: flex; border-radius: var(--radio); overflow: hidden; border: 1.5px solid var(--gris-borde); }
  .toggle-btn-group button {
    flex: 1; padding: 8px 0;
    border: none; background: var(--blanco);
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: .85rem; font-weight: 500;
    cursor: pointer; transition: background .15s, color .15s;
    color: var(--texto-suave);
  }
  .toggle-btn-group button:first-child { border-right: 1px solid var(--gris-borde); }
  .toggle-btn-group button.activo-si { background: #dcfce7; color: #16a34a; font-weight: 700; }
  .toggle-btn-group button.activo-no { background: #fee2e2; color: #dc2626; font-weight: 700; }

  table.vif-table {
    width: auto; border-collapse: collapse; font-size: .82rem; margin-top: 4px;
  }
  table.vif-table th {
    background: var(--azul); color: #fff;
    padding: 7px 12px; font-size: .73rem;
    text-align: center; font-weight: 700;
  }
  table.vif-table td { padding: 6px 8px; border-bottom: 1px solid #e2eaf4; }
  table.vif-table td.total-cell {
    background: var(--azul-claro);
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 700; color: var(--azul); text-align: center;
  }

  /* ── ESCALA 0-20 ── */
  .escala-wrap { margin-bottom: 16px; }
  .escala-label {
    font-size: .75rem; font-weight: 700; color: var(--texto-suave);
    text-transform: uppercase; letter-spacing: .6px; margin-bottom: 8px;
  }
  .escala-extremos {
    display: flex; justify-content: space-between;
    font-size: .68rem; font-weight: 600; color: var(--texto-suave);
    text-transform: uppercase; letter-spacing: .4px;
    margin-bottom: 4px; padding: 0 2px;
  }
  .escala-opciones { display: flex; flex-wrap: wrap; gap: 5px; }
  .escala-opciones label {
    display: flex; align-items: center; justify-content: center;
    background: var(--gris-bg);
    border: 1.5px solid var(--gris-borde);
    border-radius: 50%;
    width: 36px; height: 36px;
    cursor: pointer; font-size: .78rem; font-weight: 600;
    transition: all .15s; text-align: center;
    color: var(--texto); flex-shrink: 0;
  }
  .escala-opciones input[type=radio] { display: none; }
  .escala-opciones label:hover { background: #dbeafe; border-color: var(--azul-med); }
  .escala-opciones label:has(input:checked) {
    background: var(--azul-med); border-color: var(--azul-med);
    color: #fff; transform: scale(1.12);
    box-shadow: 0 2px 8px rgba(37,99,168,.35);
  }

  /* Vivienda */
  .vivienda-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 16px; }

  /* ── BOTÓN GUARDAR ── */
  .btn-guardar {
    display: block; width: 100%; padding: 16px;
    background: #d97706;
    color: #fff; border: none; border-radius: 8px;
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 1.05rem; font-weight: 700;
    cursor: pointer; letter-spacing: .3px;
    box-shadow: 0 4px 14px rgba(217,119,6,.35);
    transition: background .15s, box-shadow .15s;
    margin-top: 8px;
  }
  .btn-guardar:hover { background: #b45309; box-shadow: 0 6px 18px rgba(217,119,6,.45); }
  .btn-guardar:disabled { background: #94a3b8; cursor: not-allowed; box-shadow: none; }

  .nota { text-align: center; font-size: .78rem; color: rgba(255,255,255,.7); margin-top: 10px; }

  /* ── TOAST ── */
  #toast {
    display: none; position: fixed; bottom: 28px; left: 50%;
    transform: translateX(-50%);
    background: #16a34a; color: #fff;
    padding: 12px 28px; border-radius: 8px;
    font-size: .9rem; font-weight: 600;
    box-shadow: 0 4px 16px rgba(0,0,0,.25);
    z-index: 999; min-width: 260px; text-align: center;
  }
  #toast.error-toast { background: #dc2626; }
  #toast.warn-toast  { background: #d97706; }

  /* Estado vacío */
  #formulario-correc { display: none; }
  .estado-inicial {
    text-align: center; padding: 3rem;
    color: rgba(255,255,255,.8);
  }
  .estado-inicial .icono { font-size: 3rem; margin-bottom: 1rem; }
  .estado-inicial p { font-size: 1rem; }
  .estado-inicial small { font-size: .85rem; opacity: .7; }
</style>
</head>
<body>

<header>
  <div class="header-texto">
    <div class="titulo-form">✏️ Corrección de Registros TOP</div>
    <div class="subtitulo">Proyecto QALAT · UNODC Chile</div>
    <div class="subtitulo">© Rodrigo Portilla · UNODC</div>
  </div>
  <span class="modo-badge">✏️ Edición</span>
  <span class="pais-badge">🇵🇪 Perú</span>
</header>

<div class="contenedor">

  <!-- ══ BÚSQUEDA ══ -->
  <div class="busqueda-card">
    <div class="busqueda-header">🔍 Buscar registro a corregir</div>
    <div class="busqueda-body">
      <div class="advertencia-edicion">
        <strong>⚠️ Módulo de corrección.</strong> Busca el registro por código de paciente y fecha de entrevista. Los cambios que guardes se aplicarán directamente en la base de datos QALAT.
      </div>
      <div class="busqueda-grid">
        <div class="campo">
          <label>Código de paciente</label>
          <input type="text" id="busq_codigo" placeholder="Ej: ROPE15031985" maxlength="12"
                 style="text-transform:uppercase;font-family:'IBM Plex Mono',monospace;letter-spacing:1px">
        </div>
        <div class="campo">
          <label>Fecha de entrevista</label>
          <div class="fecha-busq-grid">
            <input type="text" id="busq_dia" placeholder="DD" maxlength="2"
                   style="font-family:'IBM Plex Mono',monospace;text-align:center;border:1.5px solid var(--gris-borde);border-radius:var(--radio);padding:9px 6px;font-size:.9rem;width:100%">
            <input type="text" id="busq_mes" placeholder="MM" maxlength="2"
                   style="font-family:'IBM Plex Mono',monospace;text-align:center;border:1.5px solid var(--gris-borde);border-radius:var(--radio);padding:9px 6px;font-size:.9rem;width:100%">
            <input type="text" id="busq_anio" placeholder="AAAA" maxlength="4"
                   style="font-family:'IBM Plex Mono',monospace;text-align:center;border:1.5px solid var(--gris-borde);border-radius:var(--radio);padding:9px 6px;font-size:.9rem;width:100%">
          </div>
          <small style="color:var(--texto-suave);font-size:.7rem;margin-top:3px">Día · Mes · Año — puedes pegar la fecha</small>
        </div>
        <button class="btn-buscar" id="btnBuscar" onclick="buscarRegistro()">🔍 Buscar</button>
      </div>
    </div>
  </div>

  <!-- ══ ESTADO INICIAL ══ -->
  <div class="estado-inicial" id="estadoInicial">
    <div class="icono">🔍</div>
    <p>Ingresa el código de paciente y la fecha de entrevista para buscar el registro</p>
    <small>Solo puedes editar registros de Perú</small>
  </div>

  <!-- ══ FORMULARIO DE CORRECCIÓN ══ -->
  <div id="formulario-correc">

    <div class="registro-encontrado" id="bannerEncontrado">
      ✅ Registro encontrado — edita los campos que necesitas corregir y presiona Guardar cambios
    </div>

    <!-- Sección 0: Identificación -->
    <div class="seccion">
      <div class="seccion-header"><span class="num">0</span> Datos de Identificación</div>
      <div class="seccion-body">
        <div class="grid-2">
          <div class="campo">
            <label>Centro de tratamiento</label>
            <input type="text" id="centro" placeholder="Nombre del centro">
          </div>
          <div class="campo">
            <label>Etapa</label>
            <select id="etapa">
              <option value="">— Seleccione —</option>
              <option value="ingreso">Ingreso (TOP1)</option>
              <option value="seguimiento1">Seguimiento 1</option>
              <option value="seguimiento2">Seguimiento 2</option>
              <option value="egreso">Egreso</option>
            </select>
          </div>
          <div class="campo">
            <label>Código de paciente</label>
            <input type="text" id="codigo_paciente" maxlength="12"
                   style="text-transform:uppercase;font-family:'IBM Plex Mono',monospace;letter-spacing:1px"
                   placeholder="LLLLDDMMAAAA">
          </div>
          <div class="campo">
            <label>Fecha de entrevista</label>
            <input type="date" id="fecha_entrevista">
          </div>
          <div class="campo">
            <label>Fecha de nacimiento</label>
            <input type="date" id="fecha_nacimiento">
          </div>
          <div class="campo">
            <label>Sexo</label>
            <select id="sexo">
              <option value="">— Seleccione —</option>
              <option value="M">Masculino</option>
              <option value="F">Femenino</option>
              <option value="O">Otro</option>
            </select>
          </div>
          <div class="campo span-2">
            <label>Nombre del entrevistador</label>
            <input type="text" id="nombre_entrevistador" placeholder="Nombre completo">
          </div>
        </div>
      </div>
    </div>

    <!-- Sección 1: Sustancias -->
    <div class="seccion">
      <div class="seccion-header"><span class="num">1</span> Sección 1: Uso de Sustancias</div>
      <div class="seccion-body">
        <div class="tabla-wrap">
          <table class="sustancias">
            <thead>
              <tr>
                <th>Sustancia</th>
                <th>Última Semana</th>
                <th>Semana 3</th>
                <th>Semana 2</th>
                <th>Semana 1</th>
                <th>Total</th>
                <th>Promedio</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>Alcohol</td>
                <td><input type="number" min="0" max="7" class="num-inp s4" data-sust="alcohol"></td>
                <td><input type="number" min="0" max="7" class="num-inp s3" data-sust="alcohol"></td>
                <td><input type="number" min="0" max="7" class="num-inp s2" data-sust="alcohol"></td>
                <td><input type="number" min="0" max="7" class="num-inp s1" data-sust="alcohol"></td>
                <td class="total-cell" id="alcohol_total"></td>
                <td><input type="number" step="any" class="prom-inp" id="alcohol_prom" placeholder="—"><br><small style="color:var(--texto-suave);font-size:.68rem;display:block;text-align:center">Tragos/día</small></td>
              </tr>
              <tr>
                <td>Marihuana</td>
                <td><input type="number" min="0" max="7" class="num-inp s4" data-sust="marihuana"></td>
                <td><input type="number" min="0" max="7" class="num-inp s3" data-sust="marihuana"></td>
                <td><input type="number" min="0" max="7" class="num-inp s2" data-sust="marihuana"></td>
                <td><input type="number" min="0" max="7" class="num-inp s1" data-sust="marihuana"></td>
                <td class="total-cell" id="marihuana_total"></td>
                <td><input type="number" step="any" class="prom-inp" id="marihuana_prom" placeholder="—"><br><small style="color:var(--texto-suave);font-size:.68rem;display:block;text-align:center">Pitos/día</small></td>
              </tr>
              <tr>
                <td>Pasta Base</td>
                <td><input type="number" min="0" max="7" class="num-inp s4" data-sust="pastabase"></td>
                <td><input type="number" min="0" max="7" class="num-inp s3" data-sust="pastabase"></td>
                <td><input type="number" min="0" max="7" class="num-inp s2" data-sust="pastabase"></td>
                <td><input type="number" min="0" max="7" class="num-inp s1" data-sust="pastabase"></td>
                <td class="total-cell" id="pastabase_total"></td>
                <td><input type="number" step="any" class="prom-inp" id="pastabase_prom" placeholder="—"><br><small style="color:var(--texto-suave);font-size:.68rem;display:block;text-align:center">Papelillos/día</small></td>
              </tr>
              <tr>
                <td>Cocaína</td>
                <td><input type="number" min="0" max="7" class="num-inp s4" data-sust="cocaina"></td>
                <td><input type="number" min="0" max="7" class="num-inp s3" data-sust="cocaina"></td>
                <td><input type="number" min="0" max="7" class="num-inp s2" data-sust="cocaina"></td>
                <td><input type="number" min="0" max="7" class="num-inp s1" data-sust="cocaina"></td>
                <td class="total-cell" id="cocaina_total"></td>
                <td><input type="number" step="any" class="prom-inp" id="cocaina_prom" placeholder="—"><br><small style="color:var(--texto-suave);font-size:.68rem;display:block;text-align:center">Gramos/día</small></td>
              </tr>
              <tr>
                <td>Sedantes o Tranquilizantes</td>
                <td><input type="number" min="0" max="7" class="num-inp s4" data-sust="sedantes"></td>
                <td><input type="number" min="0" max="7" class="num-inp s3" data-sust="sedantes"></td>
                <td><input type="number" min="0" max="7" class="num-inp s2" data-sust="sedantes"></td>
                <td><input type="number" min="0" max="7" class="num-inp s1" data-sust="sedantes"></td>
                <td class="total-cell" id="sedantes_total"></td>
                <td><input type="number" step="any" class="prom-inp" id="sedantes_prom" placeholder="—"><br><small style="color:var(--texto-suave);font-size:.68rem;display:block;text-align:center">Comprimidos/día</small></td>
              </tr>
              <tr class="otra-sust-fila">
                <td colspan="7" style="padding:8px 10px">
                  <input type="text" id="otra_sust_nombre" placeholder="Nombre de otra sustancia (si aplica)"
                         style="width:100%;border:1.5px solid var(--gris-borde);border-radius:4px;padding:6px 10px;font-family:inherit;font-size:.85rem;">
                </td>
              </tr>
              <tr>
                <td><em style="color:var(--texto-suave)">Otra sustancia problema</em></td>
                <td><input type="number" min="0" max="7" class="num-inp s4" data-sust="otra_sust"></td>
                <td><input type="number" min="0" max="7" class="num-inp s3" data-sust="otra_sust"></td>
                <td><input type="number" min="0" max="7" class="num-inp s2" data-sust="otra_sust"></td>
                <td><input type="number" min="0" max="7" class="num-inp s1" data-sust="otra_sust"></td>
                <td class="total-cell" id="otra_sust_total"></td>
                <td><input type="number" step="any" class="prom-inp" id="otra_sust_prom" placeholder="—"><br><small style="color:var(--texto-suave);font-size:.68rem;display:block;text-align:center">Medida/día</small></td>
              </tr>
            </tbody>
          </table>
        </div>
        <hr class="sep">
        <div class="campo">
          <label>Sustancia principal</label>
          <select id="sustancia_principal">
            <option value="">— Seleccione —</option>
            <option value="Alcohol">Alcohol</option>
            <option value="Marihuana">Marihuana</option>
            <option value="Pasta Base">Pasta Base</option>
            <option value="Cocaína">Cocaína</option>
            <option value="Sedantes">Sedantes</option>
            <option value="Otra">Otra sustancia</option>
          </select>
        </div>
      </div>
    </div>

    <!-- Sección 2: Transgresión -->
    <div class="seccion">
      <div class="seccion-header"><span class="num">2</span> Sección 2: Transgresión a la Norma Social</div>
      <div class="seccion-body">
        <div class="transgresion-grid">
          <div class="toggle-campo">
            <label>Hurto</label>
            <div class="toggle-btn-group" id="tg-hurto">
              <button type="button" onclick="setToggle('hurto','S',this)">Sí</button>
              <button type="button" onclick="setToggle('hurto','N',this)">No</button>
            </div>
          </div>
          <div class="toggle-campo">
            <label>Robo</label>
            <div class="toggle-btn-group" id="tg-robo">
              <button type="button" onclick="setToggle('robo','S',this)">Sí</button>
              <button type="button" onclick="setToggle('robo','N',this)">No</button>
            </div>
          </div>
          <div class="toggle-campo">
            <label>Venta de droga</label>
            <div class="toggle-btn-group" id="tg-venta_droga">
              <button type="button" onclick="setToggle('venta_droga','S',this)">Sí</button>
              <button type="button" onclick="setToggle('venta_droga','N',this)">No</button>
            </div>
          </div>
          <div class="toggle-campo">
            <label>Riña / Pelea</label>
            <div class="toggle-btn-group" id="tg-rina_pelea">
              <button type="button" onclick="setToggle('rina_pelea','S',this)">Sí</button>
              <button type="button" onclick="setToggle('rina_pelea','N',this)">No</button>
            </div>
          </div>
          <div class="toggle-campo">
            <label>Otra acción</label>
            <div class="toggle-btn-group" id="tg-otra_accion">
              <button type="button" onclick="setToggle('otra_accion','S',this);mostrarOtraAccion(true)">Sí</button>
              <button type="button" onclick="setToggle('otra_accion','N',this);mostrarOtraAccion(false)">No</button>
            </div>
          </div>
        </div>
        <div id="otra-accion-desc-wrap" class="campo" style="display:none;margin-bottom:16px;">
          <label>Descripción de otra acción</label>
          <input type="text" id="otra_accion_desc" placeholder="Describa la acción">
        </div>

        <hr class="sep">
        <p style="font-size:.75rem;font-weight:700;color:var(--texto-suave);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px;">
          Violencia intrafamiliar (VIF) — días por semana (0–7)
        </p>
        <div class="tabla-wrap">
          <table class="vif-table">
            <thead>
              <tr><th>Última Semana</th><th>Semana 3</th><th>Semana 2</th><th>Semana 1</th><th>Total</th></tr>
            </thead>
            <tbody>
              <tr>
                <td><input type="number" id="vif_s4" min="0" max="7" class="num-inp" oninput="calcVif()"></td>
                <td><input type="number" id="vif_s3" min="0" max="7" class="num-inp" oninput="calcVif()"></td>
                <td><input type="number" id="vif_s2" min="0" max="7" class="num-inp" oninput="calcVif()"></td>
                <td><input type="number" id="vif_s1" min="0" max="7" class="num-inp" oninput="calcVif()"></td>
                <td class="total-cell" id="vif_total"></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- Sección 3: Salud -->
    <div class="seccion">
      <div class="seccion-header"><span class="num">3</span> Sección 3: Salud y Funcionamiento Social</div>
      <div class="seccion-body">

        <!-- Salud psicológica -->
        <div class="escala-wrap">
          <div class="escala-label">3a. Salud Psicológica</div>
          <div class="escala-extremos"><span>0 · Pésima</span><span>20 · Excelente</span></div>
          <div class="escala-opciones">
            <label><input type="radio" name="salud_psicologica" value="0">0</label>
            <label><input type="radio" name="salud_psicologica" value="1">1</label>
            <label><input type="radio" name="salud_psicologica" value="2">2</label>
            <label><input type="radio" name="salud_psicologica" value="3">3</label>
            <label><input type="radio" name="salud_psicologica" value="4">4</label>
            <label><input type="radio" name="salud_psicologica" value="5">5</label>
            <label><input type="radio" name="salud_psicologica" value="6">6</label>
            <label><input type="radio" name="salud_psicologica" value="7">7</label>
            <label><input type="radio" name="salud_psicologica" value="8">8</label>
            <label><input type="radio" name="salud_psicologica" value="9">9</label>
            <label><input type="radio" name="salud_psicologica" value="10">10</label>
            <label><input type="radio" name="salud_psicologica" value="11">11</label>
            <label><input type="radio" name="salud_psicologica" value="12">12</label>
            <label><input type="radio" name="salud_psicologica" value="13">13</label>
            <label><input type="radio" name="salud_psicologica" value="14">14</label>
            <label><input type="radio" name="salud_psicologica" value="15">15</label>
            <label><input type="radio" name="salud_psicologica" value="16">16</label>
            <label><input type="radio" name="salud_psicologica" value="17">17</label>
            <label><input type="radio" name="salud_psicologica" value="18">18</label>
            <label><input type="radio" name="salud_psicologica" value="19">19</label>
            <label><input type="radio" name="salud_psicologica" value="20">20</label>
          </div>
        </div>

        <hr class="sep">
        <!-- Trabajo -->
        <p style="font-size:.75rem;font-weight:700;color:var(--texto-suave);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px;">3b. Trabajo remunerado — días por semana (0–7)</p>
        <div class="tabla-wrap" style="margin-bottom:16px;">
          <table class="vif-table">
            <thead><tr><th>Última Semana</th><th>Semana 3</th><th>Semana 2</th><th>Semana 1</th><th>Total</th></tr></thead>
            <tbody><tr>
              <td><input type="number" id="dias_trabajo_s4" min="0" max="7" class="num-inp" oninput="calcSimple('dias_trabajo')"></td>
              <td><input type="number" id="dias_trabajo_s3" min="0" max="7" class="num-inp" oninput="calcSimple('dias_trabajo')"></td>
              <td><input type="number" id="dias_trabajo_s2" min="0" max="7" class="num-inp" oninput="calcSimple('dias_trabajo')"></td>
              <td><input type="number" id="dias_trabajo_s1" min="0" max="7" class="num-inp" oninput="calcSimple('dias_trabajo')"></td>
              <td class="total-cell" id="dias_trabajo_total"></td>
            </tr></tbody>
          </table>
        </div>

        <!-- Educación -->
        <p style="font-size:.75rem;font-weight:700;color:var(--texto-suave);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px;">3c. Educación / Formación — días por semana (0–7)</p>
        <div class="tabla-wrap" style="margin-bottom:16px;">
          <table class="vif-table">
            <thead><tr><th>Última Semana</th><th>Semana 3</th><th>Semana 2</th><th>Semana 1</th><th>Total</th></tr></thead>
            <tbody><tr>
              <td><input type="number" id="dias_educacion_s4" min="0" max="7" class="num-inp" oninput="calcSimple('dias_educacion')"></td>
              <td><input type="number" id="dias_educacion_s3" min="0" max="7" class="num-inp" oninput="calcSimple('dias_educacion')"></td>
              <td><input type="number" id="dias_educacion_s2" min="0" max="7" class="num-inp" oninput="calcSimple('dias_educacion')"></td>
              <td><input type="number" id="dias_educacion_s1" min="0" max="7" class="num-inp" oninput="calcSimple('dias_educacion')"></td>
              <td class="total-cell" id="dias_educacion_total"></td>
            </tr></tbody>
          </table>
        </div>

        <hr class="sep">
        <!-- Salud física -->
        <div class="escala-wrap">
          <div class="escala-label">3e. Salud Física</div>
          <div class="escala-extremos"><span>0 · Pésima</span><span>20 · Excelente</span></div>
          <div class="escala-opciones">
            <label><input type="radio" name="salud_fisica" value="0">0</label>
            <label><input type="radio" name="salud_fisica" value="1">1</label>
            <label><input type="radio" name="salud_fisica" value="2">2</label>
            <label><input type="radio" name="salud_fisica" value="3">3</label>
            <label><input type="radio" name="salud_fisica" value="4">4</label>
            <label><input type="radio" name="salud_fisica" value="5">5</label>
            <label><input type="radio" name="salud_fisica" value="6">6</label>
            <label><input type="radio" name="salud_fisica" value="7">7</label>
            <label><input type="radio" name="salud_fisica" value="8">8</label>
            <label><input type="radio" name="salud_fisica" value="9">9</label>
            <label><input type="radio" name="salud_fisica" value="10">10</label>
            <label><input type="radio" name="salud_fisica" value="11">11</label>
            <label><input type="radio" name="salud_fisica" value="12">12</label>
            <label><input type="radio" name="salud_fisica" value="13">13</label>
            <label><input type="radio" name="salud_fisica" value="14">14</label>
            <label><input type="radio" name="salud_fisica" value="15">15</label>
            <label><input type="radio" name="salud_fisica" value="16">16</label>
            <label><input type="radio" name="salud_fisica" value="17">17</label>
            <label><input type="radio" name="salud_fisica" value="18">18</label>
            <label><input type="radio" name="salud_fisica" value="19">19</label>
            <label><input type="radio" name="salud_fisica" value="20">20</label>
          </div>
        </div>

        <hr class="sep">
        <!-- Vivienda -->
        <p style="font-size:.75rem;font-weight:700;color:var(--texto-suave);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px;">3f. Vivienda</p>
        <div class="vivienda-grid">
          <div class="toggle-campo">
            <label>Vivienda estable</label>
            <div class="toggle-btn-group" id="tg-vivienda_estable">
              <button type="button" onclick="setToggle('vivienda_estable','S',this)">Sí</button>
              <button type="button" onclick="setToggle('vivienda_estable','N',this)">No</button>
            </div>
          </div>
          <div class="toggle-campo">
            <label>Vivienda básica</label>
            <div class="toggle-btn-group" id="tg-vivienda_basica">
              <button type="button" onclick="setToggle('vivienda_basica','S',this)">Sí</button>
              <button type="button" onclick="setToggle('vivienda_basica','N',this)">No</button>
            </div>
          </div>
        </div>

        <hr class="sep">
        <!-- Calidad de vida -->
        <div class="escala-wrap">
          <div class="escala-label">3g. Calidad de Vida General</div>
          <div class="escala-extremos"><span>0 · Pésima</span><span>20 · Excelente</span></div>
          <div class="escala-opciones">
            <label><input type="radio" name="calidad_vida" value="0">0</label>
            <label><input type="radio" name="calidad_vida" value="1">1</label>
            <label><input type="radio" name="calidad_vida" value="2">2</label>
            <label><input type="radio" name="calidad_vida" value="3">3</label>
            <label><input type="radio" name="calidad_vida" value="4">4</label>
            <label><input type="radio" name="calidad_vida" value="5">5</label>
            <label><input type="radio" name="calidad_vida" value="6">6</label>
            <label><input type="radio" name="calidad_vida" value="7">7</label>
            <label><input type="radio" name="calidad_vida" value="8">8</label>
            <label><input type="radio" name="calidad_vida" value="9">9</label>
            <label><input type="radio" name="calidad_vida" value="10">10</label>
            <label><input type="radio" name="calidad_vida" value="11">11</label>
            <label><input type="radio" name="calidad_vida" value="12">12</label>
            <label><input type="radio" name="calidad_vida" value="13">13</label>
            <label><input type="radio" name="calidad_vida" value="14">14</label>
            <label><input type="radio" name="calidad_vida" value="15">15</label>
            <label><input type="radio" name="calidad_vida" value="16">16</label>
            <label><input type="radio" name="calidad_vida" value="17">17</label>
            <label><input type="radio" name="calidad_vida" value="18">18</label>
            <label><input type="radio" name="calidad_vida" value="19">19</label>
            <label><input type="radio" name="calidad_vida" value="20">20</label>
          </div>
        </div>

      </div>
    </div>

    <!-- Botón guardar -->
    <button type="button" class="btn-guardar" id="btnGuardar" onclick="guardarCambios()">
      ✓ &nbsp;Guardar cambios
    </button>
    <p class="nota">Los cambios se aplican directamente en la base de datos QALAT · UNODC Chile</p>

  </div><!-- fin formulario-correc -->
</div>

<div id="toast"></div>

<script>
const SUPABASE_URL = '%%SUPABASE_URL%%';
const SUPABASE_KEY = '%%SUPABASE_KEY%%';

// ID del registro encontrado (para el UPDATE)
let registroId = null;

// Toggles
const toggleValues = {
  hurto:null, robo:null, venta_droga:null,
  rina_pelea:null, otra_accion:null,
  vivienda_estable:null, vivienda_basica:null
};

function setToggle(campo, valor, btn) {
  toggleValues[campo] = valor;
  const group = btn.closest('.toggle-btn-group');
  group.querySelectorAll('button').forEach(b => b.classList.remove('activo-si','activo-no'));
  btn.classList.add(valor==='S' ? 'activo-si' : 'activo-no');
}
function mostrarOtraAccion(show) {
  document.getElementById('otra-accion-desc-wrap').style.display = show ? 'flex' : 'none';
}

// Toast
function mostrarToast(msg, tipo='ok') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = tipo === 'error' ? 'error-toast' : tipo === 'warn' ? 'warn-toast' : '';
  t.style.display = 'block';
  setTimeout(() => { t.style.display = 'none'; }, 5000);
}

// Cálculo totales sustancias
function calcAct(prefix) {
  const sems = ['s4','s3','s2','s1'];
  let total = null;
  sems.forEach(s => {
    const el = document.querySelector(`input.${s}[data-sust="${prefix}"]`);
    if (el && el.value !== '' && !el.classList.contains('inp-error')) {
      const v = parseInt(el.value);
      if (!isNaN(v)) total = (total === null ? 0 : total) + v;
    }
  });
  const tcell = document.getElementById(`${prefix}_total`);
  if (tcell) tcell.textContent = total !== null ? total : '';
}

function calcVif() {
  const ids = ['vif_s4','vif_s3','vif_s2','vif_s1'];
  let total = null;
  ids.forEach(id => {
    const el = document.getElementById(id);
    if (el && el.value !== '') {
      const v = parseInt(el.value);
      if (!isNaN(v)) total = (total === null ? 0 : total) + v;
    }
  });
  const tc = document.getElementById('vif_total');
  if (tc) tc.textContent = total !== null ? total : '';
}

function calcSimple(prefix) {
  const ids = [`${prefix}_s4`,`${prefix}_s3`,`${prefix}_s2`,`${prefix}_s1`];
  let total = null;
  ids.forEach(id => {
    const el = document.getElementById(id);
    if (el && el.value !== '') {
      const v = parseInt(el.value);
      if (!isNaN(v)) total = (total === null ? 0 : total) + v;
    }
  });
  const tc = document.getElementById(`${prefix}_total`);
  if (tc) tc.textContent = total !== null ? total : '';
}

// Inicializar listeners de sustancias
['alcohol','marihuana','pastabase','cocaina','sedantes','otra_sust'].forEach(prefix => {
  ['s4','s3','s2','s1'].forEach(s => {
    const el = document.querySelector(`input.${s}[data-sust="${prefix}"]`);
    if (!el) return;
    el.addEventListener('input', () => calcAct(prefix));
    el.addEventListener('blur', () => calcAct(prefix));
  });
});

// ── BÚSQUEDA ──────────────────────────────────────────────
async function buscarRegistro() {
  const codigo = document.getElementById('busq_codigo').value.trim().toUpperCase();
  const dia    = document.getElementById('busq_dia').value.trim().padStart(2,'0');
  const mes    = document.getElementById('busq_mes').value.trim().padStart(2,'0');
  const anio   = document.getElementById('busq_anio').value.trim();

  if (!codigo) {
    mostrarToast('⚠ Ingresa el código de paciente', 'error'); return;
  }
  if (!dia || !mes || !anio || anio.length !== 4) {
    mostrarToast('⚠ Ingresa la fecha completa (DD · MM · AAAA)', 'error'); return;
  }

  const fecha = `${anio}-${mes}-${dia}`;  // formato Supabase: YYYY-MM-DD

  const btn = document.getElementById('btnBuscar');
  btn.disabled = true;
  btn.textContent = 'Buscando…';

  try {
    const url = `${SUPABASE_URL}/rest/v1/top_registros?codigo_paciente=eq.${encodeURIComponent(codigo)}&fecha_entrevista=eq.${fecha}&pais=eq.Per%C3%BA&select=*`;
    const resp = await fetch(url, {
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Content-Type': 'application/json'
      }
    });

    const data = await resp.json();

    if (!data || data.length === 0) {
      mostrarToast(`⚠ No se encontró registro: ${codigo} · ${dia}/${mes}/${anio}`, 'warn');
      document.getElementById('formulario-correc').style.display = 'none';
      document.getElementById('estadoInicial').style.display = 'block';
    } else {
      const reg = data[0];
      registroId = reg.id;
      poblarFormulario(reg);
      document.getElementById('estadoInicial').style.display = 'none';
      document.getElementById('formulario-correc').style.display = 'block';
      document.getElementById('bannerEncontrado').textContent =
        `✅ Registro encontrado — ID ${registroId} · ${codigo} · ${dia}/${mes}/${anio} — edita lo que necesitas y guarda`;
      mostrarToast('✓ Registro cargado correctamente');
    }
  } catch(e) {
    mostrarToast('Error de conexión al buscar: ' + e.message, 'error');
  }

  btn.disabled = false;
  btn.textContent = '🔍 Buscar';
}

// ── POBLAR FORMULARIO CON DATOS DEL REGISTRO ──────────────
function poblarFormulario(r) {
  const set = (id, val) => {
    const el = document.getElementById(id);
    if (el && val !== null && val !== undefined) el.value = val;
  };
  const setNum = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.value = (val !== null && val !== undefined) ? val : '';
  };
  const setTotal = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = (val !== null && val !== undefined) ? val : '';
  };
  const setToggleBtn = (campo, val) => {
    if (!val) return;
    toggleValues[campo] = val;
    const group = document.getElementById(`tg-${campo}`);
    if (!group) return;
    group.querySelectorAll('button').forEach(b => b.classList.remove('activo-si','activo-no'));
    const btns = group.querySelectorAll('button');
    if (val === 'S' && btns[0]) btns[0].classList.add('activo-si');
    if (val === 'N' && btns[1]) btns[1].classList.add('activo-no');
  };
  const setRadio = (name, val) => {
    if (val === null || val === undefined) return;
    const radio = document.querySelector(`input[name="${name}"][value="${val}"]`);
    if (radio) radio.checked = true;
  };
  const setSustSem = (sust, sem, val) => {
    const el = document.querySelector(`input.${sem}[data-sust="${sust}"]`);
    if (el) el.value = (val !== null && val !== undefined) ? val : '';
  };

  // Identificación
  set('centro', r.centro);
  set('etapa', r.etapa);
  set('codigo_paciente', r.codigo_paciente);
  set('fecha_entrevista', r.fecha_entrevista);
  set('fecha_nacimiento', r.fecha_nacimiento);
  set('sexo', r.sexo);
  set('nombre_entrevistador', r.nombre_entrevistador);

  // Sustancias
  ['alcohol','marihuana','pastabase','cocaina','sedantes','otra_sust'].forEach(sust => {
    setSustSem(sust,'s4', r[`${sust}_s4`]);
    setSustSem(sust,'s3', r[`${sust}_s3`]);
    setSustSem(sust,'s2', r[`${sust}_s2`]);
    setSustSem(sust,'s1', r[`${sust}_s1`]);
    setTotal(`${sust}_total`, r[`${sust}_total`]);
    setNum(`${sust}_prom`, r[`${sust}_prom`]);
  });
  set('otra_sust_nombre', r.otra_sust_nombre);
  set('sustancia_principal', r.sustancia_principal);

  // Transgresión
  setToggleBtn('hurto', r.hurto);
  setToggleBtn('robo', r.robo);
  setToggleBtn('venta_droga', r.venta_droga);
  setToggleBtn('rina_pelea', r.rina_pelea);
  setToggleBtn('otra_accion', r.otra_accion);
  if (r.otra_accion === 'S') mostrarOtraAccion(true);
  set('otra_accion_desc', r.otra_accion_desc);

  // VIF
  setNum('vif_s4', r.vif_s4); setNum('vif_s3', r.vif_s3);
  setNum('vif_s2', r.vif_s2); setNum('vif_s1', r.vif_s1);
  setTotal('vif_total', r.vif_total);

  // Salud y funcionamiento
  setRadio('salud_psicologica', r.salud_psicologica);
  setNum('dias_trabajo_s4', r.dias_trabajo_s4); setNum('dias_trabajo_s3', r.dias_trabajo_s3);
  setNum('dias_trabajo_s2', r.dias_trabajo_s2); setNum('dias_trabajo_s1', r.dias_trabajo_s1);
  setTotal('dias_trabajo_total', r.dias_trabajo_total);
  setNum('dias_educacion_s4', r.dias_educacion_s4); setNum('dias_educacion_s3', r.dias_educacion_s3);
  setNum('dias_educacion_s2', r.dias_educacion_s2); setNum('dias_educacion_s1', r.dias_educacion_s1);
  setTotal('dias_educacion_total', r.dias_educacion_total);
  setRadio('salud_fisica', r.salud_fisica);
  setToggleBtn('vivienda_estable', r.vivienda_estable);
  setToggleBtn('vivienda_basica', r.vivienda_basica);
  setRadio('calidad_vida', r.calidad_vida);
}

// ── GUARDAR CAMBIOS (UPDATE) ──────────────────────────────
async function guardarCambios() {
  if (!registroId) {
    mostrarToast('⚠ Primero busca un registro', 'warn');
    return;
  }

  const btn = document.getElementById('btnGuardar');
  btn.disabled = true;
  btn.textContent = 'Guardando…';

  const n = id => { const v = parseFloat(document.getElementById(id)?.value); return isNaN(v) ? null : v; };
  const s = id => document.getElementById(id)?.value || null;
  const sv = (sust, sem) => {
    const el = document.querySelector(`input.${sem}[data-sust="${sust}"]`);
    return (el && el.value !== '') ? parseInt(el.value) : null;
  };
  const tv = id => { const el = document.getElementById(id); return el ? parseInt(el.textContent) || null : null; };
  const pv = id => { const el = document.getElementById(id); return el && el.value !== '' ? parseFloat(el.value) : null; };

  const payload = {
    centro: s('centro'), etapa: s('etapa'),
    codigo_paciente: s('codigo_paciente'),
    fecha_entrevista: s('fecha_entrevista'),
    fecha_nacimiento: s('fecha_nacimiento') || null,
    sexo: s('sexo') || null,
    nombre_entrevistador: s('nombre_entrevistador') || null,

    alcohol_s4: sv('alcohol','s4'), alcohol_s3: sv('alcohol','s3'),
    alcohol_s2: sv('alcohol','s2'), alcohol_s1: sv('alcohol','s1'),
    alcohol_total: tv('alcohol_total'), alcohol_prom: pv('alcohol_prom'),

    marihuana_s4: sv('marihuana','s4'), marihuana_s3: sv('marihuana','s3'),
    marihuana_s2: sv('marihuana','s2'), marihuana_s1: sv('marihuana','s1'),
    marihuana_total: tv('marihuana_total'), marihuana_prom: pv('marihuana_prom'),

    pastabase_s4: sv('pastabase','s4'), pastabase_s3: sv('pastabase','s3'),
    pastabase_s2: sv('pastabase','s2'), pastabase_s1: sv('pastabase','s1'),
    pastabase_total: tv('pastabase_total'), pastabase_prom: pv('pastabase_prom'),

    cocaina_s4: sv('cocaina','s4'), cocaina_s3: sv('cocaina','s3'),
    cocaina_s2: sv('cocaina','s2'), cocaina_s1: sv('cocaina','s1'),
    cocaina_total: tv('cocaina_total'), cocaina_prom: pv('cocaina_prom'),

    sedantes_s4: sv('sedantes','s4'), sedantes_s3: sv('sedantes','s3'),
    sedantes_s2: sv('sedantes','s2'), sedantes_s1: sv('sedantes','s1'),
    sedantes_total: tv('sedantes_total'), sedantes_prom: pv('sedantes_prom'),

    otra_sust_nombre: s('otra_sust_nombre') || null,
    otra_sust_s4: sv('otra_sust','s4'), otra_sust_s3: sv('otra_sust','s3'),
    otra_sust_s2: sv('otra_sust','s2'), otra_sust_s1: sv('otra_sust','s1'),
    otra_sust_total: tv('otra_sust_total'), otra_sust_prom: pv('otra_sust_prom'),

    sustancia_principal: s('sustancia_principal'),

    hurto: toggleValues.hurto, robo: toggleValues.robo,
    venta_droga: toggleValues.venta_droga, rina_pelea: toggleValues.rina_pelea,
    otra_accion: toggleValues.otra_accion,
    otra_accion_desc: s('otra_accion_desc') || null,
    vif_s4: n('vif_s4'), vif_s3: n('vif_s3'), vif_s2: n('vif_s2'), vif_s1: n('vif_s1'),
    vif_total: tv('vif_total'),

    salud_psicologica: parseInt(document.querySelector('input[name="salud_psicologica"]:checked')?.value ?? 'NaN') || null,
    dias_trabajo_s4: n('dias_trabajo_s4'), dias_trabajo_s3: n('dias_trabajo_s3'),
    dias_trabajo_s2: n('dias_trabajo_s2'), dias_trabajo_s1: n('dias_trabajo_s1'),
    dias_trabajo_total: tv('dias_trabajo_total'),
    dias_educacion_s4: n('dias_educacion_s4'), dias_educacion_s3: n('dias_educacion_s3'),
    dias_educacion_s2: n('dias_educacion_s2'), dias_educacion_s1: n('dias_educacion_s1'),
    dias_educacion_total: tv('dias_educacion_total'),
    salud_fisica: parseInt(document.querySelector('input[name="salud_fisica"]:checked')?.value ?? 'NaN') || null,
    vivienda_estable: toggleValues.vivienda_estable,
    vivienda_basica: toggleValues.vivienda_basica,
    calidad_vida: parseInt(document.querySelector('input[name="calidad_vida"]:checked')?.value ?? 'NaN') || null
  };

  try {
    const resp = await fetch(`${SUPABASE_URL}/rest/v1/top_registros?id=eq.${registroId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Prefer': 'return=minimal'
      },
      body: JSON.stringify(payload)
    });

    if (resp.ok || resp.status === 204) {
      mostrarToast('✓ Cambios guardados correctamente en la base de datos');
    } else {
      const err = await resp.json();
      mostrarToast('Error al guardar: ' + (err.message || resp.status), 'error');
    }
  } catch(e) {
    mostrarToast('Error de conexión. Verifique su internet.', 'error');
  }

  btn.disabled = false;
  btn.textContent = '✓ Guardar cambios';
}

// Enter en búsqueda
['busq_codigo','busq_dia','busq_mes','busq_anio'].forEach(id => {
  document.getElementById(id).addEventListener('keydown', e => {
    if (e.key === 'Enter') buscarRegistro();
  });
});
document.getElementById('busq_codigo').addEventListener('input', function() {
  const pos = this.selectionStart;
  this.value = this.value.toUpperCase();
  this.setSelectionRange(pos, pos);
});
</script>
</body>
</html>
"""

with tab_correccion:
    if rol not in ('Perú', 'UNODC'):
        st.info(f'El formulario de corrección para {flag} {rol} estará disponible próximamente.')
    else:
        if es_unodc:
            st.selectbox('Corregir registros de:', ['Perú'], key='corr_pais_sel')

        st.markdown(
            f'''<div style="background:#FFF8E1;border-left:4px solid #F9A825;
            padding:.7rem 1.2rem;border-radius:6px;margin-bottom:.5rem;font-size:.85rem;">
            <b>⚠ Módulo de corrección — 🇵🇪 Perú.</b>
            Los cambios se aplican directamente en la base de datos QALAT.
            </div>''', unsafe_allow_html=True
        )

        try:
            _html = _CORRECCION_HTML_TEMPLATE
            _html = _html.replace('%%SUPABASE_URL%%', st.secrets['SUPABASE_URL'])
            _html = _html.replace('%%SUPABASE_KEY%%',  st.secrets['SUPABASE_KEY'])
            _components.html(_html, height=4500, scrolling=True)
        except KeyError:
            st.error('⚠ Las credenciales de Supabase no están configuradas en Secrets.')
        except Exception as _e:
            st.error(f'❌ Error al cargar el formulario de corrección: {_e}')
