"""
app.py — QALAT · Sistema de Monitoreo de Resultados de Tratamiento
v3.0 — filtro período + filtro centro
"""
import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import tempfile, os, sys
from io import BytesIO
from datetime import datetime, date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pipeline.wide_top import procesar_wide
from pipeline.runner   import run_script

NAVY='#1F3864'; MID='#2E75B6'; ACCENT='#00B0F0'
ORANGE='#C8590A'; RED='#C00000'; GREEN='#538135'; WHITE='#FFFFFF'

st.set_page_config(page_title='QALAT · Monitoreo', page_icon='📊',
                   layout='wide', initial_sidebar_state='collapsed')
st.markdown(f"""<style>
html,body,[class*="css"]{{font-family:'Calibri',sans-serif;}}
.main{{background:#F8FAFD;}}
.qalat-hdr{{background:{NAVY};color:white;padding:1.2rem 2rem;border-radius:8px;margin-bottom:1.5rem;}}
.qalat-hdr h1{{color:white;font-size:1.6rem;margin:0;}}
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
div.stButton>button{{background:{NAVY};color:white;border:none;
    padding:.6rem 2rem;border-radius:6px;font-size:1rem;font-weight:600;width:100%;}}
div.stButton>button:hover{{background:{MID};}}
#MainMenu,footer,header{{visibility:hidden;}}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="qalat-hdr">
  <h1>📊 QALAT · Sistema de Monitoreo de Resultados de Tratamiento</h1>
  <p>Procesamiento automático TOP · Sube tu Excel, aplica filtros y descarga todos los reportes</p>
</div>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown('### 📋 Pasos')
    st.markdown('1. Sube tu Excel bruto\n2. Aplica filtros (opcional)\n3. Elige reportes\n4. Clic en **Procesar**\n5. Descarga')
    st.markdown('---')
    st.caption(f'QALAT v3.0 · {datetime.now().strftime("%d/%m/%Y")}')

LABELS = {
    'caract_excel':('📋 Tablas caracterización',   'Excel',      '11 tablas al ingreso: sexo, edad, sustancias, transgresión'),
    'seg_excel':   ('📋 Tablas seguimiento',        'Excel',      'Comparativo TOP1 vs TOP2'),
    'pdf_caract':  ('📄 Word caracterización',       'Word',       '4 secciones · gráficos · tablas'),
    'pdf_seg':     ('📄 Word seguimiento',           'Word',       'Comparativo ingreso vs seguimiento'),
    'pptx_caract': ('📑 PPT caracterización',       'PowerPoint', '6 slides · perfil al ingreso'),
    'pptx_seg':    ('📑 PPT seguimiento',           'PowerPoint', '6 slides · ingreso vs seguimiento'),
}

# ── Sección de carga ───────────────────────────────────────────────────────────
st.markdown('<div class="sec">📁 Cargar base de datos</div>', unsafe_allow_html=True)
uploaded = st.file_uploader('Arrastra tu Excel aquí o haz clic para buscar',
                             type=['xlsx','xls'],
                             help='Archivo bruto exportado de Jotform — instrumento TOP')

# ── Filtros (solo visibles si hay archivo) ─────────────────────────────────────
filtro_centro_val = None
fecha_desde_val   = None
fecha_hasta_val   = None
centros_disponibles = []

if uploaded:
    # Leer solo las columnas de centro y fecha para poblar los filtros
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

    st.markdown('<div class="sec">🔍 Filtros (opcional — por defecto procesa todo)</div>',
                unsafe_allow_html=True)

    fc1, fc2, fc3 = st.columns([1.5, 1.5, 1])

    with fc1:
        st.markdown('<div class="filter-box"><h4>🏥 Filtrar por centro</h4>', unsafe_allow_html=True)
        opciones_centro = ['Todos los centros'] + centros_disponibles
        sel_centro = st.selectbox('Centro / Servicio', opciones_centro,
                                   label_visibility='collapsed')
        if sel_centro != 'Todos los centros':
            filtro_centro_val = sel_centro
        st.markdown('</div>', unsafe_allow_html=True)

    with fc2:
        st.markdown('<div class="filter-box"><h4>📅 Filtrar por período</h4>', unsafe_allow_html=True)
        MESES = ['Ene','Feb','Mar','Abr','May','Jun',
                 'Jul','Ago','Sep','Oct','Nov','Dic']
        anio_actual = datetime.now().year

        # Rango de años disponibles en la base
        if len(fechas_serie):
            anio_min = max(fechas_serie.dt.year.min(), anio_actual - 10)
            anio_max = min(fechas_serie.dt.year.max(), anio_actual + 1)
        else:
            anio_min, anio_max = anio_actual - 3, anio_actual

        anios = list(range(int(anio_min), int(anio_max)+1))

        p1, p2 = st.columns(2)
        with p1:
            st.caption('Desde')
            mes_d  = st.selectbox('Mes inicio', MESES, index=0,   key='mes_d', label_visibility='collapsed')
            anio_d = st.selectbox('Año inicio', anios, index=0,   key='anio_d', label_visibility='collapsed')
        with p2:
            st.caption('Hasta')
            mes_h  = st.selectbox('Mes fin',   MESES, index=11,   key='mes_h', label_visibility='collapsed')
            anio_h = st.selectbox('Año fin',   anios, index=len(anios)-1, key='anio_h', label_visibility='collapsed')

        usar_periodo = st.checkbox('Aplicar filtro de período', value=False)
        if usar_periodo:
            mes_d_n  = MESES.index(mes_d)  + 1
            mes_h_n  = MESES.index(mes_h)  + 1
            fecha_desde_val = f'{anio_d}-{mes_d_n:02d}'
            fecha_hasta_val = f'{anio_h}-{mes_h_n:02d}'
        st.markdown('</div>', unsafe_allow_html=True)

    with fc3:
        st.markdown('<div class="filter-box"><h4>📄 Reportes a generar</h4>', unsafe_allow_html=True)
        cb_ce  = st.checkbox('Tablas caracterización', value=True,  key='cb_ce')
        cb_se  = st.checkbox('Tablas seguimiento',     value=True,  key='cb_se')
        cb_pc  = st.checkbox('PDF caracterización',    value=True,  key='cb_pc')
        cb_ps  = st.checkbox('PDF seguimiento',        value=True,  key='cb_ps')
        cb_ppc = st.checkbox('PPT caracterización',    value=False, key='cb_ppc')
        cb_pps = st.checkbox('PPT seguimiento',        value=False, key='cb_pps')
        st.markdown('</div>', unsafe_allow_html=True)

    SELECCION = {
        'caract_excel': cb_ce, 'seg_excel': cb_se,
        'pdf_caract':   cb_pc, 'pdf_seg':   cb_ps,
        'pptx_caract':  cb_ppc,'pptx_seg':  cb_pps,
    }

    # ── Resumen de filtros activos ─────────────────────────────────────────────
    badges = ''
    if filtro_centro_val:
        badges += f'<span class="badge badge-centro">🏥 Centro: {filtro_centro_val}</span>'
    if fecha_desde_val:
        badges += f'<span class="badge badge-periodo">📅 {fecha_desde_val} → {fecha_hasta_val}</span>'
    if not badges:
        badges = '<span style="color:#888;font-size:.85rem">Sin filtros — procesa toda la base</span>'

    st.markdown(f'**Archivo:** `{uploaded.name}` &nbsp;|&nbsp; {badges}',
                unsafe_allow_html=True)

    # ── Botón procesar ─────────────────────────────────────────────────────────
    if st.button('⚡ Procesar y generar reportes', use_container_width=True):

        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp.write(uploaded.read()); tmp_raw = tmp.name

        work_dir = tempfile.mkdtemp(prefix='qalat_')

        try:
            with st.spinner('Paso 1/7 — Procesando base Wide...'):
                result = procesar_wide(
                    tmp_raw,
                    filtro_centro = filtro_centro_val,
                    fecha_desde   = fecha_desde_val,
                    fecha_hasta   = fecha_hasta_val,
                )
                st.session_state['result']    = result
                st.session_state['filename']  = uploaded.name
                st.session_state['seleccion'] = SELECCION

                wide_path = os.path.join(work_dir, 'TOP_Base_Wide.xlsx')
                with open(wide_path,'wb') as f:
                    f.write(result['excel_bytes'].getvalue())
                st.session_state['wide_path'] = wide_path
                st.session_state['work_dir']  = work_dir

            st.success(f"✅ Base Wide — {result['stats']['N_total']} pacientes · {result['periodo']}")

            outputs = {}
            keys_sel = [k for k,v in SELECCION.items() if v]
            prog = st.progress(0, text='Generando reportes...')
            for i, key in enumerate(keys_sel):
                lbl = LABELS[key][0]
                prog.progress(i/len(keys_sel), text=f'Generando {lbl}...')
                try:
                    buf, fname, mime = run_script(key, wide_path,
                                                  filtro_centro=filtro_centro_val)
                    outputs[key] = {'ok':True,'buf':buf,'fname':fname,'mime':mime}
                except Exception as e:
                    outputs[key] = {'ok':False,'error':str(e)}
            prog.progress(1.0, text='✅ Listo')
            st.session_state['outputs'] = outputs

        except Exception as e:
            st.error(f'❌ Error: {e}')
        finally:
            os.unlink(tmp_raw)

# ══════════════════════════════════════════════════════════════════════════════
# RESULTADOS
# ══════════════════════════════════════════════════════════════════════════════
if 'result' in st.session_state:
    R = st.session_state['result']
    s = R['stats']; wide = R['wide']
    fc = R.get('filtro_centro'); fd = R.get('fecha_desde'); fh = R.get('fecha_hasta')

    # Badge de filtros aplicados
    filtro_str = ''
    if fc:   filtro_str += f' · Centro: {fc}'
    if fd:   filtro_str += f' · {fd} → {fh}'

    st.markdown('---')
    st.markdown(f'<div class="sec">📊 Resultados — {R["periodo"]}{filtro_str}</div>',
                unsafe_allow_html=True)

    # KPIs
    k1,k2,k3,k4,k5,k6 = st.columns(6)
    for col,lbl,val,sub,cls in [
        (k1,'Pacientes únicos',       s['N_total'], '',                          ''),
        (k2,'Con seguimiento TOP2',   s['N_top2'],  f"{s['pct_top2']}% del total",''),
        (k3,'Solo TOP1 (pendientes)', s['N_solo1'], '',                          ''),
        (k4,'Valores corregidos',     s['N_alertas'],'', 'red' if s['N_alertas'] else 'green'),
        (k5,'🔴 Urgentes (90+ días)', s['n_rojo'],  '',                          'red'),
        (k6,'🟠 Próximos (60–89d)',   s['n_naranja'],'',                         'orange'),
    ]:
        with col:
            st.markdown(f'<div class="kpi {cls}"><div class="kpi-lbl">{lbl}</div>'
                        f'<div class="kpi-val">{val}</div>'
                        f'{"<div class=kpi-sub>"+sub+"</div>" if sub else ""}</div>',
                        unsafe_allow_html=True)

    # Tabla centros
    centros = R.get('centros', [])
    if centros and not fc:   # Solo mostrar si no hay filtro de centro
        st.markdown('<div class="sec">🏥 Resumen por Centro / Servicio de Tratamiento</div>',
                    unsafe_allow_html=True)
        df_c = pd.DataFrame(centros)
        df_c.columns = ['Centro','Aplicaciones','Pacientes únicos',
                         'Con TOP2','Sin TOP2 (pendientes)','Valores corregidos']

        rows_html = ''
        for i, row in df_c.iterrows():
            is_total = str(row.iloc[0]) == 'TOTAL'
            bg = f'background:{NAVY};color:white;font-weight:700;' if is_total else \
                 ('background:#EEF4FB;' if i%2==0 else 'background:white;')
            cells = ''
            for j, val in enumerate(row):
                align = 'left' if j==0 else 'center'
                corr  = (j==5 and not is_total and int(val)>0)
                color = 'white' if is_total else (RED if corr else '#333')
                weight= 'font-weight:700;' if is_total or corr else ''
                cells += f'<td style="padding:7px 12px;text-align:{align};color:{color};{weight}">{val}</td>'
            rows_html += f'<tr style="{bg}">{cells}</tr>'

        hdrs = ''.join(f'<th style="padding:9px 12px;text-align:{"left" if i==0 else "center"};'
                       f'background:{NAVY};color:white;font-size:.85rem;">{c}</th>'
                       for i,c in enumerate(df_c.columns))
        st.markdown(f'<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;'
                    f'font-family:Calibri,sans-serif;font-size:.9rem;">'
                    f'<thead><tr>{hdrs}</tr></thead><tbody>{rows_html}</tbody></table></div>',
                    unsafe_allow_html=True)

    # Gráficos
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
        ax.set_facecolor('#F8FAFD');fig.patch.set_facecolor('#F8FAFD')
        ax.spines[['top','right','left']].set_visible(False);ax.yaxis.set_visible(False)
        plt.tight_layout();st.pyplot(fig);plt.close()

    with gc2:
        fig,ax=plt.subplots(figsize=(4.5,3.2))
        if sv_f:
            w,_,at=ax.pie(sv_f,colors=sc_f,autopct='%1.0f%%',startangle=90,
                wedgeprops={'edgecolor':'white','linewidth':1.5},textprops={'fontsize':9})
            for a in at: a.set_color('white');a.set_fontweight('bold')
            ax.legend(w,[f'{l} ({v})' for l,v in zip(sl_f,sv_f)],
                loc='lower center',bbox_to_anchor=(.5,-.3),fontsize=7.5,ncol=2,frameon=False)
        ax.set_title('Semáforo de seguimiento',fontsize=11,color=NAVY,fontweight='bold',pad=8)
        fig.patch.set_facecolor('#F8FAFD');plt.tight_layout();st.pyplot(fig);plt.close()

    with gc3:
        fig,ax=plt.subplots(figsize=(4.5,3.2))
        if not sd.empty:
            ax.barh(sd['S'],sd['n'],color=colors_s,height=.6)
            tot=sd['n'].sum()
            for b,v in zip(ax.patches,sd['n']):
                ax.text(b.get_width()+.3,b.get_y()+b.get_height()/2,
                        f'{v} ({round(v/tot*100,1) if tot else 0}%)',va='center',fontsize=8,color=NAVY)
            ax.spines[['top','right','bottom']].set_visible(False);ax.xaxis.set_visible(False)
        else:
            ax.text(.5,.5,'Sustancia no detectada',ha='center',va='center',
                    transform=ax.transAxes,fontsize=10,color='#888')
        ax.set_title('Sustancia principal (TOP1)',fontsize=11,color=NAVY,fontweight='bold',pad=8)
        ax.set_facecolor('#F8FAFD');fig.patch.set_facecolor('#F8FAFD')
        plt.tight_layout();st.pyplot(fig);plt.close()

    # Pendientes
    pend=wide[wide['Alerta_TOP2'].isin(['🟠 60-89 dias','🔴 90+ dias'])].copy()
    if len(pend):
        st.markdown('<div class="sec">🚨 Pendientes urgentes</div>', unsafe_allow_html=True)
        pend=pend.loc[:,~pend.columns.duplicated()]
        id_c=wide.columns[0]; cs=[id_c]
        col_c=next((c for c in pend.columns if 'centro' in c.lower() and '_TOP1' in c),None)
        col_f=next((c for c in pend.columns if 'fecha entrevista' in c.lower() and '_TOP1' in c),None)
        if col_c: cs.append(col_c)
        if col_f: cs.append(col_f)
        cs+=['Dias_desde_TOP1','Alerta_TOP2']
        cs=list(dict.fromkeys(c for c in cs if c in pend.columns))
        tab=pend[cs].copy()
        tab['_o']=tab['Alerta_TOP2'].apply(lambda x: 0 if '90' in str(x) else 1)
        tab=tab.sort_values(['_o','Dias_desde_TOP1'],ascending=[True,False]).drop(columns='_o')
        st.dataframe(tab.head(30),use_container_width=True,height=280)

    with st.expander('📋 Log de procesamiento'):
        for log in R['logs']: st.text(log)

    # DESCARGAS
    st.markdown('---')
    st.markdown('<div class="sec">⬇️ Descargar reportes</div>', unsafe_allow_html=True)

    fname_base = os.path.splitext(st.session_state.get('filename','base'))[0]
    if fc: fname_base += f'_{fc}'
    if fd: fname_base += f'_{fd}_{fh}'
    today_str = datetime.now().strftime('%Y-%m-%d')
    outputs   = st.session_state.get('outputs',{})
    sel       = st.session_state.get('seleccion',{})

    # Fila 1
    d1,d2,d3=st.columns(3)
    with d1:
        st.markdown('<div class="outcard"><h4>📊 Base Wide completa</h4>'
                    '<p>6 hojas: Wide · Resumen · Alertas · Calidad · Por Centro · Pendientes</p></div>',
                    unsafe_allow_html=True)
        st.download_button('⬇️ Base Wide (.xlsx)',
            data=R['excel_bytes'].getvalue(),
            file_name=f'TOP_Base_Wide_{fname_base}_{today_str}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True,key='dl_wide')

    for key,col,dlkey in [('caract_excel',d2,'dl_ce'),('seg_excel',d3,'dl_se')]:
        o=outputs.get(key,{}); lbl,fmt,desc=LABELS[key]
        with col:
            st.markdown(f'<div class="outcard"><h4>{lbl}</h4><p>{desc}</p></div>',unsafe_allow_html=True)
            if not sel.get(key,False):
                st.caption('No seleccionado')
            elif o.get('ok'):
                st.download_button(f'⬇️ {fmt}',data=o['buf'].getvalue(),
                    file_name=o['fname'],mime=o['mime'],use_container_width=True,key=dlkey)
            else:
                st.warning(f"⚠️ {o.get('error','Error')[:100]}")

    st.markdown('---')
    d4,d5=st.columns(2)
    for key,col,dlkey in [('pdf_caract',d4,'dl_pc'),('pdf_seg',d5,'dl_ps')]:
        o=outputs.get(key,{}); lbl,fmt,desc=LABELS[key]
        with col:
            st.markdown(f'<div class="outcard"><h4>{lbl}</h4><p>{desc}</p></div>',unsafe_allow_html=True)
            if not sel.get(key,False): st.caption('No seleccionado')
            elif o.get('ok'):
                st.download_button(f'⬇️ {fmt}',data=o['buf'].getvalue(),
                    file_name=o['fname'],mime=o['mime'],use_container_width=True,key=dlkey)
            else: st.warning(f"⚠️ {o.get('error','Error')[:100]}")

    st.markdown('---')
    d6,d7=st.columns(2)
    for key,col,dlkey in [('pptx_caract',d6,'dl_ppc'),('pptx_seg',d7,'dl_pps')]:
        o=outputs.get(key,{}); lbl,fmt,desc=LABELS[key]
        with col:
            st.markdown(f'<div class="outcard"><h4>{lbl}</h4><p>{desc}</p></div>',unsafe_allow_html=True)
            if not sel.get(key,False): st.caption('No seleccionado')
            elif o.get('ok'):
                st.download_button(f'⬇️ {fmt}',data=o['buf'].getvalue(),
                    file_name=o['fname'],mime=o['mime'],use_container_width=True,key=dlkey)
            else: st.warning(f"⚠️ {o.get('error','Error')[:100]}")

if not uploaded and 'result' not in st.session_state:
    st.markdown("""<div style="text-align:center;padding:3rem;color:#888;">
        <div style="font-size:3rem;">📤</div>
        <div style="font-size:1.1rem;margin-top:1rem;">Sube tu Excel para comenzar</div>
        <div style="font-size:.85rem;margin-top:.5rem;color:#aaa;">Base bruta exportada de Jotform · instrumento TOP</div>
    </div>""", unsafe_allow_html=True)
