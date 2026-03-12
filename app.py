"""
App Streamlit — Análisis de Informes de Mantenimiento
======================================================
Instalación (una sola vez):
    pip install streamlit plotly openpyxl pandas numpy python-docx

Ejecución:
    streamlit run app_mantenimiento.py
"""

import io
import re
import json
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Análisis de Mantenimiento",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# ESTILOS CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');

    html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

    .stApp { background-color: #0f1117; color: #e8eaf0; }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #161b27;
        border-right: 1px solid #2a3040;
    }

    /* Títulos */
    h1 { color: #4fc3f7 !important; font-weight: 700 !important; letter-spacing: -0.5px; }
    h2 { color: #81d4fa !important; font-weight: 600 !important; }
    h3 { color: #b0bec5 !important; font-weight: 600 !important; }

    /* KPI Cards */
    .kpi-card {
        background: linear-gradient(135deg, #1a2235 0%, #1e2a3a 100%);
        border: 1px solid #2a3a50;
        border-left: 4px solid #4fc3f7;
        border-radius: 8px;
        padding: 20px 24px;
        text-align: center;
        transition: border-color 0.2s;
    }
    .kpi-card:hover { border-left-color: #81d4fa; }
    .kpi-value {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 2.4rem;
        font-weight: 700;
        color: #4fc3f7;
        line-height: 1;
        margin-bottom: 6px;
    }
    .kpi-label {
        font-size: 0.78rem;
        color: #78909c;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }
    .kpi-card.danger  { border-left-color: #ef5350; }
    .kpi-card.danger .kpi-value { color: #ef5350; }
    .kpi-card.warning { border-left-color: #ffa726; }
    .kpi-card.warning .kpi-value { color: #ffa726; }
    .kpi-card.success { border-left-color: #66bb6a; }
    .kpi-card.success .kpi-value { color: #66bb6a; }

    /* Section divider */
    .section-header {
        border-bottom: 2px solid #2a3a50;
        padding-bottom: 8px;
        margin: 32px 0 16px 0;
        color: #4fc3f7;
        font-size: 1.1rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.06em;
    }

    /* Desviación badge */
    .desv-pos { color: #ef5350; font-weight: 700; font-family: 'IBM Plex Mono', monospace; }
    .desv-neg { color: #42a5f5; font-weight: 700; font-family: 'IBM Plex Mono', monospace; }
    .crit-c   { background: #b71c1c; color: white; padding: 2px 8px; border-radius: 4px; font-size: 0.8rem; font-weight: 700; }
    .crit-r   { background: #e65100; color: white; padding: 2px 8px; border-radius: 4px; font-size: 0.8rem; font-weight: 700; }
    .crit-n   { background: #263238; color: #90a4ae; padding: 2px 8px; border-radius: 4px; font-size: 0.8rem; }

    /* Upload area */
    .uploadedFile { background: #1a2235 !important; border: 1px solid #2a3a50 !important; }

    /* Dataframe */
    .stDataFrame { border-radius: 8px; overflow: hidden; }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; background: transparent; }
    .stTabs [data-baseweb="tab"] {
        background: #1a2235;
        border-radius: 6px 6px 0 0;
        color: #78909c;
        padding: 8px 20px;
        border: 1px solid #2a3a50;
        border-bottom: none;
    }
    .stTabs [aria-selected="true"] {
        background: #1e2a3a !important;
        color: #4fc3f7 !important;
        border-color: #4fc3f7 !important;
    }

    /* Selectbox, multiselect */
    .stSelectbox > div, .stMultiSelect > div { background: #1a2235 !important; }

    /* Info/warning boxes */
    .stAlert { border-radius: 8px; }

    /* Scrollbar */
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #0f1117; }
    ::-webkit-scrollbar-thumb { background: #2a3a50; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# FUNCIONES DE ANÁLISIS
# ─────────────────────────────────────────────

MONTH_ORDER = [
    'Enero','Febrero','Marzo','Abril','MAYO','JUNIO',
    'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'
]

SISTEMA_LABELS = {
    'BG':'Bogie','SLN':'Salón','EXT':'Exterior','TYC':'Tracción y Comando',
    'PM':'Partes Mecánicas','EBC':'Equipo de a Bordo','NSF':'Neumo-freno (S/F)',
    'MSF':'Neumo-freno (C/F)','DSM':'Suspensión','CAB':'Cabina',
    'ATS':'ATS','DOC':'Documentación','NGN':'Motor'
}

def parse_valor(v):
    if v is None:
        return np.nan
    try:
        return float(v)
    except Exception:
        nums = re.findall(r'[\d.]+', str(v))
        return float(nums[0]) if nums else np.nan

def calcular_desvio(row):
    vals = [v for v in [row['R1_num'], row['R2_num']] if not np.isnan(v)]
    if not vals:
        return np.nan, np.nan
    mejor_dev, mejor_val = 0, vals[0]
    for v in vals:
        if not np.isnan(row['RefMin_num']) and v < row['RefMin_num']:
            dev = v - row['RefMin_num']
            if abs(dev) > abs(mejor_dev):
                mejor_dev, mejor_val = dev, v
        if not np.isnan(row['RefMax_num']) and v > row['RefMax_num']:
            dev = v - row['RefMax_num']
            if abs(dev) > abs(mejor_dev):
                mejor_dev, mejor_val = dev, v
    return mejor_val, mejor_dev

@st.cache_data(show_spinner=False)
def cargar_y_analizar(file_bytes):
    """Lee el Excel y devuelve df principal + df de desvíos."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    rows1, rows2 = [], []
    for row in ws.iter_rows(min_row=3, values_only=True):
        r1, r2 = row[1:34], row[36:69]
        if any(v is not None for v in r1): rows1.append(r1)
        if any(v is not None for v in r2): rows2.append(r2)

    cols = [
        'Mes','Responsable','Contrato','Linea','Vehiculo','Modulo','MR',
        'Modelo','Servicio','Fecha','NroInforme','SistemaUnidad',
        'SistemaAmpliado','Item1','Item2','Descripcion',
        'RefMin','RefMax','Relevado1','Relevado2','Criticidad',
        'DescAgrupada','CritAmpliado','CodItem','FechaReInsp','NroReInsp',
        'SistUnitReInsp','SistAmpReInsp','ItemsReInsp','DescReInsp',
        'CritReInsp','DescAgrupReInsp','CodReInsp'
    ]
    df = pd.concat([
        pd.DataFrame(rows1, columns=cols),
        pd.DataFrame(rows2, columns=cols)
    ], ignore_index=True)

    # Calcular desvíos
    df_ref = df[df['RefMin'].notna() | df['RefMax'].notna()].copy()
    df_ref['R1_num']     = df_ref['Relevado1'].apply(parse_valor)
    df_ref['R2_num']     = df_ref['Relevado2'].apply(parse_valor)
    df_ref['RefMin_num'] = pd.to_numeric(df_ref['RefMin'], errors='coerce')
    df_ref['RefMax_num'] = pd.to_numeric(df_ref['RefMax'], errors='coerce')
    df_ref[['ValorRelevado','Desviacion']] = df_ref.apply(
        lambda r: pd.Series(calcular_desvio(r)), axis=1
    )
    df_out = df_ref[df_ref['Desviacion'].notna() & (df_ref['Desviacion'] != 0)].copy()

    return df, df_out


def kpi(label, value, variant="default"):
    return f"""
    <div class="kpi-card {variant}">
        <div class="kpi-value">{value}</div>
        <div class="kpi-label">{label}</div>
    </div>"""


# ─────────────────────────────────────────────
# GENERACIÓN DE WORD (descarga)
# ─────────────────────────────────────────────

def generar_word(df, df_out):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    # Título
    titulo = doc.add_heading('Informe de Fallas y Desvíos', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titulo.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

    doc.add_paragraph(f"Período analizado: {df['Mes'].dropna().nunique()} meses  |  Vehículos: {df['Vehiculo'].dropna().nunique()}")
    doc.add_paragraph()

    def add_section(title):
        h = doc.add_heading(title, level=1)
        h.runs[0].font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

    def table_from_df(dataframe, col_widths=None):
        t = doc.add_table(rows=1, cols=len(dataframe.columns))
        t.style = 'Table Grid'
        hdr = t.rows[0].cells
        for i, col in enumerate(dataframe.columns):
            hdr[i].text = str(col)
            hdr[i].paragraphs[0].runs[0].bold = True
            hdr[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            hdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for _, row in dataframe.iterrows():
            cells = t.add_row().cells
            for i, val in enumerate(row):
                cells[i].text = str(val)
                cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        return t

    # Resumen criticidad
    add_section("1. Distribución por Criticidad")
    crit_df = df['Criticidad'].value_counts().reset_index()
    crit_df.columns = ['Criticidad', 'Cantidad']
    crit_df['Porcentaje'] = (crit_df['Cantidad'] / len(df) * 100).round(1).astype(str) + '%'
    table_from_df(crit_df)
    doc.add_paragraph()

    # Top fallas
    add_section("2. Top 15 Tipos de Falla")
    top_df = df['DescAgrupada'].value_counts().head(15).reset_index()
    top_df.columns = ['Descripción', 'Cantidad']
    table_from_df(top_df)
    doc.add_paragraph()

    # Desvíos
    add_section("3. Resumen de Desvíos por Categoría")
    desv_sum = df_out.groupby('DescAgrupada').agg(
        Cantidad=('Desviacion','count'),
        Desv_Promedio=('Desviacion', lambda x: round(x.mean(),2)),
        Desv_Max_Abs=('Desviacion', lambda x: round(x.abs().max(),2)),
        Criticos=('Criticidad', lambda x: (x=='C').sum())
    ).reset_index().sort_values('Cantidad', ascending=False)
    desv_sum.columns = ['Categoría','Cantidad','Desvío Prom.','Desvío Máx.','Críticos']
    table_from_df(desv_sum)
    doc.add_paragraph()

    # Detalle
    add_section("4. Detalle de Observaciones Fuera de Parámetro")
    det = df_out[['Vehiculo','Mes','SistemaUnidad','Descripcion','RefMin','RefMax',
                  'ValorRelevado','Desviacion','Criticidad']].copy()
    det.columns = ['Vehículo','Mes','Sistema','Descripción','Ref Min','Ref Max',
                   'Relevado','Desvío','Crit.']
    det['Desvío'] = det['Desvío'].round(2)
    det['Relevado'] = det['Relevado'].round(1)
    table_from_df(det)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 🔧 Mantenimiento LSM")
    st.markdown("---")
    uploaded = st.file_uploader(
        "Subir archivo Excel",
        type=["xlsx"],
        help="Archivo de seguimiento de inspecciones estáticas"
    )
    st.markdown("---")
    st.markdown("""
    **Cómo usar:**
    1. Subí el archivo `.xlsx`
    2. Explorá los paneles
    3. Filtrá por mes, sistema o vehículo
    4. Descargá el informe Word
    """)
    st.markdown("---")
    st.caption("Análisis automático de desvíos · Línea San Martín")


# ─────────────────────────────────────────────
# PANTALLA PRINCIPAL
# ─────────────────────────────────────────────

st.title("🔧 Análisis de Informes de Mantenimiento")

if uploaded is None:
    st.markdown("""
    <div style="
        background: linear-gradient(135deg,#1a2235,#1e2a3a);
        border: 1px dashed #2a3a50;
        border-radius: 12px;
        padding: 60px;
        text-align: center;
        margin-top: 40px;
    ">
        <div style="font-size:3rem;margin-bottom:16px">📂</div>
        <div style="color:#4fc3f7;font-size:1.3rem;font-weight:700;margin-bottom:8px">
            Subí tu archivo Excel para comenzar
        </div>
        <div style="color:#546e7a;font-size:0.9rem">
            Usá el panel izquierdo para cargar el archivo de inspecciones
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Cargar y procesar ──
with st.spinner("Procesando archivo..."):
    df, df_out = cargar_y_analizar(uploaded.read())

total_obs   = len(df)
total_veh   = df['Vehiculo'].dropna().nunique()
total_crit  = int((df['Criticidad'] == 'C').sum())
total_rech  = int((df['Criticidad'] == 'R').sum())
total_desv  = len(df_out)

# ─────────────────────────────────────────────
# KPIs
# ─────────────────────────────────────────────
st.markdown('<div class="section-header">Resumen del período</div>', unsafe_allow_html=True)

c1, c2, c3, c4, c5 = st.columns(5)
c1.markdown(kpi("Total Observaciones", total_obs), unsafe_allow_html=True)
c2.markdown(kpi("Vehículos", total_veh), unsafe_allow_html=True)
c3.markdown(kpi("Críticas (C)", total_crit, "danger"), unsafe_allow_html=True)
c4.markdown(kpi("Rechazadas (R)", total_rech, "warning"), unsafe_allow_html=True)
c5.markdown(kpi("Fuera de Parámetro", total_desv, "warning"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TABS PRINCIPALES
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📊  Resúmenes",
    "📈  Gráficos",
    "🔍  Desvíos Detallados",
    "📄  Exportar Informe"
])

PLOTLY_THEME = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(26,34,53,0.6)',
    font=dict(family='IBM Plex Sans', color='#b0bec5'),
    colorway=['#4fc3f7','#81d4fa','#ef5350','#ffa726','#66bb6a','#ab47bc','#26c6da'],
    xaxis=dict(gridcolor='#1e2a3a', linecolor='#2a3a50'),
    yaxis=dict(gridcolor='#1e2a3a', linecolor='#2a3a50'),
)


# ── TAB 1: RESÚMENES ──
with tab1:
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### Criticidad")
        crit_df = df['Criticidad'].value_counts().reset_index()
        crit_df.columns = ['Criticidad','Cantidad']
        crit_df['Porcentaje'] = (crit_df['Cantidad'] / total_obs * 100).round(1).astype(str) + '%'
        st.dataframe(crit_df, use_container_width=True, hide_index=True)

        st.markdown("#### Por Modelo")
        mod_df = df['Modelo'].value_counts().reset_index()
        mod_df.columns = ['Modelo','Observaciones']
        mod_df['%'] = (mod_df['Observaciones'] / total_obs * 100).round(1).astype(str) + '%'
        st.dataframe(mod_df, use_container_width=True, hide_index=True)

    with col_b:
        st.markdown("#### Por Sistema de Unidad")
        sist_df = df['SistemaUnidad'].value_counts().reset_index()
        sist_df.columns = ['Código','Cantidad']
        sist_df['Sistema'] = sist_df['Código'].map(SISTEMA_LABELS).fillna(sist_df['Código'])
        sist_df['%'] = (sist_df['Cantidad'] / total_obs * 100).round(1).astype(str) + '%'
        st.dataframe(sist_df[['Código','Sistema','Cantidad','%']], use_container_width=True, hide_index=True)

        st.markdown("#### Top 10 Vehículos con más observaciones")
        veh_df = df['Vehiculo'].value_counts().head(10).reset_index()
        veh_df.columns = ['Vehículo','Observaciones']
        st.dataframe(veh_df, use_container_width=True, hide_index=True)

    st.markdown("#### Top 15 Tipos de Falla")
    top15 = df['DescAgrupada'].value_counts().head(15).reset_index()
    top15.columns = ['Descripción Agrupada','Cantidad']
    top15['% del total'] = (top15['Cantidad'] / total_obs * 100).round(1).astype(str) + '%'
    st.dataframe(top15, use_container_width=True, hide_index=True)


# ── TAB 2: GRÁFICOS ──
with tab2:
    row1_l, row1_r = st.columns(2)

    # Torta criticidad
    with row1_l:
        st.markdown("#### Distribución por Criticidad")
        crit_counts = df['Criticidad'].value_counts()
        fig_pie = go.Figure(go.Pie(
            labels=crit_counts.index,
            values=crit_counts.values,
            hole=0.55,
            marker=dict(colors=['#4fc3f7','#ef5350','#ffa726','#66bb6a','#ab47bc']),
            textfont=dict(size=13)
        ))
        fig_pie.update_layout(**PLOTLY_THEME, margin=dict(t=10,b=10,l=10,r=10), height=300,
                              legend=dict(orientation='h', y=-0.1))
        st.plotly_chart(fig_pie, use_container_width=True)

    # Barras sistemas
    with row1_r:
        st.markdown("#### Observaciones por Sistema")
        sist_counts = df['SistemaUnidad'].value_counts().reset_index()
        sist_counts.columns = ['Sistema','Cantidad']
        sist_counts['Label'] = sist_counts['Sistema'].map(SISTEMA_LABELS).fillna(sist_counts['Sistema'])
        fig_sist = px.bar(sist_counts, x='Cantidad', y='Label', orientation='h',
                          color='Cantidad', color_continuous_scale='Blues')
        fig_sist.update_layout(**PLOTLY_THEME, margin=dict(t=10,b=10,l=10,r=10), height=300,
                               yaxis=dict(autorange='reversed', **PLOTLY_THEME['yaxis']),
                               coloraxis_showscale=False)
        st.plotly_chart(fig_sist, use_container_width=True)

    # Evolución mensual
    st.markdown("#### Evolución Mensual de Observaciones")
    monthly = []
    for m in MONTH_ORDER:
        cnt = df[df['Mes'].str.upper() == m.upper()].shape[0]
        if cnt > 0:
            monthly.append({'Mes': m, 'Observaciones': cnt})
    monthly_df = pd.DataFrame(monthly)
    if not monthly_df.empty:
        fig_line = go.Figure()
        fig_line.add_trace(go.Scatter(
            x=monthly_df['Mes'], y=monthly_df['Observaciones'],
            mode='lines+markers',
            line=dict(color='#4fc3f7', width=3),
            marker=dict(size=8, color='#4fc3f7', line=dict(color='white', width=2)),
            fill='tozeroy',
            fillcolor='rgba(79,195,247,0.1)'
        ))
        fig_line.update_layout(**PLOTLY_THEME, margin=dict(t=10,b=10,l=10,r=10), height=280)
        st.plotly_chart(fig_line, use_container_width=True)

    # Top fallas barras
    st.markdown("#### Top 15 Tipos de Falla")
    top15_plot = df['DescAgrupada'].value_counts().head(15).reset_index()
    top15_plot.columns = ['Falla','Cantidad']
    fig_top = px.bar(top15_plot, x='Cantidad', y='Falla', orientation='h',
                     color='Cantidad', color_continuous_scale='Blues_r')
    fig_top.update_layout(**PLOTLY_THEME, height=420, margin=dict(t=10,b=10,l=10,r=10),
                          yaxis=dict(autorange='reversed', **PLOTLY_THEME['yaxis']),
                          coloraxis_showscale=False)
    st.plotly_chart(fig_top, use_container_width=True)

    # Desvíos por categoría
    if not df_out.empty:
        st.markdown("#### Desvíos: Distribución por Categoría")
        desv_plot = df_out.groupby('DescAgrupada').agg(
            Cantidad=('Desviacion','count'),
            Desv_Max=('Desviacion', lambda x: x.abs().max())
        ).reset_index().sort_values('Cantidad', ascending=False)

        fig_desv = go.Figure()
        fig_desv.add_trace(go.Bar(
            x=desv_plot['DescAgrupada'], y=desv_plot['Cantidad'],
            name='Cantidad', marker_color='#4fc3f7', yaxis='y'
        ))
        fig_desv.add_trace(go.Scatter(
            x=desv_plot['DescAgrupada'], y=desv_plot['Desv_Max'],
            name='Desvío Máx.', mode='lines+markers',
            line=dict(color='#ef5350', width=2), marker=dict(size=7),
            yaxis='y2'
        ))
        fig_desv.update_layout(
            **PLOTLY_THEME, height=380, barmode='group',
            margin=dict(t=10,b=80,l=10,r=60),
            xaxis=dict(tickangle=-35, **PLOTLY_THEME['xaxis']),
            yaxis=dict(title='Cantidad', **PLOTLY_THEME['yaxis']),
            yaxis2=dict(title='Desvío Máx.', overlaying='y', side='right',
                        gridcolor='#1e2a3a', linecolor='#2a3a50'),
            legend=dict(orientation='h', y=1.05)
        )
        st.plotly_chart(fig_desv, use_container_width=True)


# ── TAB 3: DESVÍOS DETALLADOS ──
with tab3:
    st.markdown("#### Filtros")
    fc1, fc2, fc3 = st.columns(3)

    meses_disponibles = sorted(df_out['Mes'].dropna().unique().tolist())
    sistemas_disponibles = sorted(df_out['SistemaUnidad'].dropna().unique().tolist())
    criticas_opciones = sorted(df_out['Criticidad'].dropna().unique().tolist())

    with fc1:
        sel_mes = st.multiselect("Mes", meses_disponibles, default=meses_disponibles,
                                 placeholder="Todos los meses")
    with fc2:
        sel_sist = st.multiselect("Sistema", sistemas_disponibles, default=sistemas_disponibles,
                                  placeholder="Todos los sistemas")
    with fc3:
        sel_crit = st.multiselect("Criticidad", criticas_opciones, default=criticas_opciones,
                                  placeholder="Todas")

    df_filtrado = df_out[
        df_out['Mes'].isin(sel_mes) &
        df_out['SistemaUnidad'].isin(sel_sist) &
        df_out['Criticidad'].isin(sel_crit)
    ].copy()

    # Rango referencia como string
    def ref_str(row):
        mn = row['RefMin_num'] if pd.notna(row.get('RefMin_num')) else None
        mx = row['RefMax_num'] if pd.notna(row.get('RefMax_num')) else None
        if mn is not None and mx is not None: return f"{mn} – {mx}"
        if mn is not None: return f"≥ {mn}"
        if mx is not None: return f"≤ {mx}"
        return "—"

    df_filtrado['RefMin_num'] = pd.to_numeric(df_filtrado['RefMin'], errors='coerce')
    df_filtrado['RefMax_num'] = pd.to_numeric(df_filtrado['RefMax'], errors='coerce')
    df_filtrado['Rango Ref.'] = df_filtrado.apply(ref_str, axis=1)
    df_filtrado['Desvío'] = df_filtrado['Desviacion'].round(2)
    df_filtrado['Relevado'] = df_filtrado['ValorRelevado'].round(1)

    st.markdown(f"**{len(df_filtrado)}** observaciones fuera de parámetro")

    # Resumen por categoría
    st.markdown("##### Resumen por Categoría")
    resumen = df_filtrado.groupby('DescAgrupada').agg(
        Cantidad=('Desvío','count'),
        Desvío_Promedio=('Desvío', lambda x: round(x.mean(),2)),
        Desvío_Máx=('Desvío', lambda x: round(x.abs().max(),2)),
        Críticos=('Criticidad', lambda x: (x=='C').sum())
    ).reset_index().sort_values('Cantidad', ascending=False)
    resumen.columns = ['Categoría','Cantidad','Desvío Prom.','Desvío Máx. (abs)','Críticos']
    st.dataframe(resumen, use_container_width=True, hide_index=True)

    # Tabla detalle
    st.markdown("##### Detalle completo")
    cols_show = ['Vehiculo','Mes','SistemaUnidad','Descripcion','Rango Ref.','Relevado','Desvío','Criticidad']
    rename_map = {'Vehiculo':'Vehículo','SistemaUnidad':'Sistema','Descripcion':'Descripción'}
    tabla_det = df_filtrado[cols_show].rename(columns=rename_map).reset_index(drop=True)

    st.dataframe(
        tabla_det,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Desvío": st.column_config.NumberColumn(format="%.2f"),
            "Relevado": st.column_config.NumberColumn(format="%.1f"),
            "Criticidad": st.column_config.TextColumn(),
        }
    )


# ── TAB 4: EXPORTAR ──
with tab4:
    st.markdown("#### Generación del Informe Word")
    st.markdown("""
    El informe incluye:
    - Resumen de criticidad
    - Top 15 tipos de falla
    - Tabla resumen de desvíos por categoría
    - Detalle completo de todas las observaciones fuera de parámetro
    """)

    if st.button("📄  Generar y descargar informe Word", type="primary"):
        with st.spinner("Generando documento..."):
            word_bytes = generar_word(df, df_out)

        st.success("✅ Informe generado correctamente.")
        st.download_button(
            label="⬇️  Descargar .docx",
            data=word_bytes,
            file_name="Informe_Mantenimiento.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
