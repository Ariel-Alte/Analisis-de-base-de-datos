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
    'ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
    'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'
]

SISTEMA_LABELS = {
    'BG' : 'Bogie',
    'SLN': 'Salón',
    'EXT': 'Exterior',
    'TYC': 'Tracción y Choque',
    'PM' : 'Par Montado',
    'EBC': 'Elementos bajo coche',
    'NSF': 'Sistema de freno Neumatico',
    'MSF': 'Sistema de freno Mecanico',
    'SFM': 'Sistema de freno Mecanico',
    'DSM': 'Sala de motor Diesel',
    'CAB': 'Cabina',
    'ATS': 'ATS',
    'DOC': 'Documentación',
    'NGN': 'Ninguna observación en el informe',
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
    """
    Lee el Excel y devuelve (df, df_out).
    Auto-detecta dos formatos:
      Formato A: dos bloques en paralelo, con columnas RefMin/RefMax/Relevado
      Formato B: un solo bloque, sin columnas de referencia (ej: Roca)
    """
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    # ── 1. Detectar fila de headers ──
    header_row = 2
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        if row[1] is not None:
            header_row = i
            break
    data_start = header_row + 1

    # ── 2. Leer headers reales ──
    headers_raw = [ws.cell(row=header_row, column=c).value
                   for c in range(1, ws.max_column + 1)]

    # ── 3. Detectar si hay segundo bloque (se repite el primer header) ──
    first_header = next((h for h in headers_raw if h is not None), None)
    b1_start = next((i for i, h in enumerate(headers_raw) if h is not None), 1)
    bloque2_col = None
    if first_header:
        found = False
        for i, h in enumerate(headers_raw):
            if h == first_header:
                if not found:
                    found = True
                else:
                    bloque2_col = i
                    break

    b1_end = bloque2_col if bloque2_col else len(headers_raw)

    # ── 4. Detectar formato segun headers del bloque 1 ──
    hdrs_b1 = [h for h in headers_raw[b1_start:b1_end] if h is not None]
    tiene_refs = any('Referencia' in str(h) or 'Relevado' in str(h) for h in hdrs_b1)

    # ── 5. Leer filas ──
    rows1, rows2 = [], []
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        r1 = row[b1_start:b1_end]
        if any(v is not None for v in r1):
            rows1.append(r1)
        if bloque2_col is not None:
            b2_start = bloque2_col
            b2_end = b2_start + (b1_end - b1_start)
            if b2_end <= len(row):
                r2 = row[b2_start:b2_end]
                if any(v is not None for v in r2):
                    rows2.append(r2)

    # ── 6. Definir columnas segun formato ──
    if tiene_refs:
        # Formato A: con RefMin, RefMax, Relevado1, Relevado2
        cols = [
            'Mes','Responsable','Contrato','Linea','Vehiculo','Modulo','MR',
            'Modelo','Servicio','Fecha','NroInforme','SistemaUnidad',
            'SistemaAmpliado','Item1','Item2','Descripcion',
            'RefMin','RefMax','Relevado1','Relevado2','Criticidad',
            'DescAgrupada','CritAmpliado','CodItem','FechaReInsp','NroReInsp',
            'SistUnitReInsp','SistAmpReInsp','ItemsReInsp','DescReInsp',
            'CritReInsp','DescAgrupReInsp','CodReInsp','Clasificacion'
        ]
    else:
        # Formato B: sin columnas de referencia numerica
        cols = [
            'Mes','Responsable','Contrato','Linea','Vehiculo','Modulo','MR',
            'Modelo','Servicio','Fecha','NroInforme','SistemaUnidad',
            'SistemaAmpliado','Item1','Item2','Descripcion',
            'Criticidad','DescAgrupada','CritAmpliado','CodItem',
            'FechaReInsp','NroReInsp','SistUnitReInsp','SistAmpReInsp',
            'ItemsReInsp','DescReInsp','CritReInsp','DescAgrupReInsp','CodReInsp',
            'Clasificacion'
        ]
        # Agregar columnas de referencia vacias para que el resto del codigo funcione
        # (df_out quedara vacio, tab de desvios mostrara mensaje informativo)

    def filas_a_df(rows):
        if not rows:
            return pd.DataFrame(columns=cols)
        n, n_cols = len(rows[0]), len(cols)
        if n == n_cols:
            return pd.DataFrame(rows, columns=cols)
        elif n < n_cols:
            rows_pad = [list(r) + [None] * (n_cols - n) for r in rows]
            return pd.DataFrame(rows_pad, columns=cols)
        else:
            return pd.DataFrame([r[:n_cols] for r in rows], columns=cols)

    dfs = [filas_a_df(rows1)]
    if rows2:
        dfs.append(filas_a_df(rows2))
    df = pd.concat(dfs, ignore_index=True)

    # Normalizar columna Mes (unificar mayusculas/minusculas)
    if 'Mes' in df.columns:
        df['Mes'] = df['Mes'].astype(str).str.strip().str.upper()

    # Limpiar columna Clasificacion: eliminar fórmulas Excel no calculadas
    if 'Clasificacion' in df.columns:
        df['Clasificacion'] = df['Clasificacion'].apply(
            lambda x: None if (pd.isna(x) or str(x).strip().startswith('=')) else str(x).strip()
        )

    # ── 7. Calcular desvios (solo si hay columnas de referencia) ──
    if tiene_refs and all(c in df.columns for c in ['RefMin','RefMax','Relevado1','Relevado2']):
        df_ref = df[df['RefMin'].notna() | df['RefMax'].notna()].copy()
        df_ref['R1_num']     = df_ref['Relevado1'].apply(parse_valor)
        df_ref['R2_num']     = df_ref['Relevado2'].apply(parse_valor)
        df_ref['RefMin_num'] = pd.to_numeric(df_ref['RefMin'], errors='coerce')
        df_ref['RefMax_num'] = pd.to_numeric(df_ref['RefMax'], errors='coerce')
        df_ref[['ValorRelevado','Desviacion']] = df_ref.apply(
            lambda r: pd.Series(calcular_desvio(r)), axis=1
        )
        df_out = df_ref[df_ref['Desviacion'].notna() & (df_ref['Desviacion'] != 0)].copy()
    else:
        # Formato sin referencias: df_out vacio con columnas esperadas
        df_out = pd.DataFrame(columns=list(df.columns) + ['ValorRelevado','Desviacion'])

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

def generar_word(df, df_out, config=None):
    from docx.shared import Cm
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from datetime import date

    if config is None:
        config = {}

    # Defaults
    inc_crit     = config.get('inc_crit',     True)
    inc_top15    = config.get('inc_top15',    True)
    inc_desvios  = config.get('inc_desvios',  True)
    inc_detalle  = config.get('inc_detalle',  True)
    inc_graficos = config.get('inc_graficos', False)
    inc_concl    = config.get('inc_concl',    True)
    hdr_codigo   = config.get('codigo',  '') or ''
    hdr_version  = config.get('version', 'v1.0') or 'v1.0'
    hdr_linea    = config.get('linea',   '') or ''
    hdr_subger   = config.get('subger',  'Sub Gerencia de Programación y Seguimiento de Mantenimiento de Material Rodante (SPySM)') or ''
    logo_bytes   = config.get('logo',    None)

    doc = Document()

    # ── Configuración de página A4 ──
    from docx.shared import Mm
    section = doc.sections[0]
    section.page_width  = Mm(210)
    section.page_height = Mm(297)
    section.top_margin    = Mm(25)
    section.bottom_margin = Mm(20)
    section.left_margin   = Mm(20)
    section.right_margin  = Mm(20)

    # ── Estilos ──
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(10)

    # Ancho de contenido en DXA (1 DXA = 1/20 pt; A4 con márgenes 20mm c/u = 170mm)
    PAGE_W_DXA = int((210 - 40) * 56.69)  # ≈ 9639 DXA

    # ── ENCABEZADO ──
    header = section.header
    # Limpiar párrafo por defecto
    for p in header.paragraphs:
        p.clear()

    hdr_tbl = header.add_table(rows=2, cols=3, width=Mm(170))
    hdr_tbl.style = 'Table Grid'

    def set_cell_bg(cell, hex_color):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement('w:shd')
        shd.set(qn('w:val'),   'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'),  hex_color)
        tcPr.append(shd)

    def hdr_cell_text(cell, text, bold=False, size=9, color="000000", align=WD_ALIGN_PARAGRAPH.LEFT):
        cell.paragraphs[0].clear()
        run = cell.paragraphs[0].add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = 'Arial'
        run.font.color.rgb = RGBColor.from_string(color)
        cell.paragraphs[0].alignment = align

    # Fila 0: banner/logo ocupa col 0-1, datos en col 2
    banner_cell = hdr_tbl.cell(0, 0).merge(hdr_tbl.cell(0, 1))
    if logo_bytes:
        banner_cell.paragraphs[0].clear()
        run_img = banner_cell.paragraphs[0].add_run()
        run_img.add_picture(io.BytesIO(logo_bytes), width=Mm(110))
    else:
        set_cell_bg(banner_cell, '1F3864')
        hdr_cell_text(banner_cell, 'TRENES ARGENTINOS — PISE', bold=True, size=11, color='FFFFFF', align=WD_ALIGN_PARAGRAPH.CENTER)

    data_cell = hdr_tbl.cell(0, 2)
    set_cell_bg(data_cell, 'D5E8F0')
    data_cell.paragraphs[0].clear()
    for line, bold in [
        (f"Código: {hdr_codigo or '____________'}", False),
        (f"Versión: {hdr_version}", False),
        (f"Fecha: {date.today().strftime('%d/%m/%Y')}", False),
    ]:
        p = data_cell.add_paragraph()
        run = p.add_run(line)
        run.bold = bold
        run.font.size = Pt(8)
        run.font.name = 'Arial'

    # Fila 1: subgerencia + línea
    subger_cell = hdr_tbl.cell(1, 0).merge(hdr_tbl.cell(1, 1))
    set_cell_bg(subger_cell, 'EBF3FB')
    hdr_cell_text(subger_cell, hdr_subger, bold=False, size=8, color='1F3864')

    linea_cell = hdr_tbl.cell(1, 2)
    set_cell_bg(linea_cell, 'EBF3FB')
    hdr_cell_text(linea_cell, hdr_linea or '', bold=False, size=8, color='1F3864', align=WD_ALIGN_PARAGRAPH.CENTER)

    # ── PIE DE PÁGINA ──
    footer = section.footer
    for p in footer.paragraphs:
        p.clear()
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = fp.add_run(f"Generado el {date.today().strftime('%d/%m/%Y')}  —  Pág. ")
    run_f.font.size = Pt(8)
    run_f.font.name = 'Arial'
    run_f.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    # Número de página automático
    fld = OxmlElement('w:fldChar')
    fld.set(qn('w:fldCharType'), 'begin')
    fp.runs[-1]._r.append(fld)
    instrText = OxmlElement('w:instrText')
    instrText.text = 'PAGE'
    fld2 = OxmlElement('w:fldChar')
    fld2.set(qn('w:fldCharType'), 'end')
    fp.add_run()._r.extend([fld, instrText, fld2])

    # ── TÍTULO ──
    doc.add_paragraph()
    titulo = doc.add_heading('Informe de Fallas y Desvíos', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titulo.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    doc.add_paragraph(
        f"Período analizado: {df['Mes'].dropna().nunique()} meses  "
        f"|  Vehículos inspeccionados: {df['Vehiculo'].dropna().nunique()}  "
        f"|  Total observaciones: {len(df)}"
    )
    doc.add_paragraph()

    def add_section(title):
        h = doc.add_heading(title, level=1)
        h.runs[0].font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

    def table_from_df(dataframe):
        """Tabla ajustada al ancho de página, columnas distribuidas proporcionalmente."""
        n_cols = len(dataframe.columns)
        col_w  = PAGE_W_DXA // n_cols
        t = doc.add_table(rows=1, cols=n_cols)
        t.style = 'Table Grid'
        t.autofit = False
        # Encabezado
        hdr_row = t.rows[0]
        for i, col in enumerate(dataframe.columns):
            cell = hdr_row.cells[i]
            cell.width = col_w
            cell.paragraphs[0].clear()
            run = cell.paragraphs[0].add_run(str(col))
            run.bold = True
            run.font.size = Pt(9)
            run.font.name = 'Arial'
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Fondo azul encabezado
            tc   = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd  = OxmlElement('w:shd')
            shd.set(qn('w:val'),   'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'),  '1F3864')
            tcPr.append(shd)
        # Filas de datos
        for row_idx, (_, row) in enumerate(dataframe.iterrows()):
            row_cells = t.add_row().cells
            fill = 'F2F2F2' if row_idx % 2 == 0 else 'FFFFFF'
            for i, val in enumerate(row):
                cell = row_cells[i]
                cell.width = col_w
                cell.paragraphs[0].clear()
                run = cell.paragraphs[0].add_run(str(val) if pd.notna(val) else '')
                run.font.size = Pt(9)
                run.font.name = 'Arial'
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                tc   = cell._tc
                tcPr = tc.get_or_add_tcPr()
                shd  = OxmlElement('w:shd')
                shd.set(qn('w:val'),   'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'),  fill)
                tcPr.append(shd)
        return t

    sec_num = [0]
    def next_sec(title):
        sec_num[0] += 1
        add_section(f"{sec_num[0]}. {title}")

    # Resumen criticidad
    if inc_crit:
        next_sec("Distribución por Criticidad")
        crit_df = df['CritAmpliado'].value_counts().reset_index()
        crit_df.columns = ['Criticidad', 'Cantidad']
        crit_df['Porcentaje'] = (crit_df['Cantidad'] / len(df) * 100).round(1).astype(str) + '%'
        table_from_df(crit_df)
        doc.add_paragraph()

    # Top fallas
    if inc_top15:
        next_sec("Top 15 Tipos de Falla")
        top_df = df['DescAgrupada'].value_counts().head(15).reset_index()
        top_df.columns = ['Descripción', 'Cantidad']
        table_from_df(top_df)
        doc.add_paragraph()

    # Desvíos y detalle — solo si el archivo tiene columnas de referencia
    tiene_desv = not df_out.empty and 'Desviacion' in df_out.columns

    if tiene_desv and (inc_desvios or inc_detalle):
        if inc_desvios:
            next_sec("Resumen de Desvíos por Categoría")
            desv_sum = df_out.groupby('DescAgrupada').agg(
                Cantidad=('Desviacion','count'),
                Desv_Promedio=('Desviacion', lambda x: round(x.mean(),2)),
                Desv_Max_Abs=('Desviacion', lambda x: round(x.abs().max(),2)),
                Criticos=('Criticidad', lambda x: (x=='C').sum())
            ).reset_index().sort_values('Cantidad', ascending=False)
            desv_sum.columns = ['Categoría','Cantidad','Desvío Prom.','Desvío Máx.','Críticos']
            table_from_df(desv_sum)
            doc.add_paragraph()

        if inc_detalle:
            next_sec("Detalle de Observaciones Fuera de Parámetro")
        # Columnas base siempre presentes
        cols_det = ['Vehiculo','Mes','SistemaUnidad','Descripcion','Criticidad']
        rename_det = ['Vehículo','Mes','Sistema','Descripción','Crit.']
        # Agregar columnas de referencia solo si existen
        for col in ['RefMin','RefMax','ValorRelevado','Desviacion']:
            if col in df_out.columns:
                cols_det.insert(-1, col)
                rename_det.insert(-1, {'RefMin':'Ref Min','RefMax':'Ref Max',
                                       'ValorRelevado':'Relevado','Desviacion':'Desvío'}[col])
        det = df_out[cols_det].copy()
        det.columns = rename_det
        if 'Desvío' in det.columns:
            det['Desvío'] = det['Desvío'].round(2)
        if 'Relevado' in det.columns:
            det['Relevado'] = det['Relevado'].round(1)
        table_from_df(det)
    else:
        add_section("3. Observaciones sin valores de referencia")
        doc.add_paragraph(
            "Este archivo no contiene columnas de valores de referencia paramétrica. "
            "El análisis de desvíos no está disponible para este formato."
        )

    # ── CONCLUSIONES DINÁMICAS ──
    if inc_concl:
        next_sec("Conclusiones y Observaciones Generales")
        doc.add_paragraph(
        "Del análisis de la totalidad de las inspecciones estáticas realizadas "
        "durante el período se extraen las siguientes conclusiones:"
    )
    doc.add_paragraph()

    conclusiones = []
    n = len(df)

    # 1. Sistema con más observaciones
    sist_top_cod  = df['SistemaUnidad'].value_counts().index[0]
    sist_top_n    = int(df['SistemaUnidad'].value_counts().iloc[0])
    sist_top_lbl  = SISTEMA_LABELS.get(sist_top_cod, sist_top_cod)
    sist_top_pct  = round(sist_top_n / n * 100, 1)
    conclusiones.append(
        f"El sistema de {sist_top_lbl} ({sist_top_cod}) concentra la mayor cantidad de "
        f"observaciones con {sist_top_n} casos ({sist_top_pct}% del total), lo que indica "
        f"que es el área de mayor desgaste y atención requerida."
    )

    # 2. Vehículo o módulo con más fallas
    df['_unidad'] = df['Modulo'].apply(
        lambda x: None if (x is None or str(x).strip() in ('', '0', 'nan')) else str(x).strip()
    ).fillna(df['Vehiculo'].astype(str).str.strip())
    unidad_top   = df['_unidad'].value_counts().index[0]
    unidad_top_n = int(df['_unidad'].value_counts().iloc[0])
    conclusiones.append(
        f"La unidad con mayor cantidad de observaciones es {unidad_top} con {unidad_top_n} casos, "
        f"siendo candidata prioritaria para revisión integral."
    )

    # 3. Clasificación dominante
    if 'Clasificacion' in df.columns and df['Clasificacion'].notna().any():
        clasif_top   = df['Clasificacion'].value_counts().index[0]
        clasif_top_n = int(df['Clasificacion'].value_counts().iloc[0])
        clasif_top_pct = round(clasif_top_n / n * 100, 1)
        conclusiones.append(
            f"La clasificación de falla predominante es '{clasif_top}' con {clasif_top_n} casos "
            f"({clasif_top_pct}% del total), seguida por "
            f"'{df['Clasificacion'].value_counts().index[1] if len(df['Clasificacion'].value_counts()) > 1 else '-'}' "
            f"con {int(df['Clasificacion'].value_counts().iloc[1]) if len(df['Clasificacion'].value_counts()) > 1 else 0} casos."
        )

    # 4. % criticidad alta
    total_c = int((df['Criticidad'] == 'C').sum())
    total_r = int((df['Criticidad'] == 'R').sum())
    pct_cr  = round((total_c + total_r) / n * 100, 1)
    conclusiones.append(
        f"Las observaciones de criticidad alta (Crítico + Rechazado) representan el {pct_cr}% "
        f"del total ({total_c} críticas y {total_r} rechazadas), requiriendo atención prioritaria."
    )

    # 5. Falla más recurrente
    falla_top   = df['DescAgrupada'].value_counts().index[0]
    falla_top_n = int(df['DescAgrupada'].value_counts().iloc[0])
    falla_top_pct = round(falla_top_n / n * 100, 1)
    conclusiones.append(
        f"La falla más recurrente es '{falla_top}' con {falla_top_n} casos "
        f"({falla_top_pct}% del total de observaciones)."
    )

    # 6. Desvíos — solo si hay columnas de referencia
    if tiene_desv and not df_out.empty:
        # Grupo con mayor desvío promedio absoluto
        desv_grupo = df_out.groupby('DescAgrupada').agg(
            Cantidad=('Desviacion','count'),
            DesvProm=('Desviacion', lambda x: round(x.abs().mean(), 2)),
            DesvMax=('Desviacion',  lambda x: round(x.abs().max(),  2)),
        ).sort_values('DesvMax', ascending=False)

        grp_top     = desv_grupo.index[0]
        grp_max     = desv_grupo.iloc[0]['DesvMax']
        grp_prom    = desv_grupo.iloc[0]['DesvProm']
        grp_cant    = int(desv_grupo.iloc[0]['Cantidad'])

        # Fila individual con mayor desvío absoluto
        fila_max    = df_out.loc[df_out['Desviacion'].abs().idxmax()]
        veh_max     = fila_max['Vehiculo']
        desc_max    = str(fila_max['Descripcion'])[:60]
        ref_min     = fila_max.get('RefMin', '-')
        ref_max_val = fila_max.get('RefMax', '-')
        val_relev   = round(float(fila_max['ValorRelevado']), 1) if pd.notna(fila_max['ValorRelevado']) else '-'
        desv_max_v  = round(float(fila_max['Desviacion']), 2)
        direccion   = "por encima del máximo" if desv_max_v > 0 else "por debajo del mínimo"

        conclusiones.append(
            f"En cuanto a valores fuera de parámetro, la categoría con mayor desvío es "
            f"'{grp_top}' ({grp_cant} casos), con un desvío promedio de {grp_prom} unidades "
            f"y un máximo de {grp_max} unidades. El caso más extremo corresponde a la unidad "
            f"{veh_max} ('{desc_max}'), con un valor relevado de {val_relev} "
            f"frente a un rango de referencia de [{ref_min} – {ref_max_val}], "
            f"resultando {abs(desv_max_v)} unidades {direccion}."
        )

    # Escribir conclusiones numeradas
    for idx, texto in enumerate(conclusiones, 1):
        p = doc.add_paragraph(style='List Number')
        run = p.add_run(texto)
        run.font.size = Pt(11)
        run.font.name = 'Arial'

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 🔧 Análisis de Mantenimiento")
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
    st.caption("Análisis automático de desvíos · Material Rodante")


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

total_obs        = len(df)
total_veh        = df['Vehiculo'].dropna().nunique()
total_normal     = int((df['Criticidad'] == 'N').sum())
total_crit       = int((df['Criticidad'] == 'C').sum())
total_rech       = int((df['Criticidad'] == 'R').sum())
total_corregidas = int((df['Criticidad'] == 'O').sum())
total_nrc        = int((df['Criticidad'] == 'NRC').sum())
total_desv       = len(df_out)
pct_alta         = round((total_crit + total_rech) / total_obs * 100, 1) if total_obs > 0 else 0

# ─────────────────────────────────────────────
# KPIs — fila 1: volumen general
# ─────────────────────────────────────────────
st.markdown('<div class="section-header">Resumen del período</div>', unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)
c1.markdown(kpi("Total Observaciones", total_obs),          unsafe_allow_html=True)
c2.markdown(kpi("Vehículos Inspeccionados", total_veh),     unsafe_allow_html=True)
c3.markdown(kpi("Sin Observaciones", total_nrc),      unsafe_allow_html=True)
c4.markdown(kpi("% Criticidad Alta (Rechazo + Critico)", f"{pct_alta}%", "danger"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# KPIs — fila 2: desglose por criticidad
# ─────────────────────────────────────────────
d1, d2, d3, d4 = st.columns(4)
d1.markdown(kpi("Normales",          total_normal,     "success"), unsafe_allow_html=True)
d2.markdown(kpi("Corregidas en Inspección", total_corregidas, "success"), unsafe_allow_html=True)
d3.markdown(kpi("Críticas",          total_crit,       "danger"),  unsafe_allow_html=True)
d4.markdown(kpi("Rechazadas",        total_rech,       "warning"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# TABS PRINCIPALES
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊  Resúmenes",
    "📈  Gráficos",
    "🔍  Desvíos Detallados",
    "🧩  Análisis por Clasificación",
    "🔬  Explorador Libre",
    "📄  Exportar Informe"
])

PLOTLY_THEME = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(26,34,53,0.6)',
    font=dict(family='IBM Plex Sans', color='#b0bec5'),
    colorway=['#4fc3f7','#81d4fa','#ef5350','#ffa726','#66bb6a','#ab47bc','#26c6da'],
)
# Estilos de ejes reutilizables (separados para no generar claves duplicadas en update_layout)
AXIS_STYLE = dict(gridcolor='#1e2a3a', linecolor='#2a3a50')


# ── TAB 1: RESÚMENES ──
with tab1:
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### Criticidad")
        crit_df = df['CritAmpliado'].value_counts().reset_index()
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

    st.markdown("#### Top 15 Tipos de Falla")
    top15 = df['DescAgrupada'].value_counts().head(15).reset_index()
    top15.columns = ['Descripción Agrupada','Cantidad']
    top15['% del total'] = (top15['Cantidad'] / total_obs * 100).round(1).astype(str) + '%'
    st.dataframe(top15, use_container_width=True, hide_index=True)

    # ── TOP 10 POR TIPO DE MR ──
    st.markdown("---")
    st.markdown("#### Top 10 por Tipo de Material Rodante")

    MR_CONFIG = {
        'LOC' : {'label': 'Locomotoras (LOC)',        'servicios': ['LD', 'PERIFERICO', 'AMBA']},
        'CCRR': {'label': 'Coches Remolcados (CCRR)', 'servicios': ['LD', 'PERIFERICO']},
        'CCEE': {'label': 'Coches Eléctricos (CCEE)', 'servicios': []},
        'CCMM': {'label': 'Coche Motor (CCMM)',       'servicios': ['PERIFERICO', 'AMBA']},
    }

    def top10_unidad(df_base):
        """Aplica regla Modulo > Vehiculo y devuelve top 10."""
        df_base = df_base.copy()
        df_base['_unidad'] = df_base['Modulo'].apply(
            lambda x: None if (x is None or str(x).strip() in ('', '0', 'nan')) else str(x).strip()
        ).fillna(df_base['Vehiculo'].astype(str).str.strip())
        result = df_base['_unidad'].value_counts().head(10).reset_index()
        result.columns = ['Unidad', 'Observaciones']
        return result

    # Fila 1: LOC y CCRR
    row_t1, row_t2 = st.columns(2)

    with row_t1:
        cfg = MR_CONFIG['LOC']
        st.markdown(f"##### {cfg['label']}")
        df_mr = df[df['MR'] == 'LOC']
        if df_mr.empty:
            st.caption("NO SE INSPECCIONÓ ESTE TIPO DE MR")
        else:
            servicios_disp = sorted(df_mr['Servicio'].dropna().unique().tolist())
            sel_srv = st.multiselect("Servicio", servicios_disp,
                                     default=servicios_disp, key="srv_loc")
            df_fil = df_mr[df_mr['Servicio'].isin(sel_srv)] if sel_srv else df_mr
            st.dataframe(top10_unidad(df_fil), use_container_width=True, hide_index=True)

    with row_t2:
        cfg = MR_CONFIG['CCRR']
        st.markdown(f"##### {cfg['label']}")
        df_mr = df[df['MR'] == 'CCRR']
        if df_mr.empty:
            st.caption("NO SE INSPECCIONÓ ESTE TIPO DE MR")
        else:
            servicios_disp = sorted(df_mr['Servicio'].dropna().unique().tolist())
            sel_srv = st.multiselect("Servicio", servicios_disp,
                                     default=servicios_disp, key="srv_ccrr")
            df_fil = df_mr[df_mr['Servicio'].isin(sel_srv)] if sel_srv else df_mr
            st.dataframe(top10_unidad(df_fil), use_container_width=True, hide_index=True)

    # Fila 2: CCEE y CCMM
    row_t3, row_t4 = st.columns(2)

    with row_t3:
        cfg = MR_CONFIG['CCEE']
        st.markdown(f"##### {cfg['label']}")
        df_mr = df[df['MR'] == 'CCEE']
        if df_mr.empty:
            st.caption("NO SE INSPECCIONÓ ESTE TIPO DE MR")
        else:
            # CCEE solo tiene AMBA, no necesita filtro
            st.dataframe(top10_unidad(df_mr), use_container_width=True, hide_index=True)

    with row_t4:
        cfg = MR_CONFIG['CCMM']
        st.markdown(f"##### {cfg['label']}")
        df_mr = df[df['MR'] == 'CCMM']
        if df_mr.empty:
            st.caption("NO SE INSPECCIONÓ ESTE TIPO DE MR")
        else:
            servicios_disp = sorted(df_mr['Servicio'].dropna().unique().tolist())
            sel_srv = st.multiselect("Servicio", servicios_disp,
                                     default=servicios_disp, key="srv_ccmm")
            df_fil = df_mr[df_mr['Servicio'].isin(sel_srv)] if sel_srv else df_mr
            st.dataframe(top10_unidad(df_fil), use_container_width=True, hide_index=True)


# ── TAB 2: GRÁFICOS ──
with tab2:
    row1_l, row1_r = st.columns(2)

    # Torta criticidad
    with row1_l:
        st.markdown("#### Distribución por Criticidad")
        crit_counts = df['CritAmpliado'].value_counts()
        fig_pie = go.Figure(go.Pie(
            labels=crit_counts.index,
            values=crit_counts.values,
            hole=0.55,
            marker=dict(colors=['#4fc3f7','#ef5350','#ffa726','#66bb6a','#ab47bc']),
            textfont=dict(size=13)
        ))
        fig_pie.update_layout(**PLOTLY_THEME, margin=dict(t=10,b=10,l=10,r=10), height=300,
                              legend=dict(orientation='h', y=-0.1),
                              xaxis=AXIS_STYLE, yaxis=AXIS_STYLE)
        st.plotly_chart(fig_pie, use_container_width=True, key="chart_pie")

    # Barras sistemas
    with row1_r:
        st.markdown("#### Observaciones por Sistema")
        sist_counts = df['SistemaUnidad'].value_counts().reset_index()
        sist_counts.columns = ['Sistema','Cantidad']
        sist_counts['Label'] = sist_counts['Sistema'].map(SISTEMA_LABELS).fillna(sist_counts['Sistema'])
        fig_sist = px.bar(sist_counts, x='Cantidad', y='Label', orientation='h',
                          color='Cantidad', color_continuous_scale='Blues',
                          text='Cantidad')
        fig_sist.update_traces(textposition='outside', textfont=dict(color='#b0bec5', size=12))
        fig_sist.update_layout(**PLOTLY_THEME, margin=dict(t=10,b=10,r=80,l=10), height=300,
                               xaxis=dict(range=[0, sist_counts['Cantidad'].max() * 1.15], **AXIS_STYLE),
                               yaxis=dict(autorange='reversed', **AXIS_STYLE),
                               coloraxis_showscale=False)
        st.plotly_chart(fig_sist, use_container_width=True, key="chart_sist")

    # Evolución mensual + ratio obs/vehículo
    st.markdown("#### Evolución Mensual de Observaciones")
    monthly = []
    for m in MONTH_ORDER:
        df_m = df[df['Mes'].str.upper() == m]
        cnt  = len(df_m)
        vehs = df_m['Vehiculo'].dropna().nunique()
        if cnt > 0:
            monthly.append({
                'Mes': m,
                'Observaciones': cnt,
                'Vehículos': vehs,
                'Ratio Obs/Veh': round(cnt / vehs, 2) if vehs > 0 else 0
            })
    monthly_df = pd.DataFrame(monthly)
    if not monthly_df.empty:
        fig_line = go.Figure()

        # Barras de observaciones
        fig_line.add_trace(go.Bar(
            x=monthly_df['Mes'], y=monthly_df['Observaciones'],
            name='Observaciones',
            marker_color='rgba(79,195,247,0.3)',
            text=monthly_df['Observaciones'],
            textposition='outside',
            textfont=dict(color='#4fc3f7', size=11),
            yaxis='y'
        ))

        # Línea de vehículos inspeccionados
        fig_line.add_trace(go.Scatter(
            x=monthly_df['Mes'], y=monthly_df['Vehículos'],
            name='Vehículos inspeccionados',
            mode='lines+markers+text',
            line=dict(color='#66bb6a', width=2, dash='dot'),
            marker=dict(size=7, color='#66bb6a'),
            text=monthly_df['Vehículos'].astype(str),
            textposition='top center',
            textfont=dict(color='#ffa726', size=10),            
            yaxis='y'
        ))

        # Línea ratio obs/vehículo — eje secundario
        fig_line.add_trace(go.Scatter(
            x=monthly_df['Mes'], y=monthly_df['Ratio Obs/Veh'],
            name='Ratio Obs/Veh',
            mode='lines+markers+text',
            line=dict(color='#ffa726', width=2),
            marker=dict(size=8, color='#ffa726'),
            text=monthly_df['Ratio Obs/Veh'].astype(str),
            textposition='top center',
            textfont=dict(color='#ffa726', size=10),
            yaxis='y2'
        ))

        fig_line.update_layout(
            **PLOTLY_THEME,
            height=360,
            margin=dict(t=20, b=20, l=10, r=60),
            xaxis=AXIS_STYLE,
            yaxis=dict(title='Cantidad', **AXIS_STYLE),
            yaxis2=dict(title='Ratio Obs/Veh', overlaying='y', side='right',
                        gridcolor='#1e2a3a', linecolor='#2a3a50'),
            legend=dict(orientation='h', y=1.08),
            barmode='overlay'
        )
        st.plotly_chart(fig_line, use_container_width=True, key="chart_line")

    # Top fallas barras
    st.markdown("#### Top 15 Tipos de Falla")
    top15_plot = df['DescAgrupada'].value_counts().head(15).reset_index()
    top15_plot.columns = ['Falla','Cantidad']
    fig_top = px.bar(top15_plot, x='Cantidad', y='Falla', orientation='h',
                     color='Cantidad', color_continuous_scale='Blues_r',
                     text='Cantidad')
    fig_top.update_traces(textposition='outside', textfont=dict(color='#b0bec5', size=12))
    fig_top.update_layout(**PLOTLY_THEME, height=420, margin=dict(t=10,b=10,r=80,l=10),
                          xaxis=dict(range=[0, top15_plot['Cantidad'].max() * 1.15], **AXIS_STYLE),
                          yaxis=dict(autorange='reversed', **AXIS_STYLE),
                          coloraxis_showscale=False)
    st.plotly_chart(fig_top, use_container_width=True, key="chart_top")

    # Desvíos por categoría
    if not df_out.empty and 'Desviacion' in df_out.columns:
        st.markdown("#### Desvíos: Distribución por Categoría")
        st.caption("Barras = cantidad de casos · Líneas = desvío máx/prom/mín sobre eje derecho (mismas unidades que el parámetro)")
        desv_plot = df_out.groupby('DescAgrupada').agg(
            Cantidad  =('Desviacion', 'count'),
            Desv_Max  =('Desviacion', lambda x: round(x.abs().max(),  2)),
            Desv_Prom =('Desviacion', lambda x: round(x.abs().mean(), 2)),
            Desv_Min  =('Desviacion', lambda x: round(x.abs().min(),  2)),
        ).reset_index().sort_values('Cantidad', ascending=False)

        # Etiquetas combinadas para tooltips
        desv_plot['Etiqueta'] = desv_plot.apply(
            lambda r: f"Máx: {r['Desv_Max']}  |  Prom: {r['Desv_Prom']}  |  Mín: {r['Desv_Min']}", axis=1
        )

        fig_desv = go.Figure()

        # Barras cantidad
        fig_desv.add_trace(go.Bar(
            x=desv_plot['DescAgrupada'], y=desv_plot['Cantidad'],
            name='Cantidad de casos',
            marker_color='#4fc3f7',
            text=desv_plot['Cantidad'],
            textposition='outside',
            textfont=dict(color='#4fc3f7', size=11),
            yaxis='y'
        ))

        # Línea desvío máximo
        fig_desv.add_trace(go.Scatter(
            x=desv_plot['DescAgrupada'], y=desv_plot['Desv_Max'],
            name='Desvío Máx.',
            mode='lines+markers+text',
            line=dict(color='#ef5350', width=2),
            marker=dict(size=8, color='#ef5350'),
            text=desv_plot['Desv_Max'].astype(str),
            textposition='top center',
            textfont=dict(color='#ef5350', size=10),
            yaxis='y2'
        ))

        # Línea desvío promedio
        fig_desv.add_trace(go.Scatter(
            x=desv_plot['DescAgrupada'], y=desv_plot['Desv_Prom'],
            name='Desvío Prom.',
            mode='lines+markers+text',
            line=dict(color='#ffa726', width=2, dash='dot'),
            marker=dict(size=7, color='#ffa726'),
            text=desv_plot['Desv_Prom'].astype(str),
            textposition='bottom center',
            textfont=dict(color='#ffa726', size=10),
            yaxis='y2'
        ))

        # Línea desvío mínimo
        fig_desv.add_trace(go.Scatter(
            x=desv_plot['DescAgrupada'], y=desv_plot['Desv_Min'],
            name='Desvío Mín.',
            mode='lines+markers+text',
            line=dict(color='#66bb6a', width=1, dash='dash'),
            marker=dict(size=6, color='#66bb6a'),
            text=desv_plot['Desv_Prom'].astype(str),
            textposition='bottom center',
            textfont=dict(color='#ffa726', size=10),
            yaxis='y2'
        ))

        fig_desv.update_layout(
            **PLOTLY_THEME, height=440, barmode='group',
            margin=dict(t=20, b=100, l=10, r=70),
            xaxis=dict(tickangle=-35, **AXIS_STYLE),
            yaxis=dict(title='Cantidad de casos', **AXIS_STYLE),
            yaxis2=dict(title='Desvío (unidades del parámetro)', overlaying='y', side='right',
                        gridcolor='#1e2a3a', linecolor='#2a3a50'),
            legend=dict(orientation='h', y=1.06)
        )
        st.plotly_chart(fig_desv, use_container_width=True, key="chart_desv")


# ── TAB 3: DESVÍOS DETALLADOS ──
with tab3:
    if df_out.empty or 'Desviacion' not in df_out.columns:
        st.info(
            "Este archivo no contiene columnas de valores de referencia (RefMin / RefMax / Relevado). "
            "El análisis de desvíos paramétricos no está disponible para este formato. "
            "Las observaciones cualitativas se muestran en el tab **Resúmenes**."
        )
    else:
        st.markdown("#### Filtros")
        fc1, fc2, fc3, fc4, fc5 = st.columns(5)

        meses_disponibles    = sorted(df_out['Mes'].dropna().unique().tolist())
        sistemas_disponibles = sorted(df_out['SistemaUnidad'].dropna().unique().tolist())
        criticas_opciones    = sorted(df_out['Criticidad'].dropna().unique().tolist())

        # Unidades (Modulo > Vehiculo)
        df_out['_unidad'] = df_out['Modulo'].apply(
            lambda x: None if (x is None or str(x).strip() in ('', '0', 'nan')) else str(x).strip()
        ).fillna(df_out['Vehiculo'].astype(str).str.strip())
        unidades_disponibles = sorted(df_out['_unidad'].dropna().unique().tolist())

        # Descripciones (equivale a rango ref / tipo de medición)
        desc_disponibles = sorted(df_out['DescAgrupada'].dropna().unique().tolist())

        with fc1:
            sel_mes = st.multiselect("Mes", meses_disponibles, default=meses_disponibles,
                                     placeholder="Todos los meses", key="f3_mes")
        with fc2:
            sel_sist = st.multiselect("Sistema", sistemas_disponibles, default=sistemas_disponibles,
                                      placeholder="Todos los sistemas", key="f3_sist")
        with fc3:
            sel_crit = st.multiselect("Criticidad", criticas_opciones, default=criticas_opciones,
                                      placeholder="Todas", key="f3_crit")
        with fc4:
            sel_unidad = st.multiselect("Vehículo / Módulo", unidades_disponibles,
                                        default=unidades_disponibles, placeholder="Todas las unidades",
                                        key="f3_unidad")
        with fc5:
            sel_desc = st.multiselect("Tipo de medición", desc_disponibles,
                                      default=desc_disponibles, placeholder="Todas",
                                      key="f3_desc")

        df_filtrado = df_out[
            df_out['Mes'].isin(sel_mes) &
            df_out['SistemaUnidad'].isin(sel_sist) &
            df_out['Criticidad'].isin(sel_crit) &
            df_out['_unidad'].isin(sel_unidad) &
            df_out['DescAgrupada'].isin(sel_desc)
        ].copy()

        def ref_str(row):
            mn = row['RefMin_num'] if pd.notna(row.get('RefMin_num')) else None
            mx = row['RefMax_num'] if pd.notna(row.get('RefMax_num')) else None
            if mn is not None and mx is not None: return f"{mn} - {mx}"
            if mn is not None: return f">= {mn}"
            if mx is not None: return f"<= {mx}"
            return "-"

        df_filtrado['RefMin_num'] = pd.to_numeric(df_filtrado.get('RefMin'), errors='coerce')
        df_filtrado['RefMax_num'] = pd.to_numeric(df_filtrado.get('RefMax'), errors='coerce')
        df_filtrado['Rango Ref.'] = df_filtrado.apply(ref_str, axis=1)
        df_filtrado['Desvio']     = df_filtrado['Desviacion'].round(2)
        df_filtrado['Relevado']   = df_filtrado['ValorRelevado'].round(1)

        st.markdown(f"**{len(df_filtrado)}** observaciones fuera de parámetro")

        st.markdown("##### Resumen por Categoría")
        resumen = df_filtrado.groupby('DescAgrupada').agg(
            Cantidad=('Desvio','count'),
            Desv_Promedio=('Desvio', lambda x: round(x.mean(),2)),
            Desv_Max=('Desvio', lambda x: round(x.abs().max(),2)),
            Criticos=('Criticidad', lambda x: (x=='C').sum())
        ).reset_index().sort_values('Cantidad', ascending=False)
        resumen.columns = ['Categoria','Cantidad','Desvio Prom.','Desvio Max. (abs)','Criticos']
        st.dataframe(resumen, use_container_width=True, hide_index=True)

        st.markdown("##### Detalle completo")
        cols_show  = ['Vehiculo','Mes','SistemaUnidad','Descripcion','Rango Ref.','Relevado','Desvio','Criticidad']
        rename_map = {'Vehiculo':'Vehiculo','SistemaUnidad':'Sistema','Descripcion':'Descripcion'}
        tabla_det  = df_filtrado[cols_show].rename(columns=rename_map).reset_index(drop=True)

        st.dataframe(
            tabla_det,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Desvio":    st.column_config.NumberColumn(format="%.2f"),
                "Relevado":  st.column_config.NumberColumn(format="%.1f"),
                "Criticidad": st.column_config.TextColumn(),
            }
        )




# ── TAB 4: ANÁLISIS POR CLASIFICACIÓN ──
with tab4:
    tiene_clasif = 'Clasificacion' in df.columns and df['Clasificacion'].notna().any()

    if not tiene_clasif:
        st.info("Este archivo no tiene columna 'Clasificacion'. Agregala en el Excel para habilitar este análisis.")
    else:
        # ── Tabla cruzada Sistema × Criticidad ──
        st.markdown("#### Tabla Cruzada: Sistema × Criticidad")
        st.caption("Cantidad de observaciones por sistema y nivel de criticidad")

        pivot = df.pivot_table(
            index='CritAmpliado',
            columns='SistemaUnidad',
            aggfunc='size',
            fill_value=0
        )
        # Ordenar filas por criticidad
        orden_crit = ['Rechazado','Critico','Normal','Corregida']
        pivot = pivot.reindex([r for r in orden_crit if r in pivot.index])
        st.dataframe(pivot, use_container_width=True)

        st.markdown("---")

        # ── 3 Paretos por Clasificacion ──
        st.markdown("#### Pareto por Clasificación de Falla")
        st.caption("Cada gráfico muestra qué sistemas concentran el 80% de las fallas en cada categoría")

        CLASIF_COLORS = {
            'Fuera de rango':       '#ef5350',
            'Ausencia de elementos':'#ffa726',
            'Mal estado':           '#4fc3f7',
        }

        categorias = ['Fuera de rango', 'Ausencia de elementos', 'Mal estado']
        cols_pareto = st.columns(3)

        for i, cat in enumerate(categorias):
            df_cat = df[df['Clasificacion'].str.strip().str.lower() == cat.lower()].copy()
            with cols_pareto[i]:
                st.markdown(f"##### {cat}")
                if df_cat.empty:
                    st.caption("Sin datos")
                    continue

                # Contar por sistema y calcular % acumulado
                conteo = df_cat['SistemaUnidad'].value_counts().reset_index()
                conteo.columns = ['Sistema', 'Cantidad']
                conteo['Label'] = conteo['Sistema'].map(SISTEMA_LABELS).fillna(conteo['Sistema'])
                conteo['%_acum'] = (conteo['Cantidad'].cumsum() / conteo['Cantidad'].sum() * 100).round(1)

                color = CLASIF_COLORS.get(cat, '#4fc3f7')

                fig_p = go.Figure()

                # Barras
                fig_p.add_trace(go.Bar(
                    x=conteo['Label'],
                    y=conteo['Cantidad'],
                    name='Cantidad',
                    marker_color=color,
                    text=conteo['Cantidad'],
                    textposition='outside',
                    textfont=dict(color='#e8eaf0', size=11),
                    yaxis='y'
                ))

                # Línea acumulada
                fig_p.add_trace(go.Scatter(
                    x=conteo['Label'],
                    y=conteo['%_acum'],
                    name='% Acum.',
                    mode='lines+markers+text',
                    line=dict(color='#ffffff', width=2),
                    marker=dict(size=6, color='#ffffff'),
                    text=[f"{v}%" for v in conteo['%_acum']],
                    textposition='top center',
                    textfont=dict(color='#ffffff', size=10),
                    yaxis='y2'
                ))

                # Línea de referencia 80%
                fig_p.add_hline(
                    y=80, line_dash='dash',
                    line_color='#ffa726', line_width=1,
                    annotation_text='80%',
                    annotation_font_color='#ffa726',
                    yref='y2'
                )

                fig_p.update_layout(
                    **PLOTLY_THEME,
                    height=350,
                    margin=dict(t=30, b=60, l=10, r=50),
                    showlegend=False,
                    xaxis=dict(tickangle=-35, **AXIS_STYLE),
                    yaxis=dict(title='Cantidad', **AXIS_STYLE),
                    yaxis2=dict(
                        title='% Acumulado',
                        overlaying='y', side='right',
                        range=[0, 110],
                        gridcolor='#1e2a3a', linecolor='#2a3a50'
                    ),
                )
                st.plotly_chart(fig_p, use_container_width=True, key=f"pareto_{i}")

        # ── Últimas observaciones de rechazo ──
        st.markdown("---")
        st.markdown("#### Últimas Observaciones de Rechazo")
        st.caption("Observaciones con criticidad R ordenadas por fecha descendente")

        df_rech = df[df['Criticidad'] == 'R'].copy()
        if df_rech.empty:
            st.info("No hay observaciones de rechazo en este archivo.")
        else:
            df_rech['Fecha'] = pd.to_datetime(df_rech['Fecha'], errors='coerce')
            df_rech = df_rech.sort_values('Fecha', ascending=False)
            cols_rech = ['Fecha','Modulo','Vehiculo','SistemaUnidad','Descripcion','CritAmpliado']
            cols_rech = [c for c in cols_rech if c in df_rech.columns]
            tabla_rech = df_rech[cols_rech].head(15).copy()
            tabla_rech['Fecha'] = tabla_rech['Fecha'].dt.strftime('%d/%m/%Y')
            tabla_rech.columns = [{'SistemaUnidad':'Sistema','CritAmpliado':'Criticidad',
                                    'Descripcion':'Descripción'}.get(c,c) for c in tabla_rech.columns]
            st.dataframe(tabla_rech, use_container_width=True, hide_index=True)


# ── TAB 5: EXPLORADOR LIBRE ──
with tab5:
    st.markdown("#### Explorador de Datos")
    st.caption("Seleccioná cualquier variable para los ejes y explorá relaciones entre ellas")

    # Columnas disponibles para explorar
    cols_explorar = {
        'Sistema':      'SistemaUnidad',
        'Criticidad':   'CritAmpliado',
        'Tipo MR':      'MR',
        'Modelo':       'Modelo',
        'Servicio':     'Servicio',
        'Mes':          'Mes',
        'Clasificación':'Clasificacion',
    }
    # Filtrar solo las que existen en el df
    cols_explorar = {k: v for k, v in cols_explorar.items() if v in df.columns}

    ex1, ex2, ex3 = st.columns(3)
    with ex1:
        eje_x = st.selectbox("Eje X (categorías)", list(cols_explorar.keys()), index=0, key="ex_x")
    with ex2:
        eje_color = st.selectbox("Color (segunda variable)", list(cols_explorar.keys()), index=1, key="ex_col")
    with ex3:
        tipo_graf = st.selectbox("Tipo de gráfico", ["Barras agrupadas", "Barras apiladas", "Barras apiladas %"], key="ex_tipo")

    col_x     = cols_explorar[eje_x]
    col_color = cols_explorar[eje_color]

    # Agrupar
    df_exp = df.groupby([col_x, col_color]).size().reset_index(name='Cantidad')

    # Construir kwargs sin barnorm para evitar TypeError
    bar_kwargs = dict(
        x=col_x, y='Cantidad', color=col_color,
        text_auto=True,
        labels={col_x: eje_x, col_color: eje_color, 'Cantidad': 'Observaciones'},
        color_discrete_sequence=['#4fc3f7','#ef5350','#ffa726','#66bb6a','#ab47bc','#26c6da','#81d4fa'],
    )
    if tipo_graf == "Barras apiladas %":
        bar_kwargs['barmode']  = 'stack'
        bar_kwargs['barnorm']  = 'percent'
    elif tipo_graf == "Barras apiladas":
        bar_kwargs['barmode']  = 'stack'
    else:
        bar_kwargs['barmode']  = 'group'
    fig_exp = px.bar(df_exp, **bar_kwargs)
    fig_exp.update_traces(textfont=dict(size=11, color='white'), textposition='inside')
    fig_exp.update_layout(
        **PLOTLY_THEME,
        height=450,
        margin=dict(t=20, b=80, l=10, r=10),
        xaxis=dict(tickangle=-30, **AXIS_STYLE),
        yaxis=AXIS_STYLE,
        legend=dict(orientation='h', y=-0.25),
    )
    st.plotly_chart(fig_exp, use_container_width=True, key="chart_explorador")

    # Tabla de datos del gráfico
    with st.expander("Ver datos del gráfico"):
        pivot_exp = df_exp.pivot_table(
            index=col_x, columns=col_color, values='Cantidad', fill_value=0
        )
        st.dataframe(pivot_exp, use_container_width=True)


# ── TAB 6: EXPORTAR ──
with tab6:
    st.markdown("#### Configuración del Informe Word")

    # ── Encabezado del documento ──
    st.markdown("##### Datos del encabezado")
    hdr_col1, hdr_col2 = st.columns(2)
    with hdr_col1:
        hdr_codigo   = st.text_input("Código del informe", placeholder="Ej: SGBV-INF-2025-001")
        hdr_version  = st.text_input("Versión", value="v1.0")
        hdr_linea    = st.text_input("Línea / Contrato", placeholder="Ej: Línea San Martín — 3-LA")
    with hdr_col2:
        logo_file    = st.file_uploader("Banner / Logo (JPG o PNG)", type=["jpg","jpeg","png"],
                                         help="Imagen que aparece en el encabezado de todas las páginas")
        hdr_subger   = st.text_input("Subgerencia", value="Sub Gerencia de Programación y Seguimiento de Mantenimiento de Material Rodante (SPySM)")

    st.markdown("---")

    # ── Secciones a incluir ──
    st.markdown("##### Secciones del informe")
    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        inc_crit     = st.checkbox("Distribución por Criticidad",     value=True)
        inc_top15    = st.checkbox("Top 15 Tipos de Falla",           value=True)
    with sc2:
        inc_desvios  = st.checkbox("Tabla de Desvíos",                value=True)
        inc_detalle  = st.checkbox("Detalle fuera de parámetro",      value=True)
    with sc3:
        inc_graficos = st.checkbox("Gráficos (requiere kaleido)",     value=False)
        inc_concl    = st.checkbox("Conclusiones automáticas",        value=True)

    st.markdown("---")

    if st.button("📄  Generar y descargar informe Word", type="primary"):
        logo_bytes = logo_file.read() if logo_file else None
        config = dict(
            codigo=hdr_codigo, version=hdr_version, linea=hdr_linea,
            subger=hdr_subger, logo=logo_bytes,
            inc_crit=inc_crit, inc_top15=inc_top15,
            inc_desvios=inc_desvios, inc_detalle=inc_detalle,
            inc_graficos=inc_graficos, inc_concl=inc_concl,
        )
        with st.spinner("Generando documento..."):
            word_bytes = generar_word(df, df_out, config)

        st.success("✅ Informe generado correctamente.")
        linea_safe = (hdr_linea or "Informe").replace(" ","_").replace("/","_")[:30]
        st.download_button(
            label="⬇️  Descargar .docx",
            data=word_bytes,
            file_name=f"Informe_{linea_safe}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
