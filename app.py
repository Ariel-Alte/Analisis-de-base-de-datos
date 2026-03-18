"""
App Streamlit — Análisis de Informes de Mantenimiento
======================================================
Instalación (una sola vez):
    pip install streamlit plotly openpyxl pandas numpy python-docx matplotlib kaleido

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
        cols = [
            'Mes','Responsable','Contrato','Linea','Vehiculo','Modulo','MR',
            'Modelo','Servicio','Fecha','NroInforme','SistemaUnidad',
            'SistemaAmpliado','Item1','Item2','Descripcion',
            'Criticidad','DescAgrupada','CritAmpliado','CodItem',
            'FechaReInsp','NroReInsp','SistUnitReInsp','SistAmpReInsp',
            'ItemsReInsp','DescReInsp','CritReInsp','DescAgrupReInsp','CodReInsp',
            'Clasificacion'
        ]

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

    if 'Mes' in df.columns:
        df['Mes'] = df['Mes'].astype(str).str.strip().str.upper()

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
        df_out = pd.DataFrame(columns=list(df.columns) + ['ValorRelevado','Desviacion'])

    return df, df_out


def kpi(label, value, variant="default"):
    return f"""
    <div class="kpi-card {variant}">
        <div class="kpi-value">{value}</div>
        <div class="kpi-label">{label}</div>
    </div>"""


# ─────────────────────────────────────────────
# EXPORTAR GRÁFICOS — con fallback robusto
# ─────────────────────────────────────────────

def fig_to_png_robust(fig, width=700, height=350):
    """
    Convierte una figura Plotly a bytes PNG.
    Intenta kaleido primero, luego orca, luego matplotlib como fallback completo.
    """
    # Intento 1: kaleido
    try:
        png_bytes = fig.to_image(format='png', width=width, height=height, engine='kaleido')
        if png_bytes and len(png_bytes) > 100:
            return png_bytes
    except Exception:
        pass

    # Intento 2: orca
    try:
        png_bytes = fig.to_image(format='png', width=width, height=height)
        if png_bytes and len(png_bytes) > 100:
            return png_bytes
    except Exception:
        pass

    # Intento 3: matplotlib fallback completo con leyendas, titulos y etiquetas
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import matplotlib.cm as mcm
        import matplotlib.colors as mcolors
        import matplotlib.patches as mpatches

        BG     = '#FFFFFF'
        PLOTBG = '#F5F8FA'
        TXT    = '#333333'
        GRID   = '#E0E0E0'

        fig_mpl, ax = plt.subplots(figsize=(width / 100, height / 100), dpi=150)
        fig_mpl.patch.set_facecolor(BG)
        ax.set_facecolor(PLOTBG)

        # Detectar eje secundario
        has_y2 = any(getattr(t, 'yaxis', None) == 'y2' for t in fig.data)
        ax2 = None
        if has_y2:
            ax2 = ax.twinx()
            ax2.set_facecolor('none')

        bar_x_positions = None
        legend_handles  = []
        legend_labels   = []

        # Extraer titulos de ejes del layout de Plotly
        layout = fig.layout
        y1_title = y2_title = x_title = ''
        try:
            if layout.yaxis and layout.yaxis.title and layout.yaxis.title.text:
                y1_title = layout.yaxis.title.text
        except Exception:
            pass
        try:
            if layout.yaxis2 and layout.yaxis2.title and layout.yaxis2.title.text:
                y2_title = layout.yaxis2.title.text
        except Exception:
            pass
        try:
            if layout.xaxis and layout.xaxis.title and layout.xaxis.title.text:
                x_title = layout.xaxis.title.text
        except Exception:
            pass

        # Detectar hlines (linea 80% de Pareto)
        hline_y = None
        hline_ref = None
        try:
            if layout.shapes:
                for shape in layout.shapes:
                    if shape.type == 'line' and shape.y0 == shape.y1:
                        hline_y = shape.y0
                        hline_ref = getattr(shape, 'yref', 'y')
        except Exception:
            pass

        def _safe_color(mc_raw, fallback='#2E75B6'):
            """Extrae color seguro de marker.color (str, lista, array, tuple)."""
            if mc_raw is None:
                return fallback
            if isinstance(mc_raw, str):
                if mc_raw.startswith('rgba'):
                    parts = mc_raw.replace('rgba(', '').replace(')', '').split(',')
                    try:
                        return (int(parts[0])/255, int(parts[1])/255, int(parts[2])/255, float(parts[3]))
                    except Exception:
                        return fallback
                return mc_raw
            try:
                arr = list(mc_raw)
                if arr:
                    nums = []
                    for v in arr:
                        try: nums.append(float(v))
                        except: nums.append(0)
                    if nums:
                        norm = mcolors.Normalize(vmin=min(nums), vmax=max(nums))
                        cmap = mcm.get_cmap('Blues')
                        return [cmap(norm(v)) for v in nums]
            except Exception:
                pass
            return fallback

        # Recorrer trazas
        for trace in fig.data:
            target_ax = ax2 if (ax2 and getattr(trace, 'yaxis', None) == 'y2') else ax
            trace_name = getattr(trace, 'name', None) or ''

            # PIE
            if trace.type == 'pie':
                ax.remove()
                ax = fig_mpl.add_subplot(111)
                ax.set_facecolor(BG)
                colors_pie = ['#2E75B6', '#E74C3C', '#F39C12', '#27AE60', '#8E44AD', '#16A085']
                vals = list(trace.values) if trace.values is not None else []
                lbls = list(trace.labels) if trace.labels is not None else []
                if vals:
                    wedges, texts, autotexts = ax.pie(
                        vals, labels=lbls,
                        colors=colors_pie[:len(vals)],
                        autopct='%1.1f%%', pctdistance=0.75,
                        wedgeprops=dict(width=0.45), startangle=90,
                    )
                    for t in texts:
                        t.set_color(TXT); t.set_fontsize(9)
                    for t in autotexts:
                        t.set_color('white'); t.set_fontsize(8); t.set_fontweight('bold')
                    ax.legend(wedges, lbls, loc='center left', bbox_to_anchor=(1, 0.5),
                              fontsize=7, frameon=False, labelcolor=TXT)
                break

            # BAR
            elif trace.type == 'bar':
                mc_raw = None
                try:
                    mc_raw = trace.marker.color if (hasattr(trace, 'marker') and trace.marker) else None
                except Exception:
                    pass
                color = _safe_color(mc_raw)

                x_data = list(trace.x) if trace.x is not None else []
                y_data = list(trace.y) if trace.y is not None else []
                orientation = getattr(trace, 'orientation', None)

                if orientation == 'h' and y_data and x_data:
                    positions = list(range(len(y_data)))
                    target_ax.barh(positions, x_data, color=color, alpha=0.85, label=trace_name)
                    target_ax.set_yticks(positions)
                    target_ax.set_yticklabels([str(l) for l in y_data], fontsize=7, color=TXT)
                    target_ax.invert_yaxis()
                    for j, v in enumerate(x_data):
                        try:
                            fv = float(v)
                            target_ax.text(fv + max(float(xx) for xx in x_data) * 0.015, j,
                                           str(int(fv)) if fv == int(fv) else f"{fv:.1f}",
                                           va='center', ha='left', fontsize=7, color=TXT)
                        except Exception:
                            pass
                elif x_data and y_data:
                    positions = list(range(len(x_data)))
                    bar_x_positions = positions
                    target_ax.bar(positions, y_data, color=color, alpha=0.85, label=trace_name)
                    target_ax.set_xticks(positions)
                    target_ax.set_xticklabels([str(l) for l in x_data],
                                              fontsize=6, color=TXT, rotation=35, ha='right')
                    for j, v in enumerate(y_data):
                        try:
                            fv = float(v)
                            target_ax.text(j, fv + max(float(yy) for yy in y_data) * 0.02,
                                           str(int(fv)) if fv == int(fv) else f"{fv:.1f}",
                                           ha='center', va='bottom', fontsize=7, color=TXT)
                        except Exception:
                            pass

                if trace_name:
                    c = color if isinstance(color, str) else (color[0] if isinstance(color, list) else '#2E75B6')
                    legend_handles.append(mpatches.Patch(color=c, alpha=0.85))
                    legend_labels.append(trace_name)

            # SCATTER
            elif trace.type == 'scatter':
                line_color = '#F39C12'
                try:
                    if trace.line and trace.line.color:
                        line_color = trace.line.color
                except Exception:
                    pass

                x_data = list(trace.x) if trace.x is not None else []
                y_data = list(trace.y) if trace.y is not None else []

                if x_data and y_data:
                    if bar_x_positions is not None and len(x_data) == len(bar_x_positions):
                        x_plot = list(bar_x_positions)
                    else:
                        x_plot = list(range(len(x_data)))

                    dash = None
                    try: dash = trace.line.dash if trace.line else None
                    except: pass
                    ls = '--' if dash in ('dot', 'dash', 'dashdot') else '-'

                    line_obj, = target_ax.plot(x_plot, y_data, color=line_color,
                                               marker='o', markersize=5, linewidth=2,
                                               linestyle=ls, label=trace_name)
                    # Etiquetas de valor
                    y_max = max(float(yy) for yy in y_data) if y_data else 1
                    for j, v in enumerate(y_data):
                        try:
                            fv = float(v)
                            lbl = f"{fv:.1f}" if fv != int(fv) else str(int(fv))
                            if 'acum' in (y2_title or '').lower() or '%' in (y2_title or ''):
                                lbl += '%'
                            target_ax.text(x_plot[j], fv + y_max * 0.03, lbl,
                                           ha='center', va='bottom', fontsize=6,
                                           color=line_color, fontweight='bold')
                        except Exception:
                            pass

                    if trace_name:
                        legend_handles.append(line_obj)
                        legend_labels.append(trace_name)

        # Linea horizontal de referencia (80% Pareto)
        if hline_y is not None:
            ref_ax = ax2 if (ax2 and hline_ref and 'y2' in str(hline_ref)) else ax
            ref_ax.axhline(y=hline_y, color='#F39C12', linestyle='--', linewidth=1.2, alpha=0.8)
            ref_ax.text(0.02, hline_y, f'{int(hline_y)}%',
                        transform=ref_ax.get_yaxis_transform(),
                        color='#F39C12', fontsize=8, fontweight='bold', va='bottom')

        # Titulos de ejes
        if y1_title:
            ax.set_ylabel(y1_title, fontsize=9, color=TXT, fontweight='bold')
        if y2_title and ax2:
            ax2.set_ylabel(y2_title, fontsize=9, color=TXT, fontweight='bold')
        if x_title:
            ax.set_xlabel(x_title, fontsize=9, color=TXT, fontweight='bold')

        # Titulo del grafico
        try:
            if layout.title and layout.title.text:
                fig_mpl.suptitle(layout.title.text, fontsize=11, color=TXT, fontweight='bold', y=0.98)
        except Exception:
            pass

        # Estilo de ejes
        for a in [ax] + ([ax2] if ax2 else []):
            a.tick_params(colors=TXT, labelsize=7)
            for spine in a.spines.values():
                spine.set_color(GRID)
            a.grid(True, alpha=0.3, color=GRID, linestyle='-', linewidth=0.5)

        # Leyenda combinada
        if legend_handles and legend_labels:
            ax.legend(legend_handles, legend_labels, loc='upper center',
                      bbox_to_anchor=(0.5, -0.18), ncol=min(len(legend_labels), 4),
                      fontsize=8, frameon=True, facecolor=BG, edgecolor=GRID, labelcolor=TXT)

        buf = io.BytesIO()
        fig_mpl.tight_layout(rect=[0, 0.05, 1, 0.95])
        fig_mpl.savefig(buf, format='png', facecolor=BG,
                        edgecolor='none', bbox_inches='tight', dpi=150)
        plt.close(fig_mpl)
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        try:
            st.warning(f"\u26a0\ufe0f No se pudo generar grafico. Instala kaleido (pip install kaleido) "
                       f"para mejor calidad. Error: {e}")
        except Exception:
            pass
        return None


# ─────────────────────────────────────────────
# GENERACIÓN DE WORD (descarga)
# ─────────────────────────────────────────────

def generar_word(df, df_out, config=None):
    from docx.shared import Cm, Mm, Twips, Emu
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH as WALIGN
    from datetime import date

    if config is None:
        config = {}

    inc_crit    = config.get("inc_crit",    True)
    inc_top15   = config.get("inc_top15",   True)
    inc_desvios = config.get("inc_desvios", True)
    inc_detalle = config.get("inc_detalle", True)
    inc_concl   = config.get("inc_concl",   True)
    graficos    = config.get("graficos",    {})
    hdr_codigo  = config.get("codigo",  "") or ""
    hdr_version = config.get("version", "v1.0") or "v1.0"
    hdr_linea   = config.get("linea",   "") or ""
    hdr_subger  = config.get("subger",  "Sub Gerencia de Programación y Seguimiento de Mantenimiento de Material Rodante (SPySM)") or ""
    logo_bytes  = config.get("logo",    None)

    doc = Document()

    # ── Página A4, márgenes 15mm ──
    section = doc.sections[0]
    section.page_width      = Mm(210)
    section.page_height     = Mm(297)
    section.left_margin     = Mm(15)
    section.right_margin    = Mm(15)
    section.top_margin      = Mm(35)
    section.bottom_margin   = Mm(15)
    section.header_distance = Mm(5)

    # Ancho de contenido: 210 - 30 = 180mm → DXA (1mm ≈ 56.69 DXA)
    PAGE_W = int(180 * 56.69)  # ~10204 DXA

    # ── Estilos base ──
    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Arial"
    style_normal.font.size = Pt(10)
    style_normal.paragraph_format.space_after = Pt(0)
    style_normal.paragraph_format.space_before = Pt(0)

    # ── XML helpers ──
    def set_shd(cell, hex_color):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for old in tcPr.findall(qn("w:shd")):
            tcPr.remove(old)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  hex_color)
        tcPr.append(shd)

    def set_cell_w(cell, w_dxa):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for old in tcPr.findall(qn("w:tcW")):
            tcPr.remove(old)
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:w"),    str(w_dxa))
        tcW.set(qn("w:type"), "dxa")
        tcPr.insert(0, tcW)

    def set_tbl_w(tbl, w_dxa):
        tblPr = tbl._tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            tbl._tbl.insert(0, tblPr)
        for old in tblPr.findall(qn("w:tblW")):
            tblPr.remove(old)
        tblW = OxmlElement("w:tblW")
        tblW.set(qn("w:w"),    str(w_dxa))
        tblW.set(qn("w:type"), "dxa")
        tblPr.append(tblW)

    def set_tbl_layout_fixed(tbl):
        """Fuerza layout fijo para que Word respete los anchos de columna."""
        tblPr = tbl._tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            tbl._tbl.insert(0, tblPr)
        for old in tblPr.findall(qn("w:tblLayout")):
            tblPr.remove(old)
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tblPr.append(layout)

    def set_tbl_grid(tbl, widths):
        """
        Crea el elemento <w:tblGrid> con <w:gridCol> para cada columna.
        CRÍTICO: Sin este elemento Word ignora los anchos de celda y usa autofit.
        """
        tbl_element = tbl._tbl
        # Eliminar tblGrid existente
        for old in tbl_element.findall(qn("w:tblGrid")):
            tbl_element.remove(old)
        tblGrid = OxmlElement("w:tblGrid")
        for w in widths:
            gridCol = OxmlElement("w:gridCol")
            gridCol.set(qn("w:w"), str(w))
            tblGrid.append(gridCol)
        # tblGrid debe ir justo después de tblPr
        tblPr = tbl_element.find(qn("w:tblPr"))
        if tblPr is not None:
            tblPr.addnext(tblGrid)
        else:
            tbl_element.insert(0, tblGrid)

    def disable_autofit_xml(tbl):
        """Desactiva autofit a nivel XML (más confiable que t.autofit = False)."""
        tblPr = tbl._tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            tbl._tbl.insert(0, tblPr)
        # Remover tblLook que puede interferir
        for old in tblPr.findall(qn("w:tblCellMar")):
            pass  # mantener márgenes globales si existen

    def set_cell_margins(cell, top=40, bottom=40, left=60, right=60):
        """Agrega márgenes internos a una celda para mejor legibilidad."""
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for old in tcPr.findall(qn("w:tcMar")):
            tcPr.remove(old)
        tcMar = OxmlElement("w:tcMar")
        for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:w"),    str(val))
            el.set(qn("w:type"), "dxa")
            tcMar.append(el)
        tcPr.append(tcMar)

    def set_cell_vertical_alignment(cell, align="center"):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for old in tcPr.findall(qn("w:vAlign")):
            tcPr.remove(old)
        vAlign = OxmlElement("w:vAlign")
        vAlign.set(qn("w:val"), align)
        tcPr.append(vAlign)

    def set_table_borders(tbl, color="B0B0B0", sz="4"):
        """Aplica bordes uniformes a toda la tabla."""
        tblPr = tbl._tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            tbl._tbl.insert(0, tblPr)
        for old in tblPr.findall(qn("w:tblBorders")):
            tblPr.remove(old)
        borders = OxmlElement("w:tblBorders")
        for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:val"),   "single")
            el.set(qn("w:sz"),    sz)
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)
            borders.append(el)
        tblPr.append(borders)

    def remove_borders(tbl):
        tblPr = tbl._tbl.tblPr
        if tblPr is None:
            return
        for old in tblPr.findall(qn("w:tblBorders")):
            tblPr.remove(old)
        b = OxmlElement("w:tblBorders")
        for side in ["top","left","bottom","right","insideH","insideV"]:
            s = OxmlElement(f"w:{side}")
            s.set(qn("w:val"),    "none")
            s.set(qn("w:sz"),     "0")
            s.set(qn("w:space"),  "0")
            s.set(qn("w:color"),  "auto")
            b.append(s)
        tblPr.append(b)

    def set_row_height(row, height_twips, rule="atLeast"):
        tr   = row._tr
        trPr = tr.get_or_add_trPr()
        for old in trPr.findall(qn("w:trHeight")):
            trPr.remove(old)
        trH = OxmlElement("w:trHeight")
        trH.set(qn("w:val"),   str(height_twips))
        trH.set(qn("w:hRule"), rule)
        trPr.append(trH)

    def cell_text(cell, text, bold=False, size=9, color="000000",
                  align=WD_ALIGN_PARAGRAPH.CENTER, italic=False):
        p = cell.paragraphs[0]
        p.clear()
        p.alignment = align
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after  = Pt(1)
        r = p.add_run(str(text) if pd.notna(text) else "")
        r.bold   = bold
        r.italic = italic
        r.font.size = Pt(size)
        r.font.name = "Arial"
        r.font.color.rgb = RGBColor.from_string(color)

    # ────────────────────────────────────────
    # ENCABEZADO
    # ────────────────────────────────────────
    header = section.header

    for p in list(header.paragraphs):
        p._element.getparent().remove(p._element)

    W_LEFT  = int(PAGE_W * 0.60)
    W_RIGHT = PAGE_W - W_LEFT

    htbl = header.add_table(rows=1, cols=2, width=Mm(180))
    htbl.style = "Table Grid"
    set_tbl_w(htbl, PAGE_W)
    set_tbl_layout_fixed(htbl)
    set_tbl_grid(htbl, [W_LEFT, W_RIGHT])
    remove_borders(htbl)

    # Altura de fila ~2.5cm = 1418 twips
    set_row_height(htbl.rows[0], int(2.5 * 567))

    c_left  = htbl.cell(0, 0)
    c_right = htbl.cell(0, 1)
    set_cell_w(c_left,  W_LEFT)
    set_cell_w(c_right, W_RIGHT)
    set_cell_vertical_alignment(c_left, "center")
    set_cell_vertical_alignment(c_right, "center")

    # Celda izquierda — logo o texto
    if logo_bytes:
        set_shd(c_left, "FFFFFF")
        c_left.paragraphs[0].clear()
        c_left.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_img = c_left.paragraphs[0].add_run()
        run_img.add_picture(io.BytesIO(logo_bytes), height=Cm(2.2))
    else:
        set_shd(c_left, "1F3864")
        cell_text(c_left, "TRENES ARGENTINOS — PISE",
                  bold=True, size=13, color="FFFFFF",
                  align=WD_ALIGN_PARAGRAPH.CENTER)

    # Celda derecha — sub-tabla con campos
    set_shd(c_right, "EBF3FB")
    for p in list(c_right.paragraphs):
        p._element.getparent().remove(p._element)

    fields = [
        ("Código:",      hdr_codigo  or "___________"),
        ("Versión:",     hdr_version),
        ("Fecha:",       date.today().strftime("%d/%m/%Y")),
        ("Línea:",       hdr_linea   or "___________"),
        ("Subgerencia:", hdr_subger),
    ]

    sub = c_right.add_table(rows=len(fields), cols=2)
    remove_borders(sub)
    sub_w_label = int(W_RIGHT * 0.35)
    sub_w_value = W_RIGHT - sub_w_label
    set_tbl_w(sub, W_RIGHT)
    set_tbl_layout_fixed(sub)
    set_tbl_grid(sub, [sub_w_label, sub_w_value])

    for i, (lbl, val) in enumerate(fields):
        last = (i == len(fields) - 1)
        if last:
            merged = sub.cell(i, 0).merge(sub.cell(i, 1))
            set_cell_w(merged, W_RIGHT)
            set_shd(merged, "D5E8F0")
            cell_text(merged, f"{lbl} {val}",
                      bold=False, size=7, color="1F3864",
                      align=WD_ALIGN_PARAGRAPH.LEFT, italic=True)
            set_row_height(sub.rows[i], int(0.55 * 567), "exact")
        else:
            lc = sub.cell(i, 0)
            vc = sub.cell(i, 1)
            set_cell_w(lc, sub_w_label)
            set_cell_w(vc, sub_w_value)
            set_shd(lc, "EBF3FB")
            set_shd(vc, "EBF3FB")
            cell_text(lc, lbl, bold=True,  size=8, color="1F3864", align=WD_ALIGN_PARAGRAPH.LEFT)
            cell_text(vc, val, bold=False, size=8, color="333333", align=WD_ALIGN_PARAGRAPH.LEFT)
            set_row_height(sub.rows[i], int(0.42 * 567), "exact")

    # ── PIE DE PÁGINA ──
    footer = section.footer
    for p in list(footer.paragraphs):
        p.clear()
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_f = fp.add_run(f"SPySM  —  {date.today().strftime('%d/%m/%Y')}  —  Pág. ")
    r_f.font.size = Pt(8)
    r_f.font.name = "Arial"
    r_f.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    # ────────────────────────────────────────
    # TABLE HELPER — formato profesional
    # ────────────────────────────────────────
    def table_from_df(dataframe):
        n = len(dataframe.columns)
        col_names = list(dataframe.columns)

        # Calcular anchos: columnas "anchas" (descripciones) vs estándar
        wide_cols = {"Descripción", "Descripcion", "Categoría", "Categoria",
                     "Descripción Agrupada", "Falla"}

        wide_count = sum(1 for c in col_names if c in wide_cols)
        narrow_count = n - wide_count

        if wide_count > 0 and narrow_count > 0:
            narrow_w = max(900, int(PAGE_W * 0.08))
            wide_w = (PAGE_W - narrow_w * narrow_count) // wide_count
        elif wide_count > 0:
            wide_w = PAGE_W // wide_count
            narrow_w = wide_w
        else:
            narrow_w = PAGE_W // n
            wide_w = narrow_w

        widths = []
        for c in col_names:
            widths.append(wide_w if c in wide_cols else narrow_w)

        # Normalizar para que sumen exactamente PAGE_W
        total = sum(widths)
        if total != PAGE_W:
            widths[-1] += PAGE_W - total

        t = doc.add_table(rows=1, cols=n)
        t.style   = "Table Grid"
        t.autofit = False

        # Configurar tabla: ancho fijo, bordes, layout, grilla
        set_tbl_w(t, PAGE_W)
        set_tbl_layout_fixed(t)
        set_tbl_grid(t, widths)  # CRÍTICO: define la grilla de columnas en el XML
        set_table_borders(t, color="A0A0A0", sz="4")

        # Encabezado
        for i, (cname, w) in enumerate(zip(col_names, widths)):
            cell = t.rows[0].cells[i]
            set_cell_w(cell, w)
            set_shd(cell, "1F3864")
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            set_cell_vertical_alignment(cell, "center")
            cell_text(cell, cname, bold=True, size=8,
                      color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)

        set_row_height(t.rows[0], 400, "atLeast")

        # Datos con filas alternas
        for ri, (_, row) in enumerate(dataframe.iterrows()):
            fill = "EBF3FB" if ri % 2 == 0 else "FFFFFF"
            tr = t.add_row()
            set_row_height(tr, 320, "atLeast")
            for i, (val, w) in enumerate(zip(row, widths)):
                cell = tr.cells[i]
                set_cell_w(cell, w)
                set_shd(cell, fill)
                set_cell_margins(cell, top=40, bottom=40, left=80, right=80)
                set_cell_vertical_alignment(cell, "center")
                txt = "" if pd.isna(val) else str(val)
                # Descripciones alineadas a la izquierda
                col_name = col_names[i]
                text_align = WD_ALIGN_PARAGRAPH.LEFT if col_name in wide_cols else WD_ALIGN_PARAGRAPH.CENTER
                cell_text(cell, txt, size=8, align=text_align)
        return t

    # ── Numerador automático ──
    sec_num = [0]
    def next_sec(title):
        sec_num[0] += 1
        h = doc.add_heading(f"{sec_num[0]}. {title}", level=1)
        for run in h.runs:
            run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
            run.font.name = "Arial"

    # ────────────────────────────────────────
    # TÍTULO
    # ────────────────────────────────────────
    doc.add_paragraph()
    tit = doc.add_heading("Informe de Fallas y Desvíos", 0)
    tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in tit.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
        run.font.name = "Arial"

    p_meta = doc.add_paragraph(
        f"Período: {df['Mes'].dropna().nunique()} meses  "
        f"|  Vehículos: {df['Vehiculo'].dropna().nunique()}  "
        f"|  Total observaciones: {len(df)}"
    )
    p_meta.runs[0].font.size = Pt(10)
    doc.add_paragraph()

    # ────────────────────────────────────────
    # SECCIONES
    # ────────────────────────────────────────
    if inc_crit:
        next_sec("Distribución por Criticidad")
        d = df["CritAmpliado"].value_counts().reset_index()
        d.columns = ["Criticidad", "Cantidad"]
        d["Porcentaje"] = (d["Cantidad"] / len(df) * 100).round(1).astype(str) + "%"
        table_from_df(d)
        doc.add_paragraph()

    if inc_top15:
        next_sec("Top 15 Tipos de Falla")
        d = df["DescAgrupada"].value_counts().head(15).reset_index()
        d.columns = ["Descripción", "Cantidad"]
        d["% total"] = (d["Cantidad"] / len(df) * 100).round(1).astype(str) + "%"
        table_from_df(d)
        doc.add_paragraph()

    tiene_desv = not df_out.empty and "Desviacion" in df_out.columns

    if tiene_desv and inc_desvios:
        next_sec("Resumen de Desvíos por Categoría")
        d = df_out.groupby("DescAgrupada").agg(
            Cantidad     =("Desviacion", "count"),
            Desv_Promedio=("Desviacion", lambda x: round(x.mean(),   2)),
            Desv_Max_Abs =("Desviacion", lambda x: round(x.abs().max(), 2)),
            Criticos     =("Criticidad", lambda x: (x == "C").sum())
        ).reset_index().sort_values("Cantidad", ascending=False)
        d.columns = ["Categoría", "Cantidad", "Desvío Prom.", "Desvío Máx.", "Críticos"]
        table_from_df(d)
        doc.add_paragraph()

    if tiene_desv and inc_detalle:
        next_sec("Detalle — Observaciones Fuera de Parámetro")
        cols_d  = ["Vehiculo","Mes","SistemaUnidad","Descripcion","Criticidad"]
        rnames  = ["Vehículo","Mes","Sistema","Descripción","Crit."]
        for col in ["RefMin","RefMax","ValorRelevado","Desviacion"]:
            if col in df_out.columns:
                cols_d.insert(-1, col)
                rnames.insert(-1, {"RefMin":"Ref Min","RefMax":"Ref Max",
                                   "ValorRelevado":"Relevado","Desviacion":"Desvío"}[col])
        det = df_out[cols_d].copy()
        det.columns = rnames
        if "Desvío"   in det.columns: det["Desvío"]   = det["Desvío"].round(2)
        if "Relevado" in det.columns: det["Relevado"] = det["Relevado"].round(1)
        table_from_df(det)
        doc.add_paragraph()

    if not tiene_desv and (inc_desvios or inc_detalle):
        next_sec("Observaciones sin valores de referencia")
        p = doc.add_paragraph(
            "Este archivo no contiene columnas de valores de referencia paramétrica. "
            "El análisis de desvíos no está disponible para este formato."
        )
        p.runs[0].font.size = Pt(10)

    # ────────────────────────────────────────
    # CONCLUSIONES DINÁMICAS
    # ────────────────────────────────────────
    if inc_concl:
        next_sec("Conclusiones y Observaciones Generales")
        doc.add_paragraph(
            "Del análisis de la totalidad de las inspecciones estáticas realizadas "
            "durante el período se extraen las siguientes conclusiones:"
        )
        doc.add_paragraph()
        conclusiones = []
        n_tot = len(df)

        sv  = df["SistemaUnidad"].value_counts()
        conclusiones.append(
            f"El sistema de {SISTEMA_LABELS.get(sv.index[0], sv.index[0])} ({sv.index[0]}) "
            f"concentra la mayor cantidad de observaciones con {int(sv.iloc[0])} casos "
            f"({round(sv.iloc[0]/n_tot*100,1)}% del total)."
        )

        df["_unidad"] = df["Modulo"].apply(
            lambda x: None if (x is None or str(x).strip() in ("","0","nan")) else str(x).strip()
        ).fillna(df["Vehiculo"].astype(str).str.strip())
        uv = df["_unidad"].value_counts()
        conclusiones.append(
            f"La unidad con mayor cantidad de observaciones es {uv.index[0]} "
            f"con {int(uv.iloc[0])} casos, siendo candidata prioritaria para revisión integral."
        )

        if "Clasificacion" in df.columns and df["Clasificacion"].notna().any():
            cv = df["Clasificacion"].value_counts()
            conclusiones.append(
                f"La clasificación de falla predominante es '{cv.index[0]}' con {int(cv.iloc[0])} casos "
                f"({round(cv.iloc[0]/n_tot*100,1)}%), seguida por "
                f"'{cv.index[1] if len(cv)>1 else '-'}' con {int(cv.iloc[1]) if len(cv)>1 else 0} casos."
            )

        tc = int((df["Criticidad"]=="C").sum())
        tr = int((df["Criticidad"]=="R").sum())
        conclusiones.append(
            f"Las observaciones de criticidad alta (Crítico + Rechazado) representan el "
            f"{round((tc+tr)/n_tot*100,1)}% del total ({tc} críticas y {tr} rechazadas)."
        )

        fv = df["DescAgrupada"].value_counts()
        conclusiones.append(
            f"La falla más recurrente es '{fv.index[0]}' con {int(fv.iloc[0])} casos "
            f"({round(fv.iloc[0]/n_tot*100,1)}% del total)."
        )

        if tiene_desv and not df_out.empty:
            dg = df_out.groupby("DescAgrupada").agg(
                Cant=("Desviacion","count"),
                Prom=("Desviacion", lambda x: round(x.abs().mean(), 2)),
                Max =("Desviacion", lambda x: round(x.abs().max(),  2)),
            ).sort_values("Max", ascending=False)
            fm   = df_out.loc[df_out["Desviacion"].abs().idxmax()]
            dval = round(float(fm["Desviacion"]), 2)
            conclusiones.append(
                f"La categoría con mayor desvío paramétrico es '{dg.index[0]}' "
                f"({int(dg.iloc[0]['Cant'])} casos, desvío prom. {dg.iloc[0]['Prom']}, "
                f"máx. {dg.iloc[0]['Max']} unidades). "
                f"Caso extremo: {fm['Vehiculo']} — valor relevado "
                f"{round(float(fm['ValorRelevado']),1) if pd.notna(fm['ValorRelevado']) else '-'}, "
                f"{'por encima del máximo' if dval>0 else 'por debajo del mínimo'} "
                f"en {abs(dval)} unidades."
            )

        for idx, texto in enumerate(conclusiones, 1):
            p   = doc.add_paragraph(style="List Number")
            run = p.add_run(texto)
            run.font.size = Pt(10)
            run.font.name = "Arial"

    # ── GRÁFICOS ──
    def get_png_dimensions(png_bytes):
        """Obtiene ancho y alto de un PNG sin depender de PIL."""
        try:
            from PIL import Image
            img = Image.open(io.BytesIO(png_bytes))
            return img.size  # (width, height)
        except Exception:
            pass
        # Fallback: leer header PNG directamente (bytes 16-23 = width + height)
        try:
            import struct
            if png_bytes[:8] == b'\x89PNG\r\n\x1a\n':
                w = struct.unpack('>I', png_bytes[16:20])[0]
                h = struct.unpack('>I', png_bytes[20:24])[0]
                if w > 0 and h > 0:
                    return (w, h)
        except Exception:
            pass
        return None

    def insert_graph(key, caption):
        png = graficos.get(key)
        if not png:
            return
        next_sec(f"Gráfico — {caption}")
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER

        max_width_mm = 165
        max_height_mm = 110

        dims = get_png_dimensions(png)
        if dims:
            img_w, img_h = dims
            aspect = img_h / img_w
            width_mm = max_width_mm
            height_mm = width_mm * aspect
            if height_mm > max_height_mm:
                height_mm = max_height_mm
                width_mm = height_mm / aspect
            run_img = p_img.add_run()
            run_img.add_picture(io.BytesIO(png), width=Mm(width_mm), height=Mm(height_mm))
        else:
            # Sin dimensiones: solo fijar ancho, Word calcula alto proporcional
            run_img = p_img.add_run()
            run_img.add_picture(io.BytesIO(png), width=Mm(max_width_mm))
        doc.add_paragraph()

    insert_graph('torta',      'Distribución por Criticidad')
    insert_graph('sistemas',   'Observaciones por Sistema')
    insert_graph('mensual',    'Evolución Mensual y Ratio Obs/Vehículo')
    insert_graph('top15',      'Top 15 Tipos de Falla')
    insert_graph('desvios',    'Desvíos: Distribución por Categoría')
    insert_graph('pareto_fr',  'Pareto — Fuera de Rango')
    insert_graph('pareto_aus', 'Pareto — Ausencia de Elementos')
    insert_graph('pareto_mal', 'Pareto — Mal Estado')
    insert_graph('explorador', 'Explorador Libre')

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
AXIS_STYLE = dict(gridcolor='#1e2a3a', linecolor='#2a3a50')

# Tema con fondo blanco para exportación Word
PLOTLY_THEME_EXPORT = dict(
    paper_bgcolor='#FFFFFF',
    plot_bgcolor='#F5F8FA',
    font=dict(family='Arial', color='#333333', size=12),
    colorway=['#2E75B6','#E74C3C','#F39C12','#27AE60','#8E44AD','#16A085'],
)
AXIS_STYLE_EXPORT = dict(gridcolor='#E0E0E0', linecolor='#CCCCCC')


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
        df_base = df_base.copy()
        df_base['_unidad'] = df_base['Modulo'].apply(
            lambda x: None if (x is None or str(x).strip() in ('', '0', 'nan')) else str(x).strip()
        ).fillna(df_base['Vehiculo'].astype(str).str.strip())
        result = df_base['_unidad'].value_counts().head(10).reset_index()
        result.columns = ['Unidad', 'Observaciones']
        return result

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

    row_t3, row_t4 = st.columns(2)

    with row_t3:
        cfg = MR_CONFIG['CCEE']
        st.markdown(f"##### {cfg['label']}")
        df_mr = df[df['MR'] == 'CCEE']
        if df_mr.empty:
            st.caption("NO SE INSPECCIONÓ ESTE TIPO DE MR")
        else:
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
        fig_line.add_trace(go.Bar(
            x=monthly_df['Mes'], y=monthly_df['Observaciones'],
            name='Observaciones',
            marker_color='rgba(79,195,247,0.3)',
            text=monthly_df['Observaciones'],
            textposition='outside',
            textfont=dict(color='#4fc3f7', size=11),
            yaxis='y'
        ))
        fig_line.add_trace(go.Scatter(
            x=monthly_df['Mes'], y=monthly_df['Vehículos'],
            name='Vehículos inspeccionados',
            mode='lines+markers',
            line=dict(color='#66bb6a', width=2, dash='dot'),
            marker=dict(size=7, color='#66bb6a'),
            yaxis='y'
        ))
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

    if not df_out.empty and 'Desviacion' in df_out.columns:
        st.markdown("#### Desvíos: Distribución por Categoría")
        st.caption("Barras = cantidad de casos · Líneas = desvío máx/prom/mín sobre eje derecho (mismas unidades que el parámetro)")
        desv_plot = df_out.groupby('DescAgrupada').agg(
            Cantidad  =('Desviacion', 'count'),
            Desv_Max  =('Desviacion', lambda x: round(x.abs().max(),  2)),
            Desv_Prom =('Desviacion', lambda x: round(x.abs().mean(), 2)),
            Desv_Min  =('Desviacion', lambda x: round(x.abs().min(),  2)),
        ).reset_index().sort_values('Cantidad', ascending=False)

        desv_plot['Etiqueta'] = desv_plot.apply(
            lambda r: f"Máx: {r['Desv_Max']}  |  Prom: {r['Desv_Prom']}  |  Mín: {r['Desv_Min']}", axis=1
        )

        fig_desv = go.Figure()
        fig_desv.add_trace(go.Bar(
            x=desv_plot['DescAgrupada'], y=desv_plot['Cantidad'],
            name='Cantidad de casos',
            marker_color='#4fc3f7',
            text=desv_plot['Cantidad'],
            textposition='outside',
            textfont=dict(color='#4fc3f7', size=11),
            yaxis='y'
        ))
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
        fig_desv.add_trace(go.Scatter(
            x=desv_plot['DescAgrupada'], y=desv_plot['Desv_Min'],
            name='Desvío Mín.',
            mode='lines+markers',
            line=dict(color='#66bb6a', width=1, dash='dash'),
            marker=dict(size=6, color='#66bb6a'),
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

        df_out['_unidad'] = df_out['Modulo'].apply(
            lambda x: None if (x is None or str(x).strip() in ('', '0', 'nan')) else str(x).strip()
        ).fillna(df_out['Vehiculo'].astype(str).str.strip())
        unidades_disponibles = sorted(df_out['_unidad'].dropna().unique().tolist())

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
        st.markdown("#### Tabla Cruzada: Sistema × Criticidad")
        st.caption("Cantidad de observaciones por sistema y nivel de criticidad")

        pivot = df.pivot_table(
            index='CritAmpliado',
            columns='SistemaUnidad',
            aggfunc='size',
            fill_value=0
        )
        orden_crit = ['Rechazado','Critico','Normal','Corregida']
        pivot = pivot.reindex([r for r in orden_crit if r in pivot.index])
        st.dataframe(pivot, use_container_width=True)

        st.markdown("---")

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

                conteo = df_cat['SistemaUnidad'].value_counts().reset_index()
                conteo.columns = ['Sistema', 'Cantidad']
                conteo['Label'] = conteo['Sistema'].map(SISTEMA_LABELS).fillna(conteo['Sistema'])
                conteo['%_acum'] = (conteo['Cantidad'].cumsum() / conteo['Cantidad'].sum() * 100).round(1)

                color = CLASIF_COLORS.get(cat, '#4fc3f7')

                fig_p = go.Figure()
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

    cols_explorar = {
        'Sistema':      'SistemaUnidad',
        'Criticidad':   'CritAmpliado',
        'Tipo MR':      'MR',
        'Modelo':       'Modelo',
        'Servicio':     'Servicio',
        'Mes':          'Mes',
        'Clasificación':'Clasificacion',
    }
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

    df_exp = df.groupby([col_x, col_color]).size().reset_index(name='Cantidad')

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

    with st.expander("Ver datos del gráfico"):
        pivot_exp = df_exp.pivot_table(
            index=col_x, columns=col_color, values='Cantidad', fill_value=0
        )
        st.dataframe(pivot_exp, use_container_width=True)


# ── TAB 6: EXPORTAR ──
with tab6:
    st.markdown("#### Configuración del Informe Word")

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

    st.markdown("##### Tablas a incluir")
    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        inc_crit    = st.checkbox("Distribución por Criticidad",  value=True)
        inc_top15   = st.checkbox("Top 15 Tipos de Falla",        value=True)
    with sc2:
        inc_desvios = st.checkbox("Tabla de Desvíos",             value=True)
        inc_detalle = st.checkbox("Detalle fuera de parámetro",   value=True)
    with sc3:
        inc_concl   = st.checkbox("Conclusiones automáticas",     value=True)

    st.markdown("---")

    st.markdown("##### Gráficos a incluir")
    st.caption("Se genera usando kaleido (preferido) o matplotlib como respaldo. "
               "Instalá `pip install kaleido` para mejor calidad.")

    gc1, gc2, gc3 = st.columns(3)
    with gc1:
        g_torta    = st.checkbox("Torta — Criticidad",            value=False, key="g_torta")
        g_sistemas = st.checkbox("Barras — Sistemas",             value=False, key="g_sistemas")
        g_mensual  = st.checkbox("Línea mensual + ratio",         value=False, key="g_mensual")
    with gc2:
        g_top15    = st.checkbox("Barras — Top 15 Fallas",        value=False, key="g_top15")
        g_desvios  = st.checkbox("Desvíos — distribución",        value=False, key="g_desvios")
    with gc3:
        g_pareto_fr  = st.checkbox("Pareto — Fuera de rango",     value=False, key="g_pfr")
        g_pareto_aus = st.checkbox("Pareto — Ausencia elementos", value=False, key="g_paus")
        g_pareto_mal = st.checkbox("Pareto — Mal estado",         value=False, key="g_pmal")
        g_explorador = st.checkbox("Explorador libre (actual)",   value=False, key="g_exp")

    st.markdown("---")

    if st.button("📄  Generar y descargar informe Word", type="primary"):
        logo_bytes = logo_file.read() if logo_file else None

        graficos = {}

        any_graf = any([g_torta, g_sistemas, g_mensual, g_top15, g_desvios,
                        g_pareto_fr, g_pareto_aus, g_pareto_mal, g_explorador])

        if any_graf:
            with st.spinner("Generando gráficos para el documento..."):
                # Torta criticidad
                if g_torta:
                    crit_counts = df['CritAmpliado'].value_counts()
                    fig = go.Figure(go.Pie(
                        labels=crit_counts.index, values=crit_counts.values, hole=0.55,
                        marker=dict(colors=['#2E75B6','#E74C3C','#F39C12','#27AE60','#8E44AD']),
                        textinfo='label+percent', textfont=dict(size=11, color='#333333'),
                    ))
                    fig.update_layout(**PLOTLY_THEME_EXPORT, height=350,
                                      margin=dict(t=30,b=30,l=30,r=30))
                    graficos['torta'] = fig_to_png_robust(fig, 600, 350)

                # Barras sistemas
                if g_sistemas:
                    sc = df['SistemaUnidad'].value_counts().reset_index()
                    sc.columns = ['Sistema','Cantidad']
                    sc['Label'] = sc['Sistema'].map(SISTEMA_LABELS).fillna(sc['Sistema'])
                    fig = px.bar(sc, x='Cantidad', y='Label', orientation='h',
                                 color='Cantidad', color_continuous_scale='Blues', text='Cantidad')
                    fig.update_traces(textposition='outside', textfont=dict(color='#333333', size=11))
                    fig.update_layout(**PLOTLY_THEME_EXPORT, height=350,
                                      margin=dict(t=10,b=10,l=10,r=80),
                                      xaxis=dict(range=[0, sc['Cantidad'].max()*1.15], **AXIS_STYLE_EXPORT),
                                      yaxis=dict(autorange='reversed', **AXIS_STYLE_EXPORT),
                                      coloraxis_showscale=False)
                    graficos['sistemas'] = fig_to_png_robust(fig, 700, 350)

                # Línea mensual
                if g_mensual:
                    monthly_exp = []
                    for m in MONTH_ORDER:
                        df_m = df[df['Mes'] == m]
                        cnt = len(df_m); vehs = df_m['Vehiculo'].dropna().nunique()
                        if cnt > 0:
                            monthly_exp.append({'Mes':m,'Observaciones':cnt,'Vehículos':vehs,
                                            'Ratio':round(cnt/vehs,2) if vehs>0 else 0})
                    if monthly_exp:
                        mdf = pd.DataFrame(monthly_exp)
                        fig = go.Figure()
                        fig.add_trace(go.Bar(x=mdf['Mes'], y=mdf['Observaciones'],
                                             name='Obs.', marker_color='rgba(46,117,182,0.5)',
                                             text=mdf['Observaciones'], textposition='outside',
                                             textfont=dict(color='#333333', size=10)))
                        fig.add_trace(go.Scatter(x=mdf['Mes'], y=mdf['Ratio'],
                                                 name='Ratio Obs/Veh', mode='lines+markers',
                                                 line=dict(color='#F39C12',width=2), yaxis='y2'))
                        fig.update_layout(**PLOTLY_THEME_EXPORT, height=320, barmode='overlay',
                                          margin=dict(t=20,b=20,l=10,r=60),
                                          xaxis=AXIS_STYLE_EXPORT,
                                          yaxis=dict(title='Obs.', **AXIS_STYLE_EXPORT),
                                          yaxis2=dict(title='Ratio', overlaying='y', side='right',
                                                      gridcolor='#E0E0E0', linecolor='#CCCCCC'))
                        graficos['mensual'] = fig_to_png_robust(fig, 800, 320)

                # Top 15 fallas
                if g_top15:
                    tp = df['DescAgrupada'].value_counts().head(15).reset_index()
                    tp.columns = ['Falla','Cantidad']
                    fig = px.bar(tp, x='Cantidad', y='Falla', orientation='h',
                                 color='Cantidad', color_continuous_scale='Blues_r', text='Cantidad')
                    fig.update_traces(textposition='outside', textfont=dict(color='#333333', size=11))
                    fig.update_layout(**PLOTLY_THEME_EXPORT, height=420,
                                      margin=dict(t=10,b=10,r=80,l=10),
                                      xaxis=dict(range=[0,tp['Cantidad'].max()*1.15], **AXIS_STYLE_EXPORT),
                                      yaxis=dict(autorange='reversed', **AXIS_STYLE_EXPORT),
                                      coloraxis_showscale=False)
                    graficos['top15'] = fig_to_png_robust(fig, 800, 420)

                # Desvíos distribución
                if g_desvios and not df_out.empty and 'Desviacion' in df_out.columns:
                    dp = df_out.groupby('DescAgrupada').agg(
                        Cantidad=('Desviacion','count'),
                        Desv_Max=('Desviacion', lambda x: round(x.abs().max(),2)),
                        Desv_Prom=('Desviacion', lambda x: round(x.abs().mean(),2)),
                    ).reset_index().sort_values('Cantidad', ascending=False)
                    fig = go.Figure()
                    fig.add_trace(go.Bar(x=dp['DescAgrupada'], y=dp['Cantidad'],
                                         name='Cantidad', marker_color='#2E75B6',
                                         text=dp['Cantidad'], textposition='outside',
                                         textfont=dict(color='#333333', size=10), yaxis='y'))
                    fig.add_trace(go.Scatter(x=dp['DescAgrupada'], y=dp['Desv_Max'],
                                             name='Desvío Máx.', mode='lines+markers',
                                             line=dict(color='#E74C3C',width=2), yaxis='y2'))
                    fig.add_trace(go.Scatter(x=dp['DescAgrupada'], y=dp['Desv_Prom'],
                                             name='Desvío Prom.', mode='lines+markers',
                                             line=dict(color='#F39C12',width=2,dash='dot'), yaxis='y2'))
                    fig.update_layout(**PLOTLY_THEME_EXPORT, height=400, barmode='group',
                                      margin=dict(t=20,b=100,l=10,r=60),
                                      xaxis=dict(tickangle=-35, **AXIS_STYLE_EXPORT),
                                      yaxis=dict(title='Cantidad', **AXIS_STYLE_EXPORT),
                                      yaxis2=dict(title='Desvío', overlaying='y', side='right',
                                                  gridcolor='#E0E0E0', linecolor='#CCCCCC'))
                    graficos['desvios'] = fig_to_png_robust(fig, 900, 400)

                # Paretos
                CLASIF_COLORS_EXP = {'fuera de rango':'#E74C3C','ausencia de elementos':'#F39C12','mal estado':'#2E75B6'}
                pareto_cfg = [
                    ('pareto_fr',  'fuera de rango',        g_pareto_fr),
                    ('pareto_aus', 'ausencia de elementos', g_pareto_aus),
                    ('pareto_mal', 'mal estado',            g_pareto_mal),
                ]
                if 'Clasificacion' in df.columns:
                    for key, cat, sel in pareto_cfg:
                        if not sel: continue
                        df_cat = df[df['Clasificacion'].str.strip().str.lower() == cat]
                        if df_cat.empty: continue
                        conteo = df_cat['SistemaUnidad'].value_counts().reset_index()
                        conteo.columns = ['Sistema','Cantidad']
                        conteo['Label'] = conteo['Sistema'].map(SISTEMA_LABELS).fillna(conteo['Sistema'])
                        conteo['%_acum'] = (conteo['Cantidad'].cumsum()/conteo['Cantidad'].sum()*100).round(1)
                        color = CLASIF_COLORS_EXP.get(cat, '#2E75B6')
                        fig = go.Figure()
                        fig.add_trace(go.Bar(x=conteo['Label'], y=conteo['Cantidad'],
                                             marker_color=color, text=conteo['Cantidad'],
                                             textposition='outside',
                                             textfont=dict(color='#333333', size=10), yaxis='y'))
                        fig.add_trace(go.Scatter(x=conteo['Label'], y=conteo['%_acum'],
                                                 mode='lines+markers',
                                                 line=dict(color='#333333',width=2), yaxis='y2'))
                        fig.add_hline(y=80, line_dash='dash', line_color='#F39C12',
                                      annotation_text='80%', yref='y2')
                        fig.update_layout(**PLOTLY_THEME_EXPORT, height=350,
                                          margin=dict(t=30,b=60,l=10,r=50),
                                          xaxis=dict(tickangle=-35, **AXIS_STYLE_EXPORT),
                                          yaxis=dict(title='Cantidad', **AXIS_STYLE_EXPORT),
                                          yaxis2=dict(title='% Acum.', overlaying='y', side='right',
                                                      range=[0,110], gridcolor='#E0E0E0', linecolor='#CCCCCC'),
                                          title=dict(text=cat.title(), font=dict(color='#333333')))
                        graficos[key] = fig_to_png_robust(fig, 700, 350)

                # Explorador libre
                if g_explorador:
                    try:
                        col_x_e     = st.session_state.get('ex_x',     'Sistema')
                        col_color_e = st.session_state.get('ex_col',    'Criticidad')
                        tipo_e      = st.session_state.get('ex_tipo',   'Barras agrupadas')
                        cx  = {'Sistema':'SistemaUnidad','Criticidad':'CritAmpliado','Tipo MR':'MR',
                               'Modelo':'Modelo','Servicio':'Servicio','Mes':'Mes',
                               'Clasificación':'Clasificacion'}.get(col_x_e, 'SistemaUnidad')
                        cc  = {'Sistema':'SistemaUnidad','Criticidad':'CritAmpliado','Tipo MR':'MR',
                               'Modelo':'Modelo','Servicio':'Servicio','Mes':'Mes',
                               'Clasificación':'Clasificacion'}.get(col_color_e, 'CritAmpliado')
                        df_e = df.groupby([cx, cc]).size().reset_index(name='Cantidad')
                        bk = dict(x=cx, y='Cantidad', color=cc, text_auto=True,
                                  color_discrete_sequence=['#2E75B6','#E74C3C','#F39C12',
                                                           '#27AE60','#8E44AD','#16A085'])
                        if tipo_e == "Barras apiladas %":
                            bk['barmode'] = 'stack'; bk['barnorm'] = 'percent'
                        elif tipo_e == "Barras apiladas":
                            bk['barmode'] = 'stack'
                        else:
                            bk['barmode'] = 'group'
                        fig = px.bar(df_e, **bk)
                        fig.update_layout(**PLOTLY_THEME_EXPORT, height=400,
                                          margin=dict(t=20,b=80,l=10,r=10),
                                          xaxis=dict(tickangle=-30, **AXIS_STYLE_EXPORT),
                                          yaxis=AXIS_STYLE_EXPORT)
                        graficos['explorador'] = fig_to_png_robust(fig, 800, 400)
                    except Exception:
                        pass

            # Filtrar gráficos que fallaron (None)
            graficos = {k: v for k, v in graficos.items() if v is not None}

            if any_graf and not graficos:
                st.warning(
                    "⚠️ No se pudo generar ningún gráfico. "
                    "Instalá `kaleido` (`pip install kaleido`) o `matplotlib` para habilitar la exportación de gráficos."
                )
            elif any_graf and len(graficos) < sum([g_torta, g_sistemas, g_mensual, g_top15,
                                                     g_desvios, g_pareto_fr, g_pareto_aus,
                                                     g_pareto_mal, g_explorador]):
                st.info(f"ℹ️ Se generaron {len(graficos)} de los gráficos seleccionados. "
                        "Algunos no pudieron exportarse.")

        config = dict(
            codigo=hdr_codigo, version=hdr_version, linea=hdr_linea,
            subger=hdr_subger, logo=logo_bytes,
            inc_crit=inc_crit, inc_top15=inc_top15,
            inc_desvios=inc_desvios, inc_detalle=inc_detalle,
            inc_concl=inc_concl,
            graficos=graficos,
        )
        with st.spinner("Generando documento Word..."):
            word_bytes = generar_word(df, df_out, config)

        st.success("✅ Informe generado correctamente.")
        linea_safe = (hdr_linea or "Informe").replace(" ","_").replace("/","_")[:30]
        st.download_button(
            label="⬇️  Descargar .docx",
            data=word_bytes,
            file_name=f"Informe_{linea_safe}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
