"""
App Streamlit — Análisis de Informes de Mantenimiento v3
=========================================================
Sub-tabs por MR + filtro Modelo + filtro Meses en TODOS los tabs.
Word configurable por MR.
"""
import io, re, numpy as np, pandas as pd, streamlit as st
import plotly.express as px, plotly.graph_objects as go
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Análisis de Mantenimiento", page_icon="🔧", layout="wide", initial_sidebar_state="expanded")

st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');
html,body,[class*="css"]{font-family:'IBM Plex Sans',sans-serif}
.stApp{background-color:#0f1117;color:#e8eaf0}
section[data-testid="stSidebar"]{background:#161b27;border-right:1px solid #2a3040}
h1{color:#4fc3f7!important;font-weight:700!important}h2{color:#81d4fa!important;font-weight:600!important}h3{color:#b0bec5!important;font-weight:600!important}
.kpi-card{background:linear-gradient(135deg,#1a2235,#1e2a3a);border:1px solid #2a3a50;border-left:4px solid #4fc3f7;border-radius:8px;padding:20px 24px;text-align:center}
.kpi-value{font-family:'IBM Plex Mono',monospace;font-size:2.4rem;font-weight:700;color:#4fc3f7;line-height:1;margin-bottom:6px}
.kpi-label{font-size:.78rem;color:#78909c;text-transform:uppercase;letter-spacing:.08em}
.kpi-card.danger{border-left-color:#ef5350}.kpi-card.danger .kpi-value{color:#ef5350}
.kpi-card.warning{border-left-color:#ffa726}.kpi-card.warning .kpi-value{color:#ffa726}
.kpi-card.success{border-left-color:#66bb6a}.kpi-card.success .kpi-value{color:#66bb6a}
.section-header{border-bottom:2px solid #2a3a50;padding-bottom:8px;margin:32px 0 16px;color:#4fc3f7;font-size:1.1rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em}
.stTabs [data-baseweb="tab-list"]{gap:8px;background:transparent}
.stTabs [data-baseweb="tab"]{background:#1a2235;border-radius:6px 6px 0 0;color:#78909c;padding:8px 20px;border:1px solid #2a3a50;border-bottom:none}
.stTabs [aria-selected="true"]{background:#1e2a3a!important;color:#4fc3f7!important;border-color:#4fc3f7!important}
</style>""", unsafe_allow_html=True)

# ── Constantes ──
MONTH_ORDER = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']
SISTEMA_LABELS = {'BG':'Bogie','SLN':'Salón','EXT':'Exterior','TYC':'Tracción y Choque','PM':'Par Montado','EBC':'Elem. bajo coche','NSF':'Sist. freno Neumático','MSF':'Sist. freno Mecánico','SFM':'Sist. freno Mecánico','DSM':'Sala motor Diesel','CAB':'Cabina','ATS':'ATS','DOC':'Documentación','NGN':'Sin observación'}
MR_LABELS = {'LOC':'Locomotoras','CCRR':'Coches Remolcados','CCEE':'Coches Eléctricos','CCMM':'Coche Motor'}
PT = dict(paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(26,34,53,0.6)',font=dict(family='IBM Plex Sans',color='#b0bec5'),colorway=['#4fc3f7','#81d4fa','#ef5350','#ffa726','#66bb6a','#ab47bc','#26c6da'])
AX = dict(gridcolor='#1e2a3a',linecolor='#2a3a50')
CCOL = {'Fuera de rango':'#ef5350','Ausencia de elementos':'#ffa726','Mal estado':'#4fc3f7'}

# ── Parser ──
def parse_valor(v):
    if v is None: return np.nan
    try: return float(v)
    except:
        nums = re.findall(r'[\d.]+', str(v))
        return float(nums[0]) if nums else np.nan

def calcular_desvio(row):
    vals = [v for v in [row['R1_num'], row['R2_num']] if not np.isnan(v)]
    if not vals: return np.nan, np.nan
    mejor_dev, mejor_val = 0, vals[0]
    for v in vals:
        if not np.isnan(row['RefMin_num']) and v < row['RefMin_num']:
            dev = v - row['RefMin_num']
            if abs(dev) > abs(mejor_dev): mejor_dev, mejor_val = dev, v
        if not np.isnan(row['RefMax_num']) and v > row['RefMax_num']:
            dev = v - row['RefMax_num']
            if abs(dev) > abs(mejor_dev): mejor_dev, mejor_val = dev, v
    return mejor_val, mejor_dev

@st.cache_data(show_spinner=False)
def cargar_y_analizar(file_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True); ws = wb.active
    header_row = 2
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
        if row[1] is not None: header_row = i; break
    data_start = header_row + 1
    headers_raw = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column+1)]
    first_header = next((h for h in headers_raw if h is not None), None)
    b1_start = next((i for i, h in enumerate(headers_raw) if h is not None), 1)
    bloque2_col = None
    if first_header:
        found = False
        for i, h in enumerate(headers_raw):
            if h == first_header:
                if not found: found = True
                else: bloque2_col = i; break
    b1_end = bloque2_col if bloque2_col else len(headers_raw)
    hdrs_b1 = [h for h in headers_raw[b1_start:b1_end] if h is not None]
    tiene_refs = any('Referencia' in str(h) or 'Relevado' in str(h) for h in hdrs_b1)
    rows1, rows2 = [], []
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        r1 = row[b1_start:b1_end]
        if any(v is not None for v in r1): rows1.append(r1)
        if bloque2_col:
            b2s, b2e = bloque2_col, bloque2_col+(b1_end-b1_start)
            if b2e <= len(row):
                r2 = row[b2s:b2e]
                if any(v is not None for v in r2): rows2.append(r2)
    if tiene_refs:
        cols = ['Mes','Responsable','Contrato','Linea','Vehiculo','Modulo','MR','Modelo','Servicio','Fecha','NroInforme','SistemaUnidad','SistemaAmpliado','Item1','Item2','Descripcion','RefMin','RefMax','Relevado1','Relevado2','Criticidad','DescAgrupada','CritAmpliado','CodItem','FechaReInsp','NroReInsp','SistUnitReInsp','SistAmpReInsp','ItemsReInsp','DescReInsp','CritReInsp','DescAgrupReInsp','CodReInsp','Clasificacion']
    else:
        cols = ['Mes','Responsable','Contrato','Linea','Vehiculo','Modulo','MR','Modelo','Servicio','Fecha','NroInforme','SistemaUnidad','SistemaAmpliado','Item1','Item2','Descripcion','Criticidad','DescAgrupada','CritAmpliado','CodItem','FechaReInsp','NroReInsp','SistUnitReInsp','SistAmpReInsp','ItemsReInsp','DescReInsp','CritReInsp','DescAgrupReInsp','CodReInsp','Clasificacion']
    def filas_a_df(rows):
        if not rows: return pd.DataFrame(columns=cols)
        n, nc = len(rows[0]), len(cols)
        if n == nc: return pd.DataFrame(rows, columns=cols)
        elif n < nc: return pd.DataFrame([list(r)+[None]*(nc-n) for r in rows], columns=cols)
        else: return pd.DataFrame([r[:nc] for r in rows], columns=cols)
    dfs = [filas_a_df(rows1)]
    if rows2: dfs.append(filas_a_df(rows2))
    df = pd.concat(dfs, ignore_index=True)
    if 'Mes' in df.columns: df['Mes'] = df['Mes'].astype(str).str.strip().str.upper()
    if 'Clasificacion' in df.columns:
        df['Clasificacion'] = df['Clasificacion'].apply(lambda x: None if (pd.isna(x) or str(x).strip().startswith('=')) else str(x).strip())
    df['CodItem_num'] = pd.to_numeric(df.get('CodItem'), errors='coerce').fillna(1).astype(int) if 'CodItem' in df.columns else 1
    if tiene_refs and all(c in df.columns for c in ['RefMin','RefMax','Relevado1','Relevado2']):
        df_ref = df[df['RefMin'].notna() | df['RefMax'].notna()].copy()
        df_ref['R1_num'] = df_ref['Relevado1'].apply(parse_valor)
        df_ref['R2_num'] = df_ref['Relevado2'].apply(parse_valor)
        df_ref['RefMin_num'] = pd.to_numeric(df_ref['RefMin'], errors='coerce')
        df_ref['RefMax_num'] = pd.to_numeric(df_ref['RefMax'], errors='coerce')
        df_ref[['ValorRelevado','Desviacion']] = df_ref.apply(lambda r: pd.Series(calcular_desvio(r)), axis=1)
        df_out = df_ref[df_ref['Desviacion'].notna() & (df_ref['Desviacion'] != 0)].copy()
    else:
        df_out = pd.DataFrame(columns=list(df.columns)+['ValorRelevado','Desviacion'])
    return df, df_out

def kpi(label, value, variant="default"):
    return f'<div class="kpi-card {variant}"><div class="kpi-value">{value}</div><div class="kpi-label">{label}</div></div>'

# ─────────────────────────────────────────────
# HELPER: Sub-tabs MR + Filtro Modelo + Meses
# ─────────────────────────────────────────────
def get_mr_list(df):
    mrs = df['MR'].dropna().unique().tolist()
    return [mr for mr in ['LOC','CCRR','CCEE','CCMM'] if mr in mrs]

def render_filters(df_subset, df_out_subset, key):
    """Filtros de Modelo y Meses. Devuelve (df_filtrado, df_out_filtrado)."""
    c1, c2 = st.columns(2)
    with c1:
        modelos = sorted(df_subset['Modelo'].dropna().unique().tolist())
        if len(modelos) > 1:
            sel_mod = st.multiselect("🔧 Modelo", modelos, default=modelos, key=f"{key}_mod")
            df_subset = df_subset[df_subset['Modelo'].isin(sel_mod)] if sel_mod else df_subset
            if not df_out_subset.empty:
                df_out_subset = df_out_subset[df_out_subset['Modelo'].isin(sel_mod)] if sel_mod else df_out_subset
    with c2:
        meses_pres = [m for m in MONTH_ORDER if m in df_subset['Mes'].dropna().str.upper().unique()]
        if len(meses_pres) > 1:
            sel_mes = st.multiselect("📅 Meses", meses_pres, default=meses_pres, key=f"{key}_mes")
            df_subset = df_subset[df_subset['Mes'].str.upper().isin(sel_mes)] if sel_mes else df_subset
            if not df_out_subset.empty:
                df_out_subset = df_out_subset[df_out_subset['Mes'].str.upper().isin(sel_mes)] if sel_mes else df_out_subset
    return df_subset, df_out_subset

def render_with_mr_subtabs(df, df_out, tab_id, render_fn):
    """Crea sub-tabs [General + cada MR] con filtros y llama render_fn."""
    mr_list = get_mr_list(df)
    labels = ["📋 General"] + [f"🚂 {MR_LABELS.get(m,m)}" for m in mr_list]
    subs = st.tabs(labels)
    with subs[0]:
        d, do = render_filters(df, df_out, f"{tab_id}_GEN")
        if d.empty: st.info("Sin datos para esta selección.")
        else: render_fn(d, do, f"{tab_id}_GEN")
    for i, mr in enumerate(mr_list):
        with subs[i+1]:
            d = df[df['MR'] == mr]
            do = df_out[df_out['MR'] == mr] if 'MR' in df_out.columns and not df_out.empty else pd.DataFrame()
            d, do = render_filters(d, do, f"{tab_id}_{mr}")
            if d.empty: st.info("Sin datos para esta selección.")
            else: render_fn(d, do, f"{tab_id}_{mr}")

# ═══════════════════════════════════════════════
# RENDER FUNCTIONS — contenido completo en cada sub-tab
# ═══════════════════════════════════════════════

def render_resumenes(df, df_out, k):
    n = int(df['CodItem_num'].sum())
    if n == 0: return
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### Criticidad")
        d = df.groupby('CritAmpliado')['CodItem_num'].sum().sort_values(ascending=False).reset_index()
        d.columns = ['Criticidad','Cantidad']; d['%'] = (d['Cantidad']/n*100).round(1).astype(str)+'%'
        st.dataframe(d, use_container_width=True, hide_index=True)
        st.markdown("#### Por Modelo")
        d = df.groupby('Modelo')['CodItem_num'].sum().sort_values(ascending=False).reset_index()
        d.columns = ['Modelo','Obs.']; d['%'] = (d['Obs.']/n*100).round(1).astype(str)+'%'
        st.dataframe(d, use_container_width=True, hide_index=True)
    with col_b:
        st.markdown("#### Por Sistema")
        d = df.groupby('SistemaUnidad')['CodItem_num'].sum().sort_values(ascending=False).reset_index()
        d.columns = ['Código','Cantidad']; d['Sistema'] = d['Código'].map(SISTEMA_LABELS).fillna(d['Código'])
        d['%'] = (d['Cantidad']/n*100).round(1).astype(str)+'%'
        st.dataframe(d[['Código','Sistema','Cantidad','%']], use_container_width=True, hide_index=True)
    st.markdown("#### Top 15 Tipos de Falla")
    d = df.groupby('DescAgrupada')['CodItem_num'].sum().sort_values(ascending=False).head(15).reset_index()
    d.columns = ['Falla','Cantidad']; d['%'] = (d['Cantidad']/n*100).round(1).astype(str)+'%'
    st.dataframe(d, use_container_width=True, hide_index=True)
    st.markdown("#### Top 10 Unidades")
    df2 = df.copy()
    df2['_u'] = df2['Modulo'].apply(lambda x: None if (x is None or str(x).strip() in ('','0','nan')) else str(x).strip()).fillna(df2['Vehiculo'].astype(str).str.strip())
    d = df2.groupby('_u')['CodItem_num'].sum().sort_values(ascending=False).head(10).reset_index()
    d.columns = ['Unidad','Obs.']
    st.dataframe(d, use_container_width=True, hide_index=True)


def render_graficos(df, df_out, k):
    n = int(df['CodItem_num'].sum())
    if n == 0: return
    r1l, r1r = st.columns(2)
    with r1l:
        st.markdown("#### Criticidad")
        cc = df.groupby('CritAmpliado')['CodItem_num'].sum().sort_values(ascending=False)
        fig = go.Figure(go.Pie(labels=cc.index, values=cc.values, hole=0.55,
            marker=dict(colors=['#4fc3f7','#ef5350','#ffa726','#66bb6a','#ab47bc']), textfont=dict(size=13)))
        fig.update_layout(**PT, margin=dict(t=10,b=10,l=10,r=10), height=300, legend=dict(orientation='h',y=-0.1))
        st.plotly_chart(fig, use_container_width=True, key=f"pie_{k}")
    with r1r:
        st.markdown("#### Por Sistema")
        sc = df.groupby('SistemaUnidad')['CodItem_num'].sum().sort_values(ascending=False).reset_index()
        sc.columns = ['Sistema','Cantidad']; sc['Label'] = sc['Sistema'].map(SISTEMA_LABELS).fillna(sc['Sistema'])
        fig = px.bar(sc, x='Cantidad', y='Label', orientation='h', color='Cantidad', color_continuous_scale='Blues', text='Cantidad')
        fig.update_traces(textposition='outside',textfont=dict(color='#b0bec5',size=12))
        fig.update_layout(**PT,height=300,margin=dict(t=10,b=10,r=80,l=10),
            xaxis=dict(range=[0,sc['Cantidad'].max()*1.15],**AX),yaxis=dict(autorange='reversed',**AX),coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True, key=f"sist_{k}")

    st.markdown("#### Evolución Mensual")
    monthly = []
    for m in MONTH_ORDER:
        dm = df[df['Mes'].str.upper()==m]; cnt = int(dm['CodItem_num'].sum()); vehs = dm['Vehiculo'].dropna().nunique()
        if cnt > 0: monthly.append({'Mes':m,'Obs':cnt,'Veh':vehs,'Ratio':round(cnt/vehs,2) if vehs>0 else 0})
    mdf = pd.DataFrame(monthly)
    if not mdf.empty:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=mdf['Mes'],y=mdf['Obs'],name='Observaciones',marker_color='rgba(79,195,247,0.3)',text=mdf['Obs'],textposition='outside',textfont=dict(color='#4fc3f7',size=11)))
        fig.add_trace(go.Scatter(x=mdf['Mes'],y=mdf['Veh'],name='Vehículos',mode='lines+markers',line=dict(color='#66bb6a',width=2,dash='dot'),marker=dict(size=7,color='#66bb6a')))
        fig.add_trace(go.Scatter(x=mdf['Mes'],y=mdf['Ratio'],name='Ratio',mode='lines+markers+text',line=dict(color='#ffa726',width=2),marker=dict(size=8,color='#ffa726'),text=mdf['Ratio'].astype(str),textposition='top center',textfont=dict(color='#ffa726',size=10),yaxis='y2'))
        fig.update_layout(**PT,height=360,margin=dict(t=20,b=20,l=10,r=60),xaxis=AX,yaxis=dict(title='Cantidad',**AX),yaxis2=dict(title='Ratio',overlaying='y',side='right',gridcolor='#1e2a3a',linecolor='#2a3a50'),legend=dict(orientation='h',y=1.08),barmode='overlay')
        st.plotly_chart(fig, use_container_width=True, key=f"line_{k}")

    st.markdown("#### Top 15 Fallas")
    tp = df.groupby('DescAgrupada')['CodItem_num'].sum().sort_values(ascending=False).head(15).reset_index()
    tp.columns = ['Falla','Cantidad']
    fig = px.bar(tp, x='Cantidad', y='Falla', orientation='h', color='Cantidad', color_continuous_scale='Blues_r', text='Cantidad')
    fig.update_traces(textposition='outside',textfont=dict(color='#b0bec5',size=12))
    fig.update_layout(**PT,height=420,margin=dict(t=10,b=10,r=80,l=10),xaxis=dict(range=[0,tp['Cantidad'].max()*1.15],**AX),yaxis=dict(autorange='reversed',**AX),coloraxis_showscale=False)
    st.plotly_chart(fig, use_container_width=True, key=f"top15_{k}")

    if not df_out.empty and 'Desviacion' in df_out.columns:
        st.markdown("#### Desvíos por Categoría")
        dp = df_out.groupby('DescAgrupada').agg(Cantidad=('CodItem_num','sum'),Desv_Max=('Desviacion',lambda x:round(x.abs().max(),2)),Desv_Prom=('Desviacion',lambda x:round(x.abs().mean(),2)),Desv_Min=('Desviacion',lambda x:round(x.abs().min(),2))).reset_index().sort_values('Cantidad',ascending=False)
        fig = go.Figure()
        fig.add_trace(go.Bar(x=dp['DescAgrupada'],y=dp['Cantidad'],name='Cantidad',marker_color='#4fc3f7',text=dp['Cantidad'],textposition='outside',textfont=dict(color='#4fc3f7',size=11),yaxis='y'))
        fig.add_trace(go.Scatter(x=dp['DescAgrupada'],y=dp['Desv_Max'],name='Máx.',mode='lines+markers',line=dict(color='#ef5350',width=2),marker=dict(size=8,color='#ef5350'),yaxis='y2'))
        fig.add_trace(go.Scatter(x=dp['DescAgrupada'],y=dp['Desv_Prom'],name='Prom.',mode='lines+markers',line=dict(color='#ffa726',width=2,dash='dot'),marker=dict(size=7,color='#ffa726'),yaxis='y2'))
        fig.update_layout(**PT,height=440,barmode='group',margin=dict(t=20,b=100,l=10,r=70),xaxis=dict(tickangle=-35,**AX),yaxis=dict(title='Cantidad',**AX),yaxis2=dict(title='Desvío',overlaying='y',side='right',gridcolor='#1e2a3a',linecolor='#2a3a50'),legend=dict(orientation='h',y=1.06))
        st.plotly_chart(fig, use_container_width=True, key=f"desv_{k}")


def render_desvios(df, df_out, k):
    if df_out.empty or 'Desviacion' not in df_out.columns:
        st.info("Sin datos de desvíos paramétricos."); return
    fc1, fc2 = st.columns(2)
    sist_d = sorted(df_out['SistemaUnidad'].dropna().unique().tolist())
    crit_d = sorted(df_out['Criticidad'].dropna().unique().tolist())
    with fc1: sel_s = st.multiselect("Sistema",sist_d,default=sist_d,key=f"ds_{k}")
    with fc2: sel_c = st.multiselect("Criticidad",crit_d,default=crit_d,key=f"dc_{k}")
    dff = df_out[df_out['SistemaUnidad'].isin(sel_s) & df_out['Criticidad'].isin(sel_c)].copy()
    dff['RefMin_num'] = pd.to_numeric(dff.get('RefMin'),errors='coerce')
    dff['RefMax_num'] = pd.to_numeric(dff.get('RefMax'),errors='coerce')
    def ref_str(r):
        mn = r.get('RefMin_num'); mx = r.get('RefMax_num')
        mn = mn if pd.notna(mn) else None; mx = mx if pd.notna(mx) else None
        if mn and mx: return f"{mn}-{mx}"
        if mn: return f">={mn}"
        if mx: return f"<={mx}"
        return "-"
    dff['Rango'] = dff.apply(ref_str,axis=1); dff['Desvio'] = dff['Desviacion'].round(2); dff['Relev'] = dff['ValorRelevado'].round(1)
    st.markdown(f"**{len(dff)}** observaciones fuera de parámetro")
    st.markdown("##### Resumen por Categoría")
    res = dff.groupby('DescAgrupada').agg(Cant=('CodItem_num','sum'),Prom=('Desvio',lambda x:round(x.mean(),2)),Max=('Desvio',lambda x:round(x.abs().max(),2))).reset_index().sort_values('Cant',ascending=False)
    res.columns = ['Categoría','Cant.','Desv.Prom','Desv.Máx']
    st.dataframe(res, use_container_width=True, hide_index=True)
    st.markdown("##### Detalle")
    cols_s = ['Vehiculo','Mes','SistemaUnidad','Descripcion','Rango','Relev','Desvio','Criticidad']
    cols_s = [c for c in cols_s if c in dff.columns]
    st.dataframe(dff[cols_s].rename(columns={'SistemaUnidad':'Sistema'}), use_container_width=True, hide_index=True)


def render_clasificacion(df, df_out, k):
    tiene_clasif = 'Clasificacion' in df.columns and df['Clasificacion'].notna().any()
    if not tiene_clasif: st.info("Sin columna 'Clasificacion'."); return
    n = int(df['CodItem_num'].sum())
    if n == 0: return

    st.markdown("#### Sistema × Criticidad")
    pivot = df.pivot_table(index='CritAmpliado',columns='SistemaUnidad',values='CodItem_num',aggfunc='sum',fill_value=0)
    orden = ['Rechazado','Critico','Normal','Corregida']
    pivot = pivot.reindex([r for r in orden if r in pivot.index])
    st.dataframe(pivot, use_container_width=True)

    st.markdown("---")
    st.markdown("#### Pareto por Clasificación")
    for cat in ['Fuera de rango','Ausencia de elementos','Mal estado']:
        df_cat = df[df['Clasificacion'].str.strip().str.lower()==cat.lower()]
        st.markdown(f"##### {cat}")
        if df_cat.empty: st.caption("Sin datos"); continue
        conteo = df_cat.groupby('SistemaUnidad')['CodItem_num'].sum().sort_values(ascending=False).reset_index()
        conteo.columns = ['Sistema','Cantidad']; conteo['Label'] = conteo['Sistema'].map(SISTEMA_LABELS).fillna(conteo['Sistema'])
        conteo['%_acum'] = (conteo['Cantidad'].cumsum()/conteo['Cantidad'].sum()*100).round(1)
        color = CCOL.get(cat,'#4fc3f7')
        fig = go.Figure()
        fig.add_trace(go.Bar(x=conteo['Label'],y=conteo['Cantidad'],marker_color=color,text=conteo['Cantidad'],textposition='outside',textfont=dict(color='#e8eaf0',size=11),yaxis='y'))
        fig.add_trace(go.Scatter(x=conteo['Label'],y=conteo['%_acum'],mode='lines+markers+text',line=dict(color='#ffffff',width=2),marker=dict(size=6,color='#ffffff'),text=[f"{v}%" for v in conteo['%_acum']],textposition='top center',textfont=dict(color='#ffffff',size=10),yaxis='y2'))
        fig.add_hline(y=80,line_dash='dash',line_color='#ffa726',line_width=1,annotation_text='80%',annotation_font_color='#ffa726',yref='y2')
        fig.update_layout(**PT,height=350,margin=dict(t=30,b=60,l=10,r=50),showlegend=False,xaxis=dict(tickangle=-35,**AX),yaxis=dict(title='Cantidad',**AX),yaxis2=dict(title='% Acum.',overlaying='y',side='right',range=[0,110],gridcolor='#1e2a3a',linecolor='#2a3a50'))
        st.plotly_chart(fig, use_container_width=True, key=f"par_{cat[:4]}_{k}")

    st.markdown("---"); st.markdown("#### Últimas Observaciones de Rechazo")
    df_r = df[df['Criticidad']=='R'].copy()
    if df_r.empty: st.info("Sin rechazos."); return
    df_r['Fecha'] = pd.to_datetime(df_r['Fecha'],errors='coerce'); df_r = df_r.sort_values('Fecha',ascending=False)
    cols_r = [c for c in ['Fecha','Modulo','Vehiculo','SistemaUnidad','Descripcion','CritAmpliado'] if c in df_r.columns]
    t = df_r[cols_r].head(15).copy(); t['Fecha'] = t['Fecha'].dt.strftime('%d/%m/%Y')
    t.columns = [{'SistemaUnidad':'Sistema','CritAmpliado':'Criticidad','Descripcion':'Descripción'}.get(c,c) for c in t.columns]
    st.dataframe(t, use_container_width=True, hide_index=True)


def render_explorador(df, df_out, k):
    n = int(df['CodItem_num'].sum())
    if n == 0: return
    cols_exp = {'Sistema':'SistemaUnidad','Criticidad':'CritAmpliado','Tipo MR':'MR','Modelo':'Modelo','Servicio':'Servicio','Mes':'Mes','Clasificación':'Clasificacion'}
    cols_exp = {kk:v for kk,v in cols_exp.items() if v in df.columns}
    e1,e2,e3 = st.columns(3)
    with e1: eje_x = st.selectbox("Eje X",list(cols_exp.keys()),index=0,key=f"exx_{k}")
    with e2: eje_c = st.selectbox("Color",list(cols_exp.keys()),index=min(1,len(cols_exp)-1),key=f"exc_{k}")
    with e3: tipo = st.selectbox("Tipo",["Barras agrupadas","Barras apiladas","Barras apiladas %"],key=f"ext_{k}")
    cx = cols_exp[eje_x]; cc = cols_exp[eje_c]
    de = df.groupby([cx,cc])['CodItem_num'].sum().reset_index(name='Cantidad')
    bk = dict(x=cx,y='Cantidad',color=cc,text_auto=True,labels={cx:eje_x,cc:eje_c},color_discrete_sequence=['#4fc3f7','#ef5350','#ffa726','#66bb6a','#ab47bc','#26c6da'])
    if tipo=="Barras apiladas %": bk['barmode']='stack'; bk['barnorm']='percent'
    elif tipo=="Barras apiladas": bk['barmode']='stack'
    else: bk['barmode']='group'
    fig = px.bar(de,**bk)
    fig.update_traces(textfont=dict(size=11,color='white'),textposition='inside')
    fig.update_layout(**PT,height=450,margin=dict(t=20,b=80,l=10,r=10),xaxis=dict(tickangle=-30,**AX),yaxis=AX,legend=dict(orientation='h',y=-0.25))
    st.plotly_chart(fig, use_container_width=True, key=f"exp_{k}")
    with st.expander("Ver datos"):
        st.dataframe(de.pivot_table(index=cx,columns=cc,values='Cantidad',aggfunc='sum',fill_value=0), use_container_width=True)

# ═══════════════════════════════════════════════
# WORD Y EXCEL
# ═══════════════════════════════════════════════

def generar_word(df, df_out, config=None):
    from docx.shared import Cm, Mm
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from datetime import date
    if config is None: config = {}
    secciones = config.get("secciones", {})
    mr_sel = config.get("mr_sel", ["General"])
    hdr_codigo = config.get("codigo","") or ""; hdr_version = config.get("version","v1.0") or "v1.0"
    hdr_linea = config.get("linea","") or ""; logo_bytes = config.get("logo",None)
    hdr_subger = config.get("subger","Programación y Seguimiento de Mantenimiento") or ""
    doc = Document()
    section = doc.sections[0]; section.page_width=Mm(210); section.page_height=Mm(297)
    section.left_margin=Mm(15); section.right_margin=Mm(15); section.top_margin=Mm(35)
    section.bottom_margin=Mm(15); section.header_distance=Mm(5)
    PAGE_W = int(180*56.69)
    doc.styles["Normal"].font.name="Arial"; doc.styles["Normal"].font.size=Pt(10)
    doc.styles["Normal"].paragraph_format.space_after=Pt(0); doc.styles["Normal"].paragraph_format.space_before=Pt(0)

    # XML helpers (compactos)
    def _shd(cell,hx):
        tc=cell._tc;p=tc.get_or_add_tcPr()
        for o in p.findall(qn("w:shd")):p.remove(o)
        s=OxmlElement("w:shd");s.set(qn("w:val"),"clear");s.set(qn("w:color"),"auto");s.set(qn("w:fill"),hx);p.append(s)
    def _cw(cell,w):
        tc=cell._tc;p=tc.get_or_add_tcPr()
        for o in p.findall(qn("w:tcW")):p.remove(o)
        e=OxmlElement("w:tcW");e.set(qn("w:w"),str(w));e.set(qn("w:type"),"dxa");p.insert(0,e)
    def _tw(tbl,w):
        p=tbl._tbl.tblPr
        if p is None:p=OxmlElement("w:tblPr");tbl._tbl.insert(0,p)
        for o in p.findall(qn("w:tblW")):p.remove(o)
        e=OxmlElement("w:tblW");e.set(qn("w:w"),str(w));e.set(qn("w:type"),"dxa");p.append(e)
    def _fix(tbl):
        p=tbl._tbl.tblPr
        if p is None:p=OxmlElement("w:tblPr");tbl._tbl.insert(0,p)
        for o in p.findall(qn("w:tblLayout")):p.remove(o)
        e=OxmlElement("w:tblLayout");e.set(qn("w:type"),"fixed");p.append(e)
    def _grid(tbl,widths):
        te=tbl._tbl
        for o in te.findall(qn("w:tblGrid")):te.remove(o)
        g=OxmlElement("w:tblGrid")
        for w in widths:c=OxmlElement("w:gridCol");c.set(qn("w:w"),str(w));g.append(c)
        p=te.find(qn("w:tblPr"))
        if p is not None:p.addnext(g)
        else:te.insert(0,g)
    def _brd(tbl,color="A0A0A0",sz="4"):
        p=tbl._tbl.tblPr
        if p is None:p=OxmlElement("w:tblPr");tbl._tbl.insert(0,p)
        for o in p.findall(qn("w:tblBorders")):p.remove(o)
        b=OxmlElement("w:tblBorders")
        for s in["top","left","bottom","right","insideH","insideV"]:
            e=OxmlElement(f"w:{s}");e.set(qn("w:val"),"single");e.set(qn("w:sz"),sz);e.set(qn("w:space"),"0");e.set(qn("w:color"),color);b.append(e)
        p.append(b)
    def _nobrd(tbl):
        p=tbl._tbl.tblPr
        if p is None:return
        for o in p.findall(qn("w:tblBorders")):p.remove(o)
        b=OxmlElement("w:tblBorders")
        for s in["top","left","bottom","right","insideH","insideV"]:
            e=OxmlElement(f"w:{s}");e.set(qn("w:val"),"none");e.set(qn("w:sz"),"0");e.set(qn("w:space"),"0");e.set(qn("w:color"),"auto");b.append(e)
        p.append(b)
    def _rh(row,h,rule="atLeast"):
        tr=row._tr;p=tr.get_or_add_trPr()
        for o in p.findall(qn("w:trHeight")):p.remove(o)
        e=OxmlElement("w:trHeight");e.set(qn("w:val"),str(h));e.set(qn("w:hRule"),rule);p.append(e)
    def _cm(cell,t=40,b=40,l=60,r=60):
        tc=cell._tc;p=tc.get_or_add_tcPr()
        for o in p.findall(qn("w:tcMar")):p.remove(o)
        m=OxmlElement("w:tcMar")
        for s,v in[("top",t),("bottom",b),("left",l),("right",r)]:e=OxmlElement(f"w:{s}");e.set(qn("w:w"),str(v));e.set(qn("w:type"),"dxa");m.append(e)
        p.append(m)
    def _va(cell,a="center"):
        tc=cell._tc;p=tc.get_or_add_tcPr()
        for o in p.findall(qn("w:vAlign")):p.remove(o)
        e=OxmlElement("w:vAlign");e.set(qn("w:val"),a);p.append(e)
    def _ct(cell,text,bold=False,size=9,color="000000",align=WD_ALIGN_PARAGRAPH.CENTER,italic=False):
        p=cell.paragraphs[0];p.clear();p.alignment=align;p.paragraph_format.space_before=Pt(1);p.paragraph_format.space_after=Pt(1)
        r=p.add_run(str(text) if pd.notna(text) else "");r.bold=bold;r.italic=italic;r.font.size=Pt(size);r.font.name="Arial";r.font.color.rgb=RGBColor.from_string(color)

    # Encabezado
    header=section.header
    for p in list(header.paragraphs):p._element.getparent().remove(p._element)
    WL=int(PAGE_W*0.60);WR=PAGE_W-WL
    ht=header.add_table(rows=1,cols=2,width=Mm(180));ht.style="Table Grid"
    _tw(ht,PAGE_W);_fix(ht);_grid(ht,[WL,WR]);_nobrd(ht);_rh(ht.rows[0],int(2.5*567))
    cl=ht.cell(0,0);cr=ht.cell(0,1);_cw(cl,WL);_cw(cr,WR);_va(cl);_va(cr)
    if logo_bytes:
        _shd(cl,"FFFFFF");cl.paragraphs[0].clear();cl.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.LEFT
        cl.paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes),height=Cm(2.2))
    else:
        _shd(cl,"1F3864");_ct(cl,"TRENES ARGENTINOS — PISE",bold=True,size=13,color="FFFFFF")
    _shd(cr,"EBF3FB")
    for p in list(cr.paragraphs):p._element.getparent().remove(p._element)
    fields=[("Código:",hdr_codigo or"___"),("Versión:",hdr_version),("Fecha:",date.today().strftime("%d/%m/%Y")),("Línea:",hdr_linea or"___"),("Subgerencia:",hdr_subger)]
    swl=int(WR*0.35);swv=WR-swl
    sub=cr.add_table(rows=len(fields),cols=2);_nobrd(sub);_tw(sub,WR);_fix(sub);_grid(sub,[swl,swv])
    for i,(lbl,val) in enumerate(fields):
        if i==len(fields)-1:
            mg=sub.cell(i,0).merge(sub.cell(i,1));_cw(mg,WR);_shd(mg,"D5E8F0")
            _ct(mg,f"{lbl} {val}",size=7,color="1F3864",align=WD_ALIGN_PARAGRAPH.LEFT,italic=True);_rh(sub.rows[i],int(0.55*567),"exact")
        else:
            lc=sub.cell(i,0);vc=sub.cell(i,1);_cw(lc,swl);_cw(vc,swv);_shd(lc,"EBF3FB");_shd(vc,"EBF3FB")
            _ct(lc,lbl,bold=True,size=8,color="1F3864",align=WD_ALIGN_PARAGRAPH.LEFT)
            _ct(vc,val,size=8,color="333333",align=WD_ALIGN_PARAGRAPH.LEFT);_rh(sub.rows[i],int(0.42*567),"exact")
    footer=section.footer
    for p in list(footer.paragraphs):p.clear()
    fp=footer.paragraphs[0];fp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    rf=fp.add_run(f"SPySM — {date.today().strftime('%d/%m/%Y')} — Pág. ");rf.font.size=Pt(8);rf.font.name="Arial";rf.font.color.rgb=RGBColor(0x88,0x88,0x88)

    def tbl_df(dataframe):
        nc=len(dataframe.columns);cn=list(dataframe.columns)
        wc={"Descripción","Descripcion","Categoría","Falla","Sistema"}
        wn=sum(1 for c in cn if c in wc);nn=nc-wn
        if wn>0 and nn>0:nw=max(900,int(PAGE_W*0.08));ww=(PAGE_W-nw*nn)//wn
        elif wn>0:ww=PAGE_W//wn;nw=ww
        else:nw=PAGE_W//nc;ww=nw
        widths=[ww if c in wc else nw for c in cn];widths[-1]+=PAGE_W-sum(widths)
        t=doc.add_table(rows=1,cols=nc);t.style="Table Grid";t.autofit=False
        _tw(t,PAGE_W);_fix(t);_grid(t,widths);_brd(t)
        for i,(c,w) in enumerate(zip(cn,widths)):
            cell=t.rows[0].cells[i];_cw(cell,w);_shd(cell,"1F3864");_cm(cell,50,50,80,80);_va(cell)
            _ct(cell,c,bold=True,size=8,color="FFFFFF")
        _rh(t.rows[0],400)
        for ri,(_,row) in enumerate(dataframe.iterrows()):
            fill="EBF3FB" if ri%2==0 else "FFFFFF";tr=t.add_row();_rh(tr,320)
            for i,(val,w) in enumerate(zip(row,widths)):
                cell=tr.cells[i];_cw(cell,w);_shd(cell,fill);_cm(cell,40,40,80,80);_va(cell)
                txt="" if pd.isna(val) else str(val)
                al=WD_ALIGN_PARAGRAPH.LEFT if cn[i] in wc else WD_ALIGN_PARAGRAPH.CENTER
                _ct(cell,txt,size=8,align=al)
        return t

    sn=[0]
    def ns(title):
        sn[0]+=1;h=doc.add_heading(f"{sn[0]}. {title}",level=1)
        for r in h.runs:r.font.color.rgb=RGBColor(0x2E,0x75,0xB6);r.font.name="Arial"

    # Generar para cada MR seleccionado
    for mr_key in mr_sel:
        if mr_key == "General":
            df_w, do_w = df, df_out
            mr_title = "General"
        else:
            df_w = df[df['MR']==mr_key]
            do_w = df_out[df_out['MR']==mr_key] if 'MR' in df_out.columns and not df_out.empty else pd.DataFrame()
            mr_title = MR_LABELS.get(mr_key, mr_key)
        if df_w.empty: continue
        n_tot = int(df_w['CodItem_num'].sum())
        tiene_desv = not do_w.empty and "Desviacion" in do_w.columns

        # Separador por MR
        doc.add_paragraph()
        h = doc.add_heading(f"{'═'*3} {mr_title} {'═'*3}", level=0)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in h.runs: r.font.color.rgb = RGBColor(0x1F,0x38,0x64); r.font.name = "Arial"
        pm = doc.add_paragraph(f"Observaciones: {n_tot} | Vehículos: {df_w['Vehiculo'].dropna().nunique()}")
        pm.runs[0].font.size = Pt(10); doc.add_paragraph()

        if secciones.get("crit",True):
            ns("Criticidad"); d=df_w.groupby("CritAmpliado")['CodItem_num'].sum().sort_values(ascending=False).reset_index()
            d.columns=["Criticidad","Cantidad"];d["%"]=(d["Cantidad"]/n_tot*100).round(1).astype(str)+"%"
            tbl_df(d); doc.add_paragraph()
        if secciones.get("sistemas",True):
            ns("Sistemas"); d=df_w.groupby('SistemaUnidad')['CodItem_num'].sum().sort_values(ascending=False).reset_index()
            d.columns=['Código','Cantidad'];d['Sistema']=d['Código'].map(SISTEMA_LABELS).fillna(d['Código'])
            d['%']=(d['Cantidad']/n_tot*100).round(1).astype(str)+"%"; tbl_df(d[['Código','Sistema','Cantidad','%']]); doc.add_paragraph()
        if secciones.get("mensual",True):
            ns("Evolución Mensual"); md=[]
            for m in MONTH_ORDER:
                dm=df_w[df_w['Mes'].str.upper()==m];cnt=int(dm['CodItem_num'].sum());v=dm['Vehiculo'].dropna().nunique()
                if cnt>0:md.append({'Mes':m.title(),'Obs':cnt,'Veh':v,'Ratio':round(cnt/v,2) if v>0 else 0})
            if md: tbl_df(pd.DataFrame(md)); doc.add_paragraph()
        if secciones.get("top15",True):
            ns("Top 15 Fallas"); d=df_w.groupby("DescAgrupada")['CodItem_num'].sum().sort_values(ascending=False).head(15).reset_index()
            d.columns=["Descripción","Cantidad"];d["%"]=(d["Cantidad"]/n_tot*100).round(1).astype(str)+"%"; tbl_df(d); doc.add_paragraph()
        if tiene_desv and secciones.get("desvios",True):
            ns("Desvíos"); d=do_w.groupby("DescAgrupada").agg(Cant=("CodItem_num","sum"),Prom=("Desviacion",lambda x:round(x.mean(),2)),Max=("Desviacion",lambda x:round(x.abs().max(),2)),Min=("Desviacion",lambda x:round(x.abs().min(),2))).reset_index().sort_values("Cant",ascending=False)
            d.columns=["Categoría","Cant.","Prom","Máx","Mín"]; tbl_df(d); doc.add_paragraph()
        if secciones.get("pareto",True) and 'Clasificacion' in df_w.columns and df_w['Clasificacion'].notna().any():
            for cat in ['Fuera de rango','Ausencia de elementos','Mal estado']:
                dc=df_w[df_w['Clasificacion'].str.strip().str.lower()==cat.lower()]
                if dc.empty: continue
                ns(f"Pareto — {cat}"); c=dc.groupby('SistemaUnidad')['CodItem_num'].sum().sort_values(ascending=False).reset_index()
                c.columns=['Código','Cantidad'];c['Sistema']=c['Código'].map(SISTEMA_LABELS).fillna(c['Código'])
                c['%']=(c['Cantidad']/c['Cantidad'].sum()*100).round(1).astype(str)+"%"
                c['% Acum']=(c['Cantidad'].cumsum()/c['Cantidad'].sum()*100).round(1).astype(str)+"%"
                tbl_df(c[['Código','Sistema','Cantidad','%','% Acum']]); doc.add_paragraph()
        if tiene_desv and secciones.get("detalle",True):
            ns("Detalle Fuera de Parámetro"); cd=["Vehiculo","Mes","SistemaUnidad","Descripcion"];rn=["Vehículo","Mes","Sistema","Descripción"]
            for col in ["RefMin","RefMax","ValorRelevado","Desviacion"]:
                if col in do_w.columns:cd.append(col);rn.append({"RefMin":"Ref Min","RefMax":"Ref Max","ValorRelevado":"Relev.","Desviacion":"Desvío"}[col])
            cd.append("Criticidad");rn.append("Crit.");det=do_w[cd].copy();det.columns=rn
            if "Desvío" in det.columns:det["Desvío"]=det["Desvío"].round(2)
            if "Relev." in det.columns:det["Relev."]=det["Relev."].round(1)
            tbl_df(det); doc.add_paragraph()
        if secciones.get("top10mr",True) and mr_key=="General":
            ns("Top 10 por Tipo de MR")
            for mrc,ml in [('LOC','Locomotoras'),('CCRR','Coches Remolcados'),('CCEE','Coches Eléctricos'),('CCMM','Coche Motor')]:
                dm=df_w[df_w['MR']==mrc]
                if dm.empty:continue
                h2=doc.add_heading(f"{ml} ({mrc})",level=2)
                for r in h2.runs:r.font.name="Arial"
                dm=dm.copy();dm['_u']=dm['Modulo'].apply(lambda x:None if(x is None or str(x).strip() in('','0','nan'))else str(x).strip()).fillna(dm['Vehiculo'].astype(str).str.strip())
                res=dm.groupby('_u')['CodItem_num'].sum().sort_values(ascending=False).head(10).reset_index();res.columns=['Unidad','Obs.']
                tbl_df(res);doc.add_paragraph()
        if secciones.get("concl",True):
            ns("Conclusiones"); doc.add_paragraph("Del análisis se extraen:"); doc.add_paragraph()
            concl=[]
            sv=df_w.groupby("SistemaUnidad")['CodItem_num'].sum().sort_values(ascending=False).head(3)
            items=[f"{SISTEMA_LABELS.get(s,s)} ({int(c)}, {round(c/n_tot*100,1)}%)" for s,c in sv.items()]
            concl.append(f"Sistemas principales: {', '.join(items)}.")
            df_w2=df_w.copy();df_w2['_u']=df_w2['Modulo'].apply(lambda x:None if(x is None or str(x).strip() in("","0","nan"))else str(x).strip()).fillna(df_w2['Vehiculo'].astype(str).str.strip())
            uv=df_w2.groupby("_u")['CodItem_num'].sum().sort_values(ascending=False).head(5)
            items=[f"{u} ({int(c)} obs.)" for u,c in uv.items()]; concl.append(f"Unidades prioritarias: {', '.join(items)}.")
            tc=int(df_w.loc[df_w['Criticidad']=='C','CodItem_num'].sum());tr=int(df_w.loc[df_w['Criticidad']=='R','CodItem_num'].sum())
            concl.append(f"Criticidad alta: {round((tc+tr)/n_tot*100,1)}% ({tc} C, {tr} R).")
            fv=df_w.groupby("DescAgrupada")['CodItem_num'].sum().sort_values(ascending=False).head(3)
            items=[f"'{f}' ({int(n)}, {round(n/n_tot*100,1)}%)" for f,n in fv.items()]; concl.append(f"Fallas recurrentes: {', '.join(items)}.")
            for texto in concl:
                p=doc.add_paragraph(style="List Number");r=p.add_run(texto);r.font.size=Pt(10);r.font.name="Arial"

    buf=io.BytesIO();doc.save(buf);buf.seek(0);return buf.getvalue()


def generar_excel(df, df_out):
    buf=io.BytesIO(); n_tot=int(df['CodItem_num'].sum())
    tiene_desv=not df_out.empty and "Desviacion" in df_out.columns
    with pd.ExcelWriter(buf,engine='openpyxl') as w:
        df.groupby("CritAmpliado")['CodItem_num'].sum().sort_values(ascending=False).reset_index().rename(columns={'CritAmpliado':'Criticidad','CodItem_num':'Cantidad'}).to_excel(w,sheet_name="Criticidad",index=False)
        d=df.groupby('SistemaUnidad')['CodItem_num'].sum().sort_values(ascending=False).reset_index();d.columns=['Código','Cantidad'];d['Sistema']=d['Código'].map(SISTEMA_LABELS).fillna(d['Código']);d.to_excel(w,sheet_name="Sistemas",index=False)
        md=[]
        for m in MONTH_ORDER:
            dm=df[df['Mes'].str.upper()==m];cnt=int(dm['CodItem_num'].sum());v=dm['Vehiculo'].dropna().nunique()
            if cnt>0:md.append({'Mes':m.title(),'Obs':cnt,'Veh':v,'Ratio':round(cnt/v,2) if v>0 else 0})
        if md:pd.DataFrame(md).to_excel(w,sheet_name="Mensual",index=False)
        df.groupby("DescAgrupada")['CodItem_num'].sum().sort_values(ascending=False).head(15).reset_index().rename(columns={'DescAgrupada':'Falla','CodItem_num':'Cantidad'}).to_excel(w,sheet_name="Top15",index=False)
        if tiene_desv:df_out.groupby("DescAgrupada").agg(Cant=("CodItem_num","sum"),Prom=("Desviacion",lambda x:round(x.mean(),2)),Max=("Desviacion",lambda x:round(x.abs().max(),2))).reset_index().to_excel(w,sheet_name="Desvíos",index=False)
        if 'Clasificacion' in df.columns:
            for cat in ['Fuera de rango','Ausencia de elementos','Mal estado']:
                dc=df[df['Clasificacion'].str.strip().str.lower()==cat.lower()]
                if dc.empty:continue
                c=dc.groupby('SistemaUnidad')['CodItem_num'].sum().sort_values(ascending=False).reset_index();c.columns=['Sistema','Cantidad']
                c['%Acum']=(c['Cantidad'].cumsum()/c['Cantidad'].sum()*100).round(1);c.to_excel(w,sheet_name=f"P.{cat[:13]}",index=False)
        for mrc,ml in[('LOC','LOC'),('CCRR','CCRR'),('CCEE','CCEE'),('CCMM','CCMM')]:
            dm=df[df['MR']==mrc]
            if dm.empty:continue
            dm=dm.copy();dm['_u']=dm['Modulo'].apply(lambda x:None if(x is None or str(x).strip() in('','0','nan'))else str(x).strip()).fillna(dm['Vehiculo'].astype(str).str.strip())
            dm.groupby('_u')['CodItem_num'].sum().sort_values(ascending=False).head(10).reset_index().rename(columns={'_u':'Unidad','CodItem_num':'Obs'}).to_excel(w,sheet_name=f"Top10_{ml}",index=False)
        df.pivot_table(index='CritAmpliado',columns='SistemaUnidad',values='CodItem_num',aggfunc='sum',fill_value=0).to_excel(w,sheet_name="Cruzada")
        df.to_excel(w,sheet_name="Datos",index=False)
    buf.seek(0);return buf.getvalue()

# ═══════════════════════════════════════════════
# SIDEBAR + PANTALLA PRINCIPAL
# ═══════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🔧 Análisis de Mantenimiento")
    st.markdown("---")
    uploaded = st.file_uploader("Subir archivo Excel", type=["xlsx"], help="Archivo de inspecciones")
    st.markdown("---")
    st.markdown("**Cada tab tiene sub-solapas por MR**\ncon filtro de Modelo y Meses.")
    st.markdown("---")
    st.caption("Análisis automático · Material Rodante")

st.title("🔧 Análisis de Informes de Mantenimiento")
if uploaded is None:
    st.markdown("""<div style="background:linear-gradient(135deg,#1a2235,#1e2a3a);border:1px dashed #2a3a50;border-radius:12px;padding:60px;text-align:center;margin-top:40px">
    <div style="font-size:3rem;margin-bottom:16px">📂</div>
    <div style="color:#4fc3f7;font-size:1.3rem;font-weight:700">Subí tu archivo Excel para comenzar</div></div>""", unsafe_allow_html=True)
    st.stop()

with st.spinner("Procesando..."): df, df_out = cargar_y_analizar(uploaded.read())

total_obs=int(df['CodItem_num'].sum()); total_veh=df['Vehiculo'].dropna().nunique()
total_normal=int(df.loc[df['Criticidad']=='N','CodItem_num'].sum())
total_crit=int(df.loc[df['Criticidad']=='C','CodItem_num'].sum())
total_rech=int(df.loc[df['Criticidad']=='R','CodItem_num'].sum())
total_corr=int(df.loc[df['Criticidad']=='O','CodItem_num'].sum())
total_nrc=int(df.loc[df['Criticidad']=='NRC','CodItem_num'].sum())
pct_alta=round((total_crit+total_rech)/total_obs*100,1) if total_obs>0 else 0

st.markdown('<div class="section-header">Resumen del período</div>', unsafe_allow_html=True)
c1,c2,c3,c4=st.columns(4)
c1.markdown(kpi("Total Observaciones",total_obs),unsafe_allow_html=True)
c2.markdown(kpi("Vehículos",total_veh),unsafe_allow_html=True)
c3.markdown(kpi("Sin Obs.",total_nrc),unsafe_allow_html=True)
c4.markdown(kpi("% Crit.Alta",f"{pct_alta}%","danger"),unsafe_allow_html=True)
st.markdown("<br>",unsafe_allow_html=True)
d1,d2,d3,d4=st.columns(4)
d1.markdown(kpi("Normales",total_normal,"success"),unsafe_allow_html=True)
d2.markdown(kpi("Corregidas",total_corr,"success"),unsafe_allow_html=True)
d3.markdown(kpi("Críticas",total_crit,"danger"),unsafe_allow_html=True)
d4.markdown(kpi("Rechazadas",total_rech,"warning"),unsafe_allow_html=True)
st.markdown("<br>",unsafe_allow_html=True)

# ── TABS PRINCIPALES — todos con sub-tabs MR ──
tab1,tab2,tab3,tab4,tab5,tab6=st.tabs(["📊 Resúmenes","📈 Gráficos","🔍 Desvíos","🧩 Clasificación","🔬 Explorador","📄 Exportar"])

with tab1: render_with_mr_subtabs(df, df_out, "t1", render_resumenes)
with tab2: render_with_mr_subtabs(df, df_out, "t2", render_graficos)
with tab3: render_with_mr_subtabs(df, df_out, "t3", render_desvios)
with tab4: render_with_mr_subtabs(df, df_out, "t4", render_clasificacion)
with tab5: render_with_mr_subtabs(df, df_out, "t5", render_explorador)

with tab6:
    st.markdown("#### Exportar Informe")
    st.markdown("##### Encabezado")
    h1,h2=st.columns(2)
    with h1:
        hdr_codigo=st.text_input("Código",placeholder="SGBV-INF-2025-001")
        hdr_version=st.text_input("Versión",value="v1.0")
        hdr_linea=st.text_input("Línea / Contrato",placeholder="Línea San Martín — 3-LA")
    with h2:
        logo_file=st.file_uploader("Logo (JPG/PNG)",type=["jpg","jpeg","png"])
        hdr_subger=st.text_input("Subgerencia",value="Programación y Seguimiento de Mantenimiento")
    st.markdown("---")

    st.markdown("##### Secciones a incluir")
    s1,s2,s3=st.columns(3)
    with s1: ic=st.checkbox("Criticidad",True); isi=st.checkbox("Sistemas",True); im=st.checkbox("Mensual",True)
    with s2: it=st.checkbox("Top 15 Fallas",True); idv=st.checkbox("Desvíos",True); ip=st.checkbox("Pareto",True)
    with s3: idt=st.checkbox("Detalle desvíos",True); imr=st.checkbox("Top10 MR (solo General)",True); icn=st.checkbox("Conclusiones",True)
    st.markdown("---")

    st.markdown("##### ¿Para qué MR generar el informe?")
    mr_opciones = ["General"] + get_mr_list(df)
    mr_sel = st.multiselect("Seleccioná los MR", mr_opciones, default=["General"],
                            help="Se genera una sección completa por cada MR seleccionado")
    st.markdown("---")

    b1,b2=st.columns(2)
    with b1:
        if st.button("📄 Generar Word",type="primary"):
            lb=logo_file.read() if logo_file else None
            cfg=dict(codigo=hdr_codigo,version=hdr_version,linea=hdr_linea,subger=hdr_subger,logo=lb,
                mr_sel=mr_sel,
                secciones=dict(crit=ic,sistemas=isi,mensual=im,top15=it,desvios=idv,pareto=ip,detalle=idt,top10mr=imr,concl=icn))
            with st.spinner("Generando Word..."): wb=generar_word(df,df_out,cfg)
            st.success("✅ Word generado")
            ls=(hdr_linea or"Informe").replace(" ","_").replace("/","_")[:30]
            st.download_button("⬇️ Descargar .docx",wb,f"Informe_{ls}.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with b2:
        if st.button("📊 Generar Excel",type="secondary"):
            with st.spinner("Generando Excel..."): eb=generar_excel(df,df_out)
            st.success("✅ Excel generado")
            ls=(hdr_linea or"Datos").replace(" ","_").replace("/","_")[:30]
            st.download_button("⬇️ Descargar .xlsx",eb,f"Datos_{ls}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
