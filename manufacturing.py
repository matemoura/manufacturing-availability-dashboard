import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Dashboard de Disponibilidade",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────
#  ESTILOS GLOBAIS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

    .stApp { background: #0f1117; color: #e2e8f0; }

    [data-testid="stSidebar"] {
        background: #161b27 !important;
        border-right: 1px solid #1e2d45;
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] div { color: #c9d4e8 !important; }

    h1 { color: #e2e8f0 !important; font-weight: 700 !important; letter-spacing: -0.5px; }
    h2 { color: #cbd5e1 !important; font-weight: 600 !important; }
    h3 { color: #94a3b8 !important; font-weight: 500 !important; }

    .kpi-card {
        background: #161b27;
        border: 1px solid #1e2d45;
        border-radius: 14px;
        padding: 22px 18px;
        text-align: center;
        position: relative;
        overflow: hidden;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .kpi-card:hover { transform: translateY(-2px); box-shadow: 0 8px 32px rgba(0,0,0,0.4); }
    .kpi-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
        border-radius: 14px 14px 0 0;
    }
    .kpi-card.green::before  { background: linear-gradient(90deg,#10b981,#34d399); }
    .kpi-card.orange::before { background: linear-gradient(90deg,#f59e0b,#fbbf24); }
    .kpi-card.red::before    { background: linear-gradient(90deg,#ef4444,#f87171); }
    .kpi-card.blue::before   { background: linear-gradient(90deg,#3b82f6,#60a5fa); }
    .kpi-card.purple::before { background: linear-gradient(90deg,#8b5cf6,#a78bfa); }

    .kpi-icon  { font-size:1.4rem; margin-bottom:6px; display:block; }
    .kpi-value {
        font-family: 'DM Mono', monospace;
        font-size: 1.9rem;
        font-weight: 600;
        line-height: 1.1;
        margin: 6px 0 4px;
        letter-spacing: -1px;
    }
    .kpi-value.green  { color: #34d399; }
    .kpi-value.orange { color: #fbbf24; }
    .kpi-value.red    { color: #f87171; }
    .kpi-value.blue   { color: #60a5fa; }
    .kpi-value.purple { color: #a78bfa; }
    .kpi-label { font-size:0.72rem; text-transform:uppercase; letter-spacing:1.5px; color:#64748b; font-weight:500; }
    .kpi-sub   { font-size:0.7rem; color:#475569; margin-top:4px; }

    .section-title {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 2px;
        color: #475569;
        font-weight: 600;
        margin: 28px 0 12px;
        border-left: 3px solid #3b82f6;
        padding-left: 10px;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background: #161b27;
        border-radius: 10px;
        padding: 5px;
        border: 1px solid #1e2d45;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 7px;
        color: #64748b;
        font-size: 0.85rem;
        font-weight: 500;
        padding: 8px 18px;
    }
    .stTabs [aria-selected="true"] {
        background: #1e3a5f !important;
        color: #60a5fa !important;
    }

    .info-box {
        background: #0f2744;
        border: 1px solid #1e3a5f;
        border-left: 4px solid #3b82f6;
        border-radius: 8px;
        padding: 12px 16px;
        font-size: 0.85rem;
        color: #93c5fd;
        margin: 8px 0;
    }
    .warn-box {
        background: #2d1a06;
        border: 1px solid #7c3a0a;
        border-left: 4px solid #f59e0b;
        border-radius: 8px;
        padding: 12px 16px;
        font-size: 0.85rem;
        color: #fcd34d;
        margin: 8px 0;
    }

    .stDownloadButton button {
        background: linear-gradient(135deg,#1e3a5f,#2563eb) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 10px 24px !important;
    }
    .stDownloadButton button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 16px rgba(37,99,235,0.4) !important;
    }

    hr { border-color: #1e2d45 !important; margin: 24px 0 !important; }
    [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  CONSTANTES E UTILITÁRIOS
# ─────────────────────────────────────────────

DIAS_PT = {0:'Segunda',1:'Terça',2:'Quarta',3:'Quinta',4:'Sexta',5:'Sábado',6:'Domingo'}


def decimal_para_hhmm(horas_decimal: float) -> str:
    """Converte horas decimais (ex: 1.75) para 'Xh Ym' (ex: '1h 45m')"""
    if pd.isna(horas_decimal) or horas_decimal < 0:
        return "0h"
    horas = int(horas_decimal)
    minutos = round((horas_decimal - horas) * 60)
    if minutos == 60:
        horas += 1
        minutos = 0
    if minutos == 0:
        return f"{horas}h"
    return f"{horas}h {minutos:02d}m"


def horas_jornada_dia(data, horas_padrao: float, horas_sexta: float) -> float:
    dow = pd.Timestamp(data).dayofweek
    if dow == 4:   # Sexta
        return horas_sexta
    if dow >= 5:   # Fim de semana
        return 0.0
    return horas_padrao


def categorizar_atividade(row):
    tipo = str(row.get('Tipo de registro do diário', ''))
    descricao = str(row.get('Descrição', ''))
    if tipo == 'Processo':
        return 'Tempo Produtivo'
    if tipo in ('Atividade indireta', 'Intervalo') or 'Setup' in descricao:
        return 'Paradas Planejadas'
    if tipo == 'Interrupção' and row.get('Duracao_Horas', 0) > 0:
        return 'Paradas Não Planejadas'
    return 'Outros'


def processar_dados(df, horas_padrao: float, horas_sexta: float):
    df = df.copy()
    df.columns = df.columns.str.strip()

    df['Data do perfil'] = pd.to_datetime(df['Data do perfil'], dayfirst=True, errors='coerce')

    for col in ['Hora inicial', 'Hora final']:
        parsed = pd.to_datetime(df[col], format='%I:%M:%S %p', errors='coerce')
        if parsed.isna().all():
            parsed = pd.to_datetime(df[col], format='%H:%M:%S', errors='coerce')
        df[col] = parsed.dt.time

    df.dropna(subset=['Hora inicial', 'Hora final', 'Data do perfil'], inplace=True)

    df['Inicio_ts'] = df.apply(
        lambda r: pd.Timestamp.combine(r['Data do perfil'].date(), r['Hora inicial']), axis=1)
    df['Fim_ts'] = df.apply(
        lambda r: pd.Timestamp.combine(r['Data do perfil'].date(), r['Hora final']), axis=1)
    df.loc[df['Fim_ts'] < df['Inicio_ts'], 'Fim_ts'] += pd.Timedelta(days=1)
    df['Duracao_Horas'] = (df['Fim_ts'] - df['Inicio_ts']).dt.total_seconds() / 3600

    df['Categoria'] = df.apply(categorizar_atividade, axis=1)
    df['Data'] = df['Data do perfil'].dt.date
    df['DiaSemana'] = df['Data do perfil'].dt.dayofweek
    df['DiaSemana_Nome'] = df['DiaSemana'].map(DIAS_PT)
    df['Jornada_Dia'] = df['Data'].apply(lambda d: horas_jornada_dia(d, horas_padrao, horas_sexta))

    for col in ['N° oper.', 'Referência', 'Ident. do trabalho']:
        if col not in df.columns:
            df[col] = ''

    # Garantir string limpa
    df['Referência'] = df['Referência'].fillna('').astype(str).str.strip()
    df['N° oper.'] = df['N° oper.'].fillna('').astype(str).str.strip()

    return df


def calcular_disponibilidade(df):
    df_valido = df[df['Categoria'] != 'Outros'].copy()

    pivot = df_valido.pivot_table(
        index='Nome', columns='Categoria', values='Duracao_Horas',
        aggfunc='sum', fill_value=0
    ).reset_index()
    for cat in ['Tempo Produtivo', 'Paradas Planejadas', 'Paradas Não Planejadas']:
        if cat not in pivot.columns:
            pivot[cat] = 0.0

    jornada_pessoa = (
        df[['Nome','Data','Jornada_Dia']]
        .drop_duplicates(subset=['Nome','Data'])
        .groupby('Nome')['Jornada_Dia'].sum()
        .reset_index()
        .rename(columns={'Jornada_Dia':'Jornada_Total'})
    )
    pivot = pivot.merge(jornada_pessoa, on='Nome', how='left')
    tempo_prog = pivot['Jornada_Total'] - pivot['Paradas Planejadas']
    pivot['Disponibilidade (%)'] = np.where(
        tempo_prog <= 0, 0,
        (pivot['Tempo Produtivo'] / np.where(tempo_prog <= 0, 1, tempo_prog) * 100)
    ).round(2)

    # Por data
    pivot_data = df_valido.pivot_table(
        index='Data', columns='Categoria', values='Duracao_Horas',
        aggfunc='sum', fill_value=0
    ).reset_index()
    for cat in ['Tempo Produtivo', 'Paradas Planejadas', 'Paradas Não Planejadas']:
        if cat not in pivot_data.columns:
            pivot_data[cat] = 0.0

    num_pessoas_dia = (df.groupby('Data')['Nome'].nunique()
                       .reset_index().rename(columns={'Nome':'Num_Pessoas'}))
    jornada_dia = (df[['Data','Jornada_Dia']].drop_duplicates()
                   .rename(columns={'Jornada_Dia':'Jornada_Padrao_Dia'}))
    pivot_data = pivot_data.merge(num_pessoas_dia, on='Data').merge(jornada_dia, on='Data', how='left')
    pivot_data['DiaSemana'] = pd.to_datetime(pivot_data['Data']).dt.dayofweek
    pivot_data['DiaSemana_Nome'] = pivot_data['DiaSemana'].map(DIAS_PT)

    tp_dia = pivot_data['Jornada_Padrao_Dia'] * pivot_data['Num_Pessoas'] - pivot_data['Paradas Planejadas']
    pivot_data['Disponibilidade (%)'] = np.where(
        tp_dia <= 0, 0,
        (pivot_data['Tempo Produtivo'] / np.where(tp_dia <= 0, 1, tp_dia) * 100)
    ).round(2)

    return pivot, pivot_data


def calcular_indicadores_extras(df):
    df_proc = df[df['Categoria'] == 'Tempo Produtivo'].copy()

    # Resumo por OP
    horas_op = df_proc.groupby('Referência')['Duracao_Horas'].sum().reset_index()
    horas_op.columns = ['Ordem de Produção', 'Horas Trabalhadas']
    horas_op = horas_op[horas_op['Ordem de Produção'].str.strip() != '']
    horas_op = horas_op.sort_values('Horas Trabalhadas', ascending=False).reset_index(drop=True)
    horas_op['Horas (h:m)'] = horas_op['Horas Trabalhadas'].apply(decimal_para_hhmm)

    # OP + Operação detalhado
    op_group_cols = [c for c in ['Referência','N° oper.','Ident. do trabalho','Descrição']
                     if c in df_proc.columns]
    horas_op_oper = df_proc.groupby(op_group_cols)['Duracao_Horas'].sum().reset_index()
    horas_op_oper.columns = op_group_cols + ['Horas Trabalhadas']
    horas_op_oper = horas_op_oper[horas_op_oper['Referência'].str.strip() != '']
    horas_op_oper = horas_op_oper.sort_values(
        ['Referência', 'N° oper.'], ascending=[True, True]).reset_index(drop=True)
    horas_op_oper['Horas (h:m)'] = horas_op_oper['Horas Trabalhadas'].apply(decimal_para_hhmm)

    # Atividades
    tempo_atividade = df_proc.groupby('Descrição')['Duracao_Horas'].agg(
        ['sum','mean','count']).reset_index()
    tempo_atividade.columns = ['Atividade','Total (h)','Média (h)','Ocorrências']
    tempo_atividade = tempo_atividade.sort_values('Total (h)', ascending=False).reset_index(drop=True)
    tempo_atividade['Total (h:m)'] = tempo_atividade['Total (h)'].apply(decimal_para_hhmm)
    tempo_atividade['Média (h:m)'] = tempo_atividade['Média (h)'].apply(decimal_para_hhmm)

    # Paradas
    df_par = df[df['Categoria'].isin(['Paradas Planejadas','Paradas Não Planejadas'])].copy()
    paradas = df_par.groupby(['Descrição','Categoria'])['Duracao_Horas'].sum().reset_index()
    paradas.columns = ['Descrição','Tipo','Duração Total (h)']
    paradas = paradas[paradas['Duração Total (h)'] > 0]
    paradas = paradas.sort_values('Duração Total (h)', ascending=False).reset_index(drop=True)
    paradas['Duração (h:m)'] = paradas['Duração Total (h)'].apply(decimal_para_hhmm)

    return horas_op, horas_op_oper, tempo_atividade, paradas


def plotly_dark(**kwargs):
    base = dict(
        plot_bgcolor='#0f1117',
        paper_bgcolor='#0f1117',
        font=dict(family='DM Sans', color='#94a3b8', size=12),
        xaxis=dict(gridcolor='#1e2d45', linecolor='#1e2d45',
                   tickfont=dict(color='#64748b'), title_font=dict(color='#94a3b8')),
        yaxis=dict(gridcolor='#1e2d45', linecolor='#1e2d45',
                   tickfont=dict(color='#64748b'), title_font=dict(color='#94a3b8')),
        title_font=dict(color='#e2e8f0', size=14, family='DM Sans'),
        legend=dict(bgcolor='#161b27', bordercolor='#1e2d45', borderwidth=1,
                    font=dict(color='#94a3b8')),
        margin=dict(t=50, b=40, l=40, r=20),
    )
    base.update(kwargs)
    return base


def gerar_excel(tabela_pessoa, tabela_dia, paradas, horas_op, horas_op_oper,
                tempo_atividade, horas_padrao, horas_sexta):
    wb = Workbook()
    hf = PatternFill("solid", fgColor="0F2744")
    ff = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    tf = Font(bold=True, color="1E3A5F", name="Calibri", size=12)
    af = PatternFill("solid", fgColor="EBF4FF")
    bd = Border(left=Side(style='thin',color='CCCCCC'), right=Side(style='thin',color='CCCCCC'),
                bottom=Side(style='thin',color='CCCCCC'), top=Side(style='thin',color='CCCCCC'))
    ct = Alignment(horizontal='center', vertical='center', wrap_text=True)

    def estilizar(ws, df, start_row, title):
        ws.cell(row=start_row, column=1, value=title).font = tf
        ws.row_dimensions[start_row].height = 22
        hr = start_row + 1
        for j, col in enumerate(df.columns, 1):
            c = ws.cell(row=hr, column=j, value=col)
            c.fill = hf; c.font = ff; c.alignment = ct; c.border = bd
        ws.row_dimensions[hr].height = 22
        for i, row in enumerate(df.itertuples(index=False), start=hr+1):
            fill = af if i % 2 == 0 else None
            for j, val in enumerate(row, 1):
                c = ws.cell(row=i, column=j, value=val)
                c.border = bd; c.alignment = ct
                if fill: c.fill = fill
                if isinstance(val, float): c.number_format = '0.00'
            ws.row_dimensions[i].height = 18
        for j, col in enumerate(df.columns, 1):
            try:
                mx = max(df[col].astype(str).map(len).max(), len(str(col))) + 4
            except Exception:
                mx = 20
            ws.column_dimensions[get_column_letter(j)].width = min(mx, 45)

    ws1 = wb.active
    ws1.title = "Resumo"
    ws1.sheet_view.showGridLines = False
    ws1['A1'] = f"Dashboard de Disponibilidade  |  Jornada: {horas_padrao}h  |  Sexta: {horas_sexta}h"
    ws1['A1'].font = Font(bold=True, color="FFFFFF", name="Calibri", size=13)
    ws1['A1'].fill = PatternFill("solid", fgColor="0F2744")
    ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws1.merge_cells('A1:G1')
    ws1.row_dimensions[1].height = 30
    estilizar(ws1, tabela_pessoa.round(2), 3, "Disponibilidade por Pessoa")

    ws2 = wb.create_sheet("Por Data")
    ws2.sheet_view.showGridLines = False
    estilizar(ws2, tabela_dia.round(2), 2, "Disponibilidade por Data")

    ws3 = wb.create_sheet("Paradas")
    ws3.sheet_view.showGridLines = False
    estilizar(ws3, paradas.round(2), 2, "Análise de Paradas")

    ws4 = wb.create_sheet("Horas por OP")
    ws4.sheet_view.showGridLines = False
    estilizar(ws4, horas_op.round(2), 2, "Horas por Ordem de Produção")

    ws5 = wb.create_sheet("OP por Operação")
    ws5.sheet_view.showGridLines = False
    estilizar(ws5, horas_op_oper.round(2), 2, "Horas por OP e Operação")

    ws6 = wb.create_sheet("Atividades")
    ws6.sheet_view.showGridLines = False
    estilizar(ws6, tempo_atividade.round(2), 2, "Tempo por Atividade Produtiva")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
#  INTERFACE
# ─────────────────────────────────────────────

st.markdown("""
<div style="display:flex;align-items:center;gap:14px;margin-bottom:8px">
  <span style="font-size:2.2rem">🏭</span>
  <div>
    <div style="font-size:1.6rem;font-weight:700;color:#e2e8f0;letter-spacing:-0.5px">
      Dashboard de Disponibilidade
    </div>
    <div style="font-size:0.78rem;color:#475569;letter-spacing:1.5px;text-transform:uppercase">
      Produção · Apontamentos · OEE
    </div>
  </div>
</div>
""", unsafe_allow_html=True)
st.markdown('<hr>', unsafe_allow_html=True)

# ── Sidebar ──
with st.sidebar:
    st.markdown("### ⚙️ Configurações")
    st.markdown("---")
    horas_padrao = st.number_input(
        "Jornada padrão — Seg a Qui (h)",
        min_value=1.0, max_value=24.0, value=8.0, step=0.5)
    horas_sexta = st.number_input(
        "Jornada Sexta-feira (h)",
        min_value=1.0, max_value=24.0, value=7.0, step=0.5,
        help="Sexta costuma ter 1h a menos")
    st.markdown("---")
    arquivo = st.file_uploader("📂 Planilha de apontamentos (.xlsx)", type=["xlsx"])
    st.markdown("---")
    st.markdown("""
<div style="font-size:0.78rem;color:#475569;line-height:1.8">
<b style="color:#60a5fa">Categorias:</b><br>
🟢 <b>Produtivo</b> — Tipo = Processo<br>
🟡 <b>Plan.</b> — Indireta / Intervalo / Setup<br>
🔴 <b>N.Plan.</b> — Interrupção > 0h<br>
⚫ <b>Outros</b> — Entradas/Saídas
</div>
""", unsafe_allow_html=True)

# ── Conteúdo ──
if arquivo:
    try:
        df_raw = pd.read_excel(arquivo)
        df = processar_dados(df_raw, horas_padrao, horas_sexta)

        # Filtros de período
        datas = sorted(df['Data'].unique())
        col_f1, col_f2 = st.columns(2)
        data_inicio = col_f1.date_input("📅 Data inicial", value=datas[0],
                                        min_value=datas[0], max_value=datas[-1])
        data_fim = col_f2.date_input("📅 Data final", value=datas[-1],
                                     min_value=datas[0], max_value=datas[-1])

        df_f = df[(df['Data'] >= data_inicio) & (df['Data'] <= data_fim)].copy()

        if df_f.empty:
            st.markdown('<div class="warn-box">⚠️ Nenhum dado para o período selecionado.</div>',
                        unsafe_allow_html=True)
            st.stop()

        tabela_pessoa_f, tabela_dia_f = calcular_disponibilidade(df_f)
        horas_op_f, horas_op_oper_f, tempo_atividade_f, paradas_f = calcular_indicadores_extras(df_f)

        # KPIs
        tp = tabela_pessoa_f['Tempo Produtivo'].sum()
        pp = tabela_pessoa_f['Paradas Planejadas'].sum()
        pnp = tabela_pessoa_f['Paradas Não Planejadas'].sum()
        n_pess = len(tabela_pessoa_f)
        jornada_tot = tabela_pessoa_f['Jornada_Total'].sum() if 'Jornada_Total' in tabela_pessoa_f else horas_padrao * n_pess
        tp_prog = jornada_tot - pp
        disp = round(tp / tp_prog * 100, 2) if tp_prog > 0 else 0
        cor = "green" if disp >= 85 else "orange" if disp >= 70 else "red"
        n_dias = (pd.Timestamp(data_fim) - pd.Timestamp(data_inicio)).days + 1

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.markdown(f"""<div class="kpi-card {cor}">
            <span class="kpi-icon">📊</span>
            <div class="kpi-value {cor}">{disp:.1f}%</div>
            <div class="kpi-label">Disponibilidade</div>
            <div class="kpi-sub">Meta: 85%</div></div>""", unsafe_allow_html=True)
        k2.markdown(f"""<div class="kpi-card blue">
            <span class="kpi-icon">👷</span>
            <div class="kpi-value blue">{n_pess}</div>
            <div class="kpi-label">Pessoas</div>
            <div class="kpi-sub">{n_dias} dia(s)</div></div>""", unsafe_allow_html=True)
        k3.markdown(f"""<div class="kpi-card green">
            <span class="kpi-icon">⚡</span>
            <div class="kpi-value green">{decimal_para_hhmm(tp)}</div>
            <div class="kpi-label">Tempo Produtivo</div>
            <div class="kpi-sub">{tp:.2f}h</div></div>""", unsafe_allow_html=True)
        k4.markdown(f"""<div class="kpi-card orange">
            <span class="kpi-icon">⏸️</span>
            <div class="kpi-value orange">{decimal_para_hhmm(pp)}</div>
            <div class="kpi-label">Paradas Plan.</div>
            <div class="kpi-sub">{pp:.2f}h</div></div>""", unsafe_allow_html=True)
        k5.markdown(f"""<div class="kpi-card red">
            <span class="kpi-icon">🚨</span>
            <div class="kpi-value red">{decimal_para_hhmm(pnp)}</div>
            <div class="kpi-label">N. Planejadas</div>
            <div class="kpi-sub">{pnp:.2f}h</div></div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ABAS
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📊 Disponibilidade",
            "📅 Por Data",
            "⏸️ Paradas",
            "🔧 Ordens de Produção",
            "🔩 OP + Operações",
            "📋 Dados Brutos"
        ])

        # ── TAB 1 — Disponibilidade ──
        with tab1:
            c1, c2 = st.columns([3, 2])
            with c1:
                cores_b = ['#10b981' if v >= 85 else '#f59e0b' if v >= 70 else '#ef4444'
                           for v in tabela_pessoa_f['Disponibilidade (%)']]
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=tabela_pessoa_f['Nome'], y=tabela_pessoa_f['Disponibilidade (%)'],
                    marker_color=cores_b,
                    text=[f"{v:.1f}%" for v in tabela_pessoa_f['Disponibilidade (%)']],
                    textposition='outside', textfont=dict(color='#e2e8f0', size=11)
                ))
                fig.add_hline(y=85, line_dash="dash", line_color="#34d399",
                              annotation_text="Meta 85%", annotation_font_color="#34d399",
                              annotation_position="right")
                layout_base = plotly_dark()
                layout_base['yaxis'].update({
                                                'range': [0,115],
                                                'title': "(%)"
                                            })
                fig.update_layout(
                                    title="Disponibilidade por Pessoa",
                                    xaxis_tickangle=-30,
                                    height=420,
                                    **layout_base
                                )
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig2 = go.Figure(go.Pie(
                    labels=['Tempo Produtivo','Paradas Planejadas','Paradas N.Plan.'],
                    values=[tp, pp, pnp],
                    marker_colors=['#10b981','#f59e0b','#ef4444'],
                    hole=0.5, textinfo='label+percent',
                    textfont=dict(color='#e2e8f0', size=11)
                ))
                fig2.update_layout(title="Composição da Jornada",
                                   height=420, showlegend=False, **plotly_dark())
                st.plotly_chart(fig2, use_container_width=True)

            tbl = tabela_pessoa_f[['Nome','Tempo Produtivo','Paradas Planejadas',
                                    'Paradas Não Planejadas','Disponibilidade (%)']].copy()
            tbl['Tempo Produtivo'] = tbl['Tempo Produtivo'].apply(decimal_para_hhmm)
            tbl['Paradas Planejadas'] = tbl['Paradas Planejadas'].apply(decimal_para_hhmm)
            tbl['Paradas Não Planejadas'] = tbl['Paradas Não Planejadas'].apply(decimal_para_hhmm)
            tbl['Disponibilidade (%)'] = tabela_pessoa_f['Disponibilidade (%)'].apply(lambda v: f"{v:.1f}%")
            st.dataframe(tbl, use_container_width=True, hide_index=True)

        # ── TAB 2 — Por Data ──
        with tab2:
            if len(tabela_dia_f) > 1:
                fig_ev = go.Figure()
                fig_ev.add_trace(go.Scatter(
                    x=tabela_dia_f['Data'].astype(str),
                    y=tabela_dia_f['Disponibilidade (%)'],
                    mode='lines+markers',
                    line=dict(color='#3b82f6', width=2.5, shape='spline'),
                    marker=dict(
                        color=['#10b981' if v >= 85 else '#f59e0b' if v >= 70 else '#ef4444'
                               for v in tabela_dia_f['Disponibilidade (%)']],
                        size=10, line=dict(width=2, color='#0f1117')
                    ),
                    text=[f"{r.DiaSemana_Nome}<br>{v:.1f}%"
                          for v, r in zip(tabela_dia_f['Disponibilidade (%)'],
                                          tabela_dia_f.itertuples())],
                    hovertemplate='%{text}<extra></extra>'
                ))
                fig_ev.add_hline(y=85, line_dash="dash", line_color="#34d399",
                                 annotation_text="Meta 85%", annotation_font_color="#34d399")
                fig_ev.update_layout(
                                        title="Evolução da Disponibilidade por Data",
                                        height=360,
                                        **plotly_dark()
                                    )
                fig_ev.update_yaxes(title="(%)", range=[0,115])
                st.plotly_chart(fig_ev, use_container_width=True)

                # Barras empilhadas
                fig_st = go.Figure()
                for cat, cor_c in [('Tempo Produtivo','#10b981'),
                                    ('Paradas Planejadas','#f59e0b'),
                                    ('Paradas Não Planejadas','#ef4444')]:
                    if cat in tabela_dia_f.columns:
                        fig_st.add_trace(go.Bar(
                            name=cat, x=tabela_dia_f['Data'].astype(str),
                            y=tabela_dia_f[cat], marker_color=cor_c,
                            text=tabela_dia_f[cat].apply(decimal_para_hhmm),
                            textposition='inside', insidetextanchor='middle',
                            textfont=dict(size=9, color='white')
                        ))
                fig_st.update_layout(barmode='stack', title="Composição por Data",
                                     yaxis_title="Horas", height=320, **plotly_dark())
                st.plotly_chart(fig_st, use_container_width=True)
            else:
                st.markdown('<div class="info-box">ℹ️ Apenas uma data no período — o gráfico de evolução requer múltiplos dias.</div>',
                            unsafe_allow_html=True)

            tbl_d = tabela_dia_f[['Data','DiaSemana_Nome','Num_Pessoas',
                                   'Tempo Produtivo','Paradas Planejadas',
                                   'Paradas Não Planejadas','Disponibilidade (%)']].copy()
            tbl_d = tbl_d.rename(columns={'DiaSemana_Nome':'Dia da Semana','Num_Pessoas':'Pessoas'})
            tbl_d['Tempo Produtivo'] = tbl_d['Tempo Produtivo'].apply(decimal_para_hhmm)
            tbl_d['Paradas Planejadas'] = tbl_d['Paradas Planejadas'].apply(decimal_para_hhmm)
            tbl_d['Paradas Não Planejadas'] = tbl_d['Paradas Não Planejadas'].apply(decimal_para_hhmm)
            tbl_d['Disponibilidade (%)'] = tabela_dia_f['Disponibilidade (%)'].apply(lambda v: f"{v:.1f}%")
            st.dataframe(tbl_d, use_container_width=True, hide_index=True)

        # ── TAB 3 — Paradas ──
        with tab3:
            c1, c2 = st.columns(2)
            with c1:
                fig_p = px.bar(paradas_f.head(15), x='Duração Total (h)', y='Descrição',
                               color='Tipo', orientation='h',
                               color_discrete_map={'Paradas Planejadas':'#f59e0b',
                                                   'Paradas Não Planejadas':'#ef4444'},
                               title="Top 15 Causas de Parada", text='Duração (h:m)')
                fig_p.update_traces(textposition='outside', textfont_color='#e2e8f0')
                fig_p.update_layout(height=460, **plotly_dark())
                st.plotly_chart(fig_p, use_container_width=True)
            with c2:
                fig_p2 = go.Figure(go.Pie(
                    labels=paradas_f['Descrição'], values=paradas_f['Duração Total (h)'],
                    hole=0.4, textinfo='label+percent',
                    textfont=dict(color='#e2e8f0', size=10)
                ))
                fig_p2.update_layout(title="Distribuição de Paradas",
                                     height=460, showlegend=False, **plotly_dark())
                st.plotly_chart(fig_p2, use_container_width=True)

            tbl_p = paradas_f[['Descrição','Tipo','Duração (h:m)']].copy()
            st.dataframe(tbl_p, use_container_width=True, hide_index=True)

        # ── TAB 4 — Ordens de Produção ──
        with tab4:
            c1, c2 = st.columns(2)
            with c1:
                fig_op = px.bar(horas_op_f.head(20),
                                x='Ordem de Produção', y='Horas Trabalhadas',
                                title="Horas por OP (Top 20)",
                                color='Horas Trabalhadas',
                                color_continuous_scale='Blues',
                                text='Horas (h:m)')
                fig_op.update_traces(textposition='outside', textfont_color='#e2e8f0')
                fig_op.update_layout(height=420, coloraxis_showscale=False, **plotly_dark())
                st.plotly_chart(fig_op, use_container_width=True)
            with c2:
                fig_tree = px.treemap(tempo_atividade_f.head(20),
                                      path=['Atividade'], values='Total (h)',
                                      title="Distribuição de Atividades Produtivas",
                                      color='Total (h)', color_continuous_scale='Blues')
                fig_tree.update_layout(height=420, paper_bgcolor='#0f1117',
                                       font=dict(color='#e2e8f0'),
                                       title_font=dict(color='#e2e8f0'))
                st.plotly_chart(fig_tree, use_container_width=True)

            st.markdown('<div class="section-title">Resumo por OP</div>', unsafe_allow_html=True)
            tbl_op = horas_op_f[['Ordem de Produção','Horas (h:m)']].copy()
            tbl_op['Horas (decimal)'] = horas_op_f['Horas Trabalhadas'].apply(lambda x: f"{x:.2f}h")
            st.dataframe(tbl_op, use_container_width=True, hide_index=True)

        # ── TAB 5 — OP + Operações ──
        with tab5:
            st.markdown('<div class="section-title">Detalhamento por OP e Operação</div>',
                        unsafe_allow_html=True)

            if horas_op_oper_f.empty:
                st.markdown('<div class="info-box">Nenhuma OP encontrada no período.</div>',
                            unsafe_allow_html=True)
            else:
                # Gráfico sunburst geral
                fig_sun = px.sunburst(
                    horas_op_oper_f,
                    path=['Referência','Descrição'],
                    values='Horas Trabalhadas',
                    title="OP → Operação (Sunburst)",
                    color='Horas Trabalhadas',
                    color_continuous_scale='Blues'
                )
                fig_sun.update_layout(height=480, paper_bgcolor='#0f1117',
                                      font=dict(color='#e2e8f0'),
                                      title_font=dict(color='#e2e8f0', size=14))
                st.plotly_chart(fig_sun, use_container_width=True)

                # Expanders por OP
                st.markdown('<div class="section-title">Detalhamento por OP</div>',
                            unsafe_allow_html=True)
                for op in horas_op_oper_f['Referência'].unique():
                    df_op = horas_op_oper_f[horas_op_oper_f['Referência'] == op]
                    total_op = df_op['Horas Trabalhadas'].sum()

                    with st.expander(
                        f"📦 {op}  —  Total: {decimal_para_hhmm(total_op)}  ({total_op:.2f}h)",
                        expanded=False
                    ):
                        # Gráfico de barras das operações
                        labels_op = df_op.apply(
                            lambda r: f"Op.{r.get('N° oper.','')} — {str(r.get('Descrição',''))[:45]}",
                            axis=1)
                        fig_oper = go.Figure(go.Bar(
                            x=df_op['Horas Trabalhadas'], y=labels_op,
                            orientation='h', marker_color='#3b82f6',
                            text=df_op['Horas (h:m)'],
                            textposition='outside', textfont=dict(color='#e2e8f0', size=11)
                        ))
                        fig_oper.update_layout(
                                                    height=max(180, 55*len(df_op) + 80),
                                                    xaxis_title="Horas",
                                                    **plotly_dark()
                                                )
                        fig_oper.update_layout(
                                                    margin=dict(t=10, b=30, l=10, r=60)
                                                )
                        st.plotly_chart(fig_oper, use_container_width=True)

                        cols_s = [c for c in ['N° oper.','Ident. do trabalho',
                                               'Descrição','Horas (h:m)']
                                  if c in df_op.columns]
                        st.dataframe(df_op[cols_s], use_container_width=True, hide_index=True)

        # ── TAB 6 — Dados Brutos ──
        with tab6:
            c1, c2 = st.columns(2)
            cats = ['Todos'] + sorted(df_f['Categoria'].unique().tolist())
            cat_sel = c1.selectbox("Filtrar categoria:", cats)
            nomes = ['Todos'] + sorted(df_f['Nome'].unique().tolist())
            nome_sel = c2.selectbox("Filtrar pessoa:", nomes)

            df_show = df_f.copy()
            if cat_sel != 'Todos':
                df_show = df_show[df_show['Categoria'] == cat_sel]
            if nome_sel != 'Todos':
                df_show = df_show[df_show['Nome'] == nome_sel]

            cols_ex = ['Nome','Data','DiaSemana_Nome','Hora inicial','Hora final',
                       'Tipo de registro do diário','Referência','N° oper.',
                       'Descrição','Duracao_Horas','Categoria']
            cols_ex = [c for c in cols_ex if c in df_show.columns]
            tbl_b = df_show[cols_ex].copy()
            tbl_b['Duracao_Horas'] = tbl_b['Duracao_Horas'].apply(
                lambda x: f"{decimal_para_hhmm(x)} ({x:.2f}h)")
            tbl_b = tbl_b.rename(columns={
                'DiaSemana_Nome':'Dia', 'Tipo de registro do diário':'Tipo',
                'Duracao_Horas':'Duração'
            })
            st.dataframe(tbl_b, use_container_width=True, hide_index=True)
            st.caption(f"📋 {len(df_show)} registros  |  {data_inicio} → {data_fim}")

        # ── Export ──
        st.markdown('<hr>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Exportar Relatório</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([1, 2])
        with c1:
            buf = gerar_excel(tabela_pessoa_f, tabela_dia_f, paradas_f,
                              horas_op_f, horas_op_oper_f, tempo_atividade_f,
                              horas_padrao, horas_sexta)
            st.download_button(
                "⬇️ Baixar Relatório Excel (.xlsx)", data=buf,
                file_name=f"disponibilidade_{data_inicio}_{data_fim}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c2:
            st.markdown(f"""
<div class="info-box">
📅 Período: <b>{data_inicio}</b> → <b>{data_fim}</b>  &nbsp;|&nbsp;
👷 <b>{n_pess}</b> pessoas  &nbsp;|&nbsp;
📊 Disponibilidade: <b>{disp:.1f}%</b><br>
⏱️ Jornada padrão: <b>{horas_padrao}h</b>  &nbsp;|&nbsp;
🗓️ Sexta-feira: <b>{horas_sexta}h</b>  &nbsp;|&nbsp;
📦 OPs analisadas: <b>{len(horas_op_f)}</b>
</div>""", unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
        st.exception(e)

else:
    st.markdown("""
<div style="text-align:center;padding:60px 20px 40px">
  <div style="font-size:4rem;margin-bottom:16px">🏭</div>
  <div style="font-size:1.4rem;font-weight:600;color:#e2e8f0;margin-bottom:8px">
    Carregue sua planilha para começar
  </div>
  <div style="font-size:0.9rem;color:#475569;max-width:500px;margin:0 auto">
    Importe o arquivo <code>.xlsx</code> exportado do Dynamics 365 F&amp;O
    e explore os indicadores de disponibilidade da sua equipe.
  </div>
</div>
""", unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    for col, icon, titulo, desc in [
        (c1, "📊", "Disponibilidade", "Por pessoa e por data, com meta de 85%"),
        (c2, "🔩", "OPs + Operações", "Sunburst OP→Operação com detalhamento por expander"),
        (c3, "📅", "Análise Temporal", "Evolução diária com identificação do dia da semana"),
    ]:
        col.markdown(f"""
<div style="background:#161b27;border:1px solid #1e2d45;border-radius:12px;
            padding:22px;text-align:center;height:130px">
  <div style="font-size:1.8rem;margin-bottom:10px">{icon}</div>
  <div style="font-weight:600;color:#e2e8f0;margin-bottom:6px">{titulo}</div>
  <div style="font-size:0.82rem;color:#475569">{desc}</div>
</div>""", unsafe_allow_html=True)

    st.markdown("""
<div style="margin-top:32px;background:#161b27;border:1px solid #1e2d45;
            border-radius:12px;padding:24px;max-width:680px;margin:32px auto 0">
  <div style="font-size:0.72rem;text-transform:uppercase;letter-spacing:2px;
              color:#475569;margin-bottom:12px;font-weight:600">Como usar</div>
  <ol style="color:#94a3b8;font-size:0.88rem;line-height:2;padding-left:18px">
    <li>Configure a <b style="color:#60a5fa">jornada padrão</b> (Seg–Qui) e a jornada de <b style="color:#60a5fa">Sexta</b> no menu lateral</li>
    <li>Carregue o arquivo <code>.xlsx</code> do Dynamics 365 F&amp;O</li>
    <li>Selecione o <b style="color:#60a5fa">período</b> desejado nos filtros de data</li>
    <li>Explore as <b style="color:#60a5fa">6 abas</b> de análise</li>
    <li>Exporte o relatório em <b style="color:#60a5fa">Excel</b> com 6 planilhas</li>
  </ol>
</div>
""", unsafe_allow_html=True)