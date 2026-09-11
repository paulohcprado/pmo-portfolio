import streamlit as st
import pandas as pd
import os
import altair as alt
from datetime import datetime

st.set_page_config(page_title="PMO - Portfólio", layout="wide")

st.markdown("""
    <style>
    span[data-baseweb="tag"] { background-color: #5B92E5 !important; }
    div[data-baseweb="select"] > div { border-color: #5B92E5 !important; }
    .stDataFrame { width: 100% !important; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "base_projetos.xlsx"
VETORES = ['ROI', 'Blindagem Cliente', 'Regulatório', 'Expansão', 'Backoffice', 'Melhoria Contínua', 'Inovação Digital', 'Mitigação de Risco']
STATUS_LIST = ["Backlog", "Planejado", "Em Execução", "Concluído"]

cfg_colunas = {
    "ID": st.column_config.TextColumn("ID", width="small"),
    "Projeto": st.column_config.TextColumn("Projeto", width="large"),
    "Status & Prazo": st.column_config.TextColumn("Status & Prazo", width="medium"),
    "Início": st.column_config.TextColumn("Início", width="small"),
    "Término Previsto": st.column_config.TextColumn("Previsão", width="small"),
    "Término Real": st.column_config.TextColumn("Real", width="small"),
    "Score Final": st.column_config.NumberColumn("Score Final", format="%d", width="small")
}

def formatar_moeda(valor):
    if pd.isna(valor) or valor == 0: return "-"
    return f"R$ {valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

@st.cache_data
def carregar_dados():
    if not os.path.exists(EXCEL_FILE): return pd.DataFrame()
    df = pd.read_excel(EXCEL_FILE)
    if 'Término Previsto' not in df.columns:
        st.error("⚠️ Base desatualizada. Rode o 'gerar_planilha.bat' primeiro.")
        st.stop()
        
    if 'Equipe' in df.columns:
        df['Equipe'] = df['Equipe'].apply(lambda x: sorted([p.strip() for p in str(x).split(',')]) if pd.notnull(x) and x != "" else [])
    
    def definir_status_visual(row):
        if row['Status'] == 'Concluído':
            if pd.notnull(row['Término Real']) and str(row['Término Real']) > str(row['Término Previsto']):
                return "🚨 Concluído (Atrasado)"
            return "✅ Concluído (No Prazo)"
        elif row['Status'] == 'Em Execução': return "⏳ Em Execução"
        elif row['Status'] == 'Planejado': return "📅 Planejado"
        else: return "📦 Backlog"
            
    df['Status & Prazo'] = df.apply(definir_status_visual, axis=1)
    return df

df_projetos = carregar_dados()
if df_projetos.empty: st.stop()

with st.sidebar:
    st.header("⚙️ Pesos dos Vetores")
    pesos = {}
    for vetor in VETORES: pesos[vetor] = st.number_input(f"Peso: {vetor}", value=1, step=1, key=f"peso_{vetor}")
    st.divider()
    st.caption("Versão 3.4.0 | Somatório no Fluxo de Caixa")

def calcular_score_dinamico(df, pesos_dict):
    score_base = pd.Series(0.0, index=df.index)
    for v in VETORES:
        if f"Nota {v}" in df.columns: score_base += df[f"Nota {v}"] * pesos_dict[v]
    return (score_base * (1 + (df['ROI (%)'].clip(lower=-50) / 100))).round(0).astype(int)

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Dashboards", "🏆 Ranqueamento", "📋 Matriz de Notas", "👥 Alocação e Fornecedores", "⚙️ Base Consolidada"])

with tab1:
    st.markdown("**Filtros Globais**")
    cf1, cf2, cf3, cf4 = st.columns(4)
    with cf1: f_anos = st.multiselect("Ano de Início", sorted(list(set([str(x)[:4] for x in df_projetos['Início']]))), key="d_a")
    with cf2: f_status = st.multiselect("Status", STATUS_LIST, key="d_s")
    with cf3: f_tipo = st.multiselect("Tipo de Projeto", sorted(list(set(df_projetos['Tipo']))), key="d_t")
    with cf4: f_recs = st.multiselect("Recurso Interno", sorted(list(set([p for lista in df_projetos['Equipe'] for p in lista]))), key="d_r")

    df_dash = df_projetos.copy()
    if f_anos: df_dash = df_dash[df_dash['Início'].str[:4].isin(f_anos)]
    if f_status: df_dash = df_dash[df_dash['Status'].isin(f_status)]
    if f_tipo: df_dash = df_dash[df_dash['Tipo'].isin(f_tipo)]
    if f_recs: df_dash = df_dash[df_dash['Equipe'].apply(lambda e: any(r in e for r in f_recs))]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Projetos Filtrados", len(df_dash))
    m2.metric("Orçamento Total", formatar_moeda(df_dash['Orçamento (R$)'].sum()))
    m3.metric("Retorno Esperado", formatar_moeda(df_dash['Retorno (R$)'].sum()))
    m4.metric("ROI Médio", f"{df_dash['ROI (%)'].mean():.1f}%" if not df_dash.empty else "0%")
    
    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Volume por Status**")
        if not df_dash.empty:
            st.altair_chart(alt.Chart(df_dash['Status'].value_counts().reset_index()).mark_bar(color='#5B92E5', cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
                x=alt.X('Status:N', sort=None, title=""), y=alt.Y('count:Q', title="")
            ).properties(height=400), use_container_width=True)
            
    with g2:
        st.markdown("**Orçamento por Tipo (Milhões R$)**")
        if not df_dash.empty:
            df_tipo_orc = df_dash.groupby('Tipo')['Orçamento (R$)'].sum().reset_index()
            df_tipo_orc['Orçamento (R$)'] = df_tipo_orc['Orçamento (R$)'] / 1000000
            st.altair_chart(alt.Chart(df_tipo_orc).mark_bar(color='#6CC4A1', cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
                y=alt.Y('Tipo:N', sort='-x', title="", axis=alt.Axis(labelLimit=0)),
                x=alt.X('Orçamento (R$):Q', title=""),
                tooltip=[alt.Tooltip('Tipo:N'), alt.Tooltip('Orçamento (R$):Q', format='.2f', title='Orçamento (M)')]
            ).properties(height=400), use_container_width=True)

with tab2:
    df_rank = df_projetos.copy()
    df_rank['Score Final'] = calcular_score_dinamico(df_rank, pesos)
    
    padrao_rank = ["Em Execução", "Planejado", "Backlog"]
    
    cr1, cr2, cr3 = st.columns(3)
    with cr1: f_tipo_rk = st.multiselect("Filtrar Tipo", sorted(list(set(df_rank['Tipo']))), key="rk_t")
    with cr2: f_status_rk = st.multiselect("Filtrar Status", STATUS_LIST, default=padrao_rank, key="rk_s")
    with cr3:
        m_s, M_s = float(df_rank['Score Final'].min()), float(df_rank['Score Final'].max())
        f_score = st.slider("Faixa de Score", m_s, M_s, (m_s, M_s)) if m_s < M_s else (m_s, M_s)

    if f_tipo_rk: df_rank = df_rank[df_rank['Tipo'].isin(f_tipo_rk)]
    if f_status_rk: df_rank = df_rank[df_rank['Status'].isin(f_status_rk)]
    df_rank = df_rank[(df_rank['Score Final'] >= f_score[0]) & (df_rank['Score Final'] <= f_score[1])]
    
    df_rank['Equipe'] = df_rank['Equipe'].apply(lambda x: ", ".join(x))
    st.dataframe(df_rank.sort_values(by='Score Final', ascending=False)[['ID', 'Projeto', 'Tipo', 'Score Final', 'Status & Prazo', 'Equipe']], use_container_width=True, hide_index=True, height=600, column_config=cfg_colunas)

with tab3:
    st.markdown("**Matriz de Avaliação**")
    df_notas = df_projetos[['ID', 'Projeto'] + [f"Nota {v}" for v in VETORES]].copy()
    df_notas.columns = ['ID', 'Projeto', 'N. ROI', 'N. Blind.', 'N. Reg.', 'N. Exp.', 'N. Back.', 'N. Melh.', 'N. Inov.', 'N. Risco']
    
    cfg_notas = {"ID": st.column_config.TextColumn("ID", width="small"), "Projeto": st.column_config.TextColumn("Projeto", width="large")}
    for col in df_notas.columns[2:]: cfg_notas[col] = st.column_config.NumberColumn(col, width="small")
    st.dataframe(df_notas, use_container_width=True, hide_index=True, height=600, column_config=cfg_notas)

with tab4:
    t4_1, t4_2 = st.tabs(["👥 Colaboradores", "🏢 Fornecedores (Fluxo de Caixa)"])
    
    with t4_1:
        st.caption("Legenda: 🟢 Em Execução | 🟡 Planejado | 🔵 Concluído | ⚪ Backlog")
        aloc = []
        for _, r in df_projetos.iterrows():
            try:
                dt_fim = r['Término Real'] if pd.notnull(r['Término Real']) else r['Término Previsto']
                for m in pd.period_range(pd.to_datetime(r['Início']).to_period('M'), pd.to_datetime(dt_fim).to_period('M')):
                    for p in r['Equipe']:
                        ic = '🟢' if r['Status'] == 'Em Execução' else ('🟡' if r['Status'] == 'Planejado' else ('🔵' if r['Status'] == 'Concluído' else '⚪'))
                        aloc.append({'Colaborador': p, 'Data_Mês': m.to_timestamp(), 'Status_Raw': r['Status'], 'Ícone': ic})
            except: continue
            
        if aloc:
            df_aloc = pd.DataFrame(aloc)
            df_aloc['Periodo'] = df_aloc['Data_Mês'].dt.strftime('%Y-%m')
            df_aloc['Ano'] = df_aloc['Data_Mês'].dt.year.astype(str)
            meses_pt = {1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'}
            df_aloc['Mês'] = df_aloc['Data_Mês'].dt.month.map(meses_pt)
            df_aloc['Mês'] = pd.Categorical(df_aloc['Mês'], categories=list(meses_pt.values()), ordered=True)
            
            periodos = sorted(list(set(df_aloc['Periodo'])))
            hoje_str = datetime.now().strftime('%Y-%m')
            idx_ini = periodos.index(hoje_str) if hoje_str in periodos else 0
            
            c_a1, c_a2, c_a3, c_a4 = st.columns(4)
            with c_a1: f_ini = st.selectbox("Mês Inicial", periodos, index=idx_ini, key="a_ini")
            with c_a2: f_fim = st.selectbox("Mês Final", periodos, index=len(periodos)-1, key="a_fim")
            with c_a3: f_colab = st.multiselect("Colaborador", sorted(list(set(df_aloc['Colaborador']))), key="a_c")
            with c_a4: f_stat_col = st.multiselect("Status", STATUS_LIST, key="a_s")
            
            df_aloc = df_aloc[(df_aloc['Periodo'] >= f_ini) & (df_aloc['Periodo'] <= f_fim)]
            if f_colab: df_aloc = df_aloc[df_aloc['Colaborador'].isin(f_colab)]
            if f_stat_col: df_aloc = df_aloc[df_aloc['Status_Raw'].isin(f_stat_col)]
            
            if not df_aloc.empty:
                tb_aloc = pd.pivot_table(df_aloc, index='Colaborador', columns=['Ano', 'Mês'], values='Ícone', aggfunc=lambda x: "".join(set(x)), fill_value="")
                tb_aloc = tb_aloc.dropna(axis=1, how='all')
                st.dataframe(tb_aloc, use_container_width=True, height=500)
            else: st.info("Nenhuma alocação encontrada para este período.")
            
    with t4_2:
        st.caption("ℹ️ Orçamento rateado ao longo dos meses de vigência do projeto. Valores em milhares (k).")
        f_det = []
        for _, r in df_projetos[df_projetos['Fornecedor'] != 'N/A'].iterrows():
            try:
                dt_fim = r['Término Real'] if pd.notnull(r['Término Real']) else r['Término Previsto']
                meses_proj = pd.period_range(pd.to_datetime(r['Início']).to_period('M'), pd.to_datetime(dt_fim).to_period('M'))
                if len(meses_proj) > 0:
                    valor_mensal = r['Orçamento (R$)'] / len(meses_proj)
                    for m in meses_proj: f_det.append({'Fornecedor': r['Fornecedor'], 'Data_Mês': m.to_timestamp(), 'Valor': valor_mensal})
            except: continue
            
        if f_det:
            df_f = pd.DataFrame(f_det)
            df_f['Periodo'] = df_f['Data_Mês'].dt.strftime('%Y-%m')
            df_f['Ano'] = df_f['Data_Mês'].dt.year.astype(str)
            df_f['Mês'] = df_f['Data_Mês'].dt.month.map(meses_pt)
            df_f['Mês'] = pd.Categorical(df_f['Mês'], categories=list(meses_pt.values()), ordered=True)
            
            periodos_f = sorted(list(set(df_f['Periodo'])))
            idx_ini_f = periodos_f.index(hoje_str) if hoje_str in periodos_f else 0
            
            c_f1, c_f2 = st.columns(2)
            with c_f1: f_ini_f = st.selectbox("Mês Inicial (Fornecedor)", periodos_f, index=idx_ini_f, key="f_ini")
            with c_f2: f_fim_f = st.selectbox("Mês Final (Fornecedor)", periodos_f, index=len(periodos_f)-1, key="f_fim")
            
            df_f = df_f[(df_f['Periodo'] >= f_ini_f) & (df_f['Periodo'] <= f_fim_f)]
            
            if not df_f.empty:
                tb_f = pd.pivot_table(df_f, index='Fornecedor', columns=['Ano', 'Mês'], values='Valor', aggfunc='sum', fill_value=0)
                tb_f = tb_f.dropna(axis=1, how='all')
                
                # Adiciona o somatório mensal no rodapé
                tb_f.loc['Total Mensal'] = tb_f.sum(axis=0)
                
                for col in tb_f.columns: 
                    tb_f[col] = tb_f[col].apply(lambda x: f"{x/1000:,.0f}k".replace(',', '.') if x > 0 else "-")
                st.dataframe(tb_f, use_container_width=True, height=500)
            else: st.info("Nenhum fornecedor no período selecionado.")

with tab5:
    df_view = df_projetos.copy()
    df_view['Equipe'] = df_view['Equipe'].apply(lambda x: ", ".join(x))
    df_view['Orçamento (R$)'] = df_view['Orçamento (R$)'].apply(formatar_moeda)
    df_view['Retorno (R$)'] = df_view['Retorno (R$)'].apply(formatar_moeda)
    st.dataframe(df_view[['ID', 'Projeto', 'Tipo', 'Status & Prazo', 'Início', 'Término Previsto', 'Término Real', 'Orçamento (R$)', 'Retorno (R$)', 'Equipe']], use_container_width=True, hide_index=True, height=700, column_config=cfg_colunas)