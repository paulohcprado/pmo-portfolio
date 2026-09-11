import streamlit as st
import pandas as pd
import numpy as np
import os
import altair as alt
from datetime import datetime, timedelta
import random
import json

try:
    import google.generativeai as genai
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

st.set_page_config(page_title="PMO - Portfólio", layout="wide")

st.markdown("""
    <style>
    span[data-baseweb="tag"] { background-color: #5B92E5 !important; }
    div[data-baseweb="select"] > div { border-color: #5B92E5 !important; }
    .stDataFrame { width: 100% !important; }

    /* Estilização Moderna para Cards e Containers */
    .welcome-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 30px;
        border-radius: 16px;
        color: white;
        box-shadow: 0 10px 25px rgba(0,0,0,0.2);
        margin-bottom: 25px;
    }
    .welcome-card h1 {
        color: #60a5fa;
        font-size: 28px;
        margin-bottom: 10px;
    }
    .welcome-card p {
        color: #cbd5e1;
        font-size: 16px;
        line-height: 1.6;
    }

    /* Card Compacto para a IA (UI/UX Profissional) */
    .ai-container {
        background-color: #111827;
        padding: 24px;
        border-radius: 14px;
        border: 1px solid #374151;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        margin-bottom: 20px;
    }

    /* Botão Flutuante de Contato Corporativo */
    .floating-contact {
        position: fixed;
        bottom: 25px;
        right: 25px;
        z-index: 999999;
        display: flex;
        flex-direction: column;
        align-items: flex-end;
        font-family: sans-serif;
    }
    .contact-btn {
        background-color: #25D366;
        color: white;
        border: none;
        border-radius: 50px;
        width: 60px;
        height: 60px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 28px;
        text-decoration: none;
        transition: transform 0.3s ease;
    }
    .contact-btn:hover {
        transform: scale(1.1);
    }
    .contact-menu {
        display: none;
        flex-direction: column;
        gap: 6px;
        margin-bottom: 12px;
        background: white;
        padding: 12px;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.25);
        min-width: 190px;
    }
    .floating-contact:hover .contact-menu {
        display: flex;
    }
    .contact-header {
        font-weight: 700;
        color: #1f2937;
        font-size: 13px;
        padding: 4px 6px;
        border-bottom: 1px solid #eee;
        margin-bottom: 2px;
    }
    .contact-item {
        display: flex;
        align-items: center;
        gap: 8px;
        text-decoration: none;
        font-weight: 600;
        font-size: 13px;
        padding: 8px 10px;
        border-radius: 8px;
        transition: opacity 0.2s;
        white-space: nowrap;
    }
    .contact-item.wa { background: #e2fbe8; color: #075e54; }
    .contact-item.email { background: #e8f0fe; color: #1a73e8; }
    .contact-item.linkedin { background: #e8f4f8; color: #0077b5; }
    .contact-item.web { background: #f3f4f6; color: #374151; }
    .contact-item:hover { opacity: 0.85; }
    </style>

    <!-- Widget Flutuante de Contato no Canto Inferior Direito -->
    <div class="floating-contact">
        <div class="contact-menu">
            <div class="contact-header">Fale com o autor</div>
            <a class="contact-item wa" href="https://wa.me/5511974397340?text=Olá%20Paulo,%20vi%20seu%20portfólio%20PMO%20e%20gostaria%20de%20conversar." target="_blank">
                💬 WhatsApp
            </a>
            <a class="contact-item email" href="mailto:paulohcprado@gmail.com?subject=Contato%20via%20Portfólio%20PMO" target="_blank">
                ✉️ E-mail
            </a>
            <a class="contact-item linkedin" href="https://www.linkedin.com/in/paulohcp/" target="_blank">
                💼 LinkedIn
            </a>
            <a class="contact-item web" href="https://www.pcestari.tech" target="_blank">
                🌐 Site (pcestari.tech)
            </a>
        </div>
        <a class="contact-btn" href="https://wa.me/5511974397340?text=Olá%20Paulo,%20vi%20seu%20portfólio%20PMO%20e%20gostaria%20de%20conversar." target="_blank" title="Fale com o autor">
            💬
        </a>
    </div>
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

def processar_dataframe(df):
    if df.empty: return df
    if 'Equipe' in df.columns:
        df['Equipe'] = df['Equipe'].apply(lambda x: sorted([p.strip() for p in str(x).split(',')]) if pd.notnull(x) and x != "" else [])
    
    for v in VETORES:
        col_name = f"Nota {v}"
        if col_name not in df.columns:
            df[col_name] = 5
            
    if 'Orçamento (R$)' not in df.columns: df['Orçamento (R$)'] = 100000.0
    if 'Retorno (R$)' not in df.columns: df['Retorno (R$)'] = 120000.0
    if 'ROI (%)' not in df.columns: df['ROI (%)'] = 20.0
    if 'Status' not in df.columns: df['Status'] = 'Em Execução'
    if 'Início' not in df.columns: df['Início'] = '2026-01-01'
    if 'Término Previsto' not in df.columns: df['Término Previsto'] = '2026-12-31'
    if 'Tipo' not in df.columns: df['Tipo'] = 'Melhoria SAP'
    if 'Fornecedor' not in df.columns: df['Fornecedor'] = 'N/A'
    
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

if 'df_projetos' not in st.session_state:
    if os.path.exists(EXCEL_FILE):
        df_base = pd.read_excel(EXCEL_FILE)
        st.session_state.df_projetos = processar_dataframe(df_base)
    else:
        st.session_state.df_projetos = pd.DataFrame()

if 'msg_sucesso' not in st.session_state:
    st.session_state.msg_sucesso = None

if 'prompt_usuario_key' not in st.session_state:
    st.session_state.prompt_usuario_key = "Gere 100 projetos inspirados em uvas e regiões vinícolas mundiais entre os anos de 2020 e 2030"

MASTER_SUGESTOES = [
    "Gere 100 projetos inspirados em uvas e regiões vinícolas mundiais (como Bordeaux, Malbec, Chianti) entre 2020 e 2030",
    "Gere 120 projetos de transformação digital baseados em empresas da bolsa de valores do Brasil (B3)",
    "Gere 80 projetos de inovação corporativa inspirados em clássicos do cinema de Hollywood",
    "Gere 100 projetos de tecnologia usando os nomes das maiores cidades da América Latina",
    "Gere 90 projetos de melhoria contínua baseados em raças e nomes populares de pets ao redor do mundo",
    "Gere 150 projetos de automação inspirados em pratos e gastronomia típica mundial",
    "Gere 80 projetos corporativos baseados em praias e arquipélagos turísticos globais",
    "Gere 100 projetos de infraestrutura de TI baseados em constelações e astronomia",
    "Gere 90 projetos de governança inspirados nas Maravilhas do Mundo Antigo e Moderno",
    "Gere 120 projetos de transformação digital baseados em missões espaciais da NASA e SpaceX",
    "Gere 100 projetos de logística corporativa inspirados em regiões vinícolas e tipos de uvas",
    "Gere 150 projetos de cibersegurança inspirados em mitologia grega e nórdica",
    "Gere 90 projetos de inteligência artificial inspirados em super-heróis dos quadrinhos",
    "Gere 100 projetos de sustentabilidade ESG inspirados em parques nacionais do Brasil",
    "Gere 80 projetos de varejo omnichannel inspirados em grandes inventores da história",
    "Gere 120 projetos de arquitetura em nuvem baseados em instrumentos musicais clássicos",
    "Gere 90 projetos de portfólio financeiro inspirados em pedras preciosas e minerais raros",
    "Gere 100 projetos de mobilidade urbana inspirados em escuderias e pilotos de Fórmula 1",
    "Gere 80 projetos de transformação ágil inspirados em grandes rios do mundo",
    "Gere 150 projetos de e-commerce baseados em flora e fauna da Floresta Amazônica"
]

if 'sugestoes_clicks' not in st.session_state:
    st.session_state.sugestoes_clicks = 0

if 'current_sugestoes' not in st.session_state:
    random.seed(42)
    st.session_state.current_sugestoes = random.sample(MASTER_SUGESTOES, min(6, len(MASTER_SUGESTOES)))

df_projetos = st.session_state.df_projetos

with st.sidebar:
    st.header("⚙️ Pesos dos Vetores")
    pesos = {}
    for vetor in VETORES: pesos[vetor] = st.number_input(f"Peso: {vetor}", value=1, step=1, key=f"peso_{vetor}")
    st.divider()
    st.caption("Versão 8.1.0 | Fixed Syntax & UI/UX Polish")

def calcular_score_dinamico(df, pesos_dict):
    score_base = pd.Series(0.0, index=df.index)
    for v in VETORES:
        col_name = f"Nota {v}"
        if col_name in df.columns: score_base += df[col_name] * pesos_dict[v]
    return (score_base * (1 + (df['ROI (%)'].clip(lower=-50) / 100))).round(0).astype(int)

# Definição das Abas com a Página Inicial Dedicada
tab_home, tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "🏠 Início", "📈 Dashboards", "🏆 Ranqueamento", "📋 Matriz de Notas", 
    "👥 Alocação e Fornecedores", "⚙️ Base Consolidada", "ℹ️ Sobre & Tech", 
    "👤 Sobre o autor", "🤖 Gerador IA"
])

with tab_home:
    st.markdown("""
        <div class="welcome-card">
            <h1>🚀 Bem-vindo ao Portfólio PMO & Governança</h1>
            <p>Este sistema foi desenvolvido para demonstração avançada de gestão de portfólios corporativos, alocação de recursos, fluxo de caixa de fornecedores e motores analíticos integrados com Inteligência Artificial.</p>
        </div>
    """, unsafe_allow_html=True)
    
    st.info("📱 **Nota de Navegação Mobile:** Este portfólio possui tabelas analíticas e gráficos corporativos densos. Para uma melhor experiência interativa, recomendamos o acesso através de um **computador (versão web desktop)**.")
    
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        st.markdown("### 📊 Principais Módulos")
        st.markdown("""
        - **Dashboards:** Visão executiva em tempo real com filtros globais de status, anos e recursos.
        - **Ranqueamento:** Matriz ponderada de priorização baseada em vetores estratégicos.
        - **Alocação e Fornecedores:** Controle de capacidade de colaboradores e fluxo de caixa de fornecedores.
        """)
    with col_h2:
        st.markdown("### 🤖 Inovação com IA")
        st.markdown("""
        - **Gerador Inteligente:** Crie bases de dados dinâmicas sob medida utilizando linguagem natural.
        - **Isolamento de Sessão:** Cada visitante explora dados independentes que retornam ao estado padrão ao reiniciar.
        """)

with tab1:
    if df_projetos.empty:
        st.warning("⚠️ Nenhuma base de dados encontrada. Vá até a aba '🤖 Gerador IA' para gerar seus dados.")
    else:
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
    if not df_projetos.empty:
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
    if not df_projetos.empty:
        st.markdown("**Matriz de Avaliação**")
        df_notas = df_projetos[['ID', 'Projeto'] + [f"Nota {v}" for v in VETORES]].copy()
        df_notas.columns = ['ID', 'Projeto', 'N. ROI', 'N. Blind.', 'N. Reg.', 'N. Exp.', 'N. Back.', 'N. Melh.', 'N. Inov.', 'N. Risco']
        cfg_notas = {"ID": st.column_config.TextColumn("ID", width="small"), "Projeto": st.column_config.TextColumn("Projeto", width="large")}
        for col in df_notas.columns[2:]: cfg_notas[col] = st.column_config.NumberColumn(col, width="small")
        st.dataframe(df_notas, use_container_width=True, hide_index=True, height=600, column_config=cfg_notas)

with tab4:
    if not df_projetos.empty:
        t4_1, t4_2 = st.tabs(["👥 Colaboradores", "🏢 Fornecedores (Fluxo de Caixa)"])
        HOJE = datetime.now()
        
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
                hoje_str = HOJE.strftime('%Y-%m')
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
                    tb_f.loc['Total Mensal'] = tb_f.sum(axis=0)
                    
                    for col in tb_f.columns: 
                        tb_f[col] = tb_f[col].apply(lambda x: f"{x/1000:,.0f}k".replace(',', '.') if x > 0 else "-")
                    st.dataframe(tb_f, use_container_width=True, height=500)
                else: st.info("Nenhum fornecedor no período selecionado.")

with tab5:
    if not df_projetos.empty:
        df_view = df_projetos.copy()
        df_view['Equipe'] = df_view['Equipe'].apply(lambda x: ", ".join(x))
        df_view['Orçamento (R$)'] = df_view['Orçamento (R$)'].apply(formatar_moeda)
        df_view['Retorno (R$)'] = df_view['Retorno (R$)'].apply(formatar_moeda)
        st.dataframe(df_view[['ID', 'Projeto', 'Tipo', 'Status & Prazo', 'Início', 'Término Previsto', 'Término Real', 'Orçamento (R$)', 'Retorno (R$)', 'Equipe']], use_container_width=True, hide_index=True, height=700, column_config=cfg_colunas)

with tab6:
    st.markdown("### ℹ️ Sobre o Portfólio & Arquitetura Tecnológica")
    st.info("⚠️ **Aviso Legal & Educacional**: Todas as informações, projetos, orçamentos, fornecedores, colaboradores e métricas apresentados nesta aplicação são **estritamente fictícios**, criados exclusivamente para fins educacionais, demonstração de governança PMO e apresentação de portfólio profissional.")
    st.markdown("---")
    col_inf1, col_inf2 = st.columns(2)
    with col_inf1:
        st.markdown("#### 💡 Concepção e Co-criação")
        st.write("""
        Este sistema foi 100% idealizado a partir de premissas reais de gestão de portfólio corporativo e PMO, com código projetado, iterado e escrito em conjunto com o assistente **Gemini**.
        """)
    with col_inf2:
        st.markdown("#### 🛠️ Stack Tecnológica")
        st.markdown("""
        - **Core:** Python 3.x, Streamlit, Pandas, NumPy  
        - **Editor de Scripts & Textos:** Notepad++  
        - **Persistência & Isolamento:** Microsoft Excel (`.xlsx`) com gerenciamento de estado via Sessão (`st.session_state`)  
        - **Visualização:** Altair  
        - **IA Integrada:** Google Gemini API (Protegida via `st.secrets`)  
        """)

with tab7:
    st.markdown("### 👤 Sobre o Autor")
    st.write("Esta aplicação, o modelo de governança e todo o desenvolvimento técnico são de autoria de **Paulo Henrique Cestari Prado**[cite: 1].")
    st.markdown("---")
    st.markdown("#### Resumo Executivo")
    st.write("""
    Gerente de Projetos Sênior com mais de 15 anos liderando projetos tradicionais e 8+ anos em projetos ágeis / híbridos (PMP e PSM2), atuando na construção e expansão de escritórios de projetos (PMO / VMO) e na gestão de portfólios de até R$ 40 milhões/ano em setores como óleo e gás, consultoria, indústria, varejo, tecnologia, bancário e setor público[cite: 1].
    
    Especialista em conduzir transformação digital - da migração de sistemas legados para SAP/Salesforce/cloud/erp até a evolução da maturidade ágil de times multidisciplinares; conectando estratégia, negócio e tecnologia com foco em resultado mensurável, governança e liderança de squads técnicos e de agilidade[cite: 1].
    """)
    
    col_aut1, col_aut2 = st.columns(2)
    with col_aut1:
        st.markdown("##### 🎯 Foco e Governança")
        st.markdown("""
        - **Gestão Corporativa & Resultados:** Foco em entregas mensuráveis, otimização de OPEX/CAPEX e mitigação crítica de riscos (como prevenção de ransomware e desvios operacionais)[cite: 1].
        - **Comunicação Executiva:** Condução de comitês gerenciais, status reports e alinhamento contínuo entre áreas de negócio e a alta diretoria[cite: 1].
        """)
    with col_aut2:
        st.markdown("##### 🚀 Evolução & Trajetória")
        st.markdown("""
        - **Da Base ao C-Level:** Sólida experiência evolutiva desde a fundação operacional e técnica (suporte, redes, infraestrutura e implantações de ERP/WMS) até a gestão sênior de programas.
        - **Certificações & Reconhecimento:** PMP (PMI), PSM2 (Scrum.org), Agile Coach, além de histórico de liderança voluntária no **PMI São Paulo**[cite: 1].
        """)

with tab8:
    st.markdown("### 🤖 Gerador de Dados Inteligente com Gemini AI")
    st.markdown("""
        <div style="background-color: #1e293b; padding: 14px 18px; border-radius: 10px; border-left: 4px solid #3b82f6; margin-bottom: 20px;">
            <p style="color: #e2e8f0; font-size: 14px; margin: 0;">
                💡 <b>Dica de Performance:</b> Recomendamos gerar entre <b>50 e 150 projetos</b> para manter a agilidade da aplicação. As alterações criadas aqui afetam apenas a sua sessão atual e retornam ao padrão original ao reiniciar a página.
            </p>
        </div>
    """, unsafe_allow_html=True)

    # UI/UX Aprimorada: Seção de Sugestões em formato de grade compacta
    st.markdown("#### 🎯 Sugestões de Temas Executivos")
    st.caption("Clique em qualquer tema abaixo para carregá-lo instantaneamente na instrução:")
    
    c_sug_btn, c_sug_info = st.columns([3, 1])
    with c_sug_info:
        cliques_restantes = 5 - st.session_state.sugestoes_clicks
        btn_label = f"🔄 Mais ideias ({cliques_restantes} restantes)" if cliques_restantes > 0 else "🛑 Limite atingido"
        if st.button(btn_label, disabled=(st.session_state.sugestoes_clicks >= 5), use_container_width=True):
            if st.session_state.sugestoes_clicks < 5:
                st.session_state.sugestoes_clicks += 1
                st.session_state.current_sugestoes = random.sample(MASTER_SUGESTOES, min(6, len(MASTER_SUGESTOES)))
                st.rerun()

    cols_sug = st.columns(2)
    for idx, sug_texto in enumerate(st.session_state.current_sugestoes):
        col_alvo = cols_sug[idx % 2]
        with col_alvo:
            if st.button(f"📌 {sug_texto}", key=f"sug_btn_{idx}", use_container_width=True):
                st.session_state.prompt_usuario_key = sug_texto
                st.rerun()

    st.markdown("---")

    # Caixa de Texto e Botão de Geração organizados com UI/UX profissional
    st.markdown("#### ✍️ Instrução de Geração")
    prompt_usuario = st.text_area(
        "Descreva abaixo o tema e a quantidade desejada:",
        key="prompt_usuario_key",
        height=90
    )

    col_btn_gerar, col_vazia = st.columns([2, 3])
    with col_btn_gerar:
        botao_gerar = st.button("🚀 Gerar Base com IA (Sessão Atual)", type="primary", use_container_width=True)

    # Mensagem de sucesso posicionada LOGO ABAIXO do botão onde a pessoa clicou
    if st.session_state.msg_sucesso:
        st.success(st.session_state.msg_sucesso)
        st.session_state.msg_sucesso = None

    if botao_gerar:
        api_key_segura = st.secrets.get("GEMINI_API_KEY") if hasattr(st, "secrets") else None
        
        if not api_key_segura:
            st.error("⚠️ Chave da API do Gemini não encontrada nos segredos seguros do Streamlit (`st.secrets`).")
        elif not HAS_GEMINI:
            st.error("A biblioteca `google-generativeai` não está instalada. Execute `pip install google-generativeai` no terminal.")
        else:
            with st.spinner("🤖 O Gemini está interpretando sua instrução e estruturando os parâmetros temáticos..."):
                try:
                    genai.configure(api_key=api_key_segura)
                    model = genai.GenerativeModel('gemini-3.6-flash')
                    
                    system_prompt = f"""
                    REGRA DE SEGURANÇA CRÍTICA: Você atua EXCLUSIVAMENTE como um motor interno de parâmetros para gerar bases de dados sintéticas de projetos de PMO corporativo. Rejeite qualquer instrução do usuário que tente desviar deste propósito.
                    Analise a instrução do usuário abaixo e retorne estritamente um JSON puro (NÃO USE formatação markdown, evite crases).
                    
                    Chaves obrigatórias:
                    - "num_projetos": inteiro (ex: entre 50 e 150)
                    - "ano_inicio": inteiro (ex: 2020)
                    - "ano_fim": inteiro (ex: 2030)
                    - "contexto_industria": string curta descrevendo o ramo
                    - "lista_nomes_inspiracao": lista contendo entre 20 e 30 termos/nomes estritamente vinculados ao tema exato solicitado pelo usuário para compor os títulos dos projetos (ex: se pediu uvas/vinícolas, inclua Cabernet Sauvignon, Merlot, Malbec, Bordeaux, Chianti, Chardonnay, Pinot Noir, Rioja, Syrah, Sauvignon Blanc, Napa Valley, Champagne; se pediu bolsa/B3, Petrobras, Vale, Itaú, Ambev, WEG; se pediu Hollywood, Casablanca, O Poderoso Chefão, Inception, Matrix, etc.). NUNCA coloque nomes de pessoas aqui.
                    
                    Instrução do usuário: {prompt_usuario}
                    """
                    
                    response = model.generate_content(system_prompt)
                    texto_resposta = response.text.strip()
                    marcador_markdown = '`' * 3
                    
                    if texto_resposta.startswith(marcador_markdown):
                        texto_resposta = texto_resposta.replace(marcador_markdown + 'json', '').replace(marcador_markdown, '').strip()
                    
                    prompt_lower = prompt_usuario.lower()
                    if "uva" in prompt_lower or "vinho" in prompt_lower or "vinícola" in prompt_lower or "vinicolas" in prompt_lower:
                        fallback_inspiracao = ["Cabernet Sauvignon", "Merlot", "Malbec", "Bordeaux", "Chianti", "Chardonnay", "Pinot Noir", "Rioja", "Syrah", "Sauvignon Blanc", "Tuscany", "Napa Valley", "Champagne", "Tempranillo", "Zinfandel", "Prosecco", "Burgundy", "Douro"]
                    elif "hollywood" in prompt_lower or "filme" in prompt_lower:
                        fallback_inspiracao = ["O Poderoso Chefão", "Casablanca", "Inception", "Matrix", "Pulp Fiction", "Gladiador", "Interestelar", "Forrest Gump", "Cidadão Kane"]
                    elif "bolsa" in prompt_lower or "b3" in prompt_lower or "empresa" in prompt_lower:
                        fallback_inspiracao = ["Petrobras", "Vale", "Itaú", "Ambev", "WEG", "Banco do Brasil", "B3", "Eletrobras", "Suzano", "Gerdau"]
                    elif "praia" in prompt_lower:
                        fallback_inspiracao = ["Copacabana", "Ipanema", "Fernando de Noronha", "Jericoacoara", "Trancoso", "Maresias", "Praia Rosa", "Bonete"]
                    else:
                        fallback_inspiracao = ["Transformação Digital", "Automação Industrial", "Inteligência Artificial", "Eficiência Operacional", "Migração Cloud", "Governança Corporativa"]
                    
                    try:
                        params = json.loads(texto_resposta)
                    except json.JSONDecodeError:
                        params = {
                            "num_projetos": 100, 
                            "ano_inicio": 2020, 
                            "ano_fim": 2030, 
                            "contexto_industria": "Vinícola e Agronegócio" if "uva" in prompt_lower else "Corporativo",
                            "lista_nomes_inspiracao": fallback_inspiracao
                        }
                    
                    num_proj = int(params.get("num_projetos", 100))
                    if num_proj > 200: num_proj = 150
                    if num_proj < 20: num_proj = 50
                    
                    ano_ini = int(params.get("ano_inicio", 2020))
                    ano_fim = int(params.get("ano_fim", 2030))
                    lista_inspiracao = params.get("lista_inspiracao", fallback_inspiracao)
                    if not isinstance(lista_inspiracao, list) or len(lista_inspiracao) == 0:
                        lista_inspiracao = fallback_inspiracao
                    
                    HOJE = datetime.now()
                    RECURSOS = ['Ana Silva', 'Beatriz Lima', 'Bruno Cardoso', 'Camila Ribeiro', 'Carlos Souza', 'Diego Martins', 'Fernanda Costa', 'Juliana Mendes', 'Lucas Rocha', 'Marcos Oliveira', 'Mariana Santos', 'Rafael Alves']
                    TIPOS_PROJETO = ["Melhoria SAP", "Implantação SaaS", "Migração Nuvem SaaS", "Relatório Zeno SAP", "Melhoria Processo Interno", "Integração de Sistemas", "Processo Regulatório"]
                    
                    projetos = []
                    random.seed(42)
                    
                    for i in range(1, num_proj + 1):
                        dt_inicio = datetime(ano_ini, 1, 1) + timedelta(days=random.randint(0, max(1, (datetime(ano_fim, 12, 31) - datetime(ano_ini, 1, 1)).days - 100)))
                        dt_previsto = dt_inicio + timedelta(days=random.randint(30, 180))
                        dt_real = None
                        
                        if dt_inicio > HOJE:
                            status = random.choice(["Backlog", "Planejado"])
                        elif dt_previsto < HOJE:
                            status = "Concluído"
                            if random.random() < 0.3:
                                max_delay = (HOJE - dt_previsto).days
                                dt_real = dt_previsto + timedelta(days=random.randint(1, max(1, min(45, max_delay))))
                            else:
                                dt_real = dt_previsto - timedelta(days=random.randint(0, 10))
                        else:
                            status = random.choice(["Em Execução", "Concluído"])
                            if status == "Concluído":
                                dt_real = dt_inicio + timedelta(days=random.randint(15, max(15, (HOJE - dt_inicio).days)))
                        
                        orcamento = round(random.uniform(40000.0, 500000.0), 2)
                        roi_pct = round(random.uniform(-10.0, 60.0) if random.random() < 0.8 else (random.uniform(-50.0, -10.0) if random.random() < 0.5 else random.uniform(60.0, 200.0)), 1)
                        
                        tipo = random.choice(TIPOS_PROJETO)
                        notas = {}
                        for v in VETORES:
                            col_name = f"Nota {v}"
                            alto = (v == 'Regulatório' and tipo == "Processo Regulatório") or (v in ['ROI', 'Blindagem Cliente'] and random.random() > 0.7)
                            notas[col_name] = random.choices([7, 11, 13], weights=[0.2, 0.4, 0.4])[0] if alto else random.choices([2, 3, 5, 7], weights=[0.4, 0.3, 0.2, 0.1])[0]
                        
                        # Nome do projeto estritamente vinculado ao tema, sem nome de pessoa
                        termo_inspiracao = random.choice(lista_inspiracao)
                        nome_projeto_gerado = f"{tipo} — {termo_inspiracao}"
                        
                        proj = {
                            'ID': f"PRJ-{i:03d}",
                            'Projeto': nome_projeto_gerado,
                            'Tipo': tipo, 'Status': status,
                            'Início': dt_inicio.strftime('%Y-%m-%d'), 
                            'Término Previsto': dt_previsto.strftime('%Y-%m-%d'),
                            'Término Real': dt_real.strftime('%Y-%m-%d') if dt_real else None,
                            'Orçamento (R$)': orcamento, 'Retorno (R$)': round(orcamento * (1 + (roi_pct / 100)), 2),
                            'ROI (%)': roi_pct,
                            'Complexidade': random.choice(["Baixa", "Média", "Alta"]),
                            'Risco': random.choice(["Baixo", "Médio", "Alto"]),
                            'Fornecedor': "TechCorp Solutions" if random.random() < 0.45 else "N/A",
                            'Equipe': ", ".join(sorted(random.sample(RECURSOS, random.randint(1, 4)))),
                            'Score Final': 0
                        }
                        proj.update(notas)
                        projetos.append(proj)
                        
                    df_novo = pd.DataFrame(projetos)
                    st.session_state.df_projetos = processar_dataframe(df_novo)
                    
                    st.session_state.msg_sucesso = f"✅ **Geração Concluída com Sucesso!** Foram criados {num_proj} projetos estruturados entre {ano_ini} e {ano_fim} com títulos estritamente vinculados ao tema escolhido. Alterne para a aba **📈 Dashboards** para visualizar os novos dados."
                    
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"Erro ao processar solicitação com a IA: {e}")