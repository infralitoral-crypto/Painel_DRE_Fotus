import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 1. Configurações Globais da Página
st.set_page_config(page_title="Torre de Controle DRE - Fotus", page_icon="☀️", layout="wide", initial_sidebar_state="expanded")

# CSS customizado para melhorar o visual dos cartões (Métricas)
st.markdown("""
    <style>
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        padding: 5% 5% 5% 10%;
        border-radius: 10px;
        box-shadow: 2px 2px 10px rgba(0, 0, 0, 0.05);
    }
    </style>
""", unsafe_allow_html=True)

# 2. Sidebar Profissional
col_logo1, col_logo2, col_logo3 = st.sidebar.columns([1, 4, 1])
with col_logo2:
    if os.path.exists("logo_fotus.png"):
        st.image("logo_fotus.png", use_container_width=True)
st.sidebar.markdown("---")
st.sidebar.title("Filtros Executivos")

# 3. Motor de Dados Otimizado (COM CORREÇÃO DE TIPAGEM)
@st.cache_data
def carregar_dados():
    caminho_arquivo = "DRE - POR CD.xlsx"
    df_custos = pd.read_excel(caminho_arquivo, sheet_name="BD_Cust.Desp", engine='openpyxl')
    df_faturamento = pd.read_excel(caminho_arquivo, sheet_name="BD_Fat", engine='openpyxl')
    df_hc = pd.read_excel(caminho_arquivo, sheet_name="HC", engine='openpyxl')
    
    # Normalização de Nomes: Verifica se o mês em texto está na coluna Mês1 ou MÊS
    if 'Mês1' in df_custos.columns:
        df_custos.rename(columns={'Mês1': 'Mês'}, inplace=True)
    else:
        df_custos.rename(columns={'MÊS': 'Mês'}, inplace=True, errors='ignore')
        
    df_custos.rename(columns={'Descrição (Natureza)': 'Natureza', 'Nome Parceiro (Parceiro)': 'Fornecedor'}, inplace=True, errors='ignore')
    
    for df in [df_custos, df_faturamento]:
        if 'Vlr. Nota' in df.columns:
            df['Vlr. Nota'] = pd.to_numeric(df['Vlr. Nota'], errors='coerce').fillna(0)
            
    # Ordem cronológica dos meses
    ordem_meses = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
    
    # Tratamento seguro da coluna Mês (Garante que tudo seja lido como Texto para evitar o AttributeError)
    for df in [df_custos, df_faturamento, df_hc]:
        if 'Mês' in df.columns:
            # Converte qualquer número, data ou vazio para string ANTES de usar o .str
            df['Mês'] = df['Mês'].astype(str).str.lower().str.strip()
            # Limpa valores que eram vazios (NaN) e viraram a palavra 'nan'
            df['Mês'] = df['Mês'].replace('nan', None)
            # Aplica a ordem cronológica
            df['Mês'] = pd.Categorical(df['Mês'], categories=ordem_meses, ordered=True)
            
    return df_custos, df_faturamento, df_hc

try:
    df_custos, df_faturamento, df_hc = carregar_dados()
except FileNotFoundError:
    st.error("⚠️ Ficheiro 'DRE - POR CD.xlsx' não encontrado. Verifique se está na mesma pasta.")
    st.stop()

# 4. Construção dos Filtros Dinâmicos
lista_empresas = df_custos['Empresa'].dropna().unique().tolist()
lista_meses = [m for m in ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'] if m in df_custos['Mês'].unique()]

empresa_selecionada = st.sidebar.selectbox("🏢 Filial / Unidade:", ["Visão Geral (Todas)"] + sorted(lista_empresas))
meses_selecionados = st.sidebar.multiselect("📅 Período (Meses):", options=lista_meses, default=lista_meses)

st.sidebar.markdown("---")
st.sidebar.info("💡 **Dica:** Utilize a seleção de múltiplos meses para avaliar o acumulado do trimestre ou semestre.")

# Aplicação dos Filtros em Cascata
def aplicar_filtros(df):
    if df.empty: return df
    if empresa_selecionada != "Visão Geral (Todas)" and 'Empresa' in df.columns:
        df = df[df['Empresa'] == empresa_selecionada]
    if meses_selecionados and 'Mês' in df.columns:
        df = df[df['Mês'].isin(meses_selecionados)]
    return df

df_c_filtrado = aplicar_filtros(df_custos)
df_f_filtrado = aplicar_filtros(df_faturamento)
df_h_filtrado = aplicar_filtros(df_hc)

# 5. Painel Principal (Header)
st.title("📊 Torre de Controle: Desempenho Operacional Fotus")
st.markdown("Consolidação Analítica de DRE por Centro de Distribuição.")
st.divider()

# 6. Indicadores Chave de Desempenho (KPIs com Cálculo Financeiro)
custo_total = df_c_filtrado['Vlr. Nota'].sum()
faturamento_total = df_f_filtrado['Vlr. Nota'].sum() if 'Vlr. Nota' in df_f_filtrado.columns else 0
margem_operacional = faturamento_total - custo_total
efetivo_total = df_h_filtrado['Quantidade Ativos'].sum() if 'Quantidade Ativos' in df_h_filtrado.columns else 0

def formata_brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
col_kpi1.metric("Faturamento Líquido", formata_brl(faturamento_total), delta="Entradas", delta_color="normal")
col_kpi2.metric("Despesas Operacionais", formata_brl(custo_total), delta="- Saídas", delta_color="inverse")
col_kpi3.metric("Margem (Fat - Custo)", formata_brl(margem_operacional), delta="Resultado", delta_color="normal" if margem_operacional >=0 else "inverse")
col_kpi4.metric("Efetivo (Headcount)", int(efetivo_total), delta="Colaboradores Ativos", delta_color="off")

st.markdown("<br>", unsafe_allow_html=True)

# 7. Abas de Navegação Analítica
tab_estrategia, tab_detalhamento, tab_fornecedores = st.tabs(["🎯 Visão Estratégica", "🔎 Raio-X de Custos", "🤝 Análise de Fornecedores"])

with tab_estrategia:
    col_chart1, col_chart2 = st.columns([2, 1])
    
    with col_chart1:
        st.subheader("Evolução Temporal: Faturamento vs Custo")
        if 'Mês' in df_c_filtrado.columns and 'Mês' in df_f_filtrado.columns:
            df_trend_c = df_c_filtrado.groupby('Mês', observed=True)['Vlr. Nota'].sum().reset_index()
            df_trend_c['Categoria'] = 'Custos'
            
            df_trend_f = df_f_filtrado.groupby('Mês', observed=True)['Vlr. Nota'].sum().reset_index()
            df_trend_f['Categoria'] = 'Faturamento'
            
            df_trend_completo = pd.concat([df_trend_f, df_trend_c])
            
            fig_trend = px.line(
                df_trend_completo, x='Mês', y='Vlr. Nota', color='Categoria', markers=True,
                color_discrete_map={'Faturamento': '#28a745', 'Custos': '#dc3545'},
                labels={'Vlr. Nota': 'Volume Financeiro (R$)'}
            )
            fig_trend.update_layout(plot_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis_title="")
            fig_trend.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#e0e0e0')
            st.plotly_chart(fig_trend, use_container_width=True)
            
    with col_chart2:
        st.subheader("Concentração por Unidade")
        df_unidades = df_c_filtrado.groupby('Empresa')['Vlr. Nota'].sum().reset_index()
        fig_unidades = px.bar(
            df_unidades.sort_values(by='Vlr. Nota', ascending=True), 
            x='Vlr. Nota', y='Empresa', orientation='h', text_auto='.2s',
            color_discrete_sequence=["#002040"]
        )
        fig_unidades.update_layout(plot_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis_title="")
        st.plotly_chart(fig_unidades, use_container_width=True)

with tab_detalhamento:
    col_d1, col_d2 = st.columns(2)
    
    with col_d1:
        st.subheader("Mapeamento Hierárquico de Despesas")
        st.markdown("*Clique nas fatias centrais (Tipo) para expandir as Naturezas.*")
        if 'Tipo' in df_c_filtrado.columns and 'Natureza' in df_c_filtrado.columns:
            df_hierarquia = df_c_filtrado.groupby(['Tipo', 'Natureza'])['Vlr. Nota'].sum().reset_index()
            df_hierarquia = df_hierarquia[df_hierarquia['Vlr. Nota'] > 0] # Limpa valores zerados
            
            fig_sunburst = px.sunburst(
                df_hierarquia, path=['Tipo', 'Natureza'], values='Vlr. Nota',
                color='Tipo', color_discrete_map={'Custo Fixo': '#002040', 'Custo Variável': '#FDB913'}
            )
            fig_sunburst.update_layout(margin=dict(t=0, l=0, r=0, b=0))
            st.plotly_chart(fig_sunburst, use_container_width=True)

    with col_d2:
        st.subheader("Ranking: Maiores Ofensores de Custo")
        if 'Natureza' in df_c_filtrado.columns:
            top_naturezas = df_c_filtrado.groupby('Natureza')['Vlr. Nota'].sum().nlargest(10).reset_index()
            fig_top_nat = px.bar(
                top_naturezas, x='Vlr. Nota', y='Natureza', orientation='h', text_auto='.2s',
                color='Vlr. Nota', color_continuous_scale='Reds'
            )
            fig_top_nat.update_layout(yaxis={'categoryorder':'total ascending'}, plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_top_nat, use_container_width=True)

with tab_fornecedores:
    st.subheader("Volume Transacionado por Parceiro/Fornecedor")
    if 'Fornecedor' in df_c_filtrado.columns:
        # Treemap para visualização de peso de fornecedores
        df_fornecedores = df_c_filtrado.groupby(['Tipo', 'Fornecedor'])['Vlr. Nota'].sum().reset_index()
        df_fornecedores = df_fornecedores.nlargest(30, 'Vlr. Nota') # Pega os 30 maiores para não poluir a vista
        
        fig_treemap = px.treemap(
            df_fornecedores, path=[px.Constant("Todos os Fornecedores"), 'Tipo', 'Fornecedor'], values='Vlr. Nota',
            color='Vlr. Nota', color_continuous_scale='Blues'
        )
        fig_treemap.update_layout(margin=dict(t=20, l=0, r=0, b=0))
        st.plotly_chart(fig_treemap, use_container_width=True)
        
    st.markdown("---")
    st.subheader("Base de Lançamentos Refinada")
    colunas_tabela = [c for c in ['Empresa', 'Mês', 'Nro. Nota', 'Fornecedor', 'Natureza', 'Tipo', 'Vlr. Nota'] if c in df_c_filtrado.columns]
    st.dataframe(df_c_filtrado[colunas_tabela].sort_values(by='Vlr. Nota', ascending=False), use_container_width=True, hide_index=True)