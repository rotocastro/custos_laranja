import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from custos_slp import extrair_dados

# Configuração da página
st.set_page_config(
    page_title="Dashboard de Custos Agrícolas",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Caminhos dos arquivos
CAMINHO_CUSTOS = "custos.xlsx"
CAMINHO_SAFRA = "safra.xlsx"


# Função cacheada para carregar os dados
@st.cache_data
def load_data(caminho_custos, caminho_safra):
    return extrair_dados(caminho_custos, caminho_safra)


# Função para criar gráfico de evolução dos custos
@st.cache_data
def plot_evolucao_custos(df_plot, titulo, ylabel):
    if 'Safra (cx)' in df_plot.columns:
        df_plot = df_plot.drop(columns=['Safra (cx)'])
    if 'Área (ha)' in df_plot.columns:
        df_plot = df_plot.drop(columns=['Área (ha)'])

    fig = px.bar(
        df_plot,
        barmode='group',
        title=titulo,
        labels={'value': ylabel, 'variable': 'Categoria'},
        color_discrete_sequence=['#2E86C1', '#28B463', '#D35400', '#884EA0']
    )
    fig.update_layout(
        height=400,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    return fig


# Função para formatar números
def format_number(value, format_type='money'):
    if pd.isna(value):
        return "-"
    if format_type == 'money':
        return f"R$ {value:,.2f}"
    elif format_type == 'area':
        return f"{value:,.2f} ha"
    elif format_type == 'percentage':
        return f"{value:.1f}%"
    return f"{value:,.2f}"


# Função para mostrar informações complementares
def show_complementary_info(df, tipo, fazenda_nome=""):
    if tipo == 'safra' and 'Safra (cx)' in df.columns:
        st.markdown(f"### Produção (caixas) - {fazenda_nome}")
        safra_data = df['Safra (cx)'].dropna()
        for index, value in safra_data.items():
            st.metric(f"{index}", f"{value:,.0f} cx")

    elif tipo == 'area' and 'Área (ha)' in df.columns:
        st.markdown(f"### Área - {fazenda_nome}")
        area = df['Área (ha)'].iloc[0]
        st.metric("Área Total", f"{area:,.2f} ha")


def main():
    st.title("Dashboard de Custos Agrícolas")

    try:
        # Carrega os dados
        dados = load_data(CAMINHO_CUSTOS, CAMINHO_SAFRA)
        fazendas = list(dados.keys())

        # Sidebar para seleção de fazendas
        with st.sidebar:
            st.title("🌱 Análise de Custos")
            st.markdown("---")
            fazendas_selecionadas = st.multiselect(
                "Selecione as Fazendas",
                fazendas,
                default=fazendas[-1] if fazendas else None
            )
            st.markdown("---")
            st.caption(f"Última atualização: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

        if fazendas_selecionadas:
            # Tabs principais no topo
            tab1, tab2, tab3 = st.tabs(["📊 Consolidado", "🌱 Por Área", "📦 Por Caixa"])

            # Conteúdo para cada tab
            for tab, view_type in [(tab1, 'consolidado'), (tab2, 'area'), (tab3, 'caixa')]:
                with tab:
                    for fazenda in fazendas_selecionadas:
                        dados_fazenda = dados[fazenda]
                        if 'dados_pivot' in dados_fazenda:
                            df_pivot = dados_fazenda['dados_pivot']
                            area = dados_fazenda['dados_por_area']['Área (ha)'].iloc[
                                0] if 'dados_por_area' in dados_fazenda else 0

                            # Header com métricas principais
                            st.markdown(f"## Fazenda: {fazenda} (Área: {area:,.2f} hectares)")
                            metrics_cols = st.columns(4)

                            with metrics_cols[0]:
                                total_orcado = df_pivot.loc['Orçado', 'TOTAL']
                                st.metric("Orçado Total", format_number(total_orcado))

                            ultimo_mes = df_pivot.index[-1] if len(df_pivot.index) > 1 else None
                            if ultimo_mes:
                                with metrics_cols[1]:
                                    total_realizado = df_pivot.loc[ultimo_mes, 'TOTAL']
                                    variacao = ((total_realizado - total_orcado) / total_orcado) * 100
                                    st.metric(f"Realizado {ultimo_mes}",
                                              format_number(total_realizado),
                                              f"{variacao:+.1f}%")

                            # Custos por caixa
                            if 'dados_por_safra' in dados_fazenda:
                                with metrics_cols[2]:
                                    custo_cx_orcado = dados_fazenda['dados_por_safra'].loc['Orçado', 'TOTAL']
                                    st.metric("Custo por Caixa Orçado", format_number(custo_cx_orcado, 'money'))

                                if ultimo_mes:
                                    with metrics_cols[3]:
                                        custo_cx_realizado = dados_fazenda['dados_por_safra'].loc[ultimo_mes, 'TOTAL']
                                        st.metric(f"Custo por Caixa {ultimo_mes}",
                                                  format_number(custo_cx_realizado, 'money'))

                            st.markdown("---")

                            # Conteúdo específico para cada view_type
                            if view_type == 'consolidado':
                                st.subheader(f"Visão Consolidada - {fazenda}")
                                fig = plot_evolucao_custos(
                                    dados_fazenda['dados_pivot'],
                                    f"Evolução dos Custos - {fazenda}",
                                    "Valor (R$)"
                                )
                                st.plotly_chart(fig, use_container_width=True)
                                st.dataframe(
                                    dados_fazenda['dados_pivot'].style.format("R$ {:,.2f}"),
                                    use_container_width=True
                                )

                            elif view_type == 'area' and 'dados_por_area' in dados_fazenda:
                                col1, col2 = st.columns([3, 1])
                                with col1:
                                    st.subheader(f"Visão por Área - {fazenda}")
                                    fig = plot_evolucao_custos(
                                        dados_fazenda['dados_por_area'],
                                        f"Evolução dos Custos por Área - {fazenda}",
                                        "R$/ha"
                                    )
                                    st.plotly_chart(fig, use_container_width=True)
                                    st.dataframe(
                                        dados_fazenda['dados_por_area'].style.format("R$ {:,.2f}"),
                                        use_container_width=True
                                    )
                                with col2:
                                    show_complementary_info(dados_fazenda['dados_por_area'], 'area', fazenda)

                            elif view_type == 'caixa' and 'dados_por_safra' in dados_fazenda:
                                col1, col2 = st.columns([3, 1])
                                with col1:
                                    st.subheader(f"Visão por Caixa - {fazenda}")
                                    fig = plot_evolucao_custos(
                                        dados_fazenda['dados_por_safra'],
                                        f"Evolução dos Custos por Caixa - {fazenda}",
                                        "R$/cx"
                                    )
                                    st.plotly_chart(fig, use_container_width=True)
                                    st.dataframe(
                                        dados_fazenda['dados_por_safra'].style.format("R$ {:,.2f}"),
                                        use_container_width=True
                                    )
                                with col2:
                                    show_complementary_info(dados_fazenda['dados_por_safra'], 'safra', fazenda)

    except Exception as e:
        st.error(f"Erro ao carregar dados: {str(e)}")


if __name__ == "__main__":
    main()
