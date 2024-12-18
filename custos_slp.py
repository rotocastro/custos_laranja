import pandas as pd
import openpyxl

# Configurações para a exibição das tabelas
pd.set_option('display.float_format', '{:.0f}'.format)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('future.no_silent_downcasting', True)


def encontrar_ultimo_mes_realizado(df, meses):
    """
    Encontra o último mês que tem Realizado_OPEX
    """
    linha_total = df[df['ITEM'] == 'TOTAL'].iloc[0]
    ultimo_mes = None

    for mes in meses:
        if linha_total[f'{mes}_Realizado_OPEX'] > 0:
            ultimo_mes = mes

    return ultimo_mes

def calcular_total_ajustado(df, meses):
    """
    Calcula o total ajustado usando a linha TOTAL
    """
    totais_ajustados = {}
    linha_total = df[df['ITEM'] == 'TOTAL'].iloc[0]

    # Adiciona o orçado (PREVISTO TOTAL OPEX)
    totais_ajustados['Orçado'] = linha_total['PREVISTO TOTAL OPEX']

    # Encontra o último mês com Realizado
    ultimo_mes_realizado = None
    for i, mes in enumerate(meses):
        if linha_total[f'{mes}_Realizado_OPEX'] > 0:
            ultimo_mes_realizado = i

    if ultimo_mes_realizado is not None:
        # Calcula normalmente até o último mês com Realizado
        for i, mes_atual in enumerate(meses[:ultimo_mes_realizado + 1]):
            total = 0

            # Soma Realizado_OPEX até o mês atual
            for mes in meses[:i + 1]:
                total += linha_total[f'{mes}_Realizado_OPEX']

            # Soma Previsto_OPEX dos meses seguintes
            for mes in meses[i + 1:]:
                total += linha_total[f'{mes}_Previsto_OPEX']

            totais_ajustados[mes_atual] = total

    return totais_ajustados


def calcular_total_ajustado_por_grupo(df, meses, itens_grupo):
    """
    Calcula o total ajustado para um grupo específico de itens
    """
    totais_ajustados = {}
    linhas_grupo = df[df['ITEM'].isin(itens_grupo)]

    # Adiciona o orçado
    totais_ajustados['Orçado'] = linhas_grupo['PREVISTO TOTAL OPEX'].sum()

    # Encontra o último mês com Realizado
    ultimo_mes_realizado = None
    for i, mes in enumerate(meses):
        if linhas_grupo[f'{mes}_Realizado_OPEX'].sum() > 0:
            ultimo_mes_realizado = i

    if ultimo_mes_realizado is not None:
        for i, mes_atual in enumerate(meses[:ultimo_mes_realizado + 1]):
            total = 0

            # Soma Realizado_OPEX até o mês atual
            for mes in meses[:i + 1]:
                total += linhas_grupo[f'{mes}_Realizado_OPEX'].sum()

            # Soma Previsto_OPEX dos meses seguintes
            for mes in meses[i + 1:]:
                total += linhas_grupo[f'{mes}_Previsto_OPEX'].sum()

            totais_ajustados[mes_atual] = total

            # Adiciona os meses faltantes com o valor do previsto OPEX
    for mes in meses:
        if mes not in totais_ajustados:
            totais_ajustados[mes] = linhas_grupo[f'{mes}_Previsto_OPEX'].sum()

    return totais_ajustados


def criar_dataframe_pivot_unitario(df_pivot, fazenda_info):
    """
    Cria DataFrame com valores unitários (por área e por safra)
    """
    area = fazenda_info['area'].iloc[0]
    estimativa = fazenda_info['Estimativa_Inicial'].iloc[0]
    safras = {
        'Jun24': fazenda_info['Safra Jun/24'].iloc[0],
        'Jul24': fazenda_info['Safra Jul/24'].iloc[0],
        'Ago24': fazenda_info['Safra Ago/24'].iloc[0],
        'Set24': fazenda_info['Safra Set/24'].iloc[0],
        'Out24': fazenda_info['Safra Out/24'].iloc[0],
        'Nov24': fazenda_info['Safra Nov/24'].iloc[0]
    }

    # Cria cópia do DataFrame original
    df_por_area = df_pivot.copy()
    df_por_safra = df_pivot.copy()

    # Divide todos os valores pela área
    df_por_area = df_por_area / area
    # Adiciona coluna com a área
    df_por_area['Área (ha)'] = area

    # Divide cada linha pelo valor correspondente
    if estimativa != 0:
        df_por_safra.loc['Orçado'] = df_por_safra.loc['Orçado'] / estimativa
        df_por_safra.loc['Orçado', 'Safra (cx)'] = estimativa

    # Cada mês divide pela sua safra correspondente
    for mes, safra in safras.items():
        if mes in df_por_safra.index and safra != 0:
            df_por_safra.loc[mes] = df_por_safra.loc[mes] / safra
            df_por_safra.loc[mes, 'Safra (cx)'] = safra

    return df_por_area, df_por_safra


def extrair_dados(caminho_arquivo, caminho_arquivo_safra):
    # Lê o arquivo de safra
    df_safra = pd.read_excel(caminho_arquivo_safra)

    # Lista todas as fazendas do arquivo de safra (exceto os totais)
    fazendas = df_safra[~df_safra['Fazenda'].isin(['TOTAL'])]['Fazenda'].tolist()

    # Define os grupos de itens
    insumos = ['Adubo', 'Corretivo de Solo', 'Fertilizante Orgânico', 'Semente',
               'Herbicidas', 'Fungicida', 'Inseticida', 'Acaricida', 'Óleo',
               'Reguladores Vegetais']

    dfs = {}

    for aba in pd.read_excel(caminho_arquivo, sheet_name=None).keys():
        df = pd.read_excel(
            caminho_arquivo, sheet_name=aba,
            skiprows=3, nrows=55, usecols="A:CQ")

        novos_nomes = {}

        # Ajuste das 11 primeiras colunas
        colunas_iniciais = {
            df.columns[0]: 'ITEM',
            df.columns[1]: 'PREVISTO TOTAL 2024/2025',
            df.columns[2]: 'PREVISTO TOTAL OPEX',
            df.columns[3]: 'PREVISTO TOTAL CAPEX',
            df.columns[4]: 'PREVISTO PERIODO',
            df.columns[5]: 'PREVISTO PERIODO OPEX',
            df.columns[6]: 'PREVISTO PERIODO CAPEX',
            df.columns[7]: 'REALIZADO OPEX',
            df.columns[8]: 'REALIZADO CAPEX',
            df.columns[9]: 'REALIZADO TOTAL',
            df.columns[10]: '%'
        }
        novos_nomes.update(colunas_iniciais)

        meses = ['Jun24', 'Jul24', 'Ago24', 'Set24', 'Out24', 'Nov24', 'Dez24',
                 'Jan25', 'Fev25', 'Mar25', 'Abr25', 'Mai25']

        col_atual = 11
        for mes in meses:
            novos_nomes[df.columns[col_atual]] = f'{mes}_Previsto'
            novos_nomes[df.columns[col_atual + 1]] = f'{mes}_Previsto_OPEX'
            novos_nomes[df.columns[col_atual + 2]] = f'{mes}_Previsto_CAPEX'
            novos_nomes[df.columns[col_atual + 3]] = f'{mes}_Realizado_OPEX'
            novos_nomes[df.columns[col_atual + 4]] = f'{mes}_Realizado_CAPEX'
            novos_nomes[df.columns[col_atual + 5]] = f'{mes}_Realizado_Total'
            novos_nomes[df.columns[col_atual + 6]] = f'{mes}_Percentual'
            col_atual += 7

        df = df.rename(columns=novos_nomes)

        # Identificação dos itens de colheita sem warning
        mask = (df['ITEM'].notna() &
                df['ITEM'].str.startswith(('Colheita', 'Frete'))
                .fillna(False)
                .infer_objects(copy=False))
        colheita = df[mask]['ITEM'].unique().tolist()

        # Encontra o último mês com realizado
        ultimo_mes = encontrar_ultimo_mes_realizado(df, meses)

        # Calcula os totais
        totais_geral = calcular_total_ajustado(df, meses)
        totais_insumos = calcular_total_ajustado_por_grupo(df, meses, insumos)
        totais_colheita = calcular_total_ajustado_por_grupo(df, meses, colheita)

        # Calcula CUSTOS FIXOS como TOTAL - INSUMOS - COLHEITA
        totais_fixos = {}
        totais_fixos['Orçado'] = (totais_geral['Orçado'] -
                                  totais_insumos['Orçado'] -
                                  totais_colheita['Orçado'])

        # Filtra apenas os meses até o último com realizado
        meses_filtrados = ['Orçado']  # Sempre mantém o Orçado
        if ultimo_mes:
            indice_ultimo_mes = meses.index(ultimo_mes)
            meses_filtrados.extend(meses[:indice_ultimo_mes + 1])

            for mes in meses[:indice_ultimo_mes + 1]:
                if mes in totais_geral and mes in totais_insumos and mes in totais_colheita:
                    totais_fixos[mes] = (totais_geral[mes] -
                                         totais_insumos[mes] -
                                         totais_colheita[mes])

        # Cria o DataFrame pivotado apenas com os meses que têm realizado
        df_pivot = pd.DataFrame({
            'TOTAL': {mes: totais_geral[mes] for mes in meses_filtrados},
            'INSUMOS': {mes: totais_insumos[mes] for mes in meses_filtrados},
            'COLHEITA': {mes: totais_colheita[mes] for mes in meses_filtrados},
            'FIXOS': {mes: totais_fixos[mes] for mes in meses_filtrados}
        })

        # Se a aba corresponder a uma fazenda, cria as visões unitárias
        if aba in fazendas:
            fazenda_info = df_safra[df_safra['Fazenda'] == aba]

            if not fazenda_info.empty:
                df_por_area, df_por_safra = criar_dataframe_pivot_unitario(df_pivot, fazenda_info)

                print(f"\nAba: {aba}")
                print("\nVisão Consolidada (valores totais):")
                print(df_pivot.round(0))
                print("\nVisão por Área (R$/ha):")
                print(df_por_area.round(0))
                print("\nVisão por Caixa (R$/cx):")
                pd.set_option('display.float_format', '{:.2f}'.format)
                print(df_por_safra)
                pd.set_option('display.float_format', '{:.0f}'.format)

                dfs[aba] = {
                    'dados_originais': df,
                    'dados_pivot': df_pivot,
                    'dados_por_area': df_por_area,
                    'dados_por_safra': df_por_safra
                }
            else:
                print(f"\nAtenção: Fazenda {aba} não encontrada na tabela de safra")
                dfs[aba] = {
                    'dados_originais': df,
                    'dados_pivot': df_pivot
                }
        else:
            dfs[aba] = {
                'dados_originais': df,
                'dados_pivot': df_pivot
            }

    return dfs

# def exportar_resultados(dados, nome_arquivo='resultados.xlsx'):
#     with pd.ExcelWriter(nome_arquivo) as writer:
#         for fazenda in dados.keys():
#             if 'dados_pivot' in dados[fazenda]:
#                 dados[fazenda]['dados_pivot'].to_excel(writer, sheet_name=f'{fazenda}_Totais')
#                 if 'dados_por_area' in dados[fazenda]:
#                     dados[fazenda]['dados_por_area'].to_excel(writer, sheet_name=f'{fazenda}_por_ha')
#                 if 'dados_por_safra' in dados[fazenda]:
#                     # Formata números com 2 casas decimais apenas para dados_por_safra
#                     dados[fazenda]['dados_por_safra'].round(2).to_excel(writer, sheet_name=f'{fazenda}_por_cx')


if __name__ == "__main__":
    # Código que só será executado se o arquivo for rodado diretamente
    caminho_arquivo = "custos.xlsx"
    caminho_arquivo_safra = "safra.xlsx"
    dados = extrair_dados(caminho_arquivo, caminho_arquivo_safra)
    # exportar_resultados(dados)
