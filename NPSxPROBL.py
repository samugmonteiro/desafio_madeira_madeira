import pandas as pd

# Carregar o arquivo Excel
print("Lendo arquivo 'dados_unificados_com_janela_analise.xlsx'...")
df = pd.read_excel('dados_unificados_com_janela_analise.xlsx', engine='openpyxl')

# Verificar colunas disponíveis
print(f"Colunas disponíveis: {list(df.columns)}")

# Verificar se as colunas necessárias existem
colunas_necessarias = ['janela_analise_valida', 'classificacao_nps', 'macro_problema_nps', 'tipo_negocio',
                       'data_resposta_nps']
if not all(col in df.columns for col in colunas_necessarias):
    print(
        "⚠️ Colunas 'janela_analise_valida', 'classificacao_nps', 'macro_problema_nps', 'tipo_negocio' ou 'data_resposta_nps' não encontradas.")
    print("Execute scripts anteriores para criar 'classificacao_nps' se necessário.")
    exit()

# 1. Filtrar apenas respostas dentro da janela de análise
df_filtrado = df[df['janela_analise_valida'] == 'Dentro da janela de análise'].copy()
linhas_apos_filtro_janela = len(df_filtrado)
print(f"Linhas após filtro 'Dentro da janela': {linhas_apos_filtro_janela}")

if linhas_apos_filtro_janela == 0:
    print("⚠️ Nenhum dado dentro da janela de análise. Verifique os dados.")
    exit()

# 2. Verificar distribuição de classificacao_nps no filtro
print("\nDistribuição de classificacao_nps (após filtro de janela):")
distribuicao_nps = df_filtrado['classificacao_nps'].value_counts(dropna=False)
print(distribuicao_nps)
total_com_nps = distribuicao_nps.get('Promotor', 0) + distribuicao_nps.get('Detrator', 0) + distribuicao_nps.get(
    'Neutro', 0)
print(f"Total com classificação NPS válida: {total_com_nps} (de {linhas_apos_filtro_janela})")

# 3. Remover linhas sem classificação NPS (NA)
df_analise = df_filtrado.dropna(subset=['classificacao_nps'])
linhas_apos_dropna = len(df_analise)
print(f"Linhas após remover NaN em classificacao_nps: {linhas_apos_dropna}")

if linhas_apos_dropna == 0:
    print(
        "⚠️ Nenhuma linha com classificacao_nps válida após filtros. Verifique a coluna 'nota_nps' no arquivo original.")
    print("Dica: Rode o script de complementação para criar 'classificacao_nps'.")
    exit()

# 4. Verificar distribuição de macro_problema_nps
print("\nMacro_problemas disponíveis (após filtros):")
macro_problemas = df_analise['macro_problema_nps'].value_counts(dropna=False)
print(macro_problemas)
linhas_com_macro = len(df_analise.dropna(subset=['macro_problema_nps']))
print(f"Linhas com macro_problema_nps válido: {linhas_com_macro} (de {linhas_apos_dropna})")

if linhas_com_macro == 0:
    print("⚠️ Nenhuma linha com macro_problema_nps válido. Verifique a coluna no Excel.")
    exit()

# 5. Filtrar apenas com macro_problema_nps válidos
df_analise = df_analise.dropna(subset=['macro_problema_nps'])

# 6. Extrair trimestre de data_resposta_nps
df_analise['data_resposta_nps'] = pd.to_datetime(df_analise['data_resposta_nps'], errors='coerce')
df_analise['trimestre_nps'] = df_analise['data_resposta_nps'].dt.to_period('Q')

# Filtrar apenas com trimestres válidos
df_analise = df_analise.dropna(subset=['trimestre_nps'])
linhas_com_trimestre = len(df_analise)
print(f"Linhas com trimestre válido: {linhas_com_trimestre} (de {len(df_analise)} após filtros anteriores)")

if linhas_com_trimestre == 0:
    print("⚠️ Nenhuma linha com data_resposta_nps válida para extrair trimestres.")
    exit()

# Verificar distribuição de trimestres
print("\nTrimestres disponíveis (após filtros):")
trimestres = df_analise['trimestre_nps'].value_counts().sort_index()
print(trimestres)

# 7. ANÁLISE GERAL: Contagens e % por trimestre + macro_problema + classificacao_nps
print("\n--- ANÁLISE GERAL: Macro_Problema_NPS por Classificação NPS e Trimestre ---")
agrupado_geral = df_analise.groupby(['trimestre_nps', 'macro_problema_nps', 'classificacao_nps']).size().unstack(
    fill_value=0)

# Calcular totais e percentuais por trimestre + macro_problema
agrupado_geral['total'] = agrupado_geral.sum(axis=1)
for col in ['Promotor', 'Detrator', 'Neutro']:
    if col in agrupado_geral.columns:
        agrupado_geral[f'% {col}'] = (agrupado_geral[col] / agrupado_geral['total']) * 100
    else:
        agrupado_geral[f'% {col}'] = 0.0

# Arredondar percentuais
percentuais_cols = [col for col in agrupado_geral.columns if col.startswith('%')]
agrupado_geral[percentuais_cols] = agrupado_geral[percentuais_cols].round(2)

# Exibir por trimestre (uma sub-tabela por trimestre, ordenada por total)
for trimestre in sorted(agrupado_geral.index.get_level_values('trimestre_nps').unique()):
    print(f"\n  --- Trimestre {trimestre} ---")
    sub_agrupado = agrupado_geral.loc[trimestre].sort_values('total', ascending=False)
    print(sub_agrupado[['total', 'Promotor', 'Detrator', 'Neutro', '% Promotor', '% Detrator', '% Neutro']])

# Salvar geral em CSV
agrupado_geral.to_csv('analise_macro_problema_nps_geral_por_trimestre.csv')
print("✅ Resultados gerais (por trimestre) salvos em 'analise_macro_problema_nps_geral_por_trimestre.csv'.")

# 8. ANÁLISE POR MODELO DE NEGÓCIO (com Trimestre)
print("\n--- ANÁLISE POR MODELO DE NEGÓCIO (com Trimestre) ---")

# Verificar distribuição de tipo_negocio
modelos = df_analise['tipo_negocio'].value_counts()
print("Modelos de negócio disponíveis (após filtros):")
print(modelos)

if len(modelos) == 0:
    print("⚠️ Nenhum modelo de negócio válido encontrado.")
else:
    # Agrupar por modelo + trimestre + macro_problema + classificacao_nps
    agrupado_modelo = df_analise.groupby(
        ['tipo_negocio', 'trimestre_nps', 'macro_problema_nps', 'classificacao_nps']).size().unstack(fill_value=0)

    # Calcular totais e percentuais por modelo + trimestre + macro_problema
    agrupado_modelo['total'] = agrupado_modelo.sum(axis=1)
    for col in ['Promotor', 'Detrator', 'Neutro']:
        if col in agrupado_modelo.columns:
            agrupado_modelo[f'% {col}'] = (agrupado_modelo[col] / agrupado_modelo['total']) * 100
        else:
            agrupado_modelo[f'% {col}'] = 0.0

    # Arredondar percentuais
    percentuais_cols_modelo = [col for col in agrupado_modelo.columns if col.startswith('%')]
    agrupado_modelo[percentuais_cols_modelo] = agrupado_modelo[percentuais_cols_modelo].round(2)

    # Exibir por modelo e trimestre (uma sub-tabela por modelo + trimestre, ordenada por total)
    for modelo in sorted(agrupado_modelo.index.get_level_values('tipo_negocio').unique()):
        print(f"\n  --- {modelo} ---")
        sub_modelo = agrupado_modelo.loc[(modelo, slice(None), slice(None))]
        for trimestre in sorted(sub_modelo.index.get_level_values('trimestre_nps').unique()):
            print(f"    --- Trimestre {trimestre} ---")
            sub_tabela = sub_modelo.loc[(modelo, trimestre)].sort_values('total', ascending=False)
            print(sub_tabela[['total', 'Promotor', 'Detrator', 'Neutro', '% Promotor', '% Detrator', '% Neutro']])

    # Salvar por modelo em CSV
    agrupado_modelo.to_csv('analise_macro_problema_nps_por_modelo_e_trimestre.csv')
    print(
        f"\n✅ Resultados por modelo e trimestre salvos em 'analise_macro_problema_nps_por_modelo_e_trimestre.csv' (total de linhas: {len(agrupado_modelo)}).")

# Resumo Geral Final (para referência)
total_geral = agrupado_geral['total'].sum()
print(f"\nResumo Geral Final: Total de respostas analisadas = {total_geral}")
