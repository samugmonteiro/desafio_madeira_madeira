import pandas as pd

# Carregar o arquivo Excel
print("Lendo arquivo 'dados_unificados_com_janela_analise.xlsx'...")
df = pd.read_excel('dados_unificados_com_janela_analise.xlsx', engine='openpyxl')

# Verificar colunas disponíveis
print(f"Colunas disponíveis: {list(df.columns)}")

# Garantir que data_resposta_nps esteja no formato datetime e extrair trimestre
if 'data_resposta_nps' in df.columns:
    df['data_resposta_nps'] = pd.to_datetime(df['data_resposta_nps'], errors='coerce')
    df['trimestre_resposta'] = df['data_resposta_nps'].dt.to_period('Q')
    print(f"Trimestres extraídos. Exemplo: {df['trimestre_resposta'].dropna().head()}")
else:
    print("⚠️ Coluna 'data_resposta_nps' não encontrada. Script interrompido.")
    exit()

# Verificar se as colunas necessárias existem
colunas_necessarias = ['janela_analise_valida', 'classificacao_nps']
if not all(col in df.columns for col in colunas_necessarias):
    print("⚠️ Colunas 'janela_analise_valida' ou 'classificacao_nps' não encontradas.")
    print("Execute scripts anteriores para criar 'classificacao_nps'.")
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
total_com_nps = distribuicao_nps.get('Promotor', 0) + distribuicao_nps.get('Detrator', 0) + distribuicao_nps.get('Neutro', 0)
print(f"Total com classificação NPS válida: {total_com_nps} (de {linhas_apos_filtro_janela})")

# 3. Remover linhas sem classificação NPS (NA)
df_analise = df_filtrado.dropna(subset=['classificacao_nps'])
linhas_apos_dropna = len(df_analise)
print(f"Linhas após remover NaN em classificacao_nps: {linhas_apos_dropna}")

if linhas_apos_dropna == 0:
    print("⚠️ Nenhuma linha com classificacao_nps válida após filtros. Verifique a coluna 'nota_nps' no arquivo original.")
    print("Dica: Rode o script de complementação para criar 'classificacao_nps'.")
    exit()

# 4. Verificar trimestres válidos
trimestres_validos = df_analise['trimestre_resposta'].dropna().unique()
print(f"\nTrimestres com dados válidos: {sorted(trimestres_validos)}")
linhas_com_trimestre = len(df_analise.dropna(subset=['trimestre_resposta']))
print(f"Linhas com trimestre válido: {linhas_com_trimestre} (de {linhas_apos_dropna})")

if linhas_com_trimestre == 0:
    print("⚠️ Nenhuma linha com data_resposta_nps válida para extrair trimestres.")
    exit()

# 5. Filtrar apenas com trimestres válidos e agrupar
df_analise = df_analise.dropna(subset=['trimestre_resposta'])
agrupado = df_analise.groupby(['trimestre_resposta', 'classificacao_nps']).size().unstack(fill_value=0)

# Calcular totais por trimestre
agrupado['total'] = agrupado.sum(axis=1)

# Calcular percentuais
for col in ['Promotor', 'Detrator', 'Neutro']:
    if col in agrupado.columns:
        agrupado[f'% {col}'] = (agrupado[col] / agrupado['total']) * 100
    else:
        agrupado[f'% {col}'] = 0.0

# Calcular NPS
agrupado['NPS'] = agrupado['% Promotor'] - agrupado['% Detrator']

# Arredondar
percentuais_cols = [col for col in agrupado.columns if col.startswith('%')]
agrupado[percentuais_cols] = agrupado[percentuais_cols].round(2)
agrupado['NPS'] = agrupado['NPS'].round(2)

# Exibir resumo no console
print("\nAnálise de NPS por Trimestre de Resposta:")
print(agrupado[['total', '% Promotor', '% Detrator', '% Neutro', 'NPS']])

# Salvar resultados em CSV
agrupado.to_csv('analise_nps_por_trimestre.csv')
print("\n✅ Resultados salvos em 'analise_nps_por_trimestre.csv'.")
