import pandas as pd

# Carrega a planilha
excel_file = pd.ExcelFile("Relatorio_Fiscalizacao_Completo.xlsx")

# Carrega as abas
fiscalizacao = excel_file.parse("Fiscalização")
atividades = excel_file.parse("Atividades")
infracoes = excel_file.parse("Infrações")

# Converte a coluna "Data da fiscalização" para datetime, ignorando erros
fiscalizacao["Data da fiscalização"] = pd.to_datetime(fiscalizacao["Data da fiscalização"], errors='coerce')

# Define o período de interesse
data_inicio = pd.to_datetime("01/01/2023")
data_fim = pd.to_datetime("31/12/2023")

# Filtra os dados da aba "Fiscalização"
filtro_fiscalizacao = (fiscalizacao["Município"] == "Belo Horizonte") & (fiscalizacao["Data da fiscalização"] >= data_inicio) & (fiscalizacao["Data da fiscalização"] <= data_fim)
fiscalizacao_filtrada = fiscalizacao[filtro_fiscalizacao]

# Cria DataFrames vazios para armazenar os resultados
atividades_resultado = pd.DataFrame()
infracoes_resultado = pd.DataFrame()

# Itera sobre os IDs filtrados
for id_fiscalizacao in fiscalizacao_filtrada["ID"]:
    # Filtra as atividades e infrações correspondentes
    atividades_correspondentes = atividades[atividades["ID Fiscalização"] == id_fiscalizacao]
    infracoes_correspondentes = infracoes[infracoes["ID Fiscalização"] == id_fiscalizacao]

    # Concatena os resultados nos DataFrames correspondentes
    atividades_resultado = pd.concat([atividades_resultado, atividades_correspondentes])
    infracoes_resultado = pd.concat([infracoes_resultado, infracoes_correspondentes])

# Remove IDs duplicados, se houver
fiscalizacao_filtrada = fiscalizacao_filtrada.drop_duplicates(subset="ID")

# Cria um novo arquivo Excel
with pd.ExcelWriter("Resultado_2.xlsx") as writer:
    # Escreve os DataFrames em abas separadas
    fiscalizacao_filtrada.to_excel(writer, sheet_name="Fiscalização", index=False)
    atividades_resultado.to_excel(writer, sheet_name="Atividades", index=False)
    infracoes_resultado.to_excel(writer, sheet_name="Infrações", index=False)