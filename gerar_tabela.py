import os
import pandas as pd

# Caminho da pasta onde estão os CSVs
pasta_base = "Comprovantes"

# Lista todos os arquivos CSV dentro da pasta
arquivos_csv = [f for f in os.listdir(pasta_base) if f.endswith(".csv")]

# Lista para armazenar os DataFrames
dataframes = []

# Ler cada CSV e adicionar à lista
for arquivo in arquivos_csv:
    caminho_arquivo = os.path.join(pasta_base, arquivo)
    df = pd.read_csv(caminho_arquivo)  # Lendo CSV
    df["Origem"] = arquivo  # Adiciona uma coluna indicando o arquivo de origem
    dataframes.append(df)

# Combinar todos os DataFrames em uma única tabela
if dataframes:
    tabela_final = pd.concat(dataframes, ignore_index=True)
    
    # Exibir tabela no terminal
    print(tabela_final)

    # Salvar como Excel (opcional)
    tabela_final.to_excel("Tabela_Comprovantes.xlsx", index=False)
    print("📊 Tabela salva como 'Tabela_Comprovantes.xlsx'")
else:
    print("⚠️ Nenhum arquivo CSV encontrado na pasta 'Comprovantes'.")

tabela_final.to_excel(r"C:\Users\Gabriel Souza\Desktop\Pyth.Vs\Tabela_Comprovantes.xlsx", index=False)


