import re
import pandas as pd

# Caminho do arquivo exportado do WhatsApp
arquivo_whatsapp = "teste2.txt"

# Lista para armazenar os comprovantes extraídos
comprovantes = []

# Expressões regulares para identificar os dados
padrao_valor = r"Valor: R\$ ([\d,.]+)"
padrao_destinatario = r"Destinatário: (.+)"
padrao_categoria = r"Categoria: (.+)"

# Abrir e ler o arquivo de conversa
with open(arquivo_whatsapp, "r", encoding="utf-8") as arquivo:
    linhas = arquivo.readlines()

# Variáveis temporárias para armazenar os dados de cada comprovante
valor, destinatario, categoria = None, None, None

for linha in linhas:
    if "Transferência realizada via Pix" in linha:
        valor, destinatario, categoria = None, None, None  # Resetar os dados
        
    elif re.search(padrao_valor, linha):
        valor = re.search(padrao_valor, linha).group(1)
        
    elif re.search(padrao_destinatario, linha):
        destinatario = re.search(padrao_destinatario, linha).group(1)
        
    elif re.search(padrao_categoria, linha):
        categoria = re.search(padrao_categoria, linha).group(1)

    # Quando todas as informações forem encontradas, adicionamos à lista
    if valor and destinatario and categoria:
        comprovantes.append({"Valor": valor, "Destinatário": destinatario, "Categoria": categoria})
        valor, destinatario, categoria = None, None, None  # Resetar para o próximo comprovante

# Criar um DataFrame para visualizar melhor
df = pd.DataFrame(comprovantes)

# Salvar os dados em um CSV
df.to_csv("teste2.csv", index=False)

print("Extração concluída! Os dados foram salvos em 'teste2.csv'.")
