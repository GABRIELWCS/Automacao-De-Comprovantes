from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import re
from openpyxl import Workbook, load_workbook
import pandas as pd

# Função para limpar a pasta de registros antigos
def limpar_pasta(pasta):
    if os.path.exists(pasta):
        for arquivo in os.listdir(pasta):
            caminho_arquivo = os.path.join(pasta, arquivo)
            try:
                if os.path.isfile(caminho_arquivo):
                    os.remove(caminho_arquivo)
            except Exception as e:
                print(f"⚠️ Erro ao remover {arquivo}: {e}")

# Limpa a pasta antes de iniciar se necessário
limpar_pasta("Comprovantes")
print("🧹 Todos os arquivos da pasta 'Comprovantes' foram removidos!")

# Configuração do Selenium
chrome_options = Options()
chrome_options.add_argument("--user-data-dir=C:\\Users\\Gabriel Souza\\AppData\\Local\\Google\\Chrome\\User Data")
service = Service(ChromeDriverManager().install())

# Iniciar o navegador
navegador = webdriver.Chrome(service=service, options=chrome_options)
navegador.get("https://web.whatsapp.com/")

# Aguardar login do usuário
input("📲 Escaneie o QR Code e pressione Enter para continuar...")

# Criar diretório
pasta_base = "Comprovantes"
os.makedirs(pasta_base, exist_ok=True)

# Definição dos arquivos Excel
xlsx_funcionario = os.path.join(pasta_base, "comprovantes_funcionario.xlsx")
xlsx_motoboy = os.path.join(pasta_base, "comprovantes_motoboy.xlsx")

# Função para criar arquivos Excel caso não existam
def criar_arquivo_excel(arquivo):
    if not os.path.exists(arquivo):
        wb = Workbook()
        ws = wb.active
        ws.append(["Nome", "Horário", "Valor", "Mensagem Completa"])
        wb.save(arquivo)
        print(f"📊 Arquivo criado: {arquivo}")

# Criar arquivos Excel
criar_arquivo_excel(xlsx_funcionario)
criar_arquivo_excel(xlsx_motoboy)

# Função para ajustar a largura das colunas no Excel
def ajustar_largura_colunas(arquivo):
    wb = load_workbook(arquivo)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    wb.save(arquivo)

# Função para carregar mensagens já salvas no Excel
def carregar_mensagens_anteriores(arquivo):
    wb = load_workbook(arquivo)
    ws = wb.active
    mensagens_existentes = set()

    for row in ws.iter_rows(min_row=2, values_only=True):  # Pular a primeira linha de cabeçalho
        nome, horario, valor, mensagem = row
        identificador = f"{nome}-{horario}-{valor}-{mensagem}"
        mensagens_existentes.add(identificador)

    return mensagens_existentes

# Conjunto para armazenar mensagens já processadas
mensagens_processadas = set()

# Função para classificar categoria
def classificar_categoria(mensagem):
    if "motoboy" in mensagem.lower():
        return "Motoboy"
    elif "funcionario" or "funcionário" in mensagem.lower():
        return "Funcionário"
    return "Outros"

# Função para extrair valor do Pix
def extrair_valor(mensagem):
    padrao_valor = r'R\$\s*\d{1,3}(?:\.\d{3})*(?:,\d{2})?'
    valores = re.findall(padrao_valor, mensagem)
    return valores[0].replace(" ", "") if valores else "Não encontrado"

# Função para remover erros de codificação
def corrigir_acentuacao(texto):
    return texto.encode('utf-8').decode('utf-8-sig')

# Extrair mensagens
def extrair_mensagens():
    global mensagens_processadas
    novas_mensagens = []

    bolhas = navegador.find_elements(By.XPATH, '//div[contains(@class, "message-in") or contains(@class, "message-out")]')

    for bolha in bolhas:
        try:
            nome = "Você"
            horario = "Desconhecido"

            if "message-in" in bolha.get_attribute("class"):
                nome_elemento = bolha.find_element(By.XPATH, './/div[contains(@class, "copyable-text")]')
                nome = nome_elemento.get_attribute("data-pre-plain-text")

                if "[" in nome and "]" in nome:
                    horario = nome.split("]")[0].replace("[", "").strip()
                    nome = nome.split("] ")[-1].strip()

            texto_elemento = bolha.find_elements(By.XPATH, './/span[contains(@class, "selectable-text")]')
            texto = " ".join([t.text for t in texto_elemento]).strip()

            if "Pix" in texto or "R$" in texto:
                categoria = classificar_categoria(texto)
                valor = extrair_valor(texto)

                mensagem_limpa = corrigir_acentuacao(" ".join(texto.split()))

                identificador = f"{nome}-{horario}-{valor}-{mensagem_limpa}"

                # Verificar se a mensagem já foi processada
                if identificador not in mensagens_processadas:
                    mensagens_processadas.add(identificador)
                    novas_mensagens.append((nome, horario, categoria, valor, mensagem_limpa))

        except Exception as e:
            print(f"⚠️ Erro ao extrair mensagem: {e}")

    return novas_mensagens

# Monitoramento contínuo
while True:
    mensagens = extrair_mensagens()

    if not mensagens:
        print("⏳ Nenhuma nova mensagem encontrada. Aguardando...")

    for nome, horario, categoria, valor, mensagem in mensagens:
        if categoria == "Funcionário":
            arquivo_excel = xlsx_funcionario
        elif categoria == "Motoboy":
            arquivo_excel = xlsx_motoboy
        else:
            continue

        # Carregar mensagens já salvas para evitar duplicação
        mensagens_existentes = carregar_mensagens_anteriores(arquivo_excel)

        identificador = f"{nome}-{horario}-{valor}-{mensagem}"
        if identificador not in mensagens_existentes:
            print(f"💾 Salvando no arquivo: {arquivo_excel}")

            try:
                wb = load_workbook(arquivo_excel)
                ws = wb.active
                ws.append([nome, horario, valor, mensagem])
                wb.save(arquivo_excel)
                print(f"✅ Mensagem salva: {mensagem}")

                # Ajustar a largura das colunas após salvar
                ajustar_largura_colunas(arquivo_excel)

            except Exception as e:
                print(f"⚠️ Erro ao salvar no Excel: {e}")
        else:
            print("🛑 Mensagem já salva, ignorando.")

    time.sleep(35)