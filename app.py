import csv
import time
import os
import json
import requests
import cv2
import pytesseract as ts
import re
import openpyxl
import threading
import sys
import select
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Variável global para rastrear a última mensagem processada
ultima_mensagem_processada = None

# ----- CONFIGURAÇÃO INICIAL -----
def carregar_configuracoes():
    """Carrega as configurações do arquivo JSON."""
    with open('config.json') as config_file:
        return json.load(config_file)

# Substituindo o caminho do chromedriver diretamente no código
chromedriver_path = "C:/Users/Gabriel Souza/Desktop/clonegit/Sistema-Hamburgueria-/chromedriver.exe"  # Caminho do chromedriver
chrome_profile_path = "C:/Users/Gabriel Souza/AppData/Local/Google/Chrome/User Data"  # Caminho do perfil do Chrome
base_folder_path = "C:/Users/Gabriel Souza/Desktop/clonegit/Sistema-Hamburgueria-"  # Caminho base da pasta
dowloands_folder_path = "C:/Users/Gabriel Souza/Downloads"  # Caminho da pasta de downloads
tesseract_path = "C:/Program Files/Tesseract-OCR/tesseract.exe"  # Caminho do Tesseract

# Configurar o Tesseract
ts.pytesseract.tesseract_cmd = tesseract_path

def limpar_pasta(pasta):
    """Remove todos os arquivos da pasta especificada."""
    if os.path.exists(pasta):
        for arquivo in os.listdir(pasta):
            caminho_arquivo = os.path.join(pasta, arquivo)
            try:
                if os.path.isfile(caminho_arquivo):
                    os.remove(caminho_arquivo)
            except Exception as e:
                print(f"⚠️ Erro ao remover {arquivo}: {e}")

# Limpar a pasta "Comprovantes" antes de iniciar
limpar_pasta("Comprovantes")
print("🧹 Todos os arquivos da pasta 'Comprovantes' foram removidos!")

# Configuração do Selenium
def configurar_selenium():
    """Configura e inicia o navegador Selenium."""
    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir={chrome_profile_path}")  # Usa o perfil do Chrome
    # Remover o modo headless para garantir que a interface gráfica apareça
    # chrome_options.add_argument("--headless")  # Modo headless (sem interface gráfica)
    servico = Service(chromedriver_path)
    navegador = webdriver.Chrome(service=servico, options=chrome_options)
    navegador.get("https://web.whatsapp.com/")  # abre o wpp web
    return navegador

# Esperar o carregamento do QR Code e verificar se o WhatsApp Web está carregado
def esperar_carregamento_entrada():
    """Esperar até o QR Code ser escaneado ou o WhatsApp Web carregar completamente."""
    while True:
        try:
            # Esperar até o QR Code ser escaneado (essa div aparece enquanto o QR Code é mostrado)
            WebDriverWait(navegador, 60).until(
                EC.presence_of_element_located((By.XPATH, '//div[@title="Pesquisar ou iniciar nova conversa"]'))
            )
            print("✅ WhatsApp Web carregado com sucesso! Pronto para começar.")
            break  # O WhatsApp Web carregou completamente
        except Exception as e:
            print(f"🟠 Aguardando o QR Code ser escaneado... Erro: {e}")
            time.sleep(5)  # Aguardar um pouco e tentar novamente

navegador = configurar_selenium()
esperar_carregamento_entrada()

# Diretórios e arquivos de saída
os.makedirs(base_folder_path, exist_ok=True)

csv_funcionario = os.path.join(base_folder_path, "comprovantes_funcionario.csv")
csv_motoboy = os.path.join(base_folder_path, "comprovantes_motoboy.csv")

def criar_arquivos_csv():
    """Cria arquivos CSV com cabeçalho, se não existirem."""
    for arquivo in [csv_funcionario, csv_motoboy]:
        if not os.path.exists(arquivo):
            with open(arquivo, mode="a", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(["Nome", "Horário", "Valor", "Destinatário", "Categoria"])

criar_arquivos_csv()

# Conjunto para mensagens processadas (para evitar duplicatas)
mensagens_processadas = set()

# ----- FUNÇÕES AUXILIARES -----
def classificar_categoria(mensagem):
    """Retorna a categoria se a mensagem começar com '📍 transferência', senão None."""
    if mensagem.startswith("📍 transferência"):
        mensagem_lower = mensagem.lower()
        if "motoboy" in mensagem_lower:
            return "Motoboy"
        elif "funcionário" in mensagem_lower:
            return "Funcionário"
        else:
            return "Outros"
    return None

def gerar_nome_arquivo(nome, horario):
    """Gera um nome de arquivo simples removendo caracteres especiais."""
    nome_limpo = re.sub(r'[^\w]', '', nome)[:12]  # Limita para 12 caracteres
    horario_limpo = re.sub(r'[^\d]', '', horario)
    return f"{nome_limpo}_{horario_limpo}.jpg".strip()

def extrair_mensagens():
    """Extrai mensagens e imagens do WhatsApp que estejam começando com '📍 transferência' 
       e que foram enviadas após o início da execução."""
    global ultima_mensagem_processada
    mensagens_extraidas = []
    
    if ultima_mensagem_processada is None:
        ultima_mensagem_processada = datetime.now()
    
    # Localizar todas as mensagens no WhatsApp Web
    bolhas_mensagens = navegador.find_elements(By.XPATH, '//div[contains(@class, "message-in")]')
    
    for bolha in bolhas_mensagens:
        try:
            nome = "Você"
            horario = "Desconhecido"
            data_mensagem = None
            classe_msg = bolha.get_attribute("class")
            
            if "message-in" in classe_msg:
                try:
                    nome_elemento = bolha.find_element(By.XPATH, './/div[contains(@class, "copyable-text")]')
                    nome_completo = nome_elemento.get_attribute("data-pre-plain-text")
                    if "[" in nome_completo and "]" in nome_completo:
                        horario_texto = nome_completo.split("]")[0].replace("[", "").strip()
                        nome = nome_completo.split("] ")[-1].strip()
                        try:
                            data_mensagem = datetime.strptime(horario_texto, "%H:%M, %d/%m/%Y")
                            horario = data_mensagem.strftime("%H:%M")
                        except ValueError:
                            print(f"⚠️ Formato de data não reconhecido: {horario_texto}")
                            continue
                except Exception as e:
                    print(f"⚠️ Erro ao processar mensagem recebida: {e}")
                    continue
            elif "message-out" in classe_msg:
                data_mensagem = datetime.now()
                horario = data_mensagem.strftime("%H:%M")
                nome = "Você"
            
            texto_elementos = bolha.find_elements(By.XPATH, './/span[contains(@class, "selectable-text")]')
            texto = " ".join([t.text for t in texto_elementos])
            
            if not texto.startswith("📍 transferência"):
                continue
            
            if data_mensagem and data_mensagem >= ultima_mensagem_processada:
                try:
                    imagem_elemento = bolha.find_element(By.XPATH, './/img[contains(@src, "blob:") or contains(@class, "media")]')
                    imagem_url = imagem_elemento.get_attribute("src")
                except Exception:
                    imagem_elemento = None
                    imagem_url = None
                
                categoria = classificar_categoria(texto)
                if categoria:
                    identificador = f"{nome}-{horario}-{texto}"
                    if identificador not in mensagens_processadas:
                        mensagens_processadas.add(identificador)
                        print(f"📩 Mensagem de {nome} às {horario}\n {texto}")
                        ultima_mensagem_processada = datetime.now()
                        mensagens_extraidas.append((nome, horario, categoria, imagem_url, imagem_elemento))
        except Exception as e:
            print(f"⚠️ Erro ao extrair mensagem: {e}")
    
    return mensagens_extraidas

def baixar_imagem(imagem_elemento, nome_arquivo):
    """Faz o download da imagem, usando JavaScript se for URL blob ou requests."""
    try:
        imagem_url = imagem_elemento.get_attribute("src")
        if imagem_url.startswith("blob:"):
            script = f"""
            var xhr = new XMLHttpRequest();
            xhr.open('GET', '{imagem_url}', true);
            xhr.responseType = 'blob';
            xhr.onload = function() {{
                var blob = xhr.response;
                var reader = new FileReader();
                reader.onloadend = function() {{
                    var base64data = reader.result;
                    var link = document.createElement('a');
                    link.href = base64data;
                    link.download = '{nome_arquivo}';
                    link.click();
                }}; 
                reader.readAsDataURL(blob);
            }}; 
            xhr.send();
            """
            navegador.execute_script(script)
            print(f"✅ Imagem {nome_arquivo} baixada com sucesso!")
        else:
            resposta = requests.get(imagem_url)
            with open(nome_arquivo, 'wb') as file:
                file.write(resposta.content)
            print(f"✅ Imagem {nome_arquivo} baixada com sucesso!")
    except Exception as e:
        print(f"⚠️ Erro ao baixar imagem: {e}")

# Função principal de execução
def main():
    while True:
        print("🔄 Procurando novas mensagens...")
        mensagens_extraidas = extrair_mensagens()
        if not mensagens_extraidas:
            print("🔴 Nenhuma nova mensagem encontrada.")
        for nome, horario, categoria, imagem_url, imagem_elemento in mensagens_extraidas:
            nome_arquivo = gerar_nome_arquivo(nome, horario)
            if imagem_elemento:
                baixar_imagem(imagem_elemento, nome_arquivo)
        time.sleep(10)  # Aguardar 10 segundos antes de buscar novamente

# Rodar o script
if __name__ == "__main__":
    main()
