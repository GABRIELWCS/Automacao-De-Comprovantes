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

# Variável global para rastrear a última mensagem processada
ultima_mensagem_processada = None

# ----- CONFIGURAÇÃO INICIAL -----
def carregar_configuracoes():
    """Carrega as configurações do arquivo JSON."""
    with open('config.json') as config_file:
        return json.load(config_file)

config = carregar_configuracoes()
chromedriver_path = config['chromedriver_path']
chrome_profile_path = config['chrome_profile_path']
base_folder_path = config['base_folder_path']
dowloands_folder_path = config['dowloands_folder_path']
tesseract_path = config['tesseract_path']

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
    chrome_options.add_argument(f"--user-data-dir={chrome_profile_path}")
    servico = Service(chromedriver_path)
    navegador = webdriver.Chrome(service=servico, options=chrome_options)
    navegador.get("https://web.whatsapp.com/")
    return navegador

navegador = configurar_selenium()
input("📲 Escaneie o QR Code e pressione Enter para continuar...")

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
    
    bolhas_mensagens = navegador.find_elements(By.XPATH, '//div[contains(@class, "message-in") or contains(@class, "message-out")]')
    
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
                            data_mensagem = datetime.strptime(horario_texto, "%I:%M %p, %d/%m/%Y")
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
            print(f"✅ Imagem salva como: {nome_arquivo}")
        else:
            resposta = requests.get(imagem_url)
            if resposta.status_code == 200:
                with open(nome_arquivo, 'wb') as f:
                    f.write(resposta.content)
                print(f"✅ Imagem salva como: {nome_arquivo}")
            else:
                print(f"⚠️ Erro ao baixar a imagem: Status {resposta.status_code}")
    except Exception as e:
        print(f"⚠️ Erro ao baixar a imagem: {e}")

def analisar_imagem(nome_imagem, pasta_downloads=dowloands_folder_path, tentativas_max=10, intervalo_espera=1):
    """Tenta localizar e analisar a imagem na pasta de downloads, com múltiplas tentativas."""
    for tentativa in range(tentativas_max):
        try:
            caminho_imagem = os.path.join(pasta_downloads, nome_imagem)
            arquivos_na_pasta = os.listdir(pasta_downloads)
            
            arquivo_encontrado = None
            for arquivo in arquivos_na_pasta:
                if (arquivo.lower().replace('ç', 'c') == nome_imagem.lower().replace('ç', 'c')):
                    arquivo_encontrado = os.path.join(pasta_downloads, arquivo)
                    break
            
            if not arquivo_encontrado:
                print(f"⚠️ Imagem não encontrada na tentativa {tentativa + 1}. Aguardando...")
                time.sleep(intervalo_espera)
                continue

            caminho_imagem = arquivo_encontrado
            print(f"Arquivo encontrado: {caminho_imagem}")

            if os.path.getsize(caminho_imagem) == 0:
                print(f"⚠️ Arquivo vazio na tentativa {tentativa + 1}. Aguardando...")
                time.sleep(intervalo_espera)
                continue

            img = cv2.imread(caminho_imagem)
            if img is None:
                print(f"⚠️ Erro ao carregar a imagem na tentativa {tentativa + 1}. Aguardando...")
                time.sleep(intervalo_espera)
                continue

            text_img = ts.image_to_string(img, lang='por')
            valor = re.findall(r'R\$\s*\d+[\.,]?\d*', text_img)
            destinatario = re.findall(r'(?i)(?:para|nome do favorecido)\s*:?\s*([^\n]+)', text_img)

            if valor:
                valor = valor[0].strip()
            else:
                valor = "Valor não encontrado"

            if destinatario:
                ditemp = destinatario[0].strip().replace('\n', ' ')
                destinatario = ' '.join(ditemp.split())
            else:
                destinatario = "Destinatário não encontrado"

            return valor, destinatario

        except Exception as e:
            print(f"Erro na tentativa {tentativa + 1}: {e}")
            time.sleep(intervalo_espera)

    print(f"⚠️ Falha ao processar a imagem {nome_imagem} após {tentativas_max} tentativas.")
    return None, None

def verificar_entrada():
    """Verifica se foi digitado 'e' no terminal para encerrar sem bloquear a execução."""
    if sys.stdin in select.select([sys.stdin], [], [], 0)[0]:
        linha = sys.stdin.readline().strip().lower()
        if linha == 'e':
            return True
    return False

def converter_csv_para_excel(caminho_csv, caminho_excel):
    """Converte um arquivo CSV para Excel."""
    try:
        with open(caminho_csv, 'r', encoding='utf-8') as file_csv:
            csv_reader = csv.reader(file_csv)
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            for linha in csv_reader:
                sheet.append(linha)
            for coluna in sheet.columns:
                max_length = 0
                coluna_letra = coluna[0].column_letter
                for celula in coluna:
                    try:
                        max_length = max(max_length, len(str(celula.value)))
                    except:
                        pass
                sheet.column_dimensions[coluna_letra].width = max_length + 2
            workbook.save(caminho_excel)
            print(f"✅ Arquivo Excel gerado: {caminho_excel}")
    except Exception as e:
        print(f"⚠️ Erro ao converter CSV para Excel: {e}")

def thread_verificacao_entrada():
    """Thread para verificar entrada do usuário e encerrar o programa."""
    global continuar_execucao
    while continuar_execucao:
        if verificar_entrada():
            continuar_execucao = False
            print("\n🛑 Encerrando o programa...")
            break
        time.sleep(0.5)

# Variável global para controle da execução
continuar_execucao = True
thread_entrada = threading.Thread(target=thread_verificacao_entrada)
thread_entrada.daemon = True
thread_entrada.start()

print("🚀 Programa iniciado. Digite 'e' a qualquer momento para encerrar.")

# ----- LOOP PRINCIPAL -----
try:
    while continuar_execucao:
        mensagens = extrair_mensagens()
        if not mensagens:
            print("⏳ Nenhuma nova mensagem encontrada. Aguardando...")

        for nome, horario, categoria, imagem_url, imagem_elemento in mensagens:
            if not continuar_execucao:
                break

            if categoria == "Funcionário":
                arquivo = csv_funcionario
            elif categoria == "Motoboy":
                arquivo = csv_motoboy
            else:
                continue

            print(f"💾 Salvando no arquivo: {arquivo}")
            try:
                with open(arquivo, mode="a", newline="", encoding="utf-8") as file:
                    writer = csv.writer(file)
                    if imagem_elemento:
                        nome_imagem = gerar_nome_arquivo(nome, horario)
                        baixar_imagem(imagem_elemento, nome_imagem)
                        valor, destinatario = analisar_imagem(nome_imagem)
                        print(f"Valor: {valor}, Destinatário: {destinatario}")
                        writer.writerow([nome, horario, valor, destinatario, categoria])
                print("✅ Mensagem salva")
            except Exception as e:
                print(f"⚠️ Erro ao salvar no CSV: {e}")

        time.sleep(3)

except KeyboardInterrupt:
    print("\n🛑 Interrupção do usuário detectada.")
    continuar_execucao = False

finally:
    if os.path.exists(csv_funcionario):
        converter_csv_para_excel(csv_funcionario, os.path.join(base_folder_path, "comprovantes_funcionario.xlsx"))
    if os.path.exists(csv_motoboy):
        converter_csv_para_excel(csv_motoboy, os.path.join(base_folder_path, "comprovantes_motoboy.xlsx"))
    print("📊 Conversão de CSVs para Excel concluída.")
    thread_entrada.join()