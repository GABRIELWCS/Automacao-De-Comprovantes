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

ultima_mensagem_processada = None

# Ler configurações do arquivo JSON
with open('config.json') as config_file:
    config = json.load(config_file)

chromedriver_path = config['chromedriver_path']
chrome_profile_path = config['chrome_profile_path']

# Código para limpar pasta de registro
def limpar_pasta(pasta):
    if os.path.exists(pasta):
        for arquivo in os.listdir(pasta):
            caminho_arquivo = os.path.join(pasta, arquivo)
            try:
                if os.path.isfile(caminho_arquivo):
                    os.remove(caminho_arquivo)  # Remove arquivos
            except Exception as e:
                print(f"⚠️ Erro ao remover {arquivo}: {e}")

# Exemplo de uso antes de salvar novos arquivos
limpar_pasta("Comprovantes")
print("🧹 Todos os arquivos da pasta 'Comprovantes' foram removidos!")

# Configuração do Selenium
chrome_options = Options()
chrome_options.add_argument(f"--user-data-dir={chrome_profile_path}")
servico = Service(chromedriver_path)

# Iniciar o navegador
navegador = webdriver.Chrome(service=servico, options=chrome_options)
navegador.get("https://web.whatsapp.com/")

# Aguardar o usuário escanear o QR Code
input("📲 Escaneie o QR Code e pressione Enter para continuar...")

# Criar diretório para armazenar os arquivos
pasta_base = "/home/matheus/Documents/Github/Sistema-Hamburgueria-/Comprovantes"
os.makedirs(pasta_base, exist_ok=True)

csv_funcionario = os.path.join(pasta_base, "comprovantes_funcionario.csv")
csv_motoboy = os.path.join(pasta_base, "comprovantes_motoboy.csv")

# Criar arquivos CSV se não existirem.
for arquivo in [csv_funcionario, csv_motoboy]:
    if not os.path.exists(arquivo):
        with open(arquivo, mode="a", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["Nome", "Horário", "Valor", "Destinatário", "Categoria"])

# Criar um conjunto para armazenar mensagens já processadas
mensagens_processadas = set()

# Função para classificar mensagens por categoria
def classificar_categoria(mensagem):
    mensagem_lower = mensagem.lower()
    
    # Palavras-chave para identificar comprovantes
    palavras_comprovante = [
        "comprovante pix", 
        "pix", 
        "comprovante de pagamento", 
        "comprovante", 
        "pagamento"
    ]
    
    # Verifica se a mensagem contém alguma palavra-chave de comprovante
    if any(palavra in mensagem_lower for palavra in palavras_comprovante):
        if "motoboy" in mensagem_lower:
            return "Motoboy"
        elif "funcionário" in mensagem_lower:
            return "Funcionário"
        else:
            return "Outros"
    
    return None

def extrair_mensagens():
    global ultima_mensagem_processada
    mensagens_extraidas = []

    # Se não houver registro da última mensagem, defina um limite de tempo (por exemplo, últimas 24 horas)
    if ultima_mensagem_processada is None:
        ultima_mensagem_processada = datetime.now() - timedelta(hours=0.1)

    # Localizar todas as bolhas de mensagens no chat
    bolhas_mensagens = navegador.find_elements(By.XPATH, '//div[contains(@class, "message-in") or contains(@class, "message-out")]')

    for bolha in bolhas_mensagens:
        try:
            nome = "Você"  # Para mensagens enviadas pelo usuário
            horario = "Desconhecido"  # Valor padrão caso não seja encontrado
            data_mensagem = None

            if "message-in" in bolha.get_attribute("class"):  # Se for uma mensagem recebida
                nome_elemento = bolha.find_element(By.XPATH, './/div[contains(@class, "copyable-text")]')
                nome = nome_elemento.get_attribute("data-pre-plain-text")

                # Extraindo o horário da mensagem
                if "[" in nome and "]" in nome:
                    horario_texto = nome.split("]")[0].replace("[", "").strip()
                    nome = nome.split("] ")[-1].strip()

                    try:
                        # Parsear o formato "H:MM AM/PM, dd/mm/yyyy"
                        data_mensagem = datetime.strptime(horario_texto, "%I:%M %p, %d/%m/%Y")
                        horario = data_mensagem.strftime("%H:%M")  # Formato 24h para consistência
                    except ValueError:
                        print(f"⚠️ Formato de data não reconhecido: {horario_texto}")
                        continue

            # Capturar o texto da mensagem
            texto_elemento = bolha.find_elements(By.XPATH, './/span[contains(@class, "selectable-text")]')
            texto = " ".join([t.text for t in texto_elemento])

            # Verificar se a mensagem é nova
            if data_mensagem and data_mensagem > ultima_mensagem_processada:
                # Verificar se há uma imagem anexada
                try:
                    imagem_elemento = bolha.find_element(By.XPATH, './/img[contains(@src, "blob:") or contains(@class, "media")]')
                    imagem_url = imagem_elemento.get_attribute("src")  # Extrai a URL da imagem
                except Exception:
                    imagem_elemento = None
                    imagem_url = None

                # Classificar a categoria
                categoria = classificar_categoria(texto)
                
                # Verificar se é um comprovante válido
                if categoria:
                    identificador = f"{nome}-{horario}-{texto}"

                    if identificador not in mensagens_processadas:
                        mensagens_processadas.add(identificador)
                        print(f"📩 Mensagem de {nome} às {horario} \n {texto}")
                        
                        # Atualizar a última mensagem processada
                        ultima_mensagem_processada = data_mensagem
                        
                        mensagens_extraidas.append((nome, horario, categoria, imagem_url, imagem_elemento))
        
        except Exception as e:
            print(f"⚠️ Erro ao extrair mensagem: {e}")

    return mensagens_extraidas

# Função para fazer o download da imagem
def baixar_imagem(imagem_elemento, nome_arquivo):
    try:
        # Obter a URL da imagem
        imagem_url = imagem_elemento.get_attribute("src")
        
        # Verificar se a URL é um blob
        if imagem_url.startswith("blob:"):
            # Usar JavaScript para obter os dados da imagem
            script = f"""
            var xhr = new XMLHttpRequest();
            xhr.open('GET', '{imagem_url}', true);
            xhr.responseType = 'blob';
            xhr.onload = function() {{
                var blob = xhr.response;
                var reader = new FileReader();
                reader.onloadend = function() {{
                    var base64data = reader.result;
                    // Criar um link para download
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
            # Se não for um blob, use requests como antes
            resposta = requests.get(imagem_url)
            if resposta.status_code == 200:
                with open(nome_arquivo, 'wb') as f:
                    f.write(resposta.content)
                print(f"✅ Imagem salva como: {nome_arquivo}")
            else:
                print(f"⚠️ Erro ao baixar a imagem: Status {resposta.status_code}")
    except Exception as e:
        print(f"⚠️ Erro ao baixar a imagem: {e}")

# Configurar o caminho do Tesseract
ts.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

import time
import os

def analisar_imagem(nome_imagem, pasta_downloads="/home/matheus/Downloads", tentativas_max=10, intervalo_espera=1):
    for tentativa in range(tentativas_max):
        try:
            # Constrói o caminho completo para a imagem na pasta de downloads
            caminho_imagem = os.path.join(pasta_downloads, nome_imagem)
            
            # Listar arquivos na pasta de downloads
            arquivos_na_pasta = os.listdir(pasta_downloads)
            
            # Tenta encontrar o arquivo com correspondência flexível
            arquivo_encontrado = None
            for arquivo in arquivos_na_pasta:
                # Comparação mais flexível, ignorando diferenças de acentuação e case
                if (arquivo.lower().replace('ç', 'c') == 
                    nome_imagem.lower().replace('ç', 'c')):
                    arquivo_encontrado = os.path.join(pasta_downloads, arquivo)
                    break
            
            # Se não encontrou o arquivo, espera e tenta novamente
            if not arquivo_encontrado:
                print(f"⚠️ Imagem não encontrada na tentativa {tentativa + 1}. Aguardando...")
                time.sleep(intervalo_espera)
                continue

            caminho_imagem = arquivo_encontrado
            print(f"Arquivo encontrado: {caminho_imagem}")

            # Verificar se o arquivo tem um tamanho válido
            if os.path.getsize(caminho_imagem) == 0:
                print(f"⚠️ Arquivo vazio na tentativa {tentativa + 1}. Aguardando...")
                time.sleep(intervalo_espera)
                continue

            # Ler a imagem
            img = cv2.imread(caminho_imagem)

            # Verificar se a imagem foi carregada corretamente
            if img is None:
                print(f"⚠️ Erro ao carregar a imagem na tentativa {tentativa + 1}. Aguardando...")
                time.sleep(intervalo_espera)
                continue

            # Usar Tesseract para extrair texto da imagem
            text_img = ts.image_to_string(img, lang='por')

            # Encontrar o valor e o destinatário usando expressões regulares
            valor = re.findall(r'R\$\s*\d+[\.,]?\d*', text_img)
            destinatario = re.findall(r'(?i)(?:para|nome do favorecido)\s*:?\s*([^\n]+)', text_img)

            # Limpar o valor encontrado
            if valor:
                valor_limpo = valor[0].strip()
                valor = valor_limpo
            else:
                valor = "Valor não encontrado"

            # Limpar o destinatário encontrado
            if destinatario:
                nome_limpo = destinatario[0].strip()
                nome_limpo = nome_limpo.replace('\n', ' ')
                nome_limpo = ' '.join(nome_limpo.split())
                destinatario = nome_limpo
            else:
                destinatario = "Destinatário não encontrado"

            return valor, destinatario

        except Exception as e:
            print(f"Erro na tentativa {tentativa + 1}: {e}")
            time.sleep(intervalo_espera)

    # Se todas as tentativas falharem
    print(f"⚠️ Falha ao processar a imagem {nome_imagem} após {tentativas_max} tentativas.")
    return None, None

def gerar_nome_arquivo(nome, horario):
    # Remove caracteres especiais do nome
    nome_limpo = re.sub(r'[^\w]', '', nome)

    # Remove caracteres especiais do horário
    horario_limpo = re.sub(r'[^\d]', '', horario)

    # Limita o comprimento do nome
    nome_limpo = nome_limpo[:12]  # Limita para 10 caracteres

    # Gera um nome de arquivo simples
    nome_arquivo = f"{nome_limpo}_{horario_limpo}.jpg"

    return nome_arquivo.strip()

# Função para verificar entrada do usuário
def verificar_entrada():
    while True:
        # Verifica se há entrada no terminal sem bloquear a execução
        if sys.stdin in select.select([sys.stdin], [], [], 0)[0]:
            linha = sys.stdin.readline().strip().lower()
            if linha == 'e':
                return True
        time.sleep(0.5)

# Função para converter CSV para Excel
def converter_csv_para_excel(caminho_csv, caminho_excel):
    try:
        # Ler o CSV
        with open(caminho_csv, 'r', encoding='utf-8') as file_csv:
            csv_reader = csv.reader(file_csv)
            
            # Criar um novo workbook do Excel
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            
            # Copiar dados do CSV para o Excel
            for linha in csv_reader:
                sheet.append(linha)
            
            # Ajustar largura das colunas
            for coluna in sheet.columns:
                max_length = 0
                coluna_letra = coluna[0].column_letter
                for celula in coluna:
                    try:
                        if len(str(celula.value)) > max_length:
                            max_length = len(celula.value)
                    except:
                        pass
                ajustado_width = (max_length + 2)
                sheet.column_dimensions[coluna_letra].width = ajustado_width
            
            # Salvar o workbook
            workbook.save(caminho_excel)
            print(f"✅ Arquivo Excel gerado: {caminho_excel}")
    except Exception as e:
        print(f"⚠️ Erro ao converter CSV para Excel: {e}")

# Iniciar thread de verificação de entrada
def thread_verificacao_entrada():
    global continuar_execucao
    while continuar_execucao:
        if verificar_entrada():
            continuar_execucao = False
            print("\n🛑 Encerrando o programa...")
            break

# Variável global para controle de execução
continuar_execucao = True

# Iniciar thread de verificação de entrada
thread_entrada = threading.Thread(target=thread_verificacao_entrada)
thread_entrada.daemon = True
thread_entrada.start()

# Mensagem inicial de instruções
print("🚀 Programa iniciado. Digite 'e' a qualquer momento para encerrar.")

# Loop principal
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
                continue  # Ignora categorias irrelevantes
            
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
                print(f"✅ Mensagem salva")
            except Exception as e:
                print(f"⚠️ Erro ao salvar no CSV: {e}")

        time.sleep(3)  # Espera 3 segundos antes de verificar novamente

except KeyboardInterrupt:
    print("\n🛑 Interrupção do usuário detectada.")
    continuar_execucao = False

finally:
    # Quando o programa for encerrado, converter CSVs para Excel
    if os.path.exists(csv_funcionario):
        converter_csv_para_excel(
            csv_funcionario, 
            os.path.join(pasta_base, "comprovantes_funcionario.xlsx")
        )
    
    if os.path.exists(csv_motoboy):
        converter_csv_para_excel(
            csv_motoboy, 
            os.path.join(pasta_base, "comprovantes_motoboy.xlsx")
        )
    
    print("📊 Conversão de CSVs para Excel concluída.")
    thread_entrada.join()