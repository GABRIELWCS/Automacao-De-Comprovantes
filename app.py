import csv
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



# Configuração do Selenium
chrome_options = Options()
chrome_options.add_argument("--user-data-dir=C:\\Users\\Gabriel Souza\\AppData\\Local\\Google\\Chrome\\User Data")  # Caminho do perfil do Chrome
servico = Service("C:/Users/Gabriel Souza/Desktop/Pyth.Vs/chromedriver.exe")  # Caminho do chromedriver



# Iniciar o navegador
navegador = webdriver.Chrome(service=servico, options=chrome_options)
navegador.get("https://web.whatsapp.com/")



# Aguardar o usuário escanear o QR Code
input("📲 Escaneie o QR Code e pressione Enter para continuar...")



# Criar diretório para armazenar os arquivos
pasta_base = "Comprovantes"
os.makedirs(pasta_base, exist_ok=True)

# Arquivos CSV dentro da pasta "Comprovantes"
csv_funcionario = os.path.join(pasta_base, "comprovantes_funcionario.csv")
csv_motoboy = os.path.join(pasta_base, "comprovantes_motoboy.csv")

# Criar arquivos CSV se não existirem
for arquivo in [csv_funcionario, csv_motoboy]:
    if not os.path.exists(arquivo):  # Evita sobrescrever dados existentes
        with open(arquivo, mode="a", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["Nome", "Horário", "Valor", "Mensagem Completa"])  # Cabeçalho atualizado

# Criar um conjunto para armazenar mensagens já processadas
mensagens_processadas = set()



# Função para classificar mensagens por categoria
def classificar_categoria(mensagem):
    if "motoboy" in mensagem.lower():
        return "Motoboy"
    elif "funcionário" in mensagem.lower():
        return "Funcionário"
    else:
        return "Outros"
    


# Função para extrair mensagens com nome e horário do remetente
def extrair_mensagens():
    mensagens_extraidas = []



    
    # Localizar todas as bolhas de mensagens no chat
    bolhas_mensagens = navegador.find_elements(By.XPATH, '//div[contains(@class, "message-in") or contains(@class, "message-out")]')

    for bolha in bolhas_mensagens:
        try:


            # Capturar o nome do remetente (caso seja uma mensagem recebida)
            nome = "Você"  # Para mensagens enviadas pelo usuário
            horario = "Desconhecido"  # Valor padrão caso não seja encontrado

            
            if "message-in" in bolha.get_attribute("class"):  # Se for uma mensagem recebida
                nome_elemento = bolha.find_element(By.XPATH, './/div[contains(@class, "copyable-text")]')
                nome = nome_elemento.get_attribute("data-pre-plain-text")  # ✅ Pegando corretamente



                # Extraindo o horário da mensagem
                if "[" in nome and "]" in nome:
                    horario = nome.split("]")[0].replace("[", "").strip()  # Obtendo apenas a parte do horário
                    nome = nome.split("] ")[-1].strip()  # Removendo o horário do nome



            # Capturar o texto da mensagem
            texto_elemento = bolha.find_elements(By.XPATH, './/span[contains(@class, "selectable-text")]')
            
            # Corrigir: Verificar se encontramos o texto corretamente
            texto = " ".join([t.text for t in texto_elemento])  # Combina textos longos
            
            if "Pix" in texto or "R$" in texto:  # Filtra mensagens que podem ser comprovantes
                identificador = f"{nome}-{horario}-{texto}"  # Criar um identificador único
                
                if identificador not in mensagens_processadas:  # Evitar repetição
                    mensagens_processadas.add(identificador)
                    categoria = classificar_categoria(texto)
                    
                    print(f"📩 Mensagem de {nome} às {horario} \n {texto}")
                    
                    mensagens_extraidas.append((nome, horario, categoria, texto))
        
        except Exception as e:
            print(f"⚠️ Erro ao extrair mensagem: {e}")

    return mensagens_extraidas





# Loop infinito para monitorar mensagens novas
while True:
    mensagens = extrair_mensagens()
    
    for nome, horario, categoria, mensagem in mensagens:
        if categoria == "Funcionário":
            arquivo = csv_funcionario
        elif categoria == "Motoboy":
            arquivo = csv_motoboy
        else:
            continue  # Ignora categorias irrelevantes
        
        # Salvar no arquivo CSV correspondente
        with open(arquivo, mode="a", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow([nome, horario, "", mensagem])  # Agora inclui o horário da mensagem
    
    print("⏳ Aguardando novas mensagens...")
    time.sleep(45)  # Espera 45 segundos antes de verificar novamente



######ATUALIZADO 19/03/2025 13:56#####