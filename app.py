from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import csv
import time
import os


# Código para limpar a pasta de registro
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

# Configuração do Selenium com o WebDriver Manager
chrome_options = Options()
chrome_options.add_argument("--user-data-dir=C:\\Users\\Gabriel Souza\\AppData\\Local\\Google\\Chrome\\User Data")  # Caminho do perfil do Chrome
service = Service(ChromeDriverManager().install())  # Usando o WebDriver Manager

# Iniciar o navegador
navegador = webdriver.Chrome(service=service, options=chrome_options)
navegador.get("https://web.whatsapp.com/")

# Aguardar o usuário escanear o QR Code
input("📲 Escaneie o QR Code e pressione Enter para continuar...")

# Criar diretório para armazenar os arquivos
pasta_base = "Comprovantes"
os.makedirs(pasta_base, exist_ok=True)

# Caminho dos arquivos
csv_funcionario = os.path.join(pasta_base, "comprovantes_funcionario.csv")
csv_motoboy = os.path.join(pasta_base, "comprovantes_motoboy.csv")

# Criar arquivos CSV se não existirem
for arquivo in [csv_funcionario, csv_motoboy]:
    if not os.path.exists(arquivo):  # Evita sobrescrever dados existentes
        with open(arquivo, mode="a", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["Nome", "Horário", "Valor", "Mensagem Completa"])

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
    bolhas_mensagens = navegador.find_elements(By.XPATH, '//div[contains(@class, "message-in") or contains(@class, "message-out")]')

    for bolha in bolhas_mensagens:
        try:
            nome = "Você"
            horario = "Desconhecido"

            if "message-in" in bolha.get_attribute("class"):  # Se for uma mensagem recebida
                nome_elemento = bolha.find_element(By.XPATH, './/div[contains(@class, "copyable-text")]')
                nome = nome_elemento.get_attribute("data-pre-plain-text")

                # Extraindo o horário da mensagem
                if "[" in nome and "]" in nome:
                    horario = nome.split("]")[0].replace("[", "").strip()
                    nome = nome.split("] ")[-1].strip()

            # Capturar o texto da mensagem
            texto_elemento = bolha.find_elements(By.XPATH, './/span[contains(@class, "selectable-text")]')
            texto = " ".join([t.text for t in texto_elemento])  # Combina textos longos

            if "Pix" in texto or "R$" in texto:  # Filtra mensagens que podem ser comprovantes
                identificador = f"{nome}-{horario}-{texto}"

                if identificador not in mensagens_processadas:
                    mensagens_processadas.add(identificador)
                    categoria = classificar_categoria(texto)

                    print(f"📩 Mensagem de |{nome} às |{horario}|: {texto}|")

                    mensagens_extraidas.append((nome, horario, categoria, texto))

        except Exception as e:
            print(f"⚠️ Erro ao extrair mensagem: {e}")

    return mensagens_extraidas

# Loop infinito para monitorar mensagens novas
while True:
    mensagens = extrair_mensagens()

    if not mensagens:
        print("⏳ Nenhuma nova mensagem encontrada. Aguardando...")

    for nome, horario, categoria, mensagem in mensagens:
        if categoria == "Funcionário":
            arquivo = csv_funcionario
        elif categoria == "Motoboy":
            arquivo = csv_motoboy
        else:
            continue  # Ignora categorias irrelevantes

        print(f"💾 Salvando no arquivo: {arquivo}")

        # Salvar no arquivo CSV correspondente
        try:
            with open(arquivo, mode="a", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow([nome, horario, "", mensagem])
            print(f"✅ Mensagem salva: {mensagem}")
        except Exception as e:
            print(f"⚠️ Erro ao salvar no CSV: {e}")

    time.sleep(15)  # Espera 15 segundos antes de verificar novamente




#CÓDIGO BASE: FUNCIONAMENTO OKAY(EXTRAÇÃO DE TRANSFERENCIA COM DATA/HORA,NOME DE QUEM ENVIOU A MENSAGEM,Valor, Categoria)


