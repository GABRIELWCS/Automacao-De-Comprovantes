from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import time

# Configuração do Selenium
chrome_options = Options()
chrome_options.add_argument("--user-data-dir=C:\\Users\\Gabriel Souza\\AppData\\Local\\Google\\Chrome\\User Data")  # Caminho do perfil
servico = Service("C:/Users/Gabriel Souza/Desktop/Pyth.Vs/chromedriver.exe")  # Caminho do chromedriver

# Iniciar o navegador
navegador = webdriver.Chrome(service=servico, options=chrome_options)
navegador.get("https://web.whatsapp.com/")

# Aguardar o usuário escanear o QR Code
input("📲 Escaneie o QR Code e pressione Enter para continuar...")

# 📌 Função para extrair valor de comprovantes Pix
def extrair_pix(mensagem):
    padrao_valor = r"R\$\s?\d{1,3}(?:\.\d{3})*,\d{2}"  # Exemplo: R$ 150,00 ou R$ 1.250,75
    padrao_pix = r"(Pix recebido|Transferência via Pix|Pagamento Pix|Comprovante Pix)"

    if re.search(padrao_pix, mensagem, re.IGNORECASE) and re.search(padrao_valor, mensagem):
        valor = re.search(padrao_valor, mensagem).group()
        return valor
    return None

# Esperar até que as mensagens carreguem
try:
    WebDriverWait(navegador, 30).until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'copyable-text')]"))
    )
    print("📥 Mensagens carregadas!")
except Exception as e:
    print(f"❌ Erro ao carregar as mensagens: {e}")

# 📌 Armazena mensagens já processadas
mensagens_processadas = set()

# Loop para capturar mensagens em tempo real
while True:
    try:
        # Captura as mensagens mais recentes
        mensagens = navegador.find_elements(By.XPATH, "//div[contains(@class, 'copyable-text')]")

        for msg in mensagens:
            texto_msg = msg.text.strip()  # Remove espaços extras

            if texto_msg not in mensagens_processadas:  # Verifica se já foi processada
                valor_extraido = extrair_pix(texto_msg)  

                if valor_extraido:
                    print(f"✅ Comprovante Pix detectado! Valor: {valor_extraido}")
                    print(f"📝 Mensagem completa: {texto_msg}\n")
                    mensagens_processadas.add(texto_msg)  # Adiciona ao conjunto para não repetir

        time.sleep(5)  # Aguarda 5 segundos antes de verificar novas mensagens

    except KeyboardInterrupt:
        print("🛑 Saindo do script...")
        navegador.quit()  # Fecha o navegador corretamente
        break  # Sai do loop



