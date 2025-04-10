<h1 align="center">📲 Automação de Comprovantes WhatsApp com Python</h1>

<p align="center">
  Extraia comprovantes Pix diretamente de mensagens do WhatsApp Web em tempo real! 💸<br>
  Classifique por categoria, salve os dados em Excel e mantenha tudo organizado automaticamente.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10-blue?logo=python">
  <img src="https://img.shields.io/badge/Selenium-Automation-brightgreen?logo=selenium">
  <img src="https://img.shields.io/badge/OCR-Tesseract-orange?logo=google">
  <img src="https://img.shields.io/badge/Status-Em%20Desenvolvimento-yellow">
</p>

---

## ✨ Funcionalidades

- ✅ Acessa o WhatsApp Web com Selenium
- ✅ Lê mensagens em tempo real
- ✅ Identifica e baixa imagens de comprovantes Pix
- ✅ Usa OCR para extrair informações como:
  - Valor
  - Nome do destinatário
  - Categoria (Motoboy ou Funcionário)
- ✅ Gera planilhas Excel organizadas por **mês** e **categoria**
- ✅ Evita duplicações com verificação por horário

---

## 📸 Demonstração

<p align="center">
  <img src="https://github.com/seuusuario/seurepositorio/blob/main/imagens/demo.gif" alt="Demonstração da Automação" width="80%">
</p>

---

## 🚀 Tecnologias Utilizadas

- [Python 3.10+](https://www.python.org/)
- [Selenium](https://www.selenium.dev/)
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
- [OpenCV](https://opencv.org/)
- [Pandas + ExcelWriter](https://pandas.pydata.org/)

---

## ⚙️ Como usar

```bash
# Clone o repositório
git clone https://github.com/seuusuario/seurepositorio.git
cd seurepositorio

# Instale as dependências
pip install -r requirements.txt

# Execute o script principal
python main.py

⚠️ É necessário ter o Chrome instalado e configurado com perfil local.
⚠️ Tesseract OCR também precisa estar instalado e acessível no PATH.
