import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException

# === CONFIGURAÇÕES ===

EXCEL_PATH = "CAMINHO BASE EXCEL"
FORM_URL = "LINK DO FORMULÁRIO"

CAMPOS = [
    ("Cargo", "dropdown"),
    ("Data Aprovação", "data"),
    ("Recrutador", "dropdown"),
    ("Colaborador", "texto"),
    ("Data Admissão", "data"),
    ("Frente", "dropdown"),
    ("COORDENAÇÃO", "dropdown"),
    ("Observação", "dropdown"),
]

# === FUNÇÕES ===

def selecionar_dropdown(pergunta, valor_desejado):
    try:
        # Abre o dropdown
        dropdown = pergunta.find_element(By.CSS_SELECTOR, "div[role='button']")
        dropdown.click()
        time.sleep(0.5)

        # Espera o menu global abrir
        opcoes = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[role='option']"))
        )

        valor_desejado = valor_desejado.strip().lower()

        for opcao in opcoes:
            texto_opcao = opcao.text.strip().lower()
            print(f"🔍 Verificando opção: '{texto_opcao}'")
            if valor_desejado in texto_opcao:
                print(f"✅ Selecionando: '{opcao.text.strip()}'")
                opcao.click()
                return True

        print(f"⚠️ Nenhuma opção correspondente encontrada: '{valor_desejado}' — dropdown será deixado em branco.")
        dropdown.click()  # Fecha dropdown se não encontrou
        return False

    except Exception as e:
        print(f"⚠️ Erro ao selecionar dropdown: {e}")
        return False

def preencher_formulario(driver, row):
    driver.get(FORM_URL)
    time.sleep(2)

    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Cargo')]"))
        )

        perguntas = driver.find_elements(By.XPATH, "//div[@data-automation-id='questionItem']")

    except TimeoutException:
        print("❌ Não foi possível localizar os campos do formulário.")
        return False

    for i, (coluna, tipo) in enumerate(CAMPOS):
        valor = row.get(coluna)
        if pd.isna(valor) or str(valor).strip() == "":
            continue

        try:
            pergunta = perguntas[i]

            if tipo == "texto":
                campo = pergunta.find_element(By.TAG_NAME, "input")
                campo.clear()
                campo.send_keys(str(valor))

            elif tipo == "data":
                campo = pergunta.find_element(By.TAG_NAME, "input")
                data_formatada = pd.to_datetime(valor).strftime("%d/%m/%Y")
                campo.clear()
                campo.send_keys(data_formatada)

            elif tipo == "dropdown":
                selecionar_dropdown(pergunta, str(valor))

        except Exception as e:
            print(f"⚠️ Erro ao preencher '{coluna}': {e}")
            continue

    try:
        botao_enviar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-automation-id='submitButton']"))
        )
        botao_enviar.click()
        print("✅ Formulário enviado com sucesso!")
        return True
    except Exception as e:
        print(f"❌ Erro ao enviar formulário: {e}")
        return False

# === EXECUÇÃO ===

df = pd.read_excel(EXCEL_PATH)

chrome_options = Options()
chrome_options.add_experimental_option("detach", False)
driver = webdriver.Chrome(options=chrome_options)

for index, row in df.iterrows():
    print(f"\n🔄 Preenchendo linha {index + 1}: {row.to_dict()}")
    try:
        sucesso = preencher_formulario(driver, row)
        if not sucesso:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_name = f"erro_formulario_{timestamp}.png"
            driver.save_screenshot(screenshot_name)
            print(f"📸 Screenshot salva: {screenshot_name}")
    except KeyboardInterrupt:
        print("⛔ Execução interrompida pelo usuário.")
        break
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_name = f"erro_geral_{timestamp}.png"
        driver.save_screenshot(screenshot_name)
        print(f"📸 Screenshot salva: {screenshot_name}")
        continue

driver.quit()
