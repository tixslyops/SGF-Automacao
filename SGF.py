import time
import os
import glob
import csv
import xlrd
import logging
import pandas as pd
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC

# ======= LOG ===================================================================================================================================================================
pasta_script = r"C:/Automação/Automacao_415"
os.makedirs(pasta_script, exist_ok=True)
LOG_PATH = os.path.join(pasta_script, "nome_automacao.log")

logging.basicConfig(
    filename=LOG_PATH,
    filemode='a',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ======= CREDENCIAIS ===================================================================================================================================================================

USUARIO = 'usuario'
SENHA = 'senha'
LINK = 'https://sistema-da-empresa.com/login'

PASTA_DESTINO = r"caminho\da\pasta"
PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "pasta", "downloads")

os.makedirs(PASTA_DOWNLOADS, exist_ok=True)
os.makedirs(PASTA_DESTINO, exist_ok=True)

# ======= DOWNLOAD ===================================================================================================================================================================

def esperar_download(pasta, extensao=".xls", timeout=3600):
    logging.info("Aguardando download...")
    inicio = time.time()
    while time.time() - inicio < timeout:
        arquivos = glob.glob(os.path.join(pasta, f"*{extensao}"))
        if arquivos:
            arquivo = max(arquivos, key=os.path.getmtime)
            if not arquivo.endswith(".part") and not arquivo.endswith(".crdownload"):
                if os.path.getsize(arquivo) > 1000:
                    return arquivo
        time.sleep(5)
    raise TimeoutError("Arquivo não foi baixado no tempo esperado")


# ======= INÍCIO AUTOMAÇÃO ===================================================================================================================================================================
try:
    logging.info("Iniciando automação Cubo 415...")
    print("Iniciando automação Cubo 415...")

    options = ChromeOptions()
    options.add_experimental_option("detach", True)
    
    prefs = {"download.default_directory": PASTA_DOWNLOADS}
    options.add_experimental_option("prefs", prefs)
    
    servico = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=options)
    
    driver.command_executor.set_timeout(600) 
    driver.set_page_load_timeout(3600)
    wait = WebDriverWait(driver, 60)

    driver.get(LINK)
    wait.until(EC.presence_of_element_located((By.ID, "txtLogin"))).send_keys(USUARIO)
    campo_senha = driver.find_element(By.ID, "txtSenha")
    campo_senha.send_keys(SENHA)
    time.sleep(1)
    
    print("Enviando credenciais...")
    campo_senha.send_keys(Keys.ENTER)

    logging.info("Login enviado. Aguardando a página processar...")
    
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[id$='txtDataInicio_1']")))
        print("Login efetuado com sucesso! Filtros carregados.")
    except:
        if "Login" in driver.title:
            driver.execute_script("document.getElementById('btnOk').click();")
            for _ in range(12):
                time.sleep(10)
                if len(driver.find_elements(By.CSS_SELECTOR, "input[id$='txtDataInicio_1']")) > 0:
                    break

# ======= DATAS ===================================================================================================================================================================

    hoje = datetime.today()
    inicio_mes = hoje.replace(day=1).strftime('%d/%m/%Y')
    proximo_mes = hoje.replace(day=28) + timedelta(days=4)
    fim_mes = (proximo_mes - timedelta(days=proximo_mes.day)).strftime('%d/%m/%Y')

    try:
        campo_ini = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[id$='txtDataInicio_1']")))
        campo_fim = driver.find_element(By.CSS_SELECTOR, "input[id$='txtDataFim_1']")

        print(f"Preenchendo datas via JS: {inicio_mes} até {fim_mes}")
        driver.execute_script("arguments[0].value = arguments[1];", campo_ini, inicio_mes)
        driver.execute_script("arguments[0].value = arguments[1];", campo_fim, fim_mes)
        
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", campo_ini)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", campo_fim)
        time.sleep(2)

        driver.execute_script("document.getElementById('btnOk').click();")
        logging.info(f"Filtro aplicado com sucesso: {inicio_mes} a {fim_mes}")
        
    except Exception as e:
        raise e

# ======= PROCESSAMENTO E EXPORTAÇÃO ===================================================================================================================================================================

    print("Aguardando processamento do Cubo (isso pode demorar)...")
    wait.until(lambda d: len(d.window_handles) > 1)
    
    driver.switch_to.window(driver.window_handles[-1])
    print("Foco alterado para a aba do relatório.")

    time.sleep(60)

    driver.execute_script("""
        setTimeout(function() {
            const btn = document.querySelector("input[type='submit'][value*='Exportar']");
            if (btn) {
                btn.click();
            }
        }, 100);
    """)
    
    print("Comando de clique enviado para o navegador. Monitorando início do download...")
    time.sleep(10)

# ======= DOWNLOAD E CONVERSÃO ===================================================================================================================================================================
 
    arquivo_xls = esperar_download(PASTA_DOWNLOADS)

    nome_csv = hoje.strftime('%m-%Y') + ".csv"
    caminho_csv_final = os.path.join(PASTA_DESTINO, nome_csv)

    print(f">> Convertendo arquivo para: {nome_csv}")

    try:
        tabelas = pd.read_html(arquivo_xls)
        df = tabelas[0]
        df.to_csv(caminho_csv_final, index=False, sep=';', encoding='cp1252')
        print(f"--- SUCESSO TOTAL! Arquivo salvo em: {PASTA_DESTINO} ---")

    except Exception as conv_err:
        logging.error(f"Erro na conversão Pandas: {conv_err}. Tentando via xlrd...")
        wb = xlrd.open_workbook(arquivo_xls)
        sheet = wb.sheet_by_index(0)
        with open(caminho_csv_final, 'w', newline='', encoding='cp1252') as f:
            writer = csv.writer(f, delimiter=';')
            for r in range(sheet.nrows):
                linha = [
                    str(int(sheet.cell_value(r, c))) if isinstance(sheet.cell_value(r, c), float) and sheet.cell_value(r, c).is_integer() 
                    else str(sheet.cell_value(r, c)) for c in range(sheet.ncols)
                ]
                writer.writerow(linha)
        print(f"Arquivo salvo via Fallback em: {PASTA_DESTINO}")

    if os.path.exists(arquivo_xls):
        os.remove(arquivo_xls)

    driver.quit()

except Exception as e:
    logging.exception("Erro crítico na automação 415")
    print(f"\n!!! OCORREU UM ERRO: {e}")
    if 'driver' in locals():
        driver.quit()
