import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime, timedelta
import schedule
from webdriver_manager.chrome import ChromeDriverManager
import pygetwindow as gw

def acessar_e_capturar_selenium_v2():
    print('Executando Segundo Script...')
    driver = webdriver.Chrome()
    login_url = "https://acostacontactcenter.3c.plus/login"
    ramal = "ramal"
    senha = "senha"

    driver.get(login_url)
    driver.maximize_window()
    time.sleep(2)

    tentativas = 0
    max_tentativas = 10
    login_sucesso = False

    while tentativas < max_tentativas and not login_sucesso:
        try:
            driver.find_element(By.XPATH, '//*[@id="user"]').send_keys(ramal)
            driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(senha)
            driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(Keys.ENTER)
            login_sucesso = True
            print("Login realizado com sucesso.")
        except NoSuchElementException:
            print(f"Elemento de login não encontrado. Tentativa {tentativas + 1} de {max_tentativas}.")
            time.sleep(5)
            driver.refresh()
            tentativas += 1

    if not login_sucesso:
        print("Não foi possível realizar o login após 10 tentativas.")
        driver.quit()
        return

    time.sleep(10)

    # Acessar o relatório de chamadas
    driver.get('https://acostacontactcenter.3c.plus/manager/calls-report')

    hoje = datetime.now()
    dia_anterior = hoje - timedelta(days=1)

    # Verificar se o dia anterior é domingo
    if dia_anterior.weekday() == 6:  # 6 representa domingo
        dia_anterior -= timedelta(days=1)  # Ajusta para sexta-feira (duas dias atrás)

    # Formatar a data conforme solicitado
    d_Menos_1 = dia_anterior.strftime('%d/%m/%Y')
    
    # Preencher o campo com a hora inicial e final corretas
    campo = f"{d_Menos_1} 00:00:00 até {d_Menos_1} 23:59:59"

    # Limpar o campo antes de enviar o valor
    campo_xpath = '//*[@id="app"]/div[1]/div[2]/div[1]/div[2]/div[1]/div/div[1]/input'
    input_campo = driver.find_element(By.XPATH, campo_xpath)
    input_campo.click()
    input_campo.clear()  # Limpar o campo
    time.sleep(1)  # Aguardar para garantir campo limpo (opcional)
    input_campo.send_keys(campo)  # Enviar a data formatada

    # Clicar no botão para aplicar a busca
    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[1]/div[2]/div[5]/button'))
    ).click()

    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[1]/div[7]/div[1]/div/div[1]/div/div[2]/button'))
    ).click()

    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="formCSV"]/div[1]/div[2]/label'))
    ).click()

    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="collapse4"]/div/div[1]/div[1]/div/label'))
    ).click()

    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="collapse4"]/div/div[2]/div/div/label'))
    ).click()

    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="collapse4"]/div/div[3]/div/div/label'))
    ).click()

    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="collapse4"]/div/div[1]/div[2]/div/label'))
    ).click()

    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="input-email"]'))
    ).clear()

    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="input-email"]'))
    ).send_keys('email para onde ele vai ser enviado')
    
    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="formCSV"]/div[4]/label'))
    ).click()

    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="formCSV"]/div[5]/label'))
    ).click()

    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="formCSV"]/button'))
    ).click()
    # driver.quit()
    print("Relatório puxado.")

acessar_e_capturar_selenium_v2()
# Agendar a execução do script
schedule.every().day.at("06:00").do(acessar_e_capturar_selenium_v2)
print(f"{datetime.now()}: Iniciando o agendador de scripts.")
while True:
    schedule.run_pending()
    time.sleep(60)  # Esperar um minuto antes de verificar novamente