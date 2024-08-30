from selenium import webdriver
import os
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta

def enviar_imagem_whatsapp_discagem():
    print("Enviando Prints no grupo...")
    # Configurações do usuário
    NOME_DO_GRUPO = "nome do grupo"  # Nome do grupo ou contato para enviar as imagens
    IMAGE_PATHS = [
        # Caminhos das imagens a serem enviadas
        r"caminho/para/imagem"
    ]

    # Inicializa o serviço do ChromeDriver usando o WebDriverManager
    service = Service(ChromeDriverManager().install())

    # Define o caminho do diretório para o novo perfil do Chrome
    dir_path = os.getcwd()
    profile = os.path.join(dir_path, "new_profile", "wpp")

    # Cria o diretório se ele não existir
    if not os.path.exists(profile):
        os.makedirs(profile)

    # Configura as opções do Chrome
    options = webdriver.ChromeOptions()
    options.add_argument(f"user-data-dir={profile}")

    # Inicializa o driver do Chrome com o serviço e as opções configuradas
    browser = webdriver.Chrome(service=service, options=options)

    # Abre o WhatsApp Web
    browser.get("https://web.whatsapp.com")

    print("Por favor, escaneie o código QR para acessar o WhatsApp Web.")
    time.sleep(60)

    contagem = 0

    while contagem <= 100:
        try:
            browser.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p')
            print('Objeto encontrado')                        
            break
        except NoSuchElementException:
            print('Elemento não encontrado, aguardando 20s antes de tentar')
            time.sleep(20)
            contagem += 1 

    if contagem == 100:
        print('Elemento não encontrado depois de 10 tentativas')
        browser.quit()                  

    # Encontra a barra de pesquisa e digita o nome do grupo
    search = browser.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p')
    search.send_keys(NOME_DO_GRUPO)
    search.send_keys(Keys.ENTER)

    # Espera um momento para que os resultados da pesquisa sejam carregados
    time.sleep(10)

    # Envia cada imagem na sequência
    for image_path in IMAGE_PATHS:
        # Verifica se a imagem existe
        if not os.path.exists(image_path):
            print(f"A imagem não foi encontrada: {image_path}")
            continue

        # Clica no ícone de anexo
        attach_button = browser.find_element(By.XPATH, '//div[@title="Anexar"]')
        attach_button.click()
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
        )
        
        # Seleciona a opção de enviar fotos e vídeos
        photo_video_button = browser.find_element(By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        photo_video_button.send_keys(image_path)
        
        # Espera um momento para a imagem ser carregada
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div'))
        )
        
        description = f"Texto da Imagem"
        caption_box = browser.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]/p')
        caption_box.send_keys(description)
        
        # Envia a imagem
        send_button = browser.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div')
        send_button.click()
        
        # Espera um momento para garantir que a imagem seja enviada
        WebDriverWait(browser, 10).until(
            EC.invisibility_of_element_located((By.XPATH, '//span[@data-testid="send"]'))
        )

        time.sleep(10)

    # Período para garantir que as imagens foram enviadas antes de fechar o navegador
    time.sleep(300)
    # Fecha o navegador após o envio de todas as imagens
    browser.quit()
