from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from datetime import datetime, timedelta
import pandas as pd

# Configuração do WebDriver
driver = webdriver.Chrome()  # Ou outro driver, como webdriver.Firefox()

# Acesse a página de login
driver.get("http://url.da.totalip/")
usuario = "ramal"
senha = "senha"
driver.maximize_window()

# Realiza o login
driver.find_element(By.XPATH, '//*[@id="inputEmail"]').send_keys(usuario)
driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(senha)
driver.find_element(By.XPATH, '//*[@id="bg-softphone"]/div/div[2]/div/div/form/button').click()

# Dicionário para armazenar o número de vezes que cada campanha foi ressubmetida
campaign_resubmit_count = {}
# Dicionário para armazenar a última hora de resubmissão
last_resubmit_time = {}

# Função para aguardar um elemento ser clicável
def wait_for_clickable(by, value, timeout=10):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, value))
    )

# Função para verificar e resubmeter campanhas
def verificar_e_ressubmeter():
    try:
        # Navega para a página de campanhas
        driver.get("http://url.da.totalip/campaign/list")
        wait_for_clickable(By.XPATH, '//*[@id="table_campanha"]')  # Espera até a tabela estar presente

        # Localize a tabela
        table_xpath = '//*[@id="table_campanha"]'
        table = driver.find_element(By.XPATH, table_xpath)

        # Encontra todas as linhas da tabela
        rows = table.find_elements(By.XPATH, './tbody/tr')

        def has_opacity_oscillation(row):
            styles = [row.value_of_css_property('opacity') for _ in range(5)]  # Verifica o estilo várias vezes
            return len(set(styles)) > 1  # Verifica se há variação

        # Variáveis para armazenar as campanhas que piscam
        oscillating_campaigns = []

        # Verifique cada linha
        for index, row in enumerate(rows):
            if has_opacity_oscillation(row):
                # Armazena o nome da campanha
                campaign_name = row.find_element(By.XPATH, './td[2]/a').text.strip().replace("...", "")
                oscillating_campaigns.append(campaign_name)

        # Verifica se há campanhas com oscilação
        if oscillating_campaigns:
            for campaign_name in oscillating_campaigns:
                current_time = datetime.now()

                # Verifique se a campanha foi resubmetida recentemente
                last_time = last_resubmit_time.get(campaign_name)
                if last_time:
                    time_since_last = current_time - last_time
                    if time_since_last < timedelta(seconds=60):
                        continue
                    elif time_since_last < timedelta(minutes=4):
                        continue

                try:
                    # Clica no botão para abrir o modal
                    wait_for_clickable(By.XPATH, '//*[@id="table_campanha_wrapper"]/div[1]/div[1]/div/a[5]/span').click()
                    wait_for_clickable(By.XPATH, '//*[@id="modalResubmeterStatus"]/div/div/div[2]/div[1]/button').click()

                    # Escreve o nome da campanha no campo de entrada
                    input_field = wait_for_clickable(By.XPATH, '//*[@id="modalResubmeterStatus"]/div/div/div[2]/div[1]/div/div/input')
                    input_field.clear()
                    input_field.send_keys(campaign_name)
                    input_field.send_keys(Keys.ENTER)

                    # Clica no botão "Todos" e desmarca um elemento específico
                    wait_for_clickable(By.XPATH, '//*[@id="selectAllStatus"]/button[1]').click()
                    wait_for_clickable(By.XPATH, '//*[@id="campaign_customer_statuses_34"]').click()
                    wait_for_clickable(By.XPATH, '//*[@id="confirmarResubmeterStatus"]').click()
                    wait_for_clickable(By.XPATH, '//*[@id="confirmarSalvarBtn"]/a').click()
                    
                    # Captura o horário atual e exibe
                    current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    # Atualiza o contador de ressubmissões e o timestamp
                    campaign_resubmit_count[campaign_name] = campaign_resubmit_count.get(campaign_name, 0) + 1
                    last_resubmit_time[campaign_name] = datetime.now()

                    # Exibe mensagem ao usuário
                    print(f"{campaign_name} foi ressubmetida às {current_time_str}.")  # Mensagem para o usuário

                except Exception as e:
                    print(f"Erro ao ressubmeter a campanha {campaign_name}: {e}")

            print("Todas as campanhas foram ressubmetidas.")
        else:
            print("Sem campanhas para ressubmeter.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

# Função para coletar dados das campanhas e gerar a planilha
def coletar_dados_e_gerar_planilha():
    try:
        # Navega para a página de campanhas
        driver.get("http://url.da.totalip/campaign/list")
        wait_for_clickable(By.XPATH, '//*[@id="table_campanha"]')  # Espera a tabela

        # Localiza a tabela
        table_xpath = '//*[@id="table_campanha"]'
        table = driver.find_element(By.XPATH, table_xpath)

        # Encontre todas as linhas da tabela
        rows = table.find_elements(By.XPATH, './tbody/tr')
        campaign_data = []

        # Coleta dados de cada linha
        for row_index, row in enumerate(rows):
            try:
                campaign_name = row.find_element(By.XPATH, './td[2]/a').text.strip().replace("...", "")
                clients_xpath = f'//*[@id="table_campanha"]/tbody/tr[{row_index + 1}]/td[6]'
                num_clients = driver.find_element(By.XPATH, clients_xpath).text
                num_resubmits = campaign_resubmit_count.get(campaign_name, 0)

                campaign_data.append([campaign_name, num_clients, num_resubmits])
            except Exception as inner_e:
                print(f"Erro ao coletar dados da linha: {inner_e}")

        # Cria um DataFrame do pandas
        df = pd.DataFrame(campaign_data, columns=['Campanha', 'Total de Clientes', 'Resubmissões'])

        # Verifica se existem dados para salvar
        if not df.empty:
            df.to_excel('campanhas_resubmissao.xlsx', index=False)
            print("Planilha 'campanhas_resubmissao.xlsx' gerada com sucesso.")
        else:
            print("Nenhum dado coletado para salvar na planilha.")
    except Exception as e:
        print(f"Erro ao coletar dados e gerar planilha: {e}")

# Loop infinito para verificar a cada 3 minutos
while True:
    verificar_e_ressubmeter()
    current_time = datetime.now().strftime("%H:%M")
    if current_time == "20:10":
        coletar_dados_e_gerar_planilha()
    WebDriverWait(driver, 60)  # Espera 60 segundos antes de verificar novamente

driver.quit()
