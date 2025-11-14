from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import win32com.client as win32
from PIL import ImageGrab
import schedule
import datetime
import time
import os

# --- NOVOS Imports ---
import google.generativeai as genai
from PIL import Image # Para carregar a imagem salva
import pyperclip      # Para copiar o texto da análise

# --- Configurações ---
CAMINHO_EXCEL = r""
EMAIL = ''
SENHA = ''
DESTINATARIO = ''
# DESTINATARIO = 'Plan&Gestão - Itaú PJ'

# --- NOVO: Configuração da API do Gemini ---
# !!! COLE SUA CHAVE DE API AQUI !!!
# Obtenha sua chave em: https://aistudio.google.com/app/apikey
try:
    genai.configure(api_key="CHAVE-API-AQUI")
    print("🔑 API do Gemini configurada.")
except Exception as e:
    print(f"❌ Erro ao configurar a API do Gemini: {e}. Verifique sua API Key.")
    # Considerar sair do script se a API for essencial:
    # import sys
    # sys.exit(1)


# --- Função para gerar print do Excel (sem mudanças, exceto no print final) ---
def gerar_print_excel():
    print("📊 Abrindo Excel e gerando print...")
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Open(CAMINHO_EXCEL)

    # Atualiza dados
    for i in range(1):
        print(f"🔄 Atualização {i + 1}/3...")
        wb.RefreshAll()
        excel.CalculateFullRebuild()
        excel.CalculateUntilAsyncQueriesDone()
        time.sleep(3)

    # Segmentações
    slicer_dia = wb.SlicerCaches("SegmentaçãodeDados_DIA_UTIL2")
    slicer_hora = wb.SlicerCaches("SegmentaçãodeDados_hora2")

    # Seleciona todos os dias disponíveis
    dias_disponiveis = [
        int(item.Name) for item in slicer_dia.SlicerItems
        if getattr(item, 'HasData', False)
    ]
    dias_disponiveis.sort()

    slicer_dia.ClearManualFilter()
    for item in slicer_dia.SlicerItems:
        try:
            item.Selected = getattr(item, 'HasData', False)
        except:
            item.Selected = False

    print(f"📅 Dias selecionados: {dias_disponiveis}")

    # Seleciona horas de 7 até a hora atual
    hora_atual = datetime.datetime.now().hour
    slicer_hora.ClearManualFilter()

    for item in slicer_hora.SlicerItems:
        try:
            h = int(item.Name)
            item.Selected = 7 <= h <= hora_atual-1
        except:
            item.Selected = False

    print(f"⏰ Horas selecionadas: 7 até {hora_atual}")

    # Copia área desejada como imagem
    ws = wb.Sheets["Comparativo"]
    rng = ws.Range("B2:U27")
    rng.CopyPicture(Appearance=1, Format=2)
    time.sleep(1)

    image = ImageGrab.grabclipboard()

    if image:
        pasta_destino = r""
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)

        nome_arquivo = f"Print_HxH_{datetime.datetime.now().strftime('%d-%m-%Y_%H-%M')}.png"
        caminho_imagem = os.path.join(pasta_destino, nome_arquivo)
        image.save(caminho_imagem)
        print(f"💾 Imagem salva em: {caminho_imagem}")
    else:
        print("⚠️ Nenhuma imagem encontrada na área de transferência.")
        caminho_imagem = None

    wb.Close(SaveChanges=True)
    excel.Quit()

    return caminho_imagem

# --- NOVA FUNÇÃO: Analisar imagem com Gemini ---
def analisar_imagem_com_gemini(caminho_imagem):
    print(f"🤖 Enviando imagem '{caminho_imagem}' para análise do Gemini...")
    if not caminho_imagem:
        print("⚠️ Caminho da imagem está vazio. Abortando análise.")
        return None
    try:
        img = Image.open(caminho_imagem)

        # Configura o modelo (gemini-1.5-flash é rápido e multimodal)
        model = genai.GenerativeModel('gemini-2.5-flash')

        # --- Este é o PROMPT que você pediu ---
        prompt_texto = f"""
        Você é um analista de BI e Planejamento, especialista em performance de operações de call center (cobrança/vendas).
        A hora atual é {datetime.datetime.now().strftime('%H:%M')}.
        
        Analise a imagem de dashboard (print de Excel HxH) e forneça uma análise concisa e direta, em formato de tópicos:

        1.  **Status das Carteiras (Meta Hora):**
            * Quais carteiras estão **dentro** da meta hora?
            * Quais carteiras estão **abaixo** da meta hora?

        2.  **Análise Comparativa (CPC e CPC Produtivo):**
            * Analise o CPC (Contato com Pessoa Certa) e o CPC Produtivo de hoje em comparação com os dias anteriores visíveis na imagem.

        3.  **Projeção de Fechamento (Valor Realizado):**
            * Com base no ritmo atual, qual é a projeção de "Valor Realizado" até as 18:12?

        4.  **Plano de Ação Imediato:**
            * Sugira 2-3 planos de ação **práticos e urgentes** para a operação (ex: "Focar carteira X", "Ajustar mailing", "Revisar estratégia de CPC") para reverter o cenário das carteiras que estão abaixo da meta.

        Seja objetivo e baseie-se estritamente nos dados da imagem.
        """

        # Envia o prompt de texto e a imagem
        response = model.generate_content([prompt_texto, img])

        print("✅ Análise do Gemini recebida.")
        return response.text

    except Exception as e:
        print(f"❌ Erro ao analisar imagem com Gemini: {e}")
        return f"Erro ao gerar análise do Gemini: {e}" # Retorna um erro amigável


# --- FUNÇÃO MODIFICADA: Enviar ANÁLISE (texto) pelo Teams ---
def enviar_analise_teams(texto_da_analise, caminho_imagem):
    print("💬 Abrindo Teams para enviar IMAGEM + ANÁLISE...")

    options = Options()
    options.add_argument("--log-level=3")
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 30)

    driver.get("https://teams.microsoft.com/v2/")

    # ---------------- LOGIN ESTÁVEL ----------------
    email = wait.until(EC.element_to_be_clickable((By.ID, "i0116")))
    email.send_keys(EMAIL)
    wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

    senha = wait.until(EC.element_to_be_clickable((By.ID, "passwordInput")))
    senha.send_keys(SENHA)
    wait.until(EC.element_to_be_clickable((By.ID, "submitButton"))).click()

    try:
        wait.until(EC.element_to_be_clickable((By.ID, "KmsiCheckboxField"))).click()
    except:
        pass

    wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

    print("⏳ Aguardando carregamento do Teams...")
    time.sleep(20)

    # ---------------- ENTRAR NO CHAT ----------------
    search_box = wait.until(EC.element_to_be_clickable((By.ID, "ms-searchux-input")))
    search_box.send_keys(DESTINATARIO)

    time.sleep(2)
    search_box.send_keys(Keys.ARROW_DOWN)
    time.sleep(1)
    search_box.send_keys(Keys.ENTER)

    time.sleep(10)

    # ======================================================
    #            *** PRIMEIRA MENSAGEM: IMAGEM ***
    # ======================================================

    print("📎 Enviando print...")

    # Copia a IMAGEM para o clipboard
    try:
        img = Image.open(caminho_imagem)
        img.load()
        pyperclip.copy("")  # limpa clipboard
        img.save("temp_clipboard.png")
        os.system(f"powershell.exe Set-Clipboard -Path 'temp_clipboard.png'")
    except Exception as e:
        print(f"❌ Erro ao copiar imagem: {e}")

    # Cola (CTRL+V) no chat
    actions = ActionChains(driver)
    actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()

    time.sleep(3)
    actions.send_keys(Keys.ENTER).perform()
    print("✅ Print enviado.")

    time.sleep(4)

    # ======================================================
    #        *** SEGUNDA MENSAGEM: ANÁLISE TEXTO ***
    # ======================================================

    print("✍️ Enviando análise...")

    pyperclip.copy(texto_da_analise)

    prefixo = f"📊 **Análise Automática - {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}**\n\n"

    actions = ActionChains(driver)
    actions.send_keys(prefixo)
    actions.perform()

    time.sleep(1)

    actions = ActionChains(driver)
    actions.key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()

    time.sleep(3)
    actions.send_keys(Keys.ENTER).perform()

    print("✅ Análise enviada com sucesso!")

    time.sleep(4)
    driver.quit()

# --- Função completa (MODIFICADA) ---
def tarefa_completa():
    print(f"\n--- 🚀 Iniciando tarefa completa: {datetime.datetime.now()} ---")
    try:
        # Passo 1: Gerar o print
        caminho_imagem = gerar_print_excel()
        
        if caminho_imagem:
            # Passo 2: Analisar a imagem com Gemini
            texto_analise = analisar_imagem_com_gemini(caminho_imagem)
            
            if texto_analise:
                # Passo 3: Enviar a ANÁLISE (texto) para o Teams
                enviar_analise_teams(texto_analise)
            else:
                print("⚠️ Análise do Gemini falhou ou retornou vazia. Envio para o Teams cancelado.")
        else:
            print("⚠️ Geração do print falhou. Tarefa abortada.")
    
    except Exception as e:
        print(f"❌ Erro fatal na execução da tarefa_completa: {e}")
    print(f"--- ✅ Tarefa finalizada: {datetime.datetime.now()} ---")


# --- TESTE ANTES DO AGENDAMENTO ---
print("🚀 Executando teste único antes do agendamento...")

# 1) Gera o print
caminho_imagem = gerar_print_excel()

# 2) Analisa a imagem
texto_analise = analisar_imagem_com_gemini(caminho_imagem)

# 3) Envia a análise ao Teams
if texto_analise:
    enviar_analise_teams(texto_analise)
else:
    print("⚠️ Nenhuma análise foi gerada pelo Gemini. Teste abortado.")

print("✅ Teste finalizado. Iniciando agendamentos...\n")


# --- Agendamento (sem mudanças) ---
for hora in range(9, 18):  # 9 até 17
    schedule.every().day.at(f"{hora:02d}:00").do(tarefa_completa)
schedule.every().day.at("17:12").do(tarefa_completa)

print("📅 Agendamento iniciado. Tarefas serão executadas nos seguintes horários:")
print(", ".join([f"{h:02d}:00" for h in range(9, 18)]) + ", 17:12")

# --- Loop principal (sem mudanças) ---
while True:
    schedule.run_pending()
    time.sleep(1)