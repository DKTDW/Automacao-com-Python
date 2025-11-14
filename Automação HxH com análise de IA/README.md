📊 Automação de Dashboards HxH – Excel → Análise Gemini → Envio Teams

Este projeto automatiza todo o fluxo diário de monitoramento operacional, realizando:

Abertura de um arquivo Excel

Atualização de dados e segmentações

Geração automática de um print do dashboard

Análise inteligente da imagem via IA (Gemini 2.5 Flash)

Envio automático do print + análise para o Microsoft Teams

Agendamento horário (9h às 17h + 17:12)

Ideal para operações de call center, cobrança, vendas ou qualquer área que monitora produtividade hora a hora (HxH).

🚀 Funcionalidades
✔ Atualiza e recalcula planilha Excel

Abre o Excel invisível

Força RefreshAll

Recalcula consultas assíncronas

Ajusta segmentações (slicers) automaticamente

✔ Gera screenshot automático

O script salva um print de um range fixo (B2:U27) em:

\\Servidor\...\Prints HxH\Print_HxH_DD-MM-YYYY_HH-MM.png

✔ Análise da imagem com IA

Usa o Google Gemini (modelo gemini-2.5-flash) para:

Avaliar performance por carteira

Comparar CPC/CPC Produtivo

Projetar resultados

Sugerir plano de ação

Gerar texto conciso e objetivo

✔ Envia imagem + análise para o Microsoft Teams

Login automático (email + senha)

Primeiro envia a imagem

Depois envia a análise formatada

Funciona em chats individuais ou canais

✔ Task scheduler integrado

Executa automaticamente:

09h00 – 10h00 – 11h00 – 12h00 – 13h00 – 14h00 – 15h00 – 16h00 – 17h00 – 17h12

🛠 Tecnologias Utilizadas
Tecnologia	Uso
Python 3.10+	Linguagem principal
Selenium + ChromeDriver	Automação do login/envio no Teams
win32com	Manipulação do Excel
Google Gemini API	Análise da imagem e geração de texto
schedule	Agendamentos automáticos
PIL (Pillow)	Captura e leitura da imagem
pyperclip	Clipboard para colar no Teams
📁 Estrutura do Projeto
automacao_hxh/
│
├── automacao_hxh.py          # Script principal
├── README.md                 # Este arquivo
└── /Prints HxH/              # Pasta onde os prints são salvos

🔧 Configuração
1. Instalar dependências
pip install selenium webdriver-manager pyperclip pillow schedule google-generativeai pywin32

2. Criar uma API Key do Gemini

Acesse:

👉 https://aistudio.google.com/app/apikey

Copie sua chave e cole nesta parte do código:

genai.configure(api_key="CHAVE-API-AQUI")

3. Configurar credenciais e caminhos

Edite:

CAMINHO_EXCEL = r"C:\...\sua_planilha.xlsx"
EMAIL = "seu_email"
SENHA = "sua_senha"
DESTINATARIO = "Nome do Grupo ou Pessoa"


E também:

pasta_destino = r"\\192.168....\Prints HxH"

▶ Como Executar
Modo teste

Executa:

abrir Excel

gerar print

analisar com IA

enviar para o Teams

Tudo uma única vez:

python automacao_hxh.py

Modo produção

O script já inicia automaticamente os agendamentos ao final da execução:

Agendamento iniciado. Tarefas serão executadas nos seguintes horários:
09:00, 10:00, 11:00, 12:00, 13:00, 14:00, 15:00, 16:00, 17:00, 17:12


Basta deixar a janela aberta.

⚠ Requisitos Importantes
🔹 O Chrome deve estar instalado

O Selenium usa o ChromeDriver automaticamente.

🔹 O Teams será automatizado pelo navegador

Não use o PC enquanto o script estiver enviando mensagens, pois ele utiliza teclado/clipboard.

🔹 Excel não pode estar aberto no mesmo arquivo

Evita conflito no wb.Open().

🧠 Lógica do Fluxo
[Excel] → Atualiza dados
         → Seleciona slicers
         → Gera print

[IA Gemini] → Interpreta imagem
            → Gera análise inteligente

[Teams] → Envia imagem
        → Envia análise

🐞 Troubleshooting
✓ Ele não consegue clicar no botão “Avançar”

Problema de carregamento → resolvido usando:

wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))

✓ Não colou imagem no Teams

Certifique-se de que:

clipboard está funcional

Powershell está habilitado

✓ A análise veio vazia

Verifique:

API Key ativa

Modelo “gemini-2.5-flash” está liberado na sua conta

📌 Melhorias Futuras

Envio direto no Teams usando API oficial (quando liberada)

Integração com Power BI Online

Logging em arquivo .log

Múltiplos canais simultâneos

Geração de PDF da análise

📄 Licença

Livre para uso interno e corporativo.
Não contém dados sensíveis — personalize livremente.