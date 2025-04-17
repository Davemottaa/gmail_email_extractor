2# Extração de E-mails Únicos via Gmail API com Exportação para Excel

Este projeto em **Python** autentica um usuário do Gmail, percorre toda a caixa de entrada e extrai **nomes e endereços de e-mail únicos** dos remetentes. Ele inclui uma interface gráfica para selecionar o local de salvamento da planilha `.xlsx` e uma **interface amigável no terminal**, com barras de progresso e simulação de login seguro.

## Funcionalidades

- Autenticação segura com OAuth 2.0 usando Gmail API
- Contagem total de e-mails com barra de progresso (`tqdm`)
- Extração de metadados de e-mails (`From: Nome <email>`)
- Filtro inteligente para evitar e-mails duplicados
- Exportação para planilha `.xlsx` com `pandas`
- Interface gráfica para escolha do local de salvamento com `tkinter`
- Simulação de login seguro no terminal com animações

## Tecnologias e Bibliotecas

- Python 3.x
- `google-auth`, `google-auth-oauthlib`, `google-api-python-client`
- `tkinter` (GUI embutida do Python)
- `pandas`
- `tqdm`
- `openpyxl`

## Pré-requisitos

- Uma conta do Google
- Criar um projeto no [Google Cloud Console](https://console.cloud.google.com/)
- Ativar a **Gmail API**
- Fazer download do arquivo `credentials.json` (OAuth 2.0 Client ID)
- Instalar dependências:

```bash
pip install -r requirements.txt

Como usar

1. Coloque o arquivo credentials.json na mesma pasta do script.


2. Execute o script:



python main.py

3. O navegador será aberto pedindo para autenticar sua conta Google.


4. Após login, o script:

Simula autenticação segura

Conta e processa e-mails com barras de progresso

Cria uma planilha .xlsx com nome e e-mail únicos

Pergunta onde deseja salvar a planilha




Estrutura da planilha gerada

Captura de tela (terminal)

Iniciando autenticação...
Simulando autenticação: %$@#!@*&%$@#...
Login bem-sucedido!
Contando e-mails:  100%|██████████████████| 1300/1300 [00:12<00:00]
Processando e-mails:  100%|███████████████| 1300/1300 [00:18<00:00]
Salvo 246 e-mails únicos em 'C:/Users/Davi/Documents/emails.xlsx'

Licença

MIT
