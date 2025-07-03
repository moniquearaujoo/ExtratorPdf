# 📄 Extrator de Dados de PDFs para Planilhas (SISREG III)

## 📋 Sumário

*   [Visão Geral do Projeto](#-visão-geral-do-projeto)
*   [Funcionalidades](#-funcionalidades)
*   [Tecnologias Utilizadas](#-tecnologias-utilizadas)
*   [Como Configurar e Rodar o Projeto](#-como-configurar-e-rodar-o-projeto)
    *   [Pré-requisitos](#pré-requisitos)
    *   [Instalação](#instalação)
    *   [Configuração](#configuração)
    *   [Como Rodar](#como-rodar)
*   [Uso](#-uso)
*   [Estrutura do Projeto](#-estrutura-do-projeto)
*   [Integração com Google Sheets](#-integração-com-google-sheets)
*   [Melhorias Futuras](#-melhorias-futuras)
*   [Contribuição](#-contribuição)
*   [Licença](#-licença)
*   [Contato](#-contato)

---

## 🚀 Visão Geral do Projeto

Este projeto é uma aplicação web desenvolvida em Python (Flask) que automatiza a extração de informações específicas de documentos PDF gerados pelo sistema **SISREG III** (Sistema Nacional de Regulação). Após a extração, os dados são organizados e podem ser exportados para arquivos CSV/Excel ou diretamente para uma planilha Google Sheets.

O principal objetivo é simplificar e agilizar o processo de coleta de dados de autorizações de procedimentos, que tradicionalmente seria manual e propenso a erros. A aplicação utiliza técnicas avançadas de processamento de texto e expressões regulares para lidar com as variações e ruídos comuns em documentos PDF gerados por OCR ou com layouts complexos.

## ✨ Funcionalidades

*   **Upload de Múltiplos PDFs:** Permite o envio de vários arquivos PDF de uma vez.
*   **Extração Inteligente de Dados:**
    *   Código da Solicitação
    *   CNS do Paciente
    *   Unidade Solicitante (Município de Residência)
    *   Unidade Executante
    *   Data do Exame/Atendimento
    *   Procedimento Autorizado
*   **Limpeza e Normalização de Dados:** Algoritmos robustos para limpar ruídos, caracteres indesejados e inconsistências na extração (ex: "REGULADORBAª Vez", "BRASILEIRA" colado).
*   **Exportação de Dados:**
    *   Download dos dados extraídos em formato CSV.
    *   Download dos dados extraídos em formato Excel (XLSX).
*   **Integração com Google Sheets:** Envio direto dos dados extraídos para uma planilha Google Sheets configurada.
*   **Interface Web Amigável:** Uma interface simples e intuitiva para upload e gerenciamento dos arquivos.

## 🛠️ Tecnologias Utilizadas

*   **Python 3.9+**
*   **Flask:** Microframework web para a aplicação.
*   **PyPDF2:** Para extração de texto de arquivos PDF.
*   **Pandas:** Para manipulação e exportação de dados para CSV/Excel.
*   **`re` (módulo built-in do Python):** Para expressões regulares avançadas na extração e limpeza de dados.
*   **`flask-cors`:** Para lidar com requisições Cross-Origin Resource Sharing.
*   **`werkzeug`:** Para manipulação segura de uploads de arquivos.
*   **Google Sheets API (via `google_sheets_integration_fix.py`):** Para integração com planilhas Google.

## ⚙️ Como Configurar e Rodar o Projeto

### Pré-requisitos

*   Python 3.9 ou superior instalado.
*   `pip` (gerenciador de pacotes do Python).

### Instalação

1.  **Clone o repositório:**
    ```bash
    git clone <URL_DO_SEU_REPOSITORIO>
    cd nome-do-seu-projeto # Substitua pelo nome da pasta do seu projeto
    ```

2.  **Crie e ative um ambiente virtual (recomendado):**
    ```bash
    python -m venv venv
    # No Windows
    .\venv\Scripts\activate
    # No macOS/Linux
    source venv/bin/activate
    ```

3.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```
    (Certifique-se de ter um arquivo `requirements.txt` com todas as dependências: `Flask`, `PyPDF2`, `pandas`, `flask-cors`, `gspread`, `oauth2client` ou `google-auth-oauthlib` dependendo da sua integração com Google Sheets).

### Configuração

1.  **Variáveis de Ambiente:**
    *   Seu projeto pode precisar de variáveis de ambiente para chaves de API ou configurações sensíveis. Crie um arquivo `.env` na raiz do projeto (se estiver usando `python-dotenv`) ou configure-as diretamente no seu ambiente.
    *   Exemplo: `PORT=5000` (se você não quiser usar a porta padrão do Flask).

2.  **Configuração do Google Sheets API:**
    *   Siga as instruções para configurar as credenciais da Google Sheets API. Geralmente, isso envolve:
        *   Criar um projeto no Google Cloud Console.
        *   Habilitar a Google Sheets API.
        *   Criar uma conta de serviço e baixar o arquivo `credentials.json` (ou similar).
        *   Compartilhar a planilha de destino com o e-mail da conta de serviço.
    *   Certifique-se de que o arquivo de credenciais esteja acessível pelo seu script `google_sheets_integration_fix.py`.

### Como Rodar

Após a instalação e configuração, você pode iniciar a aplicação Flask:

```bash
flask run
# Ou, se você tem um script de inicialização:
python main.py
