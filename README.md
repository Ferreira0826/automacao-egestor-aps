# Pipeline de Automação RPA: Extração de Dados e Integração Contínua com Google Sheets

## 📋 Descrição do Projeto
Este projeto consiste no desenvolvimento de uma solução robusta de **RPA (Robotic Process Automation)** e **Integração de Dados** utilizando Python. O objetivo principal é automatizar o fluxo completo de recolha, tratamento e consolidação de relatórios de financiamento público de saúde a partir do portal governamental e-Gestor APS, inserindo-os diretamente num dashboard operacional no Google Sheets.

A solução foi desenhada para eliminar processos manuais repetitivos, mitigar falhas humanas de digitação e garantir a integridade histórica dos dados através de travas de segurança antiduplicidade.

## 🚀 Funcionalidades Principais
* **Web Scraping Avançado:** Navegação e preenchimento dinâmico de filtros multinível (Tipo de Unidade, UF, Município e Período) simulando o comportamento humano via Playwright.
* **Resiliência a Instabilidades:** Implementação de estratégias de esperas explícitas e inteligência de varredura para lidar com a lentidão nativa e alterações estruturais do portal do governo.
* **Pipeline de Dados Limpo:** Utilização da biblioteca Pandas para isolar, higienizar e converter os dados brutos de ficheiros Excel (.xlsx) antes da exportação.
* **Validação Dupla Antiduplicidade:** Algoritmo rigoroso que higieniza strings (remoção de espaços invisíveis e normalização de maiúsculas/minúsculas) e valida de forma combinada o mês de referência e a parcela atual, impedindo a sobreposição de dados históricos.
* **Formatação Visual Automatizada:** Comunicação direta com a API do Google Sheets para forçar a padronização tipográfica das tabelas (Fonte Calibri, Tamanho 10) sem corromper as máscaras de dados e propriedades nativas de células de data.
* **Segurança e Boas Práticas:** Arquitetura protegida contra a exposição de caminhos de rede locais ou chaves privadas através do isolamento de variáveis de ambiente (`.env`).

## 🛠️ Tecnologias e Ferramentas Utilizadas
* **Linguagem Principal:** Python 3.10+
* **Automação Web (RPA):** Playwright
* **Manipulação e Análise de Dados:** Pandas / Openpyxl
* **Integração com Cloud API:** Gspread / Google OAuth2 (Google Cloud Console)
* **Gestão de Variáveis de Ambiente:** Python-dotenv
* **Orquestração:** Agendador de Tarefas do Windows (Execução em Background via `.bat`)

## 📁 Estrutura do Repositório
```text
├── automacao_sistema.py    # Script principal da automação RPA e pipeline
├── .env.example            # Modelo de configuração das variáveis de ambiente
├── .gitignore              # Proteção de credenciais e ficheiros locais
├── README.md               # Documentação técnica do projeto
└── executar_robo.bat       # Ficheiro de lote para orquestração local