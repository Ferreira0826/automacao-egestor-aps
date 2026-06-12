# Robô SISAB — Pipeline de Automação RPA

> **e-Gestor → Google Sheets → Planilha Local → eCIEGES**

Solução de **RPA (Robotic Process Automation)** em Python que automatiza o ciclo completo de coleta, consolidação e distribuição de dados de financiamento público de saúde. Parte do **portal e-Gestor APS** (governo federal), passa por **Google Sheets**, atualiza a **planilha consolidada no servidor** e faz upload no sistema interno **eCIEGES**.

---

## 🔄 Fluxo Completo do Pipeline

```
┌─────────────────────────────────────────────────────────────────┐
│  1. Portal e-Gestor APS                                         │
│     Robô faz login, aplica filtros (DF/Brasília/Ano/Parcela)    │
│     e baixa o XLSX da parcela mais recente                      │
└────────────────────────────────┬────────────────────────────────┘
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  2. Salva XLSX bruto no servidor                                │
│     \\------\---\-----\{ANO}\IAF\...                           │
└────────────────────────────────┬────────────────────────────────┘
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  3. Google Sheets (histórico consolidado)                       │
│     Verifica antiduplicidade e adiciona linhas da nova parcela  │
└────────────────────────────────┬────────────────────────────────┘
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  4. Planilha local consolidada                                  │
│     Baixa todos os dados do Sheets e substitui APENAS a aba     │
│     'e-Gestor' do arquivo:                                      │
│     \\----FS\----\-----\2024\PAINÉIS\IAF\...                    │
│     (preserva abas: sisab, egestor_desat, sisab_bkp)            │
└────────────────────────────────┬────────────────────────────────┘
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  5. eCIEGES (sistema interno - form 472)                        │
│     Login automatizado → botão 'Importar Arquivo' →             │
│     upload da planilha local atualizada                         │
└────────────────────────────────┬────────────────────────────────┘
                                 ▼
                          ✉️ E-mail de status
                       (sucesso / erro / indisponível)
```

---

## 🚀 Funcionalidades Principais

- **Descoberta automática da próxima parcela** — lê o Sheets, identifica a última parcela processada e busca a próxima (ex: se a última é `05/12`, busca `6/12` no site)
- **Compatibilidade de formato** — busca no site com formato natural (`6/12`) e grava no Sheets com zero à esquerda (`06/12`)
- **Detecção de parcela indisponível** — se o governo ainda não publicou, o robô encerra limpo e envia e-mail informativo
- **Antiduplicidade** — normalização de strings (zeros, aspas, espaços) impede gravação de parcelas repetidas
- **Proteção contra conversão automática de data** — injeta aspa simples nas parcelas para o Sheets não confundir `06/12` com data
- **Sincronização Sheets → planilha local** — substitui apenas a aba `e-Gestor` da planilha consolidada, sem afetar outras abas
- **Upload automático no eCIEGES** — login via Material UI, navegação até o formulário 472, importação do XLSX
- **Logging em arquivo** — `logs/robo_sisab.log` registra cada etapa com timestamp
- **Notificação por e-mail (Gmail API)** — envia avisos via HTTPS porta 443, contornando bloqueios de SMTP institucional
- **Execução em background** — `headless=True` permite rodar no Agendador de Tarefas sem interface visível
- **Ano dinâmico** — `datetime.now().year` evita edições manuais a cada virada de ano

---

## 🛠️ Tecnologias Utilizadas

| Biblioteca | Finalidade |
|---|---|
| `playwright` | Automação dos navegadores (e-Gestor e eCIEGES) |
| `gspread` | Integração com Google Sheets API |
| `pandas` / `openpyxl` | Leitura e tratamento de dados XLSX |
| `google-api-python-client` | Envio de e-mails via Gmail API |
| `python-dotenv` | Gestão de variáveis de ambiente |
| `logging` | Registro de eventos em arquivo |

---

## 📁 Estrutura do Repositório

```
Robo_SISAB/
├── automacao_sistema.py    # Script principal com todo o pipeline
├── .env.example            # Modelo das variáveis (copiar para .env)
├── .gitignore              # Proteção de credenciais e arquivos locais
├── requirements.txt        # Dependências do projeto
├── rodar_robo.bat          # Atalho para Agendador de Tarefas Windows
└── README.md               # Esta documentação
```

> **Não estão no repositório** (bloqueados pelo `.gitignore`): `.env`, `credentials.json`, `token.json`, pasta `logs/`, arquivos `.xlsx` baixados.

---

## ⚙️ Configuração Inicial

### 1. Instalar dependências

```powershell
python -m pip install -r requirements.txt
playwright install chromium
```

### 2. Configurar credenciais Google

Coloque `credentials.json` (OAuth Desktop) na raiz do projeto. Na primeira execução o navegador abrirá para autorização — autorize com a conta que tem acesso ao Google Sheets e ao Gmail. O `token.json` será gerado automaticamente.

> **Escopos necessários no Google Cloud:** Google Sheets API + Gmail API. Habilite ambas em `console.cloud.google.com`.

### 3. Configurar variáveis de ambiente

Copie `.env.example` para `.env` e preencha:

```env
# Google Sheets
ID_PLANILHA_SHEETS=seu_id_da_planilha

# Pasta dos relatórios brutos baixados do e-Gestor (ano dinâmico)
PASTA_EGESTOR=\\------\---\-----\{ANO}\IAF\... - Relatórios e-Gestor - pago

# Planilha local consolidada que alimenta o eCIEGES (caminho fixo)
PLANILHA_LOCAL=\\----FS\----\-----\2024\PAINÉIS\---\.xlsx

# eCIEGES
ECIEGES_USUARIO=seu_usuario_ecieges
ECIEGES_SENHA=sua_senha_ecieges

# E-mail (Gmail API — sem senha, usa o mesmo credentials.json)
EMAIL_DESTINATARIO=destinatario@gmail.com
```

---

## ▶️ Execução

**Manual:**
```powershell
python automacao_sistema.py
```

**Automatizada (Agendador de Tarefas do Windows):**
- Configure uma tarefa apontando para `rodar_robo.bat`
- Sugestão: agendar para rodar diariamente em horário fixo (ex: 8h da manhã)
- Quando a parcela do mês ainda não estiver disponível, o robô encerra sozinho e envia e-mail informativo

---

## 📨 Notificações por E-mail

O robô envia e-mails em três situações:

| Situação | Cor do e-mail | Ícone |
|---|---|---|
| ✅ Execução completa com sucesso | Verde | ✅ |
| ⚠️ Parcela ainda não disponível no site | Amarelo | ⚠️ |
| ❌ Erro em qualquer etapa | Vermelho | ❌ |

O envio é feito via **Gmail API (HTTPS porta 443)**, usando o mesmo `credentials.json` do Google Sheets — não precisa de senha SMTP, App Password nem libera firewall.

---

## 🔒 Segurança

- Credenciais nunca são commitadas — todas em `.env` (bloqueado pelo `.gitignore`)
- `credentials.json` e `token.json` também ficam fora do git
- Logs gravados em `logs/` (também ignorados)
- Se credenciais foram expostas antes, **revogue imediatamente** no Google Cloud Console e gere novas

---

## 📊 Frequência Recomendada

- **Agendamento:** diário, em horário comercial
- **Comportamento:** se a parcela do mês já foi processada ou ainda não saiu, o robô encerra sem fazer nada e avisa por e-mail
- **Carga real:** 1x por mês (apenas quando o governo publica nova parcela)