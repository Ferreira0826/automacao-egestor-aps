import os
import time
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

# ==========================================
# 0. CARREGAMENTO BLINDADO DE CREDENCIAIS
# ==========================================
# Descobre exatamente onde este script está salvo e força a leitura do .env na mesma pasta
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, '.env')
load_dotenv(ENV_PATH)

# ==========================================
# 1. CONFIGURAÇÕES DINÂMICAS (Atualize mensalmente aqui)
# ==========================================
PARCELA_ATUAL = "4/12"
MES_REFERENCIA = "FEV26" # Usado para o nome do arquivo
ANO_FILTRO = "2026"

# ==========================================
# 2. CONFIGURAÇÕES FIXAS DO SISTEMA
# ==========================================
URL_DIRETA = "https://relatorioaps.saude.gov.br/gerenciaaps/pagamento/incentivo-atividade-fisica"
ID_PLANILHA_SHEETS = os.getenv("ID_PLANILHA_SHEETS")
PASTA_EGESTOR = os.getenv("PASTA_EGESTOR")
NOME_ABA_SHEETS = "e-Gestor"

# Gera o nome do arquivo automaticamente (ex: Pago parcela 4de12 FEV26.xlsx)
NOME_ARQUIVO_FINAL = f"Pago parcela {PARCELA_ATUAL.replace('/', 'de')} {MES_REFERENCIA}.xlsx"
CAMINHO_COMPLETO = os.path.join(PASTA_EGESTOR, NOME_ARQUIVO_FINAL)

def baixar_relatorio_egestor():
    print(f"Iniciando robô para baixar parcela {PARCELA_ATUAL}...")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False) 
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        
        # Acesso à página principal de relatórios
        print("Acessando página inicial...")
        page.goto("https://relatorioaps.saude.gov.br/")
        page.wait_for_load_state("networkidle")
        
        # Clica no card "Financiamento APS" usando o href que você encontrou
        print("Navegando para Financiamento APS...")
        page.click("a[href='/gerenciaaps/pagamento']")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000) # Aguarda um pouco para a tela de filtros montar
        
        # ==========================================
        # PREENCHIMENTO DOS FILTROS
        # ==========================================
        print("Preenchendo filtros...")
        
        # 1. Filtro Tipo de Unidade
        page.click("span#tipo-unidade") 
        page.wait_for_timeout(1000)
        page.click("li:has-text('Município')")
        
        # ESPERA CRÍTICA: Aguarda 2 segundos para o site "entender" que mudou para Município
        # e liberar a caixa de Estado sem dar aquele erro de Timeout.
        page.wait_for_timeout(2000) 
        
        # 2. Filtro de Estado
        page.click("span#estado")
        page.wait_for_timeout(1000)
        page.click("li:has-text('DISTRITO FEDERAL')")
        
        # ESPERA VITAL: Dá tempo ao sistema para carregar a lista de cidades a partir da base de dados
        page.wait_for_timeout(3000) 
        
        # 3. Filtro de Município
        page.click("span#municipio")
        page.wait_for_timeout(1000)
        # O :visible garante que apenas interage com o texto que já foi renderizado no ecrã
        page.locator("li:has-text('BRASÍLIA'):visible").first.click()
        
        # Filtro de Ano (Usa a variável lá do topo do código)
        page.click("span#ano")
        page.wait_for_timeout(1000)
        page.click(f"li:has-text('{ANO_FILTRO}')")
        
        # 5. Filtro de Parcela Início
        page.click("span#parcela-inicio")
        page.wait_for_timeout(1000)
        # O :visible garante que ele ignore códigos antigos e clique apenas no que está na tela
        page.locator(f"li:has-text('{PARCELA_ATUAL}'):visible").first.click()
        
        # ESPERA ESTRATÉGICA: Tempo para o site recalcular as opções da caixa "Fim"
        page.wait_for_timeout(2000) 
        
        # 6. Filtro de Parcela Fim
        page.click("span#parcela-fim")
        page.wait_for_timeout(1000)
        page.locator(f"li:has-text('{PARCELA_ATUAL}'):visible").first.click()
        
        # ==========================================
        # EXECUÇÃO E NAVEGAÇÃO
        # ==========================================
        print("Executando consulta...")
        page.click("button[aria-label='Ver em tela']")
        page.wait_for_timeout(3000)
        
        # 1. Primeiro clique na seta inicial (A seta principal da UF/Município)
        print("Cobrindo a primeira camada...")
        page.locator("span.pi-chevron-right").first.click()
        page.wait_for_timeout(2500) # Espera maior para a tabela carregar os filhos
        
        # 2. TÁTICA DE VARREDURA: Tentar abrir as duas pastas possíveis
        print("Verificando pastas históricas e atuais...")
        
        # Tentativa 1: Pasta antiga
        try:
            print(" -> Buscando 'Demais programas...'")
            texto_antigo = page.get_by_text("Demais programas, serviços e equipes da Atenção Primária à Saúde", exact=True)
            texto_antigo.wait_for(state="visible", timeout=5000)
            texto_antigo.locator("xpath=ancestor::tr[1]").locator("button").first.click()
            print("    [SUCESSO] Pasta antiga localizada e aberta.")
            page.wait_for_timeout(1500)
        except Exception as e:
            print("    [IGNORADO] Pasta antiga não está na tela.")

        # Tentativa 2: Pasta nova
        try:
            print(" -> Buscando 'Promoção à saúde...'")
            texto_novo = page.get_by_text("Incentivo financeiro da APS - Promoção à saúde", exact=True)
            texto_novo.wait_for(state="visible", timeout=5000)
            texto_novo.locator("xpath=ancestor::tr[1]").locator("button").first.click()
            print("    [SUCESSO] Pasta nova localizada e aberta.")
            page.wait_for_timeout(1500)
        except Exception as e:
            print("    [IGNORADO] Pasta nova não está na tela.")

        # 3. Localizando o botão final de Detalhes
        print("Abrindo detalhes em nova aba...")
        texto_atividade = page.get_by_text("Incentivo de Atividade Física", exact=True)
        
        # Preparamos o robô para capturar a nova aba que vai abrir
        with context.expect_page() as new_page_info:
            texto_atividade.locator("xpath=ancestor::tr[1]").locator("button:has-text('Ver Detalhes')").click()
        
        # Agora o robô "pula" para a aba do relatório
        aba_relatorio = new_page_info.value
        aba_relatorio.wait_for_load_state("networkidle")
        print("Foco alterado para a aba do relatório.")

        # ==========================================
        # DOWNLOAD (NA ABA CORRETA)
        # ==========================================
        sucesso_download = False
        try:
            print("Iniciando extração na aba de detalhes...")
            # Agora procuramos o botão na 'aba_relatorio' e não na 'page'
            botao_dl = aba_relatorio.locator("button[aria-label='Download do excel']").first
            
            with aba_relatorio.expect_download(timeout=60000) as download_info:
                botao_dl.click(force=True) 
                
            download = download_info.value
            download.save_as(CAMINHO_COMPLETO)
            print(f"✅ Download concluído: {CAMINHO_COMPLETO}")
            sucesso_download = True
            
        except Exception as e:
            print(f"❌ Erro ao tentar baixar na nova aba: {e}")

        browser.close()
        return CAMINHO_COMPLETO if sucesso_download else None
    
def enviar_para_sheets(caminho_arquivo, nome_aba):
    print(f"A iniciar a integração para a aba: {nome_aba}...")
    
    # 1. Lê os dados do ficheiro Excel
    df = pd.read_excel(caminho_arquivo)
    
    # Tratamento das colunas N e O para o e-Gestor
    if nome_aba == "e-Gestor":
        try:
            df = df.drop(df.columns[[13, 14]], axis=1)
        except:
            pass

    df = df.fillna("")
    
    # HIGIENIZAÇÃO CRÍTICA: Converte tudo para texto limpo para o Sheets não travar
    valores_para_subir = []
    for linha in df.values.tolist():
        linha_limpa = [str(celula).strip() for celula in linha]
        valores_para_subir.append(linha_limpa)

    # 2. Define o Mês/Ano que o robô está a processar
    mes_ano_processado = f"fev./{ANO_FILTRO}" 

    try:
        gc = gspread.oauth(
            credentials_filename='credentials.json',
            authorized_user_filename='token.json'
        )
        
        planilha = gc.open_by_key(ID_PLANILHA_SHEETS)
        aba = planilha.worksheet(nome_aba)
        
        # ==========================================
        # TRAVA DE SEGURANÇA: VALIDAÇÃO DUPLA (D + E) BLINDADA
        # ==========================================
        print("A verificar se este lote (Mês + Parcela) já existe...")
        
        dados_no_sheets = aba.get_all_values()
        
        # Limpeza extrema do alvo: minúsculas, sem espaços, tira o ponto e tira o zero à esquerda
        mes_ano_alvo = mes_ano_processado.strip().lower().replace(".", "")
        parcela_alvo = str(PARCELA_ATUAL).strip().lower().lstrip("0")
        
        ja_existe = False
        for linha in dados_no_sheets:
            if len(linha) > 4:
                # Limpeza extrema dos dados da folha (idêntico ao alvo)
                coluna_mes = str(linha[3]).strip().lower().replace(".", "")
                coluna_parcela = str(linha[4]).strip().lower().lstrip("0")
                
                # Comparação justa e imune a formatações visuais
                if coluna_mes == mes_ano_alvo and coluna_parcela == parcela_alvo:
                    ja_existe = True
                    break
        
        if ja_existe:
            print(f"⚠️ O lote de {mes_ano_processado} com a parcela {PARCELA_ATUAL} já foi encontrado.")
            print("Envio abortado para evitar duplicidade de dados históricos.")
        else:
            print(f"Lote novo detetado ({mes_ano_processado} - {PARCELA_ATUAL}).")
            print("A adicionar informações ao final da tabela...")
            
            # O USER_ENTERED diz ao Sheets para interpretar a data com as suas próprias regras
            aba.append_rows(valores_para_subir, value_input_mode="USER_ENTERED")
            print("✅ Sucesso! Dados integrados.")
            
            # ==========================================
            # FORMATAÇÃO AUTOMÁTICA: CALIBRI 10
            # ==========================================
            print("A aplicar formatação Calibri 10 na folha...")
            
            # Pega no número total de colunas para saber até onde formatar
            total_colunas = len(df.columns)
            letra_final = gspread.utils.rowcol_to_a1(1, total_colunas).replace("1", "")
            
            # Exemplo: Se for da coluna A até à O, formata de A:O
            intervalo_formatacao = f"A:{letra_final}"
            
            aba.format(intervalo_formatacao, {
                "textFormat": {
                    "fontFamily": "Calibri",
                    "fontSize": 10
                }
            })
            print("✅ Formatação visual aplicada com sucesso.")
            # ==========================================
            
    except Exception as e:
        print(f"Erro na autenticação ou envio: {e}")

if __name__ == "__main__":
    try:
        arquivo_egestor = baixar_relatorio_egestor() 
        
        # O robô só tenta abrir o Google Sheets se a variável tiver o arquivo (não for None)
        if arquivo_egestor:
            enviar_para_sheets(arquivo_egestor, "e-Gestor")
        else:
            print("⚠️ Envio para o Google Sheets cancelado porque o download falhou.")
            
    except Exception as e:
        print(f"Erro na execução geral: {e}")