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
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, '.env')
load_dotenv(ENV_PATH)

# ==========================================
# 1. CONFIGURAÇÕES FIXAS DO SISTEMA
# ==========================================
ANO_FILTRO = "2026" # Só precisará mudar isso em 2027
URL_DIRETA = "https://relatorioaps.saude.gov.br/gerenciaaps/pagamento/incentivo-atividade-fisica"
ID_PLANILHA_SHEETS = os.getenv("ID_PLANILHA_SHEETS")
PASTA_EGESTOR = os.getenv("PASTA_EGESTOR")
NOME_ABA_SHEETS = "e-Gestor"

# ==========================================
# 2. MOTOR DE INTELIGÊNCIA: DESCOBRIR PRÓXIMA PARCELA
# ==========================================
def descobrir_proxima_parcela():
    print("🧠 A ligar ao Google Sheets para descobrir a última parcela processada...")
    
    gc = gspread.oauth(
        credentials_filename='credentials.json',
        authorized_user_filename='token.json'
    )
    
    planilha = gc.open_by_key(ID_PLANILHA_SHEETS)
    aba = planilha.worksheet(NOME_ABA_SHEETS)
    dados_no_sheets = aba.get_all_values()
    
    # Procura de baixo para cima a última parcela válida
    ultima_parcela = "0/12"
    for linha in reversed(dados_no_sheets):
        if len(linha) > 4 and "/" in str(linha[4]):
            ultima_parcela = str(linha[4]).strip()
            break
            
    print(f"📌 Última parcela na planilha: {ultima_parcela}")
    
    # Extrai o número (ex: de "4/12" tira o 4) e soma 1
    numerador_atual = int(ultima_parcela.split('/')[0])
    proximo_numerador = numerador_atual + 1
    
    if proximo_numerador > 12:
        print("🎉 Todas as 12 parcelas deste ano já foram integradas! Nada a fazer.")
        return None, None
        
    parcela_alvo = f"{proximo_numerador}/12"
    print(f"🎯 Alvo automático definido para: {parcela_alvo}")
    
    # Cria um nome dinâmico para o arquivo sem precisar saber o mês de cabeça
    nome_arquivo = f"Relatorio_Parcela_{parcela_alvo.replace('/', 'de')}_{ANO_FILTRO}.xlsx"
    caminho_completo = os.path.join(PASTA_EGESTOR, nome_arquivo)
    
    return parcela_alvo, caminho_completo

# ==========================================
# 3. EXTRAÇÃO DE DADOS (PLAYWRIGHT)
# ==========================================
def baixar_relatorio_egestor(parcela_alvo, caminho_completo):
    print(f"A iniciar o robô para descarregar a parcela {parcela_alvo}...")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False) 
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        
        print("A aceder à página inicial...")
        page.goto("https://relatorioaps.saude.gov.br/")
        page.wait_for_load_state("networkidle")
        
        print("A navegar para Financiamento APS...")
        page.click("a[href='/gerenciaaps/pagamento']")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000) 
        
        print("A preencher filtros...")
        page.click("span#tipo-unidade") 
        page.wait_for_timeout(1000)
        page.click("li:has-text('Município')")
        page.wait_for_timeout(2000) 
        
        page.click("span#estado")
        page.wait_for_timeout(1000)
        page.click("li:has-text('DISTRITO FEDERAL')")
        page.wait_for_timeout(3000) 
        
        page.click("span#municipio")
        page.wait_for_timeout(1000)
        page.locator("li:has-text('BRASÍLIA'):visible").first.click()
        
        page.click("span#ano")
        page.wait_for_timeout(1000)
        page.click(f"li:has-text('{ANO_FILTRO}')")
        
        # Clica na parcela gerada automaticamente
        page.click("span#parcela-inicio")
        page.wait_for_timeout(1000)
        page.locator(f"li:has-text('{parcela_alvo}'):visible").first.click()
        page.wait_for_timeout(2000) 
        
        page.click("span#parcela-fim")
        page.wait_for_timeout(1000)
        page.locator(f"li:has-text('{parcela_alvo}'):visible").first.click()
        
        print("A executar consulta...")
        page.click("button[aria-label='Ver em tela']")
        page.wait_for_timeout(3000)
        
        print("A cobrir a primeira camada...")
        page.locator("span.pi-chevron-right").first.click()
        page.wait_for_timeout(2500) 
        
        print("A verificar pastas históricas e atuais...")
        try:
            texto_antigo = page.get_by_text("Demais programas, serviços e equipes da Atenção Primária à Saúde", exact=True)
            texto_antigo.wait_for(state="visible", timeout=5000)
            texto_antigo.locator("xpath=ancestor::tr[1]").locator("button").first.click()
            page.wait_for_timeout(1500)
        except Exception:
            pass

        try:
            texto_novo = page.get_by_text("Incentivo financeiro da APS - Promoção à saúde", exact=True)
            texto_novo.wait_for(state="visible", timeout=5000)
            texto_novo.locator("xpath=ancestor::tr[1]").locator("button").first.click()
            page.wait_for_timeout(1500)
        except Exception:
            pass

        print("A abrir detalhes num novo separador...")
        texto_atividade = page.get_by_text("Incentivo de Atividade Física", exact=True)
        
        with context.expect_page() as new_page_info:
            texto_atividade.locator("xpath=ancestor::tr[1]").locator("button:has-text('Ver Detalhes')").click()
        
        aba_relatorio = new_page_info.value
        aba_relatorio.wait_for_load_state("networkidle")
        print("Foco alterado para o separador do relatório.")

        sucesso_download = False
        try:
            print("A iniciar extração...")
            botao_dl = aba_relatorio.locator("button[aria-label='Download do excel']").first
            
            with aba_relatorio.expect_download(timeout=60000) as download_info:
                botao_dl.click(force=True) 
                
            download = download_info.value
            download.save_as(caminho_completo)
            print(f"✅ Descarga concluída: {caminho_completo}")
            sucesso_download = True
            
        except Exception as e:
            print(f"❌ Erro ao tentar descarregar: {e}")

        browser.close()
        return caminho_completo if sucesso_download else None

# ==========================================
# 4. INTEGRAÇÃO E FORMATAÇÃO (GOOGLE SHEETS)
# ==========================================
def enviar_para_sheets(caminho_arquivo, nome_aba, parcela_alvo):
    print(f"A iniciar a integração para a aba: {nome_aba}...")
    
    df = pd.read_excel(caminho_arquivo)
    
    if nome_aba == "e-Gestor":
        try:
            df = df.drop(df.columns[[13, 14]], axis=1)
        except:
            pass

    df = df.fillna("")
    
    # ----------------------------------------------------
    # FILTRO DE HIGIENIZAÇÃO E BLOQUEIO DE DATA
    # ----------------------------------------------------
    valores_para_subir = []
    for linha in df.values.tolist():
        linha_limpa = []
        for i, celula in enumerate(linha):
            valor = str(celula).strip()
            
            # Se for a coluna E (Índice 4 no Python) e tiver uma barra
            if i == 4 and "/" in valor:
                try:
                    # Separa o numerador do denominador (ex: "5" e "12")
                    numerador, denominador = valor.split('/')
                    # Força o zero à esquerda (05) e adiciona a aspa simples (') para bloquear datas
                    valor = f"'{int(numerador):02d}/{denominador}"
                except:
                    pass
                    
            linha_limpa.append(valor)
        valores_para_subir.append(linha_limpa)
    # ----------------------------------------------------

    try:
        gc = gspread.oauth(
            credentials_filename='credentials.json',
            authorized_user_filename='token.json'
        )
        
        planilha = gc.open_by_key(ID_PLANILHA_SHEETS)
        aba = planilha.worksheet(nome_aba)
        
        print("A verificar segurança contra duplicidade...")
        dados_no_sheets = aba.get_all_values()
        
        # Limpa o alvo para garantir a comparação (tira zeros e aspas)
        parcela_alvo_limpa = parcela_alvo.strip().lower().replace("'", "").lstrip("0")
        
        ja_existe = False
        for linha in dados_no_sheets:
            if len(linha) > 4:
                # Limpa a linha da folha para comparar de forma justa
                coluna_parcela = str(linha[4]).strip().lower().replace("'", "").lstrip("0")
                if coluna_parcela == parcela_alvo_limpa:
                    ja_existe = True
                    break
        
        if ja_existe:
            print(f"⚠️ A parcela {parcela_alvo} já foi encontrada no histórico.")
            print("Envio abortado para evitar duplicidade.")
        else:
            print(f"A adicionar as informações da parcela {parcela_alvo} ao final da tabela...")
            
            # Usa o USER_ENTERED. Como injetámos a aspa ('), o Sheets não vai converter a parcela!
            aba.append_rows(valores_para_subir, value_input_option="USER_ENTERED")
            print("✅ Sucesso! Dados integrados.")
            
            print("A aplicar formatação Calibri 10 na folha...")
            total_colunas = len(df.columns)
            letra_final = gspread.utils.rowcol_to_a1(1, total_colunas).replace("1", "")
            intervalo_formatacao = f"A:{letra_final}"
            
            aba.format(intervalo_formatacao, {
                "textFormat": {
                    "fontFamily": "Calibri",
                    "fontSize": 10
                }
            })
            print("✅ Formatação visual aplicada com sucesso.")
            
    except Exception as e:
        print(f"Erro na autenticação ou envio: {e}")

# ==========================================
# 5. ORQUESTRAÇÃO FINAL
# ==========================================
if __name__ == "__main__":
    try:
        # Passo 1: O cérebro descobre o que fazer
        parcela_alvo, caminho_completo = descobrir_proxima_parcela()
        
        if parcela_alvo:
            # Passo 2: O braço vai lá e baixa
            arquivo_egestor = baixar_relatorio_egestor(parcela_alvo, caminho_completo) 
            
            # Passo 3: O sistema consolida
            if arquivo_egestor:
                enviar_para_sheets(arquivo_egestor, "e-Gestor", parcela_alvo)
            else:
                print("⚠️ Envio para o Google Sheets cancelado porque a descarga falhou.")
                
    except Exception as e:
        print(f"Erro na execução geral: {e}")