import time
import os
import glob
import win32com.client
import openpyxl
import urllib.parse
import pandas as pd
import re
import unicodedata
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta # Para lidar com datas (hoje e ontem)
from selenium.webdriver.support.ui import Select # Para caixas de <select>
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter
from selenium.webdriver.support.ui import WebDriverWait, Select
from datetime import date, timedelta
from openpyxl import load_workbook


# --- CONFIGURAÇÕES ---
EMAIL_USER = "pedro.henrsilva@mrv.com.br"
SENHA_USER = "Felipe22#" 
WAIT_TIME = 10

# --- FUNÇÃO DE APOIO: LOGIN MICROSOFT ---
def fazer_login_microsoft(driver, wait, email, senha):
    """Lida com o login da MS. Retorna True se logou, False se der erro."""
    print("--- Iniciando rotina de Login Microsoft ---")
    try:
        try:
            email_field = wait.until(EC.presence_of_element_located((By.ID, "i0116")))
            print("Preenchendo e-mail...")
            email_field.send_keys(email)
            wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
            
            password_field = wait.until(EC.presence_of_element_located((By.ID, "i0118")))
            print("Preenchendo senha...")
            password_field.send_keys(senha)
            
            clicked = False
            for _ in range(3):
                try:
                    wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
                    clicked = True
                    break
                except StaleElementReferenceException:
                    time.sleep(1)
            if not clicked: raise Exception("Não clicou em Entrar")

            print("!!! AGUARDANDO APROVAÇÃO MFA (Se necessário) !!!")
            wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click() 
            print("Login Microsoft efetuado.")
        except TimeoutException:
            print("Campo de login não apareceu. Assumindo que já estamos logados (SSO).")
        return True
    except Exception as e:
        print(f"Erro no Login Microsoft: {e}")
        return False

# --- INÍCIO DO SCRIPT ---
try:
    driver = webdriver.Chrome()
    driver.maximize_window()
    wait = WebDriverWait(driver, WAIT_TIME)
    '''
    # ==============================================================================
    # PARTE 1: PODIO
    # ==============================================================================
    print("\n=== INICIANDO PARTE 1: PODIO ===")
    driver.get("https://podio.com/login")

    try:
        wait.until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))).click()
    except: pass 

    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@data-provider='live']"))).click()

    janela_principal = driver.current_window_handle
    wait.until(EC.number_of_windows_to_be(2))
    
    for handle in driver.window_handles:
        if handle != janela_principal:
            driver.switch_to.window(handle)
            break

    fazer_login_microsoft(driver, wait, EMAIL_USER, SENHA_USER)
    driver.switch_to.window(janela_principal)
    
    print("Navegando no Podio...")
    menu_area = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'space-switcher-wrapper')]")))
    ActionChains(driver).move_to_element(menu_area).perform()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'ADM - Núcleo Contratos')]"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-app-id='22830484']"))).click()

    # --- CORREÇÃO DO FILTRO PODIO ---
    print("Aplicando filtros (Método Robusto)...")
    time.sleep(3) # Espera script da página carregar
    
    # 1. Pega o container UL
    ul_filtros = wait.until(EC.presence_of_element_located((By.XPATH, "//ul[@class='app-filter-tools']")))
    
    # 2. Pega todos os itens LI dentro dele
    itens_lista = ul_filtros.find_elements(By.TAG_NAME, "li")
    
    # 3. Passa o mouse em CADA item para garantir que o menu acorde
    actions = ActionChains(driver)
    for item in itens_lista:
        actions.move_to_element(item)
    actions.perform()
    
    # 4. Agora clica no filtro
    wait.until(EC.element_to_be_clickable((By.XPATH, ".//li[@data-original-title='Filtros']"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-id='created_on']"))).click() 
    wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-id='-1mr:-1mr']"))).click() 

    # Seleção de Views
    print("Ajustando visualização...")
    elementos_menu = wait.until(lambda d: d.find_elements(By.CSS_SELECTOR, ".app-header__app-menu"))
    if len(elementos_menu) >= 1: elementos_menu[0].click()
    time.sleep(2)
    elementos_menu = driver.find_elements(By.CSS_SELECTOR, ".app-header__app-menu") 
    if len(elementos_menu) > 1: elementos_menu[1].click()
    time.sleep(2)

    print("Exportando Excel...")
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.app-box-supermenu-v2__link.app-export-excel"))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "li.navigation-link.inbox"))).click()

    time.sleep(20)
    
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.PodioUI__Notifications__NotificationGroup"))).click()
    wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'field-type-text')]"))) 
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Mensageria - Última vista usada.xlsx"))).click()
    print("Download Podio iniciado!")
    time.sleep(5) 
    
    # ==============================================================================
    # PARTE 2: AGILIS
    # ==============================================================================
    print("\n=== INICIANDO PARTE 2: AGILIS ===")
    driver.get("https://agilis.mrv.com.br/HomePage.do?view_type=my_view")

    try:
        btn_login = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Login Integrado Microsoft']")))
        btn_login.click()
        fazer_login_microsoft(driver, wait, EMAIL_USER, SENHA_USER)
    except TimeoutException:
        print("Botão de login não apareceu, seguindo...")

    print("Navegando menus Agilis...")
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Relatórios"))).click()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Contratos - ADM"))).click()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Produtividade Contratos - ADM"))).click()
    
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "linkborder"))).click() 
    wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Coletor de custo ADM']"))).click()
    driver.find_element(By.CLASS_NAME, "moverightButton").click()

    try:
        expand_btn = wait.until(EC.presence_of_element_located((By.ID, "rcstep2src")))
        driver.execute_script("arguments[0].click();", expand_btn)
        time.sleep(1)
    except: pass

    # --- SELEÇÃO DE DATA (DROPDOWN + RÁDIO) ---
    print("Selecionando Data 'Mês Passado' no dropdown...")
    select_elem = wait.until(EC.presence_of_element_located((By.ID, "dateFilterType")))
    Select(select_elem).select_by_visible_text("Mês passado")
    
    # [INSERIDO] Código do Rádio 'Durante' que você pediu
    print("Selecionando o rádio 'Durante' (Ajuste obrigatório)...")
    try:
        selector_radio_durante = (By.CSS_SELECTOR, "input[value='predefined']")
        wait.until(EC.element_to_be_clickable(selector_radio_durante)).click()
        print(" - SUCESSO: Filtro 'Durante' selecionado.")
    except TimeoutException:
        print(" - FALHA: Rádio 'Durante' não encontrado.")
        raise

    # Executar Relatório
    wait.until(EC.element_to_be_clickable((By.ID, "addnew223222"))).click() 
    print("Relatório gerando. Aguardando 10 segundos...")
    time.sleep(10) 

    # --- 7. Baixar Relatório XLS Diretamente ---
    print("7. Iniciando o download direto do relatório XLS...")
    try:
        # Localizar o link de exportação pelo ID "exportxls" e clicar nele
        DOWNLOAD_XLS_LINK = (By.ID, "exportxls")
        wait.until(EC.element_to_be_clickable(DOWNLOAD_XLS_LINK)).click()
        print("   - Clique realizado no link 'Exportar arquivo como XLS'.")
    
        # IMPORTANTE: Adicionar uma pausa para o download começar e terminar.
        # A melhor abordagem é verificar a pasta de downloads até o arquivo aparecer.
        # Veja a explicação abaixo sobre como fazer isso.
        print("   - Aguardando o download ser concluído...")
        time.sleep(5) # Pausa simples de 15 segundos. O ideal é usar uma função de verificação.
    
        print("   - Relatório baixado com sucesso!")

    except Exception as e:
        print(f"ERRO ao tentar baixar o relatório XLS: {e}")
        # Adicione aqui o tratamento de erro
    
    print("Relatório Agilis baixado com sucesso!")
    time.sleep(5)
    
'''
    # ==============================================================================
    # PARTE 3: BÚSSOLA MRV
    # ==============================================================================
    print("Passo 1, abrir o site")

    # --- PASSO 1: Login no domínio principal (como você já faz) ---
    usuario_safe = urllib.parse.quote(EMAIL_USER)
    senha_safe = urllib.parse.quote(SENHA_USER)

    # URL autenticada para o Bússola
    url_autenticada_bussola = f"http://{usuario_safe}:{senha_safe}@bussola.mrv.com.br/Main/Big.aspx"

    print("Acessando Bússola com credenciais embutidas...")
    driver.get(url_autenticada_bussola)

    # --- PASSO 2: PRÉ-AUTENTICAÇÃO NO DOMÍNIO DO RELATÓRIO (NOVO) ---
    print("Pré-autenticando no domínio 'http://report2.mrv.com.br/ReportServer/Pages/ReportViewer.aspx?/BIG/Administrativo/ADM013%20-%20Relat%C3%B3rio%20Protocolo%20de%20Pagamento%20MRV%20PAG/REL_PRLPGTMRV&rs:Command=Render'...")

    # Monta uma URL base para o domínio do relatório com as credenciais
    url_autenticada_report = f"http://{usuario_safe}:{senha_safe}@report2.mrv.com.br/ReportServer/Pages/ReportViewer.aspx?/BIG/Administrativo/ADM013%20-%20Relat%C3%B3rio%20Protocolo%20de%20Pagamento%20MRV%20PAG/REL_PRLPGTMRV&rs:Command=Render"

    # Visita a URL base. O navegador vai processar e guardar as credenciais para este domínio.
    # A página pode dar erro ou ficar em branco, não importa. O objetivo é apenas enviar as credenciais.
    driver.get(url_autenticada_report)

    # --- PASSO 3: Volte para o Bússola para continuar a navegação ---
    print("Retornando ao Bússola para continuar a automação...")
    driver.get(url_autenticada_bussola) # Ou use driver.back() se a página anterior for a correta

    # --- PASSO 4: Continue seu script normalmente ---
    # Agora o navegador já está autenticado em ambos os domínios.
    # O código para clicar nas pastas e no relatório funcionará sem o pop-up.

    print("Procurando pasta 'Administrativo'...")
    # Se o login funcionar, essa pasta vai aparecer.
    pasta_adm = wait.until(EC.element_to_be_clickable((By.ID, "pasta2"))) # Usando o ID que corrigimos antes
    pasta_adm.click()
    time.sleep(2)

    print("Selecionando o relatório...")
    xpath_relatorio = "//div[@id='divLinha' and contains(., 'Relatório Protocolo de Pagamento MRV PAG')]"
    relatorio_link = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_relatorio)))
    relatorio_link.click()

    # Agora o relatório deve abrir diretamente, sem pedir login.
    print("Relatório aberto com sucesso!")
 

    hoje = date.today()
    primeiro_dia_mes_atual = hoje.replace(day=1)
    ultimo_dia_mes_passado = primeiro_dia_mes_atual - timedelta(days=1)
    primeiro_dia_mes_passado = ultimo_dia_mes_passado.replace(day=1)

    # Strings só para eventual fallback/validação (formato dd/mm/aaaa)
    str_inicio = primeiro_dia_mes_passado.strftime("%d/%m/%Y")
    str_fim    = ultimo_dia_mes_passado.strftime("%d/%m/%Y")
    print(f"Período a ser preenchido via calendário: {str_inicio} a {str_fim}")

    def _espera_datepicker(wait, timeout=10):
        """Espera o pop-up do calendário aparecer."""
        return wait.until(EC.visibility_of_element_located((
            By.XPATH,
            # contêiner do datepicker (varia entre versões, cobrimos alguns padrões)
            "//*[contains(@class,'ms-picker') or contains(@class,'datepicker') or contains(@class,'ms-datepicker')]"
        )))

    def _clica_prev_mes(datepicker_root, driver):
        """Clica na seta para voltar um mês no calendário."""
        # cobrindo variações comuns de seta/âmbito da seta (título/aria-label/ícone)
        candidatos = [
            ".//a[contains(@title,'Anterior') or contains(@aria-label,'Anterior')]",
            ".//a[contains(@class,'prev') or contains(.,'‹') or contains(.,'«')]",
            ".//span[contains(@class,'prev') or contains(.,'‹') or contains(.,'«')]",
        ]
        for xp in candidatos:
            elems = datepicker_root.find_elements(By.XPATH, xp)
            if elems:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elems[0])
                try:
                    elems[0].click()
                except Exception:
                    driver.execute_script("arguments[0].click();", elems[0])
                return True
        raise TimeoutException("Seta de mês anterior do datepicker não encontrada.")

    def _clica_dia(datepicker_root, driver, dia):
        """
        Seleciona um dia no calendário visível.
        - Evita dias de outros meses (classe 'dayother').
        - Considera que o número pode estar no próprio <td> ou dentro de <a>.
        """
        # Primeiro tentamos <td> com texto, excluindo 'dayother'
        xpath_opcoes = [
            f".//td[contains(@class,'ms-picker-day') and not(contains(@class,'dayother'))][normalize-space()='{dia}']",
            f".//td[contains(@class,'ms-picker-day') and not(contains(@class,'dayother'))]//a[normalize-space()='{dia}']",
            # fallback genérico
            f".//*[self::td or self::a][normalize-space()='{dia}' and not(contains(@class,'dayother'))]"
        ]

        for xp in xpath_opcoes:
            els = datepicker_root.find_elements(By.XPATH, xp)
            if els:
                el = els[0]
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                try:
                    el.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", el)
                return True

        # Último recurso: procurar pelo 'onclick' que contém o dia (SSRS costuma usar '01\\u002f12\\u002f2025')
        # Aqui apenas tentamos pelo número do dia, sem fixar o mês/ano, porque já navegamos p/ mês anterior.
        xp_onclick = f".//td[contains(@class,'ms-picker-day') and contains(@onclick, \"'{dia}\\u002f\")]"
        els = datepicker_root.find_elements(By.XPATH, xp_onclick)
        if els:
            el = els[0]
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            try:
                el.click()
            except Exception:
                driver.execute_script("arguments[0].click();", el)
            return True

        raise TimeoutException(f"Não consegui selecionar o dia {dia} no calendário.")

    def selecionar_mes_anterior_dia(aria_label_botao, dia):
        """
        Abre o calendário do parâmetro identificado por aria-label do botão,
        volta um mês e seleciona o 'dia' informado.
        """
        print(f"Abrindo calendário: {aria_label_botao} (dia {dia})")
        # Botão do calendário (o da imagem tem aria-label 'Data criação inicio'/'Data criação final')
        botao_cal = wait.until(EC.element_to_be_clickable((
            By.XPATH, f"//button[@aria-label='{aria_label_botao}' and contains(@class,'glyphui-calendar')]"
        )))
        # garantir visibilidade e clique estável
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", botao_cal)
        try:
            botao_cal.click()
        except Exception:
            driver.execute_script("arguments[0].click();", botao_cal)

        # aguarda aparecer o datepicker
        datepicker = _espera_datepicker(wait)

        # por padrão o SSRS abre no mês corrente → precisamos retroceder uma vez
        _clica_prev_mes(datepicker, driver)

        # escolhe o dia desejado
        _clica_dia(datepicker, driver, dia)

        # aguarda o datepicker sumir (indica seleção concluída)
        try:
            wait.until(EC.invisibility_of_element(datepicker))
        except Exception:
            pass  # alguns temas apenas recolhem, mas seguem visíveis; não é crítico

    # === 2) GARANTIR QUE ESTAMOS DENTRO DO IFRAME DO RELATÓRIO ===
    print("Procurando o iframe do relatório e mudando o foco...")
    iframe_relatorio = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//iframe[contains(@src, 'ReportServer')]")
    ))
    driver.switch_to.frame(iframe_relatorio)
    print("Foco alterado para o iframe com sucesso.")

    # === 3) Selecionar datas via calendário ===
    # Primeiro calendário: 1º dia do mês anterior
    selecionar_mes_anterior_dia("Data criação inicio", 1)

    # Segundo calendário: último dia do mês anterior
    ultimo_dia_num = int(ultimo_dia_mes_passado.strftime("%d"))
    selecionar_mes_anterior_dia("Data criação final", ultimo_dia_num)

    print("Datas selecionadas com sucesso pelos calendários.")

        
    # 5. Clicar em "Exibir Relatório"
    print("Gerando relatório...")
    # Procura o botão pelo valor (value)
    btn_exibir = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='Exibir Relatório']")))
    btn_exibir.click()
 
    # 6. Exportar para Excel
    print("Aguardando ícone de exportação (disquete)...")
    
    
    try:
        # Busca o botão de salvar (disquete) de várias formas possíveis
        botao_exportar = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@title='Exportar menu suspenso'] | //img[@alt='Exportar'] | //a[contains(@id, 'Export')]")))
        botao_exportar.click()
        
        print("Selecionando Excel...")
        opcao_excel = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@title='Excel'] | //a[text()='Excel']")))
        opcao_excel.click()
        
        print("Download Bússola iniciado!")
        time.sleep(10)
 
    except TimeoutException:
        print("Erro: O botão de exportar não apareceu a tempo.")

    # Opcional: Fechar a janela do relatório e voltar para a principal
    # driver.close()
    # driver.switch_to.window(janela_bussola)
    
except Exception as e:
    print(f"\nCRITICAL ERROR NA PARTE 3 (BÚSSOLA): {e}")
    driver.save_screenshot("erro_bussola.png")
    

finally:
    print("Script Completo Finalizado.")
    # driver.quit()
    print("Fim.")
''' 
#------------------------------------------------------------------------------------------
time.sleep(30)
'''

# --- FUNÇÕES AUXILIARES GERAIS ---

def find_column_ignore_case(df, column_name):
    """Encontra o nome real de uma coluna, ignorando maiúsculas/minúsculas."""
    for col in df.columns:
        if col.lower() == column_name.lower():
            return col
    return None

# --- ETAPA 1: FUNÇÕES PARA Renomear e editar as planilhas ---
'''
def processar_mensageria(filepath, new_filename):
    try:
        print(f"Processando MENSAGERIA: {filepath}")
        df = pd.read_excel(filepath, header=0) 
        df.columns = [str(col).strip() for col in df.columns]
        coluna_usuario = find_column_ignore_case(df, 'Criado por')
        coluna_data = find_column_ignore_case(df, 'Criado em')
        coluna_valores = find_column_ignore_case(df, 'Numero do chamado Agilis/Rastreio')
        df[coluna_data] = pd.to_datetime(df[coluna_data], dayfirst=True).dt.date
        pivot_table = pd.pivot_table(df, index=coluna_usuario, columns=coluna_data, values=coluna_valores, aggfunc='count', fill_value=0, margins=True, margins_name='Total Geral')
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            pivot_table.to_excel(writer, sheet_name='TabelaDinamica')
        os.rename(filepath, new_filename)
        print(f" -> Sucesso! Renomeado para '{new_filename}'")
    except Exception as e: print(f" -> ERRO (Etapa 1 - Mensageria): {e}")

def processar_produtividade(filepath, new_filename):
    try:
        print(f"Processando PRODUTIVIDADE: {filepath}")
        df = pd.read_excel(filepath, header=0)
        df.columns = [str(col).strip() for col in df.columns]
        coluna_usuario = find_column_ignore_case(df, 'Nome do usuário')
        coluna_data = find_column_ignore_case(df, 'Data de lançamento')
        coluna_valores = find_column_ignore_case(df, 'Nº doc.faturamento')
        df[coluna_data] = pd.to_datetime(df[coluna_data], dayfirst=True).dt.date
        pivot_table = pd.pivot_table(df, index=coluna_usuario, columns=coluna_data, values=coluna_valores, aggfunc='count', fill_value=0, margins=True, margins_name='Total Geral')
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            pivot_table.to_excel(writer, sheet_name='TabelaDinamica')
        os.rename(filepath, new_filename)
        print(f" -> Sucesso! Renomeado para '{new_filename}'")
    except Exception as e: print(f" -> ERRO (Etapa 1 - Produtividade): {e}")

def processar_numerico(filepath, new_filename):
    try:
        print(f"Processando ARQUIVO NUMÉRICO: {filepath}")
        df = pd.read_excel(filepath, header=8)
        df.columns = [str(col).strip() for col in df.columns]
        coluna_tecnico = find_column_ignore_case(df, 'Técnico')
        coluna_data = find_column_ignore_case(df, 'Hora de conclusão')
        coluna_valores = find_column_ignore_case(df, 'Identificação da solicitação')
        df[coluna_data] = pd.to_datetime(df[coluna_data], dayfirst=True).dt.date
        pivot_table = pd.pivot_table(df, index=coluna_tecnico, columns=coluna_data, values=coluna_valores, aggfunc='count', fill_value=0, margins=True, margins_name='Total Geral')
        with pd.ExcelWriter(new_filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='DadosOriginais', index=False)
            pivot_table.to_excel(writer, sheet_name='TabelaDinamica')
        os.remove(filepath)
        print(f" -> Sucesso! Novo arquivo '{new_filename}' criado.")
    except Exception as e: print(f" -> ERRO (Etapa 1 - Numérico): {e}")

def processar_relatorio_pedidos(filepath, new_filename):
    try:
        print(f"Processando RELATÓRIO DE PEDIDOS: {filepath}")
        df = pd.read_excel(filepath, header=1)
        df.columns = [str(col).strip() for col in df.columns]
        coluna_linhas = find_column_ignore_case(df, 'Respons. Entrega')
        coluna_colunas = find_column_ignore_case(df, 'Data Entrada NF')
        coluna_valores = find_column_ignore_case(df, 'Nro. Pedido Compra')
        df[coluna_colunas] = pd.to_datetime(df[coluna_colunas], dayfirst=True).dt.date
        pivot_table = pd.pivot_table(df, index=coluna_linhas, columns=coluna_colunas, values=coluna_valores, aggfunc='count', fill_value=0, margins=True, margins_name='Total Geral')
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            pivot_table.to_excel(writer, sheet_name='TabelaDinamica')
        os.rename(filepath, new_filename)
        print(f" -> Sucesso! Renomeado para '{new_filename}'")
    except Exception as e: print(f" -> ERRO (Etapa 1 - Pedidos): {e}")

def step_1_prepare_and_rename_reports(diretorio):
    arquivos = glob.glob(os.path.join(diretorio, '*.*'))
    for arquivo in arquivos:
        nome_arquivo = os.path.basename(arquivo)
        if nome_arquivo.startswith('Mensageria - Última vista'):
            processar_mensageria(arquivo, 'Relatório - Sedex.Malote.xlsx')
        elif nome_arquivo.startswith('export') and nome_arquivo.endswith('.xlsx'):
            processar_produtividade(arquivo, 'Relatório - SAP.xlsx')
        elif nome_arquivo.startswith('REL_PRLPGT'):
            processar_relatorio_pedidos(arquivo, 'Relatório - Lançamentos.xlsx')
        elif re.match(r'^\d+\.(xlsx|xls)$', nome_arquivo):
            processar_numerico(arquivo, 'Relatório - Agilis.xlsx')
'''

'''
# --- ETAPA 2: FUNÇÕES PARA ATUALIZAR A PLANILHA PRINCIPAL (COM novissima LÓGICA) ---

# Arquivos e abas

PROD_PATH    = "Produtividade 01 - 2026.xlsx"      # aba Plan1
AGILIS_PATH  = "Relatório - Agilis.xlsx"           # aba TabelaDinamica
SEDEX_PATH   = "Relatório - Sedex.Malote.xlsx"     # aba TabelaDinamica
LANCTOS_PATH = "Relatório - Lançamentos.xlsx"      # aba TabelaDinamica (Respons. Entrega)
SAP_PATH     = "Relatório - SAP.xlsx"              # aba TabelaDinamica (Nome do usuário)
OUT_PATH     = "Produtividade 01 - 2026 (preenchido).xlsx"

# ===================== MAPEAMENTOS =====================
# Sedex: nome na pivot -> texto exato da coluna B na produtividade
MAP_SEDEX = {
    "Alfredo Henrique Goncalves Pereira": "Alfredo.pereira MS0069532",
    "Gabriel Figueiredo Emiliano":        "gabriel.emiliano MS0073186",
    "Pedro Henrique Soares Silva":        "pedro.henrsilva MS0073814",
    "Ezequiel Viana Ferreira":            "ezequiel.ferreira",
}

# Agilis: MESMA linha do nome (coluna C), respeitando coluna mínima
AGILIS_POS = [
    {"p2": "Alfredo Henrique Goncalves Pereira", "p1": "Alfredo.pereira MS0069532",  "row_nome":  2, "min_col_letter": "CO"},
    {"p2": "Gabriel Figueiredo Emiliano",        "p1": "gabriel.emiliano MS0073186", "row_nome":  7, "min_col_letter": "CO"},
    {"p2": "Ezequiel Viana Ferreira",            "p1": "ezequiel.ferreira",          "row_nome": 12, "min_col_letter": "CO"},
    {"p2": "arthur.savio",                       "p1": "arthur.savio",               "row_nome": 17, "min_col_letter": "CO"},
    {"p2": "Pedro Henrique Soares Silva",        "p1": "pedro.henrsilva MS0073814",  "row_nome": 22, "min_col_letter": "CO"},
    {"p2": "Camilly Cristine Dos Santos",        "p1": "camilly.santos",             "row_nome": 27, "min_col_letter": "CO"},
    {"p2": "Carolina Pagnozzi Silva",            "p1": "pagnozzi.carolina",          "row_nome": 32, "min_col_letter": "CO"},
    {"p2": "Matheus Silva De Lemos",             "p1": "matheus.lemos.silva",        "row_nome": 37, "min_col_letter": "CO"},
    {"p2": "Vanessa De Brito Rodrigues",         "p1": "Vanessa",                    "row_nome": 42, "min_col_letter": "C"},  # Vanessa desde C
]

# Lançamentos 45 E 19: linhas fixas

LANCTOS_USER_MAP = {
    "alfredo.pereira":      {"p1": "Alfredo.pereira MS0069532",  "row_ativ":  4},
    "gabriel.emiliano":     {"p1": "gabriel.emiliano MS0073186", "row_ativ":  9},
    "ezequiel.ferreira":    {"p1": "ezequiel.ferreira",          "row_ativ": 14},
    "arthur.savio":         {"p1": "arthur.savio",               "row_ativ": 19},
    "pedro.henrsilva":      {"p1": "pedro.henrsilva MS0073814",  "row_ativ": 24},
    "camilly.santos":       {"p1": "camilly.santos",             "row_ativ": 29},
    "pagnozzi.carolina":    {"p1": "pagnozzi.carolina",          "row_ativ": 34},
    "matheus.lemos.silva":  {"p1": "matheus.lemos.silva",        "row_ativ": 39},
}

ATIV_LANCTOS_LABEL = "Lançamentos 45 E 19"

# SAPMIRO: códigos MS -> colaborador + linha fixa
SAP_COD_MAP = {
    "MS0069532": {"p1": "Alfredo.pereira MS0069532",  "row_ativ":  5},
    "MS0073186": {"p1": "gabriel.emiliano MS0073186", "row_ativ": 10},
    "MS0073814": {"p1": "pedro.henrsilva MS0073814",  "row_ativ": 25},
}
ATIV_SAP_LABEL = "SAPMIRO"

# ===================== FUNÇÕES AUX =====================
def norm_key(s):
    if s is None: return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    n = 0
    for ch in letter:
        if not ('A' <= ch <= 'Z'): continue
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

def date_keys(v):
    keys = []
    if isinstance(v, datetime):
        keys += [v.strftime("%Y-%m-%d 00:00:00"), v.strftime("%d/%m/%Y")]
    elif isinstance(v, str):
        s = v.strip()
        if re.match(r"^\d{4}-\d{2}-\d{2} 00:00:00$", s): keys.append(s)
        if re.match(r"^\d{2}/\d{2}/\d{4}$", s): keys.append(s)
        try:
            dt = pd.to_datetime(s, errors="raise")
            keys += [dt.strftime("%Y-%m-%d 00:00:00"), dt.strftime("%d/%m/%Y")]
        except: pass
    return list(dict.fromkeys(keys))

def build_header_map(ws):
    hdr = {}
    for c in range(1, ws.max_column+1):
        v = ws.cell(row=1, column=c).value
        if v is None: continue
        if isinstance(v, datetime):
            hdr[v.strftime("%Y-%m-%d 00:00:00")] = c
            hdr[v.strftime("%d/%m/%Y")] = c
        else:
            s = str(v).strip()
            hdr[s] = c
            for k in date_keys(s): hdr[k] = c
    return hdr

def extract_user_key(s: str) -> str:
    """
    Extrai o 'user key' no formato que você especificou:
    - remove domínio/email (parte depois de '@')
    - usa apenas letras e pontos
    - corta em espaço (primeira palavra)
    - normaliza para minúsculas sem acentos
    Ex.: 'Alfredo.Pereira ' -> 'alfredo.pereira'
         'gabriel.emiliano@mrv.com' -> 'gabriel.emiliano'
    """
    if s is None:
        return ""
    s = str(s).strip()
    # corta em '@' (se vier email)
    s = s.split('@')[0]
    # pega a primeira "palavra" até espaço
    s = s.split()[0]
    # mantém apenas letras e pontos
    s = re.sub(r'[^A-Za-z\.]', '', s)
    # normaliza e baixa caixa
    return norm_key(s)

def read_tabledinamica_with_namecol(path, name_col_hint=None):
    
    """
    Lê a aba 'TabelaDinamica' do arquivo 'path' e retorna DataFrame long com colunas:
      ['nome','data_obj','valor'].
    Se 'name_col_hint' for fornecido, usa exatamente essa coluna (normalização tolerante).
    """
    xls = pd.ExcelFile(path, engine="openpyxl")
    if "TabelaDinamica" not in xls.sheet_names:
        raise RuntimeError(f"Aba 'TabelaDinamica' não encontrada em {path}. Abas: {xls.sheet_names}")
    df = pd.read_excel(path, sheet_name="TabelaDinamica", engine="openpyxl")

    # Escolher a coluna de nomes
    nome_col = None
    if name_col_hint:
        # casar com tolerância (strip/lower e sem acento)
        hint_norm = norm_key(name_col_hint)
        for c in df.columns:
            if norm_key(c) == hint_norm:
                nome_col = c
                break
    if not nome_col:
        # fallback: tenta cabeçalhos comuns
        heads = {norm_key(c): c for c in df.columns}
        for alvo in ["criado por", "tecnico", "técnico", "respons. entrega", "nome do usuário"]:
            if alvo in heads:
                nome_col = heads[alvo]
                break
    if not nome_col:
        # fallback final: primeira coluna de strings
        for c in df.columns:
            if df[c].dtype == "O":
                nome_col = c
                break

    # Colunas de data
    day_cols = []
    for c in df.columns:
        if c == nome_col: continue
        if isinstance(c, datetime):
            day_cols.append(c)
        elif isinstance(c, str):
            s = c.strip().lower()
            if s in ("total", "total geral"): continue
            if re.match(r"^\d{2}/\d{2}/\d{4}$", c) or re.match(r"^\d{4}-\d{2}-\d{2}", c):
                day_cols.append(c)

    registros = []
    if nome_col:
        for _, row in df.iterrows():
            nome_val = str(row.get(nome_col, "")).strip()
            if not nome_val or norm_key(nome_val).startswith(norm_key("Total Geral")):
                continue
            for d in day_cols:
                val = row.get(d, 0)
                try: v = int(val) if pd.notna(val) else 0
                except Exception: v = 0
                registros.append({"nome": nome_val, "data_obj": d, "valor": v})

    return pd.DataFrame(registros, columns=["nome","data_obj","valor"])

def read_lanctos_tabledinamica(path):
    """
    Lê a aba 'TabelaDinamica' do Relatório - Lançamentos.xlsx
    usando a coluna 'Respons. Entrega' como nome-base e
    devolve DF long: ['user_key','data_obj','valor'].
    """
    xls = pd.ExcelFile(path, engine="openpyxl")
    if "TabelaDinamica" not in xls.sheet_names:
        raise RuntimeError(f"Aba 'TabelaDinamica' não encontrada em {path}. Abas: {xls.sheet_names}")
    df = pd.read_excel(path, sheet_name="TabelaDinamica", engine="openpyxl")

    # localizar a coluna 'Respons. Entrega' com tolerância
    nome_col = None
    for c in df.columns:
        if norm_key(c) == norm_key("Respons. Entrega"):
            nome_col = c
            break
    if nome_col is None:
        raise RuntimeError("Coluna 'Respons. Entrega' não encontrada na TabelaDinamica de Lançamentos.")

    # identificar colunas de data
    day_cols = []
    for c in df.columns:
        if c == nome_col: continue
        if isinstance(c, datetime):
            day_cols.append(c)
        elif isinstance(c, str):
            s = c.strip().lower()
            if s in ("total", "total geral"): continue
            if re.match(r"^\d{2}/\d{2}/\d{4}$", c) or re.match(r"^\d{4}-\d{2}-\d{2}", c):
                day_cols.append(c)

    registros = []
    for _, row in df.iterrows():
        raw_name = str(row.get(nome_col, "")).strip()
        if not raw_name or norm_key(raw_name).startswith(norm_key("Total Geral")):
            continue
        ukey = extract_user_key(raw_name)
        for d in day_cols:
            val = row.get(d, 0)
            try: v = int(val) if pd.notna(val) else 0
            except: v = 0
            registros.append({"user_key": ukey, "data_obj": d, "valor": v})

    return pd.DataFrame(registros, columns=["user_key","data_obj","valor"])

# ===================== PREENCHIMENTOS =====================
def fill_agilis_same_row(ws, header_map, df_long):
    """Agilis na MESMA linha do nome (coluna C), respeitando coluna mínima e sem criar colunas."""
    if df_long.empty:
        print("[AGILIS] Pivot vazia ou sem linhas válidas.")
        return 0

    grp = {norm_key(n): sub for n, sub in df_long.groupby(df_long['nome'].apply(norm_key))}
    total_writes = 0

    for item in AGILIS_POS:
        p2 = item["p2"]
        p1 = item["p1"]
        row_nome = item["row_nome"]
        row_agilis = row_nome              # MESMA LINHA
        min_idx = col_letter_to_index(item["min_col_letter"])

        # Se a coluna mínima não existir, desativa limite
        if min_idx > ws.max_column:
            print(f"[AGILIS] {p1}: coluna mínima {item['min_col_letter']} (idx {min_idx}) > max {ws.max_column}. Desativando limite.")
            min_idx = 1

        sub = grp.get(norm_key(p2))
        if sub is None or sub.empty:
            print(f"[AGILIS] {p1}: sem registros (deixando em branco).")
            continue

        writes = 0
        for _, reg in sub.iterrows():
            col_idx = None
            for k in date_keys(reg["data_obj"]):
                c = header_map.get(k)
                if c:
                    col_idx = c; break
            if not col_idx or col_idx < min_idx: continue

            val = int(reg["valor"])
            if val != 0:  # zeros em branco
                ws.cell(row=row_agilis, column=col_idx, value=val); writes += 1

        total_writes += writes
        print(f"[AGILIS] {p1} (linha {row_agilis}, min {item['min_col_letter']}): {writes} células escritas.")
    return total_writes

def fill_sedex(ws, header_map, df_long):
    """Sedex/Pac/Malote: encontra atividade na coluna C abaixo do nome e preenche."""
    if df_long.empty:
        print("[SEDEX] Pivot vazia ou sem linhas válidas.")
        return 0

    grp = {norm_key(n): sub for n, sub in df_long.groupby(df_long['nome'].apply(norm_key))}
    total_writes = 0

    for p2, p1 in MAP_SEDEX.items():
        # localizar linha do nome (col B)
        r_nome = None
        for r in range(2, ws.max_row+1):
            v = ws.cell(row=r, column=2).value
            if isinstance(v, str) and norm_key(v) == norm_key(p1):
                r_nome = r; break
        if r_nome is None:
            print(f"[SEDEX] {p1}: nome não encontrado na coluna B. Pulando.")
            continue

        # procurar 'Sedex/Pac/Malote' em C nas próximas linhas do bloco
        r_sedex = None
        for rr in range(r_nome, min(ws.max_row, r_nome+12)+1):
            w = ws.cell(row=rr, column=3).value
            if isinstance(w, str) and 'sedex/pac/malote' in w.strip().lower():
                r_sedex = rr; break
        if r_sedex is None:
            print(f"[SEDEX] {p1}: atividade não encontrada na coluna C. Pulando.")
            continue

        sub = grp.get(norm_key(p2))
        if sub is None or sub.empty:
            print(f"[SEDEX] {p1}: sem registros (deixando em branco).")
            continue

        writes = 0
        for _, reg in sub.iterrows():
            col_idx = None
            for k in date_keys(reg["data_obj"]):
                c = header_map.get(k)
                if c:
                    col_idx = c; break
            if not col_idx: continue
            val = int(reg["valor"])
            if val != 0:
                ws.cell(row=r_sedex, column=col_idx, value=val); writes += 1

        total_writes += writes
        print(f"[SEDEX] {p1}: {writes} células escritas.")
    return total_writes


def fill_lanctos_fixed(ws, header_map, df_long):
    """
    Preenche 'Lançamentos 45 E 19' nas linhas fixas informadas,
    casando por user_key (Respons. Entrega) -> posição do colaborador na produtividade.
    Sem criar colunas; não escreve zero; deixa em branco quando não houver valor.
    """
    if df_long.empty or "user_key" not in df_long.columns:
        print("[LANÇAMENTOS] DF vazio ou sem 'user_key'. Nada a preencher.")
        return 0

    grp = {uk: sub for uk, sub in df_long.groupby(df_long['user_key'])}
    total_writes = 0

    for ukey, meta in LANCTOS_USER_MAP.items():
        row_ativ = meta["row_ativ"]
        sub = grp.get(ukey)
        if sub is None or sub.empty:
            print(f"[LANÇAMENTOS] {ukey} → {meta['p1']}: sem registros (deixando em branco).")
            continue

        writes = 0
        for _, reg in sub.iterrows():
            # casar com o cabeçalho
            col_idx = None
            for k in date_keys(reg["data_obj"]):
                c = header_map.get(k)
                if c:
                    col_idx = c; break
            if not col_idx:
                continue
            val = int(reg["valor"])
            if val != 0:
                ws.cell(row=row_ativ, column=col_idx, value=val)
                writes += 1

        total_writes += writes
        print(f"[LANÇAMENTOS] {ukey} → {meta['p1']} (linha {row_ativ}): {writes} células escritas.")
    return total_writes


def fill_sap_fixed(ws, header_map, df_long):
    """SAPMIRO: filtra códigos MS e escreve nas linhas fixas para Alfredo/Gabriel/Pedro."""
    if df_long.empty or "nome" not in df_long.columns:
        print(f"[SAPMIRO] DataFrame vazio ou sem coluna 'nome' ({SAP_PATH}). Nada a preencher.")
        return 0

    # index por código: linhas onde 'nome' contém o MS
    cod_grp = {}
    for cod in SAP_COD_MAP.keys():
        cod_grp[cod] = df_long[df_long['nome'].str.contains(cod, case=False, regex=True, na=False)]

    total_writes = 0
    for cod, meta in SAP_COD_MAP.items():
        row_ativ = meta["row_ativ"]
        sub = cod_grp.get(cod)
        if sub is None or sub.empty:
            print(f"[SAPMIRO] {cod} → {meta['p1']}: sem registros (deixando em branco).")
            continue
        writes = 0
        for _, reg in sub.iterrows():
            col_idx = None
            for k in date_keys(reg["data_obj"]):
                c = header_map.get(k)
                if c:
                    col_idx = c; break
            if not col_idx:
                continue
            val = int(reg["valor"])
            if val != 0:
                ws.cell(row=row_ativ, column=col_idx, value=val); writes += 1
        total_writes += writes
        print(f"[SAPMIRO] {cod} → {meta['p1']} (linha {row_ativ}): {writes} células escritas.")
    return total_writes

# ===================== MAIN =====================
def main():
    wb = load_workbook(PROD_PATH)
    ws = wb["Plan1"]
    header_map = build_header_map(ws)

    # Ler pivots com o cabeçalho correto de nomes
    df_ag  = read_tabledinamica_with_namecol(AGILIS_PATH)  # detecta 'Criado por'/'Técnico'
    df_sd  = read_tabledinamica_with_namecol(SEDEX_PATH)   # detecta 'Criado por'/'Técnico'
    df_lan = read_lanctos_tabledinamica(LANCTOS_PATH)
    df_sap = read_tabledinamica_with_namecol(SAP_PATH,     name_col_hint="Nome do usuário")

    # Preencher
    ag_writes  = fill_agilis_same_row(ws, header_map, df_ag)
    sd_writes  = fill_sedex(ws, header_map, df_sd)
    ln_writes = fill_lanctos_fixed(ws, header_map, df_lan)
    sap_writes = fill_sap_fixed(ws, header_map, df_sap)

    # Salvar
    wb.save(OUT_PATH)
    print({
        "arquivo_saida": OUT_PATH,
        "agilis_celulas_escritas": ag_writes,
        "sedex_celulas_escritas": sd_writes,
        "lanctos_celulas_escritas": ln_writes,
        "sapmiro_celulas_escritas": sap_writes,
        "linhas_totais": ws.max_row,
        "colunas_totais": ws.max_column,
    })

if __name__ == "__main__":
    main()


#------------------------------------------------------------------------------------------
'''