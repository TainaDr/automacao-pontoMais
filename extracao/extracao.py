import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
import os
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env para o ambiente do sistema
load_dotenv()

URL_DO_SITE = 'https://app2.pontomais.com.br/login' 
LOGIN_EMAIL = os.getenv('LOGIN_EMAIL')
LOGIN_SENHA = os.getenv('LOGIN_SENHA')

# Verificação opcional para garantir que as variáveis foram carregadas
if not LOGIN_EMAIL or not LOGIN_SENHA:
    print("As credenciais LOGIN_EMAIL e LOGIN_SENHA não foram encontradas no arquivo .env")

# Pega o caminho absoluto do diretório onde o script está sendo executado
diretorio_do_projeto = os.path.dirname(os.path.abspath(__file__))
# Sobe um nível no caminho para chegar na pasta que contém o projeto
diretorio_superior = os.path.dirname(diretorio_do_projeto)
# Monta o caminho final para a pasta "planilhas"
pasta_download = os.path.join(diretorio_superior, "planilhas")

data_atual = datetime.datetime.now()
dia_atual = data_atual.strftime("%Y-%m-%d")
dia_seguinte = (data_atual + datetime.timedelta(days=1)).strftime("%Y-%m-%d")

def relatorio_colaboradores():
    inicio_total = time.time()
    # --- Configuração do Navegador ---
    # Criando uma instância do ChromeOptions
    options = webdriver.ChromeOptions()
    # Desativando a flag de automação do navegador
    options.add_argument("--disable-blink-features=AutomationControlled")
    # Exclui a coleção de switches "enable-automation"
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    # Desativando a extensão useAutomationExtension
    options.add_experimental_option("useAutomationExtension", False)
    # Criando a instância do navegador com as opções configuradas

    # Verifica se a pasta "planilhas" já existe no local desejado
    if not os.path.exists(pasta_download):
    # Se não existir, cria a pasta
        os.makedirs(pasta_download)

    prefs = {
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    navegador = webdriver.Chrome(options=options)
    # Alterando a propriedade navigator.webdriver para 'undefined' para evitar detecção
    navegador.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    # Criando uma instância do WebDriverWait com um tempo limite de 10 segundos
    wait = WebDriverWait(navegador, 10)

    try:
        print(f"Acessando o site: {URL_DO_SITE}")
        navegador.get(URL_DO_SITE)
        navegador.maximize_window()

        seletor_email = "input[data-testid='login-input']"
        campo_email = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, seletor_email)))
        print("Preenchendo e-mail...")
        campo_email.send_keys(LOGIN_EMAIL)
        time.sleep(2) 
        campo_email.send_keys(Keys.TAB)
        
        time.sleep(2)

        campo_ativo = navegador.switch_to.active_element
        print("Preenchendo senha...")
        campo_ativo.send_keys(LOGIN_SENHA)
        time.sleep(2)
        
        botao_login = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Entrar')]")))
        botao_login.click()
        
        # Navegação para a Página de Relatórios 
        print("Procurando o menu 'Relatórios' e clicando...")
        menu_relatorios = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Relatórios')]")))
        menu_relatorios.click()
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Tipo do relatório')]")))
        print("Página de relatórios carregada.")
        time.sleep(3)

        xpath_dropdown_container = "//pm-select[.//span[@title='Tipo do relatório']]//div[contains(@class, 'ng-select-container')]"
        dropdown_alvo = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_dropdown_container)))

        #--------------------EXTRAÇÃO COLABORADORES--------------------

        print("Abrindo dropdown Tipo de relatório...")
        navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown_alvo)
        ActionChains(navegador).move_to_element(dropdown_alvo).click().perform()

        # Verifica se o painel abriu
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'ng-dropdown-panel')]")))
            print("Dropdown aberto com sucesso!")
        except TimeoutException:
            print("Falha: dropdown ainda não abriu. Tentando clicar novamente com JavaScript...")
            navegador.execute_script("arguments[0].click();", dropdown_alvo)
            wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'ng-dropdown-panel')]")))

        print("Localizando campo de busca dentro do dropdown...")

        campo_busca_dropdown = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[contains(@placeholder, 'Digite para buscar')]")))

        print("Digitando 'Colaboradores'...")
        campo_busca_dropdown.clear()
        campo_busca_dropdown.send_keys("Colaboradores")
        time.sleep(4)

        print("Procurando e clicando na opção 'Colaboradores'...")
        opcao_colaboradores = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Colaboradores')]"))
        )
        navegador.execute_script("arguments[0].click();", opcao_colaboradores)

        print("Relatório 'Colaboradores' selecionado com sucesso!")
        
        # Etapa de Download 
        time.sleep(2) 
        print("Clicando no botão 'Baixar'...")
        botao_baixar = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), ' Baixar ')]")))
        botao_baixar.click()
        time.sleep(2)
        print("Selecionando o formato 'XLS' para download...")
        opcao_xls = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), ' XLS ')]")))
        opcao_xls.click()

        print("\nSucesso! O relatório foi solicitado para download em formato XLS.")
        print("Aguardando 20 segundos para o download ser concluído...")
        time.sleep(20)

        #--------------------EXTRAÇÃO PONTO--------------------
        xpath_dropdown_container = "//pm-select[.//span[@title='Tipo do relatório']]//div[contains(@class, 'ng-select-container')]"
        dropdown_alvo = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_dropdown_container)))

        print("Abrindo dropdown Tipo de relatório...")
        navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown_alvo)
        ActionChains(navegador).move_to_element(dropdown_alvo).click().perform()

         # Verifica se o painel abriu
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'ng-dropdown-panel')]")))
            print("Dropdown aberto com sucesso!")
        except TimeoutException:
            print("Falha: dropdown ainda não abriu. Tentando clicar novamente com JavaScript...")
            navegador.execute_script("arguments[0].click();", dropdown_alvo)
            wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'ng-dropdown-panel')]")))

        print("Localizando campo de busca dentro do dropdown...")

        # Buscar e selecionar "Registros de ponto"
        campo_busca_dropdown = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[contains(@placeholder, 'Digite para buscar')]")))
        print("Digitando 'Registros de ponto'...")
        campo_busca_dropdown.clear()
        campo_busca_dropdown.send_keys("Registros de ponto")
        time.sleep(4)

        print("Clicando na opção 'Registros de ponto'...")
        opcao_ponto = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Registros de ponto')]")))
        navegador.execute_script("arguments[0].click();", opcao_ponto)
        print(" Relatório 'Registros de ponto' selecionado com sucesso!")

        print("Selecionando o período...")
        campo_periodo = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@placeholder, 'Selecionar período')]")))

        periodo = f"{data_atual.strftime('%d/%m/%Y')} - {(data_atual + datetime.timedelta(days=1)).strftime('%d/%m/%Y')}"
        print(f"Preenchendo o período: {periodo}")

        # Força o valor com JS para garantir que o Angular detecte a mudança
        navegador.execute_script("""
        const campo = document.querySelector("input[bsdaterangepicker]");
        campo.value = arguments[0];
        campo.dispatchEvent(new Event('input', { bubbles: true }));
        campo.dispatchEvent(new Event('change', { bubbles: true }));
        """, periodo)

        print("Período preenchido com sucesso!")

        # Etapa de Download 
        time.sleep(2) 
        print("Clicando no botão 'Baixar'...")
        botao_baixar = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), ' Baixar ')]")))
        botao_baixar.click()
        time.sleep(2)
        print("Selecionando o formato 'XLS' para download...")
        opcao_xls = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), ' XLS ')]")))
        opcao_xls.click()

        print("\nSucesso! O relatório foi solicitado para download em formato XLS.")
        print("Aguardando 20 segundos para o download ser concluído...")
        time.sleep(20)

        #--------------------EXTRAÇÃO TURNO--------------------

        xpath_dropdown_container = "//pm-select[.//span[@title='Tipo do relatório']]//div[contains(@class, 'ng-select-container')]"
        dropdown_alvo = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_dropdown_container)))

        print("Abrindo dropdown Tipo de relatório...")
        navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown_alvo)
        ActionChains(navegador).move_to_element(dropdown_alvo).click().perform()

        # Verifica se o painel abriu
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'ng-dropdown-panel')]")))
            print("Dropdown aberto com sucesso!")
        except TimeoutException:
            print("Falha: dropdown ainda não abriu. Tentando clicar novamente com JavaScript...")
            navegador.execute_script("arguments[0].click();", dropdown_alvo)
            wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'ng-dropdown-panel')]")))

        print("Localizando campo de busca dentro do dropdown...")

        campo_busca_dropdown = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[contains(@placeholder, 'Digite para buscar')]")))

        print("Digitando 'Turnos'...")
        campo_busca_dropdown.clear()
        campo_busca_dropdown.send_keys("Turnos")
        time.sleep(4)

        print("Procurando e clicando na opção 'Turnos'...")
        opcao_turnos = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Turnos')]")))
        navegador.execute_script("arguments[0].click();", opcao_turnos)

        print("Relatório 'Turnos' selecionado com sucesso!")
        
        # Etapa de Download 
        time.sleep(2) 
        print("Clicando no botão 'Baixar'...")
        botao_baixar = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), ' Baixar ')]")))
        botao_baixar.click()
        time.sleep(2)
        print("Selecionando o formato 'XLS' para download...")
        opcao_xls = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), ' XLS ')]")))
        opcao_xls.click()

        print("\nSucesso! O relatório foi solicitado para download em formato XLS.")
        print("Aguardando 20 segundos para o download ser concluído...")
        time.sleep(20)

        fim_total = time.time()
        tempo_total = fim_total - inicio_total
        
        print(f"PROCESSO FINALIZADO em {tempo_total:.2f} segundos.")

    except Exception as e:
        print(f"\n Ocorreu um erro: {e}")
        time.sleep(10)

    finally:
        print("Fechando o navegador.")
        navegador.quit()

# Chama a função principal
if __name__ == '__main__':
    relatorio_colaboradores()