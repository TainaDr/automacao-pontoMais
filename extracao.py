import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

URL_DO_SITE = 'https://app2.pontomais.com.br/login' 

def relatorio_colaboradores():
    # --- Configuração do Navegador ---
    # create Chromeoptions instance 
    options = webdriver.ChromeOptions() 
    
    # adding argument to disable the AutomationControlled flag 
    options.add_argument("--disable-blink-features=AutomationControlled") 
    
    # exclude the collection of enable-automation switches 
    options.add_experimental_option("excludeSwitches", ["enable-automation"]) 
    
    # turn-off userAutomationExtension 
    options.add_experimental_option("useAutomationExtension", False) 
    
    # setting the driver path and requesting a page 
    # Criando a instância do navegador configurada
    navegador = webdriver.Chrome(options=options) 
    
    # changing the property of the navigator value for webdriver to undefined 
    navegador.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})") 

    wait = WebDriverWait(navegador, 10)

    try:
        print(f"Acessando o site: {URL_DO_SITE}")
        navegador.get(URL_DO_SITE)
        navegador.maximize_window()

        print("\n--- AÇÃO MANUAL NECESSÁRIA ---")
        input(">>> Após fazer o login no navegador, volte aqui e pressione Enter para continuar...")
        
        print("\nContinuando a automação...")
        print("Procurando o menu 'Relatórios' e clicando...")
        menu_relatorios = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Relatórios')]")))
        menu_relatorios.click()
        
        print("Aguardando a página de relatórios carregar...")
        # Procura um elemento que confirma que estamos na página certa
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Relatórios')]")))
        print("Página de relatórios carregada.")

        # Clica no dropdown "Tipo do relatório" para abri-lo
        print("Clicando no dropdown 'Tipo do relatório'...")
        dropdown_tipo_relatorio = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "ng-input")))
        
        # Aguarda o campo de busca e digitar "Colaboradores"
        print("Digitando 'Colaboradores' no campo de busca...")
        campo_busca = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@placeholder='Digite para buscar']")))
        campo_busca.send_keys("Colaboradores")

        # Aguarda a opção "Colaboradores" aparecer na lista de resultados e clicar nela
        print("Selecionando a opção 'Colaboradores'...")
        # Procura por um elemento com o texto exato
        opcao_colaboradores = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'option-label') and text()='Colaboradores']")))
        opcao_colaboradores.click()
        
        # Clica no botão dropdown "Baixar"
        print("Clicando no botão 'Baixar'...")
        # Procura por um botão que contenha o texto 'Baixar'
        botao_baixar_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Baixar')]")))
        botao_baixar_dropdown.click()
        
        # Aguarda a opção "XLS" aparecer e clicar nela
        print("Selecionando a opção de download 'XLS'...")
        # Procura por um link ou item de menu com o texto 'XLS'
        opcao_xls = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'XLS')]")))
        opcao_xls.click()
        
        print("\n Sucesso! O relatório foi solicitado para download em formato XLS.")
        print("Aguardando 20 segundos para o download...")
        time.sleep(20)

    except Exception as e:
        print(f"\n Ocorreu um erro: {e}")
        time.sleep(10)

    finally:
        print("Fechando o navegador.")
        navegador.quit()

# Chama a função principal
if __name__ == '__main__':
    relatorio_colaboradores()