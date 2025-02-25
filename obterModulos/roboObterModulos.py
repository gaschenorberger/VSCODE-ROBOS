from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from tqdm import tqdm
import os
import pandas as pd
import time

def iniciar_navegador(com_debugging_remoto=True):
    chrome_driver_path = ChromeDriverManager().install()
    #chrome_driver_path = r'C:\Users\gabriel.alvise\.wdm\drivers\chromedriver\win64\130.0.6723.91\chromedriver-win32/chromedriver.exe'
    chrome_driver_executable = os.path.join(os.path.dirname(chrome_driver_path), 'chromedriver.exe')
    #print(chrome_driver_path)
    
    if not os.path.isfile(chrome_driver_executable):
        raise FileNotFoundError(f"O ChromeDriver não foi encontrado em {chrome_driver_executable}")

    service = Service(executable_path=chrome_driver_executable)
    
    chrome_options = Options()
    if com_debugging_remoto:
        remote_debugging_port = 9222
        chrome_options.add_experimental_option("debuggerAddress", f"localhost:{remote_debugging_port}")
    
    navegador = webdriver.Chrome(service=service, options=chrome_options)
    return navegador

navegador = iniciar_navegador()

codigos_esperados = {
    "B020", "C100", "C170", "C190", "C197", "C320", "C390", "C490", "C590", "C850", "C890", "D190", "D590", "H010", "K200",
    "G125", "G140", "E110", "E111", "E113", "E116", "E210", "E250", "E310", "E520", "0200", "0150",
    "0500", "0220"
}

arquivo_excel = r"obterModulos\RelatoriosNaoEncontrados.xlsx"

def coletaDados():
    
    inputBotao = navegador.find_element(By.XPATH, '//*[@id="empresa"]/button')
    inputBotao.click()
    time.sleep(0.5)
    
    empresas = navegador.find_elements(By.XPATH, '//*[@id="empresa_panel"]/ul/li')
    total_empresas = len(empresas)
    print(f"\n🔹 Total de empresas a processar: {total_empresas}\n")
    
    indiceInicial = 1

    with tqdm(total=total_empresas, desc="Processando empresas", unit="empresa") as pbar:
        while indiceInicial <= total_empresas:  
            inputBotao = navegador.find_element(By.XPATH, '//*[@id="empresa"]/button')
            inputBotao.click()
            time.sleep(0.5)

            inputEmpresa = navegador.find_element(By.XPATH, f'//*[@id="empresa_panel"]/ul/li[{indiceInicial}]')
            nomeEmpresa = inputEmpresa.text
            inputEmpresa.click()

            print(f'\n🔹 Iniciando empresa: {nomeEmpresa}')
            time.sleep(0.5)

            informacoesData()

            while True:
                try:
                    elemento = navegador.find_element(By.XPATH, '//*[@id="j_idt108"]')
                    if not elemento.is_displayed():  
                        print("✅ Carregamento finalizado, continuando a execução...")
                        break  
                    print("⏳ Carregando...")
                except NoSuchElementException:
                    print("✅ Carregamento finalizado, continuando a execução...")
                    break  
                time.sleep(2)  

            abas = {
                "Relatório Analítico": ('//*[@id="tabs"]/ul/li[2]/a', '//*[@id="j_idt95_content"]'),
                "Relatório Apuração": ('//*[@id="tabs"]/ul/li[3]/a', '//*[@id="j_idt99_content"]'),
                "Cadastros": ('//*[@id="tabs"]/ul/li[4]/a', '//*[@id="j_idt103"]')
            }

            dados_nao_esperados = []

            for nomeAba, (xpathAba, xpathDiv) in abas.items():
                try:
                    navegador.find_element(By.XPATH, xpathAba).click()
                    time.sleep(1)

                    elemento_div = navegador.find_element(By.XPATH, xpathDiv)
                    texto_div = elemento_div.text.strip()

                    if not texto_div:
                        print(f"⚠️ {nomeAba} está vazia para {nomeEmpresa}, pulando...")
                        continue

                    print(f"\n🔸 {nomeAba}:\n{texto_div}")

                    for linha in texto_div.split("\n"):
                        partes = linha.split(" - ")
                        if len(partes) > 1:  
                            codigo = partes[0].strip()
                            descricao = " - ".join(partes[1:]).strip()
                            
                            if codigo not in codigos_esperados:
                                dados_nao_esperados.append([codigo, descricao, nomeEmpresa])
                                print(f" ❌ ➜ {codigo} - {descricao} -- {nomeEmpresa}")

                except NoSuchElementException:
                    print(f"⚠️ {nomeAba} não encontrada para {nomeEmpresa}, pulando...")

            salvarExcel(dados_nao_esperados)

            pbar.update(1)  

            indiceInicial += 1


def informacoesData():
    #------------------------MES INICIO-------------------------------
 
        inputMesInicio = navegador.find_element(By.XPATH, '//*[@id="dtIni_inline"]/div/div/div/select[1]')
        inputMesInicio.click()
        time.sleep(0.5)

        opcaoJan = navegador.find_element(By.XPATH, '//option[contains(text(), "Jan")]')
        opcaoJan.click()
        time.sleep(0.5)

        inputAnoInicio = navegador.find_element(By.XPATH, '//*[@id="dtIni_inline"]/div/div/div/select[2]')
        inputAnoInicio.click()
        time.sleep(0.5)
            
        opcao2019 = navegador.find_element(By.XPATH, '//option[contains(text(), "2019")]')
        opcao2019.click()
        time.sleep(0.5)

        #------------------------MES FIM-------------------------------

        inputMesFim = navegador.find_element(By.XPATH, '//*[@id="dtFin_inline"]/div/div/div/select[1]')
        inputMesFim.click()
        time.sleep(0.5)

        opcaoDez = navegador.find_element(By.XPATH, '//*[@id="dtFin_inline"]/div/div/div/select[1]/option[contains(text(), "Dez")]')
        opcaoDez.click()
        time.sleep(0.5)
        
        inputAnoFim = navegador.find_element(By.XPATH, '//*[@id="dtFin_inline"]/div/div/div/select[2]')
        inputAnoFim.click()
        time.sleep(0.5)
            
        opcao2024 = navegador.find_element(By.XPATH, '//*[@id="dtFin_inline"]/div/div/div/select[2]/option[contains(text(), "2024")]')
        opcao2024.click()
        time.sleep(0.5)

        inputBtnPesquisar = navegador.find_element(By.XPATH, '//*[@id="j_idt83"]')
        inputBtnPesquisar.click()
        time.sleep(0.5)

def salvarExcel(dados):
    """ Salva os códigos não esperados em um arquivo Excel """
    if not dados:
        return  # Se não houver dados, não faz nada

    df_novo = pd.DataFrame(dados, columns=["Código", "Descrição", "Empresa"])

    # Se o arquivo já existir, anexar os dados
    if os.path.exists(arquivo_excel):
        df_existente = pd.read_excel(arquivo_excel)
        df_final = pd.concat([df_existente, df_novo], ignore_index=True)
    else:
        df_final = df_novo

    df_final.to_excel(arquivo_excel, index=False)
    print(f"\n✅ Dados salvos em '{arquivo_excel}'!\n")

coletaDados()