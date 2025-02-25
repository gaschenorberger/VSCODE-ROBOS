from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from PIL import ImageGrab
from PIL import Image
import cv2
import numpy as np
import difflib
import os
import time
import pytesseract
import pyautogui
import calendar
from datetime import datetime
import shutil
import random
import mousekey

mkey = mousekey.MouseKey()

#start chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Selenium\ChromeTestProfile"
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  

def iniciar_navegador(com_debugging_remoto=True):
    chrome_driver_path = ChromeDriverManager().install()
    chrome_driver_executable = os.path.join(os.path.dirname(chrome_driver_path), 'chromedriver.exe')
    
    if not os.path.isfile(chrome_driver_executable):
        raise FileNotFoundError(f"O ChromeDriver não foi encontrado em {chrome_driver_executable}")

    service = Service(executable_path=chrome_driver_executable)
    
    chrome_options = Options()
    if com_debugging_remoto:
        remote_debugging_port = 9222
        chrome_options.add_experimental_option("debuggerAddress", f"localhost:{remote_debugging_port}")
    
    navegador = webdriver.Chrome(service=service, options=chrome_options)
    return navegador

navegador = iniciar_navegador(com_debugging_remoto=True)

def procurar_imagem(nome_arquivo, confidence=0.8, region=None, max_tentativas=60, horizontal=0, vertical=0, acao='clicar', clicks=1):
    
    def click(x, y):
        pyautogui.click(x, y)

    def doubleClick(x, y):
        pyautogui.doubleClick(x, y)

    def coordenada(x, y):
        print(f'Coordenadas da imagem: ({x}, {y})')

    def move_mouse(x,y,variationx=(-5, 5),variationy=(-5, 5),up_down=(0.2, 0.3),min_variation=-10,max_variation=10,use_every=4,sleeptime=(0.009, 0.019),linear=90,):
        mkey.left_click_xy_natural(
            int(x) - random.randint(*variationx),
            int(y) - random.randint(*variationy),
            delay=random.uniform(*up_down),
            min_variation=min_variation,
            max_variation=max_variation,
            use_every=use_every,
            sleeptime=sleeptime,
            print_coords=True,
            percent=linear,
        )
    
    acoes_validas = ['clicar', 'mover_clicar']

    if acao not in acoes_validas:
        raise ValueError(f"Ação inválida: '{acao}'. Escolha entre {acoes_validas}.")

    tentativas = 0

    while tentativas < max_tentativas:
        tentativas += 1
        try:
            img = pyautogui.locateCenterOnScreen(nome_arquivo, confidence=confidence, region=region)
            if img:
                x, y = img
                x += horizontal
                y += vertical
                #coordenada(x, y)
                if acao == 'clicar':
                    if clicks == 1:
                        click(x, y)
                    else:
                        doubleClick(x, y)
                elif acao == 'mover_clicar':
                    move_mouse(x, y)
                return True
        except pyautogui.ImageNotFoundException:
            pass
        time.sleep(1)

    print(f'Imagem não encontrada após {max_tentativas} segundos.')
    return False

def preencher_inf_inicial(matriz_filial):
    navegador.switch_to.default_content()
    navegador.switch_to.frame(navegador.find_element(By.XPATH, '/html/frameset/frame'))
    match matriz_filial:
        case 'matriz':
            ie_empresa = navegador.find_element(By.XPATH, '//*[@id="form1:j_id6_body"]/table[1]/tbody/tr[1]/td[2]/input') 
            cpf_socio = navegador.find_element(By.XPATH, '//*[@id="form1:j_id6_body"]/table[1]/tbody/tr[2]/td[2]/input')
            ult_prot_dief = navegador.find_element(By.XPATH, '//*[@id="form1:j_id6_body"]/table[1]/tbody/tr[3]/td[2]/input')

            ie_empresa.clear()
            ie_empresa.send_keys('125162103')
            cpf_socio.clear()
            cpf_socio.send_keys('52920240315')
            ult_prot_dief.clear()
            ult_prot_dief.send_keys('7704485')
            print('Inf Matriz')

        case 'filial':
            ie_empresa = navegador.find_element(By.XPATH, '//*[@id="form1:j_id6_body"]/table[1]/tbody/tr[1]/td[2]/input') 
            cpf_socio = navegador.find_element(By.XPATH, '//*[@id="form1:j_id6_body"]/table[1]/tbody/tr[2]/td[2]/input')
            ult_prot_dief = navegador.find_element(By.XPATH, '//*[@id="form1:j_id6_body"]/table[1]/tbody/tr[3]/td[2]/input')
            
            ie_empresa.clear()
            ie_empresa.send_keys('126997586')
            cpf_socio.clear()
            cpf_socio.send_keys('52920240315')
            ult_prot_dief.clear()
            ult_prot_dief.send_keys('7704946')
            print('Inf Filial')

def notas_entradas_saidas(notas_fiscais): 
    match notas_fiscais:
        case 'saidas':
            notas_emitidas = navegador.find_element(By.XPATH, '//*[@id="form1:j_id20:0"]')
            notas_emitidas.click()
            print('Saídas - Notas emitidas')
            return notas_fiscais
        case 'entradas':
            notas_recebidas = navegador.find_element(By.XPATH, '//*[@id="form1:j_id20:1"]')
            notas_recebidas.click()
            print('Entradas - Notas recebidas')
            return notas_fiscais

def tipo_de_notas(notas):
    match notas:
        case 'ambos':
            notas_ambos = navegador.find_element(By.XPATH, '//*[@id="form1:j_id25:0"]')
            notas_ambos.click()
            print('Tipo de notas - Ambos')
            return notas
        case 'nfe':
            notas_nfe = navegador.find_element(By.XPATH, '//*[@id="form1:j_id25:1"]')
            notas_nfe.click()
            print('Tipo de notas - NFE')
            return notas

def reconhecer_cod_captcha():
    screenshot = pyautogui.screenshot(region=(512, 589, 130, 35))
    screenshot.save("screenshot.png")
    img = Image.open("screenshot.png")
    text = pytesseract.image_to_string(img, lang='por')
    text = text.replace(" ", "")   
    text = text.replace("\t", "") 
    text = text.replace("\n", "") 
    text = text.strip() 

    print("Texto reconhecido:", text)

    codigo_captcha = navegador.find_element(By.XPATH, '//*[@id="form1:cap"]/tbody/tr[2]/td/div/input')
    codigo_captcha.click()
    codigo_captcha.clear()
    time.sleep(1)

    for char in text:
        codigo_captcha.click()
        pyautogui.write(char)
        time.sleep(0.3)
    
    baixar_xml = navegador.find_element(By.XPATH, '//*[@id="form1:j_id6_body"]/table[4]/tbody/tr/td[2]/input')
    baixar_xml.click()
    time.sleep(1.5)

    if procurar_imagem(r'C:\Users\gabriel.alvise\Desktop\ROBOS\robo_download_nfe\erro_download.png', max_tentativas=2) is True:
        print('Erro encontrado')
        return False
    else:
        print('Nenhum erro encontrado, continuando')
        return True

def gerar_lista_meses(ano_inicio, ano_fim, mes_inicio=1):
    lista_meses = []
    for ano in range(ano_inicio, ano_fim + 1):
        for mes in range(mes_inicio, 13):
            inicio_mes = f"01/{mes:02d}/{ano}"
            ultimo_dia_mes = calendar.monthrange(ano, mes)[1]
            fim_mes = f"{ultimo_dia_mes:02d}/{mes:02d}/{ano}"
            lista_meses.append((inicio_mes, fim_mes))
        mes_inicio = 1 
    return lista_meses

def renomear_arquivo(pasta_downloads, nome_do_arquivo):
    novo_nome_base = os.path.join(pasta_downloads, nome_do_arquivo)
    extensao = '.zip'
    
    while os.path.exists(novo_nome_base + extensao):
        novo_nome_base = os.path.join(pasta_downloads, f"{nome_do_arquivo}")

    arquivo_baixado = os.path.join(pasta_downloads, "arquivo_completo.zip")
    if os.path.exists(arquivo_baixado):
        os.rename(arquivo_baixado, novo_nome_base + extensao)
        print(f"Arquivo renomeado para: {novo_nome_base + extensao}")
    else:
        print("Arquivo para renomear não encontrado.")

def mover_arquivo(pasta_origem, pasta_destino, nome_do_arquivo):
    caminho_origem = os.path.join(pasta_origem, nome_do_arquivo)
    caminho_destino = os.path.join(pasta_destino, nome_do_arquivo)

    if os.path.exists(caminho_origem):
        shutil.move(caminho_origem, caminho_destino)
        print(f"Arquivo movido para: {caminho_destino}")
    else:
        print(f"Arquivo {nome_do_arquivo} não encontrado na pasta de origem.")

def verificar_download_temporario(pasta_downloads):
        arquivos = os.listdir(pasta_downloads)
        
        if any(arquivo.endswith('.crdownload') for arquivo in arquivos):
            print("Download em andamento (arquivo temporário encontrado).")
            return True
        else:
            print("Nenhum download em andamento.")
            return False

def dados_iniciais(inicio, fim, razao='matriz', ent_sai='saidas', tp_notas='ambos'):
    preencher_inf_inicial(razao)
    notas_fiscais = notas_entradas_saidas(ent_sai).upper()
    tipo_de_notas(tp_notas)
    time.sleep(0.5)

    data_inicio = f'{inicio}'
    data_final = f'{fim}'
    time.sleep(1)

    data_ini = navegador.find_element(By.XPATH, '//*[@id="form1:dtIniInputDate"]') 
    data_fim = navegador.find_element(By.XPATH, '//*[@id="form1:dtFinInputDate"]') 
    navegador.execute_script("arguments[0].value = arguments[1]", data_ini, f'{data_inicio}')
    navegador.execute_script("arguments[0].value = arguments[1]", data_fim, f'{data_final}')
    time.sleep(0.5)
    return notas_fiscais

def aguardando_download(pasta_download_origem, nome_arquivo):
    pasta_download_origem = r"C:\Users\gabriel.alvise\Downloads"
    nome_arquivo = 'arquivo_completo.zip'

    while any([filename.endswith('.crdownload') or filename.endswith('.part') 
               for filename in os.listdir(pasta_download_origem)]):
        print("Aguardando o download ser concluído...")
        time.sleep(2)  

    while not os.path.exists(os.path.join(pasta_download_origem, nome_arquivo)):
        print(f"Verificando se {nome_arquivo} foi baixado...")
        time.sleep(2)

    print(f"Download concluído: {nome_arquivo} está na pasta.")

def download_nfe_matriz():
    lista_meses = gerar_lista_meses(2021, 2024, 8) 
    
    for inicio, fim in lista_meses:
        print(f"'{inicio}', '{fim}'")

        # BAIXANDO SAÍDAS
        notas_fiscais = dados_iniciais(inicio, fim)
        
        for i in range(5):
            result = reconhecer_cod_captcha()
            if result is False:
                navegador.refresh()
                time.sleep(1.5)
                notas_fiscais = dados_iniciais(inicio, fim)
                result = reconhecer_cod_captcha()   
                time.sleep(1)
                arquivos = os.listdir(r"C:\Users\gabriel.alvise\Downloads")
                
                if any(arquivo.endswith('.crdownload') for arquivo in arquivos):
                    break
            else:
                break
        
        time.sleep(1.5)
        pasta_download_origem = r"C:\Users\gabriel.alvise\Downloads"
        pasta_matriz_saidas = r"C:\Users\gabriel.alvise\Desktop\ROBOS\robo_download_nfe\matriz\saidas" 
        nome_arquivo = 'arquivo_completo.zip'
        nome_arq = f'MATRIZ-{notas_fiscais}-{inicio}'
        nome_do_arquivo = nome_arq.replace('/', '_')
        nome_do_arq_mover = nome_do_arquivo + '.zip'
        print(nome_do_arquivo)

        aguardando_download(pasta_download_origem, nome_arquivo)
        time.sleep(1)

        if os.path.exists(os.path.join(pasta_download_origem, "arquivo_completo.zip")):
            renomear_arquivo(pasta_download_origem, nome_do_arquivo)
        time.sleep(0.3)

        mover_arquivo(pasta_download_origem, pasta_matriz_saidas, nome_do_arq_mover)
        time.sleep(1)

        dados_iniciais(inicio, fim, razao='matriz', ent_sai='entradas', tp_notas='nfe')

        for i in range(5):
            result = reconhecer_cod_captcha()
            if result is False:
                navegador.refresh()
                time.sleep(1.5)
                notas_fiscais = dados_iniciais(inicio, fim, ent_sai='entradas', tp_notas='nfe')
                result = reconhecer_cod_captcha()   
                time.sleep(2)
                arquivos = os.listdir(r"C:\Users\gabriel.alvise\Downloads")

                if any(arquivo.endswith('.crdownload') for arquivo in arquivos):
                    break
            else:
                break
        time.sleep(1)

        pasta_matriz_entradas = r"C:\Users\gabriel.alvise\Desktop\ROBOS\robo_download_nfe\matriz\entradas"
        nome_arq_ent = f'MATRIZ-ENTRADAS-{inicio}'
        nome_do_arquivo_entradas = nome_arq_ent.replace('/', '_')
        nome_do_arq_mover_ent = nome_do_arquivo_entradas + '.zip'

        aguardando_download(pasta_download_origem, nome_arquivo)
        time.sleep(1)

        if os.path.exists(os.path.join(pasta_download_origem, "arquivo_completo.zip")):
            renomear_arquivo(pasta_download_origem, nome_do_arquivo_entradas)
        time.sleep(1)

        mover_arquivo(pasta_download_origem, pasta_matriz_entradas, nome_do_arq_mover_ent)
        time.sleep(1)



download_nfe_matriz()
