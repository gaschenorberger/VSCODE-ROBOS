import re
import os
import shutil

def remover_cnpj_pastas(diretorio_base):
    #padrao = re.compile(r'^\d{14} - ') #CNPJ COMPLETO
    padrao = re.compile(r'^\d{8} - ') #MASCARA CNPJ

    for pasta in os.listdir(diretorio_base):
        caminho_atual = os.path.join(diretorio_base, pasta)

        if os.path.isdir(caminho_atual):  
            novo_nome = padrao.sub('', pasta).strip()  
            novo_caminho = os.path.join(diretorio_base, novo_nome)

            contador = 1
            while os.path.exists(novo_caminho):
                novo_nome_com_numero = f"{novo_nome}_{contador}"
                novo_caminho = os.path.join(diretorio_base, novo_nome_com_numero)
                contador += 1

            if pasta != novo_nome:  
                os.rename(caminho_atual, novo_caminho)
                print(f'CNPJ Retirado do nome, agora com nome: {novo_caminho}')


def remover_cnpj_pastas2(diretorio_base):
    padrao = re.compile(r'^(\d{8}) - ')  # MÁSCARA CNPJ

    for pasta in os.listdir(diretorio_base):
        caminho_atual = os.path.join(diretorio_base, pasta)
        
        if os.path.isdir(caminho_atual) and padrao.match(pasta):  
            novo_nome = padrao.sub('', pasta).strip()
            novo_caminho = os.path.join(diretorio_base, novo_nome)

            if os.path.exists(novo_caminho):  
                for item in os.listdir(caminho_atual):
                    origem = os.path.join(caminho_atual, item)
                    destino = os.path.join(novo_caminho, item)
                    
                    contador = 1
                    while os.path.exists(destino):
                        nome, extensao = os.path.splitext(item)
                        destino = os.path.join(novo_caminho, f"{nome}_{contador}{extensao}")
                        contador += 1
                    
                    if os.path.isdir(origem):
                        shutil.move(origem, destino)  
                    else:
                        shutil.move(origem, destino) 
                shutil.rmtree(caminho_atual)  
                print(f'Arquivos de "{pasta}" movidos para "{novo_nome}"')
            else:
                os.rename(caminho_atual, novo_caminho)
                print(f'CNPJ retirado do nome, agora com nome: {novo_nome}')

caminho = r'C:\Users\gabriel.alvise\Documents\Arquivos ReceitanetBX'
remover_cnpj_pastas2(caminho)


