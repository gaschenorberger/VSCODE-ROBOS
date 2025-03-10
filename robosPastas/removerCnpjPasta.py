import re
import os

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



caminho = r'C:\Users\gabriel.alvise\Documents\Arquivos ReceitanetBX'
remover_cnpj_pastas(caminho)

