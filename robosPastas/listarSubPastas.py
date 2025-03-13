'''
A FUNÇÃO DESSE SCRIPT É LISTAR PRA MIM AS SUBPASTAS QUE ESTAO DENTRO DE UMA PASTA

Percorre a pasta principal, identifica as subpastas das empresas e retorna as pastas dentro de cada uma delas

EX: 
-PASTA PRINCIPAL
--PASTAS DE EMPRESAS
*---PASTAS DENTRO DAS PASTAS DAS EMPRESAS (A QUE VAI ME RETORNAR)

        Empresa: 00704766000157 - P V SUPERMERCADO EIRELI
        Pastas dentro:
            - ECD
            - ECF
'''

import os

def listar_pastas(caminho_principal):
    if not os.path.exists(caminho_principal):
        print("O caminho especificado não existe.")
        return
    
    for empresa in os.listdir(caminho_principal):
        caminho_empresa = os.path.join(caminho_principal, empresa)
        if os.path.isdir(caminho_empresa):
            print(f"Empresa: {empresa}")
            subpastas = [pasta for pasta in os.listdir(caminho_empresa) if os.path.isdir(os.path.join(caminho_empresa, pasta))]
            if subpastas:
                print("  Pastas dentro:")
                for subpasta in subpastas:
                    print(f"    - {subpasta}")
            else:
                print("  Nenhuma subpasta encontrada.")
            print("-" * 30)

caminho_principal = r"D:\BKP_BX"
listar_pastas(caminho_principal)
