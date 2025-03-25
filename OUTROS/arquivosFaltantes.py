import pandas as pd
import os

# Caminhos do Excel e da pasta
caminho_excel = r'C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\OUTROS\arquivos.xlsx'
caminho_pasta = r'T:\CST\CST CVEL\CLIENTES\UPRESS\OPER\TRIB_FEDERAL\2024_11_08_Liq_Kepler\0000.LAUDO FINAL\01.anexos\Por Motorista\02. Contrato de Transporte de Carga'

# Ler o Excel (supondo que os nomes estão na primeira coluna)
df = pd.read_excel(caminho_excel, dtype=str)
nomes_excel = df.iloc[:, 0].str.strip().tolist()

# Obter nomes dos arquivos PDF na pasta
arquivos_pasta = [os.path.splitext(arquivo)[0] for arquivo in os.listdir(caminho_pasta) if arquivo.endswith('.pdf')]

# Identificar os que estão no Excel, mas não na pasta
faltando = [nome for nome in nomes_excel if nome not in arquivos_pasta]

# Exibir resultado
if faltando:
    print("Arquivos PDF faltando na pasta:")
    for arquivo in faltando:
        print(arquivo)
else:
    print("Todos os arquivos do Excel estão presentes na pasta.")