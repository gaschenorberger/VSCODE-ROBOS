import pandas as pd
import os

# Caminhos do Excel e da pasta
caminho_excel = r'C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\OUTROS\arquivos.xlsx'
caminho_pasta = r'T:\CST\CST CVEL\CLIENTES\UPRESS\OPER\TRIB_FEDERAL\2024_11_08_Liq_Kepler\0000.LAUDO FINAL\01.anexos\Por Motorista\02. Contrato de Transporte de Carga'

# Ler o Excel (supondo que os nomes estão na primeira coluna)
df = pd.read_excel(caminho_excel, dtype=str)
nomes_excel = df.iloc[:, 0].str.strip().tolist()

# Obter nomes dos arquivos PDF na pasta (removendo a extensão)
arquivos_pasta = [os.path.splitext(arquivo)[0].strip() for arquivo in os.listdir(caminho_pasta) if arquivo.endswith('.pdf')]

# Identificar os arquivos que estão a mais na pasta
a_mais = [arquivo for arquivo in arquivos_pasta if arquivo not in nomes_excel]

# Exibir resultado
if a_mais:
    print("📌 Arquivos que estão a mais na pasta:")
    for arquivo in a_mais:
        print(arquivo)
else:
    print("✅ Não há arquivos a mais na pasta.")
