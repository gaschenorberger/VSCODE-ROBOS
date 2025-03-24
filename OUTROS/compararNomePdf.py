import os
import re
import fitz  # PyMuPDF

# Caminhos da pasta
caminho_pasta = r'T:\CST\CST CVEL\CLIENTES\UPRESS\OPER\TRIB_FEDERAL\2024_11_08_Liq_Kepler\0000.LAUDO FINAL\01.anexos\Por Motorista\02. Contrato de Transporte de Carga'

# Função para extrair o NRO do PDF
def extrair_nro(caminho_pdf):
    doc = fitz.open(caminho_pdf)
    texto = ""
    for pagina in doc:
        texto += pagina.get_text()

    nro_match = re.search(r'Nro:\s*(\d+)', texto)
    return nro_match.group(1) if nro_match else None

# Lista para armazenar os que não batem
nros_incorretos = []

# Verificar cada arquivo na pasta
for arquivo in os.listdir(caminho_pasta):
    if arquivo.endswith('.pdf'):
        caminho_arquivo = os.path.join(caminho_pasta, arquivo)
        nro_extraido = extrair_nro(caminho_arquivo)

        if nro_extraido:
            nome_arquivo = os.path.splitext(arquivo)[0]

            if nro_extraido == nome_arquivo:
                print(f"✅ NRO {nro_extraido} confere com o arquivo: {arquivo}")
            else:
                print(f"❌ NRO {nro_extraido} não bate com o arquivo: {arquivo}")
                nros_incorretos.append((arquivo, nro_extraido))
        else:
            print(f"⚠️ NRO não encontrado no arquivo: {arquivo}")

# Exibir resumo dos NROs que não batem
if nros_incorretos:
    print("\n🔍 Arquivos com NRO incorreto:")
    for arquivo, nro in nros_incorretos:
        print(f"Arquivo: {arquivo} | NRO encontrado: {nro}")
else:
    print("\n✅ Todos os NROs estão corretos!")