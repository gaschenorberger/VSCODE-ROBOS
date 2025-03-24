import os

# Caminho da pasta com os PDFs
caminho_pasta = r'T:\CST\CST CVEL\CLIENTES\UPRESS\OPER\TRIB_FEDERAL\2024_11_08_Liq_Kepler\0000.LAUDO FINAL\01.anexos\Por Motorista\02. Contrato de Transporte de Carga'

# Percorre os arquivos da pasta
for arquivo in os.listdir(caminho_pasta):
    caminho_antigo = os.path.join(caminho_pasta, arquivo)

    # Verifica se o arquivo termina com ".pdf" e tem espaço antes do ponto
    if arquivo.endswith('.pdf') and ' .pdf' in arquivo:
        # Remove o espaço antes do ".pdf"
        nome_corrigido = arquivo.replace(' .pdf', '.pdf')
        
        caminho_novo = os.path.join(caminho_pasta, nome_corrigido)

        # Renomeia o arquivo
        os.rename(caminho_antigo, caminho_novo)
        print(f"✅ Renomeado: '{arquivo}' → '{nome_corrigido}'")
    else:
        print(f"🔍 Sem alteração: '{arquivo}'")

print("\n🚀 Processamento concluído!")