import os
from collections import Counter

# Caminhos dos arquivos e pastas
caminho_lista = r"C:\VS_CODE\BUSCA_XML\NOTAS.txt"
pasta_downloads = r"C:\Users\Certezza\Documents\3 IRMAOS"

# 1. Ler o arquivo e contar as ocorrências das linhas
with open(caminho_lista, "r", encoding="utf-8") as f:
    linhas = [line.strip() for line in f if line.strip()]  # Remove espaços extras e linhas vazias

contagem = Counter(linhas)  # Conta quantas vezes cada linha aparece
duplicatas = {item: qtd for item, qtd in contagem.items() if qtd > 1}  # Filtra apenas os duplicados

# 2. Verificar arquivos já baixados (removendo .xml dos nomes)
arquivos_baixados = {os.path.splitext(arquivo)[0] for arquivo in os.listdir(pasta_downloads) if arquivo.endswith(".xml")}

# 3. Identificar os arquivos que ainda precisam ser baixados
faltantes = [item for item in linhas if item not in arquivos_baixados]

# 4. Identificar arquivos não listados
arquivos_na_pasta = {os.path.splitext(arquivo)[0] for arquivo in os.listdir(pasta_downloads) if arquivo.endswith(".xml")}
nao_listados = arquivos_na_pasta - set(linhas)

# 5. Exibir resultados
print(f"Total de itens na lista: {len(linhas)}")
print(f"Arquivos já baixados: {len(arquivos_baixados)}")
print(f"Faltam baixar: {len(faltantes)}")

# 6. Mostrar duplicatas
if duplicatas:
    print("\n🔴 Linhas duplicadas na lista:")
    for item, qtd in duplicatas.items():
        print(f"{item} → {qtd}x")
else:
    print("\n✅ Nenhuma duplicata encontrada.")

# 7. Mostrar os arquivos que faltam baixar
if faltantes:
    print("\n🟡 Arquivos que ainda precisam ser baixados:")
    for item in faltantes:
        print(item)
else:
    print("\n✅ Todos os arquivos já foram baixados.")

# 8. Mostrar arquivos não listados
if nao_listados:
    print("\n⚠️ Arquivos presentes na pasta mas não listados:")
    for item in nao_listados:
        print(item)
else:
    print("\n✅ Não há arquivos extras na pasta.")
