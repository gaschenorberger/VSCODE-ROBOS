import os
import xml.etree.ElementTree as ET

# Caminhos
pasta_xml = r"C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\BUSCA_XML\XML"
arquivo_chaves = r"C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\BUSCA_XML\NOTAS.txt"

# Resultado final
resultados = []

# Lê as chaves do arquivo
with open(arquivo_chaves, "r") as f:
    chaves = [linha.strip() for linha in f if linha.strip()]

for chave in chaves:
    # Tenta localizar o arquivo XML correspondente
    caminho_xml = os.path.join(pasta_xml, f"{chave}.xml")

    if not os.path.exists(caminho_xml):
        resultados.append((chave, "❌ Arquivo XML não encontrado"))
        continue

    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()

        # Procura a tag <ICMSUFDest> (que pode estar em vários lugares dentro dos produtos)
        encontrou_difal = False
        for icmsufdest in root.findall(".//ICMSUFDest"):
            v_dest = float(icmsufdest.findtext("vICMSUFDest", default="0").replace(",", "."))
            v_remet = float(icmsufdest.findtext("vICMSUFRemet", default="0").replace(",", "."))
            v_fcp = float(icmsufdest.findtext("vFCPUFDest", default="0").replace(",", "."))

            if v_dest > 0 or v_remet > 0 or v_fcp > 0:
                encontrou_difal = True
                break

        if encontrou_difal:
            resultados.append((chave, "✅ Possui ICMS DIFAL"))
        else:
            resultados.append((chave, "❌ Não possui ICMS DIFAL"))

    except Exception as e:
        resultados.append((chave, f"⚠️ Erro ao processar: {e}"))

# Exibir os resultados
for chave, status in resultados:
    print(f"{chave}: {status}")