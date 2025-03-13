import xml.etree.ElementTree as ET
import csv
import os

# Pasta onde os arquivos XML estão localizados
input_folder = r'C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\0 - ECAC\DCTF - WEB\XMLs'
output_folder = r'C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\0 - ECAC\DCTF - WEB'
output_file_combined = r'C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\0 - ECAC\DCTF - WEB\Relatorio - DCTFWEB.csv'

# Criar a pasta de saída se não existir
os.makedirs(output_folder, exist_ok=True)

# Definir o namespace
namespace = {'ns': 'http://www.serpro.gov.br/dctf/v1'}

# Função para converter XML para CSV
def converter_xml_para_csv(file_path, filename):
    # Carregar o arquivo XML
    tree = ET.parse(file_path)
    root = tree.getroot()

    # Extrair os dados
    data = []

    # Acessar os elementos necessários
    for credito in root.findall('.//ns:CreditoTributarioApurado', namespace):
        row = {
            'nomeArquivo': filename,
            'codReceita': credito.find('ns:codReceita', namespace).text if credito.find('ns:codReceita', namespace) is not None else '',
            'ctDescricaoTributo': credito.find('ns:ctDescricaoTributo', namespace).text if credito.find('ns:ctDescricaoTributo', namespace) is not None else '',
            'ctCodGrupo': credito.find('ns:ctCodGrupo', namespace).text if credito.find('ns:ctCodGrupo', namespace) is not None else '',
            'ctDescGrupo': credito.find('ns:ctDescGrupo', namespace).text if credito.find('ns:ctDescGrupo', namespace) is not None else '',
            'ctValor': credito.find('ns:ctValor', namespace).text if credito.find('ns:ctValor', namespace) is not None else '',
            'paDebito': credito.find('ns:paDebito', namespace).text if credito.find('ns:paDebito', namespace) is not None else '',
            'vlTotalCred': credito.find('ns:vlTotalCred', namespace).text if credito.find('ns:vlTotalCred', namespace) is not None else '',
            'saldoaPagar': credito.find('ns:saldoaPagar', namespace).text if credito.find('ns:saldoaPagar', namespace) is not None else '',
        }
        data.append(row)

    return data

# Definir os cabeçalhos
headers = ['nomeArquivo', 'codReceita', 'ctDescricaoTributo', 'ctCodGrupo', 'ctDescGrupo', 'ctValor', 'paDebito', 'vlTotalCred', 'saldoaPagar']

# Escrever os dados combinados em um arquivo CSV
with open(output_file_combined, 'w', newline='', encoding='utf-8') as file:
    writer = csv.DictWriter(file, fieldnames=headers, delimiter=';', quoting=csv.QUOTE_MINIMAL)
    writer.writeheader()

    # Percorrer todos os arquivos XML na pasta especificada
    for filename in os.listdir(input_folder):
        if filename.endswith('.xml'):
            file_path = os.path.join(input_folder, filename)
            data = converter_xml_para_csv(file_path, filename)
            writer.writerows(data)

print(f"Todos os arquivos XML foram convertidos e combinados no CSV '{output_file_combined}'.")
