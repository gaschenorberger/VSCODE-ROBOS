import pandas as pd
import openpyxl
import re

def format_cnpj(cnpj):
    # Converte para string e remove tudo que não for número
    cnpj = re.sub(r'\D', '', str(cnpj))
    
    # Garante que o CNPJ tenha 14 dígitos preenchendo com zeros à esquerda
    cnpj = cnpj.zfill(14)
    
    # Se o CNPJ não tiver 14 dígitos, retorna como está
    if len(cnpj) != 14:
        return cnpj
    
    # Aplica a máscara XX.XXX.XXX/XXXX-XX
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

def process_excel(file_path, output_path):
    # Carregar a planilha, garantindo que os CNPJs sejam lidos como strings
    df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
    
    # Exibir as colunas disponíveis
    print("Colunas encontradas:", df.columns.tolist())
    
    # Tentar identificar a coluna correta
    possible_columns = ['CNPJ', 'cnpj', 'B']  # Ajuste conforme necessário
    for col in possible_columns:
        if col in df.columns:
            df[col] = df[col].apply(format_cnpj)
            break
    else:
        print("Nenhuma coluna correspondente encontrada.")
    
    # Salvar o arquivo corrigido
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"Arquivo salvo em {output_path}")

# Exemplo de uso

process_excel(r'Excel\planilha_corrigida1.xlsx', r'Excel\planilha_corrigida.xlsx')
