import os
import re
import time
from collections import defaultdict

def get_latest_files(file_list):
    grouped_files = defaultdict(list)
    pattern = re.compile(r'PISCOFINS_(\d{8})_(\d{8})_\d+_(Original|Retificadora)_(\d+)_')

    for file in file_list:
        match = pattern.search(file)
        if match:
            start_date, end_date, file_type, timestamp = match.groups()
            period = f"{start_date}_{end_date}"
            grouped_files[period].append((file_type, timestamp, file))

    files_to_keep = []

    for period, files in grouped_files.items():
        retificadoras = [(t, ts, f) for t, ts, f in files if t == "Retificadora"]
        if retificadoras:
            latest_file = max(retificadoras, key=lambda x: x[1])[2]  # Pegando nome do arquivo diretamente
            files_to_keep.append(latest_file)
        else:
            original = next((f for t, ts, f in files if t == "Original"), None)
            if original:
                files_to_keep.append(original)

    return files_to_keep

# Exemplo de uso
if __name__ == "__main__":
    start_time = time.time()
    
    caminho_dos_arquivos = r"C:\Users\gabriel.alvise\Documents\Arquivos ReceitanetBX"  # Caminho corrigido

    # Verifica se o diretório existe
    if not os.path.exists(caminho_dos_arquivos):
        print(f"Erro: O diretório '{caminho_dos_arquivos}' não existe.")
        exit(1)

    arquivos = [f for f in os.listdir(caminho_dos_arquivos) if f.startswith("PISCOFINS")]
    arquivos_validos = get_latest_files(arquivos)

    arquivos_excluidos = []
    for arquivo in arquivos:
        if arquivo not in arquivos_validos:
            os.remove(os.path.join(caminho_dos_arquivos, arquivo))
            arquivos_excluidos.append(arquivo)

    end_time = time.time()

    print("Arquivos filtrados com sucesso!")
    print(f"Tempo decorrido: {end_time - start_time:.2f} segundos")
    if arquivos_excluidos:
        print("Arquivos excluídos:")
        for arquivo in arquivos_excluidos:
            print(arquivo)
    else:
        print("Nenhum arquivo foi excluído.")
