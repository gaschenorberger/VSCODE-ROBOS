import os
import pandas as pd

def juntar_csvs_em_cada_fl(pasta_principal, destino):
    if not os.path.exists(destino):
        os.makedirs(destino)

    for pasta_fl in os.listdir(pasta_principal):
        caminho_fl = os.path.join(pasta_principal, pasta_fl)

        if os.path.isdir(caminho_fl) and pasta_fl.startswith("FL"):
            arquivos_csv = [os.path.join(caminho_fl, f) for f in os.listdir(caminho_fl) if f.endswith(".csv")]
            print(f'Iniciando {pasta_fl}')

            if arquivos_csv:
                lista_dfs = []
                
                numArquivo = 1
                for arquivo in arquivos_csv:
                    try:
                        with open(arquivo, "r", encoding="latin1") as file:
                            print(f'Iniciando arquivo {arquivo} -- {numArquivo}/72')
                            primeira_linha = file.readline()
                        
                        # Detecta automaticamente o delimitador
                        delimitador = ";" if ";" in primeira_linha else ","

                        df = pd.read_csv(arquivo, delimiter=delimitador, engine="python", encoding="latin1", on_bad_lines="skip")
                        lista_dfs.append(df)
                        numArquivo+=1
                    except Exception as e:
                        print(f"Erro ao ler {arquivo}: {e}")

                if lista_dfs:
                    df_final = pd.concat(lista_dfs, ignore_index=True, sort=False)
                    caminho_saida = os.path.join(destino, f"{pasta_fl}_completo.csv")
                    print(f'Consolidando para {caminho_saida}')
                    df_final.to_csv(caminho_saida, index=False, encoding="latin1", sep=delimitador)
                    print(f"Arquivo consolidado criado: {caminho_saida}")

# Exemplo de uso
pasta_principal = r"T:\COMUM\1 - PUBLICO\GABRIEL.ALVISE\Para\Erison\TUPAN"
destino = r"T:\COMUM\1 - PUBLICO\GABRIEL.ALVISE\Para\Erison\Consolidados"
juntar_csvs_em_cada_fl(pasta_principal, destino)
