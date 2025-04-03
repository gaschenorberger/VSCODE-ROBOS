import os
import pandas as pd

diretorio = r"T:\COMUM\1 - PUBLICO\GABRIEL.ALVISE\Para\Erison\TUPAN\FL 53"

dados = []

for arquivo in os.listdir(diretorio):
    if arquivo.endswith(".csv"): 
        caminho_arquivo = os.path.join(diretorio, arquivo)
        
        try:
            df = pd.read_csv(caminho_arquivo, nrows=1, encoding="ISO-8859-1", sep=";")  

            dados.append([arquivo, ", ".join(df.columns.tolist())])
        
        except Exception as e:
            print(f"Erro ao processar {arquivo}: {e}")

df_resultado = pd.DataFrame(dados, columns=["Arquivo", "Colunas"])

caminho_saida = os.path.join(diretorio, "Planilha_Validacao.xlsx")
df_resultado.to_excel(caminho_saida, index=False)

print(f"Validação salva em: {caminho_saida}")
