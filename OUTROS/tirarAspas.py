# Nome do arquivo original
arquivo_entrada = r'C:\Users\gabriel.alvise\Desktop\EMPRESAS\TERMOLAR\chaves.txt'
arquivo_saida = 'chaves2.txt'

numeros_unicos = set()

with open(arquivo_entrada, 'r') as entrada:
    for linha in entrada:
        numero = linha.strip().replace("'", "").replace('"', '')
        if numero:  # Ignora linhas vazias
            numeros_unicos.add(numero)

with open(arquivo_saida, 'w') as saida:
    for numero in sorted(numeros_unicos):
        saida.write(numero + '\n')

print('Aspas removidas e duplicados eliminados com sucesso!')