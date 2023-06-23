import pandas as pd

# Caminho do arquivo Excel
caminho_arquivo = '/Users/antonioordones/Downloads/Testes/Teste.xlsm'

# Número da coluna a ser lida
numero_coluna = int(input("Digite o número da coluna a ser lida: "))

# Lista de palavras a serem buscadas
palavras_buscadas = input("Digite as palavras a serem buscadas, separadas por vírgula: ").split(',')

# Carregar o arquivo Excel
df = pd.read_excel(caminho_arquivo)

# Nome da coluna que será lida
coluna_lida = df.columns[numero_coluna - 1]

# Criar uma nova coluna para o resultado
df['Resultado'] = df[coluna_lida].apply(lambda x: 'SIM' if any(palavra.lower() in str(x).lower() for palavra in palavras_buscadas) else 'NÃO')

# Criar um novo arquivo Excel com os resultados
novo_caminho_arquivo = '/Users/antonioordones/Downloads/Testes/NovoTeste.xlsx'
df.to_excel(novo_caminho_arquivo, index=False)

print("Segmentação concluída. Os resultados foram salvos como uma nova coluna no arquivo")
