import os
import pandas as pd
import re

# Caminho da pasta contendo os relatórios
pasta = r"V:\Setor Robô\Scripts Python\Domínio\Gera Relatorio Portal do Empregado\execução\Mescla Relatorio"

# Função para remover caracteres inválidos
def remove_illegal_characters(df):
    for col in df.columns:
        df[col] = df[col].astype(str).apply(lambda x: re.sub(r'[\x00-\x1F\x7F-\x9F]', '', x))
    return df

# Inicializa uma lista para armazenar os DataFrames
dataframes = []

# Percorre todos os arquivos na pasta
for arquivo in os.listdir(pasta):
    if arquivo.endswith(".xls"):
        # Extrai o código do nome do arquivo
        codigo = arquivo.split(" - ")[0]

        # Caminho completo do arquivo
        caminho_arquivo = os.path.join(pasta, arquivo)

        # Lê o arquivo Excel
        df = pd.read_excel(caminho_arquivo, engine='xlrd')

        # Adiciona a coluna 'Código'
        df['Código Domínio'] = codigo

        # Limpa os caracteres inválidos
        df = remove_illegal_characters(df)

        # Adiciona o DataFrame à lista
        dataframes.append(df)

# Concatena todos os DataFrames em um único DataFrame
df_geral = pd.concat(dataframes, ignore_index=True)

# Salva o DataFrame combinado em um novo arquivo Excel
df_geral.to_excel("Relatorio_Geral.xlsx", index=False)
