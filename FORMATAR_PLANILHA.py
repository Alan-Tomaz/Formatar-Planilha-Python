import pandas as pd
import re # Importar a biblioteca de expressões regulares
from unidecode import unidecode  # Biblioteca para remover acentos
from datetime import datetime

# Caminho do arquivo da planilha
caminho_planilha = 'CLIENTES.xlsx'

diretorio_original = r""

prefixo_do_arquivo = "TEXT_"

# Capturar a data de uma string ou diretamente
data_atual = datetime.now() # Data atual

# Formatar a data no formato DDMMAAAA
data_formatada = data_atual.strftime("%d%m%Y")

# Nome do arquivo de texto que será gerado
arquivo_texto = fr'{diretorio_original}{prefixo_do_arquivo}_{data_formatada}'

# Carregar os dados da planilha
df = pd.read_excel(caminho_planilha)

# Formatar os dados e salvar no arquivo de texto
with open(arquivo_texto, 'w', encoding='utf-8') as f:
    for index, row in df.iterrows():

        cliente = str(row['CLIENTE'])  # Converter a linha CLIENTE para string caso não seja
        cliente = unidecode(cliente)  # Remover acentos
        cliente = re.sub(r'[^a-zA-Z0-9\s/.&-]', '', cliente)  # Permitir letras, números, espaços, /, ., - e &

        cnpj = str(row['CNPJ']).replace(".", "").replace("-", "").replace("/", "")  # Converter a linha CNPJ para string, caso não seja e remove traços, pontos e barras

         # Verificar o tamanho da string na coluna "TELEFONE"
        telefone = str(row['TELEFONE']).replace(" ", "")  # Converter a linha telefone para string, caso não seja e remove espaços em branco
        if len(telefone) == 10: 
            telefone += " "  # Adicionar um espaço em branco ao final caso seja número fixo

        # Criar a linha formatada
        linha_formatada = f"{telefone}{cnpj}{cliente}\n"
        f.write(linha_formatada)

print(f"Arquivo de texto '{arquivo_texto}' gerado com sucesso!")