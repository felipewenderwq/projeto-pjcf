import pandas as pd

# Caminho do arquivo original
arquivo_origem = r"C:\Users\frodrigues\Downloads\Python Pasta\Divisão mt121\mata121_Pedidos_de_Compras_+Cubo_01.06_2022_à_20251004.xlsx"
planilha_origem = "Pedidos de Compras"

# Lê a planilha original
df = pd.read_excel(arquivo_origem, sheet_name=planilha_origem)

# Normaliza os nomes das colunas: maiúsculas e sem espaços extras
df.columns = df.columns.str.strip().str.upper()

# Nomes desejados (chaves)
colunas_chave = ['NUMERO', 'FILIAL', 'PRODUTO']
# Coluna de data: vamos tentar encontrar automaticamente
coluna_data = None

# Procura por coluna que tenha 'DATA' no nome
for c in df.columns:
    if 'DATA' in c:
        coluna_data = c
        break

# Verifica se encontrou todas as colunas necessárias
for col in colunas_chave:
    if col not in df.columns:
        raise ValueError(f"❌ Coluna necessária não encontrada na planilha: {col}")
if coluna_data is None:
    raise ValueError("❌ Nenhuma coluna de data encontrada na planilha (esperado algo com 'DATA').")

# Seleciona apenas as colunas corretas
df = df[colunas_chave + [coluna_data]]

# Renomeia para nomes padronizados
df.columns = colunas_chave + ['DATA ENTRADA']

# Agrupa os dados e junta as entradas por pedido/produto
df_agrupado = (
    df.groupby(['NUMERO', 'FILIAL', 'PRODUTO'])['DATA ENTRADA']
    .apply(list)
    .reset_index()
)

# Descobre o número máximo de entradas em um pedido
max_entradas = df_agrupado['DATA ENTRADA'].apply(len).max()

# Cria colunas dinâmicas de Entrada1, Entrada2, etc.
for i in range(max_entradas):
    df_agrupado[f'Entrada{i+1}'] = df_agrupado['DATA ENTRADA'].apply(
        lambda x: x[i] if i < len(x) else None
    )

# Remove a lista original
df_final = df_agrupado.drop(columns=['DATA ENTRADA'])

# Exporta o resultado
arquivo_saida = r"C:\Users\frodrigues\Downloads\Python Pasta\pedidos_com_datas(2).xlsx"
df_final.to_excel(arquivo_saida, index=False)

print(f"✅ Arquivo gerado com sucesso: {arquivo_saida}")