import xlwings as xw
from datetime import datetime
import os

# Caminho do arquivo principal
arquivo = r"C:\Users\frodrigues\Desktop\Python\Cópia de Modelo Almoxarifado.xlsx"

# Pasta onde vai salvar os arquivos separados
pasta_saida = r"C:\Users\frodrigues\Desktop\Python\Filiais"

# Garante que a pasta existe
if not os.path.exists(pasta_saida):
    os.makedirs(pasta_saida)

# Data no formato YYYYMMDD
data_hoje = datetime.today().strftime("%Y%m%d")

# Abre o Excel de forma invisível
app = xw.App(visible=False)

try:
    wb = app.books.open(arquivo)
    ws = wb.sheets["Pedidos"]

    # Descobre todas as filiais únicas (coluna "Filial")
    tabela_range = ws.range("A1").expand("table")
    tabela = tabela_range.value
    header = tabela[0]
    
    try:
        coluna_filial_idx = header.index("Filial")
    except ValueError:
        print("ERRO: A coluna 'Filial' não foi encontrada no cabeçalho.")
        app.quit()
        exit()

    filiais = sorted(set(linha[coluna_filial_idx] for linha in tabela[1:] if linha[coluna_filial_idx]))

    # Lista de filiais que DEVEM ser salvas (sem "Weslley")
    filiais_para_salvar = [f for f in filiais if "weslley" not in str(f).lower()]

    # Gera um arquivo para cada filial
    for filial in filiais_para_salvar:
        print(f"Processando filial: {filial}...")
        
        # Remove filtro anterior
        if ws.api.AutoFilterMode:
            ws.api.AutoFilterMode = False

        # Aplica filtro na coluna da filial
        # O Field do AutoFilter é baseado em 1, não em 0
        tabela_range.api.AutoFilter(Field=coluna_filial_idx + 1, Criteria1=filial)

        # --- INÍCIO DA CORREÇÃO ---

        # 1. Cria um novo arquivo Excel em branco
        wb_novo = app.books.add()
        ws_novo = wb_novo.sheets[0]
        ws_novo.name = f"Dados_{filial}" # Opcional: Renomeia a aba

        # 2. Copia apenas as células visíveis da tabela filtrada
        tabela_range.special_cells(xw.constants.SpecialCellType.xlCellTypeVisible).copy()
        
        # 3. Cola os dados na nova planilha
        ws_novo.range("A1").paste()

        # Ajusta a largura das colunas no novo arquivo (opcional, mas recomendado)
        ws_novo.autofit()

        # --- FIM DA CORREÇÃO ---

        # Limpa o nome da filial — ex: "0101 - Megatec Araçatuba" → "aracatuba"
        # A lógica abaixo pode estar resultando nos nomes "c", "d", etc. se o nome da filial for algo como "Filial C"
        nome_limpo = str(filial).split()[-1].lower().replace(' ', '_').replace('-', '_')

        # Caminho de saída
        nome_arquivo = f"mata121_{nome_limpo}_{data_hoje}.xlsx"
        caminho_saida = os.path.join(pasta_saida, nome_arquivo)

        # Salva e fecha o NOVO arquivo
        wb_novo.save(caminho_saida)
        wb_novo.close()
        
        print(f"✅ {nome_arquivo} salvo com sucesso!")

finally:
    # Fecha tudo para garantir que o processo do Excel não fique aberto
    if 'wb' in locals() and wb:
        wb.close(SaveChanges=False)
    app.quit()

print(f"\n✨ Planilhas das filiais salvas com sucesso em: {pasta_saida}")