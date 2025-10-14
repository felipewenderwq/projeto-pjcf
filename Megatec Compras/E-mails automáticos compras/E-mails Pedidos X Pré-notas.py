import win32com.client as win32
import os
import glob

# Caminho base onde ficam as pastas
pasta_base = r"C:\Users\frodrigues\Desktop\mataNOVO"

# Filiais e caminhos dos arquivos
filiais = {
    # Enviadas para o Weslley
    "weslley.henrique@megatec.com.br": [
        os.path.join(pasta_base, "Weslley", "mata121_cristalina_*.xlsx"),
        os.path.join(pasta_base, "Weslley", "mata121_itumbiara_*.xlsx"),
        os.path.join(pasta_base, "Weslley", "mata121_rioverde_*.xlsx"),
        os.path.join(pasta_base, "Weslley", "mata121_santavitoria_*.xlsx"),
        os.path.join(pasta_base, "Weslley", "mata121_uberlandia_*.xlsx"),
    ],

    # Demais filiais
    "almoxarifado.industria@megatec.com.br": [os.path.join(pasta_base, "Industria", "mata121_*.xlsx")],
    "almoxarifado.prudente@megatec.com.br": [os.path.join(pasta_base, "Prudente", "mata121_*.xlsx")],
    "almoxarifado.aracatuba@megatec.com.br": [os.path.join(pasta_base, "Araçatuba", "mata121_*.xlsx")],
    "almoxarifado.andradina@megatec.com.br": [os.path.join(pasta_base, "Andradina", "mata121_*.xlsx")],
}

# Vendedores por filial (chave = nome da filial)
vendedores = {
    "prudente": [
        "sandro.viana@megatec.com.br",
        "raphael.anderson@megatec.com.br",
        "luciano.correa@megatec.com.br",
        "lucas.melo@megatec.com.br"
    ],
    "uberlandia": [
        "joao.alves@megatec.com.br",
        "davi.lima@megatec.com.br",
        "rodrigo.duarte@megatec.com.br",
        "eliar.souza@megatec.com.br",
        "adriano.luiz@megatec.com.br",
        "gustavo.nobrega@megatec.com.br",
        "ilenicio.junior@megatec.com.br"
    ],
    "cristalina": [
        "alexandre.santana@megatec.com.br",
        "david.ribeiro@megatec.com.br"
    ],
    "aracatuba": [
        "jonathas.prado@megatec.com.br",
        "saymon.saraiva@megatec.com.br",
        "leonardo.tavares@megatec.com.br",
        "leandro.augusto@megatec.com.br",
        "adm.pecas@megatec.com.br",
        "matheus.carreira@megatec.com.br",
        "jose.marinho@megatec.com.br"
    ],
    "rioverde": [
        "wallison.rufino@megatec.com.br",
        "edgar.guimaraes@megatec.com.br",
        "nayara.camila@megatec.com.br",
        "gustavo.goncalves@megatec.com.br",
        "jean.oliveira@megatec.com.br",
        "dione.henrique@megatec.com.br"
    ],
    "itumbiara": [
        "dione.henrique@megatec.com.br",
        "wallison.rufino@megatec.com.br",
        "edgar.guimaraes@megatec.com.br",
        "nayara.camila@megatec.com.br",
        "gustavo.goncalves@megatec.com.br",
        "jean.oliveira@megatec.com.br"
    ],
    "industria": [
        "mariana.nogara@megatec.com.br",
        "bianca.rosa@megatec.com.br",
        "julia.maria@megatec.com.br" 
    ]
}

# E-mail do gestor
cc_email = "marco.ramiro@megatec.com.br"

# Inicializa o Outlook
outlook = win32.Dispatch('outlook.application')

# Loop principal
for email_destino, padroes_arquivos in filiais.items():
    for padrao in padroes_arquivos:
        arquivos = glob.glob(padrao)
        if not arquivos:
            print(f"⚠️ Nenhum arquivo encontrado para {email_destino} ({padrao})")
            continue

        arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
        nome_arquivo = os.path.basename(arquivo_mais_recente)

        # Pega o nome da filial (ex: mata121_aracatuba_20251007.xlsx → aracatuba)
        filial_nome = nome_arquivo.split("_")[1].lower()

        # Cria e-mail
        mail = outlook.CreateItem(0)
        mail.To = email_destino

        # Busca vendedores da filial
        lista_cc = vendedores.get(filial_nome, [])
        if cc_email not in lista_cc:
            lista_cc.append(cc_email)

        mail.CC = ";".join(lista_cc)

        mail.Subject = f"Relatório diário | Pedidos X Pré-notas | {filial_nome.capitalize()}"

        corpo = f"""
        <p><b>Bom dia!</b></p>
        <p>Segue relatório da filial <b>{filial_nome.capitalize()}</b> para auxílio do estoque.</p>
        <p>Atenciosamente,<br><b>Felipe Wender</b></p>
        """

        mail.HTMLBody = corpo + mail.HTMLBody
        mail.Attachments.Add(arquivo_mais_recente)
        mail.Save()  # usa .Send() quando quiser enviar de vez

        print(f"✅ Relatório preparado para {email_destino} -> {nome_arquivo}")

print("\n✨ Todos os relatórios foram preparados com sucesso!")
