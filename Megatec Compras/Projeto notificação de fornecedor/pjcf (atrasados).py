import pandas as pd
import win32com.client as win32
from datetime import datetime, timedelta

caminho_nfpa = r"C:\Users\frodrigues\Downloads\Python Pasta\Megatec Compras\Projeto notificação de Fornecedor\nfpa.xlsx"
caminho_usuarios = r"C:\Users\frodrigues\Downloads\Python Pasta\Megatec Compras\Projeto notificação de Fornecedor\nfpa.xlsx"

df = pd.read_excel(caminho_nfpa, sheet_name="Pedidos")
usuarios_df = pd.read_excel(caminho_usuarios, sheet_name="Usuários")

df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')

vendedores_df = usuarios_df[["Usuarios.2", "e-mails.2"]].dropna()
vendedores_dict = vendedores_df.groupby("Usuarios.2")["e-mails.2"].apply(list).to_dict()

df_filtrado = df[df["Data em atraso"] >= 3]
df_filtrado = df_filtrado[df_filtrado["Pre-Notas"].isna() & df_filtrado["Pre-Notas2"].isna()]

outlook = win32.Dispatch("outlook.application")

data_limite = datetime.today() + timedelta(days=3)

coligadas = [
    "RANDON S/A IMPLEMENTOS E PARTICIPACOES",
    "RANDONCORP S.A.",
    "RANDONCORP S/A",
    "FRAS-LE SA",
    "CASTERTECH FUNDICAO E TECNOLOGIA LTDA"
]

for fornecedor, dados in df_filtrado.groupby("Nome Fornec"):

    fornecedor_normalizado = str(fornecedor).strip().upper()
    if fornecedor_normalizado in [c.upper() for c in coligadas]:
        print(f"⏭️ Ignorado fornecedor coligada: {fornecedor}")
        continue

    email_fornecedor = dados["E-mail Forn"].iloc[0]
    email_comprador = dados["E-mail Comprador"].iloc[0]
    email_almoxarife = dados["E-mail Almoxarifado"].iloc[0]

    if pd.isna(email_fornecedor):
        print(f"⚠️ Fornecedor sem e-mail: {fornecedor}")
        continue

    codigo_filial = str(dados["Filial"].iloc[0]).strip()
    if codigo_filial == "0105-MEGATEC INDUSTRIA E COMERCIO":
        nome_filial = "Megatec Industria"
    else:
        nome_filial = "Megatec Randon"

    vendedores_emails = vendedores_dict.get(codigo_filial, [])

    total_pedidos = int(dados["Numero"].nunique())
    atraso_medio = int(dados["Data em atraso"].mean())
    maior_atraso = int(dados["Data em atraso"].max())

# Tabela em HTML

    tabela_html = """
    <table border="2" cellspacing="0" cellpadding="5" style="border-collapse: collapse; font-family: Arial; font-size: 12px;">
        <tr style="background-color: #001CFF;">
            <th>Código Fornecedor</th>
            <th>Pedido</th>
            <th>Filial</th>
            <th>Produto</th>
            <th>Descrição</th>
            <th>Quantidade</th>
            <th>Qtd Entregue</th>
            <th>Qtd Pendente</th>
            <th>Unidade De Med.</th>
            <th>Data Prevista</th>
            <th>Dias em atraso</th>
        </tr>
    """
    for _, row in dados.iterrows():
        numero = row.get('Numero', '')
        if pd.isna(numero):
            numero = "nan"

        tabela_html += f"""
        <tr>
            <td>{row.get('C.Prod Forne', '')}</td>
            <td>{numero}</td>
            <td>{row.get('Filial', '')}</td>
            <td>{row.get('Produto', '')}</td>
            <td>{row.get('Desc Interna', '')}</td>
            <td>{row.get('Quantidade', '')}</td>
            <td>{row.get('Qtd.Entregue', '')}</td>
            <td>{row.get('Saldo', '')}</td>
            <td>{row.get('Unidade', '')}</td>
            <td>{row['Prev Entrega'].strftime('%d/%m/%Y') if pd.notna(row.get('Prev Entrega')) else ''}</td>
            <td>{row.get('Data em atraso', '')}</td>
        </tr>
        """

    loja = int(dados["Loja"].iloc[0])

    tabela_html += "</table>"

# Corpo em HTML

    corpo_html = f"""
    <p>Prezado {fornecedor}, Loja {loja}</p>
    <p>Identificamos que existem pedidos em atraso conforme resumo abaixo:</p>
    <ul>
        <li>Total de itens: {total_pedidos}</li>
        <li>Atraso médio: {atraso_medio:.0f} dias</li>
        <li>Maior atraso: {maior_atraso} dias</li>
    </ul>
    {tabela_html}
    <p>Pedimos sua manifestação sobre estes atrasos até 3 dias úteis.<p> 
    <p>Atenciosamente,<br>
    <b>Equipe de Compras - {nome_filial}</b></p>
    """

    mail = outlook.CreateItem(0)
    mail.To = email_fornecedor

    copia_emails = set(vendedores_emails)
    if pd.notna(email_comprador):
        copia_emails.add(str(email_comprador))
    if pd.notna(email_almoxarife):
        copia_emails.add(str(email_almoxarife))
    copia_emails.add("marco.ramiro@megatec.com.br")

    mail.CC = ";".join(copia_emails)
    mail.Subject = f"{nome_filial} | Notificação de pedidos em atraso – {fornecedor} – {datetime.today().strftime('%d/%m/%Y')}"
    mail.HTMLBody = corpo_html

    mail.Send()  

print("✅ E-mails preparados com sucesso!")

# Made By Felipe Wender
# Jesus Love You!
