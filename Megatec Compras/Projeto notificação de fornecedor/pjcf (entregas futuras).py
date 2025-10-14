import pandas as pd
import win32com.client as win32
from datetime import datetime, timedelta

# Caminhos das planilhas
caminho_nfpa = r"C:\Users\frodrigues\Downloads\Python Pasta\Megatec Compras\Projeto notificação de Fornecedor\nfpa.xlsx"
caminho_usuarios = r"C:\Users\frodrigues\Downloads\Python Pasta\Megatec Compras\Projeto notificação de Fornecedor\nfpa.xlsx"

df = pd.read_excel(caminho_nfpa, sheet_name="Pedidos")
usuarios_df = pd.read_excel(caminho_usuarios, sheet_name="Usuários")

df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')

df['Nome Fornec'] = df['Nome Fornec'].astype(str).str.strip().str.upper()
df['E-mail Forn'] = df['E-mail Forn'].astype(str).str.strip()

vendedores_df = usuarios_df[["Usuarios.2", "e-mails.2"]].dropna()
vendedores_dict = vendedores_df.groupby("Usuarios.2")["e-mails.2"].apply(list).to_dict()

hoje = datetime.today()
limite = hoje + timedelta(days=15)

df["Prev Entrega"] = pd.to_datetime(df["Prev Entrega"], errors="coerce")
df_filtrado = df[(df["Prev Entrega"] <= limite) & (df["Prev Entrega"] >= hoje)]
df_filtrado = df_filtrado[df_filtrado["Pre-Notas"].isna() & df_filtrado["Pre-Notas2"].isna()]

outlook = win32.Dispatch("outlook.application")

# Fornecedores coligadas ignorados
coligadas = [
    "RANDON S/A IMPLEMENTOS E PARTICIPACOES",
    "RANDONCORP S.A.",
    "RANDONCORP S/A",
    "FRAS-LE SA",
    "CASTERTECH FUNDICAO E TECNOLOGIA LTDA"
]
coligadas = [c.upper() for c in coligadas]

for (fornecedor, filial), dados in df_filtrado.groupby(["Nome Fornec", "Filial"]):

    fornecedor_normalizado = str(fornecedor).strip().upper()
    if fornecedor_normalizado in coligadas:
        print(f"⏭️ Ignorado fornecedor coligada: {fornecedor}")
        continue

    email_fornecedor = dados["E-mail Forn"].iloc[0]
    email_comprador = dados.get("E-mail Comprador", pd.Series([None])).iloc[0]
    email_almoxarife = dados.get("E-mail Almoxarifado", pd.Series([None])).iloc[0]

    codigo_filial = str(filial).strip()
    if codigo_filial == "0105-MEGATEC INDUSTRIA E COMERCIO":
        nome_filial = "Megatec Industria"
    else:
        nome_filial = "Megatec Randon"

    vendedores_emails = vendedores_dict.get(codigo_filial, [])
    total_pedidos = dados["Numero"].nunique()

    # Montar tabela HTML
    tabela_html = """
    <table border="1" cellspacing="0" cellpadding="5" style="border-collapse: collapse; font-family: Arial; font-size: 12px;">
        <tr style="background-color: #001CFF;">
            <th>Filial</th>
            <th>Pedido</th>
            <th>Produto</th>
            <th>Código Fornecedor</th>
            <th>Descrição</th>
            <th>Quantidade</th>
            <th>Unidade De Med.</th>
            <th>Data Prevista</th>
        </tr>
    """

    for _, row in dados.iterrows():
        numero = row.get('Numero', '')
        if pd.isna(numero):
            numero = "nan"

        tabela_html += f"""
        <tr>
            <td>{row.get('Filial', '')}</td>
            <td>{numero}</td>
            <td>{row.get('Produto', '')}</td>
            <td>{row.get('C.Prod Forne', '')}</td>
            <td>{row.get('Desc Interna', '')}</td>
            <td>{row.get('Quantidade', '')}</td>
            <td>{row.get('Unidade', '')}</td>
            <td>{row['Prev Entrega'].strftime('%d/%m/%Y') if pd.notna(row.get('Prev Entrega')) else ''}</td>
        </tr>
        """

    loja = int(dados["Loja"].iloc[0])
    tabela_html += "</table>"

    corpo_html = f"""
    <p>Prezado {fornecedor}, Loja {loja}</p>
    <p>Identificamos que alguns pedidos possuem <b>entregas previstas para os próximos 15 dias</b> referentes à filial <b>{codigo_filial}</b>.<br>
    Gostaríamos de confirmar se está tudo certo para a chegada desses materiais:</p>

    {tabela_html}

    <p>Caso tenha algo divergente com a data de previsão de entrega e com a data real de entrega, por gentileza nos comunique imediatamente.</p>
    <p>Pedimos sua confirmação em até 3 dias úteis.</p>
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
    mail.Subject = f"{nome_filial} | Confirmação de pedidos futuros – {fornecedor} – {codigo_filial} – {hoje.strftime('%d/%m/%Y')}"
    mail.HTMLBody = corpo_html

    mail.Save()



print("📨 E-mails preparados com sucesso, separados por fornecedor e filial!")
print("✝️ Made by Felipe Wender — Jesus Love You!")
