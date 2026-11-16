import pandas as pd
import smtplib
import os
import keyring

from email.message import EmailMessage

# ==================================================
# 1. CAMINHOS DOS ARQUIVOS
# ==================================================

# Caminho do Excel conectado ao Power BI Dataflow
df = pd.read_excel(r'\\VCICWBSRVFLS02\Unidade_G\TI\POWER BI OFICIAL\RELATORIOS\RelatorioVendedores.xlsx')

# Caminho do relatório que será gerado
arquivo = r"\\VCICWBSRVFLS02\Unidade_G\TI\POWER BI OFICIAL\RELATORIOS\RelatorioNFs.xlsx"

# ==================================================
# 2. LEITURA E EXPORTAÇÃO DOS DADOS
# ==================================================

df.to_excel(
    arquivo,
    index=False,
    sheet_name="fato_Relatorio"
)

print("Arquivo Excel gerado com sucesso!")

# ==================================================
# 3. CONFIGURAÇÕES DO E-MAIL
# ==================================================

SMTP_SERVER = "mail.vitao.com.br"
SMTP_PORT = 587

EMAIL_REMETENTE = "bi@vitao.com.br"
EMAIL_DESTINATARIO = "lucas.mazutti@vitao.com.br"

# Recupera a senha do Gerenciador de Credenciais do Windows
EMAIL_SENHA = keyring.get_password(
    "BI_Relatorio_Email",
    "bi@vitao.com.br"
)

if not EMAIL_SENHA:
    raise Exception(
        "Senha do e-mail não encontrada no Gerenciador de Credenciais."
    )

# ==================================================
# 4. CRIAÇÃO DA MENSAGEM
# ==================================================

msg = EmailMessage()

msg["Subject"] = "Teste - Relatório de Faturamento"
msg["From"] = "bi@vitao.com.br"
msg["To"] = "lucas.mazutti@vitao.com.br"

msg.set_content("""
Bom dia!

Segue em anexo o relatório de faturamento para teste.

Atenciosamente,
Equipe BI
""")

# ==================================================
# 5. ANEXAR ARQUIVO EXCEL
# ==================================================

with open(arquivo, "rb") as anexo:
    msg.add_attachment(
        anexo.read(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(arquivo)
    )

# ==================================================
# 6. ENVIO DO E-MAIL
# ==================================================

try:
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as servidor:

        # Inicia a conexão criptografada
        servidor.starttls()

        # Autentica a conta de e-mail
        servidor.login(EMAIL_REMETENTE, EMAIL_SENHA)

        # Envia o e-mail com o Excel anexado
        servidor.send_message(msg)

    print("E-mail enviado com sucesso!")

except Exception as erro:
    print("Erro ao enviar e-mail:", erro)
    raise
