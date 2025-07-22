import os
import re
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Pastas
html_folder = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\styled_htmls"
excel_folder = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\temp_file"

# Mapeamento de nome de arquivo para título
siglas = {
    "extremo_norte": "ONMN - EXTREMO NORTE",
    "freguesia": "ONMF - FREGUESIA",
    "gopoúva": "ONMG - GOPOÚVA",
    "pimentas": "ONMM - PIMENTAS",
    "pirituba": "ONMP - PIRITUBA",
    "santana": "ONMS - SANTANA",
    "onoa": "ONOA - CARTEIRA DE ABASTECIMENTO"
}

# Lista de destinatários
destinatarios = [
    "eeyamazaki@sabesp.com.br",
    "poliveira.eficien@sabesp.com.br",
    "cridolpho@sabesp.com.br",
    "jcrodrigues@sabesp.com.br",
    "scribeiro@sabesp.com.br",
    "santoscarolina@sabesp.com.br",
    "dmsilva2@sabesp.com.br",
    "rrcamargo@sabesp.com.br",
    "alvesricardo@sabesp.com.br",
    "fjpinto@sabesp.com.br",
    "kfernandes@sabesp.com.br",
    "juliomarques@sabesp.com.br",
    "cyokoi@sabesp.com.br",
    "claudioalves@sabesp.com.br",
    "rgois@sabesp.com.br",
    "andersoncosta@sabesp.com.br",
    "jladeia@sabesp.com.br",
    "rpolonio@sabesp.com.br",
    "wofernandes@sabesp.com.br",
    "afabiano@sabesp.com.br",
    "cfernanda@sabesp.com.br",
    "emasilva@sabesp.com.br",
    "fredericolima@sabesp.com.br",
    "tgsilva@sabesp.com.br",
    "josiassilva@sabesp.com.br",
    "ialbuquerque@sabesp.com.br",
    "asalvador@sabesp.com.br",
    "denisnascimento@sabesp.com.br",
    "jfgois@sabesp.com.br",
    "rcarmo@sabesp.com.br",
    "vsena@sabesp.com.br",
    "jpmruiz@sabesp.com.br",
    "asennes@sabesp.com.br",
    "alebarreto@sabesp.com.br",
    "abalbino@sabesp.com.br",
    "renatadanielle@sabesp.com.br",
    "mrmsilva@sabesp.com.br",
    "silvarf@sabesp.com.br",
    "jbedin@sabesp.com.br",
    "mjesus@sabesp.com.br",
    "asluz@sabesp.com.br",
    "kroque@sabesp.com.br",
    "mhanjos@sabesp.com.br",
    "andervalcardoso@sabesp.com.br",
    "consultargus@gmail.com",
    "veroliveira@sabesp.com.br",
    "wfavaretto@sabesp.com.br",
    "lhonorato@sabesp.com.br",
    "aregys@sabesp.com.br",
    "irubin@sabesp.com.br",
    "isabellaoliveira@sabesp.com.br",
    "alansantos@sabesp.com.br",
    "ebmonteiro@sabesp.com.br",
    "adrianoalves@sabesp.com.br",
    "vcsilva@sabesp.com.br",
    "kelli.santana@consorcionorte.com.br",
    "ana.paula@consorcionorte.com.br",
    "aoteles@sabesp.com.br",
    "lcgaraujo@sabesp.com.br",
    "mpalhau@sabesp.com.br",
    "mfmendes@sabesp.com.br",
    "amaraglia@sabesp.com.br",
    "jvitalino@sabesp.com.br",
    "fsartorato@sabesp.com.br",
    "wesley@jobeng.com.br",
    "jessica@jobeng.com.br",
    "junior@jobeng.com.br",
    "pmmelo@sabesp.com.br",
    "mlsilva.saae@sabesp.com.br",
    "miriamyendo@sabesp.com.br",
    "edsonmacedo@sabesp.com.br",
    "luanasantana@sabesp.com.br",
    "isabelle.oliveira@consorcionorte.com.br",
    "gustavo.silva@consorcionorte.com.br",
    "sidinei.ferreira@consorcionorte.com.br",
    "gideon.reis@consorcionorte.com.br",
    "davi@jobeng.com.br",
    "prguilhem@sabesp.com.br",
    "joao.regonha@consorcionorte.com.br",
    "carlos.ferreira@consorcionorte.com.br",
    "hcsouza@sabesp.com.br",
    "silvaalex@sabesp.com.br",
    "lasouza@sabesp.com.br",
    "cerci@sabesp.com.br",
    "atcarmo@sabesp.com.br",
    "sgrillo@sabesp.com.br",
    "regina.dias@consorciogopouva.com.br",
    "elizeu.lima@consorciogopouva.com.br",
    "rcqueiroz@sabesp.com.br",
    "vcastro@sabesp.com.br",
    "lizandra.souza@consorcionorte.com.br"

]

# CSS de estilo aplicado a todas as tabelas
css_style = """
<style>
table {
  border: 1px solid #1C6EA4;
  background-color: #EEEEEE;
  width: 100%;
  text-align: left;
  border-collapse: collapse;
}
table td, table th {
  border: 1px solid #AAAAAA;
  padding: 3px 2px;
}
table tbody td {
  font-size: 13px;
  text-align: right;
}
table tbody td:first-child {
  text-align: left;
}
table thead {
  background: #5FAAD9;
  border-bottom: 2px solid #444444;
  text-align: center;
}
table thead th {
  font-size: 15px;
  font-weight: bold;
  color: #000000;
  border-left: 2px solid #D0E4F5;
}
table thead th:first-child {
  border-left: none;
}
</style>
"""

# Variável para acumular estilos dos arquivos .txt
css_style_interno = ""

# Lista para armazenar blocos de HTML por polo
blocos_html = []

# Processa os arquivos .txt
for nome_arquivo in sorted(os.listdir(html_folder)):
    if nome_arquivo.endswith(".txt") and "operação_de_água_norte" not in nome_arquivo.lower():
        nome_base = nome_arquivo.replace(".txt", "")
        chave = nome_base.lower()

        # Título da seção
        titulo = siglas.get(chave, nome_base.upper())

        # Lê e separa CSS interno
        with open(os.path.join(html_folder, nome_arquivo), "r", encoding="utf-8") as f:
            html_conteudo = f.read()

        style_match = re.search(r"<style.*?>.*?</style>", html_conteudo, flags=re.DOTALL)
        if style_match:
            css_style_interno += style_match.group(0)

        html_sem_style = re.sub(r"<style.*?>.*?</style>", "", html_conteudo, flags=re.DOTALL)
        blocos_html.append(f"<p><b>{titulo}:</b></p>{html_sem_style}<br><br>")

# Combina os estilos no topo e inicia corpo do e-mail
corpo_html = css_style + css_style_interno + "<p><b>Segue carteira de serviços operacionais:</b></p>"
corpo_html += "".join(blocos_html)

# Rodapé
corpo_html += """
<p><i>*Consultem sempre o arquivo em Excel para obter mais detalhes dos serviços.<br>
*Serviços retirados do WFM com status pendentes.<br>
*Trata-se de um e-mail automático favor não responder.</i></p>
"""

# SMTP SETTINGS
smtp_host = 'smtp.office365.com'
smtp_port = 587
smtp_user = 'poliveira.eficien@sabesp.com.br'
smtp_pass = 'zpzvbdffbmhscqsz'  # senha do aplicativo

# Monta o e-mail
msg = MIMEMultipart()
msg['From'] = smtp_user
msg['To'] = ", ".join(destinatarios)
msg['Subject'] = "Carteira de Serviços Operacionais"

msg.attach(MIMEText(corpo_html, 'html'))

# Anexa os arquivos Excel
anexos = {
    "carteira.xlsx": "carteira_polos.xlsx",
    "df_abastecimento.xlsx": "carteira_abastecimento.xlsx"
}

for original, novo_nome in anexos.items():
    caminho = os.path.join(excel_folder, original)
    if os.path.exists(caminho):
        with open(caminho, 'rb') as f:
            part = MIMEApplication(f.read(), Name=novo_nome)
        part['Content-Disposition'] = f'attachment; filename="{novo_nome}"'
        msg.attach(part)

# Envia via SMTP
with smtplib.SMTP(smtp_host, smtp_port) as server:
    server.starttls()
    server.login(smtp_user, smtp_pass)
    server.send_message(msg)
    print("✅ E-mail enviado com sucesso via SMTP Office365!")
