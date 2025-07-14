import os
import win32com.client as win32
import re
import time
import pythoncom  # biblioteca nativa para inicializar o COM corretamente

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
    "poliveira.eficien@sabesp.com.br"
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
    if nome_arquivo.endswith(".txt"):
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

# Garante que o Outlook está pronto
for tentativa in range(10):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch("Outlook.Application")
        break
    except Exception as e:
        print(f"⏳ Tentativa {tentativa+1}: Outlook ainda não está pronto...")
        time.sleep(2)
else:
    raise Exception("❌ Não foi possível iniciar o Outlook após várias tentativas.")

# Cria e envia o e-mail
mail = outlook.CreateItem(0)
mail.Subject = "Carteira de Serviços Operacionais - WFM"
mail.To = "; ".join(destinatarios)
mail.HTMLBody = corpo_html

# Anexa os arquivos Excel
anexos = {
    "carteira.xlsx": "carteira_polos.xlsx",
    "df_abastecimento.xlsx": "carteira_abastecimento.xlsx"
}

for original, novo_nome in anexos.items():
    caminho = os.path.join(excel_folder, original)
    if os.path.exists(caminho):
        mail.Attachments.Add(Source=caminho).DisplayName = novo_nome

# Envia
mail.Send()
print("✅ E-mail enviado com sucesso!")
