import time
import subprocess
from datetime import datetime

# Caminho do Python e do main.py
caminho_python = r"C:\Python312\python.exe"
caminho_main = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\main_ON.py"

# Horários alvo no formato HH:MM
horarios_programados = {"06:20", "17:30"}
executados_hoje = set()

print("⏱️ Scheduler iniciado. Aguardando horários programados...")

while True:
    agora = datetime.now()
    hora_atual = agora.strftime("%H:%M")

    if hora_atual in horarios_programados and hora_atual not in executados_hoje:
        print(f"🟢 Executando main.py às {hora_atual}...")
        try:
            subprocess.run([caminho_python, caminho_main], check=True)
            print(f"✅ main.py executado com sucesso às {hora_atual}")
        except subprocess.CalledProcessError as e:
            print(f"❌ Erro ao executar main.py: {e}")
        executados_hoje.add(hora_atual)

    # Resetar após meia-noite
    if hora_atual == "00:00":
        executados_hoje.clear()

    time.sleep(60)
