import os
import subprocess
import time
import logging
from selenium.common.exceptions import (
    TimeoutException,
    UnexpectedAlertPresentException,
    NoSuchElementException,
    StaleElementReferenceException,
    ElementClickInterceptedException
)

# === Função para limpar os arquivos de pastas usadas ===
def limpar_pastas(pastas):
    for pasta in pastas:
        if os.path.exists(pasta):
            for arquivo in os.listdir(pasta):
                caminho_arquivo = os.path.join(pasta, arquivo)
                try:
                    if os.path.isfile(caminho_arquivo) or os.path.islink(caminho_arquivo):
                        os.unlink(caminho_arquivo)
                    elif os.path.isdir(caminho_arquivo):
                        os.rmdir(caminho_arquivo)
                except Exception as e:
                    logging.error(f"Erro ao apagar {caminho_arquivo}: {e}")

# === Configuração de logging ===
error_log_path = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\error"
os.makedirs(error_log_path, exist_ok=True)

log_file = os.path.join(error_log_path, "error.log")
logging.basicConfig(filename=log_file, level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# === Limpeza das pastas antes da execução ===
pastas_para_limpar = [
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\error",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\downloads",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\n8n",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\polos",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\html_image",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\styled_htmls",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\temp_file",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\pendentes",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\lista_download",
]
limpar_pastas(pastas_para_limpar)

# === Lista de scripts a executar ===
scripts = [
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\1_wfm_listaos.py",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\2_download_relatorio.py",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\3_format_pendentes_ON.py",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\4_table_image.py",
    r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\5_send_email.py"
]

# === Execução dos scripts com controle de tempo e erro ===
overall_start_time = time.time()
total_execution_times = []

for script in scripts:
    script_start_time = time.time()

    retries = 5
    success = False
    while retries > 0:
        try:
            subprocess.run([r"C:\Python312\python.exe", script], check=True)
            success = True
            break
        except (subprocess.CalledProcessError, TimeoutException, UnexpectedAlertPresentException,
                NoSuchElementException, StaleElementReferenceException, ElementClickInterceptedException) as e:
            retries -= 1
            logging.error(f"Script failed: {script}. Retries left: {retries}. Error: {str(e)}")
            print(f"Script failed: {script}. Retries left: {retries}. Restarting...")

    script_end_time = time.time()
    elapsed = script_end_time - script_start_time
    total_execution_times.append(elapsed)

    if success:
        print(f"{script} executed in {int(elapsed // 60)} min {round(elapsed % 60, 2)} sec.")
    else:
        print(f"{script} failed after maximum retries.")
        logging.error(f"Script {script} failed after maximum retries.")

    time.sleep(10)

# === Tempo total ===
overall_elapsed_time = sum(total_execution_times)
print(f"Total time for all scripts: {int(overall_elapsed_time // 60)} min {round(overall_elapsed_time % 60, 2)} sec.")
