# Importing necessary libraries:
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from selenium.common.exceptions import (
    NoSuchElementException, TimeoutException, JavascriptException,
    StaleElementReferenceException, UnexpectedAlertPresentException, NoAlertPresentException
)
import pandas as pd
import os

# Setting prior conditions:
service = Service(r"chromedriver.exe")
options = webdriver.ChromeOptions()
options.add_argument("disable-infobars")
options.add_argument("start-maximized")
options.add_argument("disable-dev-shm-usage")
options.add_argument("no-sandbox")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_argument("disable-blink-features=AutomationControlled")

# Location of the list of names generated on WFM:
file_path = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\lista_download\lista_carteira.txt"

# Initialize a Chrome browser
driver = webdriver.Chrome(service=service, options=options)

# Navigate to the page
driver.get('https://geoprd-interno.sabesp.com.br/sabespwfm/')

# Setting time for WebDriverWait:
wait = WebDriverWait(driver, 20)

# Wait for the frames to load
wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'iframe')))

# Switch to the next frame
driver.switch_to.frame(1)  # 0-index based. If there are multiple frames, you may need to adjust the index

# Wait for the password field to load
wait.until(EC.presence_of_element_located((By.ID, 'USER')))

# Fill the password field
with open(r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\WFM\cred_wfm.txt", 'r') as file:
    login = file.readline().strip()
    password = file.readline().strip()

driver.find_element(By.ID, 'USER').send_keys(login)
driver.find_element(By.ID, 'INPUTPASS').send_keys(password)

try:
    # Click the login button
    driver.find_element(By.ID, 'submbtn').click()

    # Espera até aparecer ou o botão 'Lista OS' (login OK) ou a mensagem de erro
    WebDriverWait(driver, 20).until(
        lambda d: d.find_elements(By.CLASS_NAME, 'error-label') or
                  d.find_elements(By.XPATH, "//div[contains(text(), 'Lista OS')]")
    )

    # Verifica se houve erro de sessão ativa
    if driver.find_elements(By.CLASS_NAME, 'error-label'):
        error_message = driver.find_element(By.CLASS_NAME, 'error-label').text
        if "O usuário já está conectado" in error_message:
            print("Usuário já conectado. Selecionando 'Substituir Sessão'...")
            radio_buttons = driver.find_elements(By.XPATH, "//input[@type='RADIO' and @name='oooo']")
            for radio in radio_buttons:
                onclick_value = radio.get_attribute("onclick")
                if "form.CMD.value='RESTART'" in onclick_value:
                    radio.click()
                    enter_button = driver.find_element(By.CLASS_NAME, 'enterbutton')
                    ActionChains(driver).move_to_element(enter_button).click().perform()
                    print("Sessão substituída com sucesso.")
                    break
        else:
            print("Erro desconhecido de login:", error_message)
            driver.quit()
            exit()
    else:
        print("✅ Login realizado com sucesso.")

except TimeoutException:
    print("❌ Timeout: Login não confirmado. Página possivelmente mudou ou a estrutura está incorreta.")
    driver.quit()
    exit()

except NoSuchElementException as e:
    print(f"Error locating element: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")


# Define a list of your inputs (dropdowns from Lista OS search options)
# These are in the form (first_option, [list of second options])

ugr = ['ONMG - DIV MANUT SERV OPE GOPOÚVA', 'ONMM - DIV MANUT SERV OPE PIMENTAS',
       'ONMF - DIV MANUT SERV OPE FREGUESIA', 'ONMP - DIV MANUT SERV OPE PIRITUBA',
       'ONMN - DIV MANUT SERV OPE EXTREMO NORTE', 'ONOA - DIV OPERAÇÃO DE ÁGUA NORTE']
# Removido 'ONMB - DIV MANUT SERV OPE BRAGANÇA.
ato = [['70 ATC GOPOUVA'], ['71 ATC PIMENTAS'], [''],
       [''], [''], ['']]

# Wait for the ListaOS button:
# wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="TBB_tbm2"]/div[8]/div[2]')))
button_listaOS = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Lista OS')]")))

# Click on listaOS:
# driver.find_element(By.XPATH, '//*[@id="TBB_tbm2"]/div[8]/div[2]').click()
button_listaOS.click()
time.sleep(3)

# Wait for pesquisaOS:
wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="TBB_tbm2"]/div[1]/div[2]')))
driver.find_element(By.XPATH, '//*[@id="TBB_tbm2"]/div[1]/div[2]').click()  # Click on pesquisaOS

# Choose the first interaction for the dropdowns, after that we will include a loop to search for the rest.
ugr_item = 'ONMS - DIV MANUT SERV OPE SANTANA'
ato_item = ''

# Select the first option (UGR)
time.sleep(2)
dropdown_ugr = wait.until(EC.presence_of_element_located((By.NAME, '_lyAODLID_AFIL')))
select = Select(dropdown_ugr)
select.select_by_visible_text(ugr_item)
time.sleep(2)

# Select the second option (ATO)
dropdown_ato = wait.until(
    EC.presence_of_element_located((By.NAME, '_lyXSABAINDIDATC_XSABAETE')))
select2 = Select(dropdown_ato)
select2.select_by_visible_text(ato_item)
time.sleep(2)

# Find the span element using its text (Status da Operacao)
button_span = wait.until(
    EC.presence_of_element_located((By.XPATH, '//span[text()="Status da Operação"]')))
button_span.click()  # Click the span button
time.sleep(2)

# Click on pencil icon to search for avaiable status:
max_retries = 3
retries = 0

while retries < max_retries:
    try:
        lapis = wait.until(
            EC.presence_of_all_elements_located(
                (By.XPATH, '//img[@src="r/std/icons/control/edit55x64.png" and @alt="Editar"]'))
        )
        lapis[2].click()
        # If successful, break out of the loop
        break
    except TimeoutException:
        retries += 1
        if retries == max_retries:
            print(f"Failed to find the element after {max_retries} attempts.")
time.sleep(10)

# Wait for the button 'SELECIONAR TODOS' to be present using CSS selector
button_selectAll = wait.until(EC.element_to_be_clickable
                                                   ((By.CSS_SELECTOR, 'div.tvCaption > div > div:nth-child(2)')))
driver.execute_script("arguments[0].click();", button_selectAll)  # Use JavaScript to click the element
time.sleep(2)

# Find the 'Cancelada' cell and click on it
cancelada_cell = wait.until(
    EC.presence_of_element_located((By.XPATH, '//td[@class="tvCell" and text()="Cancelada"]')))
cancelada_cell.click()
time.sleep(2)

# Find the 'Fechada' cell and click on it
fechada_cell = wait.until(
    EC.presence_of_element_located((By.XPATH, '//td[@class="tvCell" and text()="Fechada"]')))
fechada_cell.click()
time.sleep(2)

# Find the source and destination to drag the elements into
source_element = wait.until(EC.presence_of_element_located((By.XPATH, "//td[text()='Planejada']")))
destination_element = wait.until(EC.presence_of_element_located((By.ID, "TVSCR-tvSelezionati")))

# Using JacaScript to force the movement of drag
script = """
var src=arguments[0],tgt=arguments[1];
var dataTransfer={dropEffect:'',effectAllowed:'all',files:[],items:{},types:[],setData:function(format,data){
this.items[format]=data;this.types.append(format);},getData:function(format){return this.items[format];},
clearData:function(format){}};
var emit=function(event,target){
var evt=document.createEvent('Event');
evt.initEvent(event,true,false);
evt.dataTransfer=dataTransfer;
target.dispatchEvent(evt);
};
emit('dragstart',src);
emit('drop',tgt);
emit('dragend',src);
"""

# Executing the script
driver.execute_script(script, source_element, destination_element)
time.sleep(2)

# Find the button 'Ok' by its class and text.
button_ok = wait.until(EC.presence_of_element_located(
    (By.XPATH, "//button[contains(@class, 'but butAct') and ./div[text()='Ok']]")))

# Click the button
button_ok.click()
time.sleep(2)

# Find the button 'Busca' by its class and text.
button_search = wait.until(EC.presence_of_element_located(
    (By.XPATH, "//button[contains(@class, 'but butSub butAct') and ./div[text()='Busca']]")))
driver.execute_script("arguments[0].click();", button_search)  # Click the button
time.sleep(2)

# Try to find either of the images
while True:
    try:
        # Wait for first image: if the search has more than 5k lines a popup window will appear, so we need to close it
        # if there is no popup it will move to the next block.
        first_image = wait.until(
            EC.presence_of_element_located((By.XPATH, "//img[contains(@class, 'icon Enabled icon_x')]")))
        driver.execute_script("arguments[0].click();", first_image)
    except TimeoutException:
        pass
    try:
        # Wait for second image: the second image is the Options icon to enable the file download.
        # second_image = wait.until(
        #     EC.presence_of_element_located((By.XPATH, "//img[contains(@class, 'icon Enabled icon_menu')]")))
        second_image = wait.until(EC.presence_of_element_located((By.XPATH,
                                                          '//img[contains(@class, "icon") and '
                                                          'contains(@class, "Enabled") and '
                                                          'contains(@class, "icon_menu") and @alt="Opções"]')))
        driver.execute_script("arguments[0].click();", second_image)
    except TimeoutException:
        pass
    try:
        excel = wait.until(
            EC.presence_of_element_located((By.XPATH, '//div[text()="Exportar para o Excel"]')))
        excel.click()
    except TimeoutException as ex:
        pass
    try:
        # wait for the input box to be visible and then find it
        input_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                                    'input.f.fN.fInside')))
        input_box.clear()  # clear it
        break
    except TimeoutException as ex:
        pass

# Get the current date and time in the desired formats
current_time = datetime.now().strftime('%Y%m%d%H%M')

# When you're naming the file, extract the name after 'ATO' from the second_option
file_name_part = ugr_item.split('OPE', 1)[-1].strip().lower().replace(' ', '_')
naming = "C-" + file_name_part + "-" + current_time
file_name = "33627-C-" + file_name_part + "-" + current_time

# Create an empty list to save the name of the file:
list_download = []
list_download.append(file_name)

# List to store the data for the DataFrame
export_data = []

# Function to register the export time
def register_export_time(file_name_part):
    export_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    export_data.append({'unidade': file_name_part, 'data': export_time})

# After clicking on Excel button, it will appear an alert box, we need to clear it to rename the file as we prefer.
try:
    # clear it and handle the alert
    input_box.clear()
    time.sleep(2)
    alert = wait.until(EC.alert_is_present())
    print("Alert message: " + alert.text)
    alert.accept()
    time.sleep(2)
except (UnexpectedAlertPresentException, NoAlertPresentException):
    print(f'Processando a carteira de: {file_name}...')
try:
    # try to type "pendentes"
    input_box.send_keys(f'{naming}')
except (UnexpectedAlertPresentException, NoAlertPresentException):
    print("Alert presented after typing 'pendentes'")


# Select the folder where the file will be generated
dropdown = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//select[@class='f fA fN']")))

# Initialize the Select object
select = Select(dropdown)

# Aguarda as opções ficarem disponíveis
wait.until(lambda d: len(select.options) > 1)

# Imprime todas as opções visíveis
print("Opções disponíveis no dropdown:")
for option in select.options:
    print(f"- '{option.text}'")

# Procurar a opção desejada de forma segura
found = False
for option in select.options:
    if "MNT - UG REGIONAL SANTANA" in option.text.strip():
        select.select_by_visible_text(option.text.strip())
        found = True
        print(f"Selecionado: {option.text.strip()}")
        break

if not found:
    print("⚠️ A opção 'MNT - UG REGIONAL SANTANA' não está disponível no dropdown!")


# Optionally, you can retrieve the selected option to verify
selected_option = select.first_selected_option
print("Pasta selecionada para gerar os arquivos:", selected_option.text)
time.sleep(2)

# Find and click the Ok button to export the file
button_export = wait.until(
    EC.presence_of_element_located((By.XPATH, '//button[contains(@onclick, "OkNameExport#")]')))
time.sleep(2)
driver.execute_script("arguments[0].click();", button_export)

# Register the export time
register_export_time(file_name_part)

# Find the 'X' image to close the popup
button_esc = wait.until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'img.icon.Enabled.icon_x[onclick="closeContextMenu();"]')))
driver.execute_script("arguments[0].click();", button_esc)
time.sleep(2)

# Re-open the filters options to search for the next UGR and ATO
button_filtro = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,
                                                           "//button[contains(@class, 'pbut') and "
                                                           "contains(@class, 'pbutoff') and "
                                                           "contains(@class, 'Enabled') and ./div[text()='Filtro']]")))
driver.execute_script("arguments[0].click();", button_filtro)
time.sleep(2)

# Start loop here:
for idx, ugr_item in enumerate(ugr):
    for ato_item in ato[idx]:

        # Select the first option (UGR)
        time.sleep(2)
        dropdown_ugr = wait.until(EC.presence_of_element_located((By.NAME, '_lyAODLID_AFIL')))
        select = Select(dropdown_ugr)
        select.select_by_visible_text(ugr_item)
        time.sleep(2)

        # Select the second option (ATO)
        dropdown_ato = wait.until(
            EC.presence_of_element_located((By.NAME, '_lyXSABAINDIDATC_XSABAETE')))
        select2 = Select(dropdown_ato)
        select2.select_by_visible_text(ato_item)
        time.sleep(2)

        # Find the button 'Busca' by its class and text.
        button_search = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//button[contains(@class, 'but butSub butAct') and ./div[text()='Busca']]")))
        driver.execute_script("arguments[0].click();", button_search)  # Click the button
        time.sleep(2)

        # Try to find either of the images
        while True:
            try:
                # Wait for first image: if the search has more than 5k lines a popup window will appear, so we need to close it
                # if there is no popup it will move to the next block.
                first_image = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//img[contains(@class, 'icon Enabled icon_x')]")))
                driver.execute_script("arguments[0].click();", first_image)
            except TimeoutException:
                pass
            try:
                # Wait for second image: the second image is the Options icon to enable the file download.
                # second_image = wait.until(
                #     EC.presence_of_element_located((By.XPATH, "//img[contains(@class, 'icon Enabled icon_menu')]")))
                time.sleep(5)
                second_image = wait.until(EC.presence_of_element_located((By.XPATH,
                                                                                               '//img[contains(@class, "icon") and '
                                                                                               'contains(@class, "Enabled") and '
                                                                                               'contains(@class, "icon_menu") and @alt="Opções"]')))
                driver.execute_script("arguments[0].click();", second_image)
                # second_image.click()
            except TimeoutException:
                pass
            try:
                excel = wait.until(
                    EC.presence_of_element_located((By.XPATH, '//div[text()="Exportar para o Excel"]')))
                excel.click()
            except TimeoutException as ex:
                pass
            try:
                # wait for the input box to be visible and then find it
                input_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                                            'input.f.fN.fInside')))
                input_box.clear()  # clear it
                break
            except TimeoutException as ex:
                pass

        # Check the specific 'ugr' items and adjust the file_name_part accordingly
                # Ajusta nome do arquivo
        if ugr_item == 'ONOA - DIV OPERAÇÃO DE ÁGUA NORTE':
            file_name_part = ugr_item.split('DE', 1)[-1].strip().lower().replace(' ', '_')
        else:
            file_name_part = ugr_item.split('OPE', 1)[-1].strip().lower().replace(' ', '_')

        # Timestamp e nomes
        current_time = datetime.now().strftime('%Y%m%d%H%M')
        file_name = "33627-C-" + file_name_part + "-" + current_time
        naming = "C-" + file_name_part + "-" + current_time
        list_download.append(file_name)

        # Renomeia e trata alertas
        try:
            input_box.clear()
            time.sleep(2)
            alert = wait.until(EC.alert_is_present())
            print("Alert message: " + alert.text)
            alert.accept()
            time.sleep(2)
        except (UnexpectedAlertPresentException, NoAlertPresentException):
            print(f'Processando a carteira de: {file_name}...')

        try:
            input_box.send_keys(f'{naming}')
        except (UnexpectedAlertPresentException, NoAlertPresentException):
            print("Alert presented after typing nome")

        print(f"[DEBUG] UGR atual: {ugr_item.strip()}")

        # Força exportar sempre para "MNT - UG REGIONAL SANTANA"
        pasta_desejada = "MNT - UG REGIONAL SANTANA"
        print(f"[INFO] Pasta fixa definida: {pasta_desejada}")

        try:
            pasta_dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//select[contains(@class, 'fA')]"))
            )
            select = Select(pasta_dropdown)
        except TimeoutException:
            print("⚠️ Dropdown de pasta não encontrado. Tentando reabrir o menu de exportação.")
            try:
                menu_icon = wait.until(EC.presence_of_element_located((By.XPATH, '//img[contains(@class, "icon_menu")]')))
                driver.execute_script("arguments[0].click();", menu_icon)
            except:
                pass
            try:
                export_excel = wait.until(EC.presence_of_element_located((By.XPATH, '//div[text()="Exportar para o Excel"]')))
                export_excel.click()
            except:
                pass
            pasta_dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//select[contains(@class, 'fA')]"))
            )
            select = Select(pasta_dropdown)

        pasta_encontrada = False
        for option in select.options:
            if pasta_desejada.strip() == option.text.strip():
                select.select_by_visible_text(option.text.strip())
                print(f"✅ Selecionado corretamente: {option.text.strip()}")
                pasta_encontrada = True
                break

        if not pasta_encontrada:
            print(f"❌ Pasta '{pasta_desejada}' não encontrada! Pulando este UGR.")
            continue

        # Exporta arquivo
        button_export = wait.until(
            EC.presence_of_element_located((By.XPATH, '//button[contains(@onclick, "OkNameExport#")]')))
        time.sleep(2)
        button_export.click()
        register_export_time(file_name_part)

        # Verifica e fecha alertas
        try:
            WebDriverWait(driver, 2).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            print(f"⚠️ Alerta detectado após exportação: {alert.text}")
            alert.accept()
            print("Alerta fechado com sucesso.")
            continue
        except TimeoutException:
            pass

        try:
            button_esc = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'img.icon.Enabled.icon_x[onclick="closeContextMenu();"]')))
            driver.execute_script("arguments[0].click();", button_esc)
            time.sleep(2)
        except Exception as e:
            print(f"Erro ao clicar no botão X: {e}")

        # Reabre filtros
        button_filtro = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,
                                                                                        "//button[contains(@class, 'pbut') and "
                                                                                        "contains(@class, 'pbutoff') and "
                                                                                        "contains(@class, 'Enabled') and ./div[text()='Filtro']]")))
        button_filtro.click()
        time.sleep(2)

driver.quit()

# Local do arquivo com os nomes
file_path = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\lista_download\lista_carteira.txt"

# Garante que a pasta exista
os.makedirs(os.path.dirname(file_path), exist_ok=True)

# Escreve os nomes dos arquivos exportados
with open(file_path, "w") as file:
    for name in list_download:
        file.write(name + '\n')

# Abre o arquivo e exibe no console
with open(file_path, 'r') as file:
    lines = file.readlines()
    print("\nElaborado relatório para as seguintes áreas:")
    for line in lines:
        print(line.strip())


# Create a DataFrame from the export data
df_export_times = pd.DataFrame(export_data)

# Define the path to save the DataFrame as a CSV file
export_times_csv_path = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_serviços_MBP\time_export\export_times.csv"

# Ensure the directory exists
os.makedirs(os.path.dirname(export_times_csv_path), exist_ok=True)

# Save the DataFrame to a CSV file
df_export_times.to_csv(export_times_csv_path, index=False)

print(f"Export times saved to {export_times_csv_path}")
