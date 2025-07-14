import os
import glob
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (
    NoSuchElementException, TimeoutException, JavascriptException,
    StaleElementReferenceException
)
from selenium.webdriver import ActionChains
import shutil

# Setting Chrome options
service = Service(r"chromedriver.exe")
options = webdriver.ChromeOptions()
options.add_argument("disable-infobars")
options.add_argument("start-maximized")
options.add_argument("disable-dev-shm-usage")
options.add_argument("no-sandbox")
# options.add_argument("headless")  # Run in headless mode
options.add_argument("disable-gpu")  # Disable GPU hardware acceleration
options.add_argument("disable-extensions")  # Disable extensions
options.add_argument("disable-software-rasterizer")  # Disable software rasterizer
options.add_argument("enable-automation")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--window-size=1920,1080")  # Set the window size
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("prefs", {
    "profile.managed_default_content_settings.images": 2,  # Disable images
    "disk-cache-size": 4096,  # Limit cache size
    "download.default_directory": r"C:\Users\poliveira.eficien.SBSP\Downloads"
})

# Define directories
download_dir = r"C:\Users\poliveira.eficien.SBSP\Downloads"
destination_dir = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\pendentes"

# Clean up destination folder
def clean_directory(directory):
    if os.path.exists(directory) and os.path.isdir(directory):
        print(f"Cleaning files from: {directory}")
        files = glob.glob(os.path.join(directory, "*"))  # Use os.path.join for better compatibility
        for file_path in files:
            try:
                os.remove(file_path)
                print(f"Deleted: {file_path}")
            except Exception as e:
                print(f"Error deleting {file_path}: {str(e)}")
    else:
        print(f"The specified directory {directory} does not exist or is not a directory")

# Clean up .crdownload files in the download directory
def clean_crdownload_files(directory):
    crdownload_files = glob.glob(os.path.join(directory, "*.crdownload"))
    for file_path in crdownload_files:
        try:
            os.remove(file_path)
            print(f"Deleted: {file_path}")
        except OSError as e:
            print(f"Error deleting {file_path}: {e.strerror}")

# Call cleaning functions
clean_directory(destination_dir)
clean_crdownload_files(download_dir)

# Start the browser
driver = webdriver.Chrome(service=service, options=options)
time.sleep(0)  # Wait for files to generate
print("Waiting for the .xlsx files to be generated.")

driver.get("https://geoprd-interno.sabesp.com.br/sabespwfm/")
wait = WebDriverWait(driver, 10)

# Switch to iframe
wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'iframe')))
driver.switch_to.frame(1)

wait.until(EC.presence_of_element_located((By.ID, 'USER')))

# Login
with open(r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\WFM\cred_wfm.txt", 'r') as file:
    login = file.readline().strip()
    password = file.readline().strip()

driver.find_element(By.ID, 'USER').send_keys(login)
driver.find_element(By.ID, 'INPUTPASS').send_keys(password)

try:
    # Click the submit button to attempt login
    driver.find_element(By.ID, 'submbtn').click()

    # Wait for the error message indicating the user is already connected
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'error-label')))
    error_message = driver.find_element(By.CLASS_NAME, 'error-label').text

    if "O usuário já está conectado" in error_message:
        print("User already connected. Selecting 'Substituir Sessão' option.")

        # Find all radio buttons with the same name 'oooo'
        radio_buttons = driver.find_elements(By.XPATH, "//input[@type='RADIO' and @name='oooo']")

        # Identify the correct one by inspecting its 'onclick' attribute value
        for radio in radio_buttons:
            onclick_value = radio.get_attribute("onclick")
            if "form.CMD.value='RESTART'" in onclick_value:
                radio.click()  # Select the "Substituir Sessão" option
                print("Selected 'Substituir Sessão'.")

                # Click the enter button to proceed
                enter_button = driver.find_element(By.CLASS_NAME, 'enterbutton')
                ActionChains(driver).move_to_element(enter_button).click().perform()
                print("Session replaced successfully.")
                break

except TimeoutException:
    print("Login attempt timed out. Please check your network or page structure.")
except NoSuchElementException as e:
    print(f"Error locating element: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")

# Interact with the page
button_listaOS = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Lista OS')]")))
button_listaOS.click()
time.sleep(5)

button_pesquisaOS = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='TBB_tbm2']/div[1]/div[2]")))
button_pesquisaOS.click()
time.sleep(5)

button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[./div[text()='Relatório']]")))
driver.execute_script("arguments[0].click();", button)
time.sleep(5)

# Define element-clicking function
def click_element():
    try:
        element = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@tooltip='Pasta Superior' and @class='tvMsg tvMsgEna']")))
        element.click()
        return True
    except (NoSuchElementException, TimeoutException):
        print("The element was not found or took too long to appear.")
        return False

if click_element():
    time.sleep(5)
    click_element()
    print("Localizando pasta superior...")


# Double-click cell
def find_and_double_click_cell():
    while True:
        try:
            cell = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//td[@align='left' and @class='tvCell' and contains(text(), 'MNT - UG REGIONAL SANTANA')]")))
            if cell:
                print("Localizando pasta do ONSS...")
                driver.execute_script("arguments[0].scrollIntoView(true);", cell)
                time.sleep(5)
                driver.execute_script("var evt = document.createEvent('MouseEvents');" +
                                      "evt.initMouseEvent('dblclick', true, true, window, 0, 0, 0, 0, 0, false, false, false, 0, null);" +
                                      "arguments[0].dispatchEvent(evt);", cell)
                print("Double-clicked on the desired cell.")

                while True:
                    try:
                        first_cell = wait.until(EC.presence_of_element_located((By.XPATH,
                                                                                "(//td[@align='left' and @class='tvCell'])[1]")))
                        print("Clicando na primeira celula... mes vigente")
                        time.sleep(6)
                        driver.execute_script("var evt = document.createEvent('MouseEvents');" +
                                              "evt.initMouseEvent('dblclick', true, true, window, 0, 0, 0, 0, 0, false, false, false, 0, null);" +
                                              "arguments[0].dispatchEvent(evt);", first_cell)
                        print("Double-clicked on first cell using Javascript.")
                        break
                    except StaleElementReferenceException:
                        print("Retrying to locate and double-click the first cell.")
                        continue
                break
        except(TimeoutException, NoSuchElementException):
            try:
                button = driver.find_element(By.XPATH, "//img[@class='icon Enabled icon_down' and contains(@onclick, 'loadPage')]")
                driver.execute_script("arguments[0].click();", button)
                print("Clicando em REFRESH...")
                time.sleep(3)
            except NoSuchElementException:
                print("Botao REFRESH nao encontrado.")
                break

        except JavascriptException:
            try:
                button = driver.find_element(By.XPATH, "//img[@class='icon Enabled icon_down' and contains(@onclick, 'loadPage')]")
                driver.execute_script("arguments[0].click();", button)
                print("Clicked the button to load more cells after JavascriptException.")
                time.sleep(3)
            except NoSuchElementException:
                print("The button was not found.")
                break

find_and_double_click_cell()

# Read report names from file
file_path = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\lista_download\lista_carteira.txt"
lista_relatorio = []
with open(file_path, 'r') as file:
    for line in file:
        lista_relatorio.append(line.strip())

print("File names in the list:", lista_relatorio)

for relatorio in lista_relatorio:
    while True:
        refresh_button = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[contains(@class, 'tvMsg') and contains(@class, 'tvMsgEna') and @tooltip='Atualizar']")))
        driver.execute_script("arguments[0].click();", refresh_button)
        time.sleep(5)

        try:

            cell = wait.until(EC.presence_of_element_located(
                (By.XPATH, f"//td[@align='left' and @class='tvCell' and contains(text(), '{relatorio}')]")))
            if relatorio[-1] in cell.text:
                ActionChains(driver).double_click(cell).perform()
                time.sleep(30)
                start_time = time.time()

                # Monitor download
                while any(filename.endswith(".crdownload") for filename in os.listdir(download_dir)):
                    if time.time() - start_time > 60:
                        crdownload_file = next(filename for filename in os.listdir(download_dir) if filename.endswith(".crdownload"))
                        os.remove(os.path.join(download_dir, crdownload_file))
                        break
                    time.sleep(1)
                else:
                    downloaded_file = next(filename for filename in os.listdir(download_dir) if
                                           relatorio in filename)
                    src = os.path.join(download_dir, downloaded_file)
                    dst_file_name = relatorio + '.xlsx'
                    dst = os.path.join(destination_dir, dst_file_name)
                    shutil.move(src, dst)
                    break
        except(NoSuchElementException, TimeoutException):
            continue
driver.quit()
print("Download de carteira finalizado!")