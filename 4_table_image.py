import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image

# Define the folder containing the .txt files
txt_folder = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\styled_htmls"

# Define the folder to save the images
img_folder = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\html_image"
os.makedirs(img_folder, exist_ok=True)

# Path to the external CSS file
css_file_path = r"C:\Users\poliveira.eficien.SBSP\Desktop\cod_carteira_servicos\carteira_servicos_MBP\style_css\styles.css"

# Function to convert HTML to image using Selenium
def html_to_image(html_content, image_path, width=1600, height=1200, full_screen=False):
    # Inject the CSS file into the HTML
    styled_html_content = f"""
    <html>
    <head>
    <link rel="stylesheet" type="text/css" href="file:///{css_file_path}">
    </head>
    <body>
    <div style="width: {width}px; margin: auto;">
    {html_content}
    </div>
    </body>
    </html>
    """

    # Set up Selenium with Chrome
    chrome_options = Options()
    service = Service(r"chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        # Create a temporary HTML file with CSS linked
        temp_html_path = os.path.join(img_folder, "temp.html")
        with open(temp_html_path, "w", encoding="utf-8") as f:
            f.write(styled_html_content)

        # Load the HTML file in the browser
        driver.get(f"file:///{temp_html_path}")

        # If full_screen is enabled, maximize the browser and **DO NOT resize afterward**
        if full_screen:
            driver.maximize_window()

        # Wait for the table to be fully rendered
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.TAG_NAME, "table")))

        # Extract the table's size BEFORE trying to scroll or take a screenshot
        table_element = driver.find_element(By.TAG_NAME, "table")
        table_rect = table_element.rect  # Ensure table_rect is correctly assigned

        # Scroll to the top first (ensures full-page loading)
        driver.execute_script("window.scrollTo(0, 0);")

        # Scroll to bring the table into full view
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        # Give extra time for full rendering
        time.sleep(2)

        # Take a screenshot and save it
        driver.save_screenshot(image_path)

    finally:
        driver.quit()
        os.remove(temp_html_path)  # Cleanup temp file

    # Crop the image based on detected table size (now table_rect is correctly assigned)
    crop_image(image_path, table_rect)

# Function to crop the image based on table boundaries
def crop_image(image_path, table_rect):
    with Image.open(image_path) as img:
        img = img.convert("RGB")  # Ensure RGB mode

        # Use JavaScript to dynamically get header height
        padding_top = max(50, table_rect["height"] * 0.1)  # Adjust based on table size
        padding_bottom = max(50, table_rect["height"] * 0.05)  # Ensure footer visibility

        # Define cropping boundaries
        left = table_rect["x"]
        top = max(0, table_rect["y"] - padding_top)  # Adjust to include headers
        right = left + table_rect["width"]
        bottom = min(img.height, top + table_rect["height"] + padding_bottom)  # Avoid cutting footer

        # Apply cropping and save
        img.crop((left, top, right, bottom)).save(image_path)

# Iterate over each .txt file in the folder and apply styling
for filename in os.listdir(txt_folder):
    if filename.endswith('.txt'):
        txt_path = os.path.join(txt_folder, filename)
        img_path = os.path.join(img_folder, filename.replace('.txt', '_servicos_pendentes.png'))

        # Read the HTML content from the .txt file
        with open(txt_path, "r", encoding="utf-8") as file:
            html_content = file.read()

        # Convert HTML to image with specified width and height
        html_to_image(html_content, img_path, width=1200, height=650, full_screen=True)

        print(f"Converted {txt_path} to {img_path}")
