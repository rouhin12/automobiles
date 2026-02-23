import os
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import Select

# Set up download directory
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


# Configure Edge options
edge_options = Options()
edge_options.use_chromium = True
edge_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Initialize Edge WebDriver
browser = webdriver.Edge(options=edge_options)

url = "https://vahan.parivahan.gov.in/vahan4dashboard/vahan/view/reportview.xhtml"
browser.get(url)

wait = WebDriverWait(browser, 30)

# Wait for the page to load (wait for Refresh button to appear)
wait.until(EC.presence_of_element_located((By.ID, "j_idt73")))




# Utility to select PrimeFaces dropdown by visible text
def select_primefaces_dropdown(container_id, option_text):
    # Click the dropdown container to open the menu
    container = wait.until(EC.element_to_be_clickable((By.ID, container_id)))
    container.click()
    # Wait for the dropdown panel to appear
    panel_id = container_id + '_panel'
    panel = wait.until(EC.visibility_of_element_located((By.ID, panel_id)))
    # Find the option by text and click it
    options = panel.find_elements(By.TAG_NAME, 'li')
    for opt in options:
        if opt.text.strip() == option_text:
            opt.click()
            time.sleep(0.5)
            return
    raise Exception(f'Option "{option_text}" not found in {container_id}')


# Wait for dropdowns to populate (dynamic loading)
time.sleep(5)

# Print all select element IDs and their options for debugging (after wait)
select_elements = browser.find_elements(By.TAG_NAME, 'select')
print('Available <select> elements after wait:')
for sel in select_elements:
    print('ID:', sel.get_attribute('id'), '| Name:', sel.get_attribute('name'))
    options = sel.find_elements(By.TAG_NAME, 'option')
    print('  Options:', [opt.text for opt in options])





# Only set X-Axis and Year Type if needed
select_primefaces_dropdown("xaxisVar", "Month Wise")
# Only select Year Type if not disabled
year_type_container = wait.until(EC.presence_of_element_located((By.ID, "selectedYearType")))
if "ui-state-disabled" not in year_type_container.get_attribute("class"):
    select_primefaces_dropdown("selectedYearType", "Calendar Year")

# Y-Axis options to loop through (exact visible text)
yaxis_options = [
    "Vehicle Category",
    "Vehicle Class",
    "Norms",
    "Fuel",
    "Maker",
    "State"
]

# Loop through Y-Axis options and years
import shutil
def wait_for_yaxis_change(expected):
    # Wait until the label reflects the expected Y-Axis
    for _ in range(20):
        label = browser.find_element(By.ID, "yaxisVar_label").text.strip()
        if label == expected:
            return True
        time.sleep(0.3)
    raise Exception(f"Y-Axis did not change to {expected}")

def wait_for_download(filename, timeout=30):
    # Wait for a file to appear in the download directory
    for _ in range(timeout * 2):
        files = [f for f in os.listdir(DOWNLOAD_DIR) if f.startswith("reportTable") and f.endswith(".xlsx")]
        if files:
            return sorted(files, key=lambda x: os.path.getmtime(os.path.join(DOWNLOAD_DIR, x)))[-1]
        time.sleep(0.5)
    raise Exception("Download did not complete in time")

for yaxis in yaxis_options:
    select_primefaces_dropdown("yaxisVar", yaxis)
    wait_for_yaxis_change(yaxis)
    yaxis_folder = os.path.join(DOWNLOAD_DIR, yaxis.replace(" ", "_"))
    os.makedirs(yaxis_folder, exist_ok=True)
    for year in range(2018, 2027):
        select_primefaces_dropdown("selectedYear", str(year))
        # Click Refresh button
        try:
            refresh_btn = wait.until(EC.element_to_be_clickable((By.ID, "j_idt73")))
            refresh_btn.click()
            time.sleep(2)
        except Exception as e:
            print(f"Failed to refresh for {yaxis} {year}: {e}")
            continue
        # Wait for export button (Excel) to appear after refresh
        try:
            export_btn = wait.until(EC.element_to_be_clickable((By.ID, "groupingTable:xls")))
            # Wait for overlay to disappear
            for _ in range(20):
                overlay = browser.find_element(By.ID, "j_idt135_blocker")
                if overlay.value_of_css_property("display") == "none":
                    break
                time.sleep(0.3)
            export_btn.click()
            # Wait for download and move/rename
            downloaded = wait_for_download("reportTable.xlsx")
            new_name = f"{yaxis.replace(' ', '_')}_{year}.xlsx"
            shutil.move(os.path.join(DOWNLOAD_DIR, downloaded), os.path.join(yaxis_folder, new_name))
            print(f"Saved: {os.path.join(yaxis_folder, new_name)}")
        except Exception as e:
            print(f"Failed to export for {yaxis} {year}: {e}")

browser.quit()
print("Scraping complete. Check the downloads folder.")
