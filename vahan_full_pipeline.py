import os
import time
import shutil
import pandas as pd
import argparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options

# --- Scraper Setup ---
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def select_primefaces_dropdown(browser, wait, container_id, option_text):
    container = wait.until(EC.element_to_be_clickable((By.ID, container_id)))
    container.click()
    panel_id = container_id + '_panel'
    panel = wait.until(EC.visibility_of_element_located((By.ID, panel_id)))
    options = panel.find_elements(By.TAG_NAME, 'li')
    for opt in options:
        if opt.text.strip() == option_text:
            opt.click()
            time.sleep(0.5)
            return
    raise Exception(f'Option "{option_text}" not found in {container_id}')

def wait_for_yaxis_change(browser, expected):
    for _ in range(20):
        label = browser.find_element(By.ID, "yaxisVar_label").text.strip()
        if label == expected:
            return True
        time.sleep(0.3)
    raise Exception(f"Y-Axis did not change to {expected}")

def wait_for_download(filename, timeout=30):
    for _ in range(timeout * 2):
        files = [f for f in os.listdir(DOWNLOAD_DIR) if f.startswith("reportTable") and f.endswith(".xlsx")]
        if files:
            return sorted(files, key=lambda x: os.path.getmtime(os.path.join(DOWNLOAD_DIR, x)))[-1]
        time.sleep(0.5)
    raise Exception("Download did not complete in time")

def wait_for_overlay_to_disappear(browser):
    for _ in range(20):
        overlay = browser.find_element(By.ID, "j_idt135_blocker")
        if overlay.value_of_css_property("display") == "none":
            return
        time.sleep(0.3)

# --- Scraping ---
def run_scraper(selected_types=None, start_year=2018, end_year=2026):
    edge_options = Options()
    edge_options.use_chromium = True
    edge_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    def start_browser():
        edge_options = Options()
        edge_options.use_chromium = True
        edge_options.add_experimental_option("prefs", {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        browser = webdriver.Edge(options=edge_options)
        url = "https://vahan.parivahan.gov.in/vahan4dashboard/vahan/view/reportview.xhtml"
        browser.get(url)
        wait = WebDriverWait(browser, 30)
        time.sleep(5)
        select_primefaces_dropdown(browser, wait, "xaxisVar", "Month Wise")
        year_type_container = wait.until(EC.presence_of_element_located((By.ID, "selectedYearType")))
        if "ui-state-disabled" not in year_type_container.get_attribute("class"):
            select_primefaces_dropdown(browser, wait, "selectedYearType", "Calendar Year")
        return browser, wait

    yaxis_options = [
        "Vehicle Category",
        "Vehicle Class",
        "Norms",
        "Fuel",
        "Maker",
        "State"
    ]
    if selected_types:
        yaxis_options = [y for y in yaxis_options if y in selected_types]

    browser, wait = start_browser()
    for idx, yaxis in enumerate(yaxis_options):
        # Restart browser after every 2 Y-Axis
        if idx > 0 and idx % 2 == 0:
            browser.quit()
            browser, wait = start_browser()
        for attempt_yaxis in range(3):
            try:
                select_primefaces_dropdown(browser, wait, "yaxisVar", yaxis)
                wait_for_yaxis_change(browser, yaxis)
                break
            except Exception as e:
                print(f"Retrying Y-Axis {yaxis} (attempt {attempt_yaxis+1}): {e}")
                time.sleep(2)
        yaxis_folder = os.path.join(DOWNLOAD_DIR, yaxis.replace(" ", "_"))
        os.makedirs(yaxis_folder, exist_ok=True)
        for year in range(start_year, end_year + 1):
            success = False
            for attempt in range(3):
                try:
                    # Always re-select Y-Axis and Year to get fresh elements
                    select_primefaces_dropdown(browser, wait, "yaxisVar", yaxis)
                    wait_for_yaxis_change(browser, yaxis)
                    select_primefaces_dropdown(browser, wait, "selectedYear", str(year))
                    refresh_btn = wait.until(EC.element_to_be_clickable((By.ID, "j_idt73")))
                    refresh_btn.click()
                    time.sleep(2)
                    export_btn = wait.until(EC.element_to_be_clickable((By.ID, "groupingTable:xls")))
                    wait_for_overlay_to_disappear(browser)
                    export_btn.click()
                    downloaded = wait_for_download("reportTable.xlsx")
                    new_name = f"{yaxis.replace(' ', '_')}_{year}.xlsx"
                    shutil.move(os.path.join(DOWNLOAD_DIR, downloaded), os.path.join(yaxis_folder, new_name))
                    print(f"Saved: {os.path.join(yaxis_folder, new_name)}")
                    success = True
                    break
                except Exception as e:
                    print(f"Retry {attempt+1}/3 for {yaxis} {year}: {e}")
                    if attempt < 2:
                        time.sleep(2)
            if not success:
                print(f"All attempts failed for {yaxis} {year}. Reloading page and retrying once.")
                browser.refresh()
                time.sleep(5)
                try:
                    select_primefaces_dropdown(browser, wait, "xaxisVar", "Month Wise")
                    year_type_container = wait.until(EC.presence_of_element_located((By.ID, "selectedYearType")))
                    if "ui-state-disabled" not in year_type_container.get_attribute("class"):
                        select_primefaces_dropdown(browser, wait, "selectedYearType", "Calendar Year")
                    select_primefaces_dropdown(browser, wait, "yaxisVar", yaxis)
                    wait_for_yaxis_change(browser, yaxis)
                    select_primefaces_dropdown(browser, wait, "selectedYear", str(year))
                    refresh_btn = wait.until(EC.element_to_be_clickable((By.ID, "j_idt73")))
                    refresh_btn.click()
                    time.sleep(2)
                    export_btn = wait.until(EC.element_to_be_clickable((By.ID, "groupingTable:xls")))
                    wait_for_overlay_to_disappear(browser)
                    export_btn.click()
                    downloaded = wait_for_download("reportTable.xlsx")
                    new_name = f"{yaxis.replace(' ', '_')}_{year}.xlsx"
                    shutil.move(os.path.join(DOWNLOAD_DIR, downloaded), os.path.join(yaxis_folder, new_name))
                    print(f"Saved after reload: {os.path.join(yaxis_folder, new_name)}")
                except Exception as e:
                    print(f"Final retry after reload failed for {yaxis} {year}: {e}")
    browser.quit()

# --- Master Sheet Compilation ---
def compile_master_sheet(clean_master=False):
    criteria_folders = [f for f in os.listdir(DOWNLOAD_DIR) if os.path.isdir(os.path.join(DOWNLOAD_DIR, f))]
    writer = pd.ExcelWriter(os.path.join(DOWNLOAD_DIR, 'master_sheet.xlsx'), engine='openpyxl')
    for folder in criteria_folders:
        folder_path = os.path.join(DOWNLOAD_DIR, folder)
        year_files = sorted([f for f in os.listdir(folder_path) if f.endswith('.xlsx')], key=lambda x: int(x.split('_')[-1].replace('.xlsx','')))
        header_rows = 5
        merged_data = []
        for file in year_files:
            file_path = os.path.join(folder_path, file)
            try:
                df = pd.read_excel(file_path, header=None)
                # Ensure columns are string for .str operations
                df.columns = df.columns.map(str)
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                # If first file, initialize merged_data
                if not merged_data:
                    merged_data = df.values.tolist()
                else:
                    # Find last merged column in row 0 (header row)
                    last_col = len(merged_data[0])
                    # Append new year data just after last_col
                    for i in range(len(df)):
                        # If merged_data has fewer rows, extend
                        if i >= len(merged_data):
                            merged_data.append([None]*last_col)
                        merged_data[i].extend(df.iloc[i].tolist())
            except Exception as e:
                print(f"Failed to read {file_path}: {e}")
        if merged_data:
            sheet_name = folder.replace('_', ' ')[:31]
            merged_df = pd.DataFrame(merged_data)
            if clean_master:
                # Remove first 3 rows
                merged_df = merged_df.iloc[3:].reset_index(drop=True)
                # Rename columns based on row 0 (now the header row)
                import re
                def clean_month_header(val):
                    if isinstance(val, str):
                        m = re.match(r"([A-Za-z]+)[\s\-/]*(\d{4})", val)
                        if m:
                            month = m.group(1)[:3].lower()
                            year = m.group(2)
                            return f"{month}-{year}"
                    return val
                merged_df.columns = [clean_month_header(x) for x in merged_df.iloc[0]]
                merged_df = merged_df.iloc[1:].reset_index(drop=True)
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.close()
    print('Master sheet created at:', os.path.join(DOWNLOAD_DIR, 'master_sheet.xlsx'))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Vahan Dashboard Scraper")
    parser.add_argument('--type', nargs='*', help='Y-Axis types to scrape (e.g. "Maker" "Fuel")')
    parser.add_argument('--start-year', type=int, default=2018, help='Start year (inclusive)')
    parser.add_argument('--end-year', type=int, default=2026, help='End year (inclusive)')
    parser.add_argument('--scrape', action='store_true', help='Run scraping only (no compilation)')
    parser.add_argument('--compile', action='store_true', help='Compile master sheet only (no scraping)')
    parser.add_argument('--clean-master', action='store_true', help='Clean master sheet: remove first 3 rows and rename columns to month-year format')
    parser.add_argument('--all', action='store_true', help='Run full pipeline: scrape, compile, and clean master sheet')
    args = parser.parse_args()

    # Default: run full pipeline if no specific action is given
    if args.all or (not args.scrape and not args.compile and not args.clean_master):
        run_scraper(selected_types=args.type, start_year=args.start_year, end_year=args.end_year)
        compile_master_sheet(clean_master=True)
    elif args.scrape:
        run_scraper(selected_types=args.type, start_year=args.start_year, end_year=args.end_year)
    elif args.compile:
        compile_master_sheet(clean_master=args.clean_master)
    elif args.clean_master:
        # Only clean the master sheet (assumes master_sheet.xlsx already exists)
        compile_master_sheet(clean_master=True)
