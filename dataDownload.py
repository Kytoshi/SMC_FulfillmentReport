from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import time
import os
from datetime import datetime, timedelta
import pandas as pd
import threading

import logging
import shutil


logging.basicConfig(
    filename="error_log.txt",
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

holidays = {
    "01/01",  # New Year's Day
    "01/20",  # MLK Day (example static date for demo; usually 3rd Monday in Jan)
    "02/17",  # Presidents' Day (example static date)
    "05/26",  # Memorial Day (example static date)
    "06/19",  # Juneteenth
    "07/03",  # Independence Day (observed)
    "07/04",  # Independence Day
    "09/01",  # Labor Day (example static date)
    "11/27",  # Thanksgiving (example static date)
    "11/28",  # Day after Thanksgiving
    "12/24",  # Christmas Eve
    "12/25"   # Christmas Day
}

def backup_file(source_folder, destination_folder, file_prefix, previous_date):
    """
    Finds the latest file in the source_folder that starts with file_prefix, 
    renames it with the current date, and copies it to the destination_folder.
    """
    files = [f for f in os.listdir(source_folder) if f.startswith(file_prefix)]
    
    if not files:
        print(f"{file_prefix} Not Found in {source_folder}")
        return
    
    # Get the most recent file based on modification time
    files.sort(key=lambda x: os.path.getmtime(os.path.join(source_folder, x)), reverse=True)
    latest_file = files[0]

    # Generate the new filename with previous date
    file_name, file_extension = os.path.splitext(latest_file)
    new_file_name = f"{file_name}_{previous_date}{file_extension}"

    source_path = os.path.join(source_folder, latest_file)
    destination_path = os.path.join(destination_folder, new_file_name)

    try:
        shutil.copy2(source_path, destination_path)
        print(f"Copied {latest_file} to {destination_folder} as {new_file_name}\n")
    except Exception as e:
        print(f"Error copying file: {e}")

def wait_for_element(driver, by, value, total_wait=480, check_interval=10):
    try:
        # print(f"Waiting for element: {value} for up to {total_wait} seconds...")
        element = WebDriverWait(driver, total_wait, check_interval).until(EC.presence_of_element_located((by, value)))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        return element
    except TimeoutException:
        print(f"Timeout waiting for element: {value}")
        raise TimeoutException(f"Element with {value} not found after {total_wait} seconds.")

def subtract_one_business_day(date, holidays=holidays):
    date -= timedelta(days=1)

    while True:
        date_str = date.strftime('%m/%d')  # Only month/day

        if date_str in holidays:
            print(f"{date.strftime('%m/%d/%Y')} is a holiday, going back one more day.")
            date -= timedelta(days=1)
            continue

        if date.weekday() == 5:  # Saturday
            print(f"{date.strftime('%m/%d/%Y')} is Saturday, going back one more day.")
            date -= timedelta(days=1)
            continue
        elif date.weekday() == 6:  # Sunday
            print(f"{date.strftime('%m/%d/%Y')} is Sunday, going back two days.")
            date -= timedelta(days=2)
            continue

        break  # Found valid business day

    return date

def create_driver(download_path):

    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0
    })
    chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("enable-automation")
    chrome_options.add_argument("disable-infobars")
    chrome_options.page_load_strategy = "eager"
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--max-old-space-size=4096")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("window-size=1920,1080")

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
    driver.set_page_load_timeout(600)
    driver.set_script_timeout(600)
    driver.implicitly_wait(30)
    driver.execute_cdp_cmd("Network.enable", {})
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": download_path
    })

    return driver


def DailyOS(username, password, dPath, progress_callback=None):
    """
    Original DailyOS functionality with determinate progress updates.
    """
    try:
        steps = [
            "Prepare backup folder",
            "Backup previous report",
            "Launch browser",
            "Login",
            "Navigate to report page",
            "Set report date",
            "Click report link",
            "Wait for download",
            "Wait for stable file",
            "Convert to Excel"
        ]
        total_steps = len(steps)
        def report_progress(step_index):
            if progress_callback:
                percent = ((step_index + 1) / total_steps) * 100
                progress_callback("update", percent)

        # --- Step 0: Backup folder ---
        if progress_callback:
            progress_callback("start")
        backup_path = os.path.join(dPath, "backup")
        if not os.path.exists(backup_path):
            os.makedirs(backup_path)
        report_progress(0)

        # --- Step 1: Backup previous report ---
        if os.path.exists(os.path.join(dPath, "DailyReport.xlsx")):
            backup_suffix = subtract_one_business_day(subtract_one_business_day(datetime.today())).strftime("%m-%d-%Y")
            backup_file(dPath, backup_path, "DailyReport", backup_suffix)
        report_progress(1)

        # --- Step 2: Launch browser ---
        driver = create_driver(str(dPath))
        report_progress(2)

        # --- Step 3: Login ---
        driver.get("https://pdbs.supermicro.com:18893/Home")
        username_field = driver.find_element(By.ID, "txtUserName")
        password_field = driver.find_element(By.ID, "xPWD")
        username_field.send_keys(username)
        password_field.send_keys(password)
        driver.find_element(By.ID, "btnSubmit").click()
        time.sleep(1)
        report_progress(3)

        # --- Step 4: Navigate to report page ---
        links = driver.find_elements(By.TAG_NAME, 'a')
        for link in links:
            if link.get_attribute('href') == 'javascript:onClickTaskMenu("OrdReport.asp", 65)':
                link.click()
                break
        report_progress(4)

        # --- Step 5: Set report date ---
        DailyOrders_date_field = wait_for_element(driver, By.NAME, "Date")
        DailyOrders_date_field.clear()
        prevDate = subtract_one_business_day(datetime.today())
        DailyOrders_date_field.send_keys(prevDate.strftime("%m/%d/%Y"))
        driver.execute_script("ChgDate()")
        report_progress(5)

        # --- Step 6: Click report link ---
        link = driver.find_element(By.LINK_TEXT, "Order Fulfillment Report")
        link.click()
        report_progress(6)

        # --- Step 7: Wait for download to appear ---
        timeout = 300
        start_time = time.time()
        while time.time() - start_time < timeout:
            files = [f for f in os.listdir(dPath) if f.startswith("DailyReport") and f.endswith(".xls")]
            if files:
                file_path = max([os.path.join(dPath, f) for f in files], key=os.path.getmtime)
                if not file_path.endswith(('.crdownload', '.part')):
                    break
            time.sleep(1)
        else:
            raise TimeoutError("File download timed out.")
        report_progress(7)

        # --- Step 8: Wait for stable file size ---
        def wait_for_stable_file_size(path, wait_time=3, check_interval=1):
            last_size = -1
            stable_count = 0
            while stable_count < wait_time:
                try:
                    current_size = os.path.getsize(path)
                except OSError:
                    current_size = -1
                if current_size == last_size and current_size > 0:
                    stable_count += check_interval
                else:
                    stable_count = 0
                last_size = current_size
                time.sleep(check_interval)
        wait_for_stable_file_size(file_path)
        report_progress(8)

        # --- Step 9: Convert to Excel ---
        if file_path.endswith('.xls'):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    first_line = f.readline()
                    second_line = f.readline()
                if ('<html' in first_line.lower() or '<table' in first_line.lower() or
                    '<html' in second_line.lower() or '<table' in second_line.lower()):
                    dfs = pd.read_html(file_path)
                    df = dfs[0]
                else:
                    df = pd.read_excel(file_path)
                xlsx_path = os.path.join(dPath, "DailyReport.xlsx")
                df.to_excel(xlsx_path, index=False)
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception as e:
                print(f"Error converting {file_path}: {e}")
        elif file_path.endswith('.xlsx'):
            try:
                df = pd.read_excel(file_path)
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
        else:
            print(f"Downloaded file is not .xls or .xlsx: {file_path}")

        driver.quit()
        report_progress(9)

    except Exception as e:
        logging.error("DailyOS error", exc_info=True)
    finally:
        if progress_callback:
            progress_callback("stop", 100)