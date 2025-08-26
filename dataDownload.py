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


def DailyOS(username, password, dPath, progress_callback=None) -> None:

    try:
        # Define backup folder path
        backup_path = os.path.join(dPath, "backup")

        # Check if backup folder exists, create if missing
        if not os.path.exists(backup_path):
            print(f"Backup folder not found. Creating at {backup_path}...")
            os.makedirs(backup_path)
        else:
            print(f"Backup folder already exists at {backup_path}")
        
        if not os.path.exists(dPath + "/DailyReport.xlsx"):
            print(f"DailyReport.xlsx does not exist. Moving forward with program...")
        else:
            # Format backup suffix (yesterday’s business day)
            backup_suffix = subtract_one_business_day(subtract_one_business_day(datetime.today())).strftime("%m-%d-%Y")

            # Run backup
            backup_file(dPath, backup_path, "DailyReport", backup_suffix)
            print(f"Backup complete: DailyReport_{backup_suffix}")
    except Exception as e:
        print(f"Backup failed: {e}")

    try:
        if progress_callback:
            progress_callback("start")  # Start progress animation

        driver = create_driver(str(dPath))
        driver.set_page_load_timeout(600)
        driver.set_script_timeout(600)
        driver.get("https://pdbs.supermicro.com:18893/Home")

        # Login steps
        print("Logging in...")
        username_field = driver.find_element(By.ID, "txtUserName")
        password_field = driver.find_element(By.ID, "xPWD")
        username_field.send_keys(username)
        password_field.send_keys(password)
        driver.find_element(By.ID, "btnSubmit").click()

        # Wait for the login to complete
        time.sleep(1)

        # Navigate to Daily Order Status Report Page
        links = driver.find_elements(By.TAG_NAME, 'a')

        for link in links:
            if link.get_attribute('href') == 'javascript:onClickTaskMenu("OrdReport.asp", 65)':
                link.click()
                break

        #Set date for third page (Daily Orders)
        DailyOrders_date_field = wait_for_element(driver, By.NAME, "Date")
        DailyOrders_date_field.clear()
        today = datetime.today()
        prevDate = subtract_one_business_day(today)
        DailyOrders_date_field.send_keys(prevDate.strftime("%m/%d/%Y")) # Sets date to the previous business day

        driver.execute_script("ChgDate()")

        try:
            # Find the link by its visible text and click it
            link = driver.find_element(By.LINK_TEXT, "Order Fulfillment Report")
            link.click()
            # print("Link clicked successfully!")

        except Exception as e:
            print(f"Error: {e}")

        # Wait for the file to appear and be fully downloaded
        timeout = 300  # Set a timeout in seconds (adjust as needed)
        start_time = time.time()

        while time.time() - start_time < timeout:
            files = [f for f in os.listdir(dPath) if f.startswith("DailyReport") and f.endswith(".xls")]
            if files:
                # Find the most recently modified file that matches
                file_path = max([os.path.join(dPath, f) for f in files], key=os.path.getmtime)

                if file_path.endswith(".crdownload") or file_path.endswith(".part"):
                    time.sleep(1)
                else:
                    # Only rename if another completed file with the same name exists (not the file just downloaded)
                    base_name = os.path.basename(file_path)
                    completed_files = [f for f in os.listdir(dPath)
                                    if f == base_name and f != base_name and not (f.endswith('.crdownload') or f.endswith('.part'))]
                    if completed_files:
                        unique_name = get_unique_filename(dPath, base_name)
                        unique_path = os.path.join(dPath, unique_name)
                        if file_path != unique_path:
                            os.rename(file_path, unique_path)
                            file_path = unique_path
                    # print("Download complete:", file_path)
                    break
            time.sleep(1)
        else:
            raise TimeoutError("File download timed out.")
        
        # Wait for file size to stabilize before processing
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

        # After download is complete and file_path is set
        wait_for_stable_file_size(file_path)

        # Wait for all temporary files to disappear before checking for duplicates and renaming
        def wait_for_no_temp_files(directory, check_interval=1):
            while any(f.endswith('.crdownload') or f.endswith('.part') for f in os.listdir(directory)):
                time.sleep(check_interval)

        # After download is complete and file_path is set
        wait_for_no_temp_files(dPath)

        # Now run unique filename logic and conversion
        # Convert .xls to .xlsx if needed, handling HTML disguised as .xls
        if file_path.endswith('.xls'):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    first_line = f.readline()
                    second_line = f.readline()
                if ('<html' in first_line.lower() or '<table' in first_line.lower() or
                    '<html' in second_line.lower() or '<table' in second_line.lower()):
                    # File is HTML, use read_html
                    dfs = pd.read_html(file_path)
                    df = dfs[0]  # Use the first table found
                else:
                    # File is a real Excel file
                    df = pd.read_excel(file_path)
                # Use get_unique_filename for .xlsx output
                xlsx_path = os.path.join(dPath, "DailyReport.xlsx")
                df.to_excel(xlsx_path, index=False)
                # print(f"Converted file to {xlsx_path}")
                # Delete the original file after conversion
                if os.path.exists(file_path):
                    os.remove(file_path)
                    # print(f"Deleted original file: {file_path}")
            except Exception as e:
                print(f"Error converting {file_path}: {e}")
        elif file_path.endswith('.xlsx'):
            try:
                df = pd.read_excel(file_path)
                # print(f"Read Excel file: {file_path}")
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
        else:
            print(f"Downloaded file is not .xls or .xlsx: {file_path}")
        
        driver.quit()

    finally:
        if progress_callback:
            progress_callback("stop")  # Stop progress animation