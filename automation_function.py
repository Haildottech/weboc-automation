# automation_function.py
import time
from pathlib import Path
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By

from util import Handywrapper
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
import sys
import pandas as pd

# ---------------- GET EXE OR SCRIPT DIRECTORY ----------------
def get_app_directory():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent


# ---------------- DRIVER SETUP ----------------
def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--remote-allow-origins=*")
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option(
        "prefs", {"profile.default_content_setting_values.notifications": 2}
    )

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    hw = Handywrapper(driver)
    return driver, hw


# ---------------- HELPER FUNCTIONS ----------------
def switch_to_new_window(driver, old_handles, timeout=10):
    WebDriverWait(driver, timeout).until(lambda d: len(d.window_handles) > len(old_handles))
    new_handle = (set(driver.window_handles) - set(old_handles)).pop()
    driver.switch_to.window(new_handle)
    return new_handle


def safe_click_js(hw, by, selector, retries=5, delay=1):
    for attempt in range(retries):
        try:
            elem = hw.find_element(by, selector)
            hw.driver.execute_script("arguments[0].scrollIntoView(true);", elem)
            hw.driver.execute_script("arguments[0].click();", elem)
            return True
        except:
            time.sleep(delay)
    return False


# ---------------- EXPORT TO EXCEL ----------------
def export_to_excel(rows):
    if not rows:
        print("No rows to export.")
        return

    app_dir = get_app_directory()
    export_dir = app_dir / "exports"
    export_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = export_dir / f"weboc_export_{timestamp}.xlsx"

    df = pd.DataFrame(rows)

    try:
        df.to_excel(output_file, index=False, sheet_name="Weboc Data")
        print(f"✅ Exported to Excel: {output_file}")
    except PermissionError:
        print(f"❌ Permission denied: Close the file if it is open")


def build_excel_rows(gd_header, item_details, non_duty_paid_rows):
    rows = []
    if not non_duty_paid_rows:
        rows.append({**gd_header, **item_details})
        return rows
    for ndp in non_duty_paid_rows:
        rows.append({**gd_header, **item_details, **ndp})
    return rows


# ---------------- MAIN SCRAPER FUNCTION ----------------
def start_scraping(username, password, start_date, end_date, progress_callback=None):
    driver, hw = setup_driver()
    all_rows = []
    original_window = driver.current_window_handle

    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    end_dt = datetime.strptime(end_date, "%Y-%m-%d")

    try:
        # ---------------- LOGIN WITH RETRY ----------------
        MAX_LOGIN_ATTEMPTS = 5
        attempt = 0
        logged_in = False

        while attempt < MAX_LOGIN_ATTEMPTS and not logged_in:
            attempt += 1
            try:
                driver.get("https://weboc.gov.pk/")
                hw.wait_explicitly(By.ID, "txtLogin")
                hw.find_element(By.ID, "txtLogin").clear()
                hw.find_element(By.ID, "txtLogin").send_keys(username)
                hw.find_element(By.ID, "txtPassword").clear()
                hw.find_element(By.ID, "txtPassword").send_keys(password)
                hw.Click_element(By.ID, "imgbtnLogin")
                hw.wait_explicitly(By.XPATH, "//a[.='Goods Declaration']")
                logged_in = True
                print(f"✅ Login successful on attempt {attempt}")
            except Exception as e:
                print(f"❌ Login attempt {attempt} failed: {e}")
                time.sleep(2)

        if not logged_in:
            print("❌ Failed to login after multiple attempts. Exiting...")
            return

        # ---------------- NAVIGATE TO EXPORT ----------------
        hw.Click_element(By.XPATH, "//a[.='Goods Declaration']")
        hw.wait_explicitly(By.XPATH, "//a[contains(@href,'ExportSubmit')]")
        hw.Click_element(By.XPATH, "//a[contains(@href,'ExportSubmit')]")
        hw.wait_explicitly(By.XPATH, "//td[.='GD No']")

        # ---------------- LOOP OVER DATES ----------------
        delta = end_dt - start_dt
        for i in range(delta.days + 1):
            current_date = start_dt + timedelta(days=i)
            current_date_str = current_date.strftime("%d-%m-%Y")
            if progress_callback:
                progress_callback(f"Scraping date: {current_date_str} ...")
            print(f"Scraping date: {current_date_str}")

            view_elements_xpath = f"//tr[contains(@class,'Grid')]/td[contains(.,'{current_date_str}')]/following-sibling::td/a[.='View']"
            view_elements = hw.find_elements(By.XPATH, view_elements_xpath)

            for idx, _ in enumerate(view_elements, start=1):
                try:
                    clicked = safe_click_js(hw, By.XPATH, f"({view_elements_xpath})[{idx}]")
                    if not clicked:
                        if progress_callback:
                            progress_callback("Skipped a View element due to click failure")
                        continue

                    sub_window_1 = switch_to_new_window(driver, [original_window])
                    hw.wait_explicitly(By.ID, "GdImportViewUc1_txtGDNo")

                    # -------- SCRAPE HEADER --------
                    gd_header = {
                        "GD No": hw.find_element_text(By.ID, "GdImportViewUc1_txtGDNo"),
                        "Destination": hw.find_element_text(By.ID, "GdImportViewUc1_txtDestinationCountry"),
                        "FOB Value": hw.find_element_text(By.ID, "GdImportViewUc1_lblCFRValue"),
                        "Rebate Amount": hw.find_element_text(By.ID, "GdImportViewUc1_txtRebateAmount"),
                        "Export Value": hw.find_element_text(By.ID, "GdImportViewUc1_txtImpExpValue"),
                        "Exchange Rate": hw.find_element_text(By.ID, "GdImportViewUc1_lblExchangeRate"),
                        "Bank Name": hw.find_element_text(By.ID, "GdImportViewUc1_lblBankText"),
                        "No of Packages": hw.find_element_text(By.XPATH, "//table[@id='GdImportViewUc1_dgPackages']//tr[@class='ItemStyle']/td[1]"),
                    }

                    # -------- SCRAPE ITEM DETAILS --------
                    hw.scroll_to_element(By.ID, "ItemsDetailsViewUc1_dgItems")
                    inner_elements = hw.find_elements(By.XPATH, "//table[@id='ItemsDetailsViewUc1_dgItems']//a[.='Details']")

                    for inner_idx in range(len(inner_elements)):
                        clicked = safe_click_js(hw, By.XPATH, f"(//table[@id='ItemsDetailsViewUc1_dgItems']//a[.='Details'])[ {inner_idx+1} ]")
                        if not clicked:
                            continue

                        sub_window_2 = switch_to_new_window(driver, [original_window, sub_window_1])
                        hw.wait_explicitly(By.ID, "GDItemDetailViewUc1_lblTotalValue")

                        item_details = {
                            "HS Code": hw.find_element_text(By.ID, "GDItemDetailViewUc1_lblHSCode"),
                            "Item Total Value": hw.find_element_text(By.ID, "GDItemDetailViewUc1_lblTotalValue"),
                            "Custom Value": hw.find_element_text(By.ID, "GDItemDetailViewUc1_lblImportValue"),
                            "Qty for Assessment": hw.find_element_text(By.ID, "GDItemDetailViewUc1_lblQuantity"),
                        }

                        non_duty_paid_rows = []
                        ndp_rows = hw.find_elements(By.XPATH, "//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr")
                        for r_idx in range(2, len(ndp_rows) + 1):
                            try:
                                HS_code_ndp = hw.find_element_text(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{r_idx}]/td[2]")
                                Qunatity = hw.find_element_text(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{r_idx}]/td[3]")
                                tota_value_ndp = hw.find_element_text(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{r_idx}]/td[5]")
                                Export_value_ndp = hw.find_element_text(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{r_idx}]/td[6]")
                                Import_GD_Machine_Number_ndp = hw.find_element_text(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{r_idx}]/td[7]")
                                non_duty_paid_rows.append({
                                    "NDP HS Code": HS_code_ndp,
                                    "NDP Quantity": Qunatity,
                                    "NDP Total Value": tota_value_ndp,
                                    "NDP Export Value": Export_value_ndp,
                                    "Import GD Machine No": Import_GD_Machine_Number_ndp,
                                })
                            except:
                                continue

                        rows = build_excel_rows(gd_header, item_details, non_duty_paid_rows)
                        all_rows.extend(rows)

                        driver.close()
                        driver.switch_to.window(sub_window_1)

                    driver.close()
                    driver.switch_to.window(original_window)

                except Exception as e:
                    if progress_callback:
                        progress_callback(f"Error in scraping loop: {e}")
                    continue

            if progress_callback:
                progress_callback(f"Scraping date: {current_date_str} ✅ done")

        # ---------------- EXPORT EXCEL ----------------
        if all_rows:
            export_to_excel(all_rows)

        if progress_callback:
            progress_callback("All scraping completed!")

    except Exception as e:
        print(f"❌ Fatal Error: {e}")
        if progress_callback:
            progress_callback(f"❌ Fatal Error: {e}")

    finally:
        driver.quit()
