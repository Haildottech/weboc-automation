from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
from util import Handywrapper
import pandas as pd
from pathlib import Path

USERNAME = "0676689"
PASSWORD = "123456"

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--user-data-dir=D:/weboc-automation/chrome_profile")

driver = webdriver.Chrome(options=chrome_options)

wait = WebDriverWait(driver, 20)

filename = Path("D:/weboc-automation/exports/test.xlsx")
filename.parent.mkdir(parents=True, exist_ok=True)


def build_excel_rows(
    gd_header,
    item_details,
    non_duty_paid_rows
):
    """
    Returns list of flat rows for Excel
    """
    rows = []

    # If no non-duty-paid rows exist
    if not non_duty_paid_rows:
        row = {**gd_header, **item_details}
        rows.append(row)
        return rows

    for ndp in non_duty_paid_rows:
        row = {
            **gd_header,
            **item_details,
            **ndp
        }
        rows.append(row)

    return rows

def export_to_excel(rows, filename="weboc_export.xlsx"):
    if not rows:
        print("⚠️ No data to export")
        return
    
    cleaned_rows = []
    for row in rows:
        clean_row = {}
        for k, v in row.items():
            if v is None:
                clean_row[k] = ""
            elif isinstance(v, (str, int, float)):
                clean_row[k] = v
            else:
                # convert any object (lists, dicts) to string
                clean_row[k] = str(v)
        cleaned_rows.append(clean_row)

    df = pd.DataFrame.from_records(rows)

    output = Path(filename)
    output.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False, sheet_name="Weboc Data")

    print(f"✅ Data exported to {output.resolve()}")

def switch_to_new_window(driver, old_handles, timeout=10):
    WebDriverWait(driver, timeout).until(
        lambda d: len(d.window_handles) > len(old_handles)
    )

    new_handle = (set(driver.window_handles) - set(old_handles)).pop()
    driver.switch_to.window(new_handle)
    return new_handle

try:
    driver.get("https://weboc.gov.pk/")
    hw = Handywrapper(driver)   

    hw.wait_explicitly(By.ID, "txtLogin")
    hw.find_element(By.ID, "txtLogin").send_keys(USERNAME)
    hw.find_element(By.ID, "txtPassword").send_keys(PASSWORD)
    hw.Click_element(By.ID, "imgbtnLogin")

    hw.wait_explicitly(By.XPATH, "//a[.='Goods Declaration']")  
    hw.Click_element(By.XPATH, "//a[.='Goods Declaration']")

    hw.wait_explicitly(By.XPATH, "//a[contains(@href,'ExportSubmit')]")
    hw.Click_element(By.XPATH, "//a[contains(@href,'ExportSubmit')]")

    hw.wait_explicitly(By.XPATH, "//td[.='GD No']")
    
    original_window = driver.current_window_handle
    all_rows = []
    view_elements_xpath = "//tr[contains(@class,'Grid')]/td[contains(.,'09-01-2026')]/following-sibling::td/a[.='View']"
    count = len(hw.find_elements(By.XPATH, view_elements_xpath)[:2])

    for i in range(count):
        old_handles = driver.window_handles
        view_elements = hw.find_elements(By.XPATH, view_elements_xpath)
        hw.Click_element(element=view_elements[i])
    
        sub_window_1 = switch_to_new_window(driver, old_handles)

        # Ensure page is loaded
        hw.wait_explicitly(By.ID, "GdImportViewUc1_txtGDNo")
        GD_num = hw.find_element_text(By.ID, "GdImportViewUc1_txtGDNo")
        Destination = hw.find_element_text(By.ID, "GdImportViewUc1_txtDestinationCountry")
        FOB_value = hw.find_element_text(By.ID, "GdImportViewUc1_lblCFRValue")
        Rebate_amount = hw.find_element_text(By.ID, "GdImportViewUc1_txtRebateAmount")
        Export_value = hw.find_element_text(By.ID, "GdImportViewUc1_txtImpExpValue")
        Exchange_rate = hw.find_element_text(By.ID, "GdImportViewUc1_lblExchangeRate")
        Bank_name = hw.find_element_text(By.ID, "GdImportViewUc1_lblBankText")
        No_of_packages = hw.find_element_text(By.XPATH, "//table[@id='GdImportViewUc1_dgPackages']//tr[@class='ItemStyle']/td[1]")
        
        gd_header = {
            "GD No": GD_num,
            "Destination": Destination,
            "FOB Value": FOB_value,
            "Rebate Amount": Rebate_amount,
            "Export Value": Export_value,
            "Exchange Rate": Exchange_rate,
            "Bank Name": Bank_name,
            "No of Packages": No_of_packages,
        }

        hw.scroll_to_element(By.ID, "ItemsDetailsViewUc1_dgItems")

        inner_view_elements = hw.find_elements(By.XPATH, "//table[@id='ItemsDetailsViewUc1_dgItems']//a[.='Details']")
        for inner_view in inner_view_elements:
            old_handles = driver.window_handles
            hw.Click_element(element=inner_view)
            sub_window_2 = switch_to_new_window(driver, old_handles)

            hw.wait_explicitly(By.ID, "GDItemDetailViewUc1_lblTotalValue")
            Total_value = hw.find_element_text(By.ID, "GDItemDetailViewUc1_lblTotalValue")
            Custom_value = hw.find_element_text(By.ID, "GDItemDetailViewUc1_lblImportValue")
            Qty_assessment_purpose = hw.find_element_text(By.ID, "GDItemDetailViewUc1_lblQuantity")
            HS_code = hw.find_element_text(By.ID, "GDItemDetailViewUc1_lblHSCode")
            Total_value = hw.find_element_text(By.ID, "GDItemDetailViewUc1_lblTotalValue")
            item_details = {
                "HS Code": HS_code,
                "Item Total Value": Total_value,
                "Custom Value": Custom_value,
                "Qty for Assessment": Qty_assessment_purpose,
            }
            non_duty_paid_rows = []
            Non_duty_paid_info = hw.find_elements(By.XPATH, "//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr")
            for ind in range(2, len(Non_duty_paid_info[1:])+2):
                hw.wait_explicitly(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{ind}]/td[2]")
                HS_code_ndp = hw.find_element_text(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{ind}]/td[2]")
                Qunatity = hw.find_element_text(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{ind}]/td[3]")
                tota_value_ndp = hw.find_element_text(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{ind}]/td[5]")
                Export_value_ndp = hw.find_element_text(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{ind}]/td[6]")   
                Import_GD_Machine_Number_ndp = hw.find_element_text(By.XPATH, f"//table[@id='GDItemDetailViewUc1_dgNonDutyPaidItems']/tbody/tr[{ind}]/td[7]")  
                non_duty_paid_rows.append({
                    "NDP HS Code": HS_code_ndp,
                    "NDP Quantity": Qunatity,
                    "NDP Total Value": tota_value_ndp,
                    "NDP Export Value": Export_value_ndp,
                    "Import GD Machine No": Import_GD_Machine_Number_ndp,
                })
            
            rows = build_excel_rows(
                gd_header=gd_header,
                item_details=item_details,
                non_duty_paid_rows=non_duty_paid_rows
            )

            all_rows.extend(rows)
            driver.close()                    
            driver.switch_to.window(sub_window_1)

        driver.close()

        # Switch back to original window
        driver.switch_to.window(original_window)

    if all_rows:
        export_to_excel(all_rows, filename=filename)

except Exception as e:
    print("❌ Login failed:", e)

time.sleep(10)
