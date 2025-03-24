import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

os.makedirs("Logs", exist_ok=True)
log_filename = os.path.join("Logs", f"log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt")
def write_log(message):
    with open(log_filename, "a", encoding="utf-8") as log_file:
        log_file.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")

def wait(by, value):
    """ รอให้ element ปรากฏก่อนทำงาน """
    return WebDriverWait(driver, 10).until(EC.presence_of_element_located((by, value)))

def wait_all(by, value):
    """ รอให้ element ทั้งหมดปรากฏก่อนทำงาน """
    return WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((by, value)))

# 📌 อ่าน Input จาก Excel
FILE_PATH = "user_config.xlsx"
df = pd.read_excel(FILE_PATH, usecols=[1, 2, 3, 4, 5], names=["Bank", "User_ID1", "User_ID2", "Password", "URL"])

# ตั้งค่า WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def UOB(UID1, UID2, password, url):
    driver.get(url)
    time.sleep(1.5) 

    try:
        username_field = driver.find_element(By.NAME, "username")
        username_field.send_keys(UID1)
        password_field = driver.find_element(By.NAME, "password") 
        password_field.send_keys(password)
        login_button = driver.find_element(By.XPATH, "//button[contains(text(),'Login')]")
        login_button.click()
        time.sleep(1.5)  

        my_account_button = driver.find_element(By.XPATH, "//a[contains(text(),'My Account')]")  # เปลี่ยน XPath ตามเว็บจริง
        my_account_button.click()

        success_message = f"✅ Success: {url}"
        print(success_message)
        write_log(success_message)
    
    except Exception as e:
        error_message = f"❌ Failed: {url}, Error: {e}"
        print(error_message)
        write_log(error_message)

def TTB(UID1, UID2, password, url):
    driver.get(url)
    time.sleep(1.5)

    try:
        accept_button = driver.find_element(By.CSS_SELECTOR, 'button.accept-button')
        accept_button.click()
        write_log("TTB: ✅ กดปุ่ม Accept")
    except:
        write_log("TTB: ✅ ไม่พบการ ปิดปรับปรุงเว็บ")

    try:
        uid1_field= wait(By.CSS_SELECTOR, 'input[type="text"]')
        uid1_field.send_keys(UID1)
        write_log("TTB: ✅ ใส่ uid1")

        next_button = wait(By.CSS_SELECTOR, 'button[data-e2e="login-button-next"]')
        next_button.click()
        write_log("TTB: ✅ กดปุ่มถัดไป")

        password_field = wait(By.CSS_SELECTOR, 'input[type="password"]')
        password_field.send_keys(password)
        write_log("TTB: ✅ ใส่ password")

        password_next_button = wait(By.CSS_SELECTOR, 'button[data-e2e="password-button-next"]')
        password_next_button.click()
        write_log("TTB: ✅ กดปุ่ม Login")

        try:
            _ = wait(By.ID, "main-menu-ACCOUNTS")
            write_log("TTB: ✅ เจอปุ่ม My Account")
        except:
            write_log("TTB: ❌ Timeout ไม่เจอปุ่ม My Account")
            
        driver.get("https://www.ttbbusinessone.com/main/accounts/groups/table)")
        write_log("TTB: ✅ เข้าหน้า My Account")

        view_statement = wait(By.ID, 'accounts-groups-table-history')
        view_statement.click()
        write_log("TTB: ✅ คลิกปุ่ม View Statement")
        time.sleep(2)

        date_dropdowns = wait_all(By.CLASS_NAME, 'dropdown-toggle')
        date_dropdowns[1].click()
        write_log("TTB: ✅ คลิก Date Dropdown")

        time.sleep(1)
        date_value = wait(By.XPATH, '//li[@labelvalue="Yesterday"]')
        date_value.click()
        write_log("TTB: ✅ เลือกวันที่")

        time.sleep(4)
        download_statement_button = wait(By.CSS_SELECTOR, 'button[data-e2e="download-operations-action"]')
        download_statement_button.click()
        write_log("TTB: ✅ กดปุ่ม Download statement")

        time.sleep(1)
        download_button = wait(By.CSS_SELECTOR, 'button[data-e2e="export-selector-download"]')
        download_button.click()
        write_log("TTB: ✅ กดปุ่ม Download")

        success_message = f"TTB: ✅✅ Success"
        write_log(success_message)

    except Exception as e:
        error_message = f"❌ Failed TTB Error: {e}"
        write_log(error_message)
    finally:
        time.sleep(3)

def LH_BANK(UID1, password, url):
    driver.get(url)
    time.sleep(1.5)

    try:
        accept_button = driver.find_element(By.CSS_SELECTOR, 'button.accept-button')
        accept_button.click()
        write_log("TTB: ✅ กดปุ่ม Accept")
    except:
        write_log("TTB: ✅ ไม่พบการ ปิดปรับปรุงเว็บ")

    try:
        uid1_field= wait(By.CSS_SELECTOR, 'input[type="text"]')
        uid1_field.send_keys(UID1)
        write_log("TTB: ✅ ใส่ uid1")

        next_button = wait(By.CSS_SELECTOR, 'button[data-e2e="login-button-next"]')
        next_button.click()
        write_log("TTB: ✅ กดปุ่มถัดไป")

        password_field = wait(By.CSS_SELECTOR, 'input[type="password"]')
        password_field.send_keys(password)
        write_log("TTB: ✅ ใส่ password")

        password_next_button = wait(By.CSS_SELECTOR, 'button[data-e2e="password-button-next"]')
        password_next_button.click()
        write_log("TTB: ✅ กดปุ่ม Login")

        try:
            _ = wait(By.ID, "main-menu-ACCOUNTS")
            write_log("TTB: ✅ เจอปุ่ม My Account")
        except:
            write_log("TTB: ❌ Timeout ไม่เจอปุ่ม My Account")
            
        driver.get("https://www.ttbbusinessone.com/main/accounts/groups/table)")
        write_log("TTB: ✅ เข้าหน้า My Account")

        view_statement = wait(By.ID, 'accounts-groups-table-history')
        view_statement.click()
        write_log("TTB: ✅ คลิกปุ่ม View Statement")
        time.sleep(2)

        date_dropdowns = wait_all(By.CLASS_NAME, 'dropdown-toggle')
        date_dropdowns[1].click()
        write_log("TTB: ✅ คลิก Date Dropdown")

        time.sleep(1)
        date_value = wait(By.XPATH, '//li[@labelvalue="Yesterday"]')
        date_value.click()
        write_log("TTB: ✅ เลือกวันที่")

        time.sleep(4)
        download_statement_button = wait(By.CSS_SELECTOR, 'button[data-e2e="download-operations-action"]')
        download_statement_button.click()
        write_log("TTB: ✅ กดปุ่ม Download statement")

        time.sleep(1)
        download_button = wait(By.CSS_SELECTOR, 'button[data-e2e="export-selector-download"]')
        download_button.click()
        write_log("TTB: ✅ กดปุ่ม Download")

        success_message = f"TTB: ✅✅ Success"
        write_log(success_message)

    except Exception as e:
        error_message = f"❌ Failed TTB Error: {e}"
        write_log(error_message)
    finally:
        time.sleep(3)

def K_BANK(UID1, password, url):
    driver.get(url)
    time.sleep(1.5) 

    try:
        username_field = driver.find_element(By.NAME, "username")
        username_field.send_keys(UID1)
        password_field = driver.find_element(By.NAME, "password") 
        password_field.send_keys(password)
        login_button = driver.find_element(By.XPATH, "//button[contains(text(),'Login')]")
        login_button.click()
        time.sleep(1.5)  

        my_account_button = driver.find_element(By.XPATH, "//a[contains(text(),'My Account')]")  # เปลี่ยน XPath ตามเว็บจริง
        my_account_button.click()

        success_message = f"✅ Success: {url}"
        print(success_message)
        write_log(success_message)
    
    except Exception as e:
        error_message = f"❌ Failed: {url}, Error: {e}"
        print(error_message)
        write_log(error_message)

def KTB(UID1, UID2, password, url):
    driver.get(url)
    time.sleep(1.5) 

    try:
        username_field = driver.find_element(By.NAME, "username")
        username_field.send_keys(UID1)
        password_field = driver.find_element(By.NAME, "password") 
        password_field.send_keys(password)
        login_button = driver.find_element(By.XPATH, "//button[contains(text(),'Login')]")
        login_button.click()
        time.sleep(1.5)  

        my_account_button = driver.find_element(By.XPATH, "//a[contains(text(),'My Account')]")  # เปลี่ยน XPath ตามเว็บจริง
        my_account_button.click()

        success_message = f"✅ Success: {url}"
        print(success_message)
        write_log(success_message)
    
    except Exception as e:
        error_message = f"❌ Failed: {url}, Error: {e}"
        print(error_message)
        write_log(error_message)

def SCB(UID1, UID2, password, url):
    driver.get(url)
    time.sleep(1.5) 

    try:
        username_field = driver.find_element(By.NAME, "username")
        username_field.send_keys(UID1)
        password_field = driver.find_element(By.NAME, "password") 
        password_field.send_keys(password)
        login_button = driver.find_element(By.XPATH, "//button[contains(text(),'Login')]")
        login_button.click()
        time.sleep(1.5)  

        my_account_button = driver.find_element(By.XPATH, "//a[contains(text(),'My Account')]")  # เปลี่ยน XPath ตามเว็บจริง
        my_account_button.click()

        success_message = f"✅ Success: {url}"
        print(success_message)
        write_log(success_message)
    
    except Exception as e:
        error_message = f"❌ Failed: {url}, Error: {e}"
        print(error_message)
        write_log(error_message)

def BAY(UID1, password, url):
    driver.get(url)
    time.sleep(1.5) 

    try:
        username_field = driver.find_element(By.NAME, "username")
        username_field.send_keys(UID1)
        password_field = driver.find_element(By.NAME, "password") 
        password_field.send_keys(password)
        login_button = driver.find_element(By.XPATH, "//button[contains(text(),'Login')]")
        login_button.click()
        time.sleep(1.5)  

        my_account_button = driver.find_element(By.XPATH, "//a[contains(text(),'My Account')]")  # เปลี่ยน XPath ตามเว็บจริง
        my_account_button.click()

        success_message = f"✅ Success: {url}"
        print(success_message)
        write_log(success_message)
    
    except Exception as e:
        error_message = f"❌ Failed: {url}, Error: {e}"
        print(error_message)
        write_log(error_message)

bank_functions = {
    # "UOB": lambda row: UOB(row["User_ID1"], row["User_ID2"], row["Password"], row["URL"]),
    "TTB BUSINESS ONE": lambda row: TTB(row["User_ID1"], row["User_ID2"], row["Password"], row["URL"]),
    # "LH BANK": lambda row: LH_BANK(row["User_ID1"], row["Password"], row["URL"]),
    # "KBANK": lambda row: K_BANK(row["User_ID1"], row["Password"], row["URL"]),
    # "KTB CORP.": lambda row: KTB(row["User_ID1"], row["User_ID2"], row["Password"], row["URL"]),
    # "SCB Bussiness Net": lambda row: SCB(row["User_ID1"], row["User_ID2"], row["Password"], row["URL"]),
    # "BAY-Krungsri Cash link": lambda row: BAY(row["User_ID1"], row["Password"], row["URL"]),
}

# Loop Bank
for _, row in df.iterrows():
    bank_function = bank_functions.get(row["Bank"])
    if bank_function:
        bank_function(row)


# ปิด Browser
driver.quit()

write_log("🛑 Script finished executing.")
