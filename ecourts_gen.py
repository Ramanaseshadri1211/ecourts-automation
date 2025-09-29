import time
import cv2
import pytesseract
import pandas as pd
import numpy as np
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from PIL import Image
from io import BytesIO
import requests

# ----------------------------------------------------------------------
# CONFIGURATION – EDIT THESE ONLY
# ----------------------------------------------------------------------
# Paths (keep them relative or configurable)
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # Change if needed
CHROME_DRIVER_PATH = r"C:\path\to\chromedriver.exe"               # Update with your chromedriver path
INPUT_EXCEL = r"input_cases.xlsx"                                  # Excel with columns: case_type, case_number, year
OUTPUT_FOLDER = r"output"                                          # Folder to store results/screenshots

# E-Courts portal URL
ECOURTS_URL = "https://services.ecourts.gov.in/ecourtindia_v6/"

# Optional default selections
STATE_CODE = "10"      # Tamil Nadu example
DISTRICT_CODE = "8"    # Example district code
COURT_COMPLEX_NAME = "Combined Courts, Tiruchirappalli"  # Example court name
# ----------------------------------------------------------------------

os.makedirs(OUTPUT_FOLDER, exist_ok=True)
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

driver = webdriver.Chrome(service=Service(CHROME_DRIVER_PATH))
wait = WebDriverWait(driver, 15)
driver.get(ECOURTS_URL)
driver.maximize_window()

# Click "Case Status" menu
case_status_link = wait.until(EC.element_to_be_clickable((By.ID, "leftPaneMenuCS")))
driver.execute_script("arguments[0].scrollIntoView(true);", case_status_link)
case_status_link.click()

# Select state / district / court complex
Select(wait.until(EC.presence_of_element_located((By.ID, "sess_state_code")))).select_by_value(STATE_CODE)
time.sleep(2)
wait.until(lambda d: any(
    option.get_attribute("value") == DISTRICT_CODE
    for option in d.find_element(By.ID, "sess_dist_code").find_elements(By.TAG_NAME, "option")
))
Select(driver.find_element(By.ID, "sess_dist_code")).select_by_value(DISTRICT_CODE)
time.sleep(2)
Select(wait.until(EC.presence_of_element_located((By.ID, "court_complex_code")))).select_by_visible_text(COURT_COMPLEX_NAME)

# Click Case Number tab
wait.until(EC.element_to_be_clickable((By.ID, "casenumber-tabMenu"))).click()

# Switch to iframe
iframe = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "iframe#iframeData")))
driver.switch_to.frame(iframe)

# Read Excel
df = pd.read_excel(INPUT_EXCEL)
for col in ["First Hearing Date", "Decision Date", "Case Status",
            "Nature of Disposal", "Court and Judge", "FIR Details", "Judgment PDF"]:
    if col not in df.columns:
        df[col] = ""

def preprocess_captcha(image_bytes):
    img = Image.open(BytesIO(image_bytes)).convert('L')
    open_cv_image = np.array(img)
    _, thresh = cv2.threshold(open_cv_image, 150, 255, cv2.THRESH_BINARY)
    return thresh

def solve_captcha():
    try:
        captcha_img = wait.until(EC.presence_of_element_located((By.ID, "captcha_image")))
        png = captcha_img.screenshot_as_png
        processed = preprocess_captcha(png)
        text = pytesseract.image_to_string(
            processed, config='--psm 8 -c tessedit_char_whitelist=0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
        )
        return ''.join(filter(str.isalnum, text.strip()))
    except:
        return ""

for index, row in df.iterrows():
    try:
        case_type = str(row['case_type']).strip()
        case_number = str(row['case_number']).strip()
        case_year = str(row['year']).strip()
        print(f"\nProcessing: {case_type} / {case_number} / {case_year}")

        # Select case type
        select_elem = wait.until(EC.presence_of_element_located((By.ID, "case_type")))
        select = Select(select_elem)
        for option in select.options:
            if option.text.strip().startswith(case_type):
                select.select_by_visible_text(option.text.strip())
                break

        driver.find_element(By.ID, "search_case_no").clear()
        driver.find_element(By.ID, "search_case_no").send_keys(case_number)
        driver.find_element(By.ID, "rgyear").clear()
        driver.find_element(By.ID, "rgyear").send_keys(case_year)

        # Solve captcha up to 3 tries
        for _ in range(3):
            captcha_text = solve_captcha()
            driver.find_element(By.ID, "case_captcha_code").clear()
            driver.find_element(By.ID, "case_captcha_code").send_keys(captcha_text)
            driver.execute_script("submitCaseNo();")
            time.sleep(4)
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'View')]")))
                break
            except TimeoutException:
                continue

        # Click View button
        view_btn = wait.until(
            lambda d: next((el for el in d.find_elements(By.XPATH, "//a[contains(text(),'View')]")
                            if "viewHistory" in (el.get_attribute("onclick") or "")), None)
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", view_btn)
        time.sleep(1)
        view_btn.click()

        # Wait for details
        details_wait = WebDriverWait(driver, 60)
        details_wait.until(EC.presence_of_element_located((
            By.XPATH, "//tr[td/label[contains(text(),'First Hearing Date')] or td/label/strong[contains(text(),'Decision Date')]]"
        )))

        def get_value(label_text):
            try:
                label_td = driver.find_element(
                    By.XPATH,
                    f"//tr[td/label[normalize-space(.)='{label_text}' or strong[normalize-space(.)='{label_text}']]]"
                )
                val_td = label_td.find_elements(By.XPATH, ".//following-sibling::td")[0]
                strong = val_td.find_elements(By.TAG_NAME, "strong")
                return strong[0].text.strip() if strong else val_td.text.strip()
            except:
                return ""

        df.at[index, "First Hearing Date"] = get_value("First Hearing Date")
        df.at[index, "Decision Date"] = get_value("Decision Date")
        df.at[index, "Case Status"] = get_value("Case Status")
        df.at[index, "Nature of Disposal"] = get_value("Nature of Disposal")
        df.at[index, "Court and Judge"] = get_value("Court Number and Judge")

        screenshot_path = os.path.join(OUTPUT_FOLDER, f"case_{case_type}_{case_number}_{case_year}.png")
        driver.save_screenshot(screenshot_path)

        # Clear inputs
        driver.find_element(By.ID, "search_case_no").clear()
        driver.find_element(By.ID, "rgyear").clear()
        driver.find_element(By.ID, "case_captcha_code").clear()

    except Exception as e:
        print(f"Error processing case {index + 1}: {e}")
        continue

# Save Excel
output_excel = os.path.join(OUTPUT_FOLDER, "output_cases.xlsx")
df.to_excel(output_excel, index=False)
print(f"\n✅ All cases processed. Output saved to: {output_excel}")
