import csv
import time
import sys
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'

office_id = input("Insert office ID: ")
office_name = input("Insert office name: ")
username = input("Insert MLS username: ")
password = input("Insert MLS password: ")

chrome_options = Options()
chrome_options.add_argument("--disable-logging")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_argument("--headless")

webdriver_path = './chromedriver.exe'
driver = webdriver.Chrome(service=ChromeService(executable_path=webdriver_path), options=chrome_options)

print("Browser opened successfully. Navigating to login page...")

login_url = 'https://matrix.commondataplatform.com/Matrix/Login.aspx'
target_url = 'https://matrix.commondataplatform.com/Matrix/Roster/Member'

csv_filename = f"{office_name}.csv"

driver.get(login_url)
wait = WebDriverWait(driver, 20)
actions = webdriver.ActionChains(driver)
actions.send_keys(username)
actions.send_keys(Keys.TAB)
actions.send_keys(password)
actions.send_keys(Keys.RETURN)
actions.perform()

time.sleep(1)
print("Login successful.")

driver.get(target_url)
wait.until(EC.presence_of_element_located((By.ID, 'm_ucDisplayPicker_m_ddlDisplayFormats')))
print("Roster page loaded successfully.")
print("Please wait...")

actions = webdriver.ActionChains(driver)
for _ in range(7):
    actions.send_keys(Keys.TAB)
actions.send_keys(office_id)
actions.send_keys(Keys.RETURN)
actions.perform()

time.sleep(1)

display_dropdown = Select(driver.find_element(By.ID, 'm_ucDisplayPicker_m_ddlDisplayFormats'))
display_dropdown.select_by_index(1)

def scrape_page():
    time.sleep(0.2)

    name_elements = driver.find_elements(By.CSS_SELECTOR, 'span.formula.field.d107m13')
    names = [element.text.strip() for element in name_elements]

    cell_elements = driver.find_elements(By.XPATH, '//tr[td[contains(@class, "d107m7")] and td[span[contains(text(), "Cell:")]]]/td[2]/span[@class="wrapped-field"]')
    cell_contacts = [element.text.strip() for element in cell_elements]

    email_elements = driver.find_elements(By.XPATH, '//tr[td[contains(@class, "d107m7")] and td[span[contains(text(), "Email:")]]]/td[2]/span[@class="formula field"]/a')
    emails = [element.get_attribute('href').replace('mailto:', '') for element in email_elements]

    grouped_entries = list(zip(names, cell_contacts, emails))

    return grouped_entries

with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
    csvwriter = csv.writer(csvfile)
    csvwriter.writerow(['Name', 'Cell', 'Email'])

page_count = 0
total_entries = 0

while True:
    final_groups = scrape_page()
    total_entries += len(final_groups)

    with open(csv_filename, 'a', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerows(final_groups)
    
    page_count += 1

    sys.stdout.write(f'\rScraped page {page_count}...')
    sys.stdout.flush()

    try:
        next_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'javascript:__doPostBack') and text()='Next']"))
        )
        next_button.click()

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'span.formula.field.d107m13')))
        
    except Exception as e:
        break

driver.quit()

print(f"\nScraping complete. Scraped {total_entries} data entries.")
time.sleep(5)
print("Zuk was here.")
time.sleep(2)
print("Also Cheech...")
time.sleep(3)
print("Pika too...")
time.sleep(2)
