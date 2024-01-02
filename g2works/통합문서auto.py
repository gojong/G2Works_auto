from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import os
import re
import shutil
# Play a custom MP3 sound file
import pygame

# Set up Selenium
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

# Provide the path to your chromedriver executable
CHROME_DRIVER_PATH = 'chromedriver.exe'
service = ChromeService(executable_path=CHROME_DRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Navigate to the login page
driver.get("https://whoisg.kr/")

# Find and input the username (ID)
username_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'input#emp_no'))
)
username_input.send_keys("ID")

# Find and input the password
password_input = driver.find_element(By.CSS_SELECTOR, 'input#passwd')
password_input.send_keys("PW")

# Find and click the "닫기" (Close) button if it exists
close_button = WebDriverWait(driver, 5).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'a.btn1[href="javascript:closePopup();"]'))
)
close_button.click()

# Find and click the login button
login_button = driver.find_element(By.CSS_SELECTOR, 'a[onclick="sendit();"]')
login_button.click()

# Wait for 5 seconds after a successful login
time.sleep(5)

# Now navigate to the desired URL
desired_url = "https://whoisg.kr/doc.auth.do?command=myList4Admin&menuCd=200318&parentMenuCd=100006"
driver.get(desired_url)

# Wait for the page to load
WebDriverWait(driver, 10).until(EC.url_to_be(desired_url))

# Input start date
start_date_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'input#startDate'))
)
driver.execute_script("arguments[0].value = '2020-07-01';", start_date_input)

# Input end date
end_date_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'input#endDate'))
)
driver.execute_script("arguments[0].value = '2020-09-09';", end_date_input)

# Click the 검색 (Search) button
search_button = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'a#btn_search'))
)
search_button.click()

# Add a waiting time after clicking the search button (adjust as needed)
time.sleep(5)

# Function to sanitize folder and file names
def sanitize_name(name):
    # Replace invalid characters with underscores
    sanitized_name = re.sub(r'[\/:*?"<>|]', '_', name)
    return sanitized_name

def create_folder(date_str, title):
    sanitized_title = sanitize_name(title)
    folder_name = f"{date_str} - {sanitized_title}"
    folder_path = os.path.join("E:\\2023\\2020\\7-9", folder_name)

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    return folder_path

# Access the "List" and download documents
page_number = 1  # Start with the first page

while True:
    for i in range(2, 22):  # Loop through 20 items on each page
        # Get date information
        date_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f'//form/div/div/div[2]/div/table/tbody/tr[{i}]/td[8]'))
        )
        date_str = date_element.text.split(' ')[0]  # Extract date portion

        # Get title information
        title_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f'//form/div/div/div[2]/div/table/tbody/tr[{i}]/td[2]/span/span'))
        )
        title = title_element.text

        # Create a folder based on date and title
        folder_path = create_folder(date_str, title)

        # Navigate to the tr[i] link
        tr_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f'//form/div/div/div[2]/div/table/tbody/tr[{i}]'))
        )
        tr_link.click()

        # Download all documents if available and move them to the created folder
        try:
            download_buttons = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, '//a[contains(@href, "doc.auth.do?command=download") and @class="btn1"]'))
            )

            for j, download_button in enumerate(download_buttons):
                try:
                    download_button.click()
                    time.sleep(5)  # Adjust this waiting time based on your requirements
                    print(f"Downloaded document {j + 1} for item {i}")

                    # Identify the recently downloaded file in the Downloads directory
                    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
                    downloaded_file = max(
                        [os.path.join(downloads_folder, f) for f in os.listdir(downloads_folder)],
                        key=os.path.getctime
                    )

                    # Move the downloaded file to the created folder
                    new_file_path = os.path.join(folder_path, os.path.basename(downloaded_file))
                    shutil.move(downloaded_file, new_file_path)
                    print(f"Moved document {j + 1} to folder: {folder_path}")

                except Exception as e:
                    print(f"Error downloading document {j + 1} for item {i}: {e}")

        except TimeoutException:
            print(f"Timeout: No download links found for item {i}")

        # Download the payment document if available and move it to the created folder
        try:
            download_payment_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//a[@id="btn_pdf" and contains(@class, "btn1")]'))
            )
            download_payment_button.click()
            time.sleep(5)  # Adjust this waiting time based on your requirements
            print(f"Downloaded payment document for item {i}")

            # Identify the recently downloaded file in the Downloads directory
            downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            downloaded_payment_file = max(
                [os.path.join(downloads_folder, f) for f in os.listdir(downloads_folder)],
                key=os.path.getctime
            )

            # Move the downloaded payment file to the created folder
            new_payment_path = os.path.join(folder_path, os.path.basename(downloaded_payment_file))
            shutil.move(downloaded_payment_file, new_payment_path)
            print(f"Moved payment document to folder: {folder_path}")

        except:
            print(f"No payment document download link found for item {i}")

        # Go back to the document list
        driver.back()

    try:
        # Click on the next page link
        next_page_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f'//a[contains(@href, "refreshPage({page_number + 1})")]'))
        )
        next_page_link.click()
        page_number += 1
        time.sleep(5)  # Adjust this waiting time based on your requirements

    except TimeoutException:
        print(f"No next page found. Exiting loop.")
        break

# Additional waiting time (you may need to customize this based on your specific situation)
time.sleep(5)

# Now you have navigated through all pages, downloaded documents, and clicked on the "결재문서 다운로드" link
print("Logged in. Current URL:", driver.current_url)

# Close the browser
driver.quit()

# Add a delay before playing the sound
time.sleep(10)  # Adjust as needed

# Initialize the mixer (only need to do this once)
pygame.mixer.init()

# Replace 'path/to/your/soundfile.mp3' with the path to your MP3 sound file
sound_file_path = r'일해라.mp3'

# Load and play the sound
pygame.mixer.music.load(sound_file_path)
pygame.mixer.music.play()

# Wait for the sound to finish playing
while pygame.mixer.music.get_busy():
    pygame.time.Clock().tick(10)