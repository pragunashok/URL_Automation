import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from docx import Document
from PIL import Image  # Pillow for image resizing
from docx.shared import Inches
import pyautogui

# Prompt user for API user credentials
appnav_username = input("Enter the Test_creds username: ")
appnav_password = input("Enter the Test_creds password: ")
api_username = input("Enter the API user username: ")
api_password = input("Enter the API user password: ")

# Clean up TEMP_screenshots folder and output_screenshots.docx
screenshot_dir = "TEMP_screenshots"
if os.path.exists(screenshot_dir):
    shutil.rmtree(screenshot_dir)  # Remove the folder and all its contents
os.makedirs(screenshot_dir, exist_ok=True)  # Recreate the folder

output_path = "output_screenshots.docx"
if os.path.exists(output_path):
    os.remove(output_path)  # Delete the existing output file

# Prompt the user to enter URLs line by line
print("Enter the list of URLs (one URL per line). Press Enter twice to finish:")
urls = []
while True:
    url = input()
    if url.strip() == "":
        break
    urls.append(url.strip())

# Ensure the first URL is the Application Navigator
if not urls:
    print("No URLs entered. Exiting.")
    exit()

# Add appropriate suffixes for specific URLs
for i in range(len(urls)):
    if "studentapi" in urls[i].lower() or "integrationapi" in urls[i].lower():
        urls[i] += "/api/about"
        parsed_url = urls[i].split("://")
        urls[i] = f"{parsed_url[0]}://{api_username}:{api_password}@{parsed_url[1]}"
        #adding login credentials to the studentapi and integrationapi URLs, this avoids having to manually enter them
    elif urls[i].endswith("eTranscriptAPI") or urls[i].endswith("eTranscriptApi"):
        urls[i] += "/status/system-details"


    elif "banneradmin" in urls[i] and urls[i].endswith("BannerAccessMgmt"):
        urls[i] += ".ws/saml/login"


print("Opening Browser Now")




# Setup Selenium WebDriver for Chrome in incognito mode
options = webdriver.ChromeOptions()
options.add_argument("--incognito")
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)  # Ensure ChromeDriver is in PATH or provide the full path

# Open the Application Navigator URL
print("Opening the Application Navigator. Please log in manually.")
driver.get(urls[0])
sleep(5)
pyautogui.typewrite(appnav_username)
pyautogui.typewrite("\n")
sleep(3)
pyautogui.typewrite(appnav_password)
pyautogui.typewrite("\n")











input("Press Enter after you have logged in successfully...")

# Open all remaining URLs in new tabs
for url in urls[1:]:  # Skip the first URL as it is already open
    driver.execute_script(f"window.open('{url}', '_blank');")

# Wait for 15 seconds to allow all tabs to load
print("Waiting for 15 seconds to let all tabs load...")
sleep(15)

# Create a Word document
doc = Document()
max_width = Inches(6)

# Switch to each tab, handle API credentials if needed, and take screenshots
for i, handle in enumerate(driver.window_handles):
    driver.switch_to.window(handle)
    sleep(1)  # Optional: Wait for any dynamic content to finish loading
    url = driver.current_url


    if "etranscriptapi" in url.lower():
        username_input = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.NAME, "username"))  # Replace 'identifier' with the actual attribute of the element
        )

        # Type a string into the input field
        username_input.send_keys(api_username)


        password_input = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "password"))  # Replace 'input-password' with the correct identifier
        )
        password_input.send_keys(api_password)
        sleep(5)



    # Take a screenshot before entering credentials
    screenshot_path = os.path.join(screenshot_dir, f"screenshot{i+1}.png")
    #driver.save_screenshot(screenshot_path)
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_path)
    
    doc.add_picture(screenshot_path, width=max_width)
    doc.add_page_break()
    print(f"Captured screenshots for tab {i+1}")

# Save the Word document
doc.save(output_path)
print(f"Document saved as {output_path}")

# Keep the browser open after script execution
print("Browser will remain open. Press Ctrl+C to exit the script and close the browser.")
try:
    while True:
        sleep(1)  # Keep the script alive
except KeyboardInterrupt:
    print("Exiting script. The browser will remain open.")
