import os
import shutil
from selenium import webdriver
from time import sleep
from docx import Document
from PIL import Image  # Pillow for image resizing
from docx.shared import Inches
import pyautogui

# Clear TEMP_screenshots folder and delete output_screenshots.docx
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

# Add appropriate suffixes for Studentapi, IntegrationApi, etranscriptApi, and BannerAccess Mgmt URLs
for i in range(len(urls)):
    if urls[i].endswith("StudentApi") or urls[i].endswith("StudentAPI") or urls[i].endswith("IntegrationApi") or urls[i].endswith("IntegrationAPI"):
        urls[i] += "/api/about"
    elif urls[i].endswith("eTranscriptAPI") or urls[i].endswith("eTranscriptApi"):
        urls[i] += "/status/system-details"
    elif "bam-direct" in urls[i] and urls[i].endswith("BannerAccessMgmt"):
        urls[i] += ".ws/saml/login"

# Setup Selenium WebDriver for Chrome in incognito mode
options = webdriver.ChromeOptions()
options.add_argument("--incognito")
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)  # Ensure ChromeDriver is in PATH or provide the full path

# Open the Application Navigator URL
print("Opening the Application Navigator. Please log in manually.")
driver.get(urls[0])
input("Press Enter after you have logged in successfully...")

# Open all remaining URLs in new tabs
for url in urls[1:]:  # Skip the first URL as it is already open
    driver.execute_script(f"window.open('{url}', '_blank');")

# Wait for 15 seconds to allow all tabs to load
print("Waiting for 15 seconds to let all tabs load...")
sleep(15)

# Create a Word document
doc = Document()

# Maximum width for images in the Word document (e.g., 6 inches)
max_width = Inches(6)

# Switch to each tab and take a full-screen screenshot
for i, handle in enumerate(driver.window_handles):
    driver.switch_to.window(handle)
    sleep(1)  # Optional: Wait for any dynamic content to finish loading
    screenshot_path = os.path.join(screenshot_dir, f"screenshot_{i+1}.png")

    # Capture the entire screen including the browser's search bar
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_path)

    # Open the screenshot with Pillow to resize it
    with Image.open(screenshot_path) as img:
        # Calculate the maximum dimensions in pixels
        max_width_pixels = int(max_width)
        aspect_ratio = img.height / img.width
        max_height_pixels = int(max_width_pixels * aspect_ratio)

        # Resize the image using thumbnail (in-place)
        img.thumbnail((max_width_pixels, max_height_pixels), Image.LANCZOS)  # Efficient resizing
        img.save(screenshot_path)  # Save the resized image

    # Add the resized image to the Word document
    doc.add_picture(screenshot_path, width=max_width)
    doc.add_page_break()
    print(f"Captured and resized screenshot for tab {i+1}")

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
