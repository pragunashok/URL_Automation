import os
from selenium import webdriver
from time import sleep
from docx import Document

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
screenshot_dir = "screenshots"
os.makedirs(screenshot_dir, exist_ok=True)

# Switch to each tab and take a screenshot
for i, handle in enumerate(driver.window_handles):
    driver.switch_to.window(handle)
    sleep(2)  # Optional: Wait for any dynamic content to finish loading
    screenshot_path = os.path.join(screenshot_dir, f"screenshot_{i+1}.png")
    driver.save_screenshot(screenshot_path)
    doc.add_picture(screenshot_path)
    doc.add_page_break()
    print(f"Captured screenshot for tab {i+1}")

# Save the Word document
output_path = "output_screenshots.docx"
doc.save(output_path)
print(f"Document saved as {output_path}")

# Cleanup
driver.quit()
