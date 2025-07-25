import subprocess
import sys
import os
import urllib.request
import zipfile
import time
import re
import requests
import shutil
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains

# Define necessary functions
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def get_edge_version():
    try:
        output = subprocess.check_output(["reg", "query", r"HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon", "/v", "version"], shell=True, text=True)
        match = re.search(r"version\s+REG_SZ\s+([\d.]+)", output)
        if match:
            return match.group(1)
        else:
            raise Exception("Edge version not found.")
    except Exception as e:
        print(f"Error fetching Edge version: {e}")
        return None

def get_driver_version(driver_path):
    try:
        output = subprocess.check_output([driver_path, "--version"], text=True)
        match = re.search(r"(\d+\.\d+\.\d+\.\d+)", output)
        if match:
            return match.group(1)
        else:
            raise Exception("WebDriver version not found.")
    except Exception as e:
        print(f"Error fetching WebDriver version: {e}")
        return None

def download_edge_webdriver(version, download_path):
    url = f"https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip"
    static_url = "https://msedgedriver.microsoft.com/140.0.3456.0/edgedriver_win64.zip"
    print(f"Downloading Edge WebDriver version {version} from {static_url}...")
    try:
        with requests.get(static_url, stream=True) as r:
            r.raise_for_status()
            with open(download_path, 'wb') as f:
                shutil.copyfileobj(r.raw, f)
        print("Download complete.")
    except requests.RequestException as e:
        print(f"Download failed: {e}")

def extract_zip(zip_path, extract_to):
    print(f"Extracting {zip_path} to {extract_to}...")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    print("Extraction complete.")

# Paths and variables
webdriver_dir = os.path.join(os.getcwd(), "msedgedriver")
webdriver_path = os.path.join(webdriver_dir, "msedgedriver.exe")

# Check Edge version
edge_version = get_edge_version()
print(f"Here is the Windows Edge Browser version!: {edge_version}")
if not edge_version:
    print("Unable to determine Microsoft Edge version. Exiting.")
    sys.exit(1)

# Check if WebDriver exists and matches the Edge version
if os.path.exists(webdriver_path):
    driver_version = get_driver_version(webdriver_path)
    if driver_version == edge_version:
        print(f"WebDriver version {driver_version} matches Edge version {edge_version}. Proceeding...")
    else:
        print(f"WebDriver version {driver_version} does not match Edge version {edge_version}.")
        print("\nThe existing WebDriver directory must be deleted to download the correct version.")
        print("Please manually delete the directory at the following path and rerun this script:")
        print(f"{webdriver_dir}")
        input("Press Enter to exit...")
        sys.exit(1)
else:
    print("WebDriver not found. Proceeding with download and setup...")

# Download and extract the correct WebDriver
if not os.path.exists(webdriver_path):
    os.makedirs(webdriver_dir, exist_ok=True)
    zip_path = os.path.join(webdriver_dir, "msedgedriver.zip")
    download_edge_webdriver(edge_version, zip_path)
    extract_zip(zip_path, webdriver_dir)
    os.remove(zip_path)

print("Edge WebDriver is ready.")

# Initialize WebDriver with the correct path
service = Service(webdriver_path)
edge_options = Options()
# edge_options.add_argument("--headless")  # Optional: Run in headless mode
driver = webdriver.Edge(service=service, options=edge_options)

try:
    # Step 1: Go to the login page
    driver.get("https://www.phoneburner.com/homepage/login")
    print("\nPage opened...")

    # Step 2: Locate and fill in the username and password fields
    username_field = driver.find_element(By.ID, 'f_username')
    password_field = driver.find_element(By.ID, 'f_password')

    username_given = input("Enter your email: ")
    firstname = username_given.split('@')[0].capitalize()
    password_given = input("Enter your password: ")

    username_field.send_keys(username_given)
    password_field.send_keys(password_given)

    # Step 3: Submit the login form
    login_button = driver.find_element(By.CSS_SELECTOR, 'button.btn.btn-primary.btn-lg.w-100')
    login_button.click()
    print("Logging in...")

    # Give time for the login to process and redirect
    time.sleep(3)

    # Step 4: Navigate to the page with the table
    driver.get("https://www.phoneburner.com/cm/index#params/dmlld19pZD0yNzY2MjU1JnBhZ2U9MQ==")
    print("Logged in!")

    # Locate all folder elements by their class name
    folder_elements = driver.find_elements(By.CLASS_NAME, 'contacts-folder-nav-item')

    # Create a list to store folder names and IDs
    folders = []
    final_folder = 0

    # Loop through the folder elements to extract names and IDs
    for index, folder in enumerate(folder_elements):
        folder_name = folder.find_element(By.CLASS_NAME, 'contacts-folder-nav-name').text
        folder_id = folder.get_attribute('id')
        folders.append((folder_name, folder_id))
        if folder_name:
            print(f"{index + 1}. {folder_name}")
        else:
            final_folder = index
            break

    # Ask the user to input the number corresponding to the folder they want to enter
    while True:
        try:
            user_choice = int(input("Enter the number of the folder you want to enter: ")) - 1
            if 0 <= user_choice < len(folders):
                chosen_folder = folders[user_choice]
                break
            else:
                print(f"Please enter a number from 1 to {final_folder}.")
        except ValueError:
            print("Invalid input. Please enter a number.")

    # Navigate to the chosen folder
    chosen_folder = folders[user_choice]
    folder_element = driver.find_element(By.ID, chosen_folder[1])
    folder_element.click()

    # Wait for the folder content to load
    driver.implicitly_wait(10)

    # Step 5: Locate the table by its ID and get the HTML content
    table_element = driver.find_element(By.ID, 'main_contact_grid')

    # Step 6: Extract all rows in the table
    emails = set()  # No duplicates
    rows = table_element.find_elements(By.CSS_SELECTOR, 'tr.contact-new')
    rows = [row for index, row in enumerate(rows) if index % 2 == 0]
    print("Number of DNC Contacts Found =", len(rows))
    print()

    for row in rows:
        # Find all <a> tags within the row
        links = row.find_elements(By.TAG_NAME, 'a')
        for link in links:
            # Check if the <a> tag's href attribute contains mailto:
            href = link.get_attribute('href')
            if href and href.startswith('mailto:'):
                email = href[len('mailto:'):]
                print(email)
                emails.add(email)

        # Click on the checkbox to select the contact
        checkbox = WebDriverWait(row, 15).until(
            EC.element_to_be_clickable((By.XPATH, ".//input[@type='checkbox']"))
        )
        driver.execute_script("arguments[0].click();", checkbox)

    # Select the DNC option
    move_dropdown_button = driver.find_element(By.ID, "cm_move_button")
    move_dropdown_button.click()

    option_to_select = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//ul[@id='cm_move_dropdown']//li//a[text()='DNC']"))
    )
    option_to_select.click()
    time.sleep(2)

    # Handle potential modal confirmation
    if len(rows) > 1:
        try:
            modal_present = WebDriverWait(driver, 45).until(
                EC.presence_of_element_located((By.CLASS_NAME, "btn-primary"))
            )
            driver.execute_script("arguments[0].click();", modal_present)

            confirm_input_box = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.input-group input.form-control"))
            )
            confirm_input_box.send_keys("confirm")

            okay_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Okay']"))
            )
            driver.execute_script("arguments[0].removeAttribute('disabled')", okay_button)
            driver.execute_script("arguments[0].scrollIntoView(true);", okay_button)
            actions = ActionChains(driver)
            actions.move_to_element(okay_button).click().perform()
        except (TimeoutException, NoSuchElementException):
            print("No confirmation input box needed")
        time.sleep(4)

finally:
    # Ensure the browser is closed
    driver.quit()
    print("PhoneBurner has been closed!")

emails = list(emails)
print("Opening Outlook and loading email...")
# Initialize Outlook application
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
# Create a new mail item
mail = outlook.CreateItem(0)  # 0: olMailItem
# Set email subject
mail.Subject = 'Getting your truck insurance quotes!'
# Set email body (HTML or plain text)
mail.Body = f"""Hey! This is {firstname} from The Insurance Store, I saw that your insurance renewal is coming up in a few weeks. I wanted to let you know about a new exclusive truck insurance market we have that is looking to insure high caliber trucking companies such as yours. Pricing is often 30% cheaper than the competition with superior coverage. Please let me know if you are interested and we'd be happy to get you a quote within 1-2 days!

When you have a chance, could you review the following questions/Information below and answer them the best you can:
1. Owner and Drivers Driver's License Number, Birthday, Years of Experience
2. Verify Mailing and Garage Address
3. Scheduled Vehicle/Trailer List and their Listed Values
4. Cargo Coverage value and Top 3 Types of Cargo most often hauled
5. Average working Radius/Furthest City Traveled
6. Target Premium for this year that you would like to be at, or the current premium paid.
7. If you've got any current Loss Runs or current Certificate of Insurance, those help in making our quotes more competitive!

Looking forward to hearing from you soon, thanks so much!!
"""
 
# Add email addresses to BCC
mail.BCC = ";".join(emails)
# Add email sent to yourself
mail.To = username_given
 
# Display the email (this will open the Outlook email editor with the email populated)
mail.Display(True)
