from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import zipfile
import os

def unzip_file(zip_file_path, extract_to_path):
    # Ensure the extraction directory exists
    os.makedirs(extract_to_path, exist_ok=True)

    # Open the zip file
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        # Extract all files to the specified directory
        zip_ref.extractall(extract_to_path)
        print(f"Extracted all files to: {extract_to_path}")



def download_asus_bios(model: str, current_version: int, download_path: str):
    # Set up the WebDriver (e.g., ChromeDriver)
    options = webdriver.ChromeOptions()
    chrome_driver_path = "E:/chromedriver/chromedriver.exe"  # Adjust the path to your ChromeDriver
    #options.add_argument("--headless")  # Run in headless mode if you don't need
    prefs = {"download.default_directory": download_path}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(executable_path=chrome_driver_path,options=options)

    # Navigate to the ASUS support website
    driver.get("https://www.asus.com/supportonly/"+model.lower()+"/helpdesk_bios/")
    driver.maximize_window()
    # Search for the model

    # Find all div elements with a class name that includes "roductSupportDriverBIOS__fileInfo"
    div_elements = driver.find_elements(By.CSS_SELECTOR, "div[class*='ProductSupportDriverBIOS__contentLeft'] div")
   
    for div in div_elements:
 
        # Check if the div element contains the text "BIOS for ASUS EZ Flash Utility"
        if "BIOS for ASUS EZ Flash Utility" in div.text:
            next_div = div.find_elements(By.XPATH, "following-sibling::div/div")
            for div2_ in next_div:
                # Check if the div2_ element contains the text "Version"
                if "Version" in div2_.text:
                    version = div2_.text.replace("Version", "").strip()
                    
                    #If the version is greater than the current version, print it
                    if int(version) > current_version:
                        print(f"New BIOS version available: {version}")

                        ## Find the download link
                        download_link =  div.find_element(By.XPATH, "../../div[contains(@class,'ProductSupportDriverBIOS__contentRight')]/div[contains(@class,'ProductSupportDriverBIOS__downloadBtn__')]")
                        print("Downloading BIOS...")
                        driver.execute_script("arguments[0].scrollIntoView();", download_link)
                        download_link.click()
                        time.sleep(5)  # Wait for the download to start
            
                         # Unzip the downloaded file
                        downloaded_zip = os.path.join(download_path, model.upper()+"AS"+version+".zip")  # Replace with the actual file name
                        unzip_file(downloaded_zip, download_path)

                        # Remove the original zip file
                        if os.path.exists(downloaded_zip):
                            os.remove(downloaded_zip)
                            print(f"Removed original zip file: {downloaded_zip}")

    # Close the browser
    driver.quit()


# Example usage
download_asus_bios("tn3604ya", 300 , "E:\\")