import tkinter as tk
from tkinter import ttk, messagebox
from threading import Thread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException
import time
import pyautogui

# Set Chrome preferences
chrome_options = Options()

import os
import platform

# Determine the default download directory based on the operating system
if platform.system() == "Windows":
    download_dir = os.path.join(os.environ["USERPROFILE"], "Downloads")
elif platform.system() == "Darwin":  # macOS
    download_dir = os.path.join(os.environ["HOME"], "Downloads")
else:  # For Linux or other systems
    download_dir = os.path.join(os.environ["HOME"], "Downloads")

# Chrome preferences for PDF download
prefs = {
    "download.default_directory": download_dir,  # Set default download directory
    "plugins.always_open_pdf_externally": True,  # Force PDF files to be downloaded
    "profile.default_content_settings.popups": 0,
    "safebrowsing.enabled": "false",  # Disable safe browsing (avoid unsafe download warnings)
    "safebrowsing.disable_download_protection": True,  # Disable download protection
    "download.prompt_for_download": False,  # Disable the "Save as" dialog
    "download.directory_upgrade": True,  # Allow the directory upgrade if necessary
    "safebrowsing.disable_extension": True,
}

chrome_options.add_experimental_option("prefs", prefs)


# Placeholder automation functions for the four websites
def automate_site_1(city, pcode, dept_name, section, parcel):
    # Implement the provided logic for Site 1 here
    # The final version should include `driver` initialization and scraping steps
    # Use try-except to handle errors and return status
    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.get("https://cadastre.gouv.fr/scpc/accueil.do")
        # Simulate automation steps...
        print(f"Automating Site 1 with {city}, {pcode}, {dept_name}, {section}, {parcel}")

        wait = WebDriverWait(driver, 30)  # Increased timeout for waits (30 seconds)

        # Accept preferences (click the preference button)
        prefer_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[1]/form/div[2]/p/a')))
        prefer_button.click()

        # Locate main form and partitions
        form = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[1]/form")))
        partition1 = form.find_element(By.ID, "affichageForm")
        partition2 = form.find_element(By.ID, "divparparcelle")

        city = city.replace("-", " ")

        # Fill in city and postal code fields
        city_field = partition1.find_element(By.XPATH, '/html/body/div[1]/div[1]/form/div[1]/div/p[1]/input')
        city_field.send_keys(city)
        
        pcode_field = partition1.find_element(By.XPATH, '/html/body/div[1]/div[1]/form/div[1]/div/p[2]/input')
        pcode_field.send_keys(pcode)

        # Select department name from dropdown
        dept_dropdown = partition1.find_element(By.XPATH, '/html/body/div[1]/div[1]/form/div[1]/div/p[3]/select')
        options = dept_dropdown.find_elements(By.TAG_NAME, "option")
        for option in options:
            if dept_name in option.text:
                option.click()
                break

        # Fill in section and parcel fields
        section_field = partition2.find_element(By.XPATH, '/html/body/div[1]/div[1]/form/div[2]/div/p[2]/input')
        section_field.send_keys(section)
        
        parcel_field = partition2.find_element(By.XPATH, '/html/body/div[1]/div[1]/form/div[2]/div/p[4]/input')
        parcel_field.send_keys(parcel)

        # Click the rechecker button to search
        rechecker_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#rech > p > input[type=image]")))
        rechecker_button.click()

        try:
            # # Click on the file link to open PDF page
            file_link = WebDriverWait(driver, 7).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[1]/div[2]/table/tbody/tr[2]/td[1]/a')))
            file_link.click()
        except (TimeoutException, NoSuchElementException):
            print("exception pakri gyi in file link")
            # Choose correct city option
            city_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/form/div[1]/div/p[1]/select')))
            options = city_dropdown.find_elements(By.TAG_NAME, "option")
            city_found = False
            for option in options:
                if city.lower() == option.text.lower():
                    option.click()
                    city_found = True
                    break

            # If no exact match, check for partial match
            if not city_found:
                city_parts = city.split()
                for option in options:
                    if city_parts[1].lower() in option.text.lower():
                        option.click()
                        city_found = True
                        break

            # if not city_found:
            #     city_modified = city.replace("-", " ")
            #     for option in options:
            #         if city_modified.lower() == option.text.lower():
            #             option.click()
            #             city_found = True
            #             break

            time.sleep(3)
            
            # Click search button again
            rechecker_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#rech > p > input[type=image]")))
            rechecker_button.click()

            # # Click on the file link to open PDF page
            file_link = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[1]/div[2]/table/tbody/tr[2]/td[1]/a')))
            file_link.click()


        # Switch to the new window
        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[1])

        try:
            # Click on the PDF button
            print_container = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/h3[2]")))
            print_container.click()

            # time.sleep(3)
            
            # Attempt to click on the PDF download button
            get_pdf_button = WebDriverWait(driver, 7).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/p/a')))
            get_pdf_button.click()
                
        except (NoSuchElementException, TimeoutException):
            # Retry clicking on the PDF button and download button
            try:
                print_container = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/h3[2]")))
                print_container.click()

                get_pdf_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/p/a')))
                get_pdf_button.click()
                
            except TimeoutException:
                print("Failed to locate the PDF download button after retrying.")

        # Switch to the last tab to view the PDF
        # wait.until(EC.number_of_windows_to_be(3))
        time.sleep(10)
        driver.switch_to.window(driver.window_handles[-1])

        try:
            # Click on the PDF button
            info_container = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[1]/h3[1]")))
            info_container.click()

            # USER CLICKS ON MAP
            time.sleep(10)
            
            # Attempt to click on the PDF download button
            validate_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/p/a')))
            validate_button.click()
    
            time.sleep(5)

            get_pdf_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/p/a')))
            get_pdf_button.click()

        except TimeoutException:
                print("Failed to locate the Information PDF download button.")



        # time.sleep(3)
        # driver.switch_to.window(driver.window_handles[-1])


        print("PDF opened successfully.")

        # Wait for a longer time to ensure PDF is fully loaded
        time.sleep(15)  # Increased wait time to allow PDF to load fully

        # input("Perform manual actions in browser, then press Enter here...")
        
        return "Success"
    except Exception as e:
        return f"Failure: {e}"

def automate_site_2(city, pcode, dept_name, section, parcel):
    try:
                
        driver = webdriver.Chrome(options=chrome_options)
        driver.get("https://www.geoportail-urbanisme.gouv.fr/map")

        try:
            # Wait for and click the "Prefer" button if available
            prefer_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[1]/div/div[1]/form/div[2]/div[3]/div/a"))
            )
            prefer_button.click()
        except (NoSuchElementException, TimeoutException, ElementClickInterceptedException):
            # If "Prefer" button not found, try closing any help popup and retrying
            try:
                close_help = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[8]/a[2]"))
                )
                close_help.click()

                # Retry clicking the "Prefer" button
                prefer_button = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[1]/div/div[1]/form/div[2]/div[3]/div/a"))
                )
                prefer_button.click()
            except (NoSuchElementException, TimeoutException):
                print("Prefer button or close help option not found.")

        # Select Department from dropdown
        dept_dropdown = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/form/div[2]/div[1]/select'))
        )
        dept_options = dept_dropdown.find_elements(By.TAG_NAME, "option")
        for option in dept_options:
            if dept_name.lower() in option.text.lower():
                option.click()
                break

        time.sleep(5)

        # Select City from dropdown
        city_dropdown = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/form/div[2]/div[2]/select'))
        )
        city_options = city_dropdown.find_elements(By.TAG_NAME, "option")
        for option in city_options:
            if city.lower() in option.text.lower():
                option.click()
                break

        time.sleep(5)

        # Input Section
        section_field = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/form/div[2]/div[5]/input"))
        )
        section_field.send_keys(section)
        section_dropdown = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/form/div[2]/div[5]/ul'))
        )
        section_options = section_dropdown.find_elements(By.XPATH, ".//a[@class='dropdown-item']")
        for option in section_options:
            if section.lower() == option.text.lower():
                option.click()
                break

        time.sleep(7)

        # Input Parcel
        parcel_field = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/form/div[2]/div[6]/input"))
        )
        parcel_field.send_keys(parcel)
        parcel_dropdown = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/form/div[2]/div[6]/ul'))
        )
        parcel_options = parcel_dropdown.find_elements(By.XPATH, ".//a[@class='dropdown-item']")
        for option in parcel_options:
            if parcel.lower() == option.text.lower():
                option.click()
                break


        time.sleep(2)

        # Click on Search button
        search_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/form/div[2]/div[7]/button"))
        )
        search_button.click()

        time.sleep(2)


        # Click on "Ensemble" button
        outer_button = WebDriverWait(driver, 30).until(     
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[10]/div/div[1]/div[2]/div[4]/div/div[2]/div/div[3]/a/label"))
        )
        outer_button.click()


        # Click on "Règlements" button
        inner_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'ficheinfo-pdf-collapse') and contains(text(), 'Règlements')]"))
        )
        inner_button.click()

        # Click to get PDF
        get_pdf_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'gpu-blue') and contains(text(), 'Règlement écrit')]"))
        )
        get_pdf_button.click()

        time.sleep(30)

        print("Process completed successfully.")

        # input("Perform manual actions in browser, then press Enter here...")
        
        return "Success"
    except Exception as e:
        return f"Failure: {e}"

# Similar placeholder functions for Site 3 and Site 4 can be added here

def automate_site_3(city, pcode, dept_name, section, parcel):
    # Implement the provided logic for Site 1 here
    # The final version should include `driver` initialization and scraping steps
    # Use try-except to handle errors and return status
    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.get("http://www.pprn972.fr/carto/web/")
        # Simulate automation steps...
        print(f"Automating Site 3 with {city}, {pcode}, {dept_name}, {section}, {parcel}")
        
        time.sleep(5)
        
        try:
            city_dropdown = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/main/div[4]/div/div[2]/div/div"))
            )
            city_dropdown.click()
        except TimeoutException:
            print("City dropdown not found or not clickable")
        
        time.sleep(2)
        
        try:
            options_cont = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/main/div[4]/div/div[2]/div/div/ul"))
            )
            options = options_cont.find_elements(By.TAG_NAME, "li")
            
            for option in options:
                exact_option = option.find_element(By.TAG_NAME, "span")
                if (city.lower() == (exact_option.text).lower()):
                    print(exact_option.text)
                    option.click()
                    break
        except (NoSuchElementException, TimeoutException):
            print("City options not found or clickable")

        time.sleep(5)

        try:
            section_field = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/main/div[5]/div/form/div[1]/div/div"))
            )
            section_field.click()
        except TimeoutException:
            print("Section field not found or not clickable")
        
        time.sleep(2)

        try:
            options_cont = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/main/div[5]/div/form/div[1]/div/div/ul"))
            )
            options = options_cont.find_elements(By.TAG_NAME, "li")
            
            for option in options:
                exact_option = option.find_element(By.TAG_NAME, "span")
                if (section.lower() == (exact_option.text).lower()):
                    print(exact_option.text)
                    option.click()
                    break
        except (NoSuchElementException, TimeoutException):
            print("Section options not found or clickable")
        
        try:
            parcel_field = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/main/div[5]/div/form/div[2]/div/input"))
            )
            parcel_field.send_keys(parcel)
        except TimeoutException:
            print("Parcel field not found or not interactable")

        time.sleep(3)
        
        try:
            rechercher_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/main/div[5]/div/div[2]/a[2]"))
            )
            rechercher_button.click()
        except TimeoutException:
            print("Rechercher button not found or not clickable")

        time.sleep(5)
        
        # New Page
        try:
            container = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/main/div[7]/div[2]/div[7]/div[2]/div"))
            )
            
            # Scroll to the container
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", container)
            
            container_buttons = container.find_elements(By.TAG_NAME, "p")
            
            for button in container_buttons:
                if "Localisation" in button.text:
                    get_pdf_button = button.find_element(By.TAG_NAME, "a")
                    
                    # Scroll to the button to ensure visibility
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", get_pdf_button)
                    
                    # Click the button using JavaScript to bypass interception
                    driver.execute_script("arguments[0].click();", get_pdf_button)
                    break
        except (NoSuchElementException, TimeoutException):
            print("Container or buttons not found")
        
        time.sleep(5)

        pyautogui.click(1118, 144)  # Coordinates of the "Keep" button on your screen (adjust as needed)
    
        # Give some time for the download to complete
        time.sleep(18)

        print("Process completed successfully.")

        # input("Perform manual actions in browser, then press Enter here...")
        
        return "Success"
    except Exception as e:
        return f"Failure: {e}"

def automate_site_4(city, pcode, dept_name, section, parcel):
    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.get("https://errial.georisques.gouv.fr/#/")

        # Simulate automation steps...
        print(f"Automating Site 4 with {city}, {pcode}, {dept_name}, {section}, {parcel}")
        wait = WebDriverWait(driver, 30)
        
        # Handle cookies pop-up
        try:
            accept_cookies = wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/footer/div[1]/div[3]/div/div[2]/button[2]"))
            )
            accept_cookies.click()
        except TimeoutException:
            print("No cookies pop-up found, continuing...")

        import time

        # Input city in commune field
        try:
            commune_field = wait.until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/main/section/div/div[1]/div[2]/div[3]/input[1]"))
            )
            commune_field.send_keys(city)
            time.sleep(7)

            # Wait to load dropdown and select city
            commune_dropdown_cont = wait.until(
                EC.presence_of_element_located((
                    By.XPATH, "//div/main/section/div/div[1]/div[2]/div[3]/*[contains(@class, 'autocomplete-options-wrapper')]"
                ))
            )
            commune_dropdown_ul = commune_dropdown_cont.find_element(By.TAG_NAME, "ul")
            options = commune_dropdown_ul.find_elements(By.TAG_NAME, "li")

            # print(pcode + " - " + city.lower())
            flag = 0
            for option in options:
                if (pcode + " - " + city.lower()) in option.text.lower():
                    print(option.text)
                    option.click()
                    flag = 1
                    break

            if flag==0:
                for option in options:
                    if city.lower() in option.text.lower():
                        print(option.text)
                        option.click()
                        break

        except TimeoutException:
            print("City field or dropdown not found, skipping city selection.")
        except NoSuchElementException:
            print("Dropdown elements not found, skipping city selection.")

        # Input parcel
        try:
            parcel_field = wait.until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div/div/main/section/div/div[1]/div[2]/div[4]/input"))
            )
            parcel_field.send_keys(section + "-" + parcel)
            time.sleep(3)
        except TimeoutException:
            print("Parcel field not found, skipping parcel input.")

        # Click show button
        try:
            show_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/main/section/div/div[1]/div[2]/div[6]/button"))
            )
            show_button.click()
            time.sleep(2)
        except TimeoutException:
            print("Show button not found, skipping button click.")

        # Scroll and click green button on new page
        try:
            driver.execute_script("window.scrollBy(0, 3700);")
            time.sleep(1)
            green_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/main/div[2]/div/a"))
            )
            green_button.click()
        except (TimeoutException, NoSuchElementException):
            try:
                driver.execute_script("window.scrollBy(0, 150);")
                time.sleep(1)
                green_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/main/div[2]/div/a"))
                )
                green_button.click()
            except TimeoutException:            
                print("Green button not found, skipping button click.")
            
        # Interact with checkboxes
        try:
            check_box1 = wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/main/div[3]/section/div[6]/div/div[3]/label[2]/input"))
            )
            check_box2 = wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/main/div[3]/section/div[6]/div/div[5]/label[2]/input"))
            )
            check_box3 = wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/main/div[3]/section/div[7]/div/div[2]/label[2]/input"))
            )

            check_box1.click()
            check_box2.click()
            check_box3.click()
            time.sleep(1)
        except TimeoutException:
            print("One or more checkboxes not found, skipping checkbox interaction.")

        # Scroll and click green button on next page
        try:
            driver.execute_script("window.scrollBy(0, 200);")
            time.sleep(1)
            green_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/main/div[3]/section/div[8]/a"))
            )
            green_button.click()
            time.sleep(1)
        except TimeoutException:
            print("Green button on second page not found, skipping button click.")

        # Download PDF
        try:
            get_pdf = wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/main/div[4]/section/div[5]/div/div[3]/a"))
            )
            get_pdf.click()
        except TimeoutException:
            print("Download PDF button not found, skipping PDF download.")


        # time.sleep(15)  # Wait to load and then auto download 
        
        time.sleep(55)

        print("Process completed successfully.")

        # input("Perform manual actions in browser, then press Enter here...")
        
        return "Success"
    
    except Exception as e:
        return f"Failure: {e}"


def automate_site_5(city, pcode, dept_name, section, parcel):
    try:
        
        # Set Chrome preferences
        chrome_options = Options()

        # Create the driver
        driver = webdriver.Chrome(options=chrome_options)
        
        wait = WebDriverWait(driver, 30)  # Set wait time for up to 30 seconds
        
        # Initialize the WebDriver
        driver.get("http://e-plu-martinique.com")  # Keep the URL unchanged as per the original task
        time.sleep(5)
        
        # Example automation step 1: Clicking an accept button
        try:
            accept_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div[2]/div/div/div[1]")))
            accept_button.click()

            ok_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div[2]/button")))
            ok_button.click()

        except TimeoutException:
            print("Accept button not found. Moving on.")
        
        time.sleep(3)

        try:
            get_search_bar = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[6]/div/div[3]/div/div[3]/div[1]")))
            get_search_bar.click()
        except TimeoutException:
            get_search_bar = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[7]/div/div[3]/div/div[3]/div[1]")))
            get_search_bar.click()

        time.sleep(5)
        
        # Wait and click commune empty                                          /html/body/div[2]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/ul/li/div[2]/div/div/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/div/div[1]/div[2]
        commune_empty_click = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[8]/div[2]/div/div/div[1]/ul/li/div[2]/div/div/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/div/div[1]/div[2]")))
        commune_empty_click.click()

        time.sleep(5)

        # city = "LE ROBERT"

        time.sleep(3)

        try:
            commune_input = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[5]/div[1]/div/div[2]/div[1]/div/div[2]/input")
            commune_input.send_keys(city)                  

            time.sleep(3)

            city_dropdown = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[5]/div[1]/div/div[2]/div[2]/div[1]")
            city_options = city_dropdown.find_elements(By.CLASS_NAME, "item")
        except:
            commune_input = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[3]/div[1]/div/div[2]/div[1]/div/div[2]/input")
            commune_input.send_keys(city)                  

            time.sleep(3)    

            city_dropdown = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[3]/div[1]/div/div[2]/div[2]/div[1]")
            city_options = city_dropdown.find_elements(By.CLASS_NAME, "item")
            

        city_found = False
        for option in city_options:
            if city.lower() in option.text.lower():
                option.click()
                city_found = True
                break

        if not city_found:
            city_parts = city.split()
            # commune_input = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[5]/div[1]/div/div[2]/div[1]/div/div[2]/input")
            commune_input.clear()

            try:
                commune_input = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[5]/div[1]/div/div[2]/div[1]/div/div[2]/input")
                commune_input.send_keys(city_parts[1])
            
                time.sleep(3)

                city_dropdown = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[5]/div[1]/div/div[2]/div[2]/div[1]")
                city_options = city_dropdown.find_elements(By.CLASS_NAME, "item")
            except:
                commune_input = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[3]/div[1]/div/div[2]/div[1]/div/div[2]/input")
                commune_input.send_keys(city_parts[1])

                time.sleep(3)    

                city_dropdown = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[3]/div[1]/div/div[2]/div[2]/div[1]")
                city_options = city_dropdown.find_elements(By.CLASS_NAME, "item")

            city_found = False
            for option in city_options:
                if city_parts[1].lower() in option.text.lower():
                    option.click()
                    city_found = True
                    break
                elif city_parts[0].lower() in option.text.lower():
                    option.click()
                    city_found = True
                    break

        time.sleep(3)
        
        # Section input and dropdown selection
        section_empty_click = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[8]/div[2]/div/div/div[1]/ul/li/div[2]/div/div/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td/div/div[1]/div[2]")))
        section_empty_click.click()

        time.sleep(3)

        section_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[6]/div[1]/div/div[2]/div[1]/div/div[2]/input")))
        section_input.send_keys(section)

        time.sleep(3)
        
        section_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[6]/div[1]/div/div[2]/div[2]/div[1]")))
        section_options = section_dropdown.find_elements(By.CLASS_NAME, "item")
        
        section_found = False
        for option in section_options:
            if section.lower() == option.text.lower():
                option.click()
                section_found = True
                break

        time.sleep(3)
        
        # Parcel input and dropdown selection
        parcelle_empty_click = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[8]/div[2]/div/div/div[1]/ul/li/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div[1]/div[1]")))
        parcelle_empty_click.click()

        time.sleep(3)

        parcelle_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[7]/div[1]/div/div[2]/div[1]/div/div[2]/input")))
        parcelle_input.send_keys(parcel)

        time.sleep(3)
        
        parcelle_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[7]/div[1]/div/div[2]/div[2]/div[1]")))
        parcelle_options = parcelle_dropdown.find_elements(By.CLASS_NAME, "item")
        
        parcelle_found = False
        for option in parcelle_options:
            if parcel == option.text.lower():
                option.click()
                parcelle_found = True
                break

        # MANUAL CLICK
        time.sleep(18)
        
        # Wait for next button and PDF click
        next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#map_root > div.esriPopup.esriPopupVisible > div.esriPopupWrapper > div:nth-child(1) > div > div.titleButton.next")))
        next_button.click()
        
        time.sleep(2)

        get_pdf = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/div[2]/div/div[1]/div[2]/div[3]/table/tr[5]/td[2]/a")))
        get_pdf.click()

        time.sleep(30)

        return "Site 5 automation completed successfully."
    except Exception as e:
        return f"Error in automating Site 5: {e}"
    


def run_automation(city, pcode, dept_name, section, parcel, status_label):
    # Run automation for each site in sequence
    results = []
    for i, site_func in enumerate([automate_site_1, automate_site_2, automate_site_3, automate_site_4, automate_site_5]):
    # for i, site_func in enumerate([automate_site_5]):
        status_label.config(text=f"Running automation for Site {i+1}...")
        result = site_func(city, pcode, dept_name, section, parcel)
        results.append(f"Site {i+1}: {result}")
        status_label.config(text=f"Site {i+1}: {result}")
    status_label.config(text="Automation complete. Check logs.")
    print("\n".join(results))

def start_automation(city_var, pcode_var, dept_var, section_var, parcel_var, status_label):
    # Get input values from the GUI
    city = city_var.get()
    pcode = pcode_var.get()
    dept_name = dept_var.get()
    section = section_var.get()
    parcel = parcel_var.get()

    if not all([city, pcode, dept_name, section, parcel]):
        messagebox.showerror("Input Error", "All fields are required!")
        return

    # Run automation in a separate thread to keep the GUI responsive
    Thread(target=run_automation, args=(city, pcode, dept_name, section, parcel, status_label), daemon=True).start()

# GUI setup
root = tk.Tk()
root.title("Automation Bot")
root.geometry("500x400")

frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Input fields
city_var = tk.StringVar()
pcode_var = tk.StringVar()
dept_var = tk.StringVar()
section_var = tk.StringVar()
parcel_var = tk.StringVar()

ttk.Label(frame, text="City:").grid(row=0, column=0, sticky=tk.W)
ttk.Entry(frame, textvariable=city_var).grid(row=0, column=1, sticky=(tk.W, tk.E))

ttk.Label(frame, text="Postal Code:").grid(row=1, column=0, sticky=tk.W)
ttk.Entry(frame, textvariable=pcode_var).grid(row=1, column=1, sticky=(tk.W, tk.E))

ttk.Label(frame, text="Department Name:").grid(row=2, column=0, sticky=tk.W)
ttk.Entry(frame, textvariable=dept_var).grid(row=2, column=1, sticky=(tk.W, tk.E))

ttk.Label(frame, text="Section:").grid(row=3, column=0, sticky=tk.W)
ttk.Entry(frame, textvariable=section_var).grid(row=3, column=1, sticky=(tk.W, tk.E))

ttk.Label(frame, text="Parcel:").grid(row=4, column=0, sticky=tk.W)
ttk.Entry(frame, textvariable=parcel_var).grid(row=4, column=1, sticky=(tk.W, tk.E))

# Function to clear all input fields
def clear_fields():
    city_var.set("")
    pcode_var.set("")
    dept_var.set("")
    section_var.set("")
    parcel_var.set("")
    status_label.config(text="Status: Waiting")  # Reset status label

# Clear Fields button
ttk.Button(frame, text="Clear Fields", command=clear_fields).grid(row=5, column=0, columnspan=2, pady=10)


ttk.Button(frame, text="Start Automation", command=lambda: start_automation(
    city_var, pcode_var, dept_var, section_var, parcel_var, status_label
)).grid(row=7, column=0, columnspan=2, pady=10)


# Start button
status_label = ttk.Label(frame, text="Status: Waiting", relief="sunken")
status_label.grid(row=9, column=0, columnspan=2, sticky=(tk.W, tk.E))


root.mainloop()
