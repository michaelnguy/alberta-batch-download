# Required Modules
import datetime, time, wget, os, errno, ctypes
from sys import exit

# Selenium for web automation and data extraction
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options  
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

# Url to be scraped
url = "https://maps.alberta.ca/FWIMT_Pub/Viewer/?TermsOfUseRequired=true&Viewer=FWIMT_Pub"
chrome_options = Options()  
driver = None

SHORT_TIMEOUT = 30

try:
    driver = webdriver.Chrome(options=chrome_options)
except:
    ctypes.windll.user32.MessageBoxW(None, "Please download Chromedriver and put at the root of this program.", "No Chrome Driver Found.", 0)
    exit()

# Click Export Button
def click_export_button(export_button_title):
    # Wait on export button to load
    WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.presence_of_element_located((By.XPATH, '//*[@title="{}"]'.format(export_button_title))))
    export_button = driver.find_element_by_xpath('//*[@title="{}"]'.format(export_button_title))
    # Wait on export button to be clickable
    WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.element_to_be_clickable((By.XPATH, '//*[@title="{}"]'.format(export_button_title))))
    # Click the button once clickable
    export_button.click()

# Click Search Area By Name & Code
def click_search_area():
    # Wait on select area option to load
    WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.form-item.required')))
    select_by_name = driver.find_elements_by_css_selector('div.form-item.required')
    # Wait on select area option to be clickable
    WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.form-item.required')))
    # Click the option once clickable
    select_by_name[1].click()
    
def click_next_button():
    # Wait on next button to load
    WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.presence_of_element_located((By.XPATH, '//button[contains(text(), "Next")]')))
    # Wait on next button to be clickable
    next_button = WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Next")]')))
    # Click next button once clickable
    next_button.click()

# Select which area to get file from
def select_area(option_text):
    area_list_select_title = 'Choose HUC8 By Name'
    WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.presence_of_element_located((By.XPATH, '//*[@title="{}"]'.format(area_list_select_title))))
    area_list_select = Select(WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.element_to_be_clickable((By.XPATH, '//*[@title="{}"]'.format(area_list_select_title)))))
    area_list_select.select_by_visible_text(option_text)

# Select which area to get file from
def select_output_options():
    output_options = ['CheckBox4', 'CheckBox1', 'CheckBox2', 'CheckBox3', 'CheckBox5']
    main_div = WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.module.workflow-form.active')))
    for option in output_options:
        driver.find_element_by_xpath('//input[@data-gcx-form-item="{}"]'.format(option)).click()
    export_button = main_div.find_element_by_xpath('//button[text()="Export"]')
    export_button.click()

# Select which area to get file from
def click_download(today_folder, option):
    got_file = False
    try:
        # Wait on download link to appear
        download_div = WebDriverWait(driver, SHORT_TIMEOUT * 10
                    ).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.module.display-hyperlink.active')))

        # Download if extraction successful
        success_message = WebDriverWait(driver, SHORT_TIMEOUT
                        ).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'p.form-item.form-desc')))
        # Extract download url link
        if success_message and 'The process has ended successfully. Click on the link below to download the export file.' in success_message.text:
            download_link = download_div.find_element_by_tag_name('a').get_attribute('href')
            wget.download(download_link, './{}/{}.zip'.format(today_folder, option))
            got_file = True
    except:
        pass
    return got_file

# Load the area options
def load_area_options():
    area_options = Select(WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.presence_of_element_located((By.XPATH, '//*[@title="Choose HUC8 By Name"]'))))
    return [option.text for option in area_options.options]

if __name__ == "__main__":
    # Getting today's date for folder
    today = datetime.date.today()
    today_folder = today.strftime("%B %d, %Y")

    # Make directory for files
    try:
        os.makedirs('./{}'.format(today_folder))
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise

    # Load website
    driver.get(url)
    
    # Agree to Terms & Conditions
    check_agree = driver.find_element_by_id('termsOfUseAgreedBox')
    agree_button = driver.find_element_by_id('btnAccept')
    check_agree.click()
    agree_button.click()

    export_button_title = "Export Fish Inventories to SHP or CSV file"
    
    print('Begin navigation...')
    try:
        click_export_button(export_button_title)
        click_search_area()
        click_next_button()
        area_options = load_area_options()
        print('Loading options...')
        # area_options = [area_string for area_string in area_options if area_string > 'YATES RIVER']
        driver.refresh()
        
        # Begin to iterate through each option
        for option in area_options:
            got_file = False
            print()
            print('Extracting \'{}\'...'.format(option))
            while not got_file:
                try:
                    click_export_button(export_button_title)
                    click_search_area()
                    time.sleep(1)
                    click_next_button()
                    select_area(option)
                    time.sleep(1)
                    click_next_button()
                    select_output_options()
                    time.sleep(1)
                    got_file = click_download(today_folder, option)
                except:
                    pass
        print()
        print('Completed on {}.'.format(today_folder))
        ctypes.windll.user32.MessageBoxW(None, "Your file(s) have finished downloading.", "Download Complete", 0)
        driver.quit()
        exit()
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(None, '{}'.format(e), "Internal Error", 0)
        driver.quit()
        exit()
        pass
