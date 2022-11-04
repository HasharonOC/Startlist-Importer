from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select
import pandas as pd
from csv import reader
import time
import os
import dotenv
import logging
import sys


# Compute diff between two lists of entries list
def diff(new_file, old_file):
    new_entries_list = [item for item in new_file if item not in old_file]
    deleted_entries_list = [item for item in old_file if item not in new_file]
    return [new_entries_list, deleted_entries_list]


def get_new_registered_competitors(start_list_file):
    new_entries = []
    competitors_list = read_xlsx_into_list(start_list_file)
    on_site_registered_list = [ele for ele in competitors_list if ele[9] == "[הרשמה מהירה]"]
    dotenv.load_dotenv(dotenv.find_dotenv())
    mulka_on_site_registered_list = []
    if os.getenv('REGISTERED_ON_SITE') != "":
        mulka_on_site_registered_list = os.getenv('REGISTERED_ON_SITE').split(" ")
    for entry in on_site_registered_list:
        if str(entry[0]) not in mulka_on_site_registered_list:
            new_entries.append(entry)
    return new_entries


# Read csv file
def read_csv_into_list(file):
    if os.path.exists(file):
        with open(file, encoding="utf8") as csv_file:
            # Read csv file
            csv_reader = reader(csv_file)
            # Passing the csv_reader object to list() to get a list of lists
            return list(csv_reader)[1:]
    else:
        logging.info(file + " was not found")
        return []


# Read xlsx file
def read_xlsx_into_list(file):
    if os.path.exists(file):
        with open(file, 'rb') as excel_file:
            df = pd.read_excel(excel_file, dtype=str)
            return df.values.tolist()
    else:
        logging.info(file + " was not found")
        return []


# Read initial guest start number from .env file
def get_initial_guest_start_number(dotenv_file):
    dotenv.load_dotenv(dotenv_file)
    if os.getenv('GUEST_START_NUMBER') is not None:
        guest_start_number = int(os.getenv('GUEST_START_NUMBER'))
    else:
        guest_start_number = 15000
    return guest_start_number


# Get competitor start number
def get_start_number(start_number):
    if int(start_number) < 15000:
        return start_number
    else:
        global GUEST_START_NUMBER
        start_number = GUEST_START_NUMBER
        GUEST_START_NUMBER += 1
        return start_number


# Login to mulka cloud
def mulka_cloud_login(selenium_driver, mulka_url,  uid, comp_password):
    selenium_driver.switch_to.window("mulka_tab")
    selenium_driver.get(mulka_url + "/cloud/index.jsp")
    user_id = selenium_driver.find_element_by_id("txtId")
    user_id.send_keys(uid)
    competition_password = selenium_driver.find_element_by_id("txtPassword")
    competition_password.send_keys(comp_password)
    personal_name = selenium_driver.find_element_by_id("txtPersonName")
    personal_name.send_keys("Registration Script")
    login_button = selenium_driver.find_element_by_id("btnLogin")
    login_button.click()
    WebDriverWait(selenium_driver, CONNECTION_TIMEOUT).until(
        ec.presence_of_element_located((By.ID, "btnNavbarChat"))
    )


# Register a competitor through mulka cloud
def register_competitor(selenium_driver, mulka_url, competitor_details):
    selenium_driver.switch_to.window("mulka_tab")
    selenium_driver.get(mulka_url + "/cloud/tool/direct-entry.jsp?startNumber=" +
                        str(get_start_number(competitor_details[0])))
    WebDriverWait(selenium_driver, CONNECTION_TIMEOUT).until(
        ec.presence_of_element_located((By.ID, "btnNavbarChat"))
    )
    card_number = selenium_driver.find_element_by_id("txtCardNumber")
    if card_number.get_attribute('value') == "":
        if str(competitor_details[6]) != "" and str(competitor_details[6]) != "nan":
            card_number.send_keys(str(competitor_details[6]))
    runner_name = selenium_driver.find_element_by_id("txtRunnerName1")
    if runner_name.get_attribute('value') == "":
        if str(competitor_details[1]) != "" and str(competitor_details[1]) != "nan":
            runner_name.send_keys(str(competitor_details[1]))
    club_name = selenium_driver.find_element_by_id("txtClubName1")
    if club_name.get_attribute('value') == "":
        if str(competitor_details[2]) != "" and str(competitor_details[2]) != "nan":
            club_name.send_keys(str(competitor_details[2]))
    course_class = Select(selenium_driver.find_element_by_id('selClass'))
    course_class.select_by_visible_text(competitor_details[3])
    ok_button = selenium_driver.find_element_by_id("btnOK")
    ok_button.click()
    WebDriverWait(selenium_driver, CONNECTION_TIMEOUT).until(
        ec.presence_of_element_located((By.ID, "txtStartNumber"))
    )


# Login to isoa website
def isoa_login(selenium_driver, isoa_url, user, passwd):
    selenium_driver.get(isoa_url)

    username = selenium_driver.find_element_by_id("ctl00_LoginView1_Login1_UserName")
    username.send_keys(user)
    password = selenium_driver.find_element_by_id("ctl00_LoginView1_Login1_Password")
    password.send_keys(passwd)
    password.send_keys(Keys.RETURN)

    WebDriverWait(selenium_driver, CONNECTION_TIMEOUT).until(
        ec.presence_of_element_located((By.ID, "ctl00_LoginView1_HeadLoginStatus"))
    )
    selenium_driver.get(isoa_url + "/admin/AdminEventList.aspx")


# Download new start lists from ISOA website
def download_new_start_lists(selenium_driver, competition_name):
    # Remove previous start list file
    if os.path.exists(TARGET_DOWNLOAD_DIRECTORY_PATH + '/StartList.xlsx'):
        os.remove(TARGET_DOWNLOAD_DIRECTORY_PATH + '/StartList.xlsx')
    # Switch to isoa tab
    selenium_driver.switch_to.window(selenium_driver.window_handles[0])
    selenium_driver.refresh()
    hide_old_button = WebDriverWait(selenium_driver, CONNECTION_TIMEOUT).until(
        ec.presence_of_element_located((By.ID, "_rfdSkinnedctl00_ContentPlaceHolder1_cbHideOld"))
    )
    hide_old_button.click()
    WebDriverWait(selenium_driver, CONNECTION_TIMEOUT).until(
        ec.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_rgEvents_ctl00__0"))
    )
    time.sleep(10)
    competitions_table = selenium_driver.find_element_by_id("ctl00_ContentPlaceHolder1_rgEvents_ctl00")
    for row in competitions_table.find_elements_by_css_selector('tr'):
        for cell in row.find_elements_by_tag_name('td'):
            if cell.text == competition_name:
                cell.click()
    WebDriverWait(selenium_driver, CONNECTION_TIMEOUT).until(
        ec.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_rdbExport2Xls"))
    )

    excel_export_button = selenium_driver.find_element_by_id("ctl00_ContentPlaceHolder1_rdbExport2Xls")
    excel_export_button.click()
    time.sleep(5)


def start_list_importer(selenium_driver, competition_name):
    download_new_start_lists(selenium_driver, competition_name)
    logging.info("Downloaded new start lists successfully")
    # Get diff list
    new_entries = get_new_registered_competitors(TARGET_DOWNLOAD_DIRECTORY_PATH + '/StartList.xlsx')

    if len(new_entries) > 0:
        for entry in new_entries:
            try:
                register_competitor(selenium_driver, MULKA_CLOUD_URL, entry)
            except:
                # DELETE FROM REGISTERED LIST
                logging.exception("Unable to register " + str(entry[0]) + " - Exiting")
                driver.quit()
                sys.exit(1)
            else:
                if dotenv.get_key(dotenv_file, 'REGISTERED_ON_SITE') == "":
                    dotenv.set_key(dotenv_file, 'REGISTERED_ON_SITE', dotenv.get_key(dotenv_file, 'REGISTERED_ON_SITE')
                                   + str(entry[0]))
                else:
                    dotenv.set_key(dotenv_file, 'REGISTERED_ON_SITE',
                                   dotenv.get_key(dotenv_file, 'REGISTERED_ON_SITE') + " " + str(entry[0]))
                logging.info("Registered " + str(entry[0]) + " successfully")
    else:
        logging.info("no new entries")

    dotenv.set_key(dotenv_file, 'GUEST_START_NUMBER', str(GUEST_START_NUMBER))


TARGET_DOWNLOAD_DIRECTORY_PATH = 'C:\\temp\\Orienteering'

# TARGET_DOWNLOAD_DIRECTORY_PATH = '/tmp/Orienteering'

GECKO_DRIVER_PATH = "C:/Program Files (x86)/geckodriver.exe"

# CHROME_DRIVER_PATH = "C:/Program Files (x86)/chromedriver.exe"

dotenv_file = dotenv.find_dotenv()

GUEST_START_NUMBER = get_initial_guest_start_number(dotenv_file)

# MULKA_CLOUD_URL = "http://test.mulka2.com"

MULKA_CLOUD_URL = "http://jp.mulka2.com"

MULKA_EVENT_ID = "153160"

MULKA_EVENT_PASSWORD = "450"

ISOA_WEBSITE_URL = "https://nivut.org.il"

COMPETITION_NAME = "יער חורשים"

CONNECTION_TIMEOUT = 45

TIME_BETWEEN_IMPORTS = 30

profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', TARGET_DOWNLOAD_DIRECTORY_PATH)
profile.set_preference('browser.helperApps.neverAsk.saveToDisk',
                       'text/csv; application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


options = FirefoxOptions()
options.add_argument("--headless")

# chrome_prefs = {"download.default_directory": TARGET_DOWNLOAD_DIRECTORY_PATH}
#
# chrome_options = webdriver.ChromeOptions()
# chrome_options.add_experimental_option("prefs", chrome_prefs)
# chrome_options.binary_location = os.environ.get("GOOGLE_CHROME_BIN")
# chrome_options.add_argument("--headless")
# chrome_options.add_argument("--no-sandbox")
# chrome_options.add_argument("--disable-dev-shm-usage")

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')

driver = webdriver.Firefox(firefox_profile=profile, executable_path=GECKO_DRIVER_PATH, options=options)
# driver = webdriver.Chrome(executable_path=os.environ.get("CHROMEDRIVER_PATH"), options=chrome_options)


try:
    logging.info("Started Logging in to ISOA")
    isoa_login(driver, ISOA_WEBSITE_URL, "4211", "point83")
    logging.info("Logged in to ISOA successfully")
except:
    logging.exception("Unable to login to ISOA server")
    driver.quit()
else:
    driver.execute_script("window.open('about:blank', 'mulka_tab');")
    try:
        mulka_cloud_login(driver, MULKA_CLOUD_URL, MULKA_EVENT_ID, MULKA_EVENT_PASSWORD)
        logging.info("Logged in to Mulka cloud successfully")
    except:
        logging.exception("Unable to login to Mulka cloud")
        driver.quit()
    else:
        while True:
            logging.info("Import Started")
            start_list_importer(driver, COMPETITION_NAME)
            logging.info("Import Ended")
            time.sleep(TIME_BETWEEN_IMPORTS)
