"""
Get wan ip VIETTEL modem
"""
import os
import base64
import zipfile
import shutil
import wget
import requests
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from util_git import force_push

LOCAL_HOST_ADMIN = "http://192.168.1.1/"
UPLOAD_REMOTE_URL = "git@bitbucket.org:anhdshehe/stocks.git"
USER_NAME = "YWRtaW4="
PASSWORD = "ODgxOTkwMkBBbmg="
DOWNLOADED_CHROME_DRIVER = 0
REMOTE_NAME = "origin"
REVISION = "master"

class wait_for_non_empty_text(object):
    def __init__(self, locator):
        self.locator = locator

    def __call__(self, driver):
        try:
            element_text = EC._find_element(driver, self.locator).text.strip()
            return element_text != ""
        except StaleElementReferenceException:
            return False


# Get the current date and time
def decode_base64(text: str):
    """
    Decode base 64 from string

    Args:
        text (str): String to decode

    Returns:
        decoded_message: result
    """
    base64_bytes = text.encode('ascii')
    message_bytes = base64.b64decode(base64_bytes)
    decoded_message = message_bytes.decode('ascii')

    return decoded_message


def encode_base64(text: str):
    """
    Encode base 64 from string

    Args:
        text (str): String to encode

    Returns:
        encoded_message: result
    """
    message_bytes = text.encode('ascii')
    base64_bytes = base64.b64encode(message_bytes)
    encoded_message = base64_bytes.decode('ascii')
    return encoded_message


def download_chrome_driver_latest(DOWNLOADED_CHROME_DRIVER: int):
    """
    Download latest chrome driver

    Args:
        DOWNLOADED_CHROME_DRIVER (int): _description_

    Returns:
        _type_: _description_
    """

    # get the latest chrome driver version number
    url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE'
    response = requests.get(url)
    version_number = response.text

    # build the download url
    download_url = "https://chromedriver.storage.googleapis.com/" + version_number +"/chromedriver_win32.zip"

    # download the zip file using the url built above
    latest_driver_zip = wget.download(download_url,'chromedriver.zip')

    # extract the zip file
    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        zip_ref.extractall() # you can specify the destination folder path here
    # delete the zip file downloaded above
    os.remove(latest_driver_zip)
    # Move driver to setup folder
    shutil.copy('chromedriver.exe', 'D:/Setup')
    os.remove('chromedriver.exe')
    DOWNLOADED_CHROME_DRIVER += 1

    return DOWNLOADED_CHROME_DRIVER

def get_wan_ip():
    """
    Get wan ip from VIETTEL modem
    """
    counter_time_dowload = 0
    try:
        # cd to file's directory
        file_dir = os.path.dirname(__file__)
        os.chdir(file_dir)
        print("Current dir: {}".format(os.getcwd()))
        print("Getting WAN IP from local host...")
        # Get URL source
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-ssl-errors=yes')
        options.add_argument('--ignore-certificate-errors')
        driver = webdriver.Chrome(options=options)
        driver.get(LOCAL_HOST_ADMIN)
        # Wait until web load successfully
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "username")))
        # Fill username and password
        driver.find_element_by_id('username').send_keys(decode_base64(USER_NAME))
        driver.find_element_by_id('password').send_keys(decode_base64(PASSWORD))
        driver.find_element_by_id('login-button').click()
        # Wait until IP is valid
        WebDriverWait(driver, 20, 1).until(wait_for_non_empty_text((By.ID, "ppp_ip4")))
        # Get IP WAN
        new_ip_wan = str(driver.find_element_by_id('ppp_ip4').text)
        report_changed = False
        # Save to html file
        report_name = os.path.join(os.path.dirname(__file__), 'current_ip_wan.txt')
        with open(report_name, "r", encoding="utf-8") as f:
            old_ip = f.read().strip()
            if not old_ip == new_ip_wan:
                report_changed = True
        driver.close()

        # Push file to git
        if report_changed:
            with open(report_name, "w+", encoding="utf-8") as f:
                print(f"Save IP WAN to {report_name} file")
                f.write(new_ip_wan)
                print(f"Generate successfully file {report_name}")
            force_push(REMOTE_NAME, REVISION, [report_name])

    except Exception as ex:
        print(str(ex))
        if 'Current browser version is' in str(ex):
            counter_time_dowload = download_chrome_driver_latest(counter_time_dowload)
            if counter_time_dowload < 2:
                get_wan_ip()


if __name__ == '__main__':
    get_wan_ip()
    print("Finish, have a nice day!")