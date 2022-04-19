from subprocess import check_output
import os
import re
import requests
import wget
import zipfile
import shutil
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import io
from bs4 import BeautifulSoup
import xlsxwriter
import datetime
import time 
WEBSITE_URL = "https://iboard.ssi.com.vn/bang-gia/chung-quyen"
UPLOAD_REMOTE_URL = "git@bitbucket.org:anhdshehe/stocks.git"
NO_PUSH_COMMIT = False
DOWNLOADED_CHROME_DRIVER = 0
remote_name = "origin"
revision = "master"
# Get the current date and time
TODAY_TIME = datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")
REPORT_NAME = "status/"+"stock_status_" + TODAY_TIME + ".xlsx"
WAIT_WEBSITE_LOADING_BY_ID = 'table-body-scroll'
MAPPING_DATA_WITH_ID = {
    "Ma CKhoan":0,
    "CTPH":1,
    "Ngay GDCC":2,
    "Tran":3,
    "San":4,
    "Gia TC":5,
    "Gia Khop":12,
    "Tong KLuong":23,
    "Ma CKCS":24,
    "Gia CKCS":25,
    "Gia thuc hien":28,
    "Do lech":27,
    "TL CD":29,
    "Gia hoa von":26,
}

HEADER_NAMES = [
    "Ma CKhoan",
    "CTPH",
    "Ngay GDCC",
    "Tran",
    "San",
    "Gia TC",
    "Gia Khop",
    "Tong KLuong",
    "Ma CKCS",
    "Gia CKCS",
    "Gia thuc hien",
    "Do lech",
    "TL CD",
    "Gia hoa von",
    "Do lech %",
    "Days left",
    "Gain",
]

LIST_DATA_FLOAT = [
    "Tran",
    "San",
    "Gia TC",
    "Gia Khop",
    "Tong KLuong",
    "Gia CKCS",
    "Gia thuc hien",
    "Do lech",
    "TL CD",
    "Do lech %",
    "Gia hoa von",
    "Gain",
]

LIST_SAME_FONT_COLOR = [
    "Ma CKhoan",
    "Gia Khop",
    "Tong KLuong",
]

# Format excel
FORMAT_NUMBER = '#,##0.00'
FORMAT_NUMBER_PECENTAGE = '#,##0.00%'

FORMAT_HIGHEST_PRICE = { 'font_color' : 'pink' }
FORMAT_LOWEST_PRICE = { 'font_color' : 'cyan' }

FORMAT_INCREASE_PRICE = { 'font_color' : 'green' }
FORMAT_DECREASE_PRICE = { 'font_color' : 'red' }

FORMAT_STATUS_OK = { 'bg_color' : 'green' }
FORMAT_STATUS_WARNING = { 'bg_color' : 'yellow' }
FORMAT_STATUS_DANGEROUS = { 'bg_color' : 'red' }

FORMAT_HEADER = {
    'font_size' : 12,
    'bold' : True,
    'bg_color' : '#8DB4E2',
    'bottom':1, 'top':1, 'left':1, 'right':1,
}

COLUMN_WIDTH = 15
HEADER_COLUM = 0
START_DATA_COLUM = 1

def main():
    try:
        # cd to file's directory
        file_dir = os.path.dirname(__file__)
        os.chdir(file_dir)
        print("Current dir: {}".format(os.getcwd()))
        # print("Git init")
        # git_log = check_output("git init", shell=True).decode()
        # git_log = check_output("git remote -v", shell=True).decode()
        # pattern = remote_name + r'\s+' + convert_text_to_regex(UPLOAD_REMOTE_URL) + r'.*\(fetch\)'
        # remote_exist = re.search(pattern, git_log, re.IGNORECASE)
        # if remote_exist is not None:
        #     print("Remote already exists")
        # else:
        #     pattern = remote_name + r'.*\(fetch\)'
        #     if re.search(pattern, git_log, re.IGNORECASE) is not None:
        #         print("Update URL {} to remote {}".format(UPLOAD_REMOTE_URL, remote_name,))
        #     else:
        #         git_log = check_output("git remote add {remote_name} {remote_url}".format(remote_name=remote_name, remote_url=UPLOAD_REMOTE_URL), shell=True).decode()

        # print("Clean changes")
        # git_log = check_output("git reset --hard && git clean -dfx", shell=True).decode()
        # print("Update all changes")
        # git_log = check_output("git fetch {remote_name}".format(remote_name=remote_name), shell=True).decode()
        # git_log = check_output("git fetch --all --tags", shell=True).decode()
        # git_log = check_output("git pull {remote_name} {revision}".format(remote_name=remote_name, revision=revision), shell=True).decode()
        print("Getting data from SSI website...")
        # Get URL source
        driver = webdriver.Chrome()
        driver.get(WEBSITE_URL)
        # Wait until web load successfully
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "table-body-scroll")))
        time.sleep(2)
        # Save to html file
        print("Save to html file")
        with io.open("page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        driver.close()

        # Init workbook
        workbook = xlsxwriter.Workbook(REPORT_NAME)
        worksheet = workbook.add_worksheet()
        # Write header contents
        row = HEADER_COLUM
        col = 0
        format_dict = FORMAT_HEADER
        cell_format = workbook.add_format(format_dict)
        for idx, key in enumerate(HEADER_NAMES):
            worksheet.write(row, col, key, cell_format)
            col += 1
            # Autofit cell width
            width = (len(key) + 2)
            if key == "Gain":
                width += 2
            worksheet.set_column(idx, idx, width)
        # Autofilter headers
        worksheet.autofilter(0, 0, 0, len(HEADER_NAMES) - 1)

        # Read data from website
        with io.open("page_source.html", "r", encoding="utf-8") as f:
            content = f.read()
        # Parsing data to beautifulSoup object
        print("Parsing data")
        soup = BeautifulSoup(content, "html.parser")
        # Find data table
        table = soup.find(id=WAIT_WEBSITE_LOADING_BY_ID)
        table_datas = table.find_all('tr')
        print("Save data to dictionary")
        data_dict = dict()
        row = START_DATA_COLUM
        for table_data in table_datas:
            if len(table_data.contents) == 0:
                continue
            data_dict[table_data['id']] = dict()
            col = 0
            for key, idx in MAPPING_DATA_WITH_ID.items():
                value = table_data.contents[idx].string
                if key == "TL CD":
                    value = value.split(':', 1)[0]
                if value is None and key in LIST_DATA_FLOAT:
                    value = '0'
                if key in LIST_DATA_FLOAT:
                    value = float(value.replace(',', ''))
                data_dict[table_data['id']][key] = value
                col += 1
            # Calculate do lech
            data_dict[table_data["id"]]["Do lech"] = data_dict[table_data["id"]]["Gia hoa von"] - data_dict[table_data["id"]]["Gia CKCS"]
            # Calculate diff by percentage
            if data_dict[table_data["id"]]["Gia CKCS"] != 0:
                data_dict[table_data["id"]]["Do lech %"] = round((100 * data_dict[table_data["id"]]["Do lech"] / data_dict[table_data["id"]]["Gia CKCS"]), 2)
            else:
                data_dict[table_data["id"]]["Do lech %"] = 0

            # Calculate gain with percentage
            gia_chung_quyen = data_dict[table_data["id"]]["Gia Khop"] * data_dict[table_data["id"]]["TL CD"]
            if gia_chung_quyen != 0:
                data_dict[table_data["id"]]["Gain"] = data_dict[table_data["id"]]["Gia CKCS"] / gia_chung_quyen
            else:
                data_dict[table_data["id"]]["Gain"] = 0
            # Calculate days remaining
            today = datetime.date.today()
            day = int(data_dict[table_data["id"]]["Ngay GDCC"][:2])
            month = int(data_dict[table_data["id"]]["Ngay GDCC"][3:5])
            year = int("20" + data_dict[table_data["id"]]["Ngay GDCC"][6:])
            future = datetime.date(year, month, day)
            diff = future - today
            data_dict[table_data["id"]]["Days left"] = diff.days
            row += 1


        print("Write and format data")
        row = START_DATA_COLUM
        for stock in data_dict.values():
            if stock["Gia Khop"] >= stock["Gia TC"]:
                format_dict_same_color = FORMAT_INCREASE_PRICE
                if stock["Gia Khop"] == stock["Tran"]:
                    format_dict_same_color = FORMAT_HIGHEST_PRICE
            else:
                format_dict_same_color = FORMAT_DECREASE_PRICE
                if stock["Gia Khop"] == stock["San"]:
                    format_dict_same_color = FORMAT_LOWEST_PRICE

            for header, value in stock.items():
                cell_format = workbook.add_format()
                if header in LIST_SAME_FONT_COLOR:
                    cell_format = workbook.add_format(format_dict_same_color)

                if header == "Do lech %":
                    if value < 15:
                        cell_format = workbook.add_format(FORMAT_STATUS_OK)
                    elif value > 30:
                        cell_format = workbook.add_format(FORMAT_STATUS_DANGEROUS)
                    else:
                        cell_format = workbook.add_format(FORMAT_STATUS_WARNING)

                if header == "Days left":
                    if value > 60:
                        cell_format = workbook.add_format(FORMAT_STATUS_OK)
                    elif value < 15:
                        cell_format = workbook.add_format(FORMAT_STATUS_DANGEROUS)
                    else:
                        cell_format = workbook.add_format(FORMAT_STATUS_WARNING)

                if header == "Gain":
                    if value > 2:
                        cell_format = workbook.add_format(FORMAT_STATUS_OK)
                    elif value < 1:
                        cell_format = workbook.add_format(FORMAT_STATUS_DANGEROUS)
                    else:
                        cell_format = workbook.add_format(FORMAT_STATUS_WARNING)

                if header == "Tran":
                    cell_format = workbook.add_format(FORMAT_HIGHEST_PRICE)
                if header == "San":
                    cell_format = workbook.add_format(FORMAT_LOWEST_PRICE)

                if header in LIST_DATA_FLOAT:
                    if header == "Do lech %" or header == "Gain":
                        cell_format.set_num_format(FORMAT_NUMBER_PECENTAGE)
                        value /= 100
                    else:
                        cell_format.set_num_format(FORMAT_NUMBER)
                worksheet.write(row, HEADER_NAMES.index(header), value, cell_format)
            row += 1

        # Freeze the first row, column.
        worksheet.freeze_panes(1, 2)  
        workbook.close()
        print("Generate successfully file {}".format(REPORT_NAME))
        # print(table.prettify())
        # Push file to git
        if NO_PUSH_COMMIT is not True:
            check_output("git pull {remote_name} {revision}".format(remote_name=remote_name, revision=revision), shell=True).decode()
            check_output("git add {file_name}".format(file_name=REPORT_NAME), shell=True).decode()
            check_output("git commit -m \"Add {file_name}\"".format(file_name=REPORT_NAME), shell=True).decode()
            check_output("git push {remote_name} {revision}".format(remote_name=remote_name, revision=revision), shell=True).decode()

        print("Finish, have a nice day!")
    except Exception as Ex:
        print(str(Ex))
        if 'Current browser version is' in str(Ex):
            download_chrome_driver_latest()
            if DOWNLOADED_CHROME_DRIVER < 2:
                main()

def convert_text_to_regex(message):
    """
    This function is used to add "\" before regex meta character, and replace space character
    by '\s' or multiple spaces by '\s+'
    :param message: the message that needs to replace.
    Example: 
        Message:
            "This is    an [list]+ source *.c" 
        shall be converted to:
            "This\sis\s+an\s\[list\]\+\ssource\s\*\.c" 
    """
    message = message.strip()
    message = re.sub(r'([\{\}\[\]\(\)\^\$\.\|\*\+\?\\\/])', r'\\\1', message)
    message = re.sub(r'\s{2,}', r'\\s+', message)
    message = message.replace(' ', r'\s')
 
    return message


def download_chrome_driver_latest():

    # get the latest chrome driver version number
    url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE'
    response = requests.get(url)
    version_number = response.text

    # build the donwload url
    download_url = "https://chromedriver.storage.googleapis.com/" + version_number +"/chromedriver_win32.zip"

    # download the zip file using the url built above
    latest_driver_zip = wget.download(download_url,'chromedriver.zip')

    # extract the zip file
    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        zip_ref.extractall() # you can specify the destination folder path here
    # delete the zip file downloaded above
    os.remove(latest_driver_zip)
    # Move driver to setup folder
    shutil.copy('chromedriver.exe', 'D:\Setup')
    os.remove('chromedriver.exe')
    DOWNLOADED_CHROME_DRIVER += 1


if __name__ == '__main__':
    main()