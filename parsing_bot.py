import os
import os.path
import time
from typing import Optional, Tuple, List, Any, Dict

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import openpyxl


chrome_driver_path = "/home/oleksandr/Python_projects/test_works/future-proof-technology/chromedriver"


def set_up():
    # function to take care of downloading file
    def enable_download_headless(driver, download_dir):
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_dir}}
        driver.execute("send_command", params)

    # instantiate a chrome options object, so you can set the size and headless preference
    # some of these chrome options might be unnecessary, but I just used a boilerplate
    # change the <path_to_download_default_directory> to whatever your default download folder is located
    chrome_options = Options()
    # chrome_options.add_argument("--headless")
    # chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--verbose')
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": "./output",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False
    })
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--disable-software-rasterizer')

    # initialize driver object and change the <path_to_chrome_driver>
    # depending on your directory where your chromedriver should be
    driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=chrome_driver_path)

    # change the <path_to_place_downloaded_file> to your directory where you would like to place the downloaded file
    download_dir = "./output"

    # function to handle setting up headless download
    enable_download_headless(driver, download_dir)
    print('Browser driver created.')
    return driver


def click_button(driver, xpath: str, time_slip=0) -> None:
    time.sleep(time_slip)
    try:
        driver.find_element(By.XPATH, xpath).click()
        print('The button was pressed.')
    except Exception as e:
        print(f'{e} - Button not founded')


def get_departments_amounts(driver, dep_xpath: str, amo_xpath: str) -> list:
    time.sleep(3)
    rows_list = []
    dep_list = []
    amo_list = []
    try:
        dep_rows = len(driver.find_elements(By.XPATH, dep_xpath))
        amo_rows = len(driver.find_elements(By.XPATH, amo_xpath))
        print(f'Found {dep_rows} agencies.')
        print(f'Found {amo_rows} amounts')
        departments = driver.find_elements(By.XPATH, dep_xpath)
        amounts = driver.find_elements(By.XPATH, amo_xpath)
        for department in departments:
            dep_list.append(department.text)
        for amount in amounts:
            amo_list.append(amount.text)
        for row in range(0, dep_rows):
            rows_list.append([dep_list[row], amo_list[row]])
        print('Got a list of agencies and the amount.')
        return rows_list
    except Exception as e:
        print(f'{e} - Elements not founded.')


def open_agency_page(driver, agency_name: str, xpath: str) -> None:
    try:
        agencies = driver.find_elements(By.XPATH, xpath)
        for agency in agencies:
            if agency_name in agency.text:
                print(f'Found {agency_name}.')
                link = agency.get_attribute('href')
                print(link)
                driver.get(link)
                print('Agency page opened.')
                break
    except Exception as e:
        print(f'{e} - Elements not founded')


def scrap_table(driver, rows_xpath: str, cols_xpath: str) -> Tuple[List[List[Any]], List[Any], List[Dict[Any, Any]]]:
    time.sleep(10)
    select_xpath = '//*[@id="investments-table-object_length"]/label/select'
    rows_list = []
    links_list = []
    list_for_check = []
    try:
        print('Started scraping the table.')
        select = Select(driver.find_element(By.XPATH, select_xpath))
        # select by visible text
        select.select_by_visible_text('All')
        print('Select "All" selected.')
        print('Waiting for the table to load...')
        time.sleep(20)
        rows = len(driver.find_elements(By.XPATH, rows_xpath))
        cols = len(driver.find_elements(By.XPATH, cols_xpath))
        print(f'Found {rows} lines.')
        print(f'Found {cols} columns')
        for row in range(1, rows + 1):
            row_list = []
            for col in range(1, cols + 1):
                value = driver.find_element(By.XPATH,
                                            '//*[@id="investments-table-object"]/tbody/tr[' + str(row) + ']/td[' + str(
                                                col) + ']').text
                row_list.append(value)
                if col == 1:
                    dict_for_check = {}
                    try:
                        a_xpath = '//*[@id="investments-table-object"]/tbody/tr[' + str(row) + ']/td[1]/a'
                        tag_a = driver.find_element(By.XPATH, a_xpath)
                        link = tag_a.get_attribute('href')
                        links_list.append(link)
                        print(f'Added link to list: {link}')
                        uii_for_check = tag_a.text
                        investment_title = driver.find_element(By.XPATH,
                                            '//*[@id="investments-table-object"]/tbody/tr[' + str(row) + ']/td[3]').text
                        dict_for_check[investment_title] = uii_for_check
                    except Exception as e:
                        continue
                    finally:
                        if dict_for_check != {}:
                            list_for_check.append(dict_for_check)

            rows_list.append(row_list)
        print('Scraping the table finished.')
        return rows_list, links_list, list_for_check
    except Exception as e:
        print(f'{e} - Elements not founded')


def download_file(driver, links: list) -> None:
    if len(links) > 0:
        print('Start downloading files.')
        for link in links:
            driver.get(link)
            download_xpath = '//*[@id="business-case-pdf"]/a'
            time.sleep(5)
            try:
                driver.find_element(By.XPATH, download_xpath).click()
                time.sleep(10)
            except Exception as e:
                print(f'{e} - Button not founded')
    files = [f for f in os.listdir('./output') if os.path.isfile(f) and f.endswith('.pdf')]
    while len(files) != len(links):
        pass
    else:
        print('Files upload complete.')
        driver.close()


def save_to_xlsx(data: list, file_name: str, sheet_name: str) -> None:
    print('Recording to file started.')
    if os.path.exists(file_name):
        # to load the workbook with file name
        workbook = openpyxl.load_workbook(file_name)
    else:
        workbook = openpyxl.Workbook()
    if sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
    else:
        worksheet = workbook.create_sheet(sheet_name)
    rows = len(data)
    cols = len(data[0])
    for row in range(0, rows):
        for col in range(0, cols):
            value = data[row][col]
            worksheet.cell(row=row + 1, column=col + 1).value = value
            workbook.save(file_name)
    print('Recording to file finished.')

