import os
import os.path
import time
from typing import Tuple, List, Any, Dict

import PyPDF4
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

from main import output_folder

chrome_driver_path = "driver/chromedriver"


def set_up(dir: str):
    # function to take care of downloading file
    def enable_download_headless(driver, download_dir):
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_dir}}
        driver.execute("send_command", params)

    # instantiate a chrome options object, so you can set the size and headless preference
    # some of these chrome options might be unnecessary, but I just used a boilerplate
    # change the <path_to_download_default_directory> to whatever your default download folder is located
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--verbose')
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": "",
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
    # download_dir = "output"
    download_dir = dir

    # function to handle setting up headless download
    enable_download_headless(driver, download_dir)
    print('Browser driver created.')
    return driver


def click_button(driver, xpath: str, time_slip=0) -> None:
    """
    Click "DIVE IN" on the homepage to reveal the spend amounts for each agency
    :param driver: Selenium driver for browser
    :param xpath: Xpath of element button. Example: '//*[@id="node-23"]/div/div/div/div/div/div/div/a'
    :param time_slip: Delay before clicking the button
    :return: None
    """

    time.sleep(time_slip)
    try:
        driver.find_element(By.XPATH, xpath).click()
        print('The button was pressed.')
    except Exception as e:
        print(f'{e}')


def get_departments_amounts(driver, dep_xpath: str, amo_xpath: str) -> list:
    """
    Scrapping spend amounts for each agency from a page.
    :param driver: Selenium driver for browser
    :param dep_xpath: Xpath to department's blocks
    :param amo_xpath: Xpath to amount's blocks
    :return: List of departments and amounts.
    """

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
        print(f'{e}')


def open_agency_page(driver, agency_name: str, xpath: str) -> None:
    """
    Open agency's page
    :param driver: Selenium driver for browser
    :param agency_name: name of agency
    :param xpath: xpath to agency's link
    :return: agency's page
    """
    try:
        agencies = driver.find_elements(By.XPATH, xpath)
        for agency in agencies:
            if agency_name in agency.text:
                print(f'Found {agency_name}.')
                link = agency.get_attribute('href')
                print(f'{agency_name}: {link}')
                driver.get(link)
                print('Agency page opened.')
                break
    except Exception as e:
        print(f'{e}')


def scrap_table(driver, rows_xpath: str, cols_xpath: str) -> Tuple[List[List[Any]], List[Any], List[Dict[str, Any]]]:
    """
    Scrapping data from table
    :param driver: Selenium driver for browser
    :param rows_xpath: xpath of rows
    :param cols_xpath: xpath of cols
    :return: tuple of:
    1. a list of lines to write to the xlsx file
    2. a list of download links
    3. a list to check the comparison between links and files
    """
    global dict_for_check
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
                                                               '//*[@id="investments-table-object"]/tbody/tr[' + str(
                                                                   row) + ']/td[3]').text
                        dict_for_check['investment'] = investment_title
                        dict_for_check['uii'] = uii_for_check
                    except Exception as e:
                        continue
                    finally:
                        if dict_for_check != {}:
                            list_for_check.append(dict_for_check)

            rows_list.append(row_list)
        print('Scraping the table finished.')
        return rows_list, links_list, list_for_check
    except Exception as e:
        print(f'{e}')


def latest_download_file(path=output_folder):
    """
    Get the last modified file in the folder
    :param path: path to folder
    :return: name of file
    """
    try:
        os.chdir(path)
        files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
        newest = files[-1]
        return newest
    except Exception as e:
        print(e)
    finally:
        os.chdir('..')


def download_file(driver, links: list, list_for_check: list) -> None:
    """
    Download files from list
    :param driver: Selenium driver for browser
    :param links: list of links
    :param list_for_check: list for check download
    :return:
    """
    try:
        downloaded_files = []
        if len(links) > 0:
            print('Start downloading files.')
            for link in links:
                driver.get(link)
                download_xpath = '//*[@id="business-case-pdf"]/a'
                time.sleep(5)
                try:
                    driver.find_element(By.XPATH, download_xpath).click()
                    time.sleep(10)
                    file_ends = "crdownload"
                    while "crdownload" == file_ends:
                        time.sleep(1)
                        newest_file: str = latest_download_file()
                        if "crdownload" in newest_file:
                            file_ends = "crdownload"
                        else:
                            file_ends = "none"
                    downloaded_files.append(newest_file)
                    print(f'File {newest_file} downloaded successfully.')
                    compare_data(get_data_from_file(newest_file, 0), list_for_check)
                except Exception as e:
                    print(f'{e}')
        if len(downloaded_files) == len(links):
            driver.close()
            print('File downloads complete.')
    except Exception as e:
        print(e)


def get_data_from_file(file: str, num_page: int, path=output_folder) -> Dict[str, Any]:
    """
    Get some data from file
    :param path: default "output"
    :param num_page: number of page (start from 0)
    :param file: file name
    :return: list of dictionary with some data from files
    """
    text_list = []
    file_dict = {}

    # cur_path = os.path.dirname(__file__)
    # new_path = os.path.relpath(file, cur_path)
    try:
        os.chdir(path)
        file_obj = open(file, 'rb')
        file_reader = PyPDF4.PdfFileReader(file_obj)
        page = file_reader.getPage(num_page)
        pages_text_list = page.extractText().replace(' \n', '').split("\n")
        for i in range(len(pages_text_list)):
            if 'Name of this Investment:' in pages_text_list[i]:
                file_dict['investment'] = pages_text_list[i + 1]
                file_dict['uii'] = pages_text_list[i + 3]
        text_list.append(file_dict)
        return file_dict
    except Exception as e:
        print(e)
    finally:
        os.chdir('..')


def compare_data(data_1: dict, data_2: list) -> bool:
    """
    Compare the downloaded file and data from the agency page
    :param data_1: data from download file
    :param data_2: data from site
    :return:
    """
    try:
        flag = False
        dict_from_file_uii = data_1["uii"]
        dict_2 = next(item for item in data_2 if item["uii"] == dict_from_file_uii)
        for j in data_1:
            if j not in dict_2:
                flag = False
                print('Link and download not matched.')
                break
            else:
                flag = True
        print('Link and download have been matched.')
        return flag
    except Exception as e:
        print(e)


def save_to_xlsx(data: list, file_name: str, sheet_name: str, path=output_folder) -> None:
    """
    Save data to file xlsx
    :param path: path to dir
    :param data: data to write to a file
    :param file_name: a file name
    :param sheet_name: a sheet name
    :return:
    """
    os.chdir(path)
    print('Recording to file started.')
    try:
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
    except Exception as e:
        print(e)
    finally:
        os.chdir('..')
