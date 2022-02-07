import os
import shutil
from datetime import timedelta
from typing import Any, Dict

import pandas as pd
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables
from pandas import DataFrame
from robot.libraries.String import String
from selenium.webdriver.support.select import Select

pdf = PDF()
string = String()

excel = Files()
tables = Tables()
work_folder = os.getcwd()
output_folder = f'{work_folder}/output'
tmp_output_folder = f'{work_folder}/tmp'


def create_browser():
    """
    Create browser Chrome
    :return: browser
    """
    browser = Selenium()
    browser.set_download_directory(tmp_output_folder)
    browser.open_available_browser(maximized=True)
    return browser


browser_lib = create_browser()


def open_the_webpage(url):
    browser_lib.go_to(url)


def get_departments_amounts(dep_xpath: str, amo_xpath: str) -> list:
    """
    Scrapping spend amounts for each agency from a page.
    :param dep_xpath: Xpath to department's blocks
    :param amo_xpath: Xpath to amount's blocks
    :return: List of departments and amounts.
    """
    rows_list = []
    try:
        element_for_check = '//*[@id="agency-tiles-widget"]//span[@class=" h1 w900"]'
        browser_lib.wait_until_page_contains_element(element_for_check)
        departments = browser_lib.find_elements(dep_xpath)
        amounts = browser_lib.find_elements(amo_xpath)
        print(f'Found {len(departments)} agencies.')
        print(f'Found {len(amounts)} amounts')
        for i in range(0, len(departments)):
            rows_list.append([departments[i].text, amounts[i].text])
        print('Got a list of agencies and the amount.')
        return rows_list
    except Exception as e:
        print(f'{e}')


def save_to_xlsx(data, file_name: str, sheet_name: str, path=output_folder):
    """
    Save data to .xlsx file. Used pandas
    :param data: data to save (str, list or dict)
    :param file_name: file's name
    :param sheet_name: sheet's name
    :param path: path to file
    :return: created or updated .xlsx file
    """
    print('Recording to file started.')
    try:
        os.chdir(path)
        data_frame = pd.DataFrame(data)
        mode = "w"
        if os.path.exists(file_name):  # Check whether the existing file
            workbook = pd.ExcelFile(file_name)
            sheets_list = workbook.sheet_names
            if sheet_name not in sheets_list:  # Check whether the existing sheet
                mode = "a"
        with pd.ExcelWriter(file_name, mode=mode) as writer:
            data_frame.to_excel(writer, sheet_name=sheet_name)
        print('Recording to file finished.')
    except Exception as e:
        print(e)
    finally:
        os.chdir('..')


def scrap_table(agency_name: str, rows_xpath: str) -> DataFrame:
    """
    Scraping html-table and converting it to DataFrame.
    :param agency_name: agency's name for scrapping
    :param rows_xpath: xpath to agency's table
    :return: DataFrame from html-table
    """
    show_select_xpath = '//*[@id="investments-table-object_length"]/label/select'
    try:
        agency = browser_lib.find_element(f'//*[@id="agency-tiles-widget"]//span[text()="{agency_name}"]//parent::a')
        link = agency.get_attribute('href')
        print(f'{agency_name}: {link}')
        open_the_webpage(link)
        print('Agency page opened.')
        print('Waiting for the table to load...')
        browser_lib.wait_until_page_contains_element(rows_xpath, timedelta(seconds=20))
        print('Started scraping the table.')
        select = Select(browser_lib.find_element(show_select_xpath))
        select.select_by_visible_text('All')
        print('Select "All" selected.')
        print('Waiting for the entire table to load...')

        last_button_1 = browser_lib.find_element('//*[@id="investments-table-object_last"]')
        while True:  # Waiting for the entire table to load
            try:
                last_button_1.get_attribute('data-dt-idx')
                continue
            except Exception as e:
                break

        table_xpath = '//*[@id="investments-table-object"]'
        tables_frame = pd.read_html(browser_lib.find_element(table_xpath).get_attribute('outerHTML'))
        table = tables_frame[0]
        print('Scraping the table finished.')
        return table
    except Exception as e:
        print(f'{e}')


def find_links(rows_xpath: str) -> list:
    """
    Finds a links in the table at the specified xpath
    :param rows_xpath: xpath to the link
    :return: list of links and data for comparison
    """
    rows = len(browser_lib.find_elements(rows_xpath))
    num_tags_a = len(browser_lib.find_elements('//*[@id="investments-table-object"]//a'))
    print(f'Found {num_tags_a} links.')
    links_and_data_for_check_list = []
    processed_tag_a = 0
    try:
        for i in range(1, rows + 1):
            if processed_tag_a < num_tags_a:  # processing links
                link_and_data_for_check = []
                dict_for_check = {}
                a_xpath = '//*[@id="investments-table-object"]/tbody/tr[' + str(i) + ']/td[1]/a'
                tag_a = browser_lib.find_element(a_xpath)
                link = tag_a.get_attribute('href')
                print(f'Found link: {link}')
                uii_for_check = tag_a.text
                investment_title = browser_lib.find_element(
                    '//*[@id="investments-table-object"]/tbody/tr[' + str(i) + ']/td[3]').text
                dict_for_check['investment'] = investment_title
                dict_for_check['uii'] = uii_for_check
                link_and_data_for_check.append(link)
                link_and_data_for_check.append(dict_for_check)
                links_and_data_for_check_list.append(link_and_data_for_check)
                processed_tag_a += 1
        return links_and_data_for_check_list
    except Exception as e:
        print(e)


def wait_download_file(path: str = tmp_output_folder) -> str:
    """
    Waiting for file to load
    :param path: path to the folder for download
    :return: name (str) of downloaded file
    """
    try:
        os.chdir(path)
        print("Waiting for download file...")
        flag = True
        while flag:
            for file in os.listdir(os.getcwd()):
                if file.endswith(".pdf"):
                    downloaded_file = file
                    flag = False
        shutil.copy2(f"{tmp_output_folder}/{downloaded_file}", f"{output_folder}/{downloaded_file}")
        return downloaded_file
    except Exception as e:
        print(e)
    finally:
        os.chdir('..')
        shutil.rmtree(tmp_output_folder)


def download_file(link: str) -> str:
    """
    Download the file at the link
    :param link: link for download
    :param dict_for_check: data for compare from html-table
    :return: the name of the downloaded file
    """
    try:
        os.mkdir(tmp_output_folder)
        print('Starting to download the file.')
        browser_lib.execute_javascript("window.open('');")
        browser_lib_tabs = browser_lib.get_window_handles()
        browser_lib.switch_window(browser_lib_tabs[1])
        open_the_webpage(link)
        download_xpath = '//*[@id="business-case-pdf"]/a'
        browser_lib.click_element_when_visible(download_xpath)
        downloaded_file = wait_download_file()
        print(f'File {downloaded_file} downloaded successfully.')
        return downloaded_file
    except Exception as e:
        print(e)
    finally:
        browser_lib.close_window()
        browser_lib.switch_window(browser_lib_tabs[0])


def get_data_from_pdf_file(file: str, num_page: int, path=output_folder) -> Dict[str, Any]:
    """
    Get some data from file
    :param file: file name
    :param num_page: number of page (start from 1)
    :param path: default "output"
    :return: list of dictionary with some data from files
    """
    dict_for_check_from_pdf = {}
    try:
        print('Retrieving data from the downloaded file.')
        os.chdir(path)
        text = pdf.get_text_from_pdf(file, pages=num_page, trim=False)
        line_1 = string.get_lines_containing_string(text[num_page], '1. Name of this Investment:').split(': ')
        line_2 = string.get_lines_containing_string(text[num_page], '2. Unique Investment Identifier (UII):').split(': ')
        key_1 = line_1[0].replace("1. ", "")
        value_1 = line_1[1]
        dict_for_check_from_pdf[key_1] = value_1
        key_2 = line_2[0].replace("2. ", "")
        value_2 = line_2[1]
        dict_for_check_from_pdf[key_2] = value_2
        print('Retrieving data from the downloaded file completed.')
        return dict_for_check_from_pdf
    except Exception as e:
        print(e)
    finally:
        os.chdir('..')


def compare_data(data_1: dict, data_2: dict) -> bool:
    """
    Compare the downloaded file and data from the agency page
    :rtype: Union[None, bool]
    :param data_1: data from download file
    :param data_2: data from site
    :return:
    """
    try:
        print("Starting to compare data.")
        flag = False
        if data_1['Name of this Investment'] == data_2['investment'] and data_1['Unique Investment Identifier (UII)'] == \
                data_2['uii']:
            flag = True
            print('Link and download have been matched.')
        else:
            print('Link and download not matched.')
        return flag
    except Exception as e:
        print(e)
