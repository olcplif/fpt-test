import os
import time
from typing import Any, Dict

import openpyxl
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
output_folder = os.getcwd() + '/output'


def create_browser():
    """
    Create browser Chrome
    :return: browser
    """
    browser = Selenium()
    browser.set_download_directory(output_folder)
    browser.open_available_browser(maximized=True)
    return browser


browser_lib = create_browser()


def open_the_webpage(url):
    browser_lib.go_to(url)


def wait_until_load_element(xpath: str, driver=browser_lib, timeout: int = 300) -> None:
    """
    Waiting for the item to load on the page
    :param xpath: xpath to the element
    :param driver: browser
    :param timeout: waiting time
    :return: None
    """
    while timeout > 0:
        try:
            return driver.find_element(xpath)
        except:  # if element isn't already loaded or doesn't exist
            time.sleep(0.5)
            timeout -= 1
    raise RuntimeError(f"Element loading timeout")


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
        wait_until_load_element(element_for_check)
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
        if os.path.exists(file_name):  # Check whether the existing file
            workbook = openpyxl.load_workbook(file_name)
            writer = pd.ExcelWriter(file_name, engine='openpyxl', mode='w')
            writer.book = workbook
            writer.sheets = dict((ws.title, ws) for ws in workbook.worksheets)
        else:
            writer = pd.ExcelWriter(file_name, engine='openpyxl')
        data_frame.to_excel(writer, sheet_name=sheet_name)
        writer.save()
        writer.close()
        print('Recording to file finished.')
    except Exception as e:
        print(e)
    finally:
        os.chdir('..')


def scrap_table(file: str, agency_name: str, agency_xpath: str, rows_xpath: str) -> DataFrame:
    """
    Scraping html-table and converting it to DataFrame.
    :param file: file's name
    :param agency_name: agency's name for scrapping
    :param agency_xpath: xpath agency's block
    :param rows_xpath: xpath to agency's table
    :return: DataFrame from html-table
    """
    show_select_xpath = '//*[@id="investments-table-object_length"]/label/select'
    try:
        agencies = browser_lib.find_elements(agency_xpath)
        for agency in agencies:
            if agency_name in agency.text:
                print(f'Found {agency_name}.')
                link = agency.get_attribute('href')
                print(f'{agency_name}: {link}')
                open_the_webpage(link)
                print('Agency page opened.')
                break
        print('Waiting for the table to load...')
        while not browser_lib.is_element_visible(rows_xpath):
            pass
        else:
            print('Started scraping the table.')
            last_button_1 = browser_lib.find_element('//*[@id="investments-table-object_last"]')
            select = Select(browser_lib.find_element(show_select_xpath))
            select.select_by_visible_text('All')
            print('Select "All" selected.')
            print('Waiting for the entire table to load...')

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

            save_to_xlsx(table, file, agency_name)

            rows = len(browser_lib.find_elements(rows_xpath))
            tags_a = len(browser_lib.find_elements('//*[@id="investments-table-object"]//a'))
            print(f'Found {tags_a} links.')
            processed_tag_a = 0

            for row in range(1, rows + 1):  # processing links
                if processed_tag_a < tags_a:
                    dict_for_check = {}
                    try:
                        a_xpath = '//*[@id="investments-table-object"]/tbody/tr[' + str(row) + ']/td[1]/a'
                        tag_a = browser_lib.find_element(a_xpath)
                        link = tag_a.get_attribute('href')
                        print(f'Found link: {link}')
                        uii_for_check = tag_a.text
                        investment_title = browser_lib.find_element(
                            '//*[@id="investments-table-object"]/tbody/tr[' + str(row) + ']/td[3]').text
                        dict_for_check['investment'] = investment_title
                        dict_for_check['uii'] = uii_for_check
                        download_file(link, dict_for_check)
                    except:
                        continue
                    finally:
                        processed_tag_a += 1
                else:
                    break
        return table
    except Exception as e:
        print(f'{e}')


def latest_download_file(path: str = output_folder):
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


def wait_download_file(path: str = output_folder) -> str:
    """
    Waiting for file to load
    :param path: path to the folder for download
    :return: name (str) of downloaded file
    """
    try:
        os.chdir(path)
        print("Waiting for downloads", end="")
        num_files_start = len(os.listdir(os.getcwd()))
        while len(os.listdir(os.getcwd())) == num_files_start:  # waiting for a new file in the folder
            time.sleep(0.5)
            print(".", end="")

        file_end = "crdownload"
        while "crdownload" == file_end:  # waiting for the file to load completely
            time.sleep(0.5)
            files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
            newest_file = files[-1]
            print(".", end="")
            if "crdownload" in newest_file:
                file_end = "crdownload"
            else:
                file_end = "none"
        print("done!")
        return newest_file
    except Exception as e:
        print(e)
    finally:
        os.chdir('..')


def download_file(link: str, dict_for_check: dict) -> None:
    """
    Download the file at the link
    :param link: link for download
    :param dict_for_check: data for compare from html-table
    :return: None
    """
    if len(link) > 0:
        try:
            print('Starting to download the file.')
            browser_lib.execute_javascript("window.open('');")
            browser_lib_tabs = browser_lib.get_window_handles()
            browser_lib.switch_window(browser_lib_tabs[1])
            open_the_webpage(link)
            download_xpath = '//*[@id="business-case-pdf"]/a'
            wait_until_load_element(download_xpath)
            browser_lib.click_element_if_visible(download_xpath)
            # time.sleep(15)
            # downloaded_file = latest_download_file()
            downloaded_file = wait_download_file()
            print(f'File {downloaded_file} downloaded successfully.')
            data_from_downloaded_file = get_data_from_pdf_file(downloaded_file, 1)
            compare_data(data_from_downloaded_file, dict_for_check)
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
        lines_list = string.split_to_lines(text[num_page])
        for i in range(len(lines_list)):
            if "1. Name of this Investment:" in lines_list[i]:
                line_1 = lines_list[i].split(": ")
                line_2 = lines_list[i + 1].split(": ")

                key_1 = line_1[0].replace("1. ", "")
                value_1 = line_1[1]
                dict_for_check_from_pdf[key_1] = value_1
                key_2 = line_2[0].replace("2. ", "")
                value_2 = line_2[1]
                dict_for_check_from_pdf[key_2] = value_2
                break
        print('Retrieving data from the downloaded file completed.')
        return dict_for_check_from_pdf
    except Exception as e:
        print(e)
    finally:
        os.chdir('..')


def compare_data(data_1: dict, data_2: dict) -> bool:
    """
    Compare the downloaded file and data from the agency page
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
