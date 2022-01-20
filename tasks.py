import time
from parsing_bot import *


url = 'https://itdashboard.gov/'
button = 'home-dive-in'
departments_xpath = '//*[@id="agency-tiles-widget"]/div/div[*]/div[*]/div/div/div/div[*]/a/span[1]'
amounts_xpath = '//*[@id="agency-tiles-widget"]/div/div[*]/div[*]/div/div/div/div[*]/a/span[2]'
button_xpath = '//*[@id="node-23"]/div/div/div/div/div/div/div/a'
selected_agency = 'National Science Foundation'
# selected_agency = 'General Services Administration'
agency_blocks_xpath = '//*[@id="agency-tiles-widget"]/div/div[*]/div[*]/div/div/div/div[1]/a'
row_xpath = '//*[@id="investments-table-object"]/tbody/tr[*]'
col_xpath = '//*[@id="investments-table-object"]/tbody/tr[1]/td'
file_name = 'it-dashboards.xlsx'
sheet_name = 'Agencies'


def main():
    try:
        time_start = time.time()
        print('The bot started working.')
        open_the_website(url)
        # browser_lib.find_element(button_xpath).click()
        browser_lib.click_button(browser_lib.find_element(button_xpath))
        save_to_xlsx(get_departments_amounts(departments_xpath, amounts_xpath), file_name, sheet_name)
        open_agency_page(selected_agency, agency_blocks_xpath)
        table, download_list, list_for_check = scrap_table(row_xpath, col_xpath)
        save_to_xlsx(table, file_name, selected_agency)
        download_file(download_list, list_for_check)
        
    except Exception as e:
        print(e)
    finally:
        browser_lib.close_all_browsers()
        print('The bot is finished.')
        time_finish = time.time()
        time_all = time_finish - time_start
        print(time_all)



if __name__ == '__main__':
    main()
