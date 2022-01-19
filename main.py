from parsing_bot import *

url = 'https://itdashboard.gov/'
# chrome_driver_path = "driver/chromedriverr"
departments_xpath = '//*[@id="agency-tiles-widget"]/div/div[*]/div[*]/div/div/div/div[*]/a/span[1]'
amounts_xpath = '//*[@id="agency-tiles-widget"]/div/div[*]/div[*]/div/div/div/div[*]/a/span[2]'
button_xpath = '//*[@id="node-23"]/div/div/div/div/div/div/div/a'
selected_agency = 'National Science Foundation'
# selected_agency = 'General Services Administration'
agency_blocks_xpath = '//*[@id="agency-tiles-widget"]/div/div[*]/div[*]/div/div/div/div[1]/a'
row_xpath = '//*[@id="investments-table-object"]/tbody/tr[*]'
col_xpath = '//*[@id="investments-table-object"]/tbody/tr[1]/td'
output_folder = 'output'
# file_name = './output/it-dashboards.xlsx'
file_name = 'it-dashboards.xlsx'
sheet_name = 'Agencies'


if __name__ == '__main__':
    print('The bot started working.')
    driver = set_up(output_folder)
    driver.get(url)
    click_button(driver, button_xpath, time_slip=3)
    save_to_xlsx(get_departments_amounts(driver, departments_xpath, amounts_xpath), file_name, sheet_name)
    open_agency_page(driver, selected_agency, agency_blocks_xpath)
    table, download_list, list_for_check = scrap_table(driver, row_xpath, col_xpath)
    save_to_xlsx(table, file_name, selected_agency)
    download_file(driver, download_list, list_for_check)
    print('The bot is finished.')
