from parsing_bot import (
    browser_lib,
    open_the_webpage,
    get_departments_amounts,
    save_to_xlsx,
    scrap_table,
)

url = 'https://itdashboard.gov/'
departments_xpath = '//*[@id="agency-tiles-widget"]//span[@class="h4 w200"]'
amounts_xpath = '//*[@id="agency-tiles-widget"]//span[@class=" h1 w900"]'
dive_in_button_xpath = '//*[@id="node-23"]//a[@class="btn btn-default btn-lg-2x trend_sans_oneregular"]'
selected_agency = 'National Science Foundation'
agency_blocks_xpath = '//*[@id="agency-tiles-widget"]//a'
row_xpath = '//*[@id="investments-table-object"]/tbody/tr[*]'
file_name = 'it-dashboards.xlsx'
sheet_name = 'Agencies'


def main():
    try:
        print('The bot started working.')
        open_the_webpage(url)
        browser_lib.click_button(browser_lib.find_element(dive_in_button_xpath))
        save_to_xlsx(get_departments_amounts(departments_xpath, amounts_xpath), file_name, sheet_name)
        scrap_table(file_name, selected_agency, agency_blocks_xpath, row_xpath)
    except Exception as e:
        print(e)
    finally:
        browser_lib.close_browser()
        print('The bot is finished.')


if __name__ == '__main__':
    main()
