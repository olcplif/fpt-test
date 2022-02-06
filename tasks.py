from parsing_bot import (
    browser_lib,
    open_the_webpage,
    get_departments_amounts,
    save_to_xlsx,
    scrap_table,
    find_links,
    download_file,
    get_data_from_pdf_file,
    compare_data
)

url = 'https://itdashboard.gov/'
departments_xpath = '//*[@id="agency-tiles-widget"]//span[@class="h4 w200"]'  # xpath to agency names
amounts_xpath = '//*[@id="agency-tiles-widget"]//span[@class=" h1 w900"]'  # xpath to agency amounts
dive_in_button_xpath = '//*[@id="node-23"]//a[@class="btn btn-default btn-lg-2x trend_sans_oneregular"]'
selected_agency = 'National Science Foundation'
rows_xpath = '//*[@id="investments-table-object"]/tbody/tr[*]'  # xpath to the table on the agency page
file_name = 'it-dashboards.xlsx'
sheet_name = 'Agencies'


def main():
    try:
        # orig start
        # print('The bot started working.')
        # open_the_webpage(url)
        # browser_lib.click_button(browser_lib.find_element(dive_in_button_xpath))
        # save_to_xlsx(get_departments_amounts(departments_xpath, amounts_xpath), file_name, sheet_name)
        # scrap_table(file_name, selected_agency, row_xpath)
        # orig finish

        print('The bot started working.')
        open_the_webpage(url)
        browser_lib.click_button(browser_lib.find_element(dive_in_button_xpath))
        save_to_xlsx(get_departments_amounts(departments_xpath, amounts_xpath), file_name, sheet_name)
        save_to_xlsx(scrap_table(selected_agency, rows_xpath), file_name, selected_agency)
        for link in find_links(rows_xpath):
            file = download_file(link[0])
            compare_data(get_data_from_pdf_file(file, 1), link[1])
    except Exception as e:
        print(e)
    finally:
        browser_lib.close_browser()
        print('The bot is finished.')


if __name__ == '__main__':
    main()
