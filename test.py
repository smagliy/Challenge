from RPA.Browser.Selenium import Selenium
from RPA.Browser.Selenium import ChromeOptions
import os
import time
import xlsxwriter

browser_lib = Selenium()
options = ChromeOptions()
path = os.path.dirname(os.path.abspath(__file__))
prefs = {"download.default_directory": path}
options.add_experimental_option("prefs", prefs)


def main():
    browser_lib.create_webdriver('Chrome', options=options)
    browser_lib.driver.get('https://itdashboard.gov/')
    btn = browser_lib.driver.find_element_by_css_selector('a.btn.btn-default.btn-lg-2x')
    btn.click()
    time.sleep(1)
    names = browser_lib.driver.find_elements_by_css_selector('span.h4.w200')
    total_salary = browser_lib.driver.find_elements_by_css_selector('span.h1.w900')
    name = [elem.text for elem in names if elem.text != '']
    salary = [elem.text for elem in total_salary]
    workbook = xlsxwriter.Workbook('output/agencies.xlsx')
    worksheet = workbook.add_worksheet()
    for i in range(0, len(name)-1):
        worksheet.write(f'A{i+1}', name[i])
        worksheet.write(f'B{i+1}', salary[i])
    workbook.close()


def random_company():
    main()
    href = browser_lib.driver.find_elements_by_css_selector('a.btn.btn-default.btn-sm')
    href_list = []
    for h in href:
        new_href = h.get_attribute('href')
        href_list.append(new_href)
    
    browser_lib.driver.get(href_list[2])
    time.sleep(10)
    btn_select = browser_lib.driver.find_element_by_css_selector('select.form-control.c-select[aria-controls="investments-table-object"]')
    btn_select.click()
    btn_all = browser_lib.driver.find_element_by_css_selector('select.form-control.c-select[aria-controls] option[value="-1"]')
    btn_all.click()
    time.sleep(10)
    dict_all = {
        'uii': [el.text for el in browser_lib.driver.find_elements_by_css_selector('td.left.sorting_2')],
        'uii_href': [el.get_attribute('href') for el in browser_lib.driver.find_elements_by_css_selector('td.left.sorting_2 a')],
        'total': [el_1.text for el_1 in browser_lib.driver.find_elements_by_css_selector('td.right')]
    }
    workbook = xlsxwriter.Workbook('output/uii.xlsx')
    worksheet = workbook.add_worksheet()
    for i in range(0, len(dict_all['uii']) - 1):
        worksheet.write(f'A{i + 1}', dict_all['uii'][i])
        worksheet.write(f'B{i + 1}', dict_all['total'][i])
    workbook.close()
    for href1 in dict_all['uii_href']:
        browser_lib.driver.get(href1)
        time.sleep(5)
        btn1 = browser_lib.driver.find_element_by_css_selector('div.row.top-gutter.tuck-4 a')
        btn1.click()
        time.sleep(10)


# Call the main() function, checking that we are running as a stand-alone script:
if __name__ == "__main__":
    random_company()
#   main()
