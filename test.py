from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import os
import time
import xlsxwriter

path = os.path.dirname(os.path.abspath(__file__))
prefs = {"download.default_directory": path}
options = Options()
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)


def open_browser_and_click_btn():
    driver.get('https://itdashboard.gov')
    btn = driver.find_element_by_css_selector('a.btn.btn-default.btn-lg-2x')
    btn.click()


def search_for_information():
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.btn.btn-default.btn-sm')))
    dict_el = {
        'button': [el.get_attribute('href') for el in driver.find_elements_by_css_selector(
            'a.btn.btn-default.btn-sm')],
        'name': [elem.text for elem in driver.find_elements_by_css_selector('span.h4.w200')
                 if elem.text != ''],
        'salary': [elem.text for elem in driver.find_elements_by_css_selector('span.h1.w900')]
    }
    return dict_el


def search_agent(dict_info, name_agency):
    index = dict_info['name'].index(name_agency)
    return dict_info['button'][index]


def agent_full_information(url):
    driver.get(url)
    wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR, 'select.form-control.c-select[aria-controls="investments-table-object"]')))
    btn_select = driver.find_element_by_css_selector(
        'select.form-control.c-select[aria-controls="investments-table-object"]')
    btn_select.click()
    btn_all = driver.find_element_by_css_selector(
        'select.form-control.c-select[aria-controls] option[value="-1"]')
    btn_all.click()
    time.sleep(10)
    dict_full = {
        'uii': [el.text for el in driver.find_elements_by_css_selector('td.left.sorting_2')],
        'uii_href': [el.get_attribute('href') for el in
                     driver.find_elements_by_css_selector('td.left.sorting_2 a')],
        'total': [el_1.text for el_1 in driver.find_elements_by_css_selector('td.right')]
    }
    return dict_full


def add_excel_agent(dict_full, dict_el):
    df1 = pd.DataFrame(dict_el['name'],
                       dict_el['salary'])

    df2 = pd.DataFrame(dict_full['uii'],
                       dict_full['total'])
    writer = pd.ExcelWriter('output/agencies.xlsx', engine='xlsxwriter')
    df1.to_excel(writer, sheet_name='Agencies')
    df2.to_excel(writer, sheet_name='Individual Investments')
    writer.save()
    return dict_full


def main():
    open_browser_and_click_btn()
    dict_agency = search_for_information()
#   print(dict_agency)
    url_agent = search_agent(dict_agency, 'Department of Agriculture')
    dict_only_one = agent_full_information(url_agent)
#   print(dict_only_one)
    add_excel_agent(dict_only_one, dict_agency)
    download_pdf(dict_only_one)


def download_pdf(dict_only):
    for href1 in dict_only['uii_href']:
        driver.get(href1)
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.row.top-gutter.tuck-4 a')))
        btn1 = driver.find_element_by_css_selector('div.row.top-gutter.tuck-4 a')
        btn1.click()
        time.sleep(10)


if __name__ == "__main__":
    main()