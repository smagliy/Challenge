from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
import time
import xlsxwriter

path = os.path.dirname(os.path.abspath(__file__)) + '\output'
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
    time.sleep(15)
    list_all_even = [[i.text for i in el.find_elements_by_css_selector('td')]
                for el in driver.find_elements_by_css_selector('tr.even')]
    list_all_odd = [[i.text for i in el.find_elements_by_css_selector('td')]
                for el in driver.find_elements_by_css_selector('tr.odd')]
    list_all = list_all_odd + list_all_even
    dict_full = {
        'uii_href': [el.get_attribute('href') for el in
                     driver.find_elements_by_css_selector('td.left.sorting_2 a')],
        'uii': [i[0] for i in list_all],
        'bureau': [i[1] for i in list_all],
        'investment_title': [i[2] for i in list_all],
        'total': [i[3] for i in list_all],
        'type': [i[4] for i in list_all],
        'cio': [i[5] for i in list_all],
        'projects': [i[6] for i in list_all]
    }
    return dict_full


def add_excel_agent(dict_full, dict_el):
    writer = pd.ExcelWriter('output/new.xlsx', engine='xlsxwriter')
    df1 = pd.DataFrame({'Name': dict_el['name'], 'Salary': dict_el['salary']})
    df1.to_excel(writer, 'Angencies', index=False)
    df2 = pd.DataFrame({'UII': dict_full['uii'],
                        'Bureau': dict_full['bureau'],
                        'Investment title': dict_full['investment_title'],
                        'Total': dict_full['total'],
                        'Type': dict_full['type'],
                        'CIO': dict_full['cio'],
                        'Projects': dict_full['projects']})
    df2.to_excel(writer, 'Individual Investments', index=False)
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
    driver.quit()