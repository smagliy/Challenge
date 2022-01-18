import time
from pathlib import Path
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from pdf import PDFFiles


class ItDashBoard(object):
    """Class ITDashBoard does: open the browser, add elements
    to file with 2-sheet and download pdf files in folder 'output'"""
    def __init__(self, agent):
        self.browser = Selenium()
        self.files = Files()
        self.agent = agent

    def click_dive_in(self):
        btn = self.browser.find_element('css:a.btn.btn-default.btn-lg-2x')
        btn.click()

    # Looking for information from table in site`s department and add to new_1.xlsx file
    def search_informations(self):
        self.browser.wait_until_element_is_visible('css:a.btn.btn-default.btn-sm')
        dict_elements = {
            'button': [el.get_attribute('href') for el in self.browser.find_elements('css:a.btn.btn-default.btn-sm')],
            'name': [el.text for el in self.browser.find_elements('css:span.h4.w200')],
            'salary': [el.text for el in self.browser.find_elements('css:span.h1.w900')]
        }
        workbook = self.files.create_workbook('output/new_1.xlsx')
        workbook.create_worksheet('Agencies')
        workbook.append_worksheet('Agencies', {'name': dict_elements['name'], 'salary': dict_elements['salary']})
        workbook.save()
        return dict_elements

    # Search index in dictionary['button']
    def search_agent(self, dict_info):
        index = dict_info['name'].index(self.agent)
        return dict_info['button'][index]

    # Collect information about agent and add to new_1.xlsx file
    def agent_full_information(self, url):
        self.browser.go_to(url)
        self.browser.click_element_if_visible(
            'css:select.form-control.c-select[aria-controls="investments-table-object"]')
        btn_all = self.browser.find_element(
            'css:select.form-control.c-select[aria-controls] option[value="-1"]').click()
        time.sleep(15)
        list_all = self.browser.find_elements('css:table[id="investments-table-object"] tbody tr[role="row"]')
        dict_all = {
            'uii': [],
            'bureau': [],
            'investment title': [],
            'total': [],
            'type': [],
            'cio': [],
            'projects': []
        }
        for row in list_all:
            cols = self.browser.find_elements('css:td', parent=row)
            dict_all['uii'].append(cols[0].text)
            dict_all['bureau'].append(cols[1].text)
            dict_all['investment title'].append(cols[2].text)
            dict_all['total'].append(cols[3].text)
            dict_all['type'].append(cols[4].text)
            dict_all['cio'].append(cols[5].text)
            dict_all['projects'].append(cols[6].text)
        workbook = self.files.open_workbook('output/new_1.xlsx')
        workbook.create_worksheet('Individual Investments')
        workbook.append_worksheet('Individual Investments', dict_all,
                                  header=['uii', 'bureau', 'investment title', 'total', 'type', 'cio', 'projects'])
        workbook.save()
        list_href = [el.get_attribute('href') for el in self.browser.find_elements('css:td.left.sorting_2 a')]
        return list_href

    # download pdf files
    def links_to_go(self, hrefs):
        path_to_download_folder = str(Path(Path.cwd(), 'output'))
        self.browser.set_download_directory(path_to_download_folder)
        for link in hrefs:
            self.browser.go_to(link)
            self.browser.click_element_if_visible('css:div.row.top-gutter.tuck-4 a')
            time.sleep(10)

    # all functions
    def main(self):
        self.browser.open_available_browser('https://itdashboard.gov/')
        self.click_dive_in()
        dict_values = self.search_informations()
        list_href = self.agent_full_information(self.search_agent(dict_values))
        self.links_to_go(list_href)
        self.browser.close_browser()


if __name__ == "__main__":
    main = ItDashBoard('Department of State')
    main.main()
    PDFFiles().info_from_files_pdf()
