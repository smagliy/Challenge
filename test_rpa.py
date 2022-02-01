import glob
from pathlib import Path

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import XlsWorkbook
from RPA.Browser.Selenium import ChromeOptions
from RPA.FileSystem import FileSystem

from pdf import PDFFiles


class ItDashBoard(object):
    """
    Class ITDashBoard does: open the browser, add elements
    to file with 2-sheet and download pdf files in folder 'output'
    """
    def __init__(self, agent):
        self.browser = Selenium()
        self.files = XlsWorkbook()
        self.fs = FileSystem()
        self.agent = agent
        self.output_path = str(Path(Path.cwd(), 'output'))
        self.chrome = ChromeOptions()
        self.preferences = {
            "download.default_directory": self.output_path,
            "directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
        self.chrome.add_experimental_option("prefs", self.preferences)
        self.excel_file_path = str(Path(self.output_path, 'agencies.xlsx'))
        self.fs.create_directory(self.output_path, parents=True)

    def main(self):
        """
        all functions
        """
        self.browser.open_chrome_browser('https://itdashboard.gov/', preferences=self.preferences)
        self.click_dive_in()
        self.files.create()
        agencies = self.get_list_of_all_agencies()
        total_spend = self.get_total_spend_for_agencies(agencies)
        self.add_new_sheet_agencies(total_spend)
        target_agency = self.look_for_agency(agencies)
        if target_agency:
            list_header_and_dict = self.agent_full_information(target_agency)
            list_href = self.add_new_worksheet_invest(list_header_and_dict)
            self.download_pdf_from_links(list_href)
        else:
            raise Exception('Can`t find selected agency')
        self.browser.close_browser()

    def click_dive_in(self):
        self.browser.click_element_when_visible('css:a.btn.btn-default.btn-lg-2x')
        self.browser.wait_until_element_is_visible('css:div#agency-tiles-widget')
        return

    def get_list_of_all_agencies(self):
        parent = self.browser.find_element('css:div#agency-tiles-widget')
        all_agencies = self.browser.find_elements('css:div.tuck-5', parent=parent)
        return all_agencies

    def get_total_spend_for_agencies(self, agencies):
        all_agencies_spend = []
        for agency in agencies:
            agency_name = self.browser.find_element('css:span.w200', parent=agency).text
            agency_amount = self.browser.find_element('css:span.w900', parent=agency).text
            all_agencies_spend.append([agency_name, agency_amount])
        return all_agencies_spend

    def look_for_agency(self, agencies):
        for agency in agencies:
            agency_name = self.browser.find_element('css:span.w200', parent=agency).text
            if agency_name == self.agent:
                return agency
        else:
            return None

    def add_new_sheet_agencies(self, data):
        """
        Create new sheet "Agencies" and add info
        """
        self.files.create_worksheet('Agencies')
        headers = [['Agency', 'Value']]
        self.files.append_worksheet('Agencies', headers)
        self.files.append_worksheet('Agencies', data)
        self.files.save(self.excel_file_path)

    def agent_full_information(self, agency):
        """
        Collect information about agent
        """
        url = self.browser.find_element('css: a', parent=agency)
        print(url.text)
        self.browser.click_element(url)
        self.browser.wait_until_element_is_visible(
            'css:select[aria-controls="investments-table-object"]', timeout=10)
        self.browser.find_element('css:select[name] option[value="-1"]').click()
        self.browser.wait_until_element_is_not_visible(
            'css:a.paginate_button[data-dt-idx="6"]', timeout=10
        )
        list_all = self.browser.find_elements(
            'css:table[id="investments-table-object"] tbody tr[role="row"]'
        )
        list_headers = [i.text for i in self.browser.find_elements('css:div tr[role="row"] th[tabindex]')]
        data_list = []
        for row in list_all:
            cols = self.browser.find_elements('css:td', parent=row)
            data = [col.text for col in cols]
            data_with_headers = dict(zip(list_headers, data))
            data_list.append(data_with_headers)
        return data_list

    def add_new_worksheet_invest(self, full_list):
        """
        Add new sheet in workbook "Individual Investment"
        """
        self.files.open(self.excel_file_path)
        self.files.create_worksheet('Individual Investments')
        self.files.append_worksheet('Individual Investments', full_list, header=True)
        self.files.save(self.excel_file_path)
        list_href = [el.get_attribute('href') for el in self.browser.find_elements(
            'css:td.left.sorting_2 a')]
        return list_href

    def download_pdf_from_links(self, hrefs):
        """
        Download pdf files
        """
        count = 0
        for link in hrefs:
            self.browser.go_to(link)
            self.browser.wait_until_element_is_visible('css:div.row.top-gutter.tuck-4 a')
            self.browser.find_element('css:div.row.top-gutter.tuck-4 a').click()
            count += 1
            list_files_pdf = glob.glob1(self.output_path, "*.pdf")
            while len(list_files_pdf) < count:
                list_files_pdf = glob.glob1(self.output_path, "*.pdf")


if __name__ == "__main__":
    main = ItDashBoard('Department of State')
    main.main()
    PDFFiles().info_from_files_pdf()