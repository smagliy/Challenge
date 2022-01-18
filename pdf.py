import glob
import re

from pathlib import Path

from RPA.PDF import PDF
from RPA.Excel.Files import Files


class PDFFiles(object):
    """Class PDFFiles does: compare information from excel file and pdf file"""
    def __init__(self):
        self.glob = glob
        self.pdf = PDF()
        self.file = Files()
        self.output_path = str(Path(Path.cwd(), 'output'))
        self.excel_file_path = str(Path(self.output_path, 'agencies.xlsx'))

    # Search files in folder "output"
    def search_files(self):
        list_files_pdf = self.glob.glob1(self.excel_file_path, "*.pdf")
        return list_files_pdf

    # Looking for information from excel and return list with info all uii
    def looking_for_info_in_excel(self):
        workbook = self.file.open_workbook(self.excel_file_path)
        worksheet = workbook.read_worksheet('Individual Investments', header=True, start=1)
        list_excel = []
        for file in self.search_files():
            name_file = file.replace(".pdf", '')
            for worksh in worksheet:
                if worksh['uii'] == name_file:
                    list_excel.append(worksh)
        return list_excel

    # Compare info from excel and pdf
    def info_from_files_pdf(self):
        for i in range(0, len(self.search_files())-1):
            file_path = f'output/{self.search_files()[i]}'
            text = self.pdf.get_text_from_pdf(file_path, pages=1)
            pdf_investment = re.findall(r'Investment:(.+?)2\.', text[1])[0].strip()
            pdf_uui = re.findall(r'\(UII\): (.*?)Se', text[1])[0]
            dict_values = self.looking_for_info_in_excel()[i]
            if dict_values['uii'] == pdf_uui and dict_values['investment title'] == pdf_investment:
                print(True)
            else:
                print(False)
                print(dict_values['uii'], dict_values['investment title'])
                print(pdf_uui, pdf_investment)


pdf = PDFFiles()
pdf.info_from_files_pdf()