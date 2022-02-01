import glob
import re

from pathlib import Path

from RPA.PDF import PDF
from RPA.Excel.Files import XlsWorkbook


class PDFFiles(object):
    """Class PDFFiles does: compare information from excel file and pdf file"""
    def __init__(self):
        self.glob = glob
        self.pdf = PDF()
        self.file = XlsWorkbook()
        self.output_path = str(Path(Path.cwd(), 'output'))
        self.excel_file_path = str(Path(self.output_path, 'agencies.xlsx'))

    def search_files(self):
        """
        Search files in folder "output"
        """
        list_files_pdf = self.glob.glob1(self.output_path, "*.pdf")
        return list_files_pdf

    def looking_for_info_in_excel(self):
        """
        Looking for information from excel and return list with info all uii
        """
        self.file.open(self.excel_file_path)
        worksheet = self.file.read_worksheet('Individual Investments')
        info_from_excel = []
        for file in self.search_files():
            name_file = file.replace(".pdf", '')
            for worksh in worksheet:
                if worksh['A'] == name_file:
                    info_from_excel.append(worksh)
        return info_from_excel

    def info_from_files_pdf(self):
        """
        Compare info from excel and pdf
        """
        for i in range(0, len(self.search_files())-1):
            file_path = f'output/{self.search_files()[i]}'
            text = self.pdf.get_text_from_pdf(file_path, pages=1)
            self.pdf.close_pdf()
            pdf_investment = re.findall(r'Investment:(.+?)2\.', text[1])[0].strip()
            pdf_uui = re.findall(r'\(UII\): (.*?)Se', text[1])[0]
            dict_values = self.looking_for_info_in_excel()[i]
            if dict_values['A'] == pdf_uui and dict_values['C'] == pdf_investment:
                print(True)
            else:
                print(False)
