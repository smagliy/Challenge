import glob
from PyPDF2 import PdfFileReader
import pandas as pd


def pdf_vs_xlsx():
    list_files_pdf = glob.glob1("output/", "*.pdf")
    df = pd.read_excel(io="output/new.xlsx", sheet_name='Individual Investments')
    df.set_index("UII", drop=True, inplace=True)
    dictionary = df.to_dict(orient="index")
    for file in list_files_pdf:
        print(file)
        pdf = PdfFileReader(f"output/{file}")
        page = pdf.getPage(0)
        # print(page.extractText())
        if dictionary[file[:-4]]['Bureau'] and dictionary[file[:-4]]['Investment title'] in page.extractText():
            print('Comparison result:', True)
        else:
            print('Comparison result:', False)


pdf_vs_xlsx()