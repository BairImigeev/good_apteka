
import tabula
import pandas as pd
import camelot
import ghostscript
import tabula
import pdfplumber
from hochiminh import pdf_parser
# from hochiminh.io.pdfconverter import PDFConverter
from hochiminh.io.reader import ImageReader
from pdf2image import convert_from_path
import os
from hochiminh.image_processing.lines_detector import SobelDirector

path_1 = 'C:\\apteka\\Аптека\\спецификации для хмл\\dops\\44818'
# pdf_path = 'C:\\apteka\\Аптека\\спецификации для хмл\\доп\\88847\\ДС БМФ  88847 от 17 июня 2022.pdf'
resolution = 100


def get_locate_file():
    locate_files = []
    search_files =[]
    for rootdir, dirs, files in os.walk(path_1):
        for file in files:
            if (file.split('.')[-1]) =='pdf':
                locate = os.path.join(rootdir, file)
                locate_files.append(rootdir)
                search_files.append(locate)
    # print(locate_files,search_files)
    return locate_files, search_files


def main():
    dirr = get_locate_file()[0][0]
    print(dirr)
    s = get_locate_file()[1][0]
    print(s)
    pic = convert_from_path(s, dpi=resolution, output_folder=dirr,
                                             poppler_path='C:\\Poppler\\poppler-23.01.0\\Library\\bin')

    pd = pdf_parser.PDFParser.extract_table
    print(pd)






    # dfs = tabula.read_pdf(path_1, pages='all')
    # print(dfs)

    # tables = pdf.pages[2].debug_tablefinder()
    # print(tables)
    # # table = pdf.pages[2].extract_table(
                                       #      table_settings=
                                       # {
                                       #         "vertical_strategy": "text",
                                       #         "horizontal_strategy": "text",
                                       #  }
        # )
    #
    # print(table)
    # df = pd.DataFrame(table, columns=table)
    # for i in df.columns:
    #     print(i)

    # for i in range(int(len(pdf.pages))):
    #     df = pd.DataFrame()
    #     table = pdf.pages[i].extract_table(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
    #     df = pd.DataFrame(table, columns=table)
    #     # print(df)
    #     df.to_csv('test.csv', mode="a", index=False)

    # table = dfs[1]
    # print(table.to_string())


if __name__ == '__main__':
    main()