
import tabula
import pandas as pd
import camelot
import ghostscript


def main():

    pdf_path = 'C:\\apteka\\Аптека\\спецификации для хмл\\доп\\88847\\ДС БМФ  88847 от 17 июня 2022.pdf'
    path_1 = 'C:\\apteka\\Аптека\\спецификации для хмл\\доп\\44818\\ДС ГК 44818 Витас.pdf'
    dfs = camelot.read_pdf(path_1, stream=True, pages='all')
    print(dfs)
    for i in dfs:
        print(i)

    table = dfs[1].df
    print(table.to_string())


if __name__ == '__main__':
    main()