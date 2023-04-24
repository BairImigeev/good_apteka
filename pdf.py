from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from io import StringIO


def getPDFText(pdfFilenamePath):
    retstr = StringIO()
    parser = PDFParser(open(pdfFilenamePath, 'rb'))
    try:
        document = PDFDocument(parser)
    except Exception as e:
        print(pdfFilenamePath, '')
        return ''
    if document.is_extractable:
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, retstr, laparams = LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(document):
            interpreter.process_page(page)
        return retstr.getvalue()
    else:
        print(pdfFilenamePath, "")
        return ''


if __name__ == '__main__':
    words = getPDFText('C:\\apteka\\Аптека\\спецификации для хмл\\доп\\1345\\ДС БМФ  88847 от 17 июня 2022.pdf')
    print(words)
    file = open('new.txt', 'w')
    file.write(words)