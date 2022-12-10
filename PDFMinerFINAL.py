from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
import pdfminer
import xlsxwriter
import openpyxl

import pandas as pd

from datetime import datetime
from datetime import date

from read_result import read_result

def capture_invoice(filename='invoice.pdf'):
    fatura = filename
    fp = open(fatura, 'rb')

    current_day = date.today()
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")

    df = pd.DataFrame([["Capturado em: ", str(current_day), " às: ", current_time]], index=[''], columns=['PageID','X','Y','Content'])

    writer = pd.ExcelWriter("dados_fatura.xlsx", engine='openpyxl')


    df.to_excel(writer, sheet_name='Sheet1')

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    #writer.save()

    # Create a PDF parser object associated with the file object.
    parser = PDFParser(fp)

    # Create a PDF document object that stores the document structure.
    # Password for initialization as 2nd parameter
    document = PDFDocument(parser)

    # Check if the document allows text extraction. If not, abort.
    if not document.is_extractable:
        raise PDFTextExtractionNotAllowed

    # Create a PDF resource manager object that stores shared resources.
    rsrcmgr = PDFResourceManager()

    # Create a PDF device object.
    device = PDFDevice(rsrcmgr)

    # BEGIN LAYOUT ANALYSIS
    # Set parameters for analysis.
    laparams = LAParams()

    # Create a PDF page aggregator object.
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)

    # Create a PDF interpreter object.
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    def parse_obj(lt_objs):

        # loop over the object list
        for obj in lt_objs:

            # if it's a textbox, print text and location
            if isinstance(obj, pdfminer.layout.LTTextBoxHorizontal):

                    df = pd.DataFrame([[page.pageid, int(obj.bbox[0]), int(obj.bbox[1]), str(obj.get_text().replace('\n', ''))]], index=[''], columns=['PageID','X', 'Y', 'content'])

                    df.to_excel(writer, startrow = writer.sheets['Sheet1'].max_row, index=[''], header= False)

            # if it's a container, recurse
            elif isinstance(obj, pdfminer.layout.LTFigure):
                parse_obj(obj._objs)

    # loop over all pages in the document
    for page in PDFPage.create_pages(document):

        # read the page into a layout object
        interpreter.process_page(page)
        layout = device.get_result()

        # extract text from this object
        parse_obj(layout._objs)

    writer.close()

    result_json = read_result()
    return result_json
