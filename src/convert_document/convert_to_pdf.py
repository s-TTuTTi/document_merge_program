# from docx2pdf import convert
# import comtypes.client
# import xlwings as xw
import win32com.client

class ToPdfConverter:
    def __init__(self):
        pass

    # def docx2pdf(self, input_file, output_file):
    #     convert(input_file, output_file)
    #
    # def doc2pdf(self, input_file, output_file):
    #     doc_file = input_file.replace('/', '\\')
    #
    #     word = comtypes.client.CreateObject('Word.Application')
    #     word.Visible = False
    #
    #     doc = word.Documents.Open(doc_file)
    #
    #     doc_output_file = output_file.replace('/', '\\')
    #
    #     doc.SaveAs(doc_output_file, FileFormat=17)
    #
    #     doc.Close()
    #     word.Quit()
    #
    # def excel2pdf(self, input_file, output_file):
    #     app = xw.App(visible=False)
    #     try:
    #         book = xw.Book(input_file)
    #         report_sheet = book.sheets[0]
    #         report_sheet.api.ExportAsFixedFormat(0, output_file)
    #     finally:
    #         app.quit()

    def word2pdf(self, input_file, output_file):
        doc_file = input_file.replace('/', '\\')

        wdFormatPDF = 17

        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(doc_file)

        doc_output_file = output_file.replace('/', '\\')

        doc.SaveAs(doc_output_file, FileFormat=wdFormatPDF)

        doc.Close()
        word.Quit()

    def excel2pdf(self, input_file, output_file):
        excel = win32com.client.Dispatch('Excel.Application')
        wb = excel.Workbooks.Open(input_file)

        for ws in wb.Worksheets:
            ws.Select()
            wb.ActiveSheet.ExportAsFixedFormat(0, output_file)

        wb.Close()
        excel.Quit()

    def convert_to_pdf(self, input_file, output_file):
        if input_file.endswith('.docx') or input_file.endswith('.doc'):
            self.word2pdf(input_file, output_file)
        elif input_file.endswith('.xlsx') or input_file.endswith('.xls'):
            self.excel2pdf(input_file, output_file)
        else:
            print("::ERROR::")
