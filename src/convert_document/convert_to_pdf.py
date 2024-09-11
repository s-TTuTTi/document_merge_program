import win32com.client
import os
import comtypes.client
import win32com.client as win32
class ToPdfConverter:
    def __init__(self):
        self.word = win32com.client.Dispatch('Word.Application')
        self.excel = win32com.client.Dispatch('Excel.Application')

    def __del__(self):
        if self.word:
            self.word.Quit()
        if self.excel:
            self.excel.Quit()
    def initialize_word_application(self):
        try:
            self.word = comtypes.client.CreateObject('Word.Application')
            self.word.Visible = False  # Word 창을 보이지 않게 실행
            print("Word application initialized.")
        except Exception as e:
            print(f"Error initializing Word application: {str(e)}")

    def word2pdf(self, input_file, output_file):
        wdExportFormatPDF = 17
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False  # Word 창을 보이지 않게 설정
        doc_file = input_file.replace('/', '\\')
        print(f"Opening Word file: {doc_file}")
        doc = word.Documents.Open(doc_file)
        print(f"Converting to PDF: {output_file}")
        doc.SaveAs(output_file, FileFormat=wdExportFormatPDF)
        doc.Close()
        word.Quit()

    def excel2pdf(self, input_file, output_file, sheet_name=None):
        xlExportFormatPDF = 0

        wb = self.excel.Workbooks.Open(input_file)

        if sheet_name is None:
            wb.ExportAsFixedFormat(xlExportFormatPDF, output_file)
        else:
            ws = wb.Worksheets(sheet_name)
            ws.ExportAsFixedFormat(xlExportFormatPDF, output_file)

        wb.Close()

    def convert_to_pdf(self, input_file, output_file, sheet_name=None):
        if input_file.endswith('.docx') or input_file.endswith('.doc'):
            self.word2pdf(input_file, output_file)
            if not os.path.isfile(output_file):
                print(f"PDF file was not created: {output_file}")
                return

        elif input_file.endswith('.xlsx') or input_file.endswith('.xls'):
            self.excel2pdf(input_file, output_file, sheet_name)
        else:
            print("::ERROR::")

if __name__ == '__main__':
    converter = ToPdfConverter()
    converter.convert_to_pdf('C:/Users/yean/Desktop/문서 통합/Sample DATA/test3.xlsx', 'C:/Users/yean/Desktop/문서 통합/Sample DATA/test3.pdf')
