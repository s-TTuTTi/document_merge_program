import PyPDF2 as pp2

class PdfMerger:
    def __init__(self):
        pass

    def merge_pdf(self, input_files, output_file):
        merger = pp2.PdfMerger()
        for input_file in input_files:
            merger.append(input_file)
        merger.write(output_file)
        merger.close()



# def remove_temp_files(self, file_paths):
#     for file_path in file_paths:
#         if file_path.endswith('.doc'):
#             doc_pdf = file_path.replace('.doc', '.pdf')
#             os.remove(doc_pdf)
#     os.remove('temp.pdf')
#
# def convert_to_pdf(file_path):
#     current_directory = os.getcwd()
#     output_path = os.path.join(current_directory, 'output.pdf')
#
#     if file_path.endswith('.docx'):
#         # docx to pdf
#         convert(file_path, 'temp.pdf')
#         return 'temp.pdf'
#
#     elif file_path.endswith('.doc'):
#         doc_file = file_path.replace('/', '\\')
#         # Word를 시작합니다.
#         word = comtypes.client.CreateObject('Word.Application')
#         # Word 창을 숨깁니다. (백그라운드 실행)
#         word.Visible = False
#
#         # 문서를 엽니다.
#         doc = word.Documents.Open(doc_file)
#
#         output_path = doc_file.replace('.doc', '.pdf')
#         # 문서를 PDF로 저장합니다.
#         doc.SaveAs(output_path, FileFormat=17)
#         # 문서를 닫습니다.
#
#         doc.Close()
#         # Word를 종료합니다.
#         word.Quit()
#         return output_path
#
#     elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
#         # xlsx to pdf
#         app = xw.App(visible=False)
#         try:
#             book = xw.Book(file_path)
#             report_sheet = book.sheets[0]
#             report_sheet.api.ExportAsFixedFormat(0, output_path)
#         finally:
#             app.quit()
#             return output_path
#     else:
#         print("::ERROR::")
#
# def merge_pdfs(input_pdfs, output_pdf):
#     merger = PdfMerger()
#     for pdf in input_pdfs:
#         merger.append(pdf)
#     merger.write(output_pdf)
#     merger.close()
#
# def file_organization(file_paths):
#     for file_path in file_paths:
#         if file_path.endswith('.doc'):
#             doc_pdf = file_path.replace('.doc', '.pdf')
#             os.remove(doc_pdf)
#
#     os.remove('temp.pdf')



# if __name__ == "__main__":
#     file_paths = []
#     converted_pdfs = []
#
#     while True:
#         file_path = filedialog.askopenfilename()
#         if not file_path:
#             break
#         file_paths.append(file_path)
#         print(file_path)
#
#     for file_path in file_paths:
#         converted_pdfs.append(convert_to_pdf(file_path))
#
#     save_file_path = filedialog.asksaveasfilename(defaultextension='.pdf')
#
#     merge_pdfs(converted_pdfs, save_file_path)
#     file_organization(file_paths)