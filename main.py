import src.file_io.file_io as file_io
import src.convert_document.convert_to_pdf as convert_to_pdf
import src.merge_document.merge_pdf as merge_pdf
import os

if __name__ == '__main__':
    file_selector = file_io.FileIO()
    to_pdf_converter = convert_to_pdf.ToPdfConverter()
    pdf_merger = merge_pdf.PdfMerger()

    input_files = file_selector.open_files()
    converted_files = []

    for index, input_file in enumerate(input_files, start=1):
        file_path, file_name = os.path.split(input_file)
        converted_file = os.path.join(file_path, f'temp{index}.pdf')

        to_pdf_converter.convert_to_pdf(input_file=input_file, output_file=converted_file)
        if(input_file.endswith('.doc')) or input_file.endswith('.docx'):
            to_pdf_converter.extract_page(converted_file, input_file)

        converted_files.append(converted_file)

    output_file = file_selector.save_file()

    pdf_merger.merge_pdf(input_files=converted_files, output_file=output_file)

    for converted_file in converted_files:
        os.remove(converted_file)
