
import shutil, tempfile
import os, re, sys

import comtypes
import comtypes.client

import src.gui_document.handle_gui as handle_gui
import src.file_io.file_io as file_io
import src.convert_document.convert_to_pdf as convert_to_pdf

from PyQt5.QtCore import Qt, QTimer, QRect
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtWidgets import QTableWidget, QAbstractItemView

from src.handle_document import handle_pdf, handle_docx, handle_xlsx
from src.merge_document import merge_pdf


def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

form = resource_path('mainUI.ui')
UI_main = uic.loadUiType(form)[0]



class MyWindow(QMainWindow, UI_main):
    def __init__(self):
        super().__init__()
        self.initUi()
        # Functions ===========================
        self.file_selector = file_io.FileIO()
        self.to_pdf_converter = convert_to_pdf.ToPdfConverter()
        self.pdf_handler = handle_pdf.PdfHandler()
        self.word_handler = handle_docx.WordHandler()
        self.excel_handler = handle_xlsx.ExcelHandler()
        self.pdf_merger = merge_pdf.PdfMerger()
        # Variables ===========================
        self.input_files = []
        self.selected_pages = []
        self.total_pages = []

        self.check_box = []
        self.currently_row = -1


    def initUi(self):
        self.setupUi(self)
        self.table_widget_setting()
        self.btn_guide_setting()
        self.btn_load_setting()
        self.btn_merge_setting()
        self.btn_move_setting()
        self.init_progressBar_setting()



    # =============================================================================================== Button Guide =========
    # Button Guide
    def btn_guide_setting(self):
        self.btn_guide.released.connect(self.btn_guide_click)

    def btn_guide_click(self):
        self.set_info("USER GUIDE")
        QMessageBox.about(self, 'Guide',
                          """Guide:
- To load a file, click the 'Load' button.
- To configure Word, PDF (Pages), or Excel (Worksheet), double-click on the corresponding item.
- If the file is password-protected, an error may occur during merging, so please check the file in advance.
- In rare cases where the page count cannot be detected, manually open the file to confirm the total number of pages and select the pages accordingly.
"""
)

    # =============================================================================================== Button Load =========
    # Button Load
    def btn_load_setting(self):
        self.btn_load.released.connect(self.btn_load_click)

    def btn_load_click(self):
        if not self.btn_load.isEnabled():  # 이미 클릭된 상태면 실행 방지
            return
        self.set_info("Files Loading..")

        files = self.file_selector.open_files()
        # Timer
        self.btn_load_timer_setting(False)
        QTimer.singleShot(500, lambda: self.btn_load_timer_setting(True))

        self.tableWidget.setRowCount(len(files+self.input_files))
        new_files_count = 0
        for file in files:
            if file.endswith((".docx", ".doc", ".xlsx", ".xls", ".pdf")):
                self.input_files.append(file)
                self.table_widget_insert(file, len(self.input_files) - 1)
                new_files_count += 1

        self.tableWidget.setRowCount(len(self.input_files))
        if self.selected_pages:
            self.selected_pages.extend([0] * new_files_count)
            self.total_pages.extend([0] * new_files_count)
        else:
            self.selected_pages = [0] * len(self.input_files)
            self.total_pages = [0] * len(self.input_files)
        self.set_info("Loading completed !")

        print("self.input_files : ",self.input_files)
        print("self.selected_pages : ", self.selected_pages)
        print("self.total_pages : ",self.total_pages)

    def btn_load_timer_setting(self, value):
        self.btn_load.setEnabled(value)

    # =============================================================================================== Merge Events =========

    @staticmethod
    def extract_filename_without_extension(file_path):
        file_name = os.path.basename(file_path)
        file_name_without_ext = os.path.splitext(file_name)[0]
        return file_name_without_ext

    @staticmethod
    def parse_num_ranges(text):
        nums = []
        ranges = text.split(',')
        for r in ranges:
            r = r.strip()
            if '-' in r:
                start, end = map(int, r.split('-'))
                nums.extend(range(start, end + 1))
            else:
                nums.append(int(r))
        return nums

    @staticmethod
    def decrement_array(arr):
        return [x - 1 for x in arr]



    def convert_to_pdf(self, input_file, output_file, selected_page):

        if input_file.endswith('.doc') or input_file.endswith('.docx'):
        # Word -> PDF
            self.to_pdf_converter.word2pdf(input_file, output_file)
        # What happens when you have a set page
            if selected_page != 0:
                page_array = self.parse_num_ranges(selected_page)
                decremented_array = self.decrement_array(page_array)
                self.pdf_handler.extract_page(output_file, decremented_array, output_file)

        elif input_file.endswith('.pdf'):

        # What happens when you have a set page
            if selected_page != 0:
                page_array = self.parse_num_ranges(selected_page)
                decremented_array = self.decrement_array(page_array)
                self.pdf_handler.extract_page(input_file, decremented_array, output_file)
            else:
                shutil.copyfile(input_file, output_file)

        elif input_file.endswith('.xls') or input_file.endswith('.xlsx'):
        # What happens when you have a set worksheet
            if selected_page != 0:
                self.to_pdf_converter.excel2pdf(input_file, output_file, selected_page)
            else:
                self.to_pdf_converter.excel2pdf(input_file, output_file)
        else:
            self.warning_msg("extension error")

    def convert_files(self, temp_dir):
        converted_files = []
        file_names_without_ext = []

        i = len(self.input_files) / 20
        value = 0


        for input_file, selected_page in zip(self.input_files, self.selected_pages):
            file_name_without_ext = self.extract_filename_without_extension(input_file)  # 확장명을 제외한 파일명 가져오기
            converted_file = os.path.join(temp_dir, f'{file_name_without_ext}.pdf')

            self.convert_to_pdf(input_file, converted_file, selected_page)

            file_names_without_ext.append(file_name_without_ext)
            converted_files.append(converted_file)

            self.update_progress_bar(value)
            value += i

        return converted_files, file_names_without_ext

    def get_page_numbers(self, converted_files):
        files_page_num = []
        for converted_file in converted_files:
            page_num = self.pdf_handler.extract_page_num(converted_file)
            files_page_num.append(page_num)
        return files_page_num

    def create_cover_and_index_pages(self, temp_dir, file_names_without_ext, files_page_num):
        pjt_no = self.txe_pjtno.toPlainText()
        title = self.txe_dc.toPlainText()
        dept_name = self.txe_dp.toPlainText()
        person_name = self.txe_per.toPlainText()
        cover_page_docx_path = os.path.join(temp_dir, "cover_page.docx")
        index_page_docx_path = os.path.join(temp_dir, "index_page.docx")
        self.word_handler.create_cover_page(title, pjt_no, dept_name, person_name, cover_page_docx_path)
        self.word_handler.create_index_page(file_names_without_ext, files_page_num, index_page_docx_path)
        cover_page_pdf_path = os.path.join(temp_dir, "cover_page.pdf")
        index_page_pdf_path = os.path.join(temp_dir, "index_page.pdf")
        self.to_pdf_converter.convert_to_pdf(cover_page_docx_path, cover_page_pdf_path)
        self.to_pdf_converter.convert_to_pdf(index_page_docx_path, index_page_pdf_path)
        return cover_page_pdf_path, index_page_pdf_path

    @staticmethod
    def insert_cover_and_index_pages(converted_files, cover_page_pdf_path, index_page_pdf_path):
        converted_files.insert(0, index_page_pdf_path)  # index page
        converted_files.insert(0, cover_page_pdf_path)  # cover page
        return converted_files

    def insert_cover_and_index_page_numbers(self, files_page_num, cover_page_pdf_path, index_page_pdf_path):
        cover_page_num = self.pdf_handler.extract_page_num(cover_page_pdf_path)
        index_page_num = self.pdf_handler.extract_page_num(index_page_pdf_path)
        files_page_num.insert(0, index_page_num)
        files_page_num.insert(0, cover_page_num)
        return files_page_num

    def merge_files(self, converted_files, temp_dir):
        merged_file_path = os.path.join(temp_dir, "merged.pdf")
        self.pdf_merger.merge_pdf(input_files=converted_files, output_file=merged_file_path)
        return merged_file_path

    def insert_page_numbers(self, merged_file_path, temp_dir):
        numbering_file_path = os.path.join(temp_dir, "page.pdf")
        self.pdf_handler.insert_page_number(merged_file_path, numbering_file_path, 1)
        return numbering_file_path

    def add_bookmarks(self, numbering_file_path, file_names_without_ext, files_page_num, temp_dir):
        bookmark_file_path = os.path.join(temp_dir, "page.pdf")
        file_names_without_ext.insert(0, "index_page")
        file_names_without_ext.insert(0, "cover_page")
        self.pdf_handler.add_bookmark(numbering_file_path, file_names_without_ext, files_page_num, bookmark_file_path)
        return bookmark_file_path
# =============================================================================================== Button Merge =========
# Button Merge

    def init_progressBar_setting(self):
        self.pbar.reset()
        self.pbar.setVisible(True)

    def progressBar_setting(self):
        self.pbar.setStyleSheet("QProgressBar{\n"
                                       "    background-color: rgb(98, 114, 164);\n"
                                       "    color:rgb(200,200,200);\n"
                                       "    border-style: none;\n"
                                       "    border-bottom-right-radius: 10px;\n"
                                       "    border-bottom-left-radius: 10px;\n"
                                       "    border-top-right-radius: 10px;\n"
                                       "    border-top-left-radius: 10px;\n"
                                       "    text-align: center;\n"
                                       "}\n"
                                       "QProgressBar::chunk{\n"
                                       "    border-bottom-right-radius: 10px;\n"
                                       "    border-bottom-left-radius: 10px;\n"
                                       "    border-top-right-radius: 10px;\n"
                                       "    border-top-left-radius: 10px;\n"
                                       "    background-color: qlineargradient(spread:pad, x1:0, y1:0.511364, x2:1, y2:0.523, stop:0 rgba(254, 121, 199, 255), stop:1 rgba(170, 85, 255, 255));\n"
                                       "}\n"
                                       "\n"
                                       "")

    def update_progress_bar(self, value):
        self.pbar.setValue(int(value))
        self.pbar.update()

    def btn_merge_setting(self):
        self.btn_merge.released.connect(self.btn_merge_click)

    def btn_merge_click(self):
        self.set_info("Files Merging..")
        self.progressBar_setting()
        self.update_progress_bar(0)
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                # 입력 파일 변환
                converted_files, file_names_without_ext = self.convert_files(temp_dir)
                files_page_num = self.get_page_numbers(converted_files)
                self.update_progress_bar(30)
                # 커퍼 페이지 및 목차 페이지 생성
                cover_page_pdf_path, index_page_pdf_path = self.create_cover_and_index_pages(temp_dir, file_names_without_ext, files_page_num)
                self.update_progress_bar(40)
                # 커퍼 페이지 및 목차 페이지를 변환된 파일들 맨 앞에 추가
                converted_files = self.insert_cover_and_index_pages(converted_files, cover_page_pdf_path, index_page_pdf_path)
                self.update_progress_bar(50)
                # 커버 페이지 및 목차 페이지의 페이지 수를 변환된 파일들의 페이지 수의 맨 앞에 추가
                files_page_num = self.insert_cover_and_index_page_numbers(files_page_num, cover_page_pdf_path, index_page_pdf_path)
                self.update_progress_bar(60)
                # 하나의 pdf로 병합
                merged_file_path = self.merge_files(converted_files, temp_dir)
                self.update_progress_bar(80)
                # 병합된 pdf에 번호 추가
                numbering_file_path = self.insert_page_numbers(merged_file_path, temp_dir)
                self.update_progress_bar(90)
                # 북마크 추가
                bookmark_file_path = self.add_bookmarks(numbering_file_path, file_names_without_ext, files_page_num, temp_dir)
                self.update_progress_bar(100)
                # 저장 경로 선택
                output_file = self.file_selector.save_file()

                if not output_file:
                    self.set_info("No storage path specified",1)
                else:
                    # 사용자 지정 경로 복사
                    shutil.copyfile(bookmark_file_path, output_file)
                    self.set_info("The file has been saved")

                self.pbar.reset()
                self.pbar.setVisible(True)
        except Exception as e:
            self.critical_event(f"Error during file conversion: {e}")

# =============================================================================================== Table Widget =========
# Table Widget
    def table_widget_setting(self):
        # Header size
        header = self.tableWidget.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        # Table Select
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tableWidget.setSelectionBehavior(QAbstractItemView.SelectRows)

        # ScrollBar ====================================================
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.tableWidget.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        self.tableWidget.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        # Events ====================================================
        self.tableWidget.cellClicked.connect(self.table_widget_select)
        self.tableWidget.cellDoubleClicked.connect(self.table_widget_doubleClick)


    def table_widget_insert(self, file, i):
        base_name = os.path.basename(file)
        name, extension = os.path.splitext(base_name)
        self.tableWidget.setItem(i, 1, QTableWidgetItem(name))
        self.tableWidget.setItem(i, 2, QTableWidgetItem(extension[1:]))
        #CheckBox
        item = QCheckBox()
        self.check_box.append(item)
        cellWidget = QWidget()
        layoutCB = QHBoxLayout(cellWidget)
        layoutCB.addWidget(self.check_box[i])
        layoutCB.setAlignment(Qt.AlignCenter)
        layoutCB.setContentsMargins(0,0,0,0)
        cellWidget.setLayout(layoutCB)

        self.tableWidget.setCellWidget(i,0,cellWidget)

    def total_pages_setting(self, row):
        file = self.input_files[row]

        try:
            if file.endswith('.docx') or file.endswith('.doc'):
                try:
                    page_num = self.word_handler.extract_page_num(file)
                    self.total_pages[row] = page_num  # 페이지 수 업데이트
                    if self.total_pages[row] == 0:  # 업데이트 후 체크
                        self.critical_event("User accesses directly because page information cannot be obtained")
                        self.total_pages[row] = -1
                    else:
                        self.set_info("Imported page information!")
                except Exception as e:
                    self.critical_event(f"Failed to extract page number from Word file: {e}")
                    self.total_pages[row] = -1

            elif file.endswith('.pdf'):
                try:
                    page_num = self.pdf_handler.extract_page_num(file)

                    self.total_pages[row] = page_num  # 페이지 수 업데이트
                # 페이지 수가 출력이 되지 않았을 경우
                    if self.total_pages[row] == 0:  # 업데이트 후 체크
                        self.critical_event("User accesses directly because page information cannot be obtained")
                        self.total_pages[row] = -1
                # PDF file is password
                    elif self.total_pages[row] == -5:
                        self.critical_event(f"PDF file is password protected and cannot be opened")
                        self.set_info("The pdf file has been deleted",1)
                    # 페이지 수를 잘 출력하였을 경우
                    else:
                        self.set_info("Imported page information!")
                except Exception as e:
                    self.critical_event(f"Failed to extract page number from PDF file: {e}")
                    self.total_pages[row] = -1

        except Exception as e:
            self.critical_event(f"An unexpected error occurred: {e}")
            self.total_pages[row] = -1


    #Click Event
    def table_widget_select(self, row):
        self.currently_row = row
        print(f"** Row: {row+1} clicked ** ")
        print("Number of pages set : ", self.selected_pages)
        print("Total number of pages updated : ",self.total_pages)
        print("self.input_files : ",self.input_files)


        item = self.tableWidget.item(row, 1)
        if item is not None:
            value = item.text()
            label_string = f'{row + 1} : {value}'
            self.set_info(label_string)
        else:
            print("No item at the clicked position")

    def table_widget_doubleClick(self, row):
        self.currently_row = row
        print("** Double Click TableWidget ** ")
        item = self.tableWidget.item(row, 2)
        file_type = item.text().lower()

        if file_type in ['doc', 'docx', 'pdf']:
            if self.total_pages[row] == 0:
                self.total_pages_setting(row)
            if self.total_pages[row] != -5:
                self.show_dialog_page(row)
            else:
                for lst in (self.input_files, self.selected_pages, self.total_pages, self.check_box):
                    lst.pop(row)

                for i in range(row + 1, self.tableWidget.rowCount()):
                    for col in range(4):
                        item_text = self.tableWidget.item(i, col).text() if self.tableWidget.item(i, col) else ""
                        self.tableWidget.setItem(i - 1, col, QTableWidgetItem(item_text))
                last_row = self.tableWidget.rowCount() - 1
                for col in range(4):
                    self.tableWidget.setItem(last_row, col, QTableWidgetItem(""))
                self.tableWidget.setRowCount(len(self.input_files))

        elif file_type in ['xlsx', 'xls', 'xlrd']:
            self.show_dialog_worksheet(row)
        else:
            self.set_info("Error", 1)


    # =============================================================================================== Dialog =========
    #Dialog
    def is_valid_input(self, input_value, max_value):
        pattern = r'^(\d+)(-\d+)?(,\s*(\d+)(-\d+)?)*$'

        if max_value == -1:
            return re.match(pattern, input_value) is not None or -2

        if re.match(pattern, input_value):
            ranges = input_value.split(',')
            for range_str in ranges:
                range_parts = range_str.split('-')
                if len(range_parts) == 2:
                    start, end = map(int, range_parts)
                else:
                    start = int(range_parts[0])
                    end = start

                if start > max_value or end > max_value:
                    return -1
                if start > end:
                    return -1

            return True
        else:
            return -2

    def handle_validation_error(self, validation_result, max_value, row):
        # 최대 값보다 큰 경우
        if validation_result == -1:
            self.critical_event(f"Invalid input: Values greater than max value {max_value} entered.")
            self.set_info("Please enter values within the valid range.")
            self.set_info("Not valid")
        # 형식이 잘못된 경우
        elif validation_result == -2:
            self.critical_event("Invalid input format. Please enter numeric values only.")
            self.set_info("Not number")

    def show_dialog_page(self, row):
        print("Show Dialog Word")
        # 1 파일 인식
        # 파일 인식이 안될 경우(-1)
        if self.total_pages[row] == -1:
            self.set_info("Direct document access required", 1)
            input_page, ok = QInputDialog.getText(self, "Word", "Please enter a page".format(self.total_pages[row]))
        # 파일 인식이 제대로 될 경우, 전체 페이지 표시
        elif self.total_pages[row] == -5:
            return
        else:
            input_page, ok = QInputDialog.getText(self, "Word", "Please enter a page [All Pages : {}]".format(self.total_pages[row]))

        # 2 입력한 값 유효성 검사
        validation_result = self.is_valid_input(input_page, self.total_pages[row])
        print("Validation Result: ", validation_result)

        # 파일 인식이 제대로 된 경우
        if self.total_pages[row] != -1:
            if validation_result == True or input_page == '':
                self.selected_pages[row] = input_page
                print(f"The page selected by the user : {input_page}")
            elif validation_result == -1 or validation_result == -2:
                self.handle_validation_error(validation_result, input_page, row)

        # 파일 인식이 잘못된 경우, 재검토 요청
        elif self.total_pages[row] == -1:
            if validation_result == True or input_page == '':
                self.selected_pages[row] = input_page
                print(f"The page selected by the user : {input_page}")
                self.set_info("If you turn the entire page, you may get an error", 1)
            elif validation_result == -2:
                self.handle_validation_error(validation_result, input_page, row)
        # 3 버튼 클릭 검사
        if ok:
            if validation_result == True:
                self.tableWidget.setItem(row, 3, QTableWidgetItem(input_page))
                self.set_info("Input Page : {}".format(input_page))
                print(f"Currently set information : {self.selected_pages}")

            else:
                self.selected_pages[row] = 0
                self.tableWidget.setItem(row, 3, QTableWidgetItem(""))  # 수정된 부분
                self.set_info("Input Page is not vaild")

    def show_dialog_worksheet(self, row):
        print("Show Dialog Excel")
        try:
            worksheet_names = self.excel_handler.extract_sheet_names(self.input_files[row])
        except Exception as e:
            print(f"Error: {e}")
            self.critical_event("Invalid file format. Only Excel files are supported for worksheets.")
            worksheet_names = []
            self.set_info("")

        input_worksheet, ok = QInputDialog.getItem(self, 'Excel', 'Please select a worksheet', worksheet_names)
        print(f"The worksheet selected by the user : {input_worksheet}")

        if ok:
            self.selected_pages[row] = input_worksheet
            self.tableWidget.setItem(row, 3, QTableWidgetItem(input_worksheet))
            self.set_info("Input WorkSheet : {}".format(input_worksheet))
            print(f"Currently set information : {self.selected_pages}")

    # ============================================================================================== QButton-five ========
    # Button-five
    def btn_move_setting(self):
        self.btn_top.released.connect(
            lambda: self.btn_top_click(self.input_files, self.selected_pages,self.total_pages,self.check_box))
        self.btn_up.released.connect(lambda: self.btn_up_click( self.input_files, self.selected_pages, self.total_pages,self.check_box))
        self.btn_down.released.connect(lambda: self.btn_down_click( self.input_files, self.selected_pages,self.total_pages,self.check_box))
        self.btn_bottom.released.connect(lambda: self.btn_bottom_click( self.input_files, self.selected_pages,self.total_pages,self.check_box))
        self.btn_delete.released.connect(lambda: self.btn_delete_click(self.input_files, self.selected_pages,self.total_pages,self.check_box))

    # 1. Top =========================================================
    def btn_top_click(self, *lists):
        row = self.currently_row
        check_num = 0
        for r in range(self.tableWidget.rowCount()):
            cell_widget = self.tableWidget.cellWidget(r, 0)
            checkbox = cell_widget.findChild(QCheckBox)  # QCheckBox를 찾습니다.
            if checkbox and checkbox.isChecked():
                checkbox.setChecked(False)  # 체크를 취소합니다.
                check_num += 1

        if check_num > 0:
            self.critical_event("The check box is valid only when you delete it")

        for lst in lists:
            lst.insert(0, lst.pop(row))

        # 현재 값 저장
        current_items = [self.tableWidget.item(row, col).text() if self.tableWidget.item(row, col) else "" for col in
                         range(4)]

        # 0~ROW 한 칸씩 아래로 이동
        for i in range(row - 1, -1, -1):
            for col in range(4):
                item_text = self.tableWidget.item(i, col).text() if self.tableWidget.item(i, col) else ""
                self.tableWidget.setItem(i + 1, col, QTableWidgetItem(item_text))

        # 현재 값을 맨 위로 이동
        for col in range(4):
            self.tableWidget.setItem(0, col, QTableWidgetItem(current_items[col]))

    # 2. Up =========================================================
    def btn_up_click(self, *lists):
        row = self.currently_row
        check_num = 0
        for r in range(self.tableWidget.rowCount()):
            cell_widget = self.tableWidget.cellWidget(r, 0)
            checkbox = cell_widget.findChild(QCheckBox)  # QCheckBox를 찾습니다.
            if checkbox and checkbox.isChecked():
                checkbox.setChecked(False)  # 체크를 취소합니다.
                check_num += 1

        if check_num > 0:
            self.critical_event("The check box is valid only when you delete it")
        elif row == 0:
            self.set_info("Current top value",1)
            return

        for lst in lists:
            lst.insert(row-1, lst.pop(row))

        current_row_items = [self.tableWidget.item(row, col).text()
                             if self.tableWidget.item(row, col) else "" for col in range(4)]
        above_row_items = [self.tableWidget.item(row - 1, col).text()
                           if self.tableWidget.item(row - 1, col) else "" for col in range(4)]

        for col in range(4):
            self.tableWidget.setItem(row, col, QTableWidgetItem(above_row_items[col]))
            self.tableWidget.setItem(row - 1, col, QTableWidgetItem(current_row_items[col]))

    # 3. Down =========================================================
    def btn_down_click(self, *lists):
        row = self.currently_row
        check_num = 0
        for r in range(self.tableWidget.rowCount()):
            cell_widget = self.tableWidget.cellWidget(r, 0)
            checkbox = cell_widget.findChild(QCheckBox)  # QCheckBox를 찾습니다.
            if checkbox and checkbox.isChecked():
                checkbox.setChecked(False)  # 체크를 취소합니다.
                check_num += 1

        if check_num > 0:
            self.critical_event("The check box is valid only when you delete it")
        elif row == self.tableWidget.rowCount() - 1:
            self.set_info("Current Lowest value",1)
            return
        for lst in lists:
            lst.insert(row + 1, lst.pop(row))

        current_row_items = [self.tableWidget.item(row, col).text()
                             if self.tableWidget.item(row, col) else "" for col in range(4)]
        below_row_items = [self.tableWidget.item(row + 1, col).text()
                           if self.tableWidget.item(row + 1, col) else "" for col in range(4)]

        for col in range(4):
            self.tableWidget.setItem(row, col, QTableWidgetItem(below_row_items[col]))
            self.tableWidget.setItem(row + 1, col, QTableWidgetItem(current_row_items[col]))

    # 4. Bottom =========================================================
    def btn_bottom_click(self, *lists):
        row = self.currently_row

        check_num = 0
        for r in range(self.tableWidget.rowCount()):
            cell_widget = self.tableWidget.cellWidget(r, 0)
            checkbox = cell_widget.findChild(QCheckBox)  # QCheckBox를 찾습니다.
            if checkbox and checkbox.isChecked():
                checkbox.setChecked(False)  # 체크를 취소합니다.
                check_num += 1

        if check_num > 0:
            self.critical_event("The check box is valid only when you delete it")

        for lst in lists:
            if row in lst:
                lst.append(lst.pop(lst.index(row)))

        # 현재 값 저장
        current_items = [self.tableWidget.item(row, col).text() if self.tableWidget.item(row, col) else "" for col in
                         range(4)]

        # row부터 마지막 행까지 한 칸씩 위로 이동
        for i in range(row, self.tableWidget.rowCount() - 1):
            for col in range(4):
                item_text = self.tableWidget.item(i + 1, col).text() if self.tableWidget.item(i + 1, col) else ""
                self.tableWidget.setItem(i, col, QTableWidgetItem(item_text))

        # 현재 값을 맨 마지막 행에 세팅
        last_row = self.tableWidget.rowCount() - 1
        for col in range(4):
            self.tableWidget.setItem(last_row, col, QTableWidgetItem(current_items[col]))

    # 5. Delete =========================================================
    def btn_delete_click(self, *lists):
        row = self.currently_row
        row_to_delete = []

        for r in range(self.tableWidget.rowCount()):
            cell_widget = self.tableWidget.cellWidget(r, 0)

            if cell_widget is not None:  # cell_widget이 None인지 확인
                checkbox = cell_widget.findChild(QCheckBox)  # QCheckBox를 찾습니다.
                if checkbox and checkbox.isChecked():  # isChecked()로 체크 여부를 확인합니다.
                    row_to_delete.append(r)
            else:
                print(f"Row {r} has no widget in the first cell.")  # 디버깅을 위한 메시지

        if len(row_to_delete) > 0:
            for r in reversed(row_to_delete):
                self.tableWidget.removeRow(r)
                for lst in lists:
                    lst.pop(r)

            return

        # Check 상태가 아예 없을 경우, 현재 선택된 행만 삭제
        for lst in lists:
            lst.pop(row)

        for i in range(row + 1, self.tableWidget.rowCount()):
            for col in range(4):
                item_text = self.tableWidget.item(i, col).text() if self.tableWidget.item(i, col) else ""
                self.tableWidget.setItem(i - 1, col, QTableWidgetItem(item_text))
        last_row = self.tableWidget.rowCount() - 1
        for col in range(4):
            self.tableWidget.setItem(last_row, col, QTableWidgetItem(""))
        self.tableWidget.setRowCount(len(self.input_files))

    # ============================================================================================== QMessageBox ========
    #QMessageBox
    def critical_event(self, msg) :
        QMessageBox.critical(self,'Critical',msg)
    # ============================================================================================== QLabel ========
    #QLabel
    def set_info(self, msg, state = 0):
        if state == 1:
            self.lbLocation.setText('<span style="color : red">{}</span>'.format(msg))
        else:
            self.lbLocation.setText(msg)
# ============================================================================================== Main ========
#Main

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()

    '''
    = tkinter =
class App:
    def __init__(self):
        self.gui = handle_gui.Gui()

    def run(self):
        self.gui.initialize_ui()
        self.gui.root.mainloop()

if __name__ == '__main__':
    app = App()
    app.run()
'''
