import os
import xlsxwriter

EXCEL_FILE_NAME = 'demo.xlsx'
EXCEL_LIST_SHEET_NAME = [
    '01'
]
EVIDENCE_PATH = '.'

# Create an new Excel file and add a worksheet.
# workbook = xlsxwriter.Workbook('demo.xlsx')
# worksheet = workbook.add_worksheet()

# # Widen the first column to make the text clearer.
# worksheet.set_column('A:A', 20)

# # Add a bold format to use to highlight cells.
# bold = workbook.add_format({'bold': True})

# # Write some simple text.
# worksheet.write('A1', 'Hello')

# # Text with formatting.
# worksheet.write('A2', 'World', bold)

# # Write some numbers, with row/column notation.
# worksheet.write(2, 0, 123)
# worksheet.write(3, 0, 123.456)

# # Insert an image.
# worksheet.insert_image('B5', 'logo.jpg')

# workbook.close()

def get_evidence_file(path):
    all_file = os.listdir(path)
    return [file for file in all_file if file.endswith('.png') or file.endswith('.jpg')]

def get_testcase_text(list_file):
    dict_testcase = {}
    for file in list_file:
        testcase_name = file.split('_')[0]
        if testcase_name not in dict_testcase.keys():
            dict_testcase[testcase_name] = [file]
        else:
            dict_testcase[testcase_name].append(file)
    
    return dict_testcase

if __name__ == '__main__':
    #region 0. Khởi tạo dữ liệu sang kiểu của python
    evidence_images = get_evidence_file(EVIDENCE_PATH)
    dict_testcase = get_testcase_text(evidence_images)
    print(dict_testcase)
    #endregion
    # TODO 1. Tạo mới file, tạo mới các sheet
    workbook = xlsxwriter.Workbook(EXCEL_FILE_NAME)
    for worksheet_name in EXCEL_LIST_SHEET_NAME:
        worksheet = workbook.add_worksheet(worksheet_name)
        worksheet.insert_image('A5', '01-001_1.jpg')

        last_row = worksheet.dim_rowmax
        print(last_row)
    # TODO 2. Mở các file và các sheet tương ứng

    # TODO 3. Dán data vào