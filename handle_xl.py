# handle_xl.py
#
# Handle the excel file and prep data to send to the model
#
# @author: Dung Tran, Hien Nguyen
# @date: 2025-10-17

import openpyxl as excel
from openpyxl import Workbook
import json
import os
import sys
from inference import translate_data
import regex as re
import xlrd
from time import sleep

# Constants
OUTPUT_SUFFIX = "_Dich_CN"

# Mapping of original sheet name to new sheet name
SHEET_NAME_MAP = {}


def convert_xls_to_xlsx_preserve_format(xls_file_path, xlsx_file_path=None):
    """
    Convert .xls to .xlsx while preserving more formatting
    """
    if xlsx_file_path is None:
        xlsx_file_path = xls_file_path.replace('.xls', '.xlsx')
    
    try:
        # Open the .xls file
        xls_workbook = xlrd.open_workbook(xls_file_path)
        
        # Create new .xlsx workbook
        xlsx_workbook = Workbook()
        
        # Remove default sheet
        xlsx_workbook.remove(xlsx_workbook.active)
        
        # Copy each sheet
        for sheet_name in xls_workbook.sheet_names():
            xls_sheet = xls_workbook.sheet_by_name(sheet_name)
            xlsx_sheet = xlsx_workbook.create_sheet(title=sheet_name)
            
            # Copy data
            for row in range(xls_sheet.nrows):
                for col in range(xls_sheet.ncols):
                    cell_value = xls_sheet.cell_value(row, col)
                    xlsx_sheet.cell(row=row+1, column=col+1, value=cell_value)
        
        # Save the new workbook
        xlsx_workbook.save(xlsx_file_path)
        print(f"✅ Successfully converted {xls_file_path} to {xlsx_file_path}")
        return xlsx_file_path
        
    except Exception as e:
        print(f"❌ Error converting file: {e}")
        return None


def translate_excel_step1(workbook, source_sheet_name):
    """Duyệt ô text, dịch qua OpenAI API, và ghi vào sheet mới (bản sao giữ định dạng + công thức)."""
    source_sheet = workbook[source_sheet_name]
    new_sheet_name = f"{source_sheet_name}{OUTPUT_SUFFIX}"
    SHEET_NAME_MAP[source_sheet_name] = new_sheet_name

    if new_sheet_name in workbook.sheetnames:
        print(f"⚠️ Sheet đích '{new_sheet_name}' đã tồn tại. Dữ liệu sẽ được GHI ĐÈ.")
        new_sheet = workbook[new_sheet_name]
    else:
        new_sheet = workbook.copy_worksheet(source_sheet)
        new_sheet.title = new_sheet_name

    print(f"\n--- Bắt đầu dịch Sheet: {source_sheet_name} ---")

    # Thu thập các ô text (không phải công thức)
    cells_to_translate = []
    for row in source_sheet.iter_rows(values_only=False):
        for cell in row:
            if (isinstance(cell.value, str) and cell.value.strip() and cell.data_type != 'f'):
                cells_to_translate.append({
                    'text': cell.value,
                    'row': cell.row,
                    'col': cell.column
                })
    print("Done loading data!")

    original_texts = [item['text'] for item in cells_to_translate]

    # Bạn có thể chia batch nhỏ để ổn định hơn (ví dụ 100 mục/lượt)
    BATCH_SIZE = 100
    threshold = 5
    translated_texts_all = []
    try:
        for i in range(0, len(original_texts), BATCH_SIZE):
            batch = original_texts[i:min(i+BATCH_SIZE, len(original_texts))]
            count = 0
            translated_batch = []
            # Safety check to make sure we have translated all entries
            while len(translated_batch) < len(batch) and count < threshold:
                translated_batch = translate_data(batch)
                print(len(translated_batch), len(batch))
                translated_batch = translated_batch[:len(batch)]
                count += 1
                sleep(5)
            
            if len(translated_batch) != len(batch):
                print("Failed to translate at some point!")
                sys.exit(1)
            
            translated_texts_all.extend([x[1] for x in translated_batch])
            print("Translating!")
        print(f"✅ Đã dịch thành công {len(translated_texts_all)}")
    except Exception as e:
        print(f"❌ LỖI DỊCH: {e}")
        return

    # Ghi đè kết quả vào sheet mới
    for i, item in enumerate(cells_to_translate):
        target_cell = new_sheet.cell(row=item['row'], column=item['col'])
        target_cell.value = translated_texts_all[i]

    print(f"--- Hoàn thành ghi đè Sheet: {source_sheet_name} ---")


def update_formula_references(sheet, sheet_map, output_suffix):
    """Update the formula references in the sheet to point to the new sheet."""
    for row in sheet.iter_rows(values_only=False):
        for cell in row:
            if cell.data_type == 'f' and cell.value:
                formula = cell.value
                for old_name, new_name in sheet_map.items():
                    # match both 'Sheet Name'!A1 and SheetName!A1
                    pattern = re.compile(r"([']?)" + re.escape(old_name) + r"([']?)!")
                    if pattern.search(formula):
                        new_formula = pattern.sub(rf"\1{new_name}\2!", formula)
                        cell.value = new_formula
                        formula = new_formula


def full_translate(data_file):
    wb = excel.load_workbook(data_file)
    original_sheet_names = [name for name in wb.sheetnames if OUTPUT_SUFFIX not in name]

    # Bước 1: Dịch & tạo sheet mới
    for sheet_name in original_sheet_names:
        translate_excel_step1(wb, sheet_name)

    # Bước 2: Cập nhật tham chiếu công thức giữa các sheet
    print("\n[BƯỚC 2] Bắt đầu cập nhật tham chiếu công thức…")
    for sheet_name in original_sheet_names:
        new_sheet_name = SHEET_NAME_MAP[sheet_name]
        new_sheet = wb[new_sheet_name]
        update_formula_references(new_sheet, SHEET_NAME_MAP, OUTPUT_SUFFIX)
        print(f"✅ Đã cập nhật tham chiếu cho sheet: {new_sheet_name}")

    # Bước 3: Lưu file mới
    OUTPUT_FILE_PATH = data_file.replace(".xlsx", "_DICH_CN.xlsx")
    wb.save(OUTPUT_FILE_PATH)
    print(f"\n✅ THÀNH CÔNG! File đã dịch được lưu tại: {OUTPUT_FILE_PATH}")


if __name__ == "__main__":
    data_file = "./data1.xlsx"
    if data_file.endswith(".xls"):
        data_file = convert_xls_to_xlsx_preserve_format(data_file)
    
    try:
        wb = excel.load_workbook(data_file)
    except Exception as e:
        print(e)
        sys.exit(1)
    full_translate(data_file)