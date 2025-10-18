# zh-vi-translate-gemini.py
#
# Translate with Gemini API
#
# @author: Hien Nguyen, Dung Tran
# @date: 18/10/2025

import os
import re
import time
import json
import ast
import html
import random
from typing import List

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

import google.generativeai as genai

# --- THIẾT LẬP FILE VÀ NGÔN NGỮ ---
FILE_PATH = 'data1.xlsx'
LANG_SOURCE = 'zh'
LANG_TARGET = 'vi'
OUTPUT_SUFFIX = "_Dich_CN"
BATCH_SIZE = 15             # số ô/đợt gọi API
SLEEP_BETWEEN_BATCHES = 0.5  # giãn nhịp tránh throttling
GEMINI_MODEL = "gemini-2.0-flash"  # có thể đổi sang "gemini-1.5-pro" nếu cần chất lượng cao
MAX_RETRIES = 5
# -----------------------------------

# --- Load .env và khởi tạo Gemini ---
load_dotenv()
api_key = os.getenv("GOOGLE_API_KEY", "").strip()
if not api_key:
    raise ValueError("API key GOOGLE_API_KEY chưa được thiết lập (file .env hoặc biến môi trường).")

genai.configure(api_key=api_key)
model = genai.GenerativeModel(GEMINI_MODEL)
print(f"✅ Đã khởi tạo Gemini model: {GEMINI_MODEL}")

SHEET_NAME_MAP = {}

def safe_sheet_title(base: str) -> str:
    """Tạo tên sheet mới an toàn (<=31 ký tự, loại bỏ ký tự cấm)."""
    cleaned = re.sub(r'[:\\/?*\[\]]', '_', base)
    return cleaned[:31] if len(cleaned) > 31 else cleaned

def excel_quote_sheet(name: str) -> str:
    """Trả về tên sheet theo quy tắc Excel (quote khi cần)."""
    need_quote = bool(re.search(r"[ \-+*/^&=<>,;:@\[\]]", name)) or "'" in name
    if "'" in name:
        name = name.replace("'", "''")
    return f"'{name}'" if need_quote else name

def _gemini_translate_call(texts: List[str], *, src: str, tgt: str, max_retries: int = MAX_RETRIES) -> List[str]:
    """
    Gọi Gemini để dịch danh sách texts. Trả về list cùng kích thước, giữ thứ tự.
    Bắt buộc Gemini trả về JSON array để dễ parse, tránh lẫn text phụ.
    """
    if not texts:
        return []

    system_prompt = f"""Bạn là chuyên gia dịch thuật từ tiếng Trung giản thể sang tiếng Việt chuyên dành cho các dự án xây dựng và thiết kế. Hãy dịch chính xác với ngữ cảnh và đúng từ chuyên ngành.\n"
        Ngôn ngữ nguồn: tiếng Trung giản thể\n
        Ngôn ngữ đích: tiếng Việt\n
        - Chỉ dịch phần văn bản; nếu là công thức Excel hoặc tham chiếu ô/sheet thì bỏ qua.\n
        - Giữ nguyên số, mã SKU, ký hiệu đặc biệt nếu không cần dịch.\n
        - Bảo toàn xuống dòng và khoảng trắng quan trọng.\n
        - Trả KẾT QUẢ DUY NHẤT là một JSON array các chuỗi chứa từ gốc ở vị trí thứ nhất và từ được dịch ở vị trí thứ 2. Các từ được trả theo đúng thứ tự input, không thêm giải thích. Ví dụ: bạn nhận được [\"瓷砖楼梯\", "1"], bạn trả lời [(\'瓷砖楼梯\', \'cầu thang đá men\'), (\'1\', \'1\')].
        """ + "Dữ liệu cần dịch: [" + ",".join(texts) + "]"
    user_msg = (
        system_prompt
    )

    for attempt in range(1, max_retries + 1):
        try:
            resp = model.generate_content(user_msg)
            raw = (resp.text or "").strip()
            # Loại bỏ ```json ... ``` nếu có
            raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.IGNORECASE)
            out = ast.literal_eval(raw)
            print(out)
            if not isinstance(out, list) or any(not isinstance(x, list) for x in out):
                raise ValueError("Phản hồi không phải JSON array of strings.")
            if len(out) != len(texts):
                raise ValueError(f"Kích thước kết quả ({len(out)}) != đầu vào ({len(texts)}).")
            return [html.unescape(s) for s in out]
        except Exception as e:
            if attempt == max_retries:
                raise
            sleep_s = 0.8 * attempt + random.uniform(0, 0.4)
            print(f"⚠️ Lỗi GEMINI (attempt {attempt}/{max_retries}): {e}. Thử lại sau {sleep_s:.2f}s...")
            time.sleep(sleep_s)

def translate_batch(texts: List[str]) -> List[str]:
    """Dịch danh sách texts (đã chia batch bên ngoài) bằng Gemini."""
    return _gemini_translate_call(texts, src=LANG_SOURCE, tgt=LANG_TARGET)

def update_formula_references(sheet, sheet_map):
    """Cập nhật tham chiếu công thức cho sheet đã sao chép."""
    escaped_map = {}
    for old_name, new_name in sheet_map.items():
        escaped_old = excel_quote_sheet(old_name)
        escaped_new = excel_quote_sheet(new_name)
        escaped_map[escaped_old] = escaped_new

    for row in sheet.iter_rows(values_only=False):
        for cell in row:
            if cell.data_type == 'f' and cell.value:
                formula = cell.value
                for old_q, new_q in escaped_map.items():
                    pattern = re.compile(rf"(?<!\[){re.escape(old_q)}!")
                    formula = pattern.sub(f"{new_q}!", formula)
                cell.value = formula

def translate_and_copy_sheet(workbook, source_sheet_name):
    source_sheet = workbook[source_sheet_name]
    new_sheet_name_raw = f"{source_sheet_name}{OUTPUT_SUFFIX}"
    new_sheet_name = safe_sheet_title(new_sheet_name_raw)

    SHEET_NAME_MAP[source_sheet_name] = new_sheet_name

    if new_sheet_name in workbook.sheetnames:
        print(f"⚠️ Sheet đích '{new_sheet_name}' đã tồn tại. Dữ liệu sẽ được GHI ĐÈ.")
        new_sheet = workbook[new_sheet_name]
    else:
        new_sheet = workbook.copy_worksheet(source_sheet)
        new_sheet.title = new_sheet_name

    print(f"\n--- Bắt đầu dịch Sheet: {source_sheet_name} -> {new_sheet_name} ---")

    # Thu thập text cần dịch
    cells_to_translate = []
    seen = {}  # cache: text -> translated_text
    for row in source_sheet.iter_rows(values_only=False):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            val = cell.value
            if isinstance(val, str) and val.strip() and cell.data_type != 'f':
                cells_to_translate.append({'text': val, 'row': cell.row, 'col': cell.column})

    # Chạy batch
    idx = 0
    total = len(cells_to_translate)
    while idx < total:
        chunk = cells_to_translate[idx:idx+BATCH_SIZE]

        # Chuẩn bị input + cache
        to_send = []
        map_indices = []  # ('cache', text) hoặc ('send', index-in-to_send)
        for i, item in enumerate(chunk):
            t = item['text']
            if t in seen:
                map_indices.append(('cache', t))
            else:
                map_indices.append(('send', len(to_send)))
                to_send.append(t)

        translations = []
        if to_send:
            try:
                translations = [x[1] for x in translate_batch(to_send)]
                for src_text, dst_text in zip(to_send, translations):
                    seen[src_text] = dst_text
            except Exception as e:
                print(f"❌ LỖI API GEMINI (batch offset {idx}): {e}")
                # Bỏ qua batch này để tiếp tục vòng lặp
                idx += BATCH_SIZE
                time.sleep(SLEEP_BETWEEN_BATCHES)
                continue

        # Ghi kết quả
        for i, item in enumerate(chunk):
            if map_indices[i][0] == 'cache':
                translated = seen[map_indices[i][1]]
            else:
                translated = translations[map_indices[i][1]]
            target_cell = new_sheet.cell(row=item['row'], column=item['col'])
            if isinstance(target_cell, MergedCell):
                continue
            target_cell.value = translated

        print(f"✅ Đã dịch & ghi {min(idx + len(chunk), total)} / {total} ô.")
        idx += BATCH_SIZE
        time.sleep(SLEEP_BETWEEN_BATCHES)

    print(f"--- Hoàn thành ghi đè Sheet: {source_sheet_name} ---")

# --- CHƯƠNG TRÌNH CHÍNH ---
try:
    wb = load_workbook(FILE_PATH)
    original_sheet_names = [name for name in wb.sheetnames if OUTPUT_SUFFIX not in name]

    # Bước 1: Dịch và tạo các sheet mới
    for sheet_name in original_sheet_names:
        translate_and_copy_sheet(wb, sheet_name)

    # Bước 2: Cập nhật công thức tham chiếu giữa các sheet
    print("\n[BƯỚC 2] Bắt đầu cập nhật tham chiếu công thức...")
    for sheet_name in original_sheet_names:
        new_sheet_name = SHEET_NAME_MAP[sheet_name]
        new_sheet = wb[new_sheet_name]
        update_formula_references(new_sheet, SHEET_NAME_MAP)
        print(f"✅ Đã cập nhật tham chiếu cho sheet: {new_sheet_name}")

    # Lưu Workbook mới
    OUTPUT_FILE_PATH = FILE_PATH.replace(".xlsx", "_DICH_CN.xlsx")
    wb.save(OUTPUT_FILE_PATH)
    print(f"\n✅ THÀNH CÔNG! File đã dịch được lưu tại: {OUTPUT_FILE_PATH}")

except Exception as e:
    print(f"❌ Xảy ra lỗi tổng quát khi xử lý file: {e}.")
    print("Kiểm tra: file Excel có đang mở, quyền ghi, và GOOGLE_API_KEY trong .env.")
