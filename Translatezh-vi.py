import json
import time
import re
from openpyxl import load_workbook

# ====== CẤU HÌNH CƠ BẢN ======
FILE_PATH = 'data.xlsx'
KEY_FILE_PATH = 'key.json'
LANG_SOURCE = 'vi'
LANG_TARGET = 'zh'          # 'zh' hoặc 'zh-CN' tùy bạn
OUTPUT_SUFFIX = "_Dich_CN"
LOCATION = 'asia-southeast1'  # 'us-central1' cũng được

# ====== KHỞI TẠO VERTEX AI (GEMINI) ======
from google.oauth2 import service_account
import vertexai
from vertexai.generative_models import GenerativeModel

try:
    with open(KEY_FILE_PATH, 'r', encoding='utf-8') as f:
        sa_info = json.load(f)
    PROJECT_ID = sa_info.get("project_id")
    if not PROJECT_ID:
        raise ValueError("Không tìm thấy 'project_id' trong tệp service account.")

    credentials = service_account.Credentials.from_service_account_file(KEY_FILE_PATH)
    vertexai.init(project=PROJECT_ID, location=LOCATION, credentials=credentials)
    model = GenerativeModel("gemini-2.5-pro")
    print(f"✅ Đã khởi tạo Gemini (Vertex AI) cho project: {PROJECT_ID} tại {LOCATION}")
except Exception as e:
    print(f"❌ LỖI KHỞI TẠO GEMINI/Vertex AI: {e}")
    raise SystemExit(1) from e

# Ánh xạ tên sheet gốc -> sheet mới
SHEET_NAME_MAP = {}

# ====== HÀM DỊCH BẰNG GEMINI ======

def translate_batch_gemini(texts, lang_src, lang_tgt, max_retries=3, sleep_seconds=2):
    """
    Dịch danh sách 'texts' bằng Gemini. Trả về list các chuỗi đã dịch, cùng thứ tự.
    Để gom batch an toàn, ta yêu cầu model trả về JSON array theo thứ tự.
    """
    if not texts:
        return []

    # Prompt yêu cầu trả JSON array, giữ thứ tự, không thêm chú thích
    system_instruction = (
        "Bạn là chuyên gia dịch thuật từ tiếng Trung giản thể sang tiếng Việt chuyên dành cho các dự án xây dựng và thiết kế. Hãy dịch chính xác với ngữ cảnh và đúng từ chuyên ngành.\n"
        f"Ngôn ngữ nguồn: {lang_src}\n"
        f"Ngôn ngữ đích: {lang_tgt}\n"
        "- Chỉ dịch phần văn bản; nếu là công thức Excel hoặc tham chiếu ô/sheet thì bỏ qua (nhưng ở đây đầu vào đã lọc công thức rồi).\n"
        "- Giữ nguyên số, mã SKU, ký hiệu đặc biệt nếu không cần dịch.\n"
        "- Bảo toàn xuống dòng và khoảng trắng quan trọng.\n"
        "- Trả KẾT QUẢ DUY NHẤT là một JSON array các chuỗi, theo đúng thứ tự input, không thêm giải thích. Ví dụ: bạn nhận được [\"瓷砖楼梯\"], bạn trả lời \"cầu thang đá men\"."
    )

    # Để chắc ăn mảng JSON, mình đóng gói input thành JSON array string
    payload = json.dumps(texts, ensure_ascii=False)

    content = [
        {"role": "user", "parts": [
            f"{system_instruction}\n\nĐây là danh sách câu cần dịch (JSON array):\n{payload}\n\n"
            "Hãy trả về *chỉ* một JSON array các bản dịch, theo thứ tự tương ứng."
        ]}
    ]

    for attempt in range(1, max_retries + 1):
        try:
            resp = model.generate_content(content, generation_config={
                "temperature": 0.2,
                "max_output_tokens": 2048,
                # Bạn có thể set thêm "response_mime_type": "application/json" (SDK mới hỗ trợ)
            })
            text = resp.text.strip()
            # Cố gắng parse JSON trực tiếp
            # Một số trường hợp model có thể bao code block, ta làm sạch nhẹ
            if text.startswith("```"):
                # loại bỏ ```json ... ```
                text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.DOTALL)

            out = json.loads(text)
            if not isinstance(out, list):
                raise ValueError("Phản hồi không phải JSON array.")
            # Đảm bảo độ dài khớp:
            if len(out) != len(texts):
                raise ValueError(f"Số phần tử phản hồi ({len(out)}) không khớp input ({len(texts)}).")
            # Chuẩn hóa về str
            return [str(x) if x is not None else "" for x in out]

        except Exception as e:
            print(f"⚠️ Lỗi dịch batch (lần {attempt}/{max_retries}): {e}")
            if attempt < max_retries:
                time.sleep(sleep_seconds)
            else:
                # Thử fallback: dịch từng câu để không mất kết quả toàn batch
                print("🔁 Fallback: dịch từng câu…")
                results = []
                for t in texts:
                    try:
                        single = translate_batch_gemini([t], lang_src, lang_tgt, max_retries=1, sleep_seconds=1)
                        results.append(single[0] if single else "")
                    except Exception:
                        results.append("")
                return results

def update_formula_references(sheet, sheet_map, output_suffix):
    """Cập nhật tham chiếu công thức trong sheet để trỏ đến sheet đích mới."""
    for row in sheet.iter_rows(values_only=False):
        for cell in row:
            if cell.data_type == 'f' and cell.value:
                formula = cell.value
                for old_name, new_name in sheet_map.items():
                    # khớp cả 'Sheet Name'!A1 và SheetName!A1
                    pattern = re.compile(r"([']?)" + re.escape(old_name) + r"([']?)!")
                    if pattern.search(formula):
                        new_formula = pattern.sub(rf"\1{new_name}\2!", formula)
                        cell.value = new_formula
                        formula = new_formula

def translate_and_copy_sheet(workbook, source_sheet_name):
    """Duyệt ô text, dịch bằng Gemini, và ghi vào sheet mới (bản sao giữ định dạng + công thức)."""
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

    original_texts = [item['text'] for item in cells_to_translate]

    # Bạn có thể chia batch nhỏ để ổn định hơn (ví dụ 100 mục/lượt)
    BATCH_SIZE = 100
    translated_texts_all = []
    try:
        for i in range(0, len(original_texts), BATCH_SIZE):
            batch = original_texts[i:i+BATCH_SIZE]
            translated_batch = translate_batch_gemini(batch, LANG_SOURCE, LANG_TARGET)
            translated_texts_all.extend(translated_batch)
        print(f"✅ Đã dịch thành công {len(translated_texts_all)} đoạn text bằng Gemini.")
    except Exception as e:
        print(f"❌ LỖI DỊCH GEMINI: {e}")
        return

    # Ghi đè kết quả vào sheet mới
    for i, item in enumerate(cells_to_translate):
        target_cell = new_sheet.cell(row=item['row'], column=item['col'])
        target_cell.value = translated_texts_all[i]

    print(f"--- Hoàn thành ghi đè Sheet: {source_sheet_name} ---")

# ====== CHƯƠNG TRÌNH CHÍNH ======
try:
    wb = load_workbook(FILE_PATH)
    original_sheet_names = [name for name in wb.sheetnames if OUTPUT_SUFFIX not in name]

    # Bước 1: Dịch & tạo sheet mới
    for sheet_name in original_sheet_names:
        translate_and_copy_sheet(wb, sheet_name)

    # Bước 2: Cập nhật tham chiếu công thức giữa các sheet
    print("\n[BƯỚC 2] Bắt đầu cập nhật tham chiếu công thức…")
    for sheet_name in original_sheet_names:
        new_sheet_name = SHEET_NAME_MAP[sheet_name]
        new_sheet = wb[new_sheet_name]
        update_formula_references(new_sheet, SHEET_NAME_MAP, OUTPUT_SUFFIX)
        print(f"✅ Đã cập nhật tham chiếu cho sheet: {new_sheet_name}")

    # Bước 3: Lưu file mới
    OUTPUT_FILE_PATH = FILE_PATH.replace(".xlsx", "_DICH_CN.xlsx")
    wb.save(OUTPUT_FILE_PATH)
    print(f"\n✅ THÀNH CÔNG! File đã dịch được lưu tại: {OUTPUT_FILE_PATH}")

except Exception as e:
    print(f"❌ Lỗi tổng quát khi xử lý file: {e}")
    print("Hãy kiểm tra xem file Excel có đang mở/khoá, KEY_FILE_PATH đúng và service account có quyền Vertex AI.")
