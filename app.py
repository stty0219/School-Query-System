from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
import logging
from werkzeug.utils import secure_filename # 用於安全地處理檔案名稱

app = Flask(__name__)

# --- 常數與設定 ---
EXCEL_FILE_PATH = 'data.xlsx'  # 預設/當前使用的 Excel 檔案名稱
GSAT_ID_COLUMN = '學測號碼'
DEPARTMENT_COLUMN = '校系名稱'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'} # 允許的檔案副檔名

# --- Logging 設定 ---
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# --- 全域變數 ---
PROCESSED_EXCEL_DATA = None
EXCEL_LOAD_ERROR_MSG = None
CURRENT_DATA_FILENAME = EXCEL_FILE_PATH # 用於追蹤目前資料來源的檔案名

# --- 輔助函數 ---
def _clean_gsat_id_value(id_val):
    if pd.isna(id_val):
        return ""
    if isinstance(id_val, float):
        return str(int(id_val))
    return str(id_val).strip()

def _clean_input_gsat_id(gsat_id_str):
    if not gsat_id_str:
        return ""
    cleaned_id = str(gsat_id_str).strip()
    if cleaned_id.endswith(".0"):
        cleaned_id = cleaned_id[:-2]
    return cleaned_id

def allowed_file_type(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- 資料載入與預處理 ---
def load_and_process_excel(source_filename=EXCEL_FILE_PATH):
    """
    載入並預處理指定的 Excel 檔案。
    更新全域變數 PROCESSED_EXCEL_DATA 和 EXCEL_LOAD_ERROR_MSG。
    """
    global PROCESSED_EXCEL_DATA, EXCEL_LOAD_ERROR_MSG, CURRENT_DATA_FILENAME
    try:
        base_dir = os.path.abspath(os.path.dirname(__file__))
        # source_filename 參數允許我們指定要載入的檔案，預設是 EXCEL_FILE_PATH
        # 但實際上，我們的策略是上傳時覆寫 EXCEL_FILE_PATH，所以這裡的 source_filename 主要是為了清晰
        excel_path_to_load = os.path.join(base_dir, source_filename)

        logging.info(f"嘗試載入 Excel 檔案：{excel_path_to_load} (來源: {source_filename})")

        if not os.path.exists(excel_path_to_load):
            # 如果是因為上傳了一個新檔名，但儲存後要以固定檔名 EXCEL_FILE_PATH 讀取
            # 此處邏輯是假設 source_filename 就是 EXCEL_FILE_PATH
            raise FileNotFoundError(f"檔案 '{excel_path_to_load}' 不存在。")

        df = pd.read_excel(excel_path_to_load)
        logging.info(f"Excel 檔案 '{source_filename}' 原始資料載入成功。")

        if GSAT_ID_COLUMN not in df.columns:
            raise ValueError(f"Excel 檔案 '{source_filename}' 中找不到 '{GSAT_ID_COLUMN}' 欄位。")
        if DEPARTMENT_COLUMN not in df.columns:
            raise ValueError(f"Excel 檔案 '{source_filename}' 中找不到 '{DEPARTMENT_COLUMN}' 欄位。")

        df[GSAT_ID_COLUMN] = df[GSAT_ID_COLUMN].apply(_clean_gsat_id_value)
        logging.info(f"欄位 '{GSAT_ID_COLUMN}' 清理完成。")
        df[DEPARTMENT_COLUMN] = df[DEPARTMENT_COLUMN].fillna('').astype(str).str.strip()
        logging.info(f"欄位 '{DEPARTMENT_COLUMN}' 清理完成。")

        PROCESSED_EXCEL_DATA = df
        CURRENT_DATA_FILENAME = source_filename # 更新當前使用的檔案名 (如果上傳時用了不同名字)
                                                # 在我們的覆寫策略下，它總是 EXCEL_FILE_PATH
        EXCEL_LOAD_ERROR_MSG = None
        logging.info(f"Excel 資料 ('{source_filename}') 已成功載入並預處理完畢。")
        return True # 表示成功

    except FileNotFoundError as fnf_err:
        EXCEL_LOAD_ERROR_MSG = f"載入錯誤：找不到 Excel 檔案 '{source_filename}' ({fnf_err})"
        logging.error(EXCEL_LOAD_ERROR_MSG)
    except ValueError as ve:
        EXCEL_LOAD_ERROR_MSG = f"處理錯誤：處理 Excel ('{source_filename}') 時發生錯誤 - {ve}"
        logging.error(EXCEL_LOAD_ERROR_MSG)
    except Exception as e:
        EXCEL_LOAD_ERROR_MSG = f"未知錯誤：載入 Excel ('{source_filename}') 時發生未預期錯誤 - {e}"
        logging.error(EXCEL_LOAD_ERROR_MSG, exc_info=True)
    
    PROCESSED_EXCEL_DATA = None # 確保出錯時資料為 None
    return False # 表示失敗

# --- Flask 路由 ---
@app.route('/')
def index():
    logging.debug("請求首頁 '/'，渲染 index.html。")
    # 將當前資料來源的狀態傳給前端
    data_status = {
        'filename': CURRENT_DATA_FILENAME if PROCESSED_EXCEL_DATA is not None else "無 (請上傳或檢查錯誤)",
        'error': EXCEL_LOAD_ERROR_MSG if PROCESSED_EXCEL_DATA is None else None
    }
    return render_template('index.html', data_status=data_status)

@app.route('/upload_excel', methods=['POST'])
def upload_excel_file_route():
    global EXCEL_LOAD_ERROR_MSG, CURRENT_DATA_FILENAME # 允許修改
    if 'excelFile' not in request.files:
        logging.warning("上傳請求中沒有 'excelFile' 部分。")
        return jsonify({'error': '請求中沒有檔案部分 (No file part in the request)'}), 400
    
    file = request.files['excelFile']
    if file.filename == '':
        logging.warning("上傳請求中未選取任何檔案。")
        return jsonify({'error': '沒有選取檔案 (No selected file)'}), 400
    
    if file and allowed_file_type(file.filename):
        # 使用 secure_filename 獲取一個安全的原始檔名 (主要用於顯示，我們仍用固定路徑儲存)
        original_filename = secure_filename(file.filename)
        
        base_dir = os.path.abspath(os.path.dirname(__file__))
        # 上傳的檔案將覆寫 EXCEL_FILE_PATH 所指向的檔案
        save_path = os.path.join(base_dir, EXCEL_FILE_PATH)

        try:
            file.save(save_path)
            logging.info(f"使用者上傳的檔案 '{original_filename}' 已儲存至：{save_path}")
            
            # 檔案儲存後，重新載入並處理這個檔案 (它現在位於 EXCEL_FILE_PATH)
            if load_and_process_excel(source_filename=EXCEL_FILE_PATH): # 明確指定以 EXCEL_FILE_PATH 載入
                logging.info(f"新上傳的 Excel 資料 (來自 '{original_filename}') 已成功載入並處理。")
                CURRENT_DATA_FILENAME = original_filename # 更新顯示的檔案名為使用者上傳的檔名
                return jsonify({
                    'message': f"檔案 '{original_filename}' 上傳成功並已更新資料來源。",
                    'processed_filename': original_filename
                }), 200
            else:
                # load_and_process_excel 內部已設定 EXCEL_LOAD_ERROR_MSG
                logging.error(f"檔案 '{original_filename}' 上傳後，處理資料時發生錯誤: {EXCEL_LOAD_ERROR_MSG}")
                return jsonify({'error': f"檔案上傳成功，但處理資料時發生錯誤：{EXCEL_LOAD_ERROR_MSG}"}), 500
        except Exception as e:
            logging.error(f"儲存上傳的檔案 '{original_filename}' 時發生錯誤：{e}", exc_info=True)
            EXCEL_LOAD_ERROR_MSG = f'儲存檔案時發生錯誤：{e}' # 更新全域錯誤訊息
            return jsonify({'error': EXCEL_LOAD_ERROR_MSG}), 500
    else:
        logging.warning(f"上傳了不允許的檔案類型：{file.filename}")
        return jsonify({'error': '不允許的檔案類型。請上傳 .xls 或 .xlsx 檔案。'}), 400

@app.route('/get_department', methods=['POST'])
def get_department_info():
    logging.debug("收到對 '/get_department' 的 POST 請求。")

    if PROCESSED_EXCEL_DATA is None:
        error_msg_to_show = EXCEL_LOAD_ERROR_MSG or '伺服器錯誤：Excel 資料未載入或載入失敗。'
        logging.error(f"查詢中止，因 Excel 資料未成功載入: {error_msg_to_show}")
        return jsonify({'error': error_msg_to_show}), 500

    gsat_id_input = request.form.get('gsat_id')
    if not gsat_id_input:
        logging.warning("查詢請求中的學測號碼為空。")
        return jsonify({'error': '學測號碼不得為空！'}), 400

    gsat_id_cleaned = _clean_input_gsat_id(gsat_id_input)
    logging.info(f"接收到查詢學測號碼 (原始: '{gsat_id_input}', 清理後: '{gsat_id_cleaned}')")

    try:
        result_df = PROCESSED_EXCEL_DATA[PROCESSED_EXCEL_DATA[GSAT_ID_COLUMN] == gsat_id_cleaned]
        logging.debug(f"符合學測號碼 '{gsat_id_cleaned}' 的查詢結果：\n{result_df}")

        if not result_df.empty:
            department_names_list = result_df[DEPARTMENT_COLUMN].tolist()
            department_names_list = [name for name in department_names_list if name] 

            if department_names_list:
                logging.info(f"為學測號碼 '{gsat_id_cleaned}' 找到校系：{department_names_list}")
                return jsonify({'gsat_id': gsat_id_cleaned, 'department_names': department_names_list})
            else:
                msg = f"學測號碼 '{gsat_id_cleaned}' 查有資料，但無有效校系名稱。"
                logging.warning(msg)
                return jsonify({'error': msg}), 404
        else:
            msg = f"查無此學測號碼：'{gsat_id_cleaned}'"
            logging.warning(msg)
            return jsonify({'error': msg}), 404
    except Exception as e:
        logging.error(f"查詢學測號碼 '{gsat_id_cleaned}' 時發生未預期錯誤：{e}", exc_info=True)
        return jsonify({'error': f'查詢過程中發生伺服器內部錯誤。'}), 500

# --- 應用程式啟動 ---
if __name__ == '__main__':
    # 應用程式啟動時，嘗試載入預設的 Excel 檔案
    if not load_and_process_excel(source_filename=EXCEL_FILE_PATH): # 傳入預設檔名
        logging.critical(f"！！！警告：預設 Excel 檔案未能成功載入，應用程式功能將受限。錯誤：{EXCEL_LOAD_ERROR_MSG}！！！")
    else:
        logging.info(f"應用程式資料 ({CURRENT_DATA_FILENAME}) 準備完成。")
    
    app.run(debug=True, host='0.0.0.0', port=5000)