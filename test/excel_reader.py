import pandas as pd
import os

# --- 設定 ---
# !!! 重要：請將 'your_excel_file.xlsx' 替換成您 Excel 檔案的實際名稱 !!!
EXCEL_FILE_PATH = 'data.xlsx'
# !!! 重要：如果您的 Excel 欄位名稱不同，請修改以下變數 !!!
GSAT_ID_COLUMN = '學測號碼'  # Excel 中學測號碼的欄位名稱
DEPARTMENT_COLUMN = '校系名稱' # Excel 中校系名稱的欄位名稱
# --- 設定結束 ---

def find_department_by_gsat_id(gsat_id_to_find):
    """
    根據學測號碼在 Excel 檔案中尋找校系名稱。
    處理學測號碼可能為浮點數的情況。

    Args:
        gsat_id_to_find (str): 要查詢的學測號碼 (應為字串)。

    Returns:
        str: 如果找到，返回校系名稱；否則返回提示訊息。
    """
    try:
        # 取得 Excel 檔案的絕對路徑
        base_dir = os.path.abspath(os.path.dirname(__file__))
        excel_full_path = os.path.join(base_dir, EXCEL_FILE_PATH)

        # 檢查檔案是否存在
        if not os.path.exists(excel_full_path):
            return f"錯誤：找不到 Excel 檔案 '{EXCEL_FILE_PATH}' 於路徑 '{excel_full_path}'。"

        # 讀取 Excel 檔案
        # 我們可以嘗試在讀取時就指定學測號碼欄位為字串，但这可能仍會保留 ".0"
        # 因此後續處理更為重要
        data_df = pd.read_excel(excel_full_path)

        # 檢查必要的欄位是否存在
        if GSAT_ID_COLUMN not in data_df.columns:
            return f"錯誤：Excel 檔案中找不到指定的學測號碼欄位 '{GSAT_ID_COLUMN}'。"
        if DEPARTMENT_COLUMN not in data_df.columns:
            return f"錯誤：Excel 檔案中找不到指定的校系名稱欄位 '{DEPARTMENT_COLUMN}'。"

        # --- 處理學測號碼欄位的格式 ---
        # 1. 處理缺失值 (NaN)，可以選擇填充為空字串或特殊標記，這裡填充為空字串
        data_df[GSAT_ID_COLUMN] = data_df[GSAT_ID_COLUMN].fillna('')

        # 2. 定義一個函數來清理學測號碼
        def clean_gsat_id(id_val):
            if isinstance(id_val, float):
                # 如果是浮點數，先轉成整數 (去掉 .0)，再轉成字串
                return str(int(id_val))
            # 其他情況 (例如已經是整數或字串)，直接轉成字串並去除前後空白
            return str(id_val).strip()

        # 3. 將清理函數應用到學測號碼欄位
        data_df[GSAT_ID_COLUMN] = data_df[GSAT_ID_COLUMN].apply(clean_gsat_id)
        # --- 學測號碼格式處理完畢 ---

        # 確保輸入的 gsat_id_to_find 也是標準化的字串 (去除前後空白)
        gsat_id_to_find_cleaned = str(gsat_id_to_find).strip()

        # 在 DataFrame 中搜尋學測號碼
        # 現在 data_df[GSAT_ID_COLUMN] 中的值應該是像 "10000001" 這樣的字串
        result_row = data_df[data_df[GSAT_ID_COLUMN] == gsat_id_to_find_cleaned]

        if not result_row.empty:
            department_name = result_row.iloc[0][DEPARTMENT_COLUMN]
            return f"學測號碼 '{gsat_id_to_find_cleaned}' 對應的校系是：{department_name}"
        else:
            return f"查無此學測號碼：'{gsat_id_to_find_cleaned}'"

    except FileNotFoundError:
        return f"錯誤：Excel 檔案 '{EXCEL_FILE_PATH}' 未找到。"
    except KeyError as e:
        return f"錯誤：欄位名稱設定錯誤或 Excel 檔案中缺少該欄位 - {e}。請檢查 GSAT_ID_COLUMN 和 DEPARTMENT_COLUMN 的設定。"
    except ValueError as e:
        # 這個錯誤可能在 int(id_val) 時發生，如果浮點數無法直接轉為整數 (例如 NaN，但已被 fillna 處理)
        # 或者 Excel 中的學測號碼欄位包含無法轉換為整數的文字 (例如 "ABC" 而非 "123.0")
        return f"錯誤：處理學測號碼欄位時發生數值轉換錯誤 - {e}。請檢查 '{GSAT_ID_COLUMN}' 欄位的資料格式。"
    except Exception as e:
        return f"讀取或查詢 Excel 時發生未預期的錯誤：{e}"

if __name__ == "__main__":
    input_gsat_id = input(f"請輸入要查詢的 {GSAT_ID_COLUMN}：").strip()

    if input_gsat_id:
        result_message = find_department_by_gsat_id(input_gsat_id)
        print(result_message)
    else:
        print("您沒有輸入任何學測號碼。")