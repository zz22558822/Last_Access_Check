import os
import sys
import datetime
import csv
from alive_progress import alive_bar

# 格式化檔案大小
def format_size(size):
    """格式化檔案大小為 GB, MB, KB, B"""
    if size >= 1 << 30:  # 大於等於 1GB
        return f"{size / (1 << 30):.2f} GB"
    elif size >= 1 << 20:  # 大於等於 1MB
        return f"{size / (1 << 20):.2f} MB"
    elif size >= 1 << 10:  # 大於等於 1KB
        return f"{size / (1 << 10):.2f} KB"
    else:  # 小於 1KB
        return f"{size} B"

# 記錄錯誤資訊
def log_error(message):
    """將錯誤資訊寫入錯誤日誌檔案"""
    if hasattr(sys, '_MEIPASS'):
        error_log_path = os.path.dirname(sys.executable)
    else:
        error_log_path = os.path.dirname(os.path.abspath(__file__))
    error_log_path = os.path.join(error_log_path, f"LAC_Error.txt")
    
    with open(error_log_path, 'a', encoding='utf-8') as error_log_file:
        error_log_file.write(f"{datetime.datetime.now()}: {message}\n")

def main():
# 讓用戶輸入要查詢的目錄
    while True:
        folder_path = input("請輸入要查詢的目錄: ")
        if os.path.isdir(folder_path):
            break
        else:
            print("輸入的目錄無效，請重新輸入。")
            print()

    # 讓用戶輸入檢查週期（以月為單位）
    while True:
        try:
            months = int(input("請輸入檢查週期（月）: "))
            if months > 0:
                break
            else:
                print("檢查週期必須為正整數，請重新輸入。")
                print()
        except ValueError:
            print("輸入無效，請輸入一個正整數。")
            print()
    days = months * 30  # 簡單估算每月30天

    print('----------------------------------------------------------------------------------------------------')
    print()
    print('正在計算檔案數量，建議使用管理員運行，所需時間依檔案數量，請稍後...')


    # 設定報告檔案的路徑
    # script_directory = os.path.dirname(os.path.abspath(__file__)) # py用
    # script_directory = os.path.dirname(sys.executable) # exe用

    if hasattr(sys, '_MEIPASS'):
        script_directory = os.path.dirname(sys.executable)
    else:
        script_directory = os.path.dirname(os.path.abspath(__file__))
    LAC_Log_File = os.path.join(script_directory, 'LAC_Log')
    # 如果 LAC_Log 資料夾不存在
    if not os.path.exists(LAC_Log_File):
        print('LAC_Log 資料夾不存在，已新增。')
        os.makedirs(LAC_Log_File)
    print()
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    report_path = os.path.join(LAC_Log_File, f"LAC-{timestamp}.csv")

    # 設定指定週期前的日期
    period_ago = datetime.datetime.now() - datetime.timedelta(days=days)

    # 計算總檔案數量並顯示進度條
    all_files = []
    total_files = 0
    with alive_bar(title='計算中...') as bar:
        for root, dirs, files in os.walk(folder_path):
            total_files += len(files)
            bar()  # 更新進度條

    print()
    print(f"檔案數量: {total_files} 個")
    print()
    print('----------------------------------------------------------------------------------------------------')
    print()
    print('整理檔案資訊中，所需時間依檔案數量，請稍後...')
    print()



    with alive_bar(total_files, title='整理中...') as bar:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                all_files.append(file_path)
                bar()  # 更新進度條


    print()
    print('----------------------------------------------------------------------------------------------------')
    print()
    print(f'檢查檔案中，共計 {len(all_files)} 個，請稍後...')
    print()

    # 開啟報告檔案並使用 csv 模組寫入
    with open(report_path, 'w', newline='', encoding='utf-8-sig') as report_file:
        csv_writer = csv.writer(report_file)
        # 寫入 CSV 標題
        csv_writer.writerow(["檔案名稱", "完整路徑", "檔案大小", "位元組 (bytes)", "最後開啟時間", "距離當前時間 (天)"])

        # 數量統計
        matching_files_count = 0

        # 使用進度條顯示處理進度
        with alive_bar(len(all_files), title='檢查中...') as bar:
            # 遞迴查詢並篩選出指定週期前最後開啟的檔案
            for file_path in all_files:
                try:
                    file = os.path.basename(file_path)
                    file_size = os.path.getsize(file_path)
                    formatted_size = format_size(file_size)
                    last_access_time = os.path.getatime(file_path)
                    readable_time = datetime.datetime.fromtimestamp(last_access_time)

                    # 格式化最後開啟時間
                    formatted_time = readable_time.strftime('%Y-%m-%d %H:%M:%S')
                    
                    # 檢查最後開啟時間是否超過指定週期
                    if readable_time < period_ago:
                        time_difference = datetime.datetime.now() - readable_time
                        # 寫入 CSV 文件
                        csv_writer.writerow([file, file_path, formatted_size, file_size, formatted_time, time_difference.days])
                        # 數量統計
                        matching_files_count += 1

                except PermissionError as e:
                    error_message = f"無法存取檔案: {file_path}，權限不足。\n錯誤訊息: {str(e)}"
                    print(error_message)
                    log_error(error_message)
                except FileNotFoundError as e:
                    error_message = f"檔案未找到: {file_path}。\n錯誤訊息: {str(e)}"
                    print(error_message)
                    log_error(error_message)
                except OSError as e:
                    error_message = f"操作錯誤 ({e.errno}): {file_path}，\n錯誤訊息：{e.strerror}"
                    print(error_message)
                    log_error(error_message)
                except Exception as e:
                    error_message = f"發生未知錯誤: {file_path}，\n錯誤訊息：{str(e)}"
                    print(error_message)
                    log_error(error_message)

                # 更新進度條
                bar()

    print()
    print('----------------------------------------------------------------------------------------------------')
    # 計算占比，避免除以零的錯誤
    if len(all_files) > 0:
        percentage = (matching_files_count / len(all_files)) * 100
    print(f'符合條件: {matching_files_count} 個，總檔案數: {len(all_files)} 個，百分比約: {round(percentage, 1)}%')
    print(f"報告產出: {report_path}")
    print('----------------------------------------------------------------------------------------------------')


if __name__ == "__main__":
    main()

    print()
    os.system('pause')
    sys.exit()
