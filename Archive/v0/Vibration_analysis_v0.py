import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import time

def main():
    # 创建隐藏的主窗口
    root = tk.Tk()
    root.withdraw()

    # 弹出文件选择对话框，选择要处理的 Excel 文件
    file_path = filedialog.askopenfilename(
        title="請選擇要處理的 Excel 文件",
        filetypes=[("Excel 文件", "*.xlsx *.xls")]
    )

    if not file_path:
        messagebox.showwarning("警告", "未選擇任何文件，程序將退出。")
        return

    try:
        # 读取 Excel 文件的所有工作表
        sheets_dict = pd.read_excel(file_path, sheet_name=None)
    except Exception as e:
        messagebox.showerror("錯誤", f"無法讀取 Excel 文件：{e}")
        return

    max_values_list = []

    # 获取工作表总数，用于进度计算
    total_sheets = len(sheets_dict)
    processed_sheets = 0

    # 创建进度窗口
    progress_window = tk.Toplevel()
    progress_window.title("處理進度")
    progress_label = tk.Label(progress_window, text="正在處理工作表...")
    progress_label.pack(pady=10)

    progress_bar = ttk.Progressbar(progress_window, length=300, mode='determinate')
    progress_bar.pack(pady=10)
    progress_bar['maximum'] = total_sheets

    time_label = tk.Label(progress_window, text="預估剩餘時間：計算中...")
    time_label.pack(pady=10)

    # 开始计时
    start_time = time.time()

    # 遍历每个工作表，提取 Y 的最大值
    for sheet_name, df in sheets_dict.items():
        # 检查是否存在 'Ch2, y' 列，若没有则跳过该工作表
        if 'Ch2, y' in df.columns:
            max_y = df['Ch2, y'].max()
            max_values_list.append({
                '測試名稱': sheet_name,
                '最大值 y': max_y
            })
        else:
            messagebox.showinfo("提示", f"工作表 '{sheet_name}' 中沒有找到 'Ch2, y' 列，已跳過該表。")

        # 更新进度
        processed_sheets += 1
        progress_bar['value'] = processed_sheets

        # 计算预估剩余时间
        elapsed_time = time.time() - start_time
        if processed_sheets > 0:
            avg_time_per_sheet = elapsed_time / processed_sheets
            remaining_time = avg_time_per_sheet * (total_sheets - processed_sheets)
            time_label.config(text=f"預估剩餘時間：約 {int(remaining_time)} 秒")

        # 更新进度窗口
        progress_window.update()

    # 处理完成，关闭进度窗口
    progress_window.destroy()

    if not max_values_list:
        messagebox.showwarning("警告", "未能提取任何最大值，程式將退出。")
        return

    # 将最大值列表转换为 DataFrame
    max_values_df = pd.DataFrame(max_values_list)

    # 让用户选择保存结果的文件路径
    save_file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel 文件", "*.xlsx *.xls")],
        title="請選擇保存結果的文件路徑"
    )

    if not save_file_path:
        messagebox.showwarning("警告", "未選擇保存路徑，程式將退出。")
        return

    try:
        # 将结果保存到新的 Excel 文件
        max_values_df.to_excel(save_file_path, index=False)
        messagebox.showinfo("完成", f"最大值已成功保存到：\n{save_file_path}")
    except Exception as e:
        messagebox.showerror("錯誤", f"無法保存 Excel 文件：{e}")

    # 关闭主窗口
    root.destroy()

if __name__ == "__main__":
    main()
