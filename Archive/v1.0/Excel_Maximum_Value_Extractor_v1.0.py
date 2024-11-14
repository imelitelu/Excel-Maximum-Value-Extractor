import pandas as pd
import tkinter as tk
from tkinter import ttk, simpledialog
from tkinter import filedialog, messagebox
import time
import sys

def main():
    # Create the main window (hidden)
    root = tk.Tk()
    root.withdraw()

    # Default column names
    default_condition_column = "Ch2, x"
    default_target_column = "Ch2, y"

    # Let the user input the condition column name
    while True:
        condition_column = simpledialog.askstring(
            "輸入條件列名稱",
            f"請輸入作為條件的列名稱（預設為 {default_condition_column}）：",
            initialvalue=default_condition_column,
            parent=root
        )
        if condition_column is None or condition_column.strip() == "":
            if messagebox.askyesno("退出", "未輸入條件列名稱，是否退出程序？"):
                root.destroy()
                return
            else:
                continue
        else:
            messagebox.showinfo("退出", "程序將退出。")
            condition_column = condition_column.strip()
            break

    # Let the user input the target column name
    while True:
        target_column = simpledialog.askstring(
            "輸入目標列名稱",
            f"請輸入要提取最大值的列名稱（預設為 {default_target_column}）：",
            initialvalue=default_target_column,
            parent=root
        )
        if target_column is None or target_column.strip() == "":
            if messagebox.askyesno("退出", "未輸入目標列名稱，是否退出程序？"):
                root.destroy()
                return
            else:
                continue
        else:
            target_column = target_column.strip()
            break

    # Loop until the user selects a file or cancels the operation
    while True:
        # Open the file selection dialog to choose the Excel file to process
        file_path = filedialog.askopenfilename(
            title="請選擇要處理的 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx")]
        )

        if not file_path:
            # Ask the user whether to exit
            retry = messagebox.askretrycancel("警告", "未選擇任何文件，是否要重新選擇？")
            if retry:
                continue  # Redisplay the file selection dialog
            else:
                messagebox.showinfo("退出", "程序將退出。")
                root.destroy()
                return  # Exit the program
        else:
            break  # User has selected a file, exit the loop

    try:
        # Open the Excel file in read-only mode
        wb = pd.ExcelFile(file_path)
        sheet_names = wb.sheet_names
    except Exception as e:
        messagebox.showerror("錯誤", f"無法讀取 Excel 文件：{e}")
        root.destroy()
        return

    max_values_list = []
    missing_columns_sheets = []
    error_sheets = []

    # Get the total number of worksheets for progress calculation
    total_sheets = len(sheet_names)
    processed_sheets = 0

    # Create the progress window
    progress_window = tk.Toplevel()
    progress_window.title("處理進度")

    # Flag to control loop execution
    processing = True  # Add this variable

    # Handle the close event
    def on_close():
        nonlocal processing  # Declare that we want to modify the variable in the enclosing scope
        # Ask the user whether to exit
        if messagebox.askokcancel("退出", "正在處理，是否確定要退出？"):
            processing = False  # Set the flag to False
            progress_window.destroy()
            root.destroy()
            sys.exit()  # Exit the program immediately

    progress_window.protocol("WM_DELETE_WINDOW", on_close)

    progress_label = tk.Label(progress_window, text="正在處理工作表...")
    progress_label.pack(pady=5)

    sheet_label = tk.Label(progress_window, text="")
    sheet_label.pack(pady=5)

    progress_bar = ttk.Progressbar(progress_window, length=300, mode='determinate')
    progress_bar.pack(pady=5)
    progress_bar['maximum'] = total_sheets

    time_label = tk.Label(progress_window, text="預估剩餘時間：計算中...")
    time_label.pack(pady=5)

    # Start timing
    start_time = time.time()

    # Iterate over each worksheet
    for sheet_name in sheet_names:
        if not processing:
            break  # Exit the loop if processing is False

        # Update the current sheet name
        try:
            sheet_label.config(text=f"當前工作表：{sheet_name}")
            progress_window.update()
        except tk.TclError:
            # If the progress window has been destroyed, exit the loop
            break

        try:
            # Read only the specified columns from the sheet
            df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=[condition_column, target_column])
        except ValueError:
            # If the columns do not exist, record it
            missing_columns_sheets.append(sheet_name)
            processed_sheets += 1
            continue
        except Exception as e:
            error_sheets.append((sheet_name, str(e)))
            processed_sheets += 1
            continue

        # Filter rows where condition_column > 10
        filtered_df = df[df[condition_column] > 10]

        if not filtered_df.empty:
            # Calculate the maximum value of the target column
            max_value = filtered_df[target_column].max()
        else:
            max_value = None  # No data meets the condition

        # Append the result
        max_values_list.append({'測試名稱': sheet_name, f'最大值 {target_column}': max_value})

        # Update progress
        processed_sheets += 1
        progress_bar['value'] = processed_sheets

        # Calculate estimated remaining time
        elapsed_time = time.time() - start_time
        if processed_sheets > 0:
            avg_time_per_sheet = elapsed_time / processed_sheets
            remaining_time = avg_time_per_sheet * (total_sheets - processed_sheets)
            time_label.config(text=f"預估剩餘時間：約 {int(remaining_time)} 秒")
        else:
            time_label.config(text="預估剩餘時間：計算中...")

        # Update the progress window
        try:
            progress_window.update()
        except tk.TclError:
            # If the progress window has been destroyed, exit the loop
            break

    # Processing complete, close the progress window
    if processing:
        progress_window.destroy()

    if missing_columns_sheets:
        missing_sheets_str = '\n'.join(missing_columns_sheets)
        messagebox.showinfo("提示", f"以下工作表中沒有找到指定的列，已跳過：\n{missing_sheets_str}")

    if error_sheets:
        error_sheets_str = '\n'.join([f"{name}: {error}" for name, error in error_sheets])
        messagebox.showinfo("提示", f"以下工作表處理時出錯，已跳過：\n{error_sheets_str}")

    if not max_values_list:
        messagebox.showwarning("警告", "未能提取任何最大值，程序將退出。")
        root.destroy()
        return

    # Convert the list of max values to a DataFrame
    max_values_df = pd.DataFrame(max_values_list)

    # Loop until the user selects a save path or cancels the operation
    while processing:
        # Let the user choose the file path to save the results
        save_file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            title="請選擇保存結果的文件路徑"
        )

        if not save_file_path:
            # Ask the user whether to exit
            retry = messagebox.askretrycancel("警告", "未選擇保存路徑，是否要重新選擇？")
            if retry:
                continue  # Redisplay the save dialog
            else:
                messagebox.showinfo("退出", "程式將退出。")
                root.destroy()
                return  # Exit the program
        else:
            break  # User has selected a save path, exit the loop

    if not processing:
        return  # If the user chose to exit, end the program

    try:
        # Save the results to a new Excel file
        max_values_df.to_excel(save_file_path, index=False)
        messagebox.showinfo("完成", f"最大值已成功保存到：\n{save_file_path}")
    except Exception as e:
        messagebox.showerror("錯誤", f"無法保存 Excel 文件：{e}")

    # Close the main window and exit the program
    root.destroy()
    sys.exit()

if __name__ == "__main__":
    main()
