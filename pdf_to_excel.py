import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def convert_pdf_to_excel():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not file_path:
        return
    
    save_path = os.path.splitext(file_path)[0] + ".xlsx"

    try:
        writer = pd.ExcelWriter(save_path, engine="openpyxl")

        with pdfplumber.open(file_path) as pdf:
            for i, page in enumerate(pdf.pages):
                # ① 表データを取得
                tables = page.extract_tables()
                for j, table in enumerate(tables):
                    df_table = pd.DataFrame(table)
                    sheet_name = f"Page{i+1}_Table{j+1}"
                    df_table.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

                # ② 表以外のテキストを取得
                text = page.extract_text()
                if text:
                    df_text = pd.DataFrame([text.split("\n")])
                    sheet_name = f"Page{i+1}_Text"
                    df_text.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        writer.close()
        messagebox.showinfo("成功", f"変換完了！\n{save_path}")

    except Exception as e:
        messagebox.showerror("エラー", str(e))

# GUIセットアップ
root = tk.Tk()
root.title("PDF → Excel 変換ツール（表 + テキスト）")
root.geometry("300x150")

btn_convert = tk.Button(root, text="PDFを選択して変換", command=convert_pdf_to_excel)
btn_convert.pack(pady=30)

root.mainloop()

