# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import tkinter as tk
from tkinter import messagebox
import os
import time
import traceback

# 保存先ディレクトリ（スクリプトと同じ場所）
SAVE_DIR = os.path.dirname(os.path.abspath(__file__))

def save_html():
    """現在開いている全てのウィンドウのHTMLを保存する"""
    try:
        window_handles = driver.window_handles
        timestamp = int(time.time())
        
        saved_files = []
        
        print(f"ウィンドウ数: {len(window_handles)}")
        
        for i, handle in enumerate(window_handles):
            driver.switch_to.window(handle)
            title = driver.title
            # ファイル名に使えない文字を除去
            safe_title = "".join(c if c.isalnum() or c in "._- " else "_" for c in title)[:30]
            if not safe_title:
                safe_title = f"window_{i}"
            
            filename = f"dump_{timestamp}_{i}_{safe_title}.html"
            filepath = os.path.join(SAVE_DIR, filename)
            
            print(f"保存中: {filepath}")
            
            html_content = driver.page_source
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(html_content)
            saved_files.append(filename)
            print(f"保存完了: {filename} ({len(html_content)} bytes)")
            
        messagebox.showinfo("成功", f"以下のファイルを保存しました:\n" + "\n".join(saved_files) + f"\n\n保存先: {SAVE_DIR}")
        
    except Exception as e:
        error_detail = traceback.format_exc()
        print(f"エラー詳細:\n{error_detail}")
        messagebox.showerror("エラー", f"保存に失敗しました:\n{e}\n\n詳細はコンソールを確認してください")

def on_closing():
    """終了処理"""
    if messagebox.askokcancel("終了", "ブラウザも閉じますか？"):
        try:
            driver.quit()
        except:
            pass
        root.destroy()

# Selenium初期化
options = Options()
options.add_argument("--start-maximized")
# ポップアップを許可
options.add_argument("--disable-popup-blocking")

print("Chromeを起動中...")

try:
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("https://www.mp-learning.com/Login.aspx")
    print("ログインページを開きました")
except Exception as e:
    print(f"Chrome起動エラー: {e}")
    print(traceback.format_exc())
    input("Enterキーで終了...")
    exit(1)

# GUI作成
root = tk.Tk()
root.title("HTML調査ツール")
root.geometry("350x180")
root.attributes("-topmost", True)

label = tk.Label(root, text="1. 手動で動画終了画面まで進む\n2. 「HTML保存」を押す\n\n※ポップアップも含めて保存されます", pady=10)
label.pack()

btn_save = tk.Button(root, text="HTML保存", command=save_html, height=2, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
btn_save.pack(fill=tk.X, padx=20, pady=10)

# ウィンドウ数表示用
def update_window_count():
    try:
        count = len(driver.window_handles)
        label_count.config(text=f"現在のウィンドウ数: {count}")
    except:
        label_count.config(text="ブラウザが閉じられました")
    root.after(1000, update_window_count)

label_count = tk.Label(root, text="現在のウィンドウ数: 1", fg="gray")
label_count.pack()
update_window_count()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
