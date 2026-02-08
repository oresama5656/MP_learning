# -*- coding: utf-8 -*-
import cv2
import numpy as np
import pyautogui
import time
import tkinter as tk
from threading import Thread
import webbrowser

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("画像監視ツール")
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 300
        window_height = 150
        self.configure(bg='#FF7F50')
        self.monitoring = True
        self.status_label = tk.Label(self, text="一生懸命監視中...", font=('Helvetica', 12), bg='#FF7F50', fg='white')
        self.status_label.pack(pady=20)
        self.stop_button = tk.Button(self, text="停止", command=self.stop_monitoring, width=10)
        self.stop_button.pack(side=tk.LEFT, padx=20, pady=10)
        self.login_button = tk.Button(self, text="ログイン", command=self.open_login_page, width=10)
        self.login_button.pack(side=tk.LEFT, padx=20, pady=10)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.thread = Thread(target=self.monitor, daemon=True)
        self.thread.start()

    def open_login_page(self):
        webbrowser.open("https://www.mp-learning.com/Login.aspx")

    def on_closing(self):
        self.stop_monitoring()

    def stop_monitoring(self):
        self.monitoring = False
        if self.thread:
            self.thread.join(timeout=1)
        self.destroy()

    def monitor(self):
        while self.monitoring:
            try:
                result1 = multi_scale_template_matching()
                result2 = to_test_multi_scale_template_matching()
            except Exception as e:
                print(f"エラー: {e}")
            time.sleep(0.5)

def multi_scale_template_matching():
    screenshot = pyautogui.screenshot()
    main_image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
    template = cv2.imread("click.png", cv2.IMREAD_GRAYSCALE)
    if template is None:
        return "なし"
    h, w = template.shape
    min_scale, max_scale, step_scale = 0.5, 1.5, 0.1
    max_val, max_pt = -1, None
    for scale in np.arange(min_scale, max_scale, step_scale):
        resized_template = cv2.resize(template, (int(w * scale), int(h * scale)))
        res = cv2.matchTemplate(main_image, resized_template, cv2.TM_CCOEFF_NORMED)
        min_val, temp_max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        if temp_max_val > max_val:
            max_val, max_pt = temp_max_val, (max_loc, scale)
    threshold = 0.8
    if max_val >= threshold and max_pt is not None:
        max_loc, scale = max_pt
        center_x = max_loc[0] + int(w * scale) // 2
        center_y = max_loc[1] + int(h * scale) // 2
        print(f"画像が見つかりました！ 位置: ({center_x}, {center_y})")
        pyautogui.leftClick(center_x, center_y)
        time.sleep(0.3)
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.press('space')
        return "スタート"
    return "なし"

def to_test_multi_scale_template_matching():
    screenshot = pyautogui.screenshot()
    main_image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
    template = cv2.imread("to_test.png", cv2.IMREAD_GRAYSCALE)
    if template is None:
        return "なし"
    h, w = template.shape
    min_scale, max_scale, step_scale = 0.5, 1.5, 0.1
    max_val, max_pt = -1, None
    for scale in np.arange(min_scale, max_scale, step_scale):
        resized_template = cv2.resize(template, (int(w * scale), int(h * scale)))
        res = cv2.matchTemplate(main_image, resized_template, cv2.TM_CCOEFF_NORMED)
        min_val, temp_max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        if temp_max_val > max_val:
            max_val, max_pt = temp_max_val, (max_loc, scale)
    threshold = 0.8
    if max_val >= threshold and max_pt is not None:
        max_loc, scale = max_pt
        center_x = max_loc[0] + int(w * scale) // 2
        center_y = max_loc[1] + int(h * scale) // 2
        print(f"画像が見つかりました！ 位置: ({center_x}, {center_y})")
        pyautogui.leftClick(center_x, center_y)
        time.sleep(0.3)
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.press('space')
        time.sleep(0.1)
        pyautogui.press('enter')
        return "スタート"
    return "なし"

if __name__ == "__main__":
    app = App()
    app.mainloop()
