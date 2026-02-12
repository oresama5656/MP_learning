# -*- coding: utf-8 -*-
"""
MPラーニング 自動受講ツール (Selenium版 v2)
バックグラウンドで動作可能な半自動ツール
- チカチカ問題を解消（ポップアップウィンドウのみ監視）
- 再生ボタン自動クリック機能追加
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import tkinter as tk
from tkinter import messagebox
import threading
import time

class MPLearningAutoTool:
    """MPラーニング自動受講ツール"""
    
    def __init__(self):
        self.driver = None
        self.monitoring = False
        self.monitor_thread = None
        self.main_window = None  # メインウィンドウのハンドル
        self.popup_window = None  # ポップアップウィンドウのハンドル
        self.queue_list = []  # 予約リスト
        self.continuous_mode = False  # 連続受講モード
        
        # GUI作成
        self.root = tk.Tk()
        self.root.title("MPラーニング 自動受講 v3")
        self.root.geometry("400x500")
        self.root.attributes("-topmost", True)
        self.root.configure(bg='#2C3E50')
        
        # ステータスラベル
        self.status_label = tk.Label(
            self.root, 
            text="停止中", 
            font=('Helvetica', 14, 'bold'),
            bg='#2C3E50',
            fg='#ECF0F1'
        )
        self.status_label.pack(pady=5)
        
        # 詳細ラベル
        self.detail_label = tk.Label(
            self.root, 
            text="", 
            font=('Helvetica', 10),
            bg='#2C3E50',
            fg='#BDC3C7'
        )
        self.detail_label.pack(pady=3)
        
        # ボタンフレーム
        btn_frame = tk.Frame(self.root, bg='#2C3E50')
        btn_frame.pack(pady=5)
        
        # ブラウザ起動ボタン
        self.browser_button = tk.Button(
            btn_frame, 
            text="ブラウザ起動", 
            command=self.start_browser,
            width=12, height=1,
            bg='#3498DB', fg='white',
            font=('Arial', 9, 'bold')
        )
        self.browser_button.pack(side=tk.LEFT, padx=3)
        
        # 監視開始ボタン
        self.start_button = tk.Button(
            btn_frame, 
            text="監視開始", 
            command=self.start_monitoring,
            width=12, height=1,
            bg='#27AE60', fg='white',
            font=('Arial', 9, 'bold'),
            state=tk.DISABLED
        )
        self.start_button.pack(side=tk.LEFT, padx=3)
        
        # 監視停止ボタン
        self.stop_button = tk.Button(
            btn_frame, 
            text="監視停止", 
            command=self.stop_monitoring,
            width=12, height=1,
            bg='#E74C3C', fg='white',
            font=('Arial', 9, 'bold'),
            state=tk.DISABLED
        )
        self.stop_button.pack(side=tk.LEFT, padx=3)
        
        # ウィンドウ数ラベル
        self.window_label = tk.Label(
            self.root,
            text="ウィンドウ数: -",
            font=('Helvetica', 9),
            bg='#2C3E50',
            fg='#95A5A6'
        )
        self.window_label.pack(pady=3)
        
        # === 予約機能セクション ===
        queue_frame = tk.LabelFrame(
            self.root,
            text=" 講座予約 ",
            font=('Helvetica', 10, 'bold'),
            bg='#2C3E50',
            fg='#ECF0F1'
        )
        queue_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        # 予約ボタンフレーム
        queue_btn_frame = tk.Frame(queue_frame, bg='#2C3E50')
        queue_btn_frame.pack(pady=5)
        
        # 「この講座を予約」ボタン
        self.add_queue_button = tk.Button(
            queue_btn_frame,
            text="この講座を予約",
            command=self.add_to_queue,
            width=15, height=1,
            bg='#9B59B6', fg='white',
            font=('Arial', 9, 'bold'),
            state=tk.DISABLED
        )
        self.add_queue_button.pack(side=tk.LEFT, padx=3)
        
        # 「選択削除」ボタン
        self.remove_queue_button = tk.Button(
            queue_btn_frame,
            text="選択削除",
            command=self.remove_from_queue,
            width=10, height=1,
            bg='#7F8C8D', fg='white',
            font=('Arial', 9, 'bold')
        )
        self.remove_queue_button.pack(side=tk.LEFT, padx=3)
        
        # 予約件数ラベル
        self.queue_count_label = tk.Label(
            queue_frame,
            text="予約: 0件",
            font=('Helvetica', 10, 'bold'),
            bg='#2C3E50',
            fg='#F39C12'
        )
        self.queue_count_label.pack(pady=3)
        
        # 予約リスト（Listbox）
        list_frame = tk.Frame(queue_frame, bg='#2C3E50')
        list_frame.pack(pady=5, padx=5, fill=tk.BOTH, expand=True)
        
        self.queue_listbox = tk.Listbox(
            list_frame,
            height=6,
            font=('Consolas', 9),
            selectmode=tk.SINGLE,
            bg='#34495E',
            fg='#ECF0F1',
            selectbackground='#3498DB'
        )
        self.queue_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.queue_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.queue_listbox.yview)
        
        # 「連続受講開始」ボタン
        self.continuous_button = tk.Button(
            queue_frame,
            text="連続受講開始",
            command=self.start_continuous,
            width=20, height=2,
            bg='#E67E22', fg='white',
            font=('Arial', 10, 'bold'),
            state=tk.DISABLED
        )
        self.continuous_button.pack(pady=8)
        
        # ウィンドウを閉じる際の処理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # ウィンドウ数の定期更新
        self.update_window_count()
    
    def start_browser(self):
        """ブラウザを起動してログインページを開く"""
        try:
            self.status_label.config(text="ブラウザ起動中...")
            self.root.update()
            
            options = Options()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-popup-blocking")
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.get("https://www.mp-learning.com/Login.aspx")
            
            # メインウィンドウを記録
            self.main_window = self.driver.current_window_handle
            
            self.status_label.config(text="ログイン後、講座を開いてください")
            self.detail_label.config(text="動画再生画面（ポップアップ）を開いてから監視開始")
            self.browser_button.config(state=tk.DISABLED)
            self.start_button.config(state=tk.NORMAL)
            self.add_queue_button.config(state=tk.NORMAL)
            self.continuous_button.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("エラー", f"ブラウザ起動に失敗:\n{e}")
            self.status_label.config(text="起動失敗")
    
    def find_popup_window(self):
        """ポップアップウィンドウ（レッスン画面）を特定"""
        try:
            for handle in self.driver.window_handles:
                if handle != self.main_window:
                    # メインウィンドウ以外を確認
                    self.driver.switch_to.window(handle)
                    title = self.driver.title
                    url = self.driver.current_url
                    # タイトルまたはURLで判定
                    if "レッスン" in title or "Lesson" in title or "LessonStudy" in url:
                        return handle
            return None
        except:
            return None
    
    def start_monitoring(self):
        """監視を開始"""
        if self.driver is None:
            messagebox.showwarning("警告", "先にブラウザを起動してください")
            return
        
        # ポップアップウィンドウを探す
        self.popup_window = self.find_popup_window()
        
        if self.popup_window is None:
            messagebox.showwarning("警告", "レッスン画面（ポップアップ）が見つかりません。\n講座を選んで動画再生画面を開いてください。")
            return
        
        self.monitoring = True
        self.status_label.config(text="監視中...", fg='#2ECC71')
        self.detail_label.config(text="ボタン待機中...")
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        # 監視スレッド開始
        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()
    
    def stop_monitoring(self):
        """監視を停止"""
        self.monitoring = False
        self.status_label.config(text="停止中", fg='#ECF0F1')
        self.detail_label.config(text="")
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
    
    def update_detail(self, text):
        """詳細ラベルを更新（スレッドセーフ）"""
        try:
            self.root.after(0, lambda: self.detail_label.config(text=text))
        except:
            pass
    
    def update_queue_display(self):
        """予約リスト表示を更新"""
        self.queue_listbox.delete(0, tk.END)
        for i, lesson_id in enumerate(self.queue_list):
            self.queue_listbox.insert(tk.END, f"{i+1}. L={lesson_id}")
        self.queue_count_label.config(text=f"予約: {len(self.queue_list)}件")
    
    def add_to_queue(self):
        """現在のポップアップウィンドウの講座を予約リストに追加"""
        if self.driver is None:
            messagebox.showwarning("警告", "ブラウザが起動していません")
            return
        
        try:
            # ポップアップウィンドウを探す
            popup = self.find_popup_window()
            if popup is None:
                # 確認テスト画面やレッスン画面も含めて探す
                for handle in self.driver.window_handles:
                    if handle != self.main_window:
                        self.driver.switch_to.window(handle)
                        url = self.driver.current_url
                        if "LessonStudyFixed" in url or "Lesson" in url:
                            popup = handle
                            break
            
            if popup is None:
                messagebox.showwarning("警告", "講座画面が見つかりません。\n講座を開いてから予約してください。")
                return
            
            self.driver.switch_to.window(popup)
            url = self.driver.current_url
            
            # URLから講座IDを抽出
            if "L=" in url:
                lesson_id = url.split("L=")[-1].split("&")[0]
            else:
                messagebox.showwarning("警告", "講座IDを取得できませんでした")
                return
            
            # 既に予約済みかチェック
            if lesson_id in self.queue_list:
                messagebox.showinfo("情報", "この講座は既に予約済みです")
                return
            
            # 予約リストに講座IDを追加
            self.queue_list.append(lesson_id)
            self.update_queue_display()
            
            # ポップアップを閉じる
            self.driver.close()
            self.driver.switch_to.window(self.main_window)
            
            print(f"講座を予約: L={lesson_id}")
            self.detail_label.config(text=f"予約追加: {len(self.queue_list)}件目")
            
        except Exception as e:
            print(f"予約追加エラー: {e}")
            messagebox.showerror("エラー", f"予約追加に失敗:\n{e}")
    
    def remove_from_queue(self):
        """選択した講座を予約リストから削除"""
        selection = self.queue_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "削除する講座を選択してください")
            return
        
        index = selection[0]
        removed = self.queue_list.pop(index)
        self.update_queue_display()
        print(f"予約削除: {removed}")
    
    def start_continuous(self):
        """連続受講を開始"""
        if self.driver is None:
            messagebox.showwarning("警告", "ブラウザを起動してください")
            return
        
        if len(self.queue_list) == 0:
            messagebox.showwarning("警告", "予約リストが空です。\n講座を予約してから開始してください。")
            return
        
        # 連続受講モードを有効化
        self.continuous_mode = True
        
        # 最初の講座を開く
        self.open_next_queue()
    
    def open_next_queue(self):
        """予約リストの次の講座を開く"""
        if len(self.queue_list) == 0:
            self.continuous_mode = False
            self.status_label.config(text="全講座完了！", fg='#F39C12')
            self.detail_label.config(text="予約リストの講座を全て受講しました")
            messagebox.showinfo("完了", "予約した講座を全て受講しました！")
            return
        
        # リストの先頭から講座IDを取り出す
        next_lesson_id = self.queue_list.pop(0)
        self.update_queue_display()
        
        print(f"次の講座を開く: L={next_lesson_id}")
        self.detail_label.config(text=f"残り{len(self.queue_list)}件...")
        
        try:
            # メインウィンドウに切り替える
            self.driver.switch_to.window(self.main_window)
            time.sleep(0.5)
            
            # マイページに移動
            print("マイページに移動中...")
            self.driver.get("https://www.mp-learning.com/Members/MyPage.aspx")
            time.sleep(3)
            
            # 講座リンク要素を探してクリック
            link_id = f"linkNewLessonPC-{next_lesson_id}"
            print(f"講座リンクを探しています: {link_id}")
            
            link_clicked = False
            
            # 方法1: 直接IDで要素を探してクリック
            try:
                link = self.driver.find_element(By.ID, link_id)
                print(f"リンク発見: {link_id}")
                self.driver.execute_script("arguments[0].scrollIntoView(true);", link)
                time.sleep(0.5)
                link.click()
                print("講座リンクをクリックしました")
                link_clicked = True
            except Exception as e:
                print(f"ID検索失敗: {e}")
            
            # 方法2: XPathで探す
            if not link_clicked:
                try:
                    link = self.driver.find_element(By.XPATH, f"//a[contains(@id, '{next_lesson_id}')]")
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", link)
                    time.sleep(0.5)
                    link.click()
                    print("XPathで講座リンクをクリックしました")
                    link_clicked = True
                except Exception as e2:
                    print(f"XPath検索も失敗: {e2}")
            
            # 方法3: lnkNewLesson_OnClick をJS直接呼び出し（新着以外の講座用）
            if not link_clicked:
                print("方法3: lnkNewLesson_OnClick を直接呼び出します")
                for lesson_type in [2, 3]:
                    try:
                        result = self.driver.execute_script(f"""
                            lnkNewLesson_OnClick("{next_lesson_id}", {lesson_type}, 1, "linkNewLessonPC-{next_lesson_id}", 0);
                            return 'called_type_{lesson_type}';
                        """)
                        print(f"lnkNewLesson_OnClick (type={lesson_type}): {result}")
                        link_clicked = True
                        break
                    except Exception as e3:
                        print(f"type={lesson_type} 失敗: {e3}")
            
            if not link_clicked:
                print(f"全ての方法で講座 {next_lesson_id} のオープンに失敗しました")
            
            # ポップアップウィンドウを探して監視開始
            time.sleep(4)
            self.popup_window = self.find_popup_window()
            print(f"ポップアップ検出: {self.popup_window}")
            
            if self.popup_window:
                self.driver.switch_to.window(self.popup_window)
                if not self.monitoring:
                    self.start_monitoring()
                # 監視開始後も明示的に再生ボタンをクリック
                time.sleep(3)
                print("連続受講: 再生開始を試みます")
                self.click_play_button()
            else:
                print("ポップアップが見つかりませんでした")
        except Exception as e:
            print(f"講座オープンエラー: {e}")

    
    def click_play_button(self):
        """再生ボタンをクリック（EQプレイヤーのJavaScript APIを使用）"""
        try:
            print("=== 再生開始処理 ===")
            
            # 方法1: まずプレイヤー領域をクリック（オートプレイポリシー回避）
            try:
                player_area = self.driver.find_element(By.ID, "eqPlayer")
                
                # ActionChainsで確実にクリック
                actions = ActionChains(self.driver)
                actions.move_to_element(player_area).click().perform()
                print("Step1: プレイヤー領域をクリック")
                time.sleep(0.5)
                
                # さらに普通のクリックも試す
                player_area.click()
                print("Step2: 追加クリック")
                time.sleep(0.5)
                
            except Exception as e:
                print(f"プレイヤークリック失敗: {e}")
            
            # 方法2: JavaScript APIで再生
            for attempt in range(3):
                try:
                    result = self.driver.execute_script("""
                        if (typeof player !== 'undefined' && player && player.accessor) {
                            try {
                                player.accessor.play();
                                return 'play() called';
                            } catch(e) {
                                return 'play() error: ' + e.toString();
                            }
                        }
                        return 'player not ready';
                    """)
                    print(f"Step3: 再生API試行{attempt+1}: {result}")
                    if 'called' in result:
                        break
                except Exception as e:
                    print(f"再生API例外{attempt+1}: {e}")
                time.sleep(0.5)
            
            # 方法3: video要素を直接操作
            try:
                self.driver.execute_script("""
                    // iframe内のvideo要素を探して再生
                    var iframe = document.querySelector('#eqPlayer iframe');
                    if (iframe && iframe.contentDocument) {
                        var video = iframe.contentDocument.querySelector('video');
                        if (video) {
                            video.muted = true;  // ミュートにして再生（オートプレイポリシー回避）
                            video.play();
                        }
                    }
                """)
                print("Step4: iframe内video.play()試行")
            except Exception as e:
                print(f"iframe操作失敗（クロスオリジンの可能性）: {e}")
            
            return True
        except Exception as e:
            print(f"再生ボタンエラー: {e}")
            return False
    
    def monitor_loop(self):
        """監視ループ（ポップアップウィンドウのみ監視）"""
        while self.monitoring:
            try:
                # ポップアップウィンドウがまだ存在するか確認
                if self.popup_window not in self.driver.window_handles:
                    # 連続受講モードなら次の講座を開く
                    if self.continuous_mode and len(self.queue_list) > 0:
                        print("ポップアップ閉じ検出 - 次の講座を開く")
                        self.root.after(1000, self.open_next_queue)
                        time.sleep(3)
                        continue
                    elif self.continuous_mode and len(self.queue_list) == 0:
                        # 全講座完了
                        self.continuous_mode = False
                        self.root.after(0, lambda: self.status_label.config(text="全講座完了！", fg='#F39C12'))
                        self.root.after(0, lambda: self.detail_label.config(text="予約した講座を全て受講しました"))
                        print("全講座完了！")
                        time.sleep(2)
                        continue
                    
                    # ポップアップが閉じられた → 新しいポップアップを探す
                    self.popup_window = self.find_popup_window()
                    if self.popup_window is None:
                        self.update_detail("ポップアップを探しています...")
                        time.sleep(2)
                        continue
                
                # ポップアップウィンドウに切り替え（1回だけ）
                current = self.driver.current_window_handle
                if current != self.popup_window:
                    self.driver.switch_to.window(self.popup_window)
                
                # 「次の学習へ」「テストへ」ボタンを探す
                try:
                    btn = self.driver.find_element(By.ID, "btn-next-study")
                    if btn.is_displayed():
                        link = self.driver.find_element(By.ID, "btn-next-study-link")
                        link_text = link.text
                        self.update_detail(f"検出: {link_text}")
                        print(f"ボタン検出: {link_text}")
                        
                        # クリック
                        btn.click()
                        print(f"クリック完了: {link_text}")
                        self.update_detail(f"クリック: {link_text}")
                        
                        # ページ遷移を待機
                        time.sleep(3)
                        
                        # 新しいポップアップを探す（ページ遷移後）
                        self.popup_window = self.find_popup_window()
                        if self.popup_window:
                            self.driver.switch_to.window(self.popup_window)
                            
                            # 再生ボタンをクリック
                            time.sleep(2)  # ページ読み込み待機
                            self.click_play_button()
                            self.update_detail("再生開始...")
                        
                        continue
                except:
                    pass  # 要素が見つからない場合は無視
                
                # テスト回答画面の処理（タイトルで判定）
                try:
                    page_title = self.driver.title
                    print(f"現在のページタイトル: {page_title}")
                    
                    if "確認テスト" in page_title:
                        # 答案提出ボタンを探す
                        submit_btn = self.driver.find_element(By.ID, "ctl00_examBody_lnkExamAnswerSubmit")
                        print(f"答案提出ボタン発見: displayed={submit_btn.is_displayed()}")
                        if submit_btn.is_displayed():
                            self.update_detail("テスト回答中...")
                            print("テスト回答画面検出 - 自動回答開始")
                            
                            # JavaScriptで各問の最初のラジオボタンを選択
                            result = self.driver.execute_script("""
                                // 全てのラジオボタングループを取得
                                var radioGroups = {};
                                var radios = document.querySelectorAll('input[type="radio"]');
                                radios.forEach(function(radio) {
                                    var name = radio.name;
                                    if (!radioGroups[name]) {
                                        radioGroups[name] = [];
                                    }
                                    radioGroups[name].push(radio);
                                });
                                
                                // 各グループの最初のラジオボタンを選択
                                var count = 0;
                                for (var name in radioGroups) {
                                    if (radioGroups[name].length > 0) {
                                        radioGroups[name][0].checked = true;
                                        count++;
                                    }
                                }
                                return count + '問回答';
                            """)
                            print(f"選択肢選択完了: {result}")
                            time.sleep(1)
                            
                            # 答案提出
                            submit_btn = self.driver.find_element(By.ID, "ctl00_examBody_lnkExamAnswerSubmit")
                            submit_btn.click()
                            print("答案提出完了")
                            self.update_detail("答案提出完了")
                            time.sleep(3)
                            continue
                except Exception as e:
                    print(f"テスト回答エラー: {e}")
                
                # テスト解説画面の処理（次へボタンがある場合）
                try:
                    next_btn = self.driver.find_element(By.ID, "ctl00_examBody_cmdNext")
                    if next_btn.is_displayed():
                        self.update_detail("解説画面 - 次へ")
                        print("解説画面検出 - 次へクリック")
                        next_btn.click()
                        time.sleep(3)
                        continue
                except:
                    pass
                
                # アンケート画面の処理（終了ボタンをクリック）
                try:
                    # タイトルでアンケート画面を判定
                    if "アンケート" in self.driver.title:
                        end_label = self.driver.find_element(By.ID, "panel-end-label")
                        if end_label.is_displayed():
                            self.update_detail("アンケート画面 - 終了")
                            print("アンケート画面検出 - 終了クリック")
                            end_label.click()
                            time.sleep(2)
                            continue
                except:
                    pass
                
                # 動画が停止していたら再生する（強制再生機能）
                # ただし「次の学習へ」ボタンが表示中は除外（動画終了状態なので）
                try:
                    btn_next = self.driver.find_element(By.ID, "btn-next-study")
                    btn_visible = btn_next.is_displayed()
                except:
                    btn_visible = False
                
                if not btn_visible:
                    try:
                        is_playing = self.driver.execute_script("""
                            if (typeof player !== 'undefined' && player && player.accessor) {
                                // プレイヤーが一時停止状態かチェック
                                var isPaused = player.accessor.isPaused();
                                if (isPaused === true) {
                                    player.accessor.play();
                                    return 'resumed';
                                }
                                return 'playing';
                            }
                            return 'no_player';
                        """)
                        if is_playing == 'resumed':
                            print("動画を再開しました")
                            self.update_detail("動画再開...")
                        elif is_playing == 'playing':
                            self.update_detail("再生中...")
                        else:
                            self.update_detail("ボタン待機中...")
                    except:
                        self.update_detail("ボタン待機中...")
                else:
                    self.update_detail("動画終了 - ボタン待機中...")
                
            except Exception as e:
                print(f"監視エラー: {e}")
                self.update_detail(f"エラー: {str(e)[:30]}")
            
            # 1秒待機
            time.sleep(1)
    
    def update_window_count(self):
        """ウィンドウ数を更新"""
        try:
            if self.driver:
                count = len(self.driver.window_handles)
                self.window_label.config(text=f"ウィンドウ数: {count}")
            else:
                self.window_label.config(text="ウィンドウ数: -")
        except:
            self.window_label.config(text="ウィンドウ数: ?")
        
        self.root.after(1000, self.update_window_count)
    
    def on_closing(self):
        """終了処理"""
        self.monitoring = False
        
        if self.driver:
            if messagebox.askyesno("確認", "ブラウザも終了しますか？"):
                try:
                    self.driver.quit()
                except:
                    pass
        
        self.root.destroy()
    
    def run(self):
        """メインループ開始"""
        self.root.mainloop()


if __name__ == "__main__":
    app = MPLearningAutoTool()
    app.run()
