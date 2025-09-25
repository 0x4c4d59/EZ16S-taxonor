import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import configparser
import math
import threading


class EZBioCloudScraper:
    def __init__(self):
        self.root = tk.Tk()
        self.setup_window()
        self.load_config()
        self.create_ui()
        self.is_scraping = False

    def setup_window(self):
        """设置主窗口"""
        self.root.title("EZBioCloud Scraper")
        self.root.geometry("900x790")
        self.root.configure(bg="#f5f6fa")
        self.root.resizable(True, True)
        self.root.minsize(900, 600)

    def load_config(self):
        """加载配置文件"""
        self.config = configparser.ConfigParser()
        self.config_file = 'config.ini'
        if os.path.exists(self.config_file):
            self.config.read(self.config_file)

    def create_ui(self):
        """创建用户界面"""
        # 顶部空白区域（50像素高）
        top_spacer = tk.Frame(self.root, height=50, bg="#f5f6fa")
        top_spacer.pack(side="top", fill="x")

        # 主容器
        main_frame = tk.Frame(self.root, bg="#f5f6fa")
        main_frame.pack(fill="both", expand=True, padx=40, pady=(0, 20))

        # 文件配置区域
        config_frame = tk.Frame(main_frame, bg="#ffffff", relief="solid", bd=1)
        config_frame.pack(fill="x", pady=(0, 15))

        tk.Label(config_frame, text="📁 输入配置", font=("Segoe UI", 12, "bold"),
                 bg="#ffffff", fg="#4285f4").pack(pady=(15, 10))

        # 文件选择行
        self.fasta_entry = self.create_file_row(config_frame, "FASTA 文件:")
        self.driver_entry = self.create_file_row(config_frame, "Chrome Driver:")
        self.output_entry = self.create_file_row(config_frame, "输出文件:", is_save=True)

        # 账户信息
        account_frame = tk.Frame(config_frame, bg="#ffffff")
        account_frame.pack(fill="x", padx=20, pady=(15, 0))

        # 用户名和密码行
        user_frame = tk.Frame(account_frame, bg="#ffffff")
        user_frame.pack(fill="x", pady=5)
        tk.Label(user_frame, text="用户名:", font=("Segoe UI", 10), bg="#ffffff", width=12, anchor="w").pack(
            side="left")
        self.username_entry = tk.Entry(user_frame, relief="solid", bd=1, width=25, font=("Segoe UI", 9))
        self.username_entry.pack(side="left", padx=(10, 20))

        tk.Label(user_frame, text="密码:", font=("Segoe UI", 10), bg="#ffffff", width=8, anchor="w").pack(side="left")
        self.password_entry = tk.Entry(user_frame, relief="solid", bd=1, width=25, show="*", font=("Segoe UI", 9))
        self.password_entry.pack(side="left", padx=(10, 0))

        # 等待时间和记住密码
        settings_frame = tk.Frame(account_frame, bg="#ffffff")
        settings_frame.pack(fill="x", pady=5)
        tk.Label(settings_frame, text="等待时间(秒):", font=("Segoe UI", 10), bg="#ffffff", width=12, anchor="w").pack(
            side="left")
        self.wait_time_entry = tk.Entry(settings_frame, relief="solid", bd=1, width=8, font=("Segoe UI", 9))
        self.wait_time_entry.insert(0, "3")
        self.wait_time_entry.pack(side="left", padx=(10, 20))

        self.remember_var = tk.IntVar()
        remember_cb = tk.Checkbutton(settings_frame, text="记住登录信息", variable=self.remember_var,
                                     font=("Segoe UI", 10), bg="#ffffff", relief="flat",
                                     borderwidth=2, highlightthickness=0)
        remember_cb.pack(side="left", padx=(10, 0))

        # 后台模式复选框
        self.headless_var = tk.IntVar()
        self.headless_cb = tk.Checkbutton(settings_frame, text="启用后台模式", variable=self.headless_var,
                                          font=("Segoe UI", 10), bg="#ffffff", relief="flat",
                                          borderwidth=2, highlightthickness=0,
                                          command=lambda: self.screenshot_cb.config(
                                              state="normal" if self.headless_var.get() else "disabled"))
        self.headless_cb.pack(side="left", padx=(20, 0))

        # 截图调试复选框（默认禁用）
        self.screenshot_var = tk.IntVar()
        self.screenshot_cb = tk.Checkbutton(settings_frame, text="截图调试", variable=self.screenshot_var,
                                            font=("Segoe UI", 10), bg="#ffffff", relief="flat",
                                            borderwidth=2, highlightthickness=0, state="disabled")
        self.screenshot_cb.pack(side="left", padx=(20, 0))

        tk.Label(config_frame, text="", bg="#ffffff").pack(pady=1)  # 底部间距

        # 进度监控区域
        progress_frame = tk.Frame(main_frame, bg="#ffffff", relief="solid", bd=1)
        progress_frame.pack(fill="x", pady=(0, 15))

        tk.Label(progress_frame, text="📊 进度监控", font=("Segoe UI", 12, "bold"),
                 bg="#ffffff", fg="#4285f4").pack(pady=(15, 10))

        prog_container = tk.Frame(progress_frame, bg="#ffffff")
        prog_container.pack(fill="x", padx=20, pady=(0, 20))

        # 比对进度
        self.progress_label1 = tk.Label(prog_container, text="比对进度: 0/0 (0.00%)",
                                        font=("Segoe UI", 10), bg="#ffffff", fg="#2c3e50", anchor="w")
        self.progress_label1.pack(fill="x", pady=(0, 5))

        style = ttk.Style()
        style.configure("Blue.Horizontal.TProgressbar", background='#4285f4', troughcolor='#e8f0fe')
        self.progress_bar1 = ttk.Progressbar(prog_container, length=800, mode='determinate',
                                             style='Blue.Horizontal.TProgressbar')
        self.progress_bar1.pack(fill="x", pady=(0, 10))

        # 导出进度
        self.progress_label2 = tk.Label(prog_container, text="数据导出进度: 0/0 (0.00%)",
                                        font=("Segoe UI", 10), bg="#ffffff", fg="#2c3e50", anchor="w")
        self.progress_label2.pack(fill="x", pady=(0, 5))

        style.configure("Green.Horizontal.TProgressbar", background='#28a745', troughcolor='#d4edda')
        self.progress_bar2 = ttk.Progressbar(prog_container, length=800, mode='determinate',
                                             style='Green.Horizontal.TProgressbar')
        self.progress_bar2.pack(fill="x", pady=(0, 10))

        # 状态标签
        self.status_label = tk.Label(prog_container, text="就绪", font=("Segoe UI", 8), bg="#ffffff", fg="#28a745")
        self.status_label.pack()

        # 运行时间标签
        self.runtime_label = tk.Label(prog_container, text="运行时间: 00H:00M:00S", font=("Segoe UI", 8), bg="#ffffff",
                                      fg="#28a745")
        self.runtime_label.pack(pady=(2, 0))

        # 开始按钮
        self.start_btn = tk.Button(main_frame, text="开始爬取", command=self.start_scraping,
                                   font=("Segoe UI", 13), bg="#4285f4", fg="white",
                                   relief="flat", bd=0, padx=5, pady=3, cursor="hand2")
        self.start_btn.pack(pady=0.1)

        # 加载保存的凭据
        self.load_saved_credentials()

        footer_label = tk.Label(self.root, text="by 0x4c4d59", font=("Segoe UI", 8),
                                bg="#f5f6fa", fg="#888888")
        footer_label.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-5)

    def toggle_screenshot_cb(self):
        """启用/禁用截图调试复选框"""
        if self.headless_var.get():
            self.screenshot_cb.config(state="normal")
        else:
            self.screenshot_cb.config(state="disabled")
            self.screenshot_var.set(0)

    def create_file_row(self, parent, label_text, is_save=False):
        """创建文件选择行"""
        row_frame = tk.Frame(parent, bg="#ffffff")
        row_frame.pack(fill="x", padx=20, pady=8)

        # 标签
        label = tk.Label(row_frame, text=label_text, font=("Segoe UI", 10),
                         bg="#ffffff", fg="#2c3e50", width=12, anchor="w")
        label.pack(side="left")

        # 输入框
        entry = tk.Entry(row_frame, relief="solid", bd=1, width=60, font=("Segoe UI", 9))
        entry.pack(side="left", fill="x", expand=True, padx=(10, 10))

        # 浏览按钮
        if is_save:
            browse_btn = tk.Button(row_frame, text="📁", command=lambda: self.browse_save_file(entry),
                                   font=("Segoe UI", 12), bg="#f5f6fa", relief="flat", bd=1,
                                   width=3, cursor="hand2")
        else:
            browse_btn = tk.Button(row_frame, text="📁", command=lambda: self.browse_file(entry),
                                   font=("Segoe UI", 12), bg="#f5f6fa", relief="flat", bd=1,
                                   width=3, cursor="hand2")
        browse_btn.pack(side="right")

        return entry

    def browse_file(self, entry):
        """浏览文件"""
        file_path = filedialog.askopenfilename()
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

    def browse_save_file(self, entry):
        """浏览保存文件"""
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

    def load_saved_credentials(self):
        """加载保存的凭据"""
        if 'credentials' in self.config:
            if 'username' in self.config['credentials']:
                self.username_entry.insert(0, self.config['credentials']['username'])
            if 'password' in self.config['credentials']:
                self.password_entry.insert(0, self.config['credentials']['password'])
                self.remember_var.set(1)

    def save_credentials(self):
        """保存凭据"""
        if self.remember_var.get():
            if 'credentials' not in self.config:
                self.config.add_section('credentials')
            self.config['credentials']['username'] = self.username_entry.get()
            self.config['credentials']['password'] = self.password_entry.get()
            with open(self.config_file, 'w') as configfile:
                self.config.write(configfile)

    def update_progress(self, stage, current, total, status_text):
        """更新进度条和标签"""

        def update_ui():
            if total > 0:
                percentage = (current / total) * 100
                if stage == "比对进度":
                    self.progress_bar1['value'] = percentage
                    self.progress_label1.config(text=f"{stage}: {current}/{total} ({percentage:.2f}%)")
                elif stage == "数据导出进度":
                    self.progress_bar2['value'] = percentage
                    self.progress_label2.config(text=f"{stage}: {current}/{total} ({percentage:.2f}%)")
            self.status_label.config(text=status_text)
            self.root.update_idletasks()

        self.root.after(0, update_ui)

    def update_runtime(self):
        """动态更新运行时间，显示格式 HH:MM:SS"""
        elapsed = int(time.time() - self.start_time)
        hours = elapsed // 3600
        minutes = (elapsed % 3600) // 60
        seconds = elapsed % 60
        self.runtime_label.config(text=f"运行时间: {hours:02d}H:{minutes:02d}M:{seconds:02d}S")
        if self.is_scraping:
            self.root.after(1000, self.update_runtime)

    def start_scraping(self):
        """开始爬取"""
        if self.is_scraping:
            messagebox.showwarning("警告", "爬取正在进行中，请等待完成！")
            return

        # 获取输入参数
        fasta_path = self.fasta_entry.get()
        driver_path = self.driver_entry.get()
        username = self.username_entry.get()
        password = self.password_entry.get()
        output_path = self.output_entry.get()

        try:
            wait_time = int(self.wait_time_entry.get())
        except ValueError:
            messagebox.showwarning("警告", "等待时间必须是整数！")
            return

        if not all([fasta_path, driver_path, username, password, output_path]):
            messagebox.showwarning("警告", "请填写所有字段！")
            return

        # 重置进度条
        self.progress_bar1['value'] = 0
        self.progress_bar2['value'] = 0
        self.progress_label1.config(text="比对进度: 0/0 (0.00%)")
        self.progress_label2.config(text="数据导出进度: 0/0 (0.00%)")
        self.status_label.config(text="准备开始...")

        # 保存凭据
        self.save_credentials()

        # ✅ 设置爬取状态
        self.is_scraping = True
        # ✅ 记录开始时间
        self.start_time = time.time()
        # ✅ 启动动态计时
        self.update_runtime()

        # 在新线程中运行爬虫
        threading.Thread(target=self.scrape_ezbiocloud,
                         args=(fasta_path, driver_path, username, password, output_path, wait_time),
                         daemon=True).start()

    def get_page_settings(self, total_sequences):
        """根据序列数量确定分页设置"""
        if total_sequences <= 25:
            return None, 1  # 不需要分页
        elif total_sequences <= 50:
            return 50, math.ceil(total_sequences / 50)
        else:
            return 100, math.ceil(total_sequences / 100)

    def set_page_size(self, driver, wait, page_size):
        """设置每页显示数量"""
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            dropdown_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@class, 'btn-straintable dropdown-toggle')]")))
            dropdown_button.click()
            time.sleep(1)
            option = wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//ul[@class='dropdown-menu']//a[contains(text(), '{page_size}')]")))
            option.click()
            time.sleep(3)
        except Exception as e:
            print(f"设置每页显示条数时出错: {e}")

    def collect_page_data(self, driver):
        """收集当前页面数据"""
        try:
            names = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[3]")
            top_hit_taxon = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[4]")
            top_hit_strain = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[5]")
            similarity = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[6]")
            top_hit_taxonomy = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[7]")
            completeness = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[8]")

            return [[elem.text for elem in data] for data in
                    zip(names, top_hit_taxon, top_hit_strain, similarity, top_hit_taxonomy, completeness)]
        except Exception as e:
            print(f"收集页面数据时出错: {e}")
            return []

    def scrape_ezbiocloud(self, fasta_path, driver_path, username, password, output_path, wait_time):
        """执行爬虫操作"""
        driver = None
        try:
            # 启动浏览器并登录
            service = Service(driver_path)
            options = Options()
            if self.headless_var.get():
                # 启用 headless
                options.add_argument("--headless")
                options.add_argument("--disable-gpu")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--window-size=1920,1080")
                options.add_argument("--disable-software-rasterizer")
                options.add_argument("--disable-blink-features=AutomationControlled")
                # 修改 User-Agent (换成你真实 Chrome UA)
                options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                     "AppleWebKit/537.36 (KHTML, like Gecko) "
                                     "Chrome/140.0.7339.128 Safari/537.36")
            driver = webdriver.Chrome(service=service, options=options)

            # 隐藏 webdriver 特征（防反爬）
            if self.headless_var.get():
                driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": """
                                    Object.defineProperty(navigator, 'webdriver', { get: () => undefined })
                                """
                })

            driver.get("https://eztaxon-e.ezbiocloud.net/identify")

            if self.headless_var.get() and self.screenshot_var.get():
                driver.save_screenshot("debug.png")
                print("已保存 debug.png 截图，可检查 headless 渲染效果")

            wait = WebDriverWait(driver, 10)

            # 登录
            username_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@type='email'])[1]")))
            username_input.send_keys(username)
            password_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@type='password'])[1]")))
            password_input.send_keys(password)
            login_button = driver.find_element(By.XPATH, "//button[contains(text(),'Login')]")
            login_button.click()
            time.sleep(3)

            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'16S-based ID')]"))).click()
            time.sleep(2)

            # 读取FASTA文件
            sequences = []
            with open(fasta_path, 'r') as f:
                file_content = f.read().strip().split('>')[1:]
                for seq in file_content:
                    lines = seq.split('\n')
                    seq_id = lines[0].strip()
                    sequence = ''.join(line.strip() for line in lines[1:])
                    sequences.append((seq_id, sequence))

            total_sequences = len(sequences)
            self.update_progress("比对进度", 0, total_sequences, "准备开始比对...")

            # 逐个处理序列
            for idx, (seq_id, sequence) in enumerate(sequences, 1):
                self.update_progress("比对进度", idx - 1, total_sequences, f"正在比对序列: {seq_id}")

                wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Identify new sequence')]"))).click()
                time.sleep(2)

                id_input = wait.until(EC.presence_of_element_located((By.ID, "sequenceName")))
                id_input.clear()
                id_input.send_keys(seq_id)

                seq_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//textarea)[1]")))
                seq_input.clear()
                seq_input.send_keys(sequence)

                next_button = driver.find_element(By.XPATH, "//button[contains(text(),'Next')]")
                next_button.click()
                time.sleep(2)

                submit_button = driver.find_element(By.XPATH, "//button[contains(text(),'Submit')]")
                submit_button.click()
                time.sleep(5)

                wait.until(EC.presence_of_element_located((By.XPATH, "//table[@id='idResultTable']")))
                time.sleep(wait_time)

                self.update_progress("比对进度", idx, total_sequences, f"已完成序列: {seq_id}")

            self.update_progress("比对进度", total_sequences, total_sequences, "所有序列比对完成！")

            # 根据序列数量设置分页
            page_size, total_pages = self.get_page_settings(total_sequences)

            all_results = []
            if page_size:
                # 需要分页处理
                self.set_page_size(driver, wait, page_size)

                for page in range(1, total_pages + 1):
                    self.update_progress("数据导出进度", page - 1, total_pages, f"正在收集第{page}页数据...")
                    wait.until(EC.presence_of_element_located((By.XPATH, "//table[@id='idResultTable']")))
                    time.sleep(2)

                    page_data = self.collect_page_data(driver)
                    all_results.extend(page_data)

                    if page < total_pages:
                        try:
                            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                            time.sleep(1)
                            next_page_button = driver.find_element(By.XPATH, "//li[@class='page-next']//a")
                            if next_page_button.is_enabled():
                                next_page_button.click()
                                time.sleep(3)
                            else:
                                break
                        except Exception as e:
                            print(f"翻页出错: {e}")
                            break
            else:
                # 单页处理
                self.update_progress("数据导出进度", 0, 1, "收集数据...")
                wait.until(EC.presence_of_element_located((By.XPATH, "//table[@id='idResultTable']")))
                time.sleep(2)
                all_results = self.collect_page_data(driver)
                self.update_progress("数据导出进度", 1, 1, "数据收集完成")

            # 生成DataFrame并保存
            df = pd.DataFrame(all_results, columns=['Name', 'Top-hit taxon', 'Top-hit strain', 'Similarity (%)',
                                                    'Top-hit taxonomy', 'Completeness (%)'])

            # 添加序列信息
            sequence_dict = {seq_id: seq for seq_id, seq in sequences}
            df['Sequence'] = df['Name'].map(sequence_dict)
            df.to_excel(output_path, index=False)

            # 补全导出进度到 100%
            if total_sequences <= 25:
                self.update_progress("数据导出进度", 1, 1, "数据导出完成！")
            else:
                self.update_progress("数据导出进度", total_pages, total_pages, "数据导出完成！")

            self.root.after(0, lambda: messagebox.showinfo("完成",
                                                           f"数据爬取完成！结果已保存。\n共处理 {total_sequences} 个序列\n共收集 {len(all_results)} 条结果"))

        except Exception as e:
            err_msg = f"发生错误: {e}"
            print(err_msg)
            self.root.after(0, lambda m=err_msg: messagebox.showerror("错误", m))
        finally:
            if driver:
                driver.quit()
            self.is_scraping = False

    def run(self):
        """启动应用程序"""
        self.root.mainloop()


if __name__ == "__main__":
    app = EZBioCloudScraper()
    app.run()