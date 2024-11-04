import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import configparser


# Function to perform scraping
def scrape_ezbiocloud(fasta_path, driver_path, username, password, output_path, wait_time):
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service)

    driver.get("https://eztaxon-e.ezbiocloud.net/identify")
    wait = WebDriverWait(driver, 10)

    try:
        username_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@type='email'])[1]")))
        username_input.send_keys(username)

        password_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@type='password'])[1]")))
        password_input.send_keys(password)

        login_button = driver.find_element(By.XPATH, "//button[contains(text(),'Login')]")
        login_button.click()

        time.sleep(3)

        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'16S-based ID')]"))).click()
        time.sleep(2)

        results = []
        with open(fasta_path, 'r') as f:
            sequences = f.read().strip().split('>')[1:]

            for seq in sequences:
                lines = seq.split('\n')
                seq_id = lines[0].strip()
                sequence = ''.join(line.strip() for line in lines[1:])

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
                time.sleep(wait_time)  # 使用用户自定义的等待时间

                names = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[3]")
                top_hit_taxon = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[4]")
                top_hit_strain = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[5]")
                similarity = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[6]")
                top_hit_taxonomy = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[7]")
                completeness = driver.find_elements(By.XPATH, "//table[@id='idResultTable']//td[8]")

            for data in zip(names, top_hit_taxon, top_hit_strain, similarity, top_hit_taxonomy, completeness):
                result_row = [elem.text for elem in data]
                results.append(result_row)

            # 生成 DataFrame
            df = pd.DataFrame(results, columns=['Name', 'Top-hit taxon', 'Top-hit strain', 'Similarity (%)',
                                                'Top-hit taxonomy', 'Completeness (%)'])

            # 重新读取FASTA文件，将编号和Name匹配，添加序列到数据中
            sequence_dict = {}
            with open(fasta_path, 'r') as f:
                sequences = f.read().strip().split('>')[1:]
                for seq in sequences:
                    lines = seq.split('\n')
                    seq_id = lines[0].strip()  # 序列编号
                    sequence2 = ''.join(line.strip() for line in lines[1:])  # 序列内容
                    sequence_dict[seq_id] = sequence2

            # 根据Name列匹配FASTA文件中的序列
            df['Sequence'] = df['Name'].map(sequence_dict)

            # 保存到Excel文件
            df.to_excel(output_path, index=False)

    except Exception as e:
        print(f"发生错误: {e}")
        messagebox.showerror("Error", f"发生错误: {e}")

    finally:
        driver.quit()
        messagebox.showinfo("完成", "数据爬取完成！结果已保存。")


# Function to open file dialogs and start the scraping process
def start_scraping():
    fasta_path = fasta_entry.get()
    driver_path = driver_entry.get()
    username = username_entry.get()
    password = password_entry.get()
    output_path = output_entry.get()

    # 获取用户自定义的等待时间
    try:
        wait_time = int(wait_time_entry.get())
    except ValueError:
        messagebox.showwarning("警告", "等待时间必须是整数！")
        return

    if not all([fasta_path, driver_path, username, password, output_path]):
        messagebox.showwarning("警告", "请填写所有字段！")
        return

    save_credentials(username, password)
    scrape_ezbiocloud(fasta_path, driver_path, username, password, output_path, wait_time)


def browse_file(entry):
    file_path = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, file_path)


def browse_save_file(entry):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, file_path)


# Save credentials
def save_credentials(username, password):
    if remember_var.get():
        config['credentials'] = {'username': username, 'password': password}
        with open(config_file, 'w') as configfile:
            config.write(configfile)


# 读取配置文件
config = configparser.ConfigParser()
config_file = 'config.ini'
if os.path.exists(config_file):
    config.read(config_file)

# Set up the GUI using tkinter
root = tk.Tk()
root.title("EZBioCloud Scraper")

# 定义记住密码的复选框变量
remember_var = tk.IntVar()

# Define and place GUI elements
tk.Label(root, text="FASTA 文件路径:").grid(row=0, column=0)
fasta_entry = tk.Entry(root, width=50)
fasta_entry.grid(row=0, column=1)
fasta_button = tk.Button(root, text="浏览", command=lambda: browse_file(fasta_entry))
fasta_button.grid(row=0, column=2)

tk.Label(root, text="Chrome Driver 路径:").grid(row=1, column=0)
driver_entry = tk.Entry(root, width=50)
driver_entry.grid(row=1, column=1)
driver_button = tk.Button(root, text="浏览", command=lambda: browse_file(driver_entry))
driver_button.grid(row=1, column=2)

tk.Label(root, text="用户名:").grid(row=2, column=0)
username_entry = tk.Entry(root, width=50)
username_entry.grid(row=2, column=1)
if 'credentials' in config and 'username' in config['credentials']:
    username_entry.insert(0, config['credentials']['username'])

tk.Label(root, text="密码:").grid(row=3, column=0)
password_entry = tk.Entry(root, show="*", width=50)
password_entry.grid(row=3, column=1)
if 'credentials' in config and 'password' in config['credentials']:
    password_entry.insert(0, config['credentials']['password'])

remember_checkbox = tk.Checkbutton(root, text="记住密码", variable=remember_var)
remember_checkbox.grid(row=3, column=2)

tk.Label(root, text="输出文件路径:").grid(row=4, column=0)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=4, column=1)
output_button = tk.Button(root, text="浏览", command=lambda: browse_save_file(output_entry))
output_button.grid(row=4, column=2)

# 添加等待时间输入框
tk.Label(root, text="比对后等待时间 (秒):").grid(row=5, column=0)
wait_time_entry = tk.Entry(root, width=5)
wait_time_entry.grid(row=5, column=1)
wait_time_entry.insert(0, "3")  # 默认值为3秒

start_button = tk.Button(root, text="开始爬取", command=start_scraping)
start_button.grid(row=7, column=1)

root.mainloop()
