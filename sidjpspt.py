import pandas as pd
from datetime import datetime
import os, sys
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, JavascriptException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from PyQt5.QtWidgets import QApplication, QFileDialog, QTextEdit, QMessageBox, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel, QLineEdit, QComboBox, QDialog


class MyDownloader(QMainWindow):
    def __init__ (self):
        super().__init__()
        current_date = datetime.now()
        formatted_date = current_date.strftime("%d-%m-%Y-%H_%M")
        folder_name = f"Data_SIDJP_{formatted_date}"
        parent_folder = r"A:\proyek\scrapper sidjp\Data Downloadan"
        folder_path = os.path.join(parent_folder, folder_name)        
        os.makedirs(folder_path)
        self.path = folder_path

        self.init_webdriver()
        self.initUI() 
    
    def init_webdriver(self):
        executable_path = r"A:\proyek\mpninfo update\chromedriver.exe" #Windows
        service = Service(executable_path)

        options = webdriver.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        # options.add_argument("--headless=new")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False) 
        options.add_argument('--ignore-certificate-errors')

        # options.add_experimental_option("prefs",{"download.default_directory":path}) 

        self.driver = webdriver.Chrome(service=service, options=options)

        # options.add_argument("--headless=new")

    def get_folder(self):
        caption = "Select folder"
        initial_dir = ""
        dialog = QFileDialog()
        dialog.setWindowTitle(caption)
        dialog.setDirectory(initial_dir)
        dialog.setFileMode(QFileDialog.FileMode.Directory)
        if dialog.exec():
            self.selected_folder = dialog.selectedFiles()[0]
            print("Selected Folder:", self.selected_folder)

    def initUI(self):
        self.setWindowTitle("SI.050.DP3")
        self.setGeometry(400, 400, 250, 100)
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        layout = QVBoxLayout()
        self.central_widget.setLayout(layout)
        self.username_label = QLabel("Username:")
        self.username_input = QLineEdit()
        layout.addWidget(self.username_label)
        layout.addWidget(self.username_input)
        self.password_label = QLabel("Password:")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.password_label)
        layout.addWidget(self.password_input)
        self.login_button = QPushButton("Login")
        layout.addWidget(self.login_button)
        self.login_button.clicked.connect(self.on_login_sidjp_click)
        self.exit_button = QPushButton("Exit")
        layout.addWidget(self.exit_button) 
        self.exit_button.clicked.connect(self.close)

    def on_login_sidjp_click(self):
        username = self.username_input.text()
        password = self.password_input.text()
        try:
            self.start_sidjp(username, password)
            current_url = self.driver.current_url
            if current_url == "http://sidjp:7777/SIDJP/sipt_web.main":
                dlg = QMessageBox(self)
                dlg.setWindowTitle("Notification")
                dlg.setIcon(QMessageBox.Information)  # Menetapkan ikon informasi
                dlg.setText("Berhasil Login")
                dlg.addButton(QMessageBox.Ok)
                dlg.exec_()
                self.main_window = MainWindow(self.driver)
                self.main_window.show()
                self.close()
            else:
                dlg = QMessageBox(self)
                dlg.setWindowTitle("Notification")
                dlg.setIcon(QMessageBox.Critical)  # Menetapkan ikon informasi
                dlg.setText("Gagal Login")
                dlg.addButton(QMessageBox.Ok)  # Menambahkan tombol OK
                dlg.exec_()
                print("Terjadi kesalahan:", str(e))
        except Exception as e:
            print("Terjadi kesalahan:", str(e))


    def start_sidjp(self, username, password):
        self.driver.get("http://sidjp:7777/") 
        executable_path = r"A:\proyek\mpninfo update\chromedriver.exe" #Windows
        service = Service(executable_path)
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False) 
        options.add_argument('--ignore-certificate-errors')
        time.sleep(2)
        username_el = self.driver.find_element(By.XPATH, '//*[@id="table18"]/tbody/tr[3]/td/input')
        password_el = self.driver.find_element(By.XPATH, '//*[@id="table18"]/tbody/tr[5]/td/input')
        username_el.clear()
        username_el.send_keys(username)
        password_el.clear()
        password_el.send_keys(password)
        password_el.send_keys(Keys.RETURN)
        time.sleep(1)

  

class MainWindow(QMainWindow):
    def __init__(self, driver):
        super().__init__()
        self.driver = driver

        self.initUI()

    def initUI(self):
        self.setWindowTitle("SI.050.DP3")
        self.setGeometry(400, 400, 250, 150)
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        layout = QVBoxLayout()
        self.central_widget.setLayout(layout)

        self.folder_label = QLabel("Folder Path:")
        self.folder_input = QLineEdit()
        layout.addWidget(self.folder_label)
        layout.addWidget(self.folder_input)

        self.browse_button = QPushButton("Browse")
        layout.addWidget(self.browse_button)
        self.browse_button.clicked.connect(self.browse_folder)

        self.submit_button = QPushButton("Submit")
        layout.addWidget(self.submit_button)
        self.submit_button.clicked.connect(self.proses)

        self.exit_button = QPushButton("Exit")
        layout.addWidget(self.exit_button) 
        self.exit_button.clicked.connect(self.close)

        self.text_edit = QTextEdit(self)
        layout.addWidget(self.text_edit)

    def browse_folder(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select CSV File", "", "CSV Files (*.csv);;All Files (*)")
        if file_path:
            self.folder_input.setText(file_path)
            self.selected_file_path = file_path  # Simpan path folder yang dipilih

    def proses(self):
        wait = WebDriverWait(self.driver, 5)
        profil_wp = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[5]/td/table/tbody/tr/td[1]/table/tbody/tr[2]/td/div/a')))
        profil_wp.click()
        time.sleep(3)

        # Dapatkan semua handle jendela saat ini
        window_handles = self.driver.window_handles
        # Beralih ke tab baru (jendela terakhir dalam daftar)
        new_tab_handle = window_handles[-1]
        self.driver.switch_to.window(new_tab_handle)
        time.sleep(3)

        df_sumber = pd.read_csv(self.selected_file_path, sep=";", dtype=str)

        for index, row in df_sumber.iterrows():
            try:
                npwp = row['npwp']
                kpp = row['kpp']
                cabang = row['cabang']
                tahun = row['tahun']

                self.cari_data_per_wp(npwp,kpp,cabang,tahun)
            except IndexError:
                print("IndexError: list index out of range. Lanjut ke item berikutnya.")
                window_handles = self.driver.window_handles
                tab_to_close = window_handles[1]

                self.driver.switch_to.window(tab_to_close)
                self.driver.close()

                self.driver.switch_to.window(window_handles[0])

                main = 'http://sidjp:7777/SIDJP/sipt_web.main'
                self.driver.get(main)
                time.sleep(1)

                wait = WebDriverWait(self.driver, 30)
                profil_wp = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[5]/td/table/tbody/tr/td[1]/table/tbody/tr[2]/td/div/a')))
                profil_wp.click()
                window_handles = self.driver.window_handles
                # Beralih ke tab baru (jendela terakhir dalam daftar)
                new_tab_handle = window_handles[-1]
                self.driver.switch_to.window(new_tab_handle)
                time.sleep(3) 
                continue
            except NoSuchElementException:
                print("Bentuk SPT beda")
                window_handles = self.driver.window_handles
                tab_to_close = window_handles[1]

                self.driver.switch_to.window(tab_to_close)
                self.driver.close()

                self.driver.switch_to.window(window_handles[0])

                main = 'http://sidjp:7777/SIDJP/sipt_web.main'
                self.driver.get(main)
                time.sleep(1)

                wait = WebDriverWait(self.driver, 30)
                profil_wp = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[5]/td/table/tbody/tr/td[1]/table/tbody/tr[2]/td/div/a')))
                profil_wp.click()

                window_handles = self.driver.window_handles
                # Beralih ke tab baru (jendela terakhir dalam daftar)
                new_tab_handle = window_handles[-1]
                self.driver.switch_to.window(new_tab_handle)
                time.sleep(3) 
                continue

    def cari_data_per_wp(self,npwp,kpp,cabang,tahun):

        wait = WebDriverWait(self.driver, 30)
        time.sleep(3)
        npwp_value = npwp
        kpp_value = kpp
        cabang_value = cabang
        tahun_value = tahun
        npwp = self.driver.find_element(By.XPATH, '//*[@id="AutoNumber3"]/tbody/tr/td/table/tbody/tr[2]/td[3]/input[1]')
        kpp = self.driver.find_element(By.XPATH, '//*[@id="AutoNumber3"]/tbody/tr/td/table/tbody/tr[2]/td[3]/input[2]')
        cabang = self.driver.find_element(By.XPATH, '//*[@id="AutoNumber3"]/tbody/tr/td/table/tbody/tr[2]/td[3]/input[3]')
        npwp.clear()
        npwp.send_keys(npwp_value)
        kpp.clear()
        kpp.send_keys(kpp_value)
        cabang.clear()
        cabang.send_keys(cabang_value)
        cari = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[10]/td/span/a/img')))
        cari.click()
        time.sleep(1)

        npwp_link_el = self.driver.find_element(By.CSS_SELECTOR, 'body > form > table > tbody > tr:nth-child(9) > td > table > tbody > tr:nth-child(2) > td:nth-child(2) > span > a')
        npwp_id_link = npwp_link_el.get_attribute('href')
        npwp_id = npwp_id_link.split('idcabang=')[1].split('&')[0]
        time.sleep(1)

        detil_pelaporan = 'http://sidjp:7777/SIDJP/PPLSTSPT_MODV3?idcabang={}&tahun={}'.format(npwp_id, tahun_value)
        self.driver.get(detil_pelaporan)
        time.sleep(1)

        detil_pelaporan_table = self.driver.find_element(By.CSS_SELECTOR, 'body > iframe')
        url_detil_pelaporan = detil_pelaporan_table.get_attribute('src')
        self.driver.get(url_detil_pelaporan)
        time.sleep(1)

        # OP
        spt_tahunan = self.driver.find_element(By.CSS_SELECTOR, 'body > table > tbody > tr:nth-child(1) > td:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(2) > table > tbody > tr:nth-child(1) > td:nth-child(3) > span > a:nth-child(7)')
        spt_tahunan.click()
        time.sleep(1)

        halaman = self.driver.find_element(By.CSS_SELECTOR, 'body > table > tbody > tr:nth-child(5) > td:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(5) > span > a')
        href_halaman = halaman.get_attribute('href')
        self.driver.get(href_halaman)
        time.sleep(1)


        url_terbaru = self.driver.current_url
        href_halaman_spt = url_terbaru.split('idspt=')[1].split('&')[0]
        kuncian = url_terbaru.split('&sess=')[1]

        formulir_1770_lampiran_1 = 'http://sidjp:7777/SIDJP/spt/spt_profile_pphop_2014.f113216l11?idspt={}&sess={}'.format(href_halaman_spt, kuncian)
        formulir_1770_lampiran_4 = 'http://sidjp:7777/SIDJP/spt/spt_profile_pphop_2022.f113216l4?idspt={}&sess={}'.format(href_halaman_spt, kuncian)


        self.driver.get(formulir_1770_lampiran_1)
        time.sleep(1)

        peredaran_usaha = self.driver.find_element(By.XPATH, '//*[@id="table1"]/tbody/tr[45]/td[6]').text
        hpp = self.driver.find_element(By.XPATH, '//*[@id="table1"]/tbody/tr[48]/td[6]').text
        laba_rugi = self.driver.find_element(By.XPATH, '//*[@id="table1"]/tbody/tr[51]/td[6]').text
        biaya = self.driver.find_element(By.XPATH, '//*[@id="table1"]/tbody/tr[54]/td[6]').text

        time.sleep(1)
        self.driver.get(formulir_1770_lampiran_4)
        time.sleep(1)

        harta = self.driver.find_element(By. XPATH, '//*[@id="table1"]/tbody/tr[25]/td[6]/font').text
        utang = self.driver.find_element(By. XPATH, '//*[@id="table1"]/tbody/tr[36]/td[7]').text
        
        # Menyusun data ke dalam DataFrame
        data = {
            "PEREDARAN USAHA": [peredaran_usaha],
            "HPP": [hpp],
            "LABA RUGI": [laba_rugi],
            "BIAYA": [biaya],
            "HARTA": [harta],
            "UTANG": [utang],
        }

        df = pd.DataFrame(data, index=[npwp_value+kpp_value+cabang_value])
        time.sleep(1)

        df.to_csv(f'{npwp_value}_{tahun_value}.csv')

        self.text_edit.setPlainText(f'Berhasil bentuk file {npwp_value}_{tahun_value} csv ')

        
        window_handles = self.driver.window_handles
        tab_to_close = window_handles[1]

        self.driver.switch_to.window(tab_to_close)
        self.driver.close()

        self.driver.switch_to.window(window_handles[0])

        main = 'http://sidjp:7777/SIDJP/sipt_web.main'
        self.driver.get(main)
        time.sleep(1)

        wait = WebDriverWait(self.driver, 30)
        profil_wp = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[5]/td/table/tbody/tr/td[1]/table/tbody/tr[2]/td/div/a')))
        profil_wp.click()

        # Beralih ke tab baru (jendela terakhir dalam daftar)
        window_handles = self.driver.window_handles
        new_tab_handle = window_handles[-1]
        self.driver.switch_to.window(new_tab_handle)

        

def main():
    app = QApplication(sys.argv)
    window = MyDownloader()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()