# rpa_invoice_downloader.py

import time
import logging

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# cấu hình logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
class InvoiceDownloader:
    def __init__(self, driver_path: str = None):
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        prefs = {"download.prompt_for_download": False}
        chrome_options.add_experimental_option("prefs", prefs)

        service = Service(driver_path) if driver_path else Service()
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.wait = WebDriverWait(self.driver, 10)

    # Mở trang tra cứu hóa đơn
    def open_lookup_page(self):
        url = "https://www.meinvoice.vn/tra-cuu"
        self.driver.get(url)
        logging.info("Mở trang tra cứu hóa đơn.")

    # Nhập mã tra cứu hóa đơn
    def enter_lookup_code(self, code: str):
        try:
            input_box = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'input[placeholder="Nhập mã tra cứu hóa đơn"]')
                )
            )
            input_box.clear()
            input_box.send_keys(code)
            logging.info(f"Nhập mã tra cứu: {code}")
        except TimeoutException:
            logging.error("Không tìm thấy ô nhập mã tra cứu.")
            raise

    # Nhấn nút tra cứu
    def click_search(self):
        try:
            search_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, "btnSearchInvoice"))
            )
            search_button.click()
            logging.info("Nhấn nút tra cứu.")
        except TimeoutException:
            logging.error("Không tìm thấy nút tìm kiếm.")
            raise

    # Tải hóa đơn dưới dạng PDF
    def download_invoice_pdf(self):
        try:
            # Mở menu tải hóa đơn
            open_menu = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "span.download-invoice"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", open_menu)
            time.sleep(1)
            open_menu.click()
            logging.info("Mở menu tải hóa đơn.")

            # Click tải hóa đơn dạng PDF
            download_pdf = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div.dm-item.pdf.txt-download-pdf"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", download_pdf)
            time.sleep(1)
            download_pdf.click()
            logging.info("Đã tải hóa đơn dạng PDF thành công.")
            time.sleep(3)
        except TimeoutException:
            logging.error("Không tìm thấy nút tải hóa đơn.")
            raise

    def close(self):
        self.driver.quit()
        logging.info("Đóng trình duyệt.")


def main():
    lookup_code = "mã muôn tra cứu" 
    downloader = InvoiceDownloader(driver_path=r"D:\chromedriver-win64\chromedriver-win64\chromedriver-win64\chromedriver.exe")
    try:
        downloader.open_lookup_page()
        downloader.enter_lookup_code(lookup_code)
        downloader.click_search()
        downloader.download_invoice_pdf()
    except Exception as e:
        logging.error(f"Lỗi xảy ra: {e}")
    finally:
        downloader.close()


if __name__ == "__main__":
    main()
