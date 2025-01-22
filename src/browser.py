from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from pywinauto.application import Application
from time import sleep
import os
from pynput.keyboard import Controller, Key
import pyperclip
from bs4 import BeautifulSoup

from config import *
from utils import *

import sys

# thay đôi môi trường tiếng Việt
sys.stdout.reconfigure(encoding="utf-8")

# tìm vị trí ô tin nhắn
position_message_box = [1229, 1004]

# Các XPath cố định
# WHATSAPP
XPATHS_WHATSAPP = {
    "message_box": "//div[@contenteditable='true' and @data-tab='10']",
    "search_box": "//div[@contenteditable='true']",  # "search_box": "//div[@contenteditable='true']" dia chi cu
    "result_list": "//div[@role='grid']//span[@title]",
    "send_image_button": '//*[@id="app"]/div/div[3]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div',
    "attached_button": '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button',
    "image_input": '//input[@type="file" and @accept="image/*"]',
    "send_button": '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div',
    "audio_dow_icon": "//span[@data-icon='audio-download']",
    "caption_image": '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]/p',
    "group_title": '//*[@id="main"]/header/div[2]/div[1]/div/span',
    "join_group": '//*[@id="action-button"]',
    "use_web": '//*[@id="fallback_block"]/div/h4/a',
    "group_name_join_chat": '//*[@id="main_block"]/h3',
    "send_button_an": '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[2]/button/span',
}
# OUTLOOK
XPATHS_OUTLOOK = {
    "icon_outLook": '//*[@id="ddea774c-382b-47d7-aab5-adc2139a802b"]/span',
    "new_mail_button": '//*[@id="114-group"]/div/div[1]/div/div/span/button[1]',
    "send_to_box": "/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[3]/div[1]/div/div/div/div/div[3]/div[1]/div/div[3]/div/span/span[2]/div/div[1]",
    "send_cc_box": "/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[3]/div[1]/div/div/div/div/div[3]/div[1]/div/div[4]/div/span/span[2]/div/div[1]",
    "add_subject_box": '//*[@id="docking_InitVisiblePart_0"]/div/div[3]/div[2]/span/input',
    "add_context_box": '//*[@id="editorParent_1"]/div',
    "send_mail_button": '//*[@id="splitButton-ro__primaryActionButton"]',
    "insert_button": "/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/span/div[1]/div/div/div[5]/div/button",
    "attach_file_button": "/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/div/div/div/div/div[1]/div/div/div/div[1]/button",
    "browse_this_computer": "/html/body/div[2]/div[3]/div/div/div/div/div/div/ul/li/div/ul/li[1]/button",
    "send_mail_button": "/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[3]/div[1]/div/div/div/div/div[2]/div[1]/button[1]",  #'//*[@id="splitButton-ro__primaryActionButton"]',
    "popoup_mail": '//*[@id="popoutCompose"]',
}
XPATHS_OUTLOOK_POPUP = {
    "send_to_box": '//*[@id="docking_InitVisiblePart_0"]/div/div[3]/div[1]/div/div[3]/div/span/span[2]/div/div[1]',
    "send_cc_box": '//*[@id="docking_InitVisiblePart_0"]/div/div[3]/div[1]/div/div[4]/div/span/span[2]/div/div[1]',
    "subject_box": '//*[@id="docking_InitVisiblePart_0"]/div/div[3]/div[2]/span/input',
    "context_box": '//*[@id="editorParent_1"]/div',
}

XPATHS_ZALO = {
    "search_box": '//*[@id="contact-search-input"]',
    "message_box": '//*[@id="input_line_0"]',
    "send_button": '//*[@id="chat-input-container-id"]/div[2]/div[2]/div[2]',
    "img_attached_button": '//*[@id="chat-box-bar-id"]/div[1]/ul/li[2]/div/i',
    "file_attached_button": '//*[@id="chat-box-bar-id"]/div[1]/ul/li[3]/div',
}


class BrowserManager:
    def __init__(self):
        self.driver = None

    def start_browser(self, profile_path):
        print("dang mo trinh duyet")
        chrome_option = webdriver.ChromeOptions()
        chrome_option.add_argument(f"user-data-dir={profile_path}")
        self.driver = webdriver.Chrome(options=chrome_option)
        print("mo trinh duyet thanh cong")

    def open_url(self, url):
        self.driver.get(url)

    def switch_to_tab(self, tab_index):
        windows = self.driver.window_handles
        if tab_index < len(windows):
            self.driver.switch_to.window(windows[tab_index])
        else:
            raise Exception("Invalid tab index")

    def close(self):
        self.driver.quit()
        self.driver = None


# Lớp WhatsAppBot
class WhatsAppBot(BrowserManager):
    def access_whatsapp(self):
        self.open_url(WHATSAPP_URL)
        try:
            WebDriverWait(self.driver, 200).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '//*[@id="app"]/div/div[3]/div/div[3]/header/header/div/div[1]/h1',
                    )
                )
            )
            print("WhatsApp loaded successfully!")
        except Exception as e:
            print("Error loading WhatsApp:", e)

    def reload_web(self):
        self.driver.refresh()  # Tải lại trang
        sleep(3)  # Chờ trang tải xong

    def find_name(self, object_name):
        try:
            # Tìm kiếm ô tìm kiếm nhóm
            search_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["search_box"])
                )
            )

            # Nhập tên nhóm vào ô tìm kiếm
            search_box.click()
            search_box.send_keys(Keys.CONTROL, "a")  # Chọn tất cả nội dung
            search_box.send_keys(Keys.BACKSPACE)  # Xoá toàn bộ nội dung
            search_box.send_keys(object_name)

            try:
                # Chờ danh sách kết quả xuất hiện
                results = WebDriverWait(self.driver, 20).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, XPATHS_WHATSAPP["result_list"])
                    )
                )

                # Duyệt qua từng kết quả để tìm tên chính xác
                for result in results:
                    if result.get_attribute("title") == object_name:
                        # Cuộn đến phần tử và nhấp
                        self.driver.execute_script(
                            "arguments[0].scrollIntoView(true);", result
                        )
                        WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable(result)
                        )
                        result.click()
                        check_group = WhatsAppBot.check_group_name(self, object_name)
                        if check_group:
                            print(f"Đã tìm và mở nhóm '{object_name}' thành công!")
                            return True
                        else:
                            print(
                                f"Không tìm thấy nhóm '{object_name}' với tên chính xác."
                            )
                            return False

            except Exception as e:
                # Nếu không tìm thấy danh sách kết quả, thử nhấn Enter
                search_box.send_keys(Keys.ENTER)
                print(
                    f"Không tìm thấy danh sách kết quả. Đã thử nhấn Enter để mở nhóm '{object_name}'."
                )
                check_group = WhatsAppBot.check_group_name(self, object_name)
                if check_group:
                    print(f"Đã tìm và mở nhóm '{object_name}' thành công!")
                    return True
                else:
                    print(f"Không tìm thấy nhóm '{object_name}' với tên chính xác.")
                    return False

        except Exception as e:
            print(f"Đã xảy ra lỗi: {e}")
            return False

    def find_group_name(self, link):
        self.open_url(link)
        try:
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["group_name_join_chat"])
                )
            )
            group_name = element.text

            join_group = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["join_group"])
                )
            )
            join_group.click()

            try:
                use_web = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, XPATHS_WHATSAPP["use_web"])
                    )
                )
                # Get the HTML of the 'use_web' element
                use_web_html = use_web.get_attribute("outerHTML")

                # Parse the HTML with BeautifulSoup
                soup = BeautifulSoup(use_web_html, "html.parser")

                # Find the 'a' tag and extract the href attribute
                href_value = soup.find("a")["href"]

                self.open_url(href_value)

                try:
                    WebDriverWait(self.driver, 200).until(
                        EC.presence_of_element_located(
                            (
                                By.XPATH,
                                '//*[@id="app"]/div/div[3]/div/div[3]/header/header/div/div[1]/h1',
                            )
                        )
                    )

                    if self.check_group_name(group_name):
                        print(f"Mở nhóm [{group_name}] thành công!!!")
                        return True
                    else:
                        print("Mở nhóm thất bại")
                        return False
                except Exception as e:
                    print(e)
            except Exception as e:
                print(e)
        except Exception as e:
            print(e)

    def check_group_name(self, group_name):
        try:
            # hàm xác định group đã mở đúng không?
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["group_title"])
                )
            )
            text_content = element.text
            if text_content == group_name:
                return True
            else:
                return False

        except Exception as e:
            print(f"Không có group nào mở: {e}")

    def send_message(self, message):
        # tìm ô tin nhắn
        message_box = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, XPATHS_WHATSAPP["message_box"]))
        )
        message_box.click()
        message_box.send_keys(message)
        message_box.send_keys(Keys.ENTER)
        sleep(1)

    def send_attached_file(self, file_path):
        # gắn file đính kèm
        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_WHATSAPP["attached_button"],
                )
            )
        )
        try:
            attached_button = self.driver.find_element(
                By.XPATH, XPATHS_WHATSAPP["attached_button"]
            )
            print("đã tìm thấy nút")

            attached_button.click()
            try:
                # Chọn nút tài liệu
                file_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '//input[@type="file"]',
                        )  # Đường dẫn thường dùng để gửi tài liệu
                    )
                )
                sleep(2)
                try:
                    absolute_path = os.path.abspath(
                        file_path
                    )  # Chuyển đường dẫn thành tuyệt đối
                    file_input.send_keys(absolute_path)
                    print("xong buoc lua hinh")
                    try:
                        # nút gửi tin nhắn
                        send_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable(
                                (
                                    By.XPATH,
                                    XPATHS_WHATSAPP["send_button"],
                                )
                            )
                        )
                        send_button.click()
                        print("gui thanh cong")
                        sleep(5)
                    except Exception as e:
                        print(e)
                except Exception as e:
                    print(e)
            except Exception as e:
                print(e)
        except Exception as e:
            print(e)

    def send_attached_img_message(self, message, file_path, tag_name=None):

        message_box = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, XPATHS_WHATSAPP["message_box"]))
        )
        message_box.click()
        message_box.send_keys(Keys.CONTROL, "a")  # Chọn tất cả nội dung
        message_box.send_keys(Keys.BACKSPACE)  # Xoá toàn bộ nội dung
        message_box.send_keys(message)
        if tag_name:
            message_box.send_keys(": @")
            message_box.send_keys(remove_accents(tag_name))
            sleep(3)
            message_box.send_keys(Keys.TAB)
            sleep(2)

        # gắn file đính kèm
        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_WHATSAPP["attached_button"],
                )
            )
        )
        try:
            attached_button = self.driver.find_element(
                By.XPATH, XPATHS_WHATSAPP["attached_button"]
            )
            print("đã tìm thấy nút")

            attached_button.click()
            try:
                # Chọn nút tài liệu
                file_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]/li/div/input',
                        )
                    )
                )
                sleep(2)
                try:
                    absolute_path = os.path.abspath(
                        file_path
                    )  # Chuyển đường dẫn thành tuyệt đối
                    file_input.send_keys(absolute_path)
                    print("xong buoc lua hinh")

                    try:
                        # nút gửi tin nhắn
                        send_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable(
                                (
                                    By.XPATH,
                                    XPATHS_WHATSAPP["send_button"],
                                )
                            )
                        )
                        send_button.click()
                        print("gui thanh cong")
                        return True
                        sleep(3)
                    except Exception as e:
                        print(e)
                        return False

                except Exception as e:
                    print(e)
                    return False

            except Exception as e:
                print(e)
                return False

        except Exception as e:
            print(e)
            return False

    def send_attached_img(self, file_path):
        # gắn file đính kèm
        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_WHATSAPP["attached_button"],
                )
            )
        )
        try:
            attached_button = self.driver.find_element(
                By.XPATH, XPATHS_WHATSAPP["attached_button"]
            )
            print("đã tìm thấy nút")

            attached_button.click()
            try:
                # Chọn nút tài liệu
                file_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]/li/div/input',
                        )
                    )
                )
                sleep(2)
                try:
                    absolute_path = os.path.abspath(
                        file_path
                    )  # Chuyển đường dẫn thành tuyệt đối
                    file_input.send_keys(absolute_path)
                    print("xong buoc lua hinh")
                except Exception as e:
                    print(e)
            except Exception as e:
                print(e)
        except Exception as e:
            print(e)

    def send_message_CDBR(self, message):
        try:
            # tìm ô tin nhắn
            message_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["message_box"])
                )
            )
            message_box.click()
            message_box.send_keys(message)
            message_box.send_keys()
            message_box.send_keys(Keys.CONTROL, "v")
            sleep(5)
            try:
                # nút gửi tin nhắn
                send_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            XPATHS_WHATSAPP["send_button"],
                        )
                    )
                )
                send_button.click()
                print("gui thanh cong")
                sleep(3)
            except Exception as e:
                print(e)
        except Exception as e:
            print(f"Không tìm thấy ô tin nhắn: {e}")

    def send_Error_Notification(self, phone_number, message):
        if not phone_number.startswith("+84"):
            phone_number = "+84" + phone_number.lstrip(
                "0"
            )  # Loại bỏ số 0 đầu tiên và thêm mã quốc gia
            if phone_number.endswith(".0"):
                phone_number = phone_number.rstrip(".0")  # Loai bo dau thap phan

        self.driver.get(
            f"https://web.whatsapp.com/send?phone={phone_number}&text={message}"
        )
        try:
            send_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["send_button_an"])
                )
            )
            send_button.click()
            sleep(3)
        except Exception as e:
            print(f"lỗi gửi tin nhắn báo lỗi cho bản thân: {e}")


# Lớp ZaloBot
class ZaloBot(BrowserManager):

    def access_zalo(self):
        self.open_url(ZALO_URL)
        print("Zalo loaded successfully!")

    def find_name(self, object_name, xpath_address):
        """
        tìm tên người - nhóm theo tên và địa chỉ xpath

        Args:
            object_name (str): tên đối tượng cần tìm kiếm.
            xpath_address (str): địa chỉ xpath tương đương trên tìm kiếm
        """
        try:
            # Tìm kiếm ô tìm kiếm nhóm
            search_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["search_box"]))
            )

            # Nhập tên nhóm vào ô tìm kiếm
            search_box.click()
            search_box.clear()
            search_box.send_keys(object_name)

            # # Chờ danh sách kết quả xuất hiện
            # results = WebDriverWait(self.driver, 20).until(
            #     EC.presence_of_all_elements_located(
            #         (By.XPATH, XPATHS_WHATSAPP["result_list"])
            #     )
            # )

            # tìm tên đúng theo xpath
            object_click = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"{xpath_address}"))
            )
            object_click.click()

            print(f"Tìm {object_name} thành công!!!")

        except Exception as e:
            print(f"Đã xảy ra lỗi: {e}")
            return False

    def find_name_no_xpath(self, object_name):
        """
        tìm tên người - nhóm theo tên và địa chỉ xpath

        Args:
            object_name (str): tên đối tượng cần tìm kiếm.
            xpath_address (str): địa chỉ xpath tương đương trên tìm kiếm
        """
        try:
            # Tìm kiếm ô tìm kiếm nhóm
            search_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["search_box"]))
            )

            # Nhập tên nhóm vào ô tìm kiếm
            search_box.click()
            search_box.clear()
            search_box.send_keys(object_name)

            search_box.send_keys(Keys.ENTER)
            print(f"Tìm {object_name} thành công!!!")

        except Exception as e:
            print(f"Đã xảy ra lỗi: {e}")
            return False

    def send_message(self, message):
        # tìm ô tin nhắn
        message_box = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["message_box"]))
        )
        message_box.click()
        message_box.send_keys(message)
        message_box.send_keys(Keys.ENTER)
        sleep(1)

    def run_macro_and_send_message(self, driver, excel, macro, message):
        try:
            # excel.Application.Run(macro)
            print("Run Macro thành công!!!")

            try:
                # tìm ô tin nhắn
                message_box = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, XPATHS_ZALO["message_box"])
                    )
                )
                message_box.click()
                message_box.send_keys(message)
                message_box.send_keys(Keys.CONTROL, "V")

                try:
                    # tìm ô tin nhắn
                    send_button = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, XPATHS_ZALO["send_button"])
                        )
                    )
                    send_button.click()

                except Exception as e:
                    print(e)

            except Exception as e:
                print(e)

        except Exception as e:
            print(e)

    def send_attached_img_message(self, message, file_path, tag_name=None):
        try:
            # Kiểm tra đường dẫn ảnh
            if not os.path.isfile(file_path):
                raise FileNotFoundError(f"File không tồn tại: {file_path}")

            # Copy hình ảnh vào clipboard
            os.startfile(file_path)  # Mở hình ảnh trong trình xem mặc định
            sleep(1)  # Chờ ứng dụng mở
            pyautogui.hotkey("ctrl", "c")  # Copy ảnh vào clipboard
            sleep(1)  # Chờ hoàn tất thao tác copy
            print("Đã copy hình ảnh vào clipboard.")

            alt_F4()  # TẮT HÌNH

            try:

                pyautogui.moveTo(
                    position_message_box, duration=0.5
                )  # Thêm duration để di chuyển chuột mượt hơn
                pyautogui.click()
                # Dán hình ảnh từ clipboard
                ctrl_V()
                sleep(2)  # Chờ hình ảnh được dán
                print("Hình ảnh đã được dán vào message_box.")

                # Tìm ô message_box trong Zalo
                message_box = WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located(
                        (By.XPATH, XPATHS_ZALO["message_box"])
                    )
                )
                message_box.click()
                print("Đã tìm thấy ô message_box.")
                message_box.send_keys(message)
                message_box.send_keys(" @")
                message_box.send_keys(tag_name)
                sleep(2)
                message_box.send_keys(Keys.ARROW_DOWN)
                message_box.send_keys(Keys.ENTER)
                sleep(2)
                print("GÕ TIN NHẮN THÀNH CÔNG!!!")

                # Nhấn nút gửi tin nhắn
                send_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, XPATHS_ZALO["send_button"]))
                )
                send_button.click()
                print("Tin nhắn gửi thành công!!!")
                sleep(5)

            except Exception as e:
                print(e)

        except Exception as e:
            print(e)

    def send_attached_img(self, file_path):
        try:
            # Kiểm tra đường dẫn ảnh
            if not os.path.isfile(file_path):
                raise FileNotFoundError(f"File không tồn tại: {file_path}")

            # Copy hình ảnh vào clipboard
            os.startfile(file_path)  # Mở hình ảnh trong trình xem mặc định
            sleep(1)  # Chờ ứng dụng mở
            pyautogui.hotkey("ctrl", "c")  # Copy ảnh vào clipboard
            sleep(1)  # Chờ hoàn tất thao tác copy
            print("Đã copy hình ảnh vào clipboard.")

            alt_F4()  # TẮT HÌNH

            try:

                pyautogui.moveTo(
                    position_message_box, duration=0.5
                )  # Thêm duration để di chuyển chuột mượt hơn
                pyautogui.click()
                # Dán hình ảnh từ clipboard
                ctrl_V()
                sleep(2)  # Chờ hình ảnh được dán
                print("Hình ảnh đã được dán vào message_box.")

            except Exception as e:
                print(e)

        except Exception as e:
            print(e)

    def send_message_CDBR(self, message):
        try:
            # tìm ô tin nhắn
            message_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["message_box"]))
            )
            message_box.click()
            message_box.send_keys(message)
            message_box.send_keys(Keys.CONTROL, "v")
            sleep(2)
            message_box.send_keys(Keys.ENTER)
            sleep(5)
        except Exception as e:
            print(e)


class OutLookBot(BrowserManager):

    def access_outlook(self):
        self.open_url(OUT_LOOK_URL)
        try:
            WebDriverWait(self.driver, 200).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        XPATHS_OUTLOOK["icon_outLook"],
                    )
                )
            )
            print("OutLook loaded successfully!")
        except Exception as e:
            print("Error loading OutLook:", e)

    # nhập tên người gửi
    def find_send_to_box(self):
        send_to_box = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_OUTLOOK["send_to_box"],
                )
            )
        )
        return send_to_box

    def to_user(self, user_name):
        send_to_box = self.find_send_to_box()
        send_to_box.click()
        sleep(1)
        try:
            send_to_box.send_keys(user_name)
            sleep(3)
            send_to_box.send_keys(Keys.TAB)
        except Exception as e:
            print(e)

    # nhập tên người cc
    def find_send_cc_box(self):
        send_cc_box = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_OUTLOOK["send_cc_box"],
                )
            )
        )
        return send_cc_box

    def cc_user(self, user_name):
        send_cc_box = self.find_send_cc_box()
        send_cc_box.click()
        sleep(1)
        try:
            send_cc_box.send_keys(user_name)
            sleep(3)
            send_cc_box.send_keys(Keys.TAB)
        except Exception as e:
            print(e)

    # tìm nút tạo thư mới
    def find_new_mail_button(self):
        new_mail_button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, XPATHS_OUTLOOK["new_mail_button"])
            )
        )
        return new_mail_button

    def find_popup_mail_button(self):
        popup_mail_button = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_OUTLOOK["popoup_mail"],
                )
            )
        )
        return popup_mail_button

    def find_send_TO_box_popup(self):
        element = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_OUTLOOK_POPUP["send_to_box"],
                )
            )
        )
        return element

    def send_TO_user_popup(self, user_name):
        send_TO_box = self.find_send_TO_box_popup()
        send_TO_box.click()
        sleep(1)
        try:
            send_TO_box.send_keys(user_name)
            sleep(3)
            send_TO_box.send_keys(Keys.TAB)
        except Exception as e:
            print(e)

    def find_send_CC_box_popup(self):
        element = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_OUTLOOK_POPUP["send_cc_box"],
                )
            )
        )
        return element

    def send_CC_user_popup(self, user_name):
        send_CC_box = self.find_send_CC_box_popup()
        send_CC_box.click()
        sleep(1)
        try:
            send_CC_box.send_keys(user_name)
            sleep(3)
            send_CC_box.send_keys(" ")
            send_CC_box.send_keys(Keys.TAB)
        except Exception as e:
            print(e)

    def input_subject_popup(self, subject):
        element = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_OUTLOOK_POPUP["subject_box"],
                )
            )
        )
        element.click()
        element.send_keys(subject)

    def find_context_box_popup(self):
        element = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_OUTLOOK_POPUP["context_box"],
                )
            )
        )
        return element

    # nhập tiêu đề thư
    def find_subject_box(self):
        subject_box = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, XPATHS_OUTLOOK["add_subject_box"])
            )
        )
        return subject_box

    def input_subject_mail(self, subject):
        subject_box = self.find_subject_box()
        subject_box.click()
        subject_box.send_keys(subject)

    # nhập nội dung văn bản
    def find_context_box(self):
        context_box = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, XPATHS_OUTLOOK["add_context_box"])
            )
        )
        return context_box

    def input_context_mail(self, context):
        context_box = self.find_context_box()
        context_box.click()
        context_box.send_keys(context)

    # tìm nút gửi
    def click_send_mail_button(self):
        send_mail_button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, XPATHS_OUTLOOK["send_mail_button"])
            )
        )
        send_mail_button.click()
        return send_mail_button

    def send_attach_file(self, file_path):
        try:
            insert_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_OUTLOOK["insert_button"])
                )
            )
            insert_button.click()
            sleep(1)

            attach_file_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_OUTLOOK["attach_file_button"])
                )
            )
            attach_file_button.click()
            sleep(1)

            browse_this_computer = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_OUTLOOK["browse_this_computer"])
                )
            )
            browse_this_computer.click()
            sleep(1)

            # Tương tác với cửa sổ file picker bằng pywinauto
            app = Application(backend="win32").connect(title_re="Open", found_index=0)
            dialog = app.window(title_re="Open")
            dialog.set_focus()  # Kích hoạt cửa sổ
            dialog["Edit"].type_keys(file_path)
            dialog["Open"].click()
            keyboard = Controller()
            keyboard.press(Key.enter)
            keyboard.release(Key.enter)

            sleep(3)
        except Exception as e:
            print(f"Lỗi trong quá trình đính kèm file: {e}")
