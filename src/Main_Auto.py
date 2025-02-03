from config import *
from database import *
from utils import *
from openVPN import *
from excel_handler import *
from browser import *

from Vo_tuyen import *


from pynput.keyboard import Controller, Key
import schedule
import sys
from selenium.webdriver.common.action_chains import ActionChains

# thay đôi môi trường tiếng Việt
sys.stdout.reconfigure(encoding="utf-8")

file_path = r"D:\2-Job\Viettel\project_thu_viec\Auto_tool_offical\data\excel\data_didong\gnoc.xlsx"


def auto_WSA():
    try:
        if not getDB_to_excel(DATA_GNOC_RAW_PATH):
            browser.start_browser(CHROME_PROFILE_CDBR_PATH)
            whatsapp.driver = browser.driver
            whatsapp.send_Error_Notification(
                PHONE_NUMBER, "Lỗi xảy ra khi lấy dữ liệu từ SQL"
            )
            return False

        try:
            tool_status = excel_transition_and_run_macro(
                data_gnoc_manager, EXCEL_tool_manager
            )

            if not tool_status:
                print(
                    "Xảy ra lỗi trong quá trình chạy dữ liệu, sẽ không gửi tin nhắn cho các tỉnh. Chờ thời gian đến lần chạy tiếp theo"
                )
                browser.start_browser(CHROME_PROFILE_CDBR_PATH)
                whatsapp.driver = browser.driver
                whatsapp.send_Error_Notification(
                    PHONE_NUMBER, "Lỗi xảy ra khi xử lý dữ liệu"
                )
                browser.close()
                return False
            else:
                try:
                    send_mes_status = send_message()
                    if send_mes_status:
                        message = (
                            f"Đã gửi thông tin cho các tỉnh vào lúc: {datetime.now()}"
                        )
                    else:
                        message = "Xảy ra lỗi khi gửi tin nhắn"

                    browser.start_browser(CHROME_PROFILE_CDBR_PATH)
                    whatsapp.driver = browser.driver
                    whatsapp.send_Error_Notification(PHONE_NUMBER, message)
                    browser.close()
                    return True

                except Exception as e:
                    print(e)

        except Exception as e:
            print(e)

    except Exception as e:
        print(e)


def auto_WSA_nofi():
    if not auto_WSA():
        return
    print("========================")
    print("Chờ đến TÁC VỤ tiếp theo")
    print("========================")


def auto_WSA_mail():
    auto_WSA()

    try:
        send_email_status = send_email()
        if send_email_status:
            print("Gửi mail thành công!")
            message = f"Đã gửi mail cho các tỉnh vào lúc: {datetime.now()}"
        else:
            print("Gửi mail thất bại!")
            message = "Gửi mail thất bại!"

        browser.start_browser(CHROME_PROFILE_CDBR_PATH)
        whatsapp.driver = browser.driver
        whatsapp.send_Error_Notification(PHONE_NUMBER, message)
        browser.close()
        print("========================")
        print("Chờ đến TÁC VỤ tiếp theo")
        print("========================")

    except Exception as e:
        print(f"Lỗi Khi gửi mail {e}")


if __name__ == "__main__":

    # schedule.every().day.at("08:00").do(auto_WSA_mail)
    # schedule.every().day.at("13:30").do(auto_WSA_nofi)

    # print("========================")
    # print("Chờ đến TÁC VỤ tiếp theo")
    # print("========================")

    # while True:
    #     schedule.run_pending()
    #     sleep(1)

    # browser.start_browser(CHROME_PROFILE_CDBR_PATH)
    # sleep(5000)

    getDB_to_excel(DATA_GNOC_RAW_PATH)
