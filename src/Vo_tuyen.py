from config import *
from database import *
from utils import *
from openVPN import *
from excel_handler import *
from browser import *

import shutil
from sqlalchemy import text
from pynput.keyboard import Controller, Key
import schedule
import sys
import win32com.client as win32

# thay đôi môi trường tiếng Việt
sys.stdout.reconfigure(encoding="utf-8")

# excel variable
# |-- DI ĐỘNG

# browser variable
browser = BrowserManager()
whatsapp = WhatsAppBot()
outlook = OutLookBot()
zalo = ZaloBot()

date_obj = datetime.now()
day = date_obj.day  # Lấy ngày
month = date_obj.month
year = date_obj.year
subject_mail = (
    f"FW: KV3 báo cáo tồn WO PAKH, WO_WFM TKTU đến ngày: {day}/{month}/{year}"
)

data_gnoc_manager = ExcelManager(DATA_GNOC_RAW_PATH)
EXCEL_tool_manager = ExcelManager(EXCEL_TOOL_PATH)


def getDB_to_excel(excel_gnoc_path):
    max_retries = 5
    retries = 0

    while retries < max_retries:
        try:
            # Kiểm tra file cấu hình Excel
            if not os.path.exists(DATA_CONFIG_PATH):
                raise FileNotFoundError(f"File {DATA_CONFIG_PATH} không tồn tại.")
            print("File cấu hình Excel tồn tại.")

            # Kết nối OpenVPN
            if not on_openvpn():
                raise ConnectionError("Kết nối OpenVPN thất bại.")
            print("Kết nối OpenVPN thành công.")
            sleep(5)

            # Kết nối cơ sở dữ liệu
            connection = connect_to_db()
            if connection is None:
                raise ConnectionError("Kết nối Database thất bại.")
            print("Kết nối Database thành công.")

            # Đọc file Excel, lấy dữ liệu từ sheet "Sheet1"
            df = pd.read_excel(DATA_CONFIG_PATH, sheet_name="Sheet1", header=0)

            # Xử lý dữ liệu nhóm
            groups = {
                col: df[col].dropna().astype(str).tolist()
                for col in ["Nhóm điều phối", "Loại công việc", "Đơn vị tạo"]
            }
            groups_sql = {
                col: ", ".join(f"'{group}'" for group in groups[col]) for col in groups
            }

            # Tạo câu truy vấn SQL
            query_pakh = f"""
                SELECT *
                FROM gnoc.gnoc_open_90d
                WHERE `Nhóm điều phối` IN ({groups_sql["Nhóm điều phối"]})
                AND `Loại công việc` IN ({groups_sql["Loại công việc"]})
                AND `Đơn vị tạo` IN ({groups_sql["Đơn vị tạo"]});
            """
            print("Câu truy vấn SQL đã được tạo thành công.")

            # Xuất dữ liệu ra Excel
            query_to_excel(connection, query_pakh, excel_gnoc_path)
            print("Lấy file database thành công!")
            return

        except FileNotFoundError as fe:
            print(f"Lỗi file: {fe}")
        except ConnectionError as ce:
            print(f"Lỗi kết nối: {ce}")
        except Exception as e:
            print(f"Lỗi không xác định: {e}")
        finally:
            # Đảm bảo đóng kết nối database
            if "connection" in locals() and connection:
                connection.close()
                print("Đóng kết nối Database.")

            # Đảm bảo tắt VPN
            off_openvpn()
            print("Tắt kết nối OpenVPN.")

        # Tăng số lần thử và thời gian chờ
        retries += 1
        print(f"Thử lại lần thứ {retries} sau 5 giây...")
        sleep(5)

    print("Không thể hoàn thành tác vụ sau nhiều lần thử.")


def delete_data_folder(folder_path):
    # Kiểm tra nếu thư mục tồn tại
    if os.path.exists(folder_path):
        # Duyệt qua tất cả các tệp và thư mục trong thư mục
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)

            # Nếu là thư mục, xóa thư mục cùng tất cả các tệp trong đó
            if os.path.isdir(file_path):
                shutil.rmtree(file_path)
            else:
                os.remove(file_path)
        print("Đã xóa hết dữ liệu trong thư mục.")
    else:
        print("Thư mục không tồn tại.")


def excel_transition_and_run_macro(
    excel_gnoc_manager: ExcelManager, excel_tool_manager: ExcelManager
):
    """
    Chuyển dữ liệu từ file excel gnoc raw qua qua file tool để xử lý
    """
    # xóa data cũ trước khi chạy file xử lý
    delete_data_folder(IMG_TINH_PATH)
    delete_data_folder(IMG_MAIL_PATH)

    # Try opening the tool file
    if not excel_tool_manager.open_file():
        print("Không mở được file tool.")
        # excel_gnoc_manager.close_all_file()
        return False

    # Try clearing old data in the tool file
    try:
        excel_tool_manager.clear_data(
            "data_Gnoc", start_row=3, start_col="AD", end_col="BC"
        )
    except Exception as e:
        print(f"Lỗi khi xóa dữ liệu trong file tool: {e}")
        excel_tool_manager.save_file()
        excel_gnoc_manager.close_all_file()
        return False

    # Try pasting new data into the tool file
    try:
        if not excel_tool_manager.run_macro("Module1.DanDuLieu"):
            return False

        print("Chuyển dữ liệu thành công!!!")

        try:
            if not excel_tool_manager.run_macro("Module1.pic_cum_huyen_loop"):
                return False
            
            if not excel_tool_manager.run_macro("Module1.pic_TH_tinh"):
                return False
            
            if not excel_tool_manager.run_macro("Module1.pic_tienPhat"):
                return False
            
            excel_tool_manager.run_macro("Module2.Xuatdataguitinh")
               
            excel_tool_manager.save_file()
            excel_tool_manager.close_all_file()
            return True

        except Exception as e:
            print(f"Lỗi khi chạy macro trong file tool: {e}")
            excel_tool_manager.save_file()
            excel_gnoc_manager.close_all_file()
            return False

    except Exception as e:
        print(f"Lỗi khi dán dữ liệu vào file tool: {e}")
        excel_tool_manager.save_file()
        excel_gnoc_manager.close_all_file()
        return False


def send_message():
    browser.start_browser(CHROME_PROFILE_CDBR_PATH)
    whatsapp.driver = browser.driver

    try:
        df = pd.read_excel(DATA_CONFIG_PATH, sheet_name="Sheet2", header=0)

        for index, row in df.iterrows():
            link_group = row["Link group"]
            message = row["Message"]
            img_name = row["img name"]

            # Kiểm tra dữ liệu
            if not link_group or not message or not img_name:
                print(f"Dữ liệu không hợp lệ tại dòng {index + 1}. Bỏ qua.")
                browser.close()
                return False

            img_path = f"{IMG_TINH_PATH}/{img_name}.jpg"

            if not os.path.exists(img_path):
                print(f"Ảnh {img_path} không tồn tại. Đóng tiến trình gửi tin nhắn")
                browser.close()
                return False

            temp = whatsapp.find_group_name(link_group)
            retries = 0
            max_retries = 3
            while retries < max_retries:
                if temp:
                    try:
                        send_mess_status = whatsapp.send_attached_img_message(
                            message, img_path, tag_name=None
                        )
                        sleep(5)
                        if send_mess_status:
                            print(f"Gửi tin nhắn đén nhóm [{img_name}] thành công")
                            sleep(5)
                            break
                    except Exception as e:
                        print(f"Lỗi khi gửi tin nhắn đến nhóm [{img_name}]: {e}")

                else:
                    retries += 1
                    print(
                        f"Không tìm thấy nhóm {img_name}. Thử lại lần {retries}/{max_retries}..."
                    )
                    whatsapp.access_whatsapp()  # Hàm tải lại trang (giả định bạn có hàm này)
                    temp = whatsapp.find_group_name(link_group)  # Thử tìm lại nhóm

            if retries >= max_retries:
                print(
                    f"Không thể gửi tin nhắn đến nhóm {img_name} sau {max_retries} lần thử."
                )

        print("Đã gửi tin nhắn cho các tỉnh. Kết thúc tiến trình")
        browser.close()
        return True

    except Exception as e:
        print(f"lỗi xảy ra khi gửi tin nhắn cho các tỉnh: {e}")
        browser.close()
        return False


def send_email():
    if not os.path.exists(IMG_TONG_PATH):
        print(f"Ảnh {IMG_TONG_PATH} không tồn tại. Đóng tiến trình gửi tin nhắn")
        return False

    if not os.path.exists(IMG_TIEN_PHAT_PATH):
        print(f"Ảnh {IMG_TIEN_PHAT_PATH} không tồn tại. Đóng tiến trình gửi tin nhắn")
        return False

    # xác định ngày hiện tại
    date_time = datetime.now().strftime("%d/%m/%Y")

    # Khởi tạo ứng dụng Outlook
    outlook = win32.Dispatch("Outlook.Application")

    # Tạo email mới
    mail = outlook.CreateItem(0)  # 0 là mã loại cho email

    # Đọc file Excel
    df = pd.read_excel(EXCEL_TOOL_PATH, "TH", header=3)

    # Giới hạn dữ liệu và xử lý To_list
    df_limited2 = df.iloc[20:]
    gmail_list = (
        df_limited2.loc[df_limited2["tiền phạt"] != 0, "gmail"]
        .dropna()
        .astype(str)
        .tolist()
    )
    # Loại bỏ các giá trị không phải email hợp lệ
    gmail_list = [email for email in gmail_list if "@" in email and "." in email]
    To_list = "; ".join(gmail_list)

    # Giới hạn dữ liệu từ dòng 6 đến dòng 23
    df_limited = df.iloc[1:19]

    # Trích xuất danh sách email từ cột "GD" và "CD". Loại bỏ NaN, chuyển sang chuỗi và nối bằng ";"
    cc_gd_list = df_limited["gmail"].dropna().astype(str).tolist()
    cc_cd_list = df_limited["CD"].dropna().astype(str).tolist()

    # Kết hợp các email từ cột "GD" và "CD"
    CC_list = "; ".join(cc_gd_list + cc_cd_list)

    bonus = ["thepv1@viettel.com.vn", "nguyenbatung@viettel.com.vn", "quannt1@viettel.com.vn"]
    CC_list = "; ".join([CC_list] + bonus) if CC_list else "; ".join(bonus)

    # Cấu hình thông tin cơ bản
    mail.To = To_list  # Danh sách email từ cột "GD"
    mail.CC = CC_list  # Danh sách email từ cột "CD"
    mail.Subject = f"FW: KV3 báo cáo tồn WO PAKH, WO_WFM TKTU đến ngày {date_time}"

    # Giới hạn dữ liệu từ dòng 2 đến dòng 19 (1-based index)
    df_limited = df.iloc[1:19]

    # Danh sách để lưu các tỉnh có tiền phạt khác 0
    tinh_list = []
    gd_name_list = []

    # Duyệt qua từng dòng của DataFrame
    for index, row in df_limited.iterrows():
        tinh_name = row["MÃ Tỉnh"]
        tienPhat = row["tiền phạt"]
        gd_name = row["GD_name"]

        if tienPhat != 0:  # Nếu tiền phạt khác 0
            tinh_list.append(tinh_name)  # Thêm tỉnh vào danh sách
            gd_name_list.append(gd_name)

    # Nối các tỉnh thành chuỗi, cách nhau bởi dấu phẩy
    tinh_list_str = ", ".join(tinh_list)
    gd_name_list_str = ", ".join(gd_name_list)
    # Nội dung email (HTML format) với biến
    body_html = f"""
    <html>
    <body>
        <h1>K/g các đc CNCT Tỉnh!</h1>
        <h3>KV3 báo cáo tồn WO PAKH, WO_WFM TKTU đến ngày {date_time}</h3>
        <p>Các tỉnh tồn, quá hạn nhiều {tinh_list_str} đề nghị các PGĐ KT {gd_name_list_str} nắm thông tin hỗ trợ, điều hành tránh tăng tiền phạt. <br>
        Truyền thông đến FT nắm rõ hướng dẫn xử lý PAKH theo CT36 (tóm tắt theo file word đính kèm)</p>
        <img src="cid:image1">
        <h3>Chi tiết huyện tồn:</h3>
        <img src="cid:image2">
    </body>
    </html>
    """

    mail.HTMLBody = body_html  # Thiết lập nội dung HTML

    # Đính kèm hình ảnh
    attachment1_path = str(IMG_TONG_PATH)  # Thay bằng đường dẫn ảnh
    attachment1 = mail.Attachments.Add(attachment1_path)
    attachment1.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image1"
    )  # Gắn định danh "image1" cho ảnh

    attachment2_path = str(IMG_TIEN_PHAT_PATH)  # Thay bằng đường dẫn ảnh
    attachment2 = mail.Attachments.Add(attachment2_path)
    attachment2.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image2"
    )  # Gắn định danh "image2" cho ảnh

    # đính kèm file excel
    mail.Attachments.Add(str(EXCEL_GuiTinh_PATH))
    mail.Attachments.Add(r"D:\A_The\data\doc\HUONG DAN NHANH XU LY PHAN ANH KHACH HANG Vo TUYEN CHO CNKT.docx")

    try:
        # Gửi email
        mail.Send()
        print("Email đã được gửi thành công!")
        return True

    except Exception as e:
        print(f"xảy ra lỗi khi gửi mail: {e}")
        return False


# def send_mail_process():
#     browser.start_browser(CHROME_PROFILE_CDBR_PATH)
#     browser.driver.maximize_window()
#     outlook.driver = browser.driver
#     outlook.access_outlook()
#     try:
#         new_mail_button = outlook.find_new_mail_button()
#         new_mail_button.click()
#         popup_mail_button = outlook.find_popup_mail_button()
#         popup_mail_button.click()
#         # Lấy tất cả các handles (ID cửa sổ)
#         all_handles = outlook.driver.window_handles

#         # Chuyển sang cửa sổ con (pop-up)
#         for handle in all_handles:
#             if handle != outlook.driver.current_window_handle:
#                 print("Chuyển sang cửa sổ pop-up:", handle)
#                 outlook.driver.switch_to.window(handle)
#                 break

#         outlook.driver.maximize_window()

#         try:
#             # """
#             # xử lý gửi tên CC
#             # """
#             # df = pd.read_excel(EXCEL_TOOL_PATH, "TH", header=24)
#             # for index, row in df.iloc[1:].iterrows():
#             #     count = row["status"]
#             #     if count == 1:
#             #         user_gmail = row[
#             #             "*Các huyện tồn WO quá hạn đang phạt tiền. Đề nghị GĐH hỗ trợ FT Dong WO tránh tăng tiền phạt*"
#             #         ]
#             #         outlook.send_CC_user_popup(user_gmail)

#             # print("CC tên mail tên liên quan xong")

#             """
#             CC tất cả giám đốc
#             """
#             names_to_send = []
#             df = pd.read_excel(EXCEL_TOOL_PATH, "TH", header=3)
#             for index, row in df.iloc[1:19].iterrows():
#                 gd_name = row["GD"]
#                 cd_name = row["CD"]
#                 names_to_send.append(gd_name)
#                 names_to_send.append(cd_name)

#             outlook.send_CC_user_popup(names_to_send)

#             print("Đã CC cho tất cả cd và giám đốc huyện")

#             # # try:
#             # #     # nhập tiêu đề mail
#             # #     outlook.input_subject_popup(subject_mail)
#             # #     print("Nhập tiêu đề mail, thành công!!!")
#             # #     sleep(5)

#             # #     try:
#             # #         # nhập nội dung mail
#             # #         context_box = outlook.find_context_box_popup()
#             # #         edit_context_mail(context_box)

#             # #         paste_image(IMG_TONG_PATH)
#             # #         context_box.send_keys(Keys.ENTER)
#             # #         context_box.send_keys("Chi tiết huyện tồn:")
#             # #         context_box.send_keys(Keys.ENTER)
#             # #         paste_image(IMG_TIEN_PHAT_PATH)
#             # #         sleep(5000)

#             # #     except Exception as e:
#             # #         print(e)
#             # except Exception as e:
#             #     print(e)
#         except Exception as e:
#             print(e)
#     except Exception as e:
#         print(e)


# def paste_image(file_path):
#     context_box_position = [1170, 986]
#     try:
#         # Kiểm tra đường dẫn ảnh
#         if not os.path.isfile(file_path):
#             raise FileNotFoundError(f"File không tồn tại: {file_path}")

#         # Copy hình ảnh vào clipboard
#         os.startfile(file_path)  # Mở hình ảnh trong trình xem mặc định
#         sleep(1)  # Chờ ứng dụng mở
#         pyautogui.hotkey("ctrl", "c")  # Copy ảnh vào clipboard
#         sleep(1)  # Chờ hoàn tất thao tác copy
#         print("Đã copy hình ảnh vào clipboard.")
#         alt_F4()  # TẮT HÌNH

#         pyautogui.moveTo(context_box_position, duration=0.5)
#         pyautogui.click()

#         ctrl_V()
#         sleep(2)  # Chờ hình ảnh được dán
#         print("Hình ảnh đã được dán vào")

#         try:
#             pyautogui.moveTo(context_box_position, duration=0.5)
#             pyautogui.click()
#             sleep(2)
#             picture_format = WebDriverWait(outlook.driver, 10).until(
#                 EC.presence_of_element_located(
#                     (
#                         By.XPATH,
#                         '//*[@id="tablist"]/div/div[6]',
#                     )
#                 )
#             )
#             picture_format.click()

#             size_button_path = str(IMAGE_PATH / "size.png")
#             size_button = pyautogui.locateOnScreen(size_button_path, confidence=0.8)
#             pyautogui.click(pyautogui.center(size_button))
#             sleep(1)

#             small_button_path = str(IMAGE_PATH / "small.png")
#             small_button = pyautogui.locateOnScreen(small_button_path, confidence=0.8)
#             pyautogui.click(pyautogui.center(small_button))
#             key_end()

#         except Exception as e:
#             print(e)

#     except Exception as e:
#         print(e)


# def edit_context_mail(context_box):
#     context_box = outlook.find_context_box_popup()
#     context_box.click()
#     text_1 = f"KV3 báo cáo tồn WO PAKH, WO_WFM TKTU đến ngày : {day}/{month}/{year}"
#     context_box.send_keys(text_1)
#     context_box.send_keys(Keys.ENTER)

#     text_2 = "Các tỉnh tồn, quá hạn nhiều: "
#     text_3 = " đề nghi các PGĐ KT "
#     df = pd.read_excel(EXCEL_TOOL_PATH, "TH", header=3)

#     # Lặp qua các hàng
#     for index, (i, row) in enumerate(df.iloc[1:18].iterrows()):
#         ma_tinh = row["Mã tỉnh"]
#         gd_kt_name = row["TT huyện"]
#         count = row["Tổng"]
#         if count >= 30:
#             text_2 += f"{ma_tinh}, "
#             text_3 += f"{gd_kt_name}, "

#     text_5 = "Truyền thông đến FT nắm rõ hướng dẫn xử lý PAKH theo CT36 (tóm tắt theo file word đính kèm)!!!"
#     text_4 = f"{text_2} {text_3} nắm thông tin hỗ trợ, điều hành tránh tăng tiền phạt."
#     context_box.send_keys(text_4)
#     context_box.send_keys(Keys.ENTER)
#     context_box.send_keys(text_5)
#     context_box.send_keys(Keys.ENTER)
