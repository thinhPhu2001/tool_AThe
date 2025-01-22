import win32com.client as win32
from config import *

import sys
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd

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


def send_email_with_image():
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
    df = pd.read_excel(EXCEL_TOOL_PATH_copy, "TH", header=3)

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
    print(To_list)
    # Giới hạn dữ liệu từ dòng 6 đến dòng 23
    df_limited = df.iloc[1:19]

     # Trích xuất danh sách email từ cột "GD" và "CD". Loại bỏ NaN, chuyển sang chuỗi và nối bằng ";"
    cc_gd_list = df_limited["gmail"].dropna().astype(str).tolist()
    cc_cd_list = df_limited["CD"].dropna().astype(str).tolist()

    # Kết hợp các email từ cột "GD" và "CD"
    CC_list = "; ".join(cc_gd_list + cc_cd_list)
    print(CC_list)

    # # Cấu hình thông tin cơ bản
    # mail.To = To_list  # Danh sách email từ cột "GD"
    # mail.CC = CC_list  # Danh sách email từ cột "CD"
    mail.To = "thinhlanhuong2001@gmail.com"  # Danh sách email từ cột "GD"
    # mail.CC = CC_list  # Danh sách email từ cột "CD"

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
        <p>KV3 báo cáo tồn WO PAKH, WO_WFM TKTU đến ngày {date_time}</p>
        <p>Các tỉnh tồn, quá hạn nhiều {tinh_list_str} đề nghị các PGĐ KT {gd_name_list_str} nắm thông tin hỗ trợ, điều hành tránh tăng tiền phạt. <br>
        Truyền thông đến FT nắm rõ hướng dẫn xử lý PAKH theo CT36 (tóm tắt theo file word đính kèm)</p>
        <img src="cid:image1">
        <p>Chi tiết huyện tồn:</p>
        <img src="cid:image2">
    </body>
    </html>
    """
    print(body_html)
    mail.HTMLBody = body_html  # Thiết lập nội dung HTML
    
    # Đính kèm hình ảnh
    attachment1_path = r"D:\A_The\data\doc\img\mail\tong.jpg" #IMG_TONG_PATH  # Thay bằng đường dẫn ảnh
    attachment1 = mail.Attachments.Add(attachment1_path)
    attachment1.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image1"
    )  # Gắn định danh "image1" cho ảnh

    attachment2_path = r"D:\A_The\data\doc\img\mail\tien_phat.jpg" #IMG_TIEN_PHAT_PATH  # Thay bằng đường dẫn ảnh
    attachment2 = mail.Attachments.Add(attachment2_path)
    attachment2.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image2"
    )  # Gắn định danh "image2" cho ảnh

    # đính kèm file excel
    mail.Attachments.Add(r"D:\A_The\data\doc\VCC3-tool cap nhat WO PAKH - Copy.xlsb")
    try:
        # Gửi email
        mail.Send()
        print("Email đã được gửi thành công!")
        return True

    except Exception as e:
        print(f"xảy ra lỗi khi gửi mail: {e}")
        return False

def send_simple_email():
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = "example@example.com"  # Thay bằng email người nhận
        mail.Subject = "Test Email"
        mail.Body = "This is a test email."
        mail.Send()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error: {e}")

send_simple_email()