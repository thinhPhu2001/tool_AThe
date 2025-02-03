### CÀI ĐẶT CẤU HÌNH - VD: TÀI KHOẢN WHATSAPP, PHONE - NUMBER

from pathlib import Path

# Đường dẫn URL
WHATSAPP_URL = "https://web.whatsapp.com/"
ZALO_URL = "https://id.zalo.me/account?continue=https%3A%2F%2Fchat.zalo.me%2F"
OUT_LOOK_URL = "https://outlook.office.com/mail/"

# Lưu profile chrome

# CHROME_PROFILE_CDBR_PATH = r"C:\Users\Admin\AppData\Local\Google\Chrome\User Data\Default"
# CHROME_PROFILE_DI_DONG_PATH = (
#     r"C:\Users\Admin\AppData\Local\Google\Chrome\User Data\profile_di_dong"
# )

# Lấy đường dẫn thư mục gốc của project
project_dir = Path(__file__).resolve().parent.parent  # Thư mục cha của 'src'

# Đường dẫn thư mục
OPENVN_PATH = project_dir / "data" / "openvn"
EXCEL_PATH = project_dir / "data" / "doc"
IMAGE_PATH = project_dir / "data" / "img"


IMG_MAIL_PATH = EXCEL_PATH / "img" / "mail"

IMG_TONG_PATH = IMG_MAIL_PATH / "tong.jpg"
IMG_TIEN_PHAT_PATH = IMG_MAIL_PATH / "tien_phat.jpg"

IMG_TINH_PATH = EXCEL_PATH / "img" / "tinh"


DATA_CONFIG_PATH = EXCEL_PATH / "gnoc" / "config.xlsx"
DATA_GNOC_RAW_PATH = EXCEL_PATH / "gnoc" / "gnoc.xlsx"
EXCEL_TOOL_PATH = EXCEL_PATH / "VCC3-tool cap nhat WO PAKH.xlsb"
EXCEL_TOOL_PATH_copy = EXCEL_PATH / "VCC3-tool cap nhat WO PAKH - Copy.xlsb"
EXCEL_GuiTinh_PATH = EXCEL_PATH / "Data_gui_tinh.xlsx"

# Đọc file cấu hình OpenVPN
OPEN_VPN_CONFIG_PATH = project_dir / "config.txt"
with open(OPEN_VPN_CONFIG_PATH, "r") as file:
    # Lặp qua các dòng của file ngay khi mở
    for line in file:
        if "phone:" in line:
            PHONE_NUMBER = line.strip().replace("phone:", "").strip()

        if "pwd:" in line:
            OTP_SECRET = line.strip().replace("pwd:", "").strip()

        if "SQL:" in line:
            MySQL_DB = line.strip().replace("SQL:", "").strip()

        if "opvn_profile:" in line:
            OPEN_VPN_CONFIG_PATH = line.strip().replace("opvn_profile:", "").strip()

        if "opvn_path:" in line:
            OPEN_VPN_PATH = line.strip().replace("opvn_path:", "").strip()

        if "cdbr:" in line:
            CHROME_PROFILE_CDBR_PATH = line.strip().replace("cdbr:", "").strip()
