from config import *

### HÀM XỬ LÝ CHUNG
import subprocess
import os
import platform
from pynput.keyboard import Controller, Key
import pyautogui
from time import sleep
import pygetwindow as gw

# xóa dấu tiếng việt
s1 = "ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝàáâãèéêìíòóôõùúýĂăĐđĨĩŨũƠơƯưẠạẢảẤấẦầẨẩẪẫẬậẮắẰằẲẳẴẵẶặẸẹẺẻẼẽẾếỀềỂểỄễỆệỈỉỊịỌọỎỏỐốỒồỔổỖỗỘộỚớỜờỞởỠỡỢợỤụỦủỨứỪừỬửỮữỰựỲỳỴỵỶỷỸỹ"
s0 = "AAAAEEEIIOOOOUUYaaaaeeeiioooouuyAaDdIiUuOoUuAaAaAaAaAaAaAaAaAaAaAaAaEeEeEeEeEeEeEeEeIiIiOoOoOoOoOoOoOoOoOoOoOoOoUuUuUuUuUuUuUuYyYyYyYy"

position_center = [1064, 507]
position_context_box = [1245, 791]


def remove_accents(input_str):
    s = ""
    input_str.encode("utf-8")
    for c in input_str:
        if c in s1:
            s += s0[s1.index(c)]
        else:
            s += c
    return s


# mở file hình và đóng
def open_image(file_path):
    try:
        # Kiểm tra hệ điều hành và sử dụng lệnh phù hợp
        if platform.system() == "Windows":
            os.startfile(file_path)  # Windows
        elif platform.system() == "Darwin":
            subprocess.run(["open", file_path])  # macOS
        else:
            subprocess.run(["xdg-open", file_path])  # Linux
    except Exception as e:
        print(f"Không thể mở file: {e}")


# def close_image():
#     try:
#         # Kiểm tra hệ điều hành và sử dụng lệnh phù hợp
#         if platform.system() == "Windows":
#             # Đóng ứng dụng xem hình mặc định, ví dụ: "Microsoft.Photos.exe"
#             subprocess.run(
#                 ["taskkill", "/IM", "Photos.exe", "/F"], check=True
#             )
#         elif platform.system() == "Darwin":
#             # Đóng ứng dụng xem hình trên macOS (Preview là ứng dụng mặc định)
#             subprocess.run(["pkill", "Preview"], check=True)
#         else:
#             # Đóng ứng dụng trên Linux (tùy vào ứng dụng mặc định, ví dụ: eog)
#             subprocess.run(["pkill", "eog"], check=True)  # eog: Eye of GNOME
#     except Exception as e:
#         print(f"Không thể tắt ứng dụng: {e}")


# phím copy
def ctrl_C():
    # Sao chép (Ctrl + C)
    keyboard = Controller()
    keyboard.press(Key.ctrl)
    keyboard.press("c")
    keyboard.release("c")
    keyboard.release(Key.ctrl)
    sleep(2)


def ctrl_V():
    # Sao chép (Ctrl + C)
    keyboard = Controller()
    keyboard.press(Key.ctrl)
    keyboard.press("v")
    keyboard.release("v")
    keyboard.release(Key.ctrl)
    sleep(2)


def alt_F4():
    # Đóng cửa sổ hình ảnh (Alt + F4)
    keyboard = Controller()
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.f4)
    keyboard.release(Key.alt)


def alt_Tab():
    # Đóng cửa sổ hình ảnh (Alt + F4)
    keyboard = Controller()
    keyboard.press(Key.alt)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    keyboard.release(Key.alt)


def key_end():
    keyboard = Controller()
    keyboard.press(Key.end)
    keyboard.release(Key.end)


def open_and_copy_img(file_path):
    open_image(file_path)

    # Di chuyển chuột và click
    sleep(2)
    pyautogui.moveTo(
        position_center, duration=0.5
    )  # Thêm duration để di chuyển chuột mượt hơn
    pyautogui.click()
    ctrl_C()
    # close_image()
    alt_F4()
    sleep(1)
    pyautogui.moveTo(position_context_box, duration=0.5)
    pyautogui.scroll(-100)
    pyautogui.click()
    ctrl_V()
