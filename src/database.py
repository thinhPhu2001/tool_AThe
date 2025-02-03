import pandas as pd
import pymysql
from sqlalchemy import create_engine
import traceback
import subprocess
import pyautogui
import pyotp
import time
from pynput.keyboard import Controller, Key
from time import sleep

from config import *


import sys

# Đảm bảo in tiếng Việt không bị lỗi
sys.stdout.reconfigure(encoding="utf-8")


def connect_to_db():
    """Hàm kết nối đến cơ sở dữ liệu MySQL."""
    pymysql.install_as_MySQLdb()
    try:
        engine = create_engine(MySQL_DB)
        connection = engine.connect()
        return connection
    except Exception as e:
        print("Lỗi kết nối Database:")
        traceback.print_exc()
        return None


def query_to_excel(connection, query, output_path):
    """Hàm thực hiện truy vấn SQL và xuất ra file Excel."""
    try:
        # Thực hiện truy vấn
        result = pd.read_sql(query, connection)

        # Xuất kết quả ra file Excel
        result.to_excel(output_path, index=False, engine="openpyxl")
        print(f"Dữ liệu đã được xuất ra file: {output_path}")
    except Exception as e:
        print("Lỗi khi thực hiện truy vấn hoặc xuất file:")
        traceback.print_exc()
