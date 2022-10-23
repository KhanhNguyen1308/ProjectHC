import csv
import pandas as pd
from tkinter import *

data = pd.read_excel("D:/file Tên Chung/Duy Khánh/Project_KhoBaiTongHop/File/THEO DÕI HÀNG HOÁ NHẬP- XUẤT-TỒN 10-2022.xlsx")
file=pd.ExcelFile("D:/file Tên Chung/Duy Khánh/Project_KhoBaiTongHop/File/THEO DÕI HÀNG HOÁ NHẬP- XUẤT-TỒN 10-2022.xlsx")
print(file.sheet_names)
print(file.parse(0))

def ReadWareHouseData(filename, wh_Name):
    data=pd.read_excel(filename, sheet_name=wh_Name)
    data=data.to_csv(index = None,header=True)
    print(data)

filename="D:/file Tên Chung/Duy Khánh/Project_KhoBaiTongHop/File/THEO DÕI HÀNG HOÁ NHẬP- XUẤT-TỒN 10-2022.xlsx"
ReadWareHouseData(filename, "ĐIỀU KHO 1.1")