import os
import csv
import tkinter
import pandas as pd
from tkinter import *
from tkinter import ttk, filedialog
from function import *
filename = ""
label_create = []
date_label = []
tkbill_label = []
kh_label = []

wh_Data=[]
page = 0
maxPage = 0
window = tkinter.Tk()
window.geometry("1920x1080")
all_btnWH = []
var = tkinter.StringVar()
var1 = tkinter.StringVar()

kho11Data=[];kho12Data=[];kho13Data=[];kho21Data=[];kho22Data=[];kho23Data=[];kho31Data=[];kho32Data=[]

def GetWareHouse_Data(WareHouseName):
    data = pd.read_excel(var.get(), sheet_name=WareHouseName)
    folder = filename.split("/")[:-1]
    folder_link=""
    for i in folder:
        folder_link=folder_link+i+"/"
    data = data.to_csv(folder_link+str(WareHouseName)+".csv")
    stt,date,stk,khachhang,tonkydau,nhap_nb,vaokho,xuat_nb,rakho,toncuoi,ghichu,toncuoiky=[],[],[],[],[],[],[],[],[],[],[],[]
    with open(folder_link+str(WareHouseName)+".csv", encoding="utf8") as e:
        data = csv.reader(e)
        count = 0
        for line in data:
            count+=1
            if count >= 6:
                stt.append(line[0]);date.append(line[1]);stk.append(line[2]);khachhang.append(line[3]);tonkydau.append(line[4])
                nhap_nb.append(line[5]);vaokho.append(line[6]);xuat_nb.append(line[7]);rakho.append(line[8])
                toncuoi.append(line[9]);ghichu.append(line[10]);toncuoiky.append(line[11])
    Data=[stt,date,stk,khachhang,tonkydau,nhap_nb,vaokho,xuat_nb,rakho,toncuoi,ghichu,toncuoiky]
    return Data


def CreateLabelManager(canvas):
    lbstt = tkinter.Label(canvas, text="STT", bg='lightgreen')
    lbstt.place(x=10,y=30)
    lbNT = tkinter.Label(canvas, text="NGÀY THÁNG", bg='lightgreen')
    lbNT.place(x=40,y=30)
    lbSTK = tkinter.Label(canvas, text="                         SỐ TỜ KHAI                         ", bg='lightgreen')
    lbSTK.place(x=130,y=30)
    lbKH = tkinter.Label(canvas, text="KHÁCH HÀNG", bg='lightgreen')
    lbKH.place(x=360,y=30)
    lbDVT = tkinter.Label(canvas, text="ĐVT", bg='lightgreen')
    lbDVT.place(x=450,y=30)
    lbTKD = tkinter.Label(canvas, text="   TỒN KỲ ĐẦU   ", bg='lightgreen')
    lbTKD.place(x=485,y=30)


    lbNNB_ = tkinter.Label(canvas, text="     NHẬP NỘI BỘ     ", bg='#00ff00')
    lbNNB_.place(x=585,y=10)
    lbNNB = tkinter.Label(canvas, text="NỘI BỘ", bg='lightgreen')
    lbNNB.place(x=585,y=30)
    lbVK = tkinter.Label(canvas, text="VÀO KHO", bg='lightgreen')
    lbVK.place(x=640,y=30)

    lbXNB_ = tkinter.Label(canvas, text="     XUẤT NỘI BỘ     ", bg='#00ff00')
    lbXNB_.place(x=705,y=10)
    lbXNB = tkinter.Label(canvas, text="NỘI BỘ", bg='lightgreen')
    lbXNB.place(x=705,y=30)
    lbRK = tkinter.Label(canvas, text=" RA KHO ", bg='lightgreen')
    lbRK.place(x=760,y=30)


    lbTC = tkinter.Label(canvas, text="  TỒN CUỐI  ", bg='lightgreen')
    lbTC.place(x=825,y=30)
    lbGC = tkinter.Label(canvas, text="GHI CHÚ", bg='lightgreen')
    lbGC.place(x=905,y=30)
    lbTCK = tkinter.Label(canvas, text="   TỒN CUỐI KỲ   ", bg='#daa520')
    lbTCK.place(x=965,y=30)

def UploadAction(event=None):
    global kho11Data,kho12Data,kho13Data,kho21Data,kho22Data,kho23Data,kho31Data,kho32Data
    filename = filedialog.askopenfilename()
    var.set(filename)
    CheckImportStatus()
    kho11Data=GetWareHouse_Data("ĐIỀU KHO 1.1")
    kho12Data=GetWareHouse_Data("ĐIỀU KHO 1.2")
    kho13Data=GetWareHouse_Data("ĐIỀU KHO 1.3")
    kho21Data=GetWareHouse_Data("ĐIỀU KHO 2.1")
    kho22Data=GetWareHouse_Data("ĐIỀU KHO 2.2")
    kho23Data=GetWareHouse_Data("ĐIỀU KHO 2.3")
    kho31Data=GetWareHouse_Data("ĐIỀU KHO 3.1")
    kho32Data=GetWareHouse_Data("ĐIỀU KHO 3.2")

def CreateWareHouse(status):
    global all_btnWH
    btn11 = tkinter.Button(window, text="Kho 1.1", command= lambda: ShowWareHouseData(btn=btn11, file=var.get()), state=status)
    btn11.place(x=10,y=50)
    btn12 = tkinter.Button(window, text="Kho 1.2", command= lambda: ShowWareHouseData(btn=btn12, file=var.get()), state=status)
    btn12.place(x=60,y=50)
    btn13 = tkinter.Button(window, text="Kho 1.3", command= lambda: ShowWareHouseData(btn=btn13, file=var.get()), state=status)
    btn13.place(x=110,y=50)
    btn21 = tkinter.Button(window, text="Kho 2.1", command= lambda: ShowWareHouseData(btn=btn21, file=var.get()), state=status)
    btn21.place(x=160,y=50)
    btn22 = tkinter.Button(window, text="Kho 2.2", command= lambda: ShowWareHouseData(btn=btn22, file=var.get()), state=status)
    btn22.place(x=210,y=50)
    btn23 = tkinter.Button(window, text="Kho 2.3", command= lambda: ShowWareHouseData(btn=btn23, file=var.get()), state=status)
    btn23.place(x=260,y=50)
    btn31 = tkinter.Button(window, text="Kho 3.1", command= lambda: ShowWareHouseData(btn=btn31, file=var.get()), state=status)
    btn31.place(x=310,y=50)
    btn32 = tkinter.Button(window, text="Kho 3.2", command= lambda: ShowWareHouseData(btn=btn32, file=var.get()), state=status)
    btn32.place(x=360,y=50)
    all_btnWH = [btn11, btn12, btn13, btn21, btn22, btn23, btn31, btn32]

def CheckImportStatus():
    if var.get() == "":
        CreateWareHouse(tkinter.DISABLED)
    else:
        CreateWareHouse(tkinter.ACTIVE)

def ShowWareHouseData(btn, file, event=None):
    global page, maxPage, all_btnWH
    for bt in all_btnWH:
        if bt == btn:
            bt['bg'] = 'red'
        else:bt['bg'] = 'lightgray'

    # var1.set(btn['text'])
    wh_Name = "ĐIỀU KHO "+ str(btn['text'][-3:])
    data_ex=pd.read_excel(file, sheet_name=wh_Name)
    data = data_ex.to_csv("D:/file Tên Chung/Duy Khánh/Project_KhoBaiTongHop/File/"+wh_Name+".csv",index = None,header=True)
    link="D:/file Tên Chung/Duy Khánh/Project_KhoBaiTongHop/File/" + str(wh_Name) + ".csv"
    if len(label_create) > 0:
        for lb in label_create:
            lb.destroy()
    if len(date_label) > 0:
        for lb in date_label:
            lb.destroy()
    if len(tkbill_label) > 0:
        for lb in tkbill_label:
            lb.destroy()
    with open(link, encoding="utf8") as e:
        data = csv.reader(e)
        count=0
        wh_Data = []
        for line in data:
            count+=1
            if count >= 6:
                wh_Data.append(line)
        
    if len(wh_Data) > 0:
        maxPage=round(len(wh_Data)/17)+1
        for i in range(page*17,(page+1)*17):
            if(i == len(wh_Data)-1):
                break
            lb=tkinter.Label( canvas1, text=wh_Data[i][0])
            lb.place(x=10,y=(50+20*i))
            date = wh_Data[i][1].split(".")
            date_text= DateEdit(date)
            dtlb=tkinter.Label(canvas1, text=date_text)
            dtlb.place(x=50,y=(50+20*i))
            
            if wh_Data[i][2] != '':
                
                tkbill_text=wh_Data[i][2].replace(' ','')
                tk=tkbill_text.split('(')[0];bill=tkbill_text.split('(')[1][-5:].replace(')','')
                khlb_text=wh_Data[i][3]
                try:
                    data_bill = pd.read_excel(file, sheet_name=bill)
                    print(data_bill)
                except Exception as e:
                    print(e)
            else:
                tkbill_text ='                                                                       '
            tklb=tkinter.Label(canvas1, text=tkbill_text)
            tklb.place(x=130,y=(50+20*i))
            tkbill_label.append(tklb)
            date_label.append(dtlb)
            label_create.append(lb)


def PageChange(axis, event=None):
    global page, maxPage
    print("maxpage: ", maxPage)
    if axis == 0:
        if page == 0:
            pass
        else:
            page=page-1
            if len(label_create) > 0:
                for lb in label_create:
                    lb.after(0, lb.destroy)
            if len(date_label) > 0:
                for lb in date_label:
                    lb.after(0, lb.destroy)
            
    elif axis == 1:
        if page == maxPage:
            pass
        else:
            page=page+1
            if len(label_create) > 0:
                for lb in label_create:
                    lb.after(0, lb.destroy)
            if len(date_label) > 0:
                for lb in date_label:
                    lb.after(0, lb.destroy)
    

def show_like_excel(data_ex, tree):
    data_ex = data_ex.dropna(how='all')
    tree.delete(*tree.get_children())
    tree["column"] = list(data_ex.columns)
    tree["show"] = "headings"
    for col in tree["column"]:
        tree.heading(col, text=col)
    df_rows = data_ex.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values=row)
    tree.pack()
    


# frame = Frame(window, width=400, height=400)
# frame.place(x=10, y=80)
# tree = ttk.Treeview(frame)

Openbtn = tkinter.Button(window, text="Chọn file", command=UploadAction)
Openbtn.place(x=10,y=20)
lbLink = tkinter.Label( window, textvariable = var, relief = tkinter.RAISED )
lbLink.place(x=70,y=22)

canvas1 = tkinter.Canvas(window, width=1100, height=400, bg="lightgray")
canvas1.place(x=10, y=80)
# lbKho = tkinter.Label(window, textvariable = var1, relief = tkinter.RAISED )
# lbKho.place(x=15,y=85)
CreateLabelManager(canvas=canvas1)
lbL = tkinter.Button( window, text="<", command=lambda:PageChange(axis=0))
lbL.place(x=10,y=480)
lbR = tkinter.Button( window, text=">", command=lambda:PageChange(axis=1))
lbR.place(x=30,y=480)

CheckImportStatus()
window.mainloop()