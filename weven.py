# -*- coding: utf-8 -*-

from tkinter import *
import tkinter.filedialog
import tkinter.messagebox
import re
from EmailUtil import EmailUtil
import xlrd
import requests

filename = "";


def xz():
    global filename
    filename = tkinter.filedialog.askopenfilename()
    if filename != '':
        if re.compile(u'[\u4e00-\u9fa5]+').search(filename.split("/")[-1]):
            lb.config(text="您的文件名包含中文，为防止编码错误，请更改为非中文")
        else:
            lb.config(text="您选择的文件是：" + filename)
    else:
        lb.config(text="您没有选择任何文件");


def load_data():
    if filename.strip() != '':
        if re.compile(u'[\u4e00-\u9fa5]+').search(filename.split("/")[-1]):
            tkinter.messagebox.showinfo("提示", "您的文件名包含中文，为防止编码错误，请更改为非中文")
        else:
            workbook = xlrd.open_workbook(filename)
            sheet_names = workbook.sheet_names()
            for sheet_name in sheet_names:
                sheet2 = workbook.sheet_by_name(sheet_name)
                for index in range(5, len(sheet2.col_values(1))):
                    # print(sheet2.row_values(index))
                    eg = EmailUtil(sheet2.row_values(index), entry1.get())
                    if eg.send_by_email():
                        print("发送至" + sheet2.row_values(index)[3] + "的邮箱成功")
                    else:
                        print("发送至" + sheet2.row_values(index)[3] + "的邮箱失败")
              
                tkinter.messagebox.showinfo("发送结果", "已全部发送成功")
    else:
        tkinter.messagebox.showinfo("提示", "请选择文件")


def about():
    tkinter.messagebox.showinfo("关于", "一键助手1.0版本")


root = Tk()
root.title("工资单一键助手")
# 配置菜单栏
menu = Menu(root)
# menu1 = Menu(menu, tearoff=0)
# menu1.add_command(label="", command=hello)
# menu1.add_command(label="")
# menu.add_cascade(label="", menu=menu1)
#菜单栏子窗口
menu2 = Menu(menu, tearoff=0)
menu2.add_command(label="About", command=about)
menu2.add_command(label="Quit", command=root.quit)
menu.add_cascade(label="Help", menu=menu2)
root.config(menu=menu)
# 显示框
lb = Label(root, text='')
lb.pack()
btn = Button(root, text="请选择excel文件", command=xz)
btn.pack()
e = StringVar()
entry1 = tkinter.Entry(root, width=20, bg="white", fg="black", textvariable=e)
e.set("2018年02月份工资表")
entry1.pack(pady=8)
btn1 = Button(root, text="发送至邮箱", command=load_data)
btn1.pack()
# 设置窗口出现在正中间
curWidth = root.winfo_reqwidth()
curHeight = root.winfo_height()
scnWidth, scnHeight = root.maxsize()
tmpcnf = '%dx%d+%d+%d' % (curWidth, curHeight, (scnWidth - curWidth) / 2, (scnHeight - curHeight) / 2)
root.geometry(tmpcnf)
root.maxsize(400, 300)
root.minsize(300, 200)
root.mainloop()
