import tkinter as tk
from tkinter import filedialog,messagebox
import pandas as pd
import matplotlib.pyplot as plt
import warnings
import xlwt

warnings.filterwarnings('ignore')

root = tk.Tk()
root.title("图表生成工具")
root.geometry('800x600')

filename =tk.Entry(root,width=100)
slotsid = tk.Entry(root,width=50)
startdate =tk.Entry(root,width=50)
enddate =tk.Entry(root,width=50)

result1 = tk.Entry(root,width=100)
result2 = tk.Entry(root,width=100)
result3 = tk.Entry(root,width=100)
result4 = tk.Entry(root,width=100)
result5 = tk.Entry(root,width=100)

filename.pack()
slotsid.pack()
startdate.pack()
enddate.pack()
result1.pack()
result2.pack()
result3.pack()
result4.pack()
result5.pack()

lb = tk.Label(root,text="请依次输入机台id，想要查询的起始日期，想要查询的终止日期(格式为：xx-xx)").pack()
def get_path():
    a =[]
    path = filedialog.askopenfilename(title='选择文件')
    a.append(path)
    filename.insert("insert",path)
    return path

buttonselect = tk.Button(root,text='选择文件',command=get_path).pack()

def sum5(a):#前五个机台平均数
    s=0
    a.sort()
    b = a[-6:-1]
    for i in range(0,5):
        s += b[i]
    return s/5

def sum10(a):#前十个机台平均数
    a.sort()
    s=0
    b = a[-11:-1]
    for i in range(0,10):
        s += b[i]
    return s/10

def sum20(a):#前二十个机台平均数
    a.sort()
    s=0
    b = a[-21:-1]
    for i in range(0,20):
        s += b[i]
    return s/20

def sumall(a): #所有机台平均数
    s=0
    for i in range(0,len(a)-1):
        s += a[i]
    return s/(len(a)-1)

def jiequ(a,b,c):
    indexs = 0
    indexe = 0
    for i in range(0,len(a)):
        if a[i] == b:
            indexs = i
            break
    for i in range(indexs,len(a)):
        if a[i] == c:
            indexe = i
            break
    return(a[indexs:indexe+1])

def length(a,b,c): #取指定日期长度和下角标
    d = []
    indexs = 0
    indexe = 0
    for i in range(0,len(a)):
        if a[i] == b:
            indexs = i
            break
    for i in range(indexs,len(a)):
        if a[i] == c:
            indexe = i
            break
    d.append(indexs)
    d.append(indexe)
    return(d)

def percentage(a): #小数转百分数
    d = []
    for i in range(0,len(a)):
        d.append(format(a[i],'.2%'))
    return d

def clickbutton1():
    data = pd.read_excel(filename.get())
    id = slotsid.get()
    date = data.iloc[:, 0]  # 日期列
    count = data.iloc[:, 6]  # 启动人数列
    slotid = data.iloc[:, 2]  # 机台id列
    date_start = startdate.get() #开始日期
    date_end = enddate.get() #结束日期

    newdate = list(set(date))
    newdate.sort()
    riqi = []
    for i in range(0,len(newdate)):
        riqi.append(newdate[i][5:])

    tool = []  # 总点击人数下角标集合
    j = 0
    while j < len(slotid):
        if len(slotid[j]) == 2:
            tool.append(j)
        j += 1

    renshu = []  # 每一天总启动人数
    k = 0
    while k < len(tool):
        renshu.append(count[tool[k]])
        k += 1

    merge = list(zip(slotid, count))
    small = []
    slot = []
    perslot = []
    sumsplayer = []
    per5 = []
    per10 = []
    per20 = []
    perall= []
    l = 0
    for j in range(0, len(count)):  # 前N机台点击人数占比
        if len(merge[j][0]) == 2:
            for i in range(l , j+1):
                small.append(count[i])
            sumsplayer.append(count[j])
            per5.append(sum5(small)/count[j])
            per10.append(sum10(small)/count[j])
            per20.append(sum20(small)/count[j])
            perall.append(sumall(small)/count[j])
            small.clear()
            l = j+1
    for k in range(0,len(count)):
        if merge[k][0] == id:
            slot.append(count[k])
    for o in range(0,len(slot)):
        perslot.append(slot[o]/sumsplayer[o])

    result1.insert("insert", percentage(perslot[length(riqi,date_start,date_end)[0]:length(riqi,date_start,date_end)[1]+1]))
    result2.insert("insert", percentage(per5[length(riqi,date_start,date_end)[0]:length(riqi,date_start,date_end)[1]+1]))
    result3.insert("insert", percentage(per10[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1]))
    result4.insert("insert", percentage(per20[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1]))
    result5.insert("insert", percentage(perall[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1]))

def clickbutton2():
    data = pd.read_excel(filename.get())

    id = slotsid.get()
    date = data.iloc[:, 0]  # 日期列
    numofspin =data.iloc[:,7]  # SPIN次数数列
    slotid = data.iloc[:, 2]  # 机台id列
    newdate = list(set(date))
    date_start = startdate.get()  # 开始日期
    date_end = enddate.get()  # 结束日期
    newdate.sort()
    riqi = []
    i = len(newdate)
    for i in range(0, len(newdate)):
        riqi.append(newdate[i][5:])

    tool = []  # 总点击人数下角标集合
    j = 0
    while j < len(slotid):
        if len(slotid[j]) == 2:
            tool.append(j)
        j += 1

    renshu = []  # 每一天总启动人数
    k = 0
    while k < len(tool):
        renshu.append(numofspin[tool[k]])
        k += 1

    merge = list(zip(slotid, numofspin))
    small = []
    slot = []
    perslot = []
    sumsplayer = []
    per5 = []
    per10 = []
    per20 = []
    perall = []
    l = 0
    for j in range(0, len(numofspin)):  # 前N机台点击人数占比
        if len(merge[j][0]) == 2:
            for i in range(l, j + 1):
                small.append(numofspin[i])
            sumsplayer.append(numofspin[j])
            per5.append(sum5(small)/numofspin[j])
            per10.append(sum10(small)/numofspin[j])
            per20.append(sum20(small)/numofspin[j])
            perall.append(sumall(small)/numofspin[j])
            small.clear()
            l = j + 1
    for k in range(0, len(numofspin)):
        if merge[k][0] == id:
            slot.append(numofspin[k])
    for o in range(0, len(slot)):
        perslot.append(slot[o] / sumsplayer[o])

    result1.insert("insert", percentage(perslot[length(riqi,date_start,date_end)[0]:length(riqi,date_start,date_end)[1]+1]))
    result2.insert("insert", percentage(per5[length(riqi,date_start,date_end)[0]:length(riqi,date_start,date_end)[1]+1]))
    result3.insert("insert", percentage(per10[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1]))
    result4.insert("insert", percentage(per20[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1]))
    result5.insert("insert", percentage(perall[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1]))

def clickbutton3():
    data = pd.read_excel(filename.get())
    id = slotsid.get()
    date = data.iloc[:, 0]  # 日期列
    numofavg =data.iloc[:,8]  # 平均SPIN次数数列
    slotid = data.iloc[:, 2]  # 机台id列
    date_start = startdate.get()  # 开始日期
    date_end = enddate.get()  # 结束日期
    newdate = list(set(date))
    newdate.sort()
    riqi = []
    i = len(newdate)
    for i in range(0, len(newdate)):
        riqi.append(newdate[i][5:])

    tool = []  # 总点击人数下角标集合
    j = 0
    while j < len(slotid):
        if len(slotid[j]) == 2:
            tool.append(j)
        j += 1

    renshu = []  # 每一天总启动人数
    k = 0
    while k < len(tool):
        renshu.append(numofavg[tool[k]])
        k += 1

    merge = list(zip(slotid, numofavg))
    small = []
    slot = []
    per5 = []
    per10 = []
    per20 = []
    perall = []
    l = 0
    for j in range(0, len(numofavg)):  # 前N机台点击人数占比
        if len(merge[j][0]) == 2:
            for i in range(l, j + 1):
                small.append(numofavg[i])

            per5.append(sum5(small))
            per10.append(sum10(small))
            per20.append(sum20(small))
            perall.append(sumall(small))
            small.clear()
            l = j + 1
    for k in range(0, len(numofavg)):
        if merge[k][0] == id:
            slot.append(numofavg[k])

    result1.insert("insert", slot[length(riqi,date_start,date_end)[0]:length(riqi,date_start,date_end)[1]+1])
    result2.insert("insert", per5[length(riqi,date_start,date_end)[0]:length(riqi,date_start,date_end)[1]+1])
    result3.insert("insert", per10[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1])
    result4.insert("insert", per20[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1])
    result5.insert("insert", perall[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1])

def clickbutton4():
    data = pd.read_excel(filename.get())
    id = slotsid.get()
    date = data.iloc[:, 0]  # 日期列
    numofcost =data.iloc[:,15]  # SPIN消耗钱数列
    slotid = data.iloc[:, 2]  # 机台id列
    date_start = startdate.get()  # 开始日期
    date_end = enddate.get()  # 结束日期
    newdate = list(set(date))
    newdate.sort()
    riqi = []
    i = len(newdate)
    for i in range(0, len(newdate)):
        riqi.append(newdate[i][5:])

    tool = []  # 总点击人数下角标集合
    j = 0
    while j < len(slotid):
        if len(slotid[j]) == 2:
            tool.append(j)
        j += 1

    renshu = []  # 每一天总启动人数
    k = 0
    while k < len(tool):
        renshu.append(numofcost[tool[k]])
        k += 1

    merge = list(zip(slotid, numofcost))
    small = []
    slot = []
    perslot = []
    sumsplayer = []
    per5 = []
    per10 = []
    per20 = []
    perall = []
    l = 0
    for j in range(0, len(numofcost)):  # 前N机台点击人数占比
        if len(merge[j][0]) == 2:
            for i in range(l, j + 1):
                small.append(numofcost[i])
            sumsplayer.append(numofcost[j])
            per5.append(sum5(small) / numofcost[j])
            per10.append(sum10(small) / numofcost[j])
            per20.append(sum20(small) / numofcost[j])
            perall.append(sumall(small) / numofcost[j])
            small.clear()
            l = j + 1
    for k in range(0, len(numofcost)):
        if merge[k][0] == id:
            slot.append(numofcost[k])
    for o in range(0, len(slot)):
        perslot.append(slot[o] / sumsplayer[o])

    result1.insert("insert", percentage(perslot[length(riqi,date_start,date_end)[0]:length(riqi,date_start,date_end)[1]+1]))
    result2.insert("insert", percentage(per5[length(riqi,date_start,date_end)[0]:length(riqi,date_start,date_end)[1]+1]))
    result3.insert("insert", percentage(per10[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1]))
    result4.insert("insert", percentage(per20[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1]))
    result5.insert("insert", percentage(perall[length(riqi, date_start, date_end)[0]:length(riqi, date_start, date_end)[1] + 1]))

def clickbutton5():
    result1.delete(0,'end')
    result2.delete(0, 'end')
    result3.delete(0, 'end')
    result4.delete(0, 'end')
    result5.delete(0, 'end')

button1 = tk.Button(root,text="点击人数占比",command=clickbutton1).pack(padx=0,pady=10)
button2 = tk.Button(root,text="SPIN次数占比",command=clickbutton2).pack(padx=20,pady=10)
button3 = tk.Button(root,text="平均SPIN次数",command=clickbutton3).pack(padx=0,pady=20)
button4 = tk.Button(root,text="金币花费次数占比",command=clickbutton4).pack(padx=20,pady=20)
button5 = tk.Button(root,text="清除所有",command=clickbutton5).pack(padx=20,pady=20)

root.mainloop()

