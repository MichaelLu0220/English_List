import os
import openpyxl
import time
import tkinter as tk
import win32gui
import random
wb = openpyxl.load_workbook('test.xlsx') 
names = wb.sheetnames   
s1 = wb['file1']  


global i
global m
global lock
global mod
i = 0
m = 0
lock = ''
mod = 0
def refreshText():
    global i
    global s
    global m
    global n
    global mod
    if mod ==1:
        i =random.randint(1,int(s1.max_row))

    else:
        i+=1
        if i > int(s1.max_row):
            i = 1
    text1.delete(0.0,tk.END)
    v = s1.iter_rows(min_row=i, min_col=1, max_col=3, max_row=i)
    a = ''
    for k in v:
        for j in k:
            a+=j.value+"  "
    text1.insert(tk.INSERT,a)
    text1.tag_add("center", 1.0, "end")
    text1.update()
    s = windows.after(n,refreshText)
    

def button_event():
    global m
    global s
    global lock
    global mod 
    if m ==0:
        m=1
        lock+='123'
        if lock =='123456789123':
            print("YES")
            if mod==0:
                a.set("Change Mod to Random")
                windows.after(1000,clear)
                mod = 1
            else:
                a.set("Change Mod to Normal")
                windows.after(1000,clear)
                mod = 0
            
            lock = ''
        elif lock=='123'or lock=='123123':
            lock = '123'
        else:
            lock = ''
    else:
        m = 0
    stop()


def stop():
    global m
    global s
    global i
    global n
    if m==1:
        text1.after_cancel(s)
        text1.delete(0.0,tk.END)
        i+=1
        if i > int(s1.max_row):
            i = 1
        v = s1.iter_rows(min_row=i, min_col=1, max_col=3, max_row=i)
        a = ''
        for k in v:
            for j in k:
                a+=j.value+"  "
        text1.insert(tk.INSERT,a)
        text1.tag_add("center", 1.0, "end")
        m = 0
        stop()
    else:
        s = windows.after(n,refreshText)

def plus_():
    global n
    global lock
    lock+='456'
    if '456456' in lock:
            lock = ''
    n = n + 1000
    show()
    windows.after(1000,clear)
    
def min_():
    global n
    global lock
    lock+='789'
    if '789789' in lock:
            lock = ''
    n = n - 1000
    if n<1000:
        n=1000
        a.set("Too fast")
    else:
        show()
    windows.after(1000,clear)

def show():
    a.set("Time set:"+str(n/1000)+"sec")

def clear():
    a.set('')

windows = tk.Tk()
windows.lift()
windows.attributes('-topmost',True)
windows.title('EnglishList')
windows.geometry('470x80')#400*50
windows.resizable(False,False)

n = 3000
a = tk.StringVar() 
a.set('')
label = tk.Label(windows, textvariable=a)  #  Label


text1 = tk.Text(windows,width = 31,height=1,font=('Times New Roman',17,'bold'))
text1.tag_configure("center", justify='center')
mybutton = tk.Button(windows, text='change', command=button_event)
bt_plus = tk.Button(windows, text='+1', command=plus_)
bt_min = tk.Button(windows, text='-1', command=min_)
text1.grid(row = 0,column=1,rowspan=2,padx=10,pady=10)
mybutton.grid(row = 0,column=2,columnspan=2,padx=10,pady=10)
bt_plus.grid(row = 1,column=2)
bt_min.grid(row = 1,column=3)
label.grid(row = 1,column=1,padx=5,pady=5)
refreshText()


windows.mainloop()
