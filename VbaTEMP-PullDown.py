from asyncore import write
from cgitb import text
from email import message
from math import comb
from multiprocessing import AuthenticationError
import tkinter, tkinter.messagebox
from tkinter import ttk
from numpy import true_divide
import pyautogui
import pyperclip
import subprocess

root = tkinter.Tk()
root.title("VbaTEMPLATES")
root.geometry('250x270')
root.attributes("-topmost", True)
root.resizable(0,0)

mtxtbox = tkinter.Text(font=("", 16))
mtxtbox.place(x=60, y=10, width=150, height=30)
mtxtbox.focus_set()
mlbl = tkinter.Label(text='メイン')
mlbl.place(x=25, y=12)

otxtbox = tkinter.Text(font=("", 16))
otxtbox.place(x=60, y=50, width=150, height=30)
olbl = tkinter.Label(text='オプション')
olbl.place(x=10, y=52)

mckb = tkinter.BooleanVar()
mckb.set(True)

delckb = tkinter.BooleanVar()
delckb.set(True)

list = ("Sub", "Range", "AutoFill", "Dim", "MessageBox")
combobox = ttk.Combobox(
    root, 
    values=list, 
    height= 5,
    state="readonly",
    textvariable= tkinter.StringVar(),
    )
combobox.set(list[0])

def exe():
    print(combobox.get())

    msell = ""
    osell = ""

    msell = mtxtbox.get("1.0", "end-1c")
    osell = otxtbox.get("1.0", "end-1c")
    pyautogui.click(48, 0)

    if not msell == "" :
        
        if combobox.get() == "Range":
            try:
                if not osell == "" :
                    Lmain = "Range(\""
                    Rmain =  "\")"      

                    pyautogui.click(48, 0)
                    pyautogui.write(Lmain + msell + Rmain)

                    if not osell=="" :
                        pyperclip.copy(osell)
                        pyautogui.hotkey("ctrl", "v")
                    
                    else :
                        pass

                    pyautogui.press("Return")

                    if delckb.get() == True:
                        mtxtbox.delete("1.0", "end-1c")
                        otxtbox.delete("1.0", "end-1c")
                    else:
                        pass

            except Exception as e :
                tkinter.messagebox.showerror("ERROR", "範囲：Range")

        elif combobox.get()=="Sub":
            try:
                Copy = ""
                pyperclip.copy(msell)
                print(Copy)

                pyautogui.write("Sub ")
                pyautogui.hotkey("ctrl", "v")

                pyautogui.write(" ()")
                pyautogui.press("Return")
                pyautogui.press("Tab")

                if mckb.get()==True :
                    pyautogui.write("Cells.delet")
                
                pyautogui.press("Return")

                if delckb.get() == True:
                        mtxtbox.delete("1.0", "end-1c")
                        otxtbox.delete("1.0", "end-1c")
                else:
                    pass
            
            except Exception as e :
                tkinter.messagebox.showerror("ERROR", "範囲：sub")

        elif combobox.get()=="AutoFill":

            try:
                Lmain = "Range(\""
                Nmain = ".Autofill Destination:="
                Rmain = "\")"

                pyautogui.write(Lmain)

                pyperclip.copy(msell)
                pyautogui.hotkey("ctrl", "v")

                pyautogui.write(Rmain)
                pyautogui.write(Nmain)
                pyautogui.write(Lmain)

                pyperclip.copy(osell)
                pyautogui.hotkey("ctrl", "v")
                    
                pyautogui.write(Rmain)
                pyautogui.press("Return")              
            
            except Exception as e :
                tkinter.messagebox.showerror("ERROR", "範囲：AutoFill")
            
            if delckb.get() == True:
                        mtxtbox.delete("1.0", "end-1c")
                        otxtbox.delete("1.0", "end-1c")
            else:
                pass 
        elif combobox.get() == "Dim":
            try:

                pyautogui.write("Dim ")
                pyperclip.copy(msell)
                pyautogui.hotkey("ctrl", "v")

                if not osell == "" :
                    pyautogui.write(" As " + osell)
                else :
                    pass
                
                pyautogui.press("Return")

            except Exception as e :
                tkinter.messagebox.showerror("ERROR", "範囲：Dim")
            
            if delckb.get() == True :
                mtxtbox.delete("1.0", "end-1c")
                otxtbox.delete("1.0", "end-1c")

        elif combobox.get() == "MessageBox" :
            try:
                pyautogui.write("Msgbox \"")
                pyperclip.copy(msell)
                pyautogui.hotkey("ctrl", "v")
                pyautogui.write("\"")
                pyautogui.press("Return")
            
            except Exception as e :
                tkinter.messagebox.showerror("ERROR", "範囲：MessageBox")

        subprocess.run("echo off | clip", shell=True)
        mtxtbox.focus()
    else:
        tkinter.messagebox.showerror("ERROR", "文字を入力してください。")

def mdelete():
    mtxtbox.delete("1.0", "end-1c")
    mtxtbox.focus()

def odelete():
    otxtbox.delete("1.0", "end-1c")
    otxtbox.focus()

Execute_button = tkinter.Button(text='Execute',command=exe,width=16,height=3)
Execute_button.place(x=70,y=140)

combobox.place(x=60, y=100, width=150, height=25)

mdelete_button = tkinter.Button(text="削除",command=mdelete,width=3,height=1)
mdelete_button.place(x=215,y=12)

odelete_button = tkinter.Button(text="削除",command=odelete,width=3,height=1)
odelete_button.place(x=215,y=52)

ckbox = tkinter.Checkbutton(root, variable=mckb, text="Cells.deleteの入力")
ckbox.place(x=10, y=200)

delckbox = tkinter.Checkbutton(root, variable=delckb, text="実行後のクリア")
delckbox.place(x=140, y=200)

root.mainloop()