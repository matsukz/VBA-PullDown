from asyncore import write
from cgitb import text
from multiprocessing import AuthenticationError
import tkinter, tkinter.messagebox
from tkinter import ttk
import tkinter.font as f
import pyautogui
import pyperclip
import subprocess

root = tkinter.Tk()
root.title("Vba-AUTOMATICENTRY")
root.geometry("250x270")
root.attributes("-topmost", True)
root.resizable(0,0)

mtxtbox = tkinter.Text(font=("", 16))
mtxtbox.place(
    x=60,
    y=10,
    width=150,height=30
    )
mtxtbox.focus_set()
mlbl = tkinter.Label(text="オブジェクト")
mlbl.place(x=0, y=12)

otxtbox = tkinter.Text(font=("", 16))
otxtbox.place(
    x=60,
    y=50,
    width=150,
    height=30
    )
olbl = tkinter.Label(text="メソッド")
olbl.place(x=10, y=52)

mckb = tkinter.BooleanVar()
mckb.set(True)

delckb = tkinter.BooleanVar()
delckb.set(True)

sCount = 0

FXlist = ("Sub", "Range", "AutoFill", "Dim", "MessageBox")
combobox = ttk.Combobox(
    root, 
    values=FXlist, 
    height= 5,
    state="readonly",
    textvariable= tkinter.StringVar(),
    )
combobox.set(FXlist[0])

def exe():
    print(combobox.get())

    msell = ""
    osell = ""

    msell = mtxtbox.get("1.0", "end-1c")
    osell = otxtbox.get("1.0", "end-1c")
    
    if not msell == "" :

        pyautogui.click(48, 0)
        if combobox.get() == "Range":
            try:
                Lmain = "Range(\""
                pyperclip.copy(msell)
                Rmain =  "\")"    

                pyautogui.write(Lmain)
                pyautogui.hotkey("ctrl", "v")
                pyautogui.write(Rmain)

                if not osell=="" :
                    pyperclip.copy(osell)
                    pyautogui.hotkey("ctrl", "v")
                else :
                        pass

                pyautogui.write(Rmain)
                pyautogui.press("Return")

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
        else:
            tkinter.messagebox.showerror("ERROR", "Combobox RANGE ERROR")

        if delckb.get() == True:
            mtxtbox.delete("1.0", "end-1c")
            otxtbox.delete("1.0", "end-1c")
        else:
            pass

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

def Hojyo():
    
    global sCount
    sRadio = tkinter.IntVar()
    sRadio.set([0])
    if sCount == 0 :
        sCount = sCount + 1
        sWindow = tkinter.Toplevel()

        sWindow.title("関数の補助")
        sWindow.geometry("400x400")
        sWindow.protocol(
            "WM_DELETE_WINDOW", 
            (lambda: "pass")()
            )
        sWindow.attributes("-topmost", True)
        sWindow.resizable(0,0)

        FRadio = tkinter.Radiobutton(
            sWindow,
            value=0,
            variable=sRadio
        )
        FRadio.place(x=10, y=30)

        sfont = tkinter.Label(sWindow, text="色の設定")
        sfont.place(x=30, y=30)

        Flist = ("文字の色（Font）", "セルの色（Interior）")
        Fcombobox = ttk.Combobox(
            sWindow, 
            values=Flist, 
            height= 6,
            width= 20,
            state="readonly",
            textvariable= tkinter.StringVar(),
            )
        Fcombobox.set(Flist[0])
        Fcombobox.place(x=100, y=30)

        cIquT = tkinter.Label(sWindow, text="=", font=("nomal", "14", "bold"))
        cIquT.place(x=260, y=27)

        cIquB = tkinter.Text(
            sWindow,
            font=("", "16"),
            width=3,
            height=1
        )
        cIquB.place(x=280, y=30)

        DRadio = tkinter.Radiobutton(
            sWindow,
            value=1,
            variable=sRadio
        )
        DRadio.place(x=10, y=60)

        sdim = tkinter.Label(sWindow, text="変数の型")
        sdim.place(x=30, y=60)

        Dlist = (
            "長整数型（Long）", 
            "倍精度浮動小数点型（Double）", 
            "文字列型（String）", 
            "ALLデータ型（Variant）"
            )
        Dcombobox = ttk.Combobox(
            sWindow, 
            values = Dlist, 
            height = 6,
            width = 25,
            state ="readonly",
            textvariable = tkinter.StringVar(),
            )
        Dcombobox.set(Dlist[0])
        Dcombobox.place(x=100, y=60)

        def mainclip():
            psRadio = sRadio.get()
            print(psRadio)

        def close():
            global sCount
            sCount = sCount - 1
            sWindow.destroy()

        mainclip_button = tkinter.Button(sWindow, text="クリップ",command=mainclip,width=16,height=3)
        mainclip_button.place(x=180,y=130)

        calc_button = tkinter.Button(sWindow, text="閉じる",command=close,width=16,height=3)
        calc_button.place(x=60,y=130)

    else:
        print("sCount ERROR!! sCount is " + str(sCount))
        pass

Execute_button = tkinter.Button(text="Execute",command=exe,width=16,height=3)
Execute_button.place(x=70,y=140)

combobox.place(x=35, y=100, width=150, height=25)

Hojyo_button = tkinter.Button(text="関数補助",command=Hojyo,width=7,height=1)
Hojyo_button.place(x=190,y=100)

mdelete_button = tkinter.Button(text="削除",command=mdelete,width=3,height=1)
mdelete_button.place(x=215,y=12)

odelete_button = tkinter.Button(text="削除",command=odelete,width=3,height=1)
odelete_button.place(x=215,y=52)

ckbox = tkinter.Checkbutton(root, variable=mckb, text="Cells.deleteの入力")
ckbox.place(x=10, y=200)

delckbox = tkinter.Checkbutton(root, variable=delckb, text="実行後のクリア")
delckbox.place(x=140, y=200)

root.mainloop()
