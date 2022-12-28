import tkinter, tkinter.messagebox
from tkinter import ttk
import tkinter .font as f
from tkinter import filedialog

import os
import subprocess
import json

AskAgree = tkinter.Tk()
AskAgree.title("システムを起動する前に")
AskAgree.geometry("580x300")
AskAgree.resizable(0,0)

Warning_message_title = tkinter.Label(
    AskAgree,
    text="このソフトウェアを利用するにあたって",
    font=(
        "",
        "15",
        "bold"
    )
)
Warning_message_title.place(x=70,y=10)

Warning_message_1 = tkinter.Label(
    AskAgree,
    text="以下の事項を必ずご確認ください",
    font=(
        "",
        "10",
        "bold"
    )
)
Warning_message_1.place(x=160,y=40)

Warning_message_2 = tkinter.Label(
    AskAgree,
    text="・利用するPCをインターネットに接続してください。（通信料は利用者負担）",
    font=(
        "",
        "10"
    )
)
Warning_message_2.place(x=60,y=70)

Warning_message_3 = tkinter.Label(
    AskAgree,
    text="・Windows10のみ動作テストを行いました。(Windows以外では正常に動作しません)",
    font=(
        "",
        "10"
    )
)
Warning_message_3.place(x=60,y=100)

Warning_message_4 = tkinter.Label(
    AskAgree,
    text="・Aviutlをすでに導入されているならば必ずバックアップを行ってください。",
    font=(
        "",
        "10"
    )
)
Warning_message_4.place(x=60,y=130)

Warning_message_5 = tkinter.Label(
    AskAgree,
    text="・いかなる損害にも製作者は対応しません。",
    font=(
        "",
        "10"
    )
)
Warning_message_5.place(x=60,y=160)

Warning_message_6 = tkinter.Label(
    AskAgree,
    text="・その他のルールは製作者のGitHubにある本プログラムのReadmeをご確認ください。",
    font=(
        "",
        "10"
    )
)
Warning_message_6.place(x=60,y=160)

Warning_message_7 = tkinter.Label(
    AskAgree,
    text="上記の項目を確認の上、「同意」を選択してください。",
    font=(
        "",
        "10",
        "bold"
    )
)
Warning_message_7.place(x=80,y=190)


def main():

    AskAgree.destroy()

    root = tkinter.Tk()
    root.title("Aviutl Auto Setup")
    root.geometry("300x350")
    root.resizable(0,0)

    Select_version_lavel = tkinter.Label(
        root,
        text="1.ソフトウェアの選択",
        font=(
            "",
            "10",
            "bold"
        )
    )
    Select_version_lavel.place(x=10,y=10)

    Aviutl_exe_ckb = tkinter.BooleanVar()
    Aviutl_exe_ckb.set(True)
    Aviutl_exe_checkbox = tkinter.Checkbutton(
        root,
        variable=Aviutl_exe_ckb
    )
    Aviutl_exe_checkbox.place(x=30,y=50)

    Aviutl_exe_checkbox = tkinter.Checkbutton(
        root,
        variable=Aviutl_exe_ckb
    )
    Aviutl_exe_checkbox.place(x=30,y=50)

    Aviutl_exe_List = (
        "Aviutl Version1.10",
    )

    Aviutl_exe_Combobox = ttk.Combobox(
        root,
        values=Aviutl_exe_List,
        height=5,
        width=30,
        state="readonly",
        textvariable= tkinter.StringVar()
    )
    Aviutl_exe_Combobox.set(Aviutl_exe_List[0])
    Aviutl_exe_Combobox.place(x=60,y=50)

    Extend_Editor_ckb = tkinter.BooleanVar()
    Extend_Editor_ckb.set(True)
    Extend_Editor_checkbox = tkinter.Checkbutton(
        root,
        variable=Extend_Editor_ckb
    )
    Extend_Editor_checkbox.place(x=30,y=80)

    Extend_Editor_List = (
        "拡張編集プラグイン Version0.92",
    )

    Extend_Editor_Combbox = ttk.Combobox(
        root,
        values=Extend_Editor_List,
        height=10,
        width=30,
        state="readonly",
        textvariable=tkinter.StringVar()
    )
    Extend_Editor_Combbox.set(Extend_Editor_List[0])
    Extend_Editor_Combbox.place(x=60,y=80)

    LSMASH_ckb = tkinter.BooleanVar()
    LSMASH_ckb.set(True)
    LSMASH_checkbox = tkinter.Checkbutton(
        root,
        variable=LSMASH_ckb,
    )
    LSMASH_checkbox.place(x=30,y=110)

    LSMASH_List = (
        "L-SMASH-Works_Rev1100",
    )

    LSMASH_combobox = ttk.Combobox(
        root,
        values=LSMASH_List,
        height=10,
        width=30,
        state="readonly",
        textvariable=tkinter.StringVar()
    )
    LSMASH_combobox.set(LSMASH_List[0])
    LSMASH_combobox.place(x=60,y=110)

    Encoder_ckb = tkinter.BooleanVar()
    Encoder_ckb.set(True)
    Encoder_checkbox = tkinter.Checkbutton(
        root,
        variable=Encoder_ckb
    )
    Encoder_checkbox.place(x=30,y=140)

    Encoder_List = (
        "x264guiEx",
        "かんたんMP4出力"
    )

    Encoder_combobox = ttk.Combobox(
        root,
        values=Encoder_List,
        height=10,
        width=30,
        state="readonly",
        textvariable=tkinter.StringVar()
    )
    Encoder_combobox.set(Encoder_List[0])
    Encoder_combobox.place(x=60,y=140)

    Setup_folder_path_lavel = tkinter.Label(
        root,
        text="2.構築先フォルダを選択",
        font=(
            "",
            "10",
            "bold"
        )
    )
    Setup_folder_path_lavel.place(x=10,y=170)

    Setup_folder_path_textbox = tkinter.Entry(
        root,
        font=(
            "",
            12
        )
    )
    Setup_folder_path_textbox.place(
        x=30,
        y=200,
        width=180,
        height=30
    )

    def Select_Setup_folder_path():
        Setup_folder_path = (filedialog.askdirectory())

        if not Setup_folder_path == "":

            if not Setup_folder_path_textbox.get() == "":
                Setup_folder_path_textbox.delete("0","end")
            else:
                pass

            Setup_folder_path_textbox.insert(
                0,
                Setup_folder_path.replace("/","\\") + "\\"
            )
        else:
            pass

    def Execute():

        path = Setup_folder_path_textbox.get()

        if path =="":
            tkinter.messagebox.showerror("エラー", "構築先を選択してください。")
        else:
            if os.path.exists("Cache") == True:
                Cache_Check = tkinter.messagebox.askquestion(
                    "エラー",
                    "一時フォルダ\"Cache\"がすでに存在します。\n 削除してもよろしいですか？"
                )
                if Cache_Check == "yes":
                    subprocess.run("rd /S /Q Cache",shell=True)
                else:
                    tkinter.messagebox.showerror("エラー","処理を継続できません")
                    return
            else:
                pass

            subprocess.run("mkdir Cache",shell=True)

            if os.path.exists(path + "Plugins") == False:
                subprocess.run("mkdir " + path + "Plugins",shell=True)
            else:
                pass

            URL_json = json.load(open("DownloadLinks.json","r"))

            if Aviutl_exe_ckb.get() == True:

                print("Downloading Aviutl...")
                subprocess.run(
                    "powershell -Command \"(New-Object Net.WebClient).DownloadFile(\'" + URL_json["Aviutlexe"]["URL"] + "\', 'Cache\\Aviutl110.zip')\"",
                    shell=True
                )

                print("Unzip Aviutl.zip...")
                subprocess.run(
                    "powershell -command \"Expand-Archive Cache\\Aviutl110.zip Cache\\Aviutl110\"",
                    shell=True
                )

                print("Moving Aviutl.exe...")
                subprocess.run(
                    "move /Y Cache\\aviutl110\\* " + path,
                    shell=True
                )
                if os.path.exists(path+"Aviutl.exe") == True:
                    print("Aviutl.exe Check OK")
                else:
                    tkinter.messagebox.showerror("エラー","Aviutl.exeの取得に失敗しました")
                    return

            else:
                print("NOT Check")

            if Extend_Editor_ckb.get() == True:
                print("Downloading Extend_Editor...")
                subprocess.run(
                    "powershell -Command \"(New-Object Net.WebClient).DownloadFile(\'" + URL_json["ExtendEditor"]["URL"] + "\', 'Cache\\exedit92.zip')\"",
                    shell=True
                )

                print("Unzip exedit92.zip...")
                subprocess.run(
                    "powershell -command \"Expand-Archive Cache\\exedit92.zip Cache\\exedit92\"",
                    shell=True
                )

                print("Moving exedit92...")
                subprocess.run(
                    "move /Y Cache\\exedit92\\* " + path,
                    shell=True
                )
                if os.path.exists(path+"exedit.auf") == True:
                    print("exedit.auf Check OK")
                else:
                    tkinter.messagebox.showerror("エラー","拡張編集プラグインの取得に失敗しました")
                    return
            else:
                print("NOT Check")

            if LSMASH_ckb.get() == True:
                print("Downloading L-SMASH...")
                subprocess.run(
                    "powershell -Command \"(New-Object Net.WebClient).DownloadFile(\'" + URL_json["L-SMASH"]["URL"] + "\', 'Cache\\L-SMASH-Works.zip')\"",  
                    shell=True 
                )

                print("Unzip L-SMASH...")
                subprocess.run(
                    "powershell -command \"Expand-Archive Cache\\L-SMASH-Works.zip Cache\\L-SMASH-Works\"",
                    shell=True
                )

                print("Moving L-SMASH...")
                plugins_path = path + "Plugins\\"
                subprocess.run(
                    "move /Y Cache\\L-SMASH-Works\\*.auf " + plugins_path,
                    shell=True
                )
                subprocess.run(
                    "move /Y Cache\\L-SMASH-Works\\lwcolor.auc " + plugins_path,
                    shell=True
                )
                subprocess.run(
                    "move /Y Cache\\L-SMASH-Works\\lwinput.aui " + plugins_path,
                    shell=True
                )

                if os.path.exists(plugins_path+"lwinput.aui") == True:
                    print("lwinput.aui Check OK")
                else:
                    tkinter.messagebox.showerror("エラー","L-SMASH-Worksの取得に失敗しました")
                    return
            else:
                print("NOT Check")

            if Encoder_ckb.get() == True:
                if Encoder_combobox.get() == "x264guiEx":

                    print("Downloading x264guiEx...")
                    subprocess.run(
                        "powershell -Command \"(New-Object Net.WebClient).DownloadFile(\'" + URL_json["x264guiEx"]["URL"] + "\', 'Cache\\x264guiEx.zip')\"",  
                        shell=True 
                    )

                    print("Unzip x264guiEx...")
                    subprocess.run(
                        "powershell -command \"Expand-Archive Cache\\x264guiEx.zip Cache\\x264guiEx\"",
                        shell=True
                    )

                    print("Moving x264guiEx...")

                    subprocess.run(
                        "move /Y Cache\\x264guiEx\\x264gui_3.16\\x264guiEx_readme.txt " + path,
                        shell=True
                    )
                    if os.path.exists(path + "exe_files") == True:
                        pass
                    else:
                        subprocess.run(
                            "mkdir " + path + "exe_files",
                            shell=True
                        )

                    subprocess.run(
                        "echo D | xcopy Cache\\x264guiEx\\x264guiEx_3.16\\exe_files\\* " + path + "exe_files /E /H /C /I",
                        shell=True 
                    )

                    if os.path.exists(path+"Plugins") == True:
                        pass
                    else:
                        subprocess.run(
                            "mkdir " + path + "Plugins",
                            shell=True
                        )

                    subprocess.run(
                        "echo D | xcopy Cache\\x264guiEx\\x264guiEx_3.16\\plugins\\* " + path + "Plugins /E /H /C /I",
                        shell=True
                    )

                    if os.path.exists(path + "Plugins\\x264guiEx.auo") == True:
                        print("x264guiEx.auo Check Ok")
                    else:
                        tkinter.messagebox.showerror("エラー", "x264guiExの取得に失敗しました")
                        return

                elif Encoder_combobox.get() == "かんたんMP4出力":
                    if os.path.exists(path + "Plugins") == True:
                        pass
                    else:
                        subprocess.run(
                            "mkdir " + path + "Plugins",
                            shell=True
                        )
                    
                    print("Downloading easymp4...")
                    subprocess.run(
                        "powershell -Command \"(New-Object Net.WebClient).DownloadFile(\'" + URL_json["easyMP4"]["URL"] + "\', 'Cache\\easymp4.zip')\"",  
                        shell=True 
                    )

                    print("Unzip easymp4...")
                    subprocess.run(
                        "powershell -command \"Expand-Archive Cache\\easymp4.zip Cache\\easymp4\"",
                        shell=True
                    )

                    print("Moving easymp4...")

                    subprocess.run(
                        "move /Y Cache\\easymp4\\easymp4.auo " + path + "Plugins",
                        shell=True
                    )

                    if os.path.exists(path + "Plugins\\easymp4.auo") == True:
                        print("easymp4 Check OK")
                    else:
                        tkinter.messagebox.showerror("エラー", "easymp4の取得に失敗しました")
                        return
                else:
                    tkinter.messagebox.showerror("エラー", "\"Encoder\"にて致命的なエラーが発生しました")
                    return
            else:
                print("NOT Check")

            if os.path.exists("Cache") == True:
                print("Delete Cache")
                subprocess.run(
                    "rd /S /Q Cache",
                    shell=True
                )
            else:
                pass

            tkinter.messagebox.showinfo("Aviutl Auto Setup","指定されたソフトウェアの設定が完了しました")


    Setup_folder_path_button = tkinter.Button(
        root,
        text="参照",
        command=Select_Setup_folder_path,
        width=6,
        height=2
    )
    Setup_folder_path_button.place(x=220,y=196)

    Exceute_button = tkinter.Button(
        root,
        text="実行",
        command=Execute,
        width=20,
        height=3
    )
    Exceute_button.place(x=60,y=250)

    root.mainloop()

def Refuse():
    AskAgree.destroy()

Agree_button = tkinter.Button(
    AskAgree,
    text="同意して進む",
    command=main,
    width=20,
    height=3
)
Agree_button.place(x=300,y=220)

Agree_button = tkinter.Button(
    AskAgree,
    text="拒否（終了）",
    command=Refuse,
    width=20,
    height=3
)
Agree_button.place(x=100,y=220)

AskAgree.mainloop()


