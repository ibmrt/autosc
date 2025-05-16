import tkinter
import tkinter.messagebox
import keyboard
import os
import sys
from tkinter import Tk, StringVar
from ctypes import windll
from autosc_events import *
from autosc_engine import Engine, Process

font_type = ("MingLiU-ExtB", 12)
btn_color = "PaleGreen3"

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class AutoUI:

    def __init__(self, eng:Engine):
        self.engine = eng
        self.running = False
        self.session = None

    def save_path(self, path_input, settings):
        print(f"SENT ENGINE: {SaveFolderPath(path_input.get())}")
        result = self.engine.process(SaveFolderPath(path_input.get()))
        print(f"ENGINE: {result}")
        if isinstance(result, PathSaved):
            settings.configure(text = self.engine.current_setting())
        elif isinstance(result, SavePathFailed):
            tkinter.messagebox.showerror("Error", result.reason)

    def save_window(self, window_size, settings):
        print(f"SENT ENGINE: {StartAcquireWindow()}")
        result = self.engine.process(StartAcquireWindow())
        print(f"ENGINE: {result}")
        if isinstance(result, WindowSizeResult):
            settings.configure(text = self.engine.current_setting())
            window_size.set(f"左上角坐标：({result.left}, {result.top})\t宽高({result.width}, {result.height})")
        elif isinstance(result, SaveWindowSizeFailed):
            tkinter.messagebox.showerror("Error", result.reason)

    def save_settings(self):
        print(f"SENT ENGINE: {SaveSettings()}")
        result = self.engine.process(SaveSettings())
        print(f"ENGINE: {result}")
        if isinstance(result, SettingSaved):
            print(result)
        elif isinstance(result, SaveSettingFailed):
            tkinter.messagebox.showerror("Error", result.reason)

    def toggle(self, btn, status, img_path):
        try:
            if self.running:
                self.session.terminate()
                self.running = False
                self.session = None
                status.configure(text = "未启动", fg = "red")
                btn.configure(text = "启动", bg=btn_color)
            else:
                self.session = Process()
                self.running = True
                print(f"SENT ENGINE: {StartProcess(self.session)}")
                result = self.engine.process(StartProcess(self.session))
                print(f"ENGINE: {result}")
                if isinstance(result, FailedToStartProcess):
                    tkinter.messagebox.showerror("Error", result.reason)
                    self.session = None
                    self.running = False
                elif isinstance(result, StartedProcess):
                    status.configure(text="运行中", fg=btn_color)
                    btn.configure(text="结束", bg="red")
                    img_path.set(result)

        except Exception as e:
            status.configure(text="运行中", fg=btn_color)
            btn.configure(text="结束", bg="red")
            tkinter.messagebox.showerror("Error", f"意料之外的错误: {e}")

    def create_ppt(self, mat_path):
        print(f"SENT ENGINE: {StartCreatePPT(mat_path)}")
        result = self.engine.process(StartCreatePPT(mat_path.get()))
        print(f"ENGINE: {result}")
        if isinstance(result, PPTCreated):
            pass
        elif isinstance(result, FailedToCreatePPT):
            tkinter.messagebox.showerror("Error", result.reason)

    def run_ui(self):
        windll.shcore.SetProcessDpiAwareness(1)

        root = Tk()
        root.title("自动截屏")
        icon = tkinter.PhotoImage(file=resource_path("jie.png"))
        root.iconphoto(False, icon)

        main = tkinter.Frame(root)
        main.configure(padx=50, pady=50)
        main.grid(column=0, row=0)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # row 1
        tkinter.Label(main, text="AutoSC lh", font=("Georgia", 12)).grid(column=1, row=1)
        status = tkinter.Label(main, text="未启动", font=font_type, fg="red")
        status.grid(column=2, row=1)
        execute_button = tkinter.Button(main, text="启动", bg=btn_color, font=font_type, command=lambda: self.toggle(execute_button, status, img_path))
        execute_button.grid(column=3, row=1)


        # row 2
        path = StringVar()
        path_entry = tkinter.Entry(main, width=40, textvariable=path, font=("Georgia", 12))
        path_entry.grid(column=2, row=2, sticky="N")
        tkinter.Label(main, text="文件夹：", font=font_type).grid(column=1, row=2, sticky="N")
        tkinter.Button(main, text="保存路径", bg=btn_color, font=font_type, command=lambda: self.save_path(path_entry, current_setting)).grid(column=3, row=2, sticky="N")

        # row 3
        tkinter.Label(main, text="取景框：", font=font_type).grid(column=1, row=3, sticky="N")
        size = StringVar()
        tkinter.Label(main, textvariable=size, font=("Georgia", 12)).grid(column=2, row=3, sticky="N")
        tkinter.Button(main, text="设置取景框大小", bg=btn_color, font=font_type, command=lambda: self.save_window(size, current_setting)).grid(column=3, row=3, sticky="N")

        # row 4
        tkinter.Label(main, text="目前设置：", font=font_type).grid(column=1, row=4, sticky="N")
        current_setting = tkinter.Label(main, text="", font=("Georgia", 12))
        current_setting.grid(column=2, row=4, sticky="N")
        current_setting.configure(text=self.engine.current_setting())
        tkinter.Button(main, text="保存为常用设置", font=font_type, bg=btn_color, command=lambda: self.save_settings()).grid(column=3, row=4, sticky="N")

        # row 5
        tkinter.Label(main, text="演示稿取材路径：", font=font_type).grid(column=1, row=5, sticky="N")
        img_path = StringVar()
        tkinter.Entry(main, textvariable=img_path, width=40, font=("Georgia", 12)).grid(column=2, row=5, sticky="N")
        tkinter.Button(main, text = "制作", font=font_type, bg="gold", command= lambda: self.create_ppt(img_path)).grid(column=3, row=5, sticky="N")


        for child in main.winfo_children():
            child.grid_configure(padx=50, pady=30)

        path_entry.focus_set()

        keyboard.add_hotkey("Ctrl+Enter", lambda: self.toggle(execute_button, status, img_path))
        root.mainloop()

if __name__ == "__main__":

    engine = Engine()
    engine.load_setting()
    auto = AutoUI(engine)
    auto.run_ui()