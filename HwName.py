import os
import shutil
import subprocess
import sys
import time
import tkinter

import pywintypes
import winshell
import tkinter.messagebox
import ttkbootstrap as ttk
from ttkbootstrap import *

from WinReg import WinReg

CONFIG_CPU = [
    ("HARDWARE\\DESCRIPTION\\System", "CentralProcessor", "ProcessorNameString"),
    ("SYSTEM\\ControlSet001\\Enum\\ACPI", "GenuineIntel", "FriendlyName")
]


class FindResource:
    @staticmethod
    def get(relative_path) -> str:
        if hasattr(sys, '_MEIPASS'):  # 如果是打包后的程序
            return str(os.path.join(sys._MEIPASS, relative_path))
        return os.path.join(os.path.abspath("."), relative_path)


class HardwareName:
    def __init__(self, in_args):
        # 创建窗口 ------------------------------------------------------------
        self.root = tk.Tk()
        self.head = ttk.Style()
        self.root.geometry("353x90")
        self.root.title("CPU名称修改工具 for Windows")
        # 字体设置 ------------------------------------------------------------
        self.font = tkinter.font.Font(family="MapleMono SC NF", size=20)
        self.head.configure("TNotebook.Tab", font=self.font)
        self.head.configure("TFrame", font=self.font)
        self.head.configure("TLabel", font=self.font)
        # 界面配置 ============================================================
        self.cpu_name_var = tk.StringVar()
        self.cpu_name_tab = ttk.Label(self.root, text="CPU名称：")
        self.cpu_name_con = ttk.Entry(self.root, width=36)
        self.exe_once_btn = ttk.Button(self.root, text="本次临时修改CPU名称")
        self.exe_full_btn = ttk.Button(self.root, text="设置开机修改CPU名称")
        self.cpu_name_tab.grid(column=0, row=0, pady=10, padx=5, sticky="nsew")
        self.cpu_name_con.grid(column=1, row=0, pady=10, padx=5, columnspan=3)
        self.exe_once_btn.grid(column=0, row=1, padx=5, columnspan=2, sticky="nsew")
        self.exe_full_btn.grid(column=3, row=1, padx=5, columnspan=2, sticky="nsew")
        self.cpu_name_con.config(textvariable=self.cpu_name_var)
        self.exe_once_btn.config(bootstyle="success")
        self.exe_full_btn.config(bootstyle="info")
        self.exe_once_btn.config(command=self.change)
        self.exe_full_btn.config(command=self.server)
        self.detect()
        if len(in_args) >= 3 and in_args[1] == "server":
            self.change(in_args[2])
            exit(0)
        self.root.mainloop()

    def detect(self):
        cpu_data = WinReg(CONFIG_CPU[0][0])
        cpu_data = cpu_data.open_sub(CONFIG_CPU[0][1])
        cpu_last = cpu_data.open_sub("0")
        cpu_name = cpu_last.read_var(CONFIG_CPU[0][2])[0]
        self.cpu_name_var.set(cpu_name)
        srv_text = HardwareName.NssmUI("status HwName")
        if srv_text.find("Can't open service") == -1:
            self.exe_full_btn.config(text="取消开机修改CPU名称")
            self.exe_full_btn.config(bootstyle="danger")
        else:
            self.exe_full_btn.config(text="设置开机修改CPU名称")
            self.exe_full_btn.config(bootstyle="info")

    def change(self, in_name=None, in_show=True):
        cpu_name = self.cpu_name_var.get()
        if in_name is not None:
            cpu_name = in_name.replace("`", " ")
        for cpu_item in CONFIG_CPU:
            cpu_data = WinReg(cpu_item[0])
            cpu_real = cpu_data.find_sub(cpu_item[1])
            cpu_data = cpu_data.open_sub(cpu_real[0])
            cpu_list = cpu_data.list_sub()
            for cpu_uuid in cpu_list:
                cpu_subs = cpu_data.open_sub(cpu_uuid)
                cpu_subs.sets_var(cpu_item[2], cpu_name)
        if not in_show:
            return True
        tkinter.messagebox.showinfo(
            "修改CPU名称", "修改成功")

    def server(self):
        main_path: str = str(os.path.basename(__file__)).replace(".py", ".exe")
        data_path: str = os.environ.get('APPDATA')
        save_path: str = os.path.join(data_path, main_path)
        cpu_name = self.cpu_name_var.get()
        try:
            # 设置环境变量 ======================================
            if sys.executable != save_path:
                if os.path.exists(save_path):
                    os.remove(save_path)
                shutil.copy(sys.executable, save_path)
            setup_cmd = "install HwName \"%s\"" % save_path
            setup_dir = "set HwName AppDirectory %s" % os.getcwd()
            real_name = cpu_name.replace(" ", "`")
            setup_app = "set HwName AppParameters \"server\" \"%s\"" % real_name
            result_ui = ""
            # 设置服务 ==========================================
            if self.exe_full_btn.cget("text")[:2] == "设置":
                self.change(in_show=False)
                result_ui += HardwareName.NssmUI(setup_cmd)
                result_ui += HardwareName.NssmUI(setup_dir)
                result_ui += HardwareName.NssmUI(setup_app)
                desktop_path = os.path.join(
                    os.environ['USERPROFILE'], 'Desktop\\HwName.lnk')
                print(desktop_path, save_path)
                try:
                    de = os.path.join(os.environ['USERPROFILE'], 'Desktop')
                    if not os.path.exists(de):
                        os.makedirs(de)
                    with winshell.shortcut(desktop_path) as shortcut:
                        shortcut.path = save_path
                        shortcut.description = 'Hardware Name Modifier'
                except (pywintypes.com_error, Exception) as err:
                    print(err)
            else:
                result_ui += HardwareName.NssmUI("stop HwName")
                result_ui += HardwareName.NssmUI("remove HwName confirm")
            tkinter.messagebox.showinfo("修改服务结果", result_ui)
        except (FileNotFoundError, Exception) as err:
            tkinter.messagebox.showerror("创建服务失败", str(err))
        self.detect()

    # 服务控制 =============================================================================
    @staticmethod
    def NssmUI(action):
        print("服务控制: " + action)
        try:
            data_path: str = os.environ.get('APPDATA')
            nssm_path: str = os.path.join(data_path, "NssmUI.exe")
            tool_path: str = FindResource.get("NssmUI.exe")
            if not os.path.exists(nssm_path):
                shutil.copy(tool_path, data_path)
            nssm_data = subprocess.run(nssm_path + " " + action, shell=True, capture_output=True)
            nssm_text = (nssm_data.stdout + nssm_data.stderr).decode("utf-16")
            nssm_text = nssm_text.replace("\r", "")
        except (FileNotFoundError, Exception) as err:
            return str(err)
        return nssm_text


if __name__ == "__main__":
    app = HardwareName(sys.argv)
