import json
import os
import shutil
import subprocess
import sys
import time
import tkinter
import webbrowser
import winreg

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
CONFIG_GPU = [
    "wmic path win32_VideoController get PNPDeviceID",
    "SYSTEM\\CurrentControlSet\\Enum\\",
    "DISM /Online /Add-Capability /CapabilityName:WMIC~~~~",
    "没有检测到GPU，可能缺少WMIC\n是否立即联网安装?"
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
        self.gpus = {
            # "pci_path": {"old", "new"}
        }
        self.cpus = {"old": "", "new": ""}
        self.root = tk.Tk()
        self.head = ttk.Style()
        self.root.geometry("340x190")
        self.root.title("CPU名称修改工具 for Windows")
        # 字体设置 ------------------------------------------------------------
        self.font = tkinter.font.Font(family="MapleMono SC NF", size=20)
        self.head.configure("TNotebook.Tab", font=self.font)
        self.head.configure("TFrame", font=self.font)
        self.head.configure("TLabel", font=self.font)
        # 界面配置 ============================================================
        self.gpu_last_var = ""
        self.cpu_name_var = tk.StringVar()
        self.gpu_name_var = tk.StringVar()
        self.gpu_list_var = tk.StringVar()
        self.cpu_name_tab = ttk.Label(self.root, text="CPU名称")
        self.cpu_name_con = ttk.Entry(self.root, width=35)
        self.gpu_list_tab = ttk.Label(self.root, text="GPU选择")
        self.gpu_list_con = ttk.Combobox(self.root, width=33)
        self.gpu_name_tab = ttk.Label(self.root, text="GPU名称")
        self.gpu_name_con = ttk.Entry(self.root, width=35)
        self.exe_once_btn = ttk.Button(self.root, text="本次临时修改")
        self.exe_full_btn = ttk.Button(self.root, text="设置开机自启")
        self.exe_info_btn = ttk.Button(self.root, text="🐱 Github")
        self.cpu_name_tab.grid(column=0, row=0, pady=10, padx=5, sticky="nsew")
        self.cpu_name_con.grid(column=1, row=0, pady=10, padx=5, columnspan=5)
        self.gpu_list_tab.grid(column=0, row=1, pady=10, padx=5, sticky="nsew")
        self.gpu_list_con.grid(column=1, row=1, pady=10, padx=5, columnspan=5)
        self.gpu_name_tab.grid(column=0, row=2, pady=10, padx=5, sticky="nsew")
        self.gpu_name_con.grid(column=1, row=2, pady=10, padx=5, columnspan=5)
        self.exe_once_btn.grid(column=0, row=3, padx=5, columnspan=2, sticky="nsew")
        self.exe_full_btn.grid(column=2, row=3, padx=5, columnspan=2, sticky="nsew")
        self.exe_info_btn.grid(column=4, row=3, padx=5, columnspan=2, sticky="nsew")
        self.cpu_name_con.config(textvariable=self.cpu_name_var)
        self.gpu_name_con.config(textvariable=self.gpu_name_var)
        self.gpu_list_con.config(textvariable=self.gpu_list_var)
        self.gpu_list_con.config(state="readonly")
        self.gpu_list_con.bind('<<ComboboxSelected>>', self.select)
        self.exe_once_btn.config(bootstyle="success")
        self.exe_full_btn.config(bootstyle="info")
        self.exe_info_btn.config(bootstyle="dark")
        self.exe_once_btn.config(command=self.change)
        self.exe_full_btn.config(command=self.server)
        self.exe_info_btn.config(command=self.github)
        self.loader()
        self.detect()
        if len(in_args) >= 2 and in_args[1] == "server":
            if len(in_args) >= 3:
                self.change(in_args[2])
            else:
                self.change()
            exit(0)
        self.root.mainloop()

    # 查找显卡 ###################################################
    def github(self, event=None):
        webbrowser.open("https://github.com/PIKACHUIM/HardwareNameModifier")

    # 查找显卡 ###################################################
    def finder(self, in_name):
        for gpu_path in self.gpus:
            if in_name == self.gpus[gpu_path]['old']:
                return gpu_path
        return None

    # 选中数据 ###################################################
    def select(self, event=None):
        # 保存之前的选项 ===================================
        set_name = self.gpu_name_var.get()
        if self.gpu_last_var != "" and set_name != "":
            self.gpus[self.gpu_last_var]["new"] = set_name
        # 设置当前的选项 ===================================
        old_name = self.gpu_list_var.get()
        gpu_path = self.finder(old_name)
        self.gpu_name_var.set(old_name)
        if gpu_path is not None:
            self.gpu_last_var = gpu_path
            self.gpu_name_var.set(self.gpus[gpu_path]["new"])

    def detect(self):
        # CPU =============================================
        cpu_data = WinReg(CONFIG_CPU[0][0])
        cpu_data = cpu_data.open_sub(CONFIG_CPU[0][1])
        cpu_last = cpu_data.open_sub("0")
        cpu_name = cpu_last.read_var(CONFIG_CPU[0][2])[0]
        self.cpu_name_var.set(cpu_name)
        if self.cpus['old'] != "":
            self.cpu_name_var.set(cpu_name)
        self.cpus['new'] = self.cpu_name_var.get()
        self.cpus['old'] = self.cpu_name_var.get()
        # GPU =============================================
        gpu_proc = subprocess.run(
            CONFIG_GPU[0],
            shell=True, capture_output=True, text=True)
        gpu_last = gpu_proc.stdout.replace("\n\n", "\n")
        gpu_last = gpu_last.replace(" ", "").splitlines()[1:]
        gpu_last = [i for i in gpu_last if len(i) > 0]
        gpu_list = []
        if len(gpu_last) == 0:
            answer = tkinter.messagebox.askquestion(
                "安装WMIC", CONFIG_GPU[3])
            if answer == "yes":
                subprocess.run(CONFIG_GPU[2])
                tkinter.messagebox.showinfo(
                    "安装WMIC", "安装成功")
                os.startfile(sys.executable)
                exit(0)
        for gpu_path in gpu_last:
            gpu_data = WinReg(CONFIG_GPU[1] + gpu_path)
            gpu_name = gpu_data.read_var("DeviceDesc")
            gpu_name = gpu_name[0].split(";")[-1]
            if gpu_path not in self.gpus:
                self.gpus[gpu_path] = {}
                self.gpus[gpu_path]["old"] = gpu_name
            self.gpus[gpu_path]["new"] = gpu_name
            gpu_list.append(self.gpus[gpu_path]["old"])
        self.gpu_list_con.configure(values=gpu_list)
        # BTN =============================================
        self.export()
        srv_text = HardwareName.NssmUI("status HwName")
        if srv_text.find("Can't open service") == -1:
            self.exe_full_btn.config(text="取消开机自启")
            self.exe_full_btn.config(bootstyle="danger")
        else:
            self.exe_full_btn.config(text="设置开机自启")
            self.exe_full_btn.config(bootstyle="info")

    # 保存数据 =======================================================
    def export(self, in_path="HwName.json"):
        out_data = {"cpu": self.cpus, "gpu": self.gpus}
        with open(in_path, "w", encoding="utf-8") as out_file:
            out_file.write(json.dumps(out_data, ensure_ascii=True))

    # 保存数据 =======================================================
    def loader(self, in_path="HwName.json"):
        if not os.path.exists(in_path):
            return
        with open(in_path, "r", encoding="utf-8") as out_file:
            out_data = json.loads(out_file.read())
            self.cpus = out_data["cpu"]
            self.gpus = out_data["gpu"]

    # 执行修改 =======================================================
    def change(self, in_gpu_name=None, in_show=True):
        self.select()
        cpu_name = self.cpu_name_var.get()
        # 修改CPU ====================================================
        self.cpus['new'] = cpu_name
        if in_gpu_name is not None:
            cpu_name = in_gpu_name.replace("`", " ")
        for cpu_item in CONFIG_CPU:
            cpu_data = WinReg(cpu_item[0])
            cpu_real = cpu_data.find_sub(cpu_item[1])
            cpu_data = cpu_data.open_sub(cpu_real[0])
            cpu_list = cpu_data.list_sub()
            for cpu_uuid in cpu_list:
                cpu_subs = cpu_data.open_sub(cpu_uuid)
                cpu_subs.sets_var(cpu_item[2], cpu_name)
        # 修改GPU ====================================================
        for gpu_path in self.gpus:
            new_name = self.gpus[gpu_path]["new"]
            # if new_name == self.gpus[gpu_path]["old"]:
            #     continue
            gpu_data = WinReg(CONFIG_GPU[1] + gpu_path,
                              in_type=winreg.KEY_WRITE)
            gpu_data.sets_var("DeviceDesc", new_name)
        # 输出 =======================================================
        self.export()
        if in_show:
            tkinter.messagebox.showinfo("修改结果", "修改成功")

    # 设置服务 =======================================================
    def server(self):
        main_path: str = str(os.path.basename(__file__)).replace(".py", ".exe")
        data_path: str = os.environ.get('APPDATA')
        save_path: str = os.path.join(data_path, main_path)
        conf_path: str = os.path.join(data_path, main_path)
        # sets_name: str = self.cpu_name_var.get()
        try:
            # 设置环境变量 ======================================
            if self.exe_full_btn.cget("text")[:2] == "设置":
                if os.path.exists(conf_path):
                    os.remove(conf_path)
                shutil.copy("HwName.json", data_path)
                if sys.executable != save_path:
                    if os.path.exists(save_path):
                        os.remove(save_path)
                    shutil.copy(sys.executable, save_path)
            setup_cmd = "install HwName \"%s\"" % save_path
            # setup_dir = "set HwName AppDirectory %s" % os.getcwd()
            setup_dir = "set HwName AppDirectory %s" % data_path
            # real_name = sets_name.replace(" ", "`")
            # setup_app = "set HwName AppParameters \"server\" \"%s\"" % real_name
            setup_app = "set HwName AppParameters \"server\""
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
