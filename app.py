import tkinter
import ttkbootstrap as ttk
from ttkbootstrap import *

class HardwareName:
    def __init__(self):
        # 创建窗口 ------------------------------------------------------------
        self.area = None
        self.text = None
        self.view = None
        self.root = tk.Tk()
        self.head = ttk.Style()
        # 字体设置 ------------------------------------------------------------
        self.font = tkinter.font.Font(family="MapleMono SC NF", size=12)
        # 界面配置 ============================================================
        self.root.mainloop()



if __name__ == "__main__":
    app = HardwareName()