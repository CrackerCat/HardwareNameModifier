import ctypes
import winreg


class WinReg:
    def __init__(self,
                 in_path: str = None,
                 in_type: int = winreg.KEY_READ,
                 in_area: int = winreg.HKEY_LOCAL_MACHINE):
        self.path: str = in_path
        self.area: int = in_area
        self.type: int = in_type
        self.data = None

    # 打开注册表 ####################################################
    def open_reg(self):
        if self.path is not None and len(self.path) > 0:
            self.data = winreg.OpenKey(
                self.area, self.path,0,
                ctypes.c_int32(self.type).value)
            return True
        return False

    # 检查注册表 ####################################################
    def load_reg(self):
        # 检查 ======================================================
        if self.data is None and not self.open_reg():  # 检查打开情况
            raise OSError("Can't open registry path: %s" % self.path)

    # 列出子项目 ####################################################
    def list_sub(self):
        self.load_reg()
        # 获取数据 ==================================================
        sub_nums, var_nums, set_last = winreg.QueryInfoKey(self.data)
        return [winreg.EnumKey(self.data, i) for i in range(sub_nums)]

    # 列出键值对 ####################################################
    def list_var(self):
        self.load_reg()
        # 获取数据 ==================================================
        sub_numbers, var_i, set_last = winreg.QueryInfoKey(self.data)
        return [winreg.EnumValue(self.data, i) for i in range(var_i)]

    # 查找键值对 ####################################################
    def find_var(self, in_name):
        var_list = self.list_var()
        return [i for i in var_list if i[0].find(in_name) != -1]

    # 查找键值对 ####################################################
    def find_sub(self, in_name):
        var_list = self.list_sub()
        return [i for i in var_list if i.find(in_name) != -1]

    # 读取键值对 ####################################################
    def read_var(self, in_name):
        self.load_reg()
        # 获取数据 ==================================================
        value, regtype = winreg.QueryValueEx(self.data, in_name)
        return value, regtype

    # 设置键值对 ####################################################
    def sets_var(self, in_name, in_data, in_type=winreg.REG_SZ):
        self.load_reg()
        try:
            winreg.SetValueEx(self.data, in_name, 0, in_type, in_data)
        except winreg.error as e:
            print("Can't set value: %s" % in_name, "due to", e)
            return False
        return True

    # 删除键值对 ####################################################
    def kill_var(self, in_name):
        self.load_reg()
        winreg.DeleteValue(self.data, in_name)
        return True

    # 获取子项目 ####################################################
    def open_sub(self, in_name):
        if in_name in self.list_sub():
            return WinReg(self.path + "\\" + in_name, self.area)
        raise OSError("Can't open registry path: %s" % in_name)
