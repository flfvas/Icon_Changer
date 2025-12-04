import ctypes
import os
import tkinter as tk
from tkinter import filedialog, messagebox


def create_desktop_ini(folder_path, icon_path):
    """创建 desktop.ini 文件来设置文件夹图标"""
    desktop_ini_path = os.path.join(folder_path, 'desktop.ini')

    # desktop.ini 内容
    ini_content = f"""[.ShellClassInfo]
IconResource={icon_path},0
[ViewState]
Mode=
Vid=
FolderType=Generic
"""

    try:
        # 写入 desktop.ini
        with open(desktop_ini_path, 'w', encoding='utf-8') as f:
            f.write(ini_content)

        # 设置 desktop.ini 为隐藏和系统文件
        ctypes.windll.kernel32.SetFileAttributesW(
            desktop_ini_path,
            0x02 | 0x04  # FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM
        )

        # 设置文件夹为只读属性(这样 Windows 才会读取 desktop.ini)
        ctypes.windll.kernel32.SetFileAttributesW(
            folder_path,
            0x01 | 0x10  # FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_DIRECTORY
        )

        return True
    except Exception as e:
        return False, str(e)


def refresh_explorer():
    """刷新资源管理器以显示新图标"""
    try:
        # 通知系统图标已更改
        SHChangeNotify = ctypes.windll.shell32.SHChangeNotifyW
        SHChangeNotify(0x08000000, 0x0000, None, None)  # SHCNE_ASSOCCHANGED
    except Exception as e:
        pass


class FolderIconSetterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("文件夹图标设置工具")
        self.root.geometry("500x400")
        self.root.resizable(False, False)

        self.folder_path = ""
        self.icon_path = ""

        self.create_widgets()

    def create_widgets(self):
        # 标题
        title_label = tk.Label(
            self.root,
            text="Windows 文件夹图标设置工具",
            font=("Arial", 16, "bold"),
            pady=20
        )
        title_label.pack()

        # 文件夹选择区域
        folder_frame = tk.Frame(self.root, pady=10)
        folder_frame.pack(fill='x', padx=20)

        tk.Label(
            folder_frame,
            text="步骤 1: 选择目标文件夹",
            font=("Arial", 11, "bold")
        ).pack(anchor='w')

        self.folder_label = tk.Label(
            folder_frame,
            text="未选择文件夹",
            fg="gray",
            wraplength=450,
            justify='left'
        )
        self.folder_label.pack(anchor='w', pady=5)

        tk.Button(
            folder_frame,
            text="浏览文件夹...",
            command=self.select_folder,
            width=20,
            height=2
        ).pack(pady=5)

        # 分隔线
        tk.Frame(self.root, height=2, bg="lightgray").pack(fill='x', padx=20, pady=10)

        # 图标选择区域
        icon_frame = tk.Frame(self.root, pady=10)
        icon_frame.pack(fill='x', padx=20)

        tk.Label(
            icon_frame,
            text="步骤 2: 选择图标文件 (.ico, .exe, .dll)",
            font=("Arial", 11, "bold")
        ).pack(anchor='w')

        self.icon_label = tk.Label(
            icon_frame,
            text="未选择图标文件",
            fg="gray",
            wraplength=450,
            justify='left'
        )
        self.icon_label.pack(anchor='w', pady=5)

        tk.Button(
            icon_frame,
            text="浏览图标文件...",
            command=self.select_icon,
            width=20,
            height=2
        ).pack(pady=5)

        # 分隔线
        tk.Frame(self.root, height=2, bg="lightgray").pack(fill='x', padx=20, pady=10)

        # 执行按钮
        self.apply_button = tk.Button(
            self.root,
            text="设置文件夹图标",
            command=self.apply_icon,
            width=25,
            height=2,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 11, "bold"),
            state='disabled'
        )
        self.apply_button.pack(pady=20)

        # 状态栏
        self.status_label = tk.Label(
            self.root,
            text="请选择文件夹和图标文件",
            fg="blue",
            font=("Arial", 9)
        )
        self.status_label.pack(side='bottom', pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory(title="选择目标文件夹")
        if folder:
            self.folder_path = folder
            self.folder_label.config(
                text=f"✓ {folder}",
                fg="green"
            )
            self.check_ready()

    def select_icon(self):
        icon = filedialog.askopenfilename(
            title="选择图标文件",
            filetypes=[
                ("图标文件", "*.ico"),
                ("可执行文件", "*.exe"),
                ("动态链接库", "*.dll"),
                ("所有文件", "*.*")
            ]
        )
        if icon:
            self.icon_path = icon
            self.icon_label.config(
                text=f"✓ {icon}",
                fg="green"
            )
            self.check_ready()

    def check_ready(self):
        if self.folder_path and self.icon_path:
            self.apply_button.config(state='normal')
            self.status_label.config(
                text="准备就绪,点击按钮开始设置",
                fg="green"
            )
        else:
            self.apply_button.config(state='disabled')

    def apply_icon(self):
        if not self.folder_path or not self.icon_path:
            messagebox.showerror("错误", "请先选择文件夹和图标文件!")
            return

        # 确认操作
        confirm = messagebox.askyesno(
            "确认操作",
            f"将为以下文件夹设置图标:\n\n"
            f"文件夹: {self.folder_path}\n"
            f"图标: {self.icon_path}\n\n"
            f"是否继续?"
        )

        if not confirm:
            return

        # 执行设置
        self.status_label.config(text="正在设置图标...", fg="orange")
        self.root.update()

        result = create_desktop_ini(self.folder_path, self.icon_path)

        if result is True:
            refresh_explorer()
            messagebox.showinfo(
                "成功",
                "文件夹图标设置成功!\n\n"
                "提示:\n"
                "1. 如果图标未立即显示,请按 F5 刷新\n"
                "2. 或者关闭并重新打开文件夹\n"
                "3. 不要移动或删除图标文件"
            )
            self.status_label.config(text="设置成功!", fg="green")
        else:
            error_msg = result[1] if isinstance(result, tuple) else "未知错误"
            messagebox.showerror(
                "失败",
                f"设置失败!\n\n错误信息: {error_msg}\n\n"
                "请检查是否有足够的权限"
            )
            self.status_label.config(text="设置失败", fg="red")


def main():
    root = tk.Tk()
    app = FolderIconSetterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"程序错误: {e}")
        input("按回车键退出...")