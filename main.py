import ctypes
import os


def enable_drag_drop():
    """启用控制台窗口的拖放功能"""
    try:
        import win32console
        import win32gui
        hwnd = win32console.GetConsoleWindow()
        import win32api
        win32gui.DragAcceptFiles(hwnd, True)
    except ImportError:
        pass


def get_dropped_path(prompt_text):
    """获取拖放的文件或文件夹路径"""
    print(f"\n{prompt_text}")
    print("请将文件或文件夹拖放到此窗口,然后按回车...")

    while True:
        user_input = input().strip()

        # 移除可能的引号
        user_input = user_input.strip('"').strip("'")

        if os.path.exists(user_input):
            return user_input
        else:
            print(f"路径无效: {user_input}")
            print("请重新拖放文件或文件夹:")


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
        print(f"创建 desktop.ini 时出错: {e}")
        return False


def refresh_explorer():
    """刷新资源管理器以显示新图标"""
    try:
        # 通知系统图标已更改
        SHChangeNotify = ctypes.windll.shell32.SHChangeNotifyW
        SHChangeNotify(0x08000000, 0x0000, None, None)  # SHCNE_ASSOCCHANGED
    except Exception as e:
        print(f"刷新资源管理器时出错: {e}")


def main():
    print("=" * 60)
    print("      Windows 文件夹图标设置工具")
    print("=" * 60)
    print("\n此工具可以为指定文件夹设置自定义图标")
    print("支持的图标格式: .ico, .exe, .dll")

    # 获取文件夹路径
    folder_path = get_dropped_path("步骤 1/2: 选择目标文件夹")

    if not os.path.isdir(folder_path):
        print(f"\n错误: '{folder_path}' 不是一个有效的文件夹!")
        input("\n按回车键退出...")
        return

    print(f"✓ 已选择文件夹: {folder_path}")

    # 获取图标路径
    icon_path = get_dropped_path("\n步骤 2/2: 选择图标文件")

    # 检查图标文件格式
    icon_ext = os.path.splitext(icon_path)[1].lower()
    if icon_ext not in ['.ico', '.exe', '.dll']:
        print(f"\n警告: '{icon_ext}' 可能不是有效的图标文件格式")
        print("建议使用 .ico, .exe 或 .dll 文件")
        choice = input("是否继续? (y/n): ").strip().lower()
        if choice != 'y':
            print("操作已取消")
            input("\n按回车键退出...")
            return

    if not os.path.isfile(icon_path):
        print(f"\n错误: '{icon_path}' 不是一个有效的文件!")
        input("\n按回车键退出...")
        return

    print(f"✓ 已选择图标: {icon_path}")

    # 设置文件夹图标
    print("\n正在设置文件夹图标...")
    if create_desktop_ini(folder_path, icon_path):
        print("✓ desktop.ini 创建成功")

        # 刷新资源管理器
        print("正在刷新资源管理器...")
        refresh_explorer()

        print("\n" + "=" * 60)
        print("✓ 文件夹图标设置成功!")
        print("=" * 60)
        print("\n提示:")
        print("1. 如果图标未立即显示,请按 F5 刷新文件夹")
        print("2. 或者关闭并重新打开文件夹窗口")
        print("3. 图标文件不要移动或删除,否则图标会失效")
    else:
        print("\n× 设置失败! 请检查是否有足够的权限")

    input("\n按回车键退出...")


if __name__ == "__main__":
    try:
        enable_drag_drop()
        main()
    except KeyboardInterrupt:
        print("\n\n程序已被用户中断")
    except Exception as e:
        print(f"\n发生错误: {e}")
        input("\n按回车键退出...")