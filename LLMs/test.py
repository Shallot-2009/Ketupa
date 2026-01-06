import win32com.client
import os


def start_hfss():
    """启动 HFSS 并返回应用程序对象"""
    try:
        oAnsoftApp = win32com.client.Dispatch("AnsoftHfss.HfssScriptInterface")
        oDesktop = oAnsoftApp.GetAppDesktop()
        oDesktop.RestoreWindow()
        print("HFSS 已成功启动")
        return oAnsoftApp, oDesktop
    except Exception as e:
        print(f"启动 HFSS 失败: {e}")
        return None, None


def open_script(oDesktop, script_path):
    """在 HFSS 中打开并运行脚本"""
    if not os.path.exists(script_path):
        print(f"错误: 脚本文件 {script_path} 不存在")
        return False

    try:
        oProject = oDesktop.GetActiveProject()
        oDesign = oProject.GetActiveDesign()
        oDesign.RunScript(script_path)
        print(f"脚本 {script_path} 已成功运行")
        return True
    except Exception as e:
        print(f"运行脚本失败: {e}")
        return False


def main():
    # 启动 HFSS
    oAnsoftApp, oDesktop = start_hfss()
    if not oDesktop:
        return

    # 指定您的脚本路径
    script_path = r"E:\\00_Asenjo\\00_Project\\Ketupa\\main.py"  # 替换为您的实际脚本路径

    # 运行脚本
    open_script(oDesktop, script_path)


if __name__ == "__main__":
    main()