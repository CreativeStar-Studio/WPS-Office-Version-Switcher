import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QComboBox, QVBoxLayout, QWidget
from PyQt5.QtWidgets import QLabel, QMessageBox, QPushButton, QCheckBox, QSystemTrayIcon
from PyQt5.QtGui import QIcon
import json
import os
import winreg
import subprocess
import datetime
import time

args = sys.argv[1:]
print(args)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("WPS Office 版本切换器") 
        self.setGeometry(100, 100, 500, 250)  # x, y, width, height
        menubar = self.menuBar()
        fileMenu = menubar.addMenu("文件")
        
        exitAction = fileMenu.addAction("退出")
        exitAction.triggered.connect(self.close)
        
        helpAction = fileMenu.addAction("帮助")
        helpAction.setEnabled(False)  # Disable the "帮助" action
        helpAction.triggered.connect(self.showHelp)

        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)  # Add margins

        label = QLabel("选择版本")
        layout.addWidget(label)

        # 将 QComboBox 设置为可编辑，并设置只读及提示文字
        self.versionSelector = QComboBox()
        self.versionSelector.setEditable(True)
        self.versionSelector.lineEdit().setReadOnly(True)
        self.versionSelector.lineEdit().setPlaceholderText("选择版本")
        layout.addWidget(self.versionSelector)
        layout.addStretch()  # Add stretch to push content downwards

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        # 添加确定按钮
        confirmButton = QPushButton("确定")
        confirmButton.clicked.connect(self.onConfirm)
        layout.addWidget(confirmButton)
        # 获取版本列表和更新下拉框
        wps_dir = self.getVersionList()
        versions = []
        if wps_dir:
            for folder in os.listdir(wps_dir):
                if os.path.isdir(os.path.join(wps_dir, folder)) and any(char.isdigit() or char == '.' for char in folder):
                    versions.append(folder)
        print(versions)
        self.versionSelector.addItems(versions)
        self.versionSelector.setCurrentIndex(-1)  # 取消默认选中，显示placeholder
        self.versionSelector.setEditText('')        # 清空编辑框，让提示文字显示

        # 添加复选框到下拉框下面
        self.disableLoginCheckbox = QCheckBox("禁用登录")
        layout.addWidget(self.disableLoginCheckbox)

        # 添加伸缩项推送其它内容
        layout.addStretch()  # Add stretch to push content downwards

    def showHelp(self):
        print("This is the help information.")
    def closeEvent(self, event):
        reply = QMessageBox.question(self, '退出', '确定要退出吗？',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
    def getVersionList(self):
        settings_path = "./settings.json"
        if os.path.exists(settings_path):
            with open(settings_path, "r", encoding="utf-8") as file:
                try:
                    data = json.load(file)
                    data = data.get("dir", "")
                    return data
                except json.JSONDecodeError:
                    return "D:\\WPS Office"
        else:
            return "D:\\WPS Office"
        
    def onConfirm(self):
        print("确定按钮被点击")
        selected_version = self.versionSelector.currentText()
        if selected_version:
            print(f"选择的版本是: {selected_version}")
        else:
            print("未选择任何版本")
            QMessageBox.warning(self, "警告", "请选择一个版本")
            return
        if self.disableLoginCheckbox.isChecked():
            print("禁用登录已选中")
            try:
                key_path = r"Software\kingsoft\Office\6.0\plugins\officespace\flogin"
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
                winreg.SetValueEx(key, "enableForceLoginForFirstInstallDevice", 0, winreg.REG_SZ, "false")
                winreg.CloseKey(key)
                print("注册表已成功更新。")
            except Exception as e:
                print(f"更新注册表时出错: {e}")
        else:
            print("禁用登录未选中")
            try:
                key_path = r"Software\kingsoft\Office\6.0\plugins\officespace\flogin"
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
                winreg.SetValueEx(key, "enableForceLoginForFirstInstallDevice", 0, winreg.REG_SZ, "true")
                winreg.CloseKey(key)
                print("注册表已成功更新。")
            except Exception as e:
                print(f"更新注册表时出错: {e}")

        # 用args变量全部内容通过subprocess启动版本目录下的/wps.exe
        global wps_dir
        wps_dir = self.getVersionList()
        executable_path = os.path.join(wps_dir, selected_version, "office6", "wps.exe")
        try:
            subprocess.Popen([executable_path] + args)
            print(f"成功启动：{executable_path}")
        except Exception as e:
            print(f"启动失败：{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # 检查命令行参数是否包含 "/wpsupdate", "-from=task", "-source=task"
    update_args = {"/wpsupdate", "-from=task", "-source=task"}
    if any(arg in update_args for arg in args):
        tray = QSystemTrayIcon()
        tray.setIcon(app.style().standardIcon(QApplication.style().SP_ComputerIcon))
        tray.show()
        tray.showMessage("本启动器不支持升级WPS", "敬请谅解", QSystemTrayIcon.Information, 5000)
        tray.hide()
        sys.exit()
    elif any(arg in {"/run_app"} for arg in args):
        print("运行应用程序")
        tray = QSystemTrayIcon()
        tray.setIcon(app.style().standardIcon(QApplication.style().SP_ComputerIcon))
        tray.show()
        tray.showMessage("本启动器不支持启动小程序", "敬请谅解", QSystemTrayIcon.Information, 5000)
        tray.hide()
        sys.exit()
    path = "./settings.json"
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as file:
            try:
                data = json.load(file)
                data = data.get("debug", False)
                print(data)
                print(f"Debug mode is {'enabled' if data else 'disabled'}")
            except json.JSONDecodeError as e:
                print(f"JSON decode error: {e}")
                data = False
    else:
        data = False
    if data:
        print("Debug mode is enabled")
        # 通过临时 MainWindow 实例来获取版本路径
        tempMain = MainWindow()
        current_wps_dir = tempMain.getVersionList()
        with open("./debug.log", "w", encoding="utf-8") as file:
            file.write(f"{datetime.datetime.now()} - Args: {str(args)}.\n")
            file.write(f"{datetime.datetime.now()} - Versions: {str(current_wps_dir)}.\n")
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())