import sys
import os
import time
import ctypes
import psutil
import cv2
import numpy as np
import pydirectinput
import keyboard
from PIL import ImageGrab

os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "0"
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import QApplication, QVBoxLayout, QWidget, QHBoxLayout
from qfluentwidgets import (FluentWindow, PushButton, TextEdit,
                            StrongBodyLabel, CheckBox, InfoBar, CardWidget, LineEdit,
                            FluentIcon)



def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def get_active_window_process_name():
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        if hwnd == 0: return None
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        process = psutil.Process(pid.value)
        return process.name()
    except Exception:
        return None


class AutomationWorker(QThread):
    log_signal = pyqtSignal(str)

    def __init__(self, target_exe="Endfield.exe"):
        super().__init__()
        self.target_exe = target_exe
        self.is_running = False
        self.enable_skip = False
        self.last_focus_status = True
        self.templates = {}
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.load_resources()

    def load_resources(self):
        files = {
            'skip': 'tpl_skip.png',
            'confirm': 'tpl_confirm.png'
        }
        for key, filename in files.items():
            full_path = os.path.join(self.base_dir, filename)
            if os.path.exists(full_path):
                img = cv2.imread(full_path)
                if img is not None:
                    self.templates[key] = {
                        'data': cv2.cvtColor(img, cv2.COLOR_BGR2GRAY),
                        'w': img.shape[1],
                        'h': img.shape[0]
                    }
                else:
                    self.templates[key] = None
                    self.log_signal.emit(f"Error: 图片损坏 {filename}")
            else:
                self.templates[key] = None

    def update_config(self, enable_skip, target_exe):
        self.enable_skip = enable_skip
        self.target_exe = target_exe

    def find_on_screen(self, screen_gray, template_key, threshold=0.8):
        temp = self.templates.get(template_key)
        if not temp:
            return None
        try:
            res = cv2.matchTemplate(screen_gray, temp['data'], cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)
            if max_val >= threshold:
                return (max_loc[0] + temp['w'] // 2, max_loc[1] + temp['h'] // 2)
        except Exception:
            pass
        return None

    def capture_screen(self):
        try:
            screenshot = ImageGrab.grab()
            img = np.array(screenshot)
            # PIL 返回 RGB，转灰度
            img_gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
            return img_gray
        except Exception as e:
            self.log_signal.emit(f"截图异常: {e}")
            return None

    def run(self):
        self.log_signal.emit(">>> 核心服务启动 (OpenCV模式)")
        self.is_running = True
        skip_cooldown = 0

        while self.is_running:
            try:
                # 1. 焦点检测
                current_process = get_active_window_process_name()
                if not current_process or (self.target_exe.lower() not in current_process.lower()):
                    if self.last_focus_status:
                        self.log_signal.emit(f"挂起: 等待游戏前台 ({current_process})")
                        self.last_focus_status = False
                    time.sleep(1.0)
                    continue

                if not self.last_focus_status:
                    self.log_signal.emit("恢复: 监测中...")
                    self.last_focus_status = True

                # 2. 截图
                screen_gray = self.capture_screen()
                if screen_gray is None:
                    time.sleep(0.2)
                    continue

                current_time = time.time()

                # 3. 剧情跳过
                if self.enable_skip and current_time - skip_cooldown > 1.0:
                    skip_pos = self.find_on_screen(screen_gray, 'skip', 0.8)
                    if skip_pos:
                        self.log_signal.emit("[剧情] 触发跳过")
                        pydirectinput.click(skip_pos[0], skip_pos[1])
                        time.sleep(0.3)

                        # 再次截图寻找确认按钮
                        sg2 = self.capture_screen()
                        if sg2 is not None:
                            confirm_pos = self.find_on_screen(sg2, 'confirm', 0.8)
                            if confirm_pos:
                                pydirectinput.click(confirm_pos[0], confirm_pos[1])
                                self.log_signal.emit("[剧情] 确认跳过")

                        skip_cooldown = current_time

                # PIL截图较慢，适当放宽间隔
                time.sleep(0.1)

            except Exception as e:
                self.log_signal.emit(f"循环异常: {e}")
                time.sleep(1)

        self.log_signal.emit(">>> 服务已停止")

    def stop(self):
        self.is_running = False
        self.wait()



class LogOverlay(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool |
            Qt.WindowType.WindowTransparentForInput
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.log_display = TextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setStyleSheet(
            "TextEdit { background-color: rgba(0,0,0,100); color: #00FFCC; border: none; "
            "font-size: 14px; font-weight: bold; font-family: 'Microsoft YaHei'; }"
        )
        self.log_display.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        layout.addWidget(self.log_display)
        self.resize(350, 120)
        self.move(20, 100)

    def append_log(self, text):
        current_text = self.log_display.toPlainText()
        if len(current_text) > 2000:
            self.log_display.clear()
        self.log_display.append(f"[{time.strftime('%M:%S')}] {text}")
        cursor = self.log_display.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.log_display.setTextCursor(cursor)


class HomePage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("homePage")

        self.vbox = QVBoxLayout(self)
        self.vbox.setContentsMargins(20, 20, 20, 20)

        # 进程名
        self.vbox.addWidget(StrongBodyLabel("游戏进程名 (部分匹配)"))
        self.exe_input = LineEdit()
        self.exe_input.setText("Endfield")
        self.vbox.addWidget(self.exe_input)

        # 功能卡片
        self.vbox.addWidget(StrongBodyLabel("功能配置"))
        card = CardWidget(self)
        card_layout = QVBoxLayout(card)

        self.chk_skip = CheckBox("剧情自动跳过 (Skip)")
        self.chk_skip.setChecked(True)

        # ★ 自动拾取：取消勾选 + 禁用 + 标注开发中
        self.chk_loot = CheckBox("大世界自动拾取 (F)  [开发中]")
        self.chk_loot.setChecked(False)
        self.chk_loot.setEnabled(False)

        card_layout.addWidget(self.chk_skip)
        card_layout.addWidget(self.chk_loot)
        self.vbox.addWidget(card)

        # 热键
        self.vbox.addWidget(StrongBodyLabel("控制热键"))
        hk_layout = QHBoxLayout()
        self.hk_input = LineEdit()
        self.hk_input.setText("F10")
        self.hk_btn = PushButton("更新热键")
        hk_layout.addWidget(self.hk_input)
        hk_layout.addWidget(self.hk_btn)
        self.vbox.addLayout(hk_layout)

        # 按钮
        btn_layout = QHBoxLayout()
        self.start_btn = PushButton("启动服务")
        self.overlay_btn = PushButton("显示/隐藏日志")
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.overlay_btn)
        self.vbox.addLayout(btn_layout)

        self.vbox.addStretch(1)


class MainWindow(FluentWindow):
    toggle_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("终末地助手")
        self.resize(500, 500)

        self.worker = AutomationWorker()

        self.home_page = HomePage(self)
        self.addSubInterface(self.home_page, FluentIcon.HOME, "主页")

        self.overlay = LogOverlay()
        self._current_hotkey = None

        # 信号
        self.toggle_signal.connect(self.toggle_start)
        self.worker.log_signal.connect(self.overlay.append_log)
        self.home_page.start_btn.clicked.connect(self.toggle_start)
        self.home_page.overlay_btn.clicked.connect(self.toggle_overlay)
        self.home_page.hk_btn.clicked.connect(self.register_hotkey)
        self.home_page.chk_skip.stateChanged.connect(self.sync_config)

        self.register_hotkey()

        screen = QApplication.primaryScreen().geometry()
        self.move((screen.width() - self.width()) // 2,
                  (screen.height() - self.height()) // 2)

    @property
    def hk_input(self):
        return self.home_page.hk_input

    @property
    def start_btn(self):
        return self.home_page.start_btn

    def sync_config(self):
        target = self.home_page.exe_input.text().strip()
        self.worker.update_config(
            self.home_page.chk_skip.isChecked(),
            target
        )

    def register_hotkey(self):
        key = self.hk_input.text().strip()
        if not key:
            return
        try:
            if self._current_hotkey is not None:
                try:
                    keyboard.remove_hotkey(self._current_hotkey)
                except (ValueError, KeyError, AttributeError):
                    pass
                self._current_hotkey = None

            self._current_hotkey = keyboard.add_hotkey(
                key,
                lambda: self.toggle_signal.emit(),
                suppress=False
            )
            InfoBar.success("热键就绪", f"按 [{key}] 开启/停止", parent=self)
            self.start_btn.setText(f"启动服务 ({key})")
        except Exception as e:
            InfoBar.error("热键错误", str(e), parent=self)

    def toggle_start(self):
        if not self.worker.isRunning():
            self.sync_config()
            self.worker.start()
            self.start_btn.setText(f"停止运行 ({self.hk_input.text()})")
            if not self.overlay.isVisible():
                self.overlay.show()
        else:
            self.worker.stop()
            self.start_btn.setText(f"启动服务 ({self.hk_input.text()})")
            self.overlay.append_log(">>> 服务已暂停")

    def toggle_overlay(self):
        if self.overlay.isVisible():
            self.overlay.hide()
        else:
            self.overlay.show()

    def closeEvent(self, event):
        self.worker.stop()
        if self._current_hotkey is not None:
            try:
                keyboard.remove_hotkey(self._current_hotkey)
            except (ValueError, KeyError, AttributeError):
                pass
        try:
            keyboard.unhook_all()
        except (AttributeError, Exception):
            pass
        self.overlay.close()
        super().closeEvent(event)


# ==========================================
# 主入口
# ==========================================
def main():
    if not is_admin():
        print(">>> 请求管理员权限...")
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, __file__, None, 1)
        except Exception as e:
            print(f"提权失败: {e}")
        finally:
            sys.exit(0)

    try:
        if hasattr(Qt.HighDpiScaleFactorRoundingPolicy, 'PassThrough'):
            QApplication.setHighDpiScaleFactorRoundingPolicy(
                Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except:
        pass

    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()