import sys
import os
import json
import subprocess
import shutil
import win32com.client
import psutil
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QMessageBox,
    QSpinBox,
    QCheckBox,
    QTabWidget,
    QGroupBox,
    QGridLayout,
    QMenu,
    QInputDialog,
    QComboBox,
    QDialog,
)
from PyQt6.QtCore import Qt, QTimer
import threading
import time
import keyboard
import mouse
import win32gui
import win32api
import win32con
import win32process
import ctypes
from ctypes import wintypes
import math
from PIL import Image, ImageDraw, ImageFont
import random


def is_admin():
    # 检查是否具有管理员权限
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        print(e)
        return False


def check_admin_config():
    """检查管理员权限配置"""
    try:
        if os.path.exists("settings.json"):
            with open("settings.json", "r", encoding="utf-8") as f:
                config = json.load(f)
                return config.get("always_run_as_admin", False)
    except Exception as e:
        print(f"读取管理员配置失败: {str(e)}")
    return False


def save_admin_config():
    """保存管理员权限配置"""
    try:
        # 读取现有配置
        config = {}
        if os.path.exists("settings.json"):
            with open("settings.json", "r", encoding="utf-8") as f:
                config = json.load(f)

        # 更新管理员配置
        config["always_run_as_admin"] = True

        # 保存配置
        with open("settings.json", "w", encoding="utf-8") as f:
            json.dump(config, f)
    except Exception as e:
        print(f"保存管理员配置失败: {str(e)}")


def run_as_admin():
    # 以管理员权限重新运行程序
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )


class ChromeManager(QMainWindow):
    def __init__(self):
        super().__init__()
        if not is_admin():
            # 检查是否已经配置为始终以管理员权限运行
            if check_admin_config():
                run_as_admin()
                sys.exit()
                return

            # 如果没有配置，询问用户
            if (
                QMessageBox.question(
                    self,
                    "权限不足",
                    "需要管理员权限才能运行同步功能。\n是否以管理员身份重新启动程序?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                == QMessageBox.StandardButton.Yes
            ):
                # 保存用户的选择
                save_admin_config()
                run_as_admin()
                sys.exit()
                return
        self.setWindowTitle("Chrome多开管理器 V1.0    By:可达鸭    TG: @kedaya_798")
        self.setMinimumSize(1000, 600)

        # 初始化变量
        self.chrome_path = (
            r"C:\Program Files\Google\Chrome\Application\chrome.exe"  # Chrome路径
        )
        self.user_data_dir = "./ChromeProfiles"  # 用户数据目录
        self.browser_processes = {}  # 存储浏览器进程信息
        self.sync_timer = None  # 同步定时器
        self.sync_btn = None  # 初始化同步按钮变量

        # 添加同步功能相关变量
        self.is_syncing = False  # 是否正在同步
        self.master_window = None  # 主控窗口句柄
        self.sync_windows = []  # 同步窗口句柄列表
        self.hook_thread = None  # 钩子线程
        self.keyboard_hook = None  # 键盘钩子
        self.mouse_hook_id = None  # 鼠标钩子
        self.popup_monitor_thread = None  # 插件窗口监控线程
        self.popup_mappings = {}  # 插件窗口映射
        self.debug_ports = {}  # 调试端口映射

        # DWM API常量
        self.DWMWA_BORDER_COLOR = 34  # DWM窗口属性：边框颜色

        # 屏幕信息
        self.screens = []
        self.screen_var = None
        self.screen_combo = None

        # 鼠标移动优化参数
        self.last_mouse_position = (0, 0)
        self.last_move_time = 0
        self.mouse_threshold = 5  # 鼠标移动阈值
        self.move_interval = 0.05  # 移动间隔时间

        # 图标相关
        self.icon_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "icons"
        )
        self.profile_icons = {}  # 存储配置文件编号到图标路径的映射

        # 设置黑色主题
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1e1e1e;
            }
            QWidget {
                background-color: #1e1e1e;
                color: #ffffff;
            }
            QPushButton {
                background-color: #2d2d2d;
                border: 1px solid #3d3d3d;
                border-radius: 3px;
                padding: 5px 10px;
                color: #ffffff;
            }
            QPushButton:hover {
                background-color: #3d3d3d;
            }
            QPushButton:pressed {
                background-color: #4d4d4d;
            }
            QLineEdit {
                border: 1px solid #3d3d3d;
                border-radius: 3px;
                padding: 3px;
                background-color: #2d2d2d;
                color: #ffffff;
            }
            QTreeWidget {
                border: 1px solid #3d3d3d;
                border-radius: 3px;
                background-color: #2d2d2d;
                outline: none;
                color: #ffffff;
            }
            QTreeWidget::item {
                padding: 5px;
                border: none;
            }
            QTreeWidget::item:selected {
                background-color: transparent;
                color: #ffffff;
            }
            QTreeWidget::item:focus {
                background-color: transparent;
                color: #ffffff;
                border: none;
            }
            QTreeWidget::indicator {
                width: 16px;
                height: 16px;
            }
            QTreeWidget::indicator:unchecked {
                border: 2px solid #ffffff;
                background: transparent;
                border-radius: 3px;
            }
            QTreeWidget::indicator:checked {
                border: 2px solid #ffffff;
                background: #ffffff;
                border-radius: 3px;
                image: url(data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24'%3E%3Cpath fill='%232d2d2d' d='M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z'/%3E%3C/svg%3E);
            }
            QTabWidget::pane {
                border: 1px solid #3d3d3d;
                border-radius: 3px;
                background-color: #2d2d2d;
            }
            QTabBar::tab {
                background-color: #2d2d2d;
                border: 1px solid #3d3d3d;
                border-bottom: none;
                border-top-left-radius: 3px;
                border-top-right-radius: 3px;
                padding: 5px 10px;
                color: #ffffff;
            }
            QTabBar::tab:selected {
                background-color: #3d3d3d;
                border-bottom: none;
            }
            QGroupBox {
                border: 1px solid #3d3d3d;
                border-radius: 3px;
                margin-top: 10px;
                padding-top: 15px;
                color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #ffffff;
            }
            QSpinBox {
                border: 1px solid #3d3d3d;
                border-radius: 3px;
                padding: 3px;
                background-color: #2d2d2d;
                color: #ffffff;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                background-color: #3d3d3d;
            }
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {
                background-color: #4d4d4d;
            }
            QCheckBox {
                color: #ffffff;
            }
            QCheckBox::indicator {
                border: 1px solid #ffffff;
                background: transparent;
            }
            QCheckBox::indicator:checked {
                background: #ffffff;
            }
            QMenu {
                background-color: #2d2d2d;
                border: 1px solid #3d3d3d;
                color: #ffffff;
            }
            QMenu::item {
                padding: 5px 20px;
            }
            QMenu::item:selected {
                background-color: #3d3d3d;
            }
            QHeaderView::section {
                background-color: #2d2d2d;
                color: #ffffff;
                border: 1px solid #3d3d3d;
                padding: 5px;
            }
        """)

        # 创建图标目录
        if not os.path.exists(self.icon_dir):
            os.makedirs(self.icon_dir)

        # 创建UI
        self.create_ui()

        # 加载设置
        self.load_settings()

        # 连接信号
        self.connect_signals()

        # 更新屏幕列表
        self.update_screen_list()

        # 初始化刷新浏览器列表
        self.refresh_browser_list()

    # 添加图标生成和设置相关方法
    def generate_color_icon(self, number, size=48):
        """生成带数字的彩色图标"""

        try:
            # 随机但基于编号的颜色（确保相同编号总是有相同颜色）
            random.seed(number)
            r = random.randint(30, 220)  # 避免太亮或太暗
            g = random.randint(30, 220)
            b = random.randint(30, 220)

            # 创建图像
            img = Image.new("RGBA", (size, size), color=(0, 0, 0, 0))
            draw = ImageDraw.Draw(img)

            # 使用@chrome.png作为背景
            bg_image_path = "./chrome.png"
            if os.path.exists(bg_image_path):
                bg_image = Image.open(bg_image_path).resize((size, size))
                img.paste(bg_image, (0, 0))

            # 绘制横向扁圆背景
            ellipse_width = size * 0.85  # 椭圆宽度为图标尺寸的80%
            ellipse_height = size * 0.5  # 椭圆高度为图标尺寸的50%
            ellipse_left = (size - ellipse_width) / 2
            ellipse_top = (size - ellipse_height) / 2 + 12
            ellipse_right = ellipse_left + ellipse_width
            ellipse_bottom = ellipse_top + ellipse_height

            # 绘制扁椭圆背景
            draw.ellipse(
                (ellipse_left, ellipse_top, ellipse_right, ellipse_bottom),
                fill=(r, g, b, 255),
            )

            # 添加数字
            try:
                # 尝试加载系统字体
                font_size = 24  # 调整字体大小
                font_path = os.path.join(os.environ["WINDIR"], "Fonts", "Arial.ttf")
                if os.path.exists(font_path):
                    font = ImageFont.truetype(font_path, font_size)
                else:
                    # 尝试直接加载Arial字体
                    font = ImageFont.truetype("Arial", font_size)
            except Exception as font_error:
                print(f"加载字体失败: {str(font_error)}")
                # 使用默认字体
                font = ImageFont.load_default()
                font_size = 24  # 默认字体大小

            # 处理文本位置
            text = str(number)

            # 计算文本大小并居中放置
            try:
                # PIL 9.0.0 及以上版本
                if hasattr(font, "getbbox"):
                    bbox = font.getbbox(text)
                    text_width = bbox[2] - bbox[0]
                    text_height = bbox[3] - bbox[1]
                # 较老版本的PIL
                elif hasattr(draw, "textsize"):
                    text_width, text_height = draw.textsize(text, font=font)
                else:
                    # 最基本的估计
                    text_width = font_size * len(text) * 0.6
                    text_height = font_size

                # 计算位置使文本居中在椭圆内
                x = (size - text_width) / 2
                y = (size - text_height) / 2 + 10

                # 绘制文本，固定使用白色
                text_color = (255, 255, 255, 255)  # 白色文字
                draw.text((x, y), text, fill=text_color, font=font)
            except Exception as text_error:
                print(f"绘制文本失败: {str(text_error)}")
                # 简单位置处理
                draw.text(
                    (size // 4, size // 4), text, fill=(255, 255, 255, 255), font=font
                )

            # 保存到文件
            icon_path = os.path.join(self.icon_dir, f"chrome_icon_{number}.ico")

            # 确保图标目录存在
            os.makedirs(os.path.dirname(icon_path), exist_ok=True)

            # 保存为ICO文件
            try:
                img.save(icon_path, format="ICO")
            except Exception as save_error:
                print(f"保存图标失败: {str(save_error)}")
                # 尝试保存为PNG然后转换
                png_path = os.path.join(self.icon_dir, f"chrome_icon_{number}.png")
                img.save(png_path, format="PNG")
                icon_path = png_path

            return icon_path
        except Exception as e:
            print(f"生成图标失败: {str(e)}")
            return None

    def set_chrome_icon(self, hwnd, icon_path):
        """为Chrome窗口设置自定义图标"""
        try:
            # 加载图标
            big_icon = win32gui.LoadImage(
                0, icon_path, win32con.IMAGE_ICON, 32, 32, win32con.LR_LOADFROMFILE
            )
            small_icon = win32gui.LoadImage(
                0, icon_path, win32con.IMAGE_ICON, 16, 16, win32con.LR_LOADFROMFILE
            )

            # 设置窗口图标
            win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, big_icon)
            win32gui.SendMessage(
                hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, small_icon
            )

            return True
        except Exception as e:
            print(f"设置图标失败: {str(e)}")
            return False

    def apply_icons_to_chrome_windows(self):
        """为所有打开的Chrome窗口应用自定义图标"""
        for number, hwnd in self.browser_processes.items():
            # 检查是否已有图标
            if number not in self.profile_icons:
                icon_path = self.generate_color_icon(number)
                if icon_path:
                    self.profile_icons[number] = icon_path

            # 应用图标
            if number in self.profile_icons:
                self.set_chrome_icon(hwnd, self.profile_icons[number])

    def update_icons(self):
        """更新所有窗口的图标"""
        self.apply_icons_to_chrome_windows()

    def create_ui(self):
        # 创建主窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 创建顶部工具栏
        toolbar = QHBoxLayout()

        # 浏览器列表
        self.browser_list = QTreeWidget()
        self.browser_list.setHeaderLabels(["选择", "编号", "标题", "主控", "状态"])
        self.browser_list.setColumnWidth(0, 50)
        self.browser_list.setColumnWidth(1, 50)
        self.browser_list.setColumnWidth(2, 300)
        self.browser_list.setColumnWidth(3, 50)
        self.browser_list.setColumnWidth(4, 100)

        # 设置选择框
        self.browser_list.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)
        self.browser_list.setSelectionBehavior(QTreeWidget.SelectionBehavior.SelectRows)

        # 添加以下设置
        self.browser_list.setIndentation(0)  # 设置缩进为0
        self.browser_list.setAllColumnsShowFocus(True)  # 所有列都显示焦点
        self.browser_list.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )  # 启用自定义右键菜单

        self.browser_list.setStyleSheet("""
            QTreeWidget {
                border: 1px solid #3d3d3d;
                border-radius: 3px;
                background-color: #2d2d2d;
                outline: none;  /* 移除选中时的虚线框 */
                color: #ffffff;
            }
            QTreeWidget::item {
                padding: 5px;
                border: none;
            }
            QTreeWidget::item:selected {
                background-color: transparent;
                color: #ffffff;
            }
            QTreeWidget::item:focus {
                background-color: transparent;
                color: #ffffff;
                border: none;
            }
            QTreeWidget::indicator {
                width: 16px;
                height: 16px;
            }
            QTreeWidget::indicator:unchecked {
                border: 2px solid #ffffff;
                background: transparent;
                border-radius: 3px;
            }
            QTreeWidget::indicator:checked {
                border: 2px solid #ffffff;
                background: #ffffff;
                border-radius: 3px;
                image: url(data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24'%3E%3Cpath fill='%232d2d2d' d='M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z'/%3E%3C/svg%3E);
            }
            QHeaderView::section {
                background-color: #2d2d2d;
                color: #ffffff;
                border: 1px solid #3d3d3d;
                padding: 5px;
            }
        """)

        # 按钮
        self.refresh_btn = QPushButton("刷新列表")
        self.open_btn = QPushButton("打开选中")
        self.close_btn = QPushButton("关闭选中")
        self.copy_btn = QPushButton("复制选中")
        self.delete_btn = QPushButton("删除选中")
        self.sync_btn = QPushButton("▶ 开始同步")

        # 添加按钮到工具栏
        toolbar.addWidget(self.refresh_btn)
        toolbar.addWidget(self.open_btn)
        toolbar.addWidget(self.close_btn)
        toolbar.addWidget(self.copy_btn)
        toolbar.addWidget(self.delete_btn)
        toolbar.addWidget(self.sync_btn)

        # 添加屏幕选择下拉框
        screen_frame = QWidget()
        screen_layout = QHBoxLayout(screen_frame)
        screen_layout.addWidget(QLabel("屏幕:"))
        self.screen_combo = QComboBox()
        screen_layout.addWidget(self.screen_combo)
        toolbar.addWidget(screen_frame)

        # 添加自动排列和自定义排列按钮
        self.auto_arrange_btn = QPushButton("自动排列")
        toolbar.addWidget(self.auto_arrange_btn)

        # 添加自定义排列按钮
        self.popup_custom_arrange_btn = QPushButton("自定义排列")
        toolbar.addWidget(self.popup_custom_arrange_btn)

        toolbar.addStretch()

        # 创建标签页
        tab_widget = QTabWidget()

        # 环境设置标签页
        env_tab = QWidget()
        env_layout = QVBoxLayout(env_tab)

        # Chrome路径设置
        chrome_group = QGroupBox("Chrome路径设置")
        chrome_layout = QHBoxLayout()
        self.chrome_path_input = QLineEdit()
        self.chrome_path_input.setText(self.chrome_path)
        self.chrome_browse_btn = QPushButton("浏览")
        chrome_layout.addWidget(QLabel("Chrome路径:"))
        chrome_layout.addWidget(self.chrome_path_input)
        chrome_layout.addWidget(self.chrome_browse_btn)
        chrome_group.setLayout(chrome_layout)

        # 数据目录设置
        data_group = QGroupBox("数据目录设置")
        data_layout = QHBoxLayout()
        self.data_path = QLineEdit()
        self.data_path.setText("./Data")
        self.data_browse_btn = QPushButton("浏览")
        data_layout.addWidget(QLabel("数据保存目录:"))
        data_layout.addWidget(self.data_path)
        data_layout.addWidget(self.data_browse_btn)
        data_group.setLayout(data_layout)

        # 窗口数量设置
        count_group = QGroupBox("窗口数量设置")
        count_layout = QHBoxLayout()
        self.window_count = QSpinBox()
        self.window_count.setRange(1, 100)
        self.window_count.setValue(5)

        # 添加slow_mo设置
        slow_mo_layout = QHBoxLayout()
        self.slow_mo_input = QSpinBox()
        self.slow_mo_input.setRange(0, 5000)
        self.slow_mo_input.setValue(0)
        self.slow_mo_input.setSingleStep(100)
        slow_mo_layout.addWidget(QLabel("操作延迟(ms):"))
        slow_mo_layout.addWidget(self.slow_mo_input)

        self.create_btn = QPushButton("创建环境")
        count_layout.addWidget(QLabel("创建窗口数量:"))
        count_layout.addWidget(self.window_count)
        count_layout.addLayout(slow_mo_layout)
        count_layout.addWidget(self.create_btn)
        count_group.setLayout(count_layout)

        # 隐藏变量，用于自定义排列
        self.start_x = QLineEdit()
        self.start_x.setText("0")
        self.start_y = QLineEdit()
        self.start_y.setText("0")
        self.window_width = QLineEdit()
        self.window_width.setText("500")
        self.window_height = QLineEdit()
        self.window_height.setText("400")
        self.h_spacing = QLineEdit()
        self.h_spacing.setText("0")
        self.v_spacing = QLineEdit()
        self.v_spacing.setText("0")
        self.windows_per_row = QLineEdit()
        self.windows_per_row.setText("5")
        self.windows_per_column = QLineEdit()
        self.windows_per_column.setText("0")  # 0表示不限制
        self.auto_fill_screen = QCheckBox("自动铺满屏幕")
        self.auto_fill_screen.setChecked(True)

        # 添加组件到环境设置标签页
        env_layout.addWidget(chrome_group)
        env_layout.addWidget(data_group)
        env_layout.addWidget(count_group)
        env_layout.addStretch()

        # 批量操作标签页
        batch_tab = QWidget()
        batch_layout = QVBoxLayout(batch_tab)

        # URL设置
        url_group = QGroupBox("批量打开网页")
        url_layout = QGridLayout()

        self.url_input = QLineEdit()
        self.url_input.setText("www.google.com")

        self.new_page_checkbox = QCheckBox("在新窗口打开")
        self.new_page_checkbox.setChecked(True)

        self.open_url_btn = QPushButton("批量打开")

        url_layout.addWidget(QLabel("网址:"), 0, 0)
        url_layout.addWidget(self.url_input, 0, 1)
        url_layout.addWidget(self.new_page_checkbox, 0, 2)
        url_layout.addWidget(self.open_url_btn, 0, 3)
        url_group.setLayout(url_layout)

        # 批量点击设置
        click_group = QGroupBox("批量点击操作")
        click_layout = QGridLayout()

        self.selector_input = QLineEdit()
        self.selector_input.setPlaceholderText("输入CSS选择器，如: button.submit-btn")

        self.click_delay = QSpinBox()
        self.click_delay.setRange(0, 10000)
        self.click_delay.setValue(1000)
        self.click_delay.setSingleStep(100)

        self.click_btn = QPushButton("批量点击")

        click_layout.addWidget(QLabel("选择器:"), 0, 0)
        click_layout.addWidget(self.selector_input, 0, 1)
        click_layout.addWidget(QLabel("点击延迟(ms):"), 0, 2)
        click_layout.addWidget(self.click_delay, 0, 3)
        click_layout.addWidget(self.click_btn, 0, 4)
        click_group.setLayout(click_layout)

        # 批量输入设置
        input_group = QGroupBox("批量输入文本")
        input_layout = QGridLayout()

        self.input_selector = QLineEdit()
        self.input_selector.setPlaceholderText(
            "输入CSS选择器，如: input[name='username']"
        )

        self.input_text = QLineEdit()
        self.input_text.setPlaceholderText("要输入的文本")

        self.input_btn = QPushButton("批量输入")

        input_layout.addWidget(QLabel("选择器:"), 0, 0)
        input_layout.addWidget(self.input_selector, 0, 1)
        input_layout.addWidget(QLabel("文本:"), 1, 0)
        input_layout.addWidget(self.input_text, 1, 1)
        input_layout.addWidget(self.input_btn, 1, 2)
        input_group.setLayout(input_layout)

        # 添加组件到批量操作标签页
        batch_layout.addWidget(url_group)
        batch_layout.addWidget(click_group)
        batch_layout.addWidget(input_group)
        batch_layout.addStretch()
        batch_tab.setLayout(batch_layout)

        # 添加标签页
        tab_widget.addTab(env_tab, "环境设置")
        tab_widget.addTab(batch_tab, "批量操作")

        # 添加所有组件到主布局
        layout.addLayout(toolbar)
        layout.addWidget(self.browser_list)
        layout.addWidget(tab_widget)

    def connect_signals(self):
        # 连接按钮信号
        self.refresh_btn.clicked.connect(self.refresh_browser_list)
        self.open_btn.clicked.connect(self.open_selected_browsers)
        self.close_btn.clicked.connect(self.close_selected_browsers)
        self.copy_btn.clicked.connect(self.copy_selected_profiles)
        self.delete_btn.clicked.connect(self.delete_selected_profiles)
        self.create_btn.clicked.connect(self.create_environment)
        self.sync_btn.clicked.connect(self.toggle_sync)

        # 添加排列按钮信号
        self.auto_arrange_btn.clicked.connect(self.auto_arrange_windows)
        self.popup_custom_arrange_btn.clicked.connect(self.show_custom_arrange_dialog)

        # 浏览按钮信号
        self.chrome_browse_btn.clicked.connect(self.browse_chrome_path)
        self.data_browse_btn.clicked.connect(self.browse_data_path)

        # 批量操作
        self.open_url_btn.clicked.connect(self.batch_open_url)
        self.click_btn.clicked.connect(self.batch_click)
        self.input_btn.clicked.connect(self.batch_input)

        # 设置
        self.browser_list.itemDoubleClicked.connect(self.open_browser)
        self.browser_list.customContextMenuRequested.connect(self.show_context_menu)

        # 添加右键设置主控窗口选项
        self.browser_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)

    def show_custom_arrange_dialog(self):
        """显示自定义排列对话框"""
        # 创建对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("自定义排列设置")
        dialog.setMinimumWidth(400)

        # 创建布局
        layout = QVBoxLayout(dialog)

        # 创建网格布局
        grid_layout = QGridLayout()

        # 第一行
        grid_layout.addWidget(QLabel("起始X坐标:"), 0, 0)
        start_x = QLineEdit()
        start_x.setText(self.start_x.text())
        grid_layout.addWidget(start_x, 0, 1)

        grid_layout.addWidget(QLabel("起始Y坐标:"), 0, 2)
        start_y = QLineEdit()
        start_y.setText(self.start_y.text())
        grid_layout.addWidget(start_y, 0, 3)

        # 第二行
        grid_layout.addWidget(QLabel("窗口宽度:"), 1, 0)
        window_width = QLineEdit()
        window_width.setText(self.window_width.text())
        grid_layout.addWidget(window_width, 1, 1)

        grid_layout.addWidget(QLabel("窗口高度:"), 1, 2)
        window_height = QLineEdit()
        window_height.setText(self.window_height.text())
        grid_layout.addWidget(window_height, 1, 3)

        # 添加自动铺满屏幕选项
        auto_fill_screen = QCheckBox("自动铺满屏幕")
        auto_fill_screen.setObjectName("auto_fill_screen")
        auto_fill_screen.setChecked(self.auto_fill_screen.isChecked())
        grid_layout.addWidget(auto_fill_screen, 4, 0, 1, 4)

        # 第三行
        grid_layout.addWidget(QLabel("水平间距:"), 2, 0)
        h_spacing = QLineEdit()
        h_spacing.setText(self.h_spacing.text())
        grid_layout.addWidget(h_spacing, 2, 1)

        grid_layout.addWidget(QLabel("垂直间距:"), 2, 2)
        v_spacing = QLineEdit()
        v_spacing.setText(self.v_spacing.text())
        grid_layout.addWidget(v_spacing, 2, 3)

        # 第四行
        grid_layout.addWidget(QLabel("每行窗口数:"), 3, 0)
        windows_per_row = QLineEdit()
        windows_per_row.setText(self.windows_per_row.text())
        grid_layout.addWidget(windows_per_row, 3, 1)

        grid_layout.addWidget(QLabel("每列窗口数:"), 3, 2)
        windows_per_column = QLineEdit()
        windows_per_column.setText(self.windows_per_column.text())
        windows_per_column.setToolTip("0表示不限制列数")
        grid_layout.addWidget(windows_per_column, 3, 3)

        # 将网格布局添加到主布局
        layout.addLayout(grid_layout)

        # 添加按钮
        button_layout = QHBoxLayout()
        ok_button = QPushButton("应用并排列")
        cancel_button = QPushButton("取消")

        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        # 连接按钮信号
        ok_button.clicked.connect(
            lambda: self.apply_custom_arrange(
                dialog,
                start_x.text(),
                start_y.text(),
                window_width.text(),
                window_height.text(),
                h_spacing.text(),
                v_spacing.text(),
                windows_per_row.text(),
                windows_per_column.text(),
            )
        )
        cancel_button.clicked.connect(dialog.reject)

        # 显示对话框
        dialog.exec()

    def apply_custom_arrange(
        self,
        dialog,
        start_x,
        start_y,
        width,
        height,
        h_spacing,
        v_spacing,
        windows_per_row,
        windows_per_column,
    ):
        """应用自定义排列设置并排列窗口"""
        try:
            # 更新主窗口中的设置
            self.start_x.setText(start_x)
            self.start_y.setText(start_y)
            self.window_width.setText(width)
            self.window_height.setText(height)
            self.h_spacing.setText(h_spacing)
            self.v_spacing.setText(v_spacing)
            self.windows_per_row.setText(windows_per_row)
            self.windows_per_column.setText(windows_per_column)
            # 记录对话框中的自动铺满屏幕选项状态
            self.auto_fill_screen.setChecked(
                dialog.findChild(QCheckBox, "auto_fill_screen").isChecked()
            )

            # 关闭对话框
            dialog.accept()

            # 执行自定义排列
            self.custom_arrange_windows()

        except Exception as e:
            QMessageBox.critical(dialog, "错误", f"应用设置失败: {str(e)}")

    # 添加浏览文件夹功能
    def browse_chrome_path(self):
        """浏览并选择Chrome可执行文件"""
        from PyQt6.QtWidgets import QFileDialog

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Chrome可执行文件",
            self.chrome_path_input.text(),
            "可执行文件 (*.exe)",
        )

        if file_path:
            self.chrome_path_input.setText(file_path)
            self.chrome_path = file_path
            self.save_settings()

    def browse_data_path(self):
        """浏览并选择数据保存目录"""
        from PyQt6.QtWidgets import QFileDialog

        folder_path = QFileDialog.getExistingDirectory(
            self, "选择数据保存目录", self.data_path.text()
        )

        if folder_path:
            self.data_path.setText(folder_path)
            self.save_settings()

    def load_settings(self):
        try:
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                if "chrome_path" in settings:
                    self.chrome_path = settings["chrome_path"]
                    self.chrome_path_input.setText(settings["chrome_path"])
                if "data_path" in settings:
                    self.data_path.setText(settings["data_path"])
                if "window_count" in settings:
                    self.window_count.setValue(settings["window_count"])
                if "url" in settings:
                    self.url_input.setText(settings["url"])
                # 加载自定义排列设置
                if "start_x" in settings:
                    self.start_x.setText(str(settings["start_x"]))
                if "start_y" in settings:
                    self.start_y.setText(str(settings["start_y"]))
                if "window_width" in settings:
                    self.window_width.setText(str(settings["window_width"]))
                if "window_height" in settings:
                    self.window_height.setText(str(settings["window_height"]))
                if "h_spacing" in settings:
                    self.h_spacing.setText(str(settings["h_spacing"]))
                if "v_spacing" in settings:
                    self.v_spacing.setText(str(settings["v_spacing"]))
                if "windows_per_row" in settings:
                    self.windows_per_row.setText(str(settings["windows_per_row"]))
                if "windows_per_column" in settings:
                    self.windows_per_column.setText(str(settings["windows_per_column"]))
                if "auto_fill_screen" in settings:
                    self.auto_fill_screen.setChecked(settings["auto_fill_screen"])
        except Exception as e:
            print(f"加载设置失败: {str(e)}")

    def save_settings(self):
        try:
            settings = {
                "chrome_path": self.chrome_path,
                "data_path": self.data_path.text(),
                "window_count": self.window_count.value(),
                "url": self.url_input.text(),
                "start_x": self.start_x.text(),
                "start_y": self.start_y.text(),
                "window_width": self.window_width.text(),
                "window_height": self.window_height.text(),
                "h_spacing": self.h_spacing.text(),
                "v_spacing": self.v_spacing.text(),
                "windows_per_row": self.windows_per_row.text(),
                "windows_per_column": self.windows_per_column.text(),
                "auto_fill_screen": self.auto_fill_screen.isChecked(),
            }
            with open("settings.json", "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            QMessageBox.warning(self, "保存设置失败", f"保存设置时出错: {str(e)}")

    def create_environment(self):
        """创建Chrome多开环境"""
        try:
            # 获取设置
            chrome_path = self.chrome_path_input.text()
            data_path = os.path.abspath(self.data_path.text())
            count = self.window_count.value()

            # 检查Chrome是否存在
            if not os.path.exists(chrome_path):
                QMessageBox.warning(self, "错误", f"Chrome路径不存在: {chrome_path}")
                return

            # 创建数据目录
            if not os.path.exists(data_path):
                os.makedirs(data_path)

            # 创建Data文件夹
            data_dir = os.path.join(data_path, "Data")
            if not os.path.exists(data_dir):
                os.makedirs(data_dir)

            # 创建快捷方式
            shell = win32com.client.Dispatch("WScript.Shell")

            for i in range(1, count + 1):
                # 创建用户数据目录
                profile_dir = os.path.join(data_dir, str(i))
                if not os.path.exists(profile_dir):
                    os.makedirs(profile_dir)

                # 创建快捷方式
                shortcut_path = os.path.join(data_path, f"{i}.lnk")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.TargetPath = chrome_path
                shortcut.Arguments = f'--user-data-dir="{profile_dir}"'
                shortcut.WorkingDirectory = os.path.dirname(chrome_path)
                shortcut.save()

            self.user_data_dir = data_path
            self.chrome_path = chrome_path
            self.save_settings()

            self.refresh_browser_list()

        except Exception as e:
            QMessageBox.warning(self, "创建环境失败", f"创建环境时出错: {str(e)}")

    def refresh_browser_list(self):
        """刷新浏览器列表"""
        try:
            # 存储当前选中状态
            selected_numbers = self.get_selected_profiles()

            # 记录主控窗口
            master_number = None
            for i in range(self.browser_list.topLevelItemCount()):
                item = self.browser_list.topLevelItem(i)
                if item.text(3) == "是":
                    master_number = int(item.text(1))
                    break

            self.browser_list.clear()

            # 获取数据目录
            data_path = os.path.abspath(self.data_path.text())

            data_dir = os.path.join(data_path, "Data")

            if not os.path.exists(data_path):
                return

            # 查找所有快捷方式
            profiles = []
            for filename in os.listdir(data_path):
                if filename.endswith(".lnk") and filename[:-4].isdigit():
                    number = int(filename[:-4])
                    profiles.append(number)

            # 按编号排序
            profiles.sort()

            # 获取正在运行的Chrome进程
            running_chrome = self.get_running_chrome_processes()

            # 添加到列表
            for number in profiles:
                item = QTreeWidgetItem()

                # 设置复选框
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)

                # 恢复选中状态
                if number in selected_numbers:
                    item.setCheckState(0, Qt.CheckState.Checked)
                else:
                    item.setCheckState(0, Qt.CheckState.Unchecked)

                # 设置编号
                item.setText(1, str(number))

                # 设置标题
                item.setText(2, f"Chrome实例 {number}")

                # 设置是否为主控
                is_master = master_number is not None and number == master_number
                item.setText(3, "是" if is_master else "否")

                # 设置状态
                if number in running_chrome:
                    item.setText(4, "已运行")
                    # 存储窗口句柄
                    hwnd = running_chrome[number]
                    self.browser_processes[number] = hwnd

                    # 如果是主控窗口，更新主控窗口句柄
                    if is_master:
                        self.master_window = hwnd
                else:
                    item.setText(4, "未运行")

                self.browser_list.addTopLevelItem(item)

            # 使界面响应更快
            QApplication.processEvents()

            # 应用自定义图标
            self.apply_icons_to_chrome_windows()

        except Exception as e:
            QMessageBox.warning(self, "刷新列表失败", f"刷新浏览器列表时出错: {str(e)}")

    def get_running_chrome_processes(self):
        """获取正在运行的Chrome进程"""
        running_chrome = {}

        try:

            def callback(hwnd, chrome_dict):
                if win32gui.IsWindowVisible(hwnd):
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    try:
                        handle = win32api.OpenProcess(
                            win32con.PROCESS_QUERY_INFORMATION
                            | win32con.PROCESS_VM_READ,
                            False,
                            pid,
                        )
                        path = win32process.GetModuleFileNameEx(handle, 0)
                        win32api.CloseHandle(handle)

                        if "chrome.exe" in path.lower():
                            cmdline = None
                            try:
                                proc = psutil.Process(pid)
                                cmdline = proc.cmdline()
                            except Exception as e:
                                print(f"获取进程信息失败: {str(e)}")

                            if cmdline:
                                for cmd in cmdline:
                                    if "--user-data-dir=" in cmd:
                                        # 从命令行参数中提取数据目录
                                        data_dir = cmd.split("--user-data-dir=")[
                                            1
                                        ].strip("\"'")
                                        # 从数据目录中提取编号
                                        dirname = os.path.basename(data_dir)
                                        if dirname.isdigit():
                                            number = int(dirname)
                                            running_chrome[number] = hwnd
                    except Exception as e:
                        print(f"获取进程信息失败: {str(e)}")
                return True

            win32gui.EnumWindows(callback, running_chrome)
        except Exception as e:
            print(f"获取Chrome进程失败: {str(e)}")

        return running_chrome

    def get_selected_profiles(self):
        """获取选中的配置文件编号"""
        selected = []
        for i in range(self.browser_list.topLevelItemCount()):
            item = self.browser_list.topLevelItem(i)
            if item.checkState(0) == Qt.CheckState.Checked:
                number = int(item.text(1))
                selected.append(number)
        return selected

    def open_browser(self, item):
        """打开单个浏览器"""
        number = int(item.text(1))
        self.open_browser_by_number(number)

    def open_browser_by_number(self, number):
        """根据编号打开浏览器"""
        try:
            data_path = os.path.abspath(self.data_path.text())
            shortcut_path = os.path.join(data_path, f"{number}.lnk")

            if os.path.exists(shortcut_path):
                subprocess.Popen(f'start "" "{shortcut_path}"', shell=True)

                # 稍等一下再刷新列表和应用图标
                QTimer.singleShot(2000, lambda: self.after_browser_opened(number))
            else:
                QMessageBox.warning(self, "错误", f"快捷方式不存在: {shortcut_path}")
        except Exception as e:
            QMessageBox.warning(self, "打开浏览器失败", f"打开浏览器时出错: {str(e)}")

    def after_browser_opened(self, number):
        """浏览器打开后的处理"""
        self.refresh_browser_list()

        # 确保当前编号在browser_processes中
        if number in self.browser_processes:
            hwnd = self.browser_processes[number]

            # 如果没有为此编号生成图标，则生成
            if number not in self.profile_icons:
                icon_path = self.generate_color_icon(number)
                if icon_path:
                    self.profile_icons[number] = icon_path

            # 应用图标
            if number in self.profile_icons:
                self.set_chrome_icon(hwnd, self.profile_icons[number])

    def open_selected_browsers(self):
        """打开选中的浏览器"""
        selected = self.get_selected_profiles()
        if not selected:
            QMessageBox.information(self, "提示", "请先选择要打开的浏览器")
            return

        for number in selected:
            self.open_browser_by_number(number)

    def close_browser_by_number(self, number):
        """根据编号关闭浏览器"""
        try:
            # 如果存在窗口句柄，直接使用窗口句柄关闭
            if number in self.browser_processes:
                hwnd = self.browser_processes[number]
                if hwnd:
                    # 发送关闭消息到窗口
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    # 更新进程记录
                    del self.browser_processes[number]
            else:
                # 如果没有存储窗口句柄，尝试通过命令行查找
                for proc in psutil.process_iter(["pid", "name", "cmdline"]):
                    if proc.info["name"] and "chrome.exe" in proc.info["name"].lower():
                        cmdline = proc.info.get("cmdline", [])
                        for cmd in cmdline:
                            if "--user-data-dir=" in cmd:
                                data_dir = cmd.split("--user-data-dir=")[1].strip("\"'")
                                dirname = os.path.basename(data_dir)
                                if dirname.isdigit() and int(dirname) == number:
                                    try:
                                        proc.terminate()
                                        break
                                    except Exception as e:
                                        print(e)

            # 稍等一下再刷新列表
            QTimer.singleShot(1000, self.refresh_browser_list)
        except Exception as e:
            print(f"关闭浏览器时出错: {str(e)}")

    def close_selected_browsers(self):
        """关闭选中的浏览器"""
        selected = self.get_selected_profiles()
        if not selected:
            QMessageBox.information(self, "提示", "请先选择要关闭的浏览器")
            return

        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

            for number in selected:
                self.close_browser_by_number(number)

            # 所有浏览器都已经处理完毕，恢复光标并刷新列表
            QApplication.restoreOverrideCursor()
            self.refresh_browser_list()

        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "关闭浏览器失败", f"关闭浏览器时出错: {str(e)}")

    def copy_selected_profiles(self):
        """复制选中的配置文件"""
        selected = self.get_selected_profiles()
        if not selected:
            QMessageBox.information(self, "提示", "请先选择要复制的浏览器配置")
            return

        # 检查是否有正在运行的浏览器
        running_browsers = []
        for number in selected:
            if number in self.browser_processes:
                running_browsers.append(number)

        if running_browsers:
            QMessageBox.warning(
                self,
                "警告",
                f"以下浏览器正在运行中,请先关闭:\n{', '.join(map(str, running_browsers))}",
            )
            return

        # 添加输入对话框，询问用户要创建几份副本
        copy_count, ok = QInputDialog.getInt(
            self, "复制数量", "请输入要创建的副本数量:", 1, 1, 100, 1
        )

        if not ok:
            return  # 用户取消了操作

        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

            data_path = os.path.abspath(self.data_path.text())
            data_dir = os.path.join(data_path, "Data")

            # 获取最大编号
            max_number = 0
            for filename in os.listdir(data_path):
                if filename.endswith(".lnk") and filename[:-4].isdigit():
                    number = int(filename[:-4])
                    max_number = max(max_number, number)

            # 记录新创建的配置
            new_profiles = []

            # 创建新的配置文件
            for number in selected:
                for _ in range(copy_count):
                    max_number += 1
                    new_profiles.append(max_number)

                    # 复制数据目录
                    src_dir = os.path.join(data_dir, str(number))
                    dst_dir = os.path.join(data_dir, str(max_number))

                    if os.path.exists(src_dir):
                        shutil.copytree(src_dir, dst_dir)

                    # 创建快捷方式
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shortcut_path = os.path.join(data_path, f"{max_number}.lnk")
                    shortcut = shell.CreateShortCut(shortcut_path)
                    shortcut.TargetPath = self.chrome_path_input.text()
                    shortcut.Arguments = f'--user-data-dir="{dst_dir}"'
                    shortcut.WorkingDirectory = os.path.dirname(
                        self.chrome_path_input.text()
                    )
                    shortcut.save()

            # 先恢复光标，然后刷新列表和显示消息
            QApplication.restoreOverrideCursor()
            self.refresh_browser_list()
            QMessageBox.information(
                self,
                "成功",
                f"已成功复制 {len(selected)} 个浏览器配置，共创建了 {len(new_profiles)} 个新配置",
            )

        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "复制配置失败", f"复制浏览器配置时出错: {str(e)}")

    def delete_selected_profiles(self):
        """删除选中的配置文件"""
        selected = self.get_selected_profiles()
        if not selected:
            QMessageBox.information(self, "提示", "请先选择要删除的浏览器配置")
            return

        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除选中的 {len(selected)} 个浏览器配置吗？此操作不可恢复！",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

            data_path = os.path.abspath(self.data_path.text())
            data_dir = os.path.join(data_path, "Data")

            for number in selected:
                # 先关闭浏览器
                self.close_browser_by_number(number)

                # 删除快捷方式
                shortcut_path = os.path.join(data_path, f"{number}.lnk")
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)

                # 删除数据目录
                profile_dir = os.path.join(data_dir, str(number))
                if os.path.exists(profile_dir):
                    shutil.rmtree(profile_dir)

            # 先恢复光标，然后刷新列表和显示消息
            QApplication.restoreOverrideCursor()
            self.refresh_browser_list()

        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "删除配置失败", f"删除浏览器配置时出错: {str(e)}")

    def start_sync(self, selected_items):
        """开始同步"""
        try:
            # 确保主控窗口存在
            if not self.master_window:
                QMessageBox.warning(
                    self, "错误", "未设置主控窗口，请先设置一个主控窗口"
                )
                return False

            # 保存选中的窗口列表
            self.sync_windows = []

            # 收集所有选中的窗口
            for number in selected_items:
                if number in self.browser_processes:
                    hwnd = self.browser_processes[number]
                    if hwnd != self.master_window:  # 排除主控窗口
                        self.sync_windows.append(hwnd)

            if not self.sync_windows:
                QMessageBox.warning(
                    self, "错误", "没有可同步的窗口，请确保至少有一个非主控窗口"
                )
                return False

            # 启动键盘和鼠标钩子
            if not self.hook_thread:
                self.is_syncing = True
                self.hook_thread = threading.Thread(target=self.message_loop)
                self.hook_thread.daemon = True
                self.hook_thread.start()

                # 设置钩子
                self.keyboard_hook = keyboard.hook(self.on_keyboard_event)
                self.mouse_hook_id = mouse.hook(self.on_mouse_event)

                # 启动插件窗口监控线程
                self.popup_monitor_thread = threading.Thread(target=self.monitor_popups)
                self.popup_monitor_thread.daemon = True
                self.popup_monitor_thread.start()

                print(
                    f"已启动同步，主控窗口: {self.master_window}, 同步窗口: {self.sync_windows}"
                )
                return True

        except Exception as e:
            self.stop_sync()  # 确保清理资源
            QMessageBox.warning(self, "开启同步失败", f"开启同步时出错: {str(e)}")
            print(f"开启同步失败: {str(e)}")
            return False

    def stop_sync(self):
        """停止同步"""
        try:
            self.is_syncing = False

            # 移除键盘钩子
            if self.keyboard_hook:
                keyboard.unhook(self.keyboard_hook)
                self.keyboard_hook = None

            # 移除鼠标钩子
            if self.mouse_hook_id:
                mouse.unhook(self.mouse_hook_id)
                self.mouse_hook_id = None

            # 等待监控线程结束
            if self.hook_thread and self.hook_thread.is_alive():
                self.hook_thread.join(timeout=1.0)
                self.hook_thread = None

            # 等待插件窗口监控线程结束
            if self.popup_monitor_thread and self.popup_monitor_thread.is_alive():
                self.popup_monitor_thread.join(timeout=1.0)
                self.popup_monitor_thread = None

            # 清理资源（保留主窗口设置）
            self.sync_windows.clear()
            self.popup_mappings.clear()

            print("同步已停止")
            return True

        except Exception as e:
            QMessageBox.warning(self, "停止同步失败", f"停止同步时出错: {str(e)}")
            print(f"停止同步失败: {str(e)}")
            return False

    def toggle_sync(self):
        """切换同步状态"""
        if not self.is_syncing:
            # 获取选中的浏览器
            selected = self.get_selected_profiles()
            if not selected:
                QMessageBox.information(self, "提示", "请选择要同步的浏览器")
                return

            # 检查主控窗口
            master_set = False
            for i in range(self.browser_list.topLevelItemCount()):
                item = self.browser_list.topLevelItem(i)
                if item.text(3) == "是":
                    master_set = True
                    break

            if not master_set and selected:
                # 自动将第一个选中的窗口设为主控窗口
                for i in range(self.browser_list.topLevelItemCount()):
                    item = self.browser_list.topLevelItem(i)
                    if int(item.text(1)) == selected[0]:
                        self.set_master_window(item)
                        break

            # 启动同步
            if self.start_sync(selected):
                # 更新按钮状态
                self.sync_btn.setText("■ 停止同步")

                # 启动定时刷新
                self.sync_timer = QTimer()
                self.sync_timer.timeout.connect(self.refresh_browser_list)
                self.sync_timer.start(5000)  # 每5秒刷新一次
        else:
            # 停止同步
            if self.stop_sync():
                # 更新按钮状态
                self.sync_btn.setText("▶ 开始同步")

                # 停止定时刷新
                if self.sync_timer:
                    self.sync_timer.stop()
                    self.sync_timer = None

    def show_context_menu(self, pos):
        """显示右键菜单"""
        item = self.browser_list.itemAt(pos)
        menu = QMenu(self)

        # 如果在列表项上右击，显示特定于项目的菜单
        if item:
            number = int(item.text(1))

            open_action = menu.addAction("打开")
            close_action = menu.addAction("关闭")
            copy_action = menu.addAction("复制")
            delete_action = menu.addAction("删除")

            # 添加设置主控窗口选项
            set_master_action = menu.addAction("设为主控窗口")

            menu.addSeparator()

        # 无论是否在列表项上，都添加全选和反选选项
        select_all_action = menu.addAction("全选")
        invert_selection_action = menu.addAction("反选")

        action = menu.exec(self.browser_list.mapToGlobal(pos))

        # 如果菜单关闭后没有选择任何操作，直接返回
        if not action:
            return

        # 处理全选和反选动作
        if action == select_all_action:
            self.select_all_browsers()
            return
        elif action == invert_selection_action:
            self.invert_selection()
            return

        # 在菜单关闭后再次检查item是否有效，处理特定于项目的操作
        if item:
            if action == open_action:
                self.open_browser_by_number(number)
            elif action == close_action:
                self.close_browser_by_number(number)
            elif action == copy_action:
                # 不改变选中状态，直接复制
                self.copy_selected_profiles()
            elif action == delete_action:
                # 不改变选中状态，直接删除
                self.delete_selected_profiles()
            elif action == set_master_action:
                # 设置为主控窗口
                self.set_master_window(item)

    def select_all_browsers(self):
        """全选所有浏览器"""
        for i in range(self.browser_list.topLevelItemCount()):
            item = self.browser_list.topLevelItem(i)
            item.setCheckState(0, Qt.CheckState.Checked)

    def invert_selection(self):
        """反选浏览器"""
        for i in range(self.browser_list.topLevelItemCount()):
            item = self.browser_list.topLevelItem(i)
            if item.checkState(0) == Qt.CheckState.Checked:
                item.setCheckState(0, Qt.CheckState.Unchecked)
            else:
                item.setCheckState(0, Qt.CheckState.Checked)

    def batch_open_url(self):
        """批量打开URL"""
        selected = self.get_selected_profiles()
        if not selected:
            QMessageBox.information(self, "提示", "请先选择要操作的浏览器")
            return

        url = self.url_input.text()
        if not url:
            QMessageBox.information(self, "提示", "请输入要打开的URL")
            return

        if not url.startswith(("http://", "https://")):
            url = "https://" + url

        new_tab = self.new_page_checkbox.isChecked()

        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

            # 保存当前URL到设置
            self.save_settings()

            # 分批处理，避免UI阻塞
            self.process_batch_url_open(selected, url, new_tab, 0)

        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "批量打开URL失败", f"批量打开URL时出错: {str(e)}")

    def process_batch_url_open(self, selected, url, new_tab, index):
        """分批处理URL打开，避免UI阻塞"""
        if index >= len(selected):
            # 所有浏览器都已处理完毕
            QApplication.restoreOverrideCursor()

            return

        # 处理当前浏览器
        number = selected[index]

        # 检查浏览器是否已经运行
        running = False
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            if proc.info["name"] and "chrome.exe" in proc.info["name"].lower():
                cmdline = proc.info.get("cmdline", [])
                for cmd in cmdline:
                    if "--user-data-dir=" in cmd:
                        data_dir = cmd.split("--user-data-dir=")[1].strip("\"'")
                        dirname = os.path.basename(data_dir)
                        if dirname.isdigit() and int(dirname) == number:
                            running = True
                            break

        if not running:
            # 如果浏览器未运行，打开它并稍后继续处理下一个
            self.open_browser_by_number(number)
            QTimer.singleShot(
                2000,
                lambda: self.open_url_and_continue(
                    number, url, new_tab, selected, index
                ),
            )
        else:
            # 浏览器已经运行，打开URL并继续处理下一个
            self.open_url_in_browser(number, url, new_tab)
            QTimer.singleShot(
                300,
                lambda: self.process_batch_url_open(selected, url, new_tab, index + 1),
            )

    def open_url_and_continue(self, number, url, new_tab, selected, index):
        """打开URL并继续处理下一个浏览器"""
        self.open_url_in_browser(number, url, new_tab)
        QTimer.singleShot(
            300, lambda: self.process_batch_url_open(selected, url, new_tab, index + 1)
        )

    def open_url_in_browser(self, number, url, new_tab=True):
        """在指定浏览器中打开URL"""
        try:
            # 获取数据目录
            data_path = os.path.abspath(self.data_path.text())
            profile_dir = os.path.join(data_path, "Data", str(number))

            if new_tab:
                # 在新标签页中打开URL
                command = f'"{self.chrome_path}" --user-data-dir="{profile_dir}" {url}'
            else:
                # 在当前页面打开URL
                command = f'"{self.chrome_path}" --user-data-dir="{profile_dir}" --new-window {url}'

            subprocess.Popen(command, shell=True)

        except Exception as e:
            print(f"在浏览器 {number} 中打开URL时出错: {str(e)}")

    def batch_click(self):
        """批量点击操作"""
        selected = self.get_selected_profiles()
        if not selected:
            QMessageBox.information(self, "提示", "请先选择要操作的浏览器")
            return

        selector = self.selector_input.text()
        if not selector:
            QMessageBox.information(self, "提示", "请输入CSS选择器")
            return

        delay = self.click_delay.value()

        try:
            # 创建一个简单的JavaScript脚本来点击元素
            js_code = f"""
            (function() {{
                let elements = document.querySelectorAll('{selector}');
                if (elements.length > 0) {{
                    elements[0].click();
                    return "已点击元素: " + elements.length;
                }}
                return "未找到匹配的元素";
            }})();
            """

            # 在每个选中的浏览器中执行脚本
            for number in selected:
                # 检查浏览器是否已经运行
                for proc in psutil.process_iter(["pid", "name", "cmdline"]):
                    if proc.info["name"] and "chrome.exe" in proc.info["name"].lower():
                        cmdline = proc.info.get("cmdline", [])
                        for cmd in cmdline:
                            if "--user-data-dir=" in cmd:
                                data_dir = cmd.split("--user-data-dir=")[1].strip("\"'")
                                dirname = os.path.basename(data_dir)
                                if dirname.isdigit() and int(dirname) == number:
                                    # 实际执行点击操作需要使用如Chrome DevTools Protocol
                                    # 或Selenium等工具，这里只是示例
                                    pass

            QMessageBox.information(
                self,
                "提示",
                f"已在选中的浏览器中点击元素: {selector}\n实际执行需集成Selenium或CDP",
            )

        except Exception as e:
            QMessageBox.warning(self, "批量点击失败", f"批量点击操作时出错: {str(e)}")

    def batch_input(self):
        """批量输入文本"""
        selected = self.get_selected_profiles()
        if not selected:
            QMessageBox.information(self, "提示", "请先选择要操作的浏览器")
            return

        selector = self.input_selector.text()
        if not selector:
            QMessageBox.information(self, "提示", "请输入CSS选择器")
            return

        text = self.input_text.text()

        try:
            # 创建一个简单的JavaScript脚本来设置输入框的值
            js_code = f"""
            (function() {{
                let elements = document.querySelectorAll('{selector}');
                if (elements.length > 0) {{
                    elements[0].value = '{text}';
                    // 触发一个input事件让表单知道值已更新
                    let event = new Event('input', {{ bubbles: true }});
                    elements[0].dispatchEvent(event);
                    return "已设置输入文本: " + elements.length;
                }}
                return "未找到匹配的元素";
            }})();
            """

            # 在每个选中的浏览器中执行脚本
            for number in selected:
                # 检查浏览器是否已经运行
                for proc in psutil.process_iter(["pid", "name", "cmdline"]):
                    if proc.info["name"] and "chrome.exe" in proc.info["name"].lower():
                        cmdline = proc.info.get("cmdline", [])
                        for cmd in cmdline:
                            if "--user-data-dir=" in cmd:
                                data_dir = cmd.split("--user-data-dir=")[1].strip("\"'")
                                dirname = os.path.basename(data_dir)
                                if dirname.isdigit() and int(dirname) == number:
                                    # 实际执行输入操作需要使用如Chrome DevTools Protocol
                                    # 或Selenium等工具，这里只是示例
                                    pass

            QMessageBox.information(
                self,
                "提示",
                f"已在选中的浏览器中输入文本: {text}\n实际执行需集成Selenium或CDP",
            )
        except Exception as e:
            QMessageBox.warning(self, "批量输入失败", f"批量输入操作时出错: {str(e)}")

    def set_master_window(self, item):
        """设置主控窗口"""
        number = int(item.text(1))
        # 检查浏览器是否运行
        if number not in self.browser_processes:
            QMessageBox.warning(self, "错误", "请先打开该浏览器")
            return

        # 清除之前的主控标记
        for i in range(self.browser_list.topLevelItemCount()):
            prev_item = self.browser_list.topLevelItem(i)
            prev_item.setText(3, "否")

            # 如果之前是主控窗口，恢复窗口标题
            if prev_item.text(3) == "是":
                prev_number = int(prev_item.text(1))
                if prev_number in self.browser_processes:
                    prev_hwnd = self.browser_processes[prev_number]
                    try:
                        # 获取当前标题
                        title = win32gui.GetWindowText(prev_hwnd)
                        if title.startswith("[主控]"):
                            # 移除[主控]前缀
                            new_title = title.replace("[主控]", "").strip()
                            win32gui.SetWindowText(prev_hwnd, new_title)
                    except Exception as e:
                        print(f"恢复窗口标题失败: {str(e)}")

        # 设置新的主控窗口
        item.setText(3, "是")

        # 直接使用存储的窗口句柄
        hwnd = self.browser_processes[number]
        if hwnd:
            self.master_window = hwnd

            # 更新窗口标题，添加[主控]前缀
            try:
                title = win32gui.GetWindowText(hwnd)
                if not title.startswith("[主控]"):
                    new_title = f"[主控] {title}"
                    win32gui.SetWindowText(hwnd, new_title)
            except Exception as e:
                print(f"更新窗口标题失败: {str(e)}")

            # 添加红色边框
            self.add_border_to_master(hwnd)
        else:
            QMessageBox.warning(self, "错误", "无法获取浏览器窗口句柄")
            return

        # 更新界面
        QMessageBox.information(self, "成功", f"已将浏览器 {number} 设为主控窗口")

    def get_chrome_window_by_pid(self, pid):
        """通过进程ID获取Chrome窗口句柄"""

        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd):
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid:
                    try:
                        handle = win32api.OpenProcess(
                            win32con.PROCESS_QUERY_INFORMATION
                            | win32con.PROCESS_VM_READ,
                            False,
                            found_pid,
                        )
                        path = win32process.GetModuleFileNameEx(handle, 0)
                        win32api.CloseHandle(handle)

                        if "chrome.exe" in path.lower():
                            title = win32gui.GetWindowText(hwnd)
                            if title and not title.startswith("Chrome 传递"):
                                class_name = win32gui.GetClassName(hwnd)
                                if "Chrome_WidgetWin_1" in class_name:
                                    hwnds.append(hwnd)
                    except Exception as e:
                        print(f"获取进程信息失败: {str(e)}")
            return True

        window_handles = []
        win32gui.EnumWindows(callback, window_handles)
        if window_handles:
            return window_handles[0]  # 返回第一个匹配的窗口
        return None

    def add_border_to_master(self, hwnd):
        """为主控窗口添加红色边框"""
        try:
            # 使用DWM API设置红色边框
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,  # 窗口句柄
                self.DWMWA_BORDER_COLOR,  # 窗口边框颜色属性
                ctypes.byref(ctypes.c_int(0x000000FF)),  # 设置为红色(0x000000FF)
                ctypes.sizeof(ctypes.c_int),  # 参数大小
            )

            # 重绘窗口
            win32gui.SetWindowPos(
                hwnd,
                0,
                0,
                0,
                0,
                0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_FRAMECHANGED,
            )
        except Exception as e:
            print(f"添加边框失败: {str(e)}")

    def get_chrome_popups(self, chrome_hwnd):
        """获取Chrome弹出窗口"""
        popups = []

        def enum_windows_callback(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return

                class_name = win32gui.GetClassName(hwnd)
                title = win32gui.GetWindowText(hwnd)
                _, chrome_pid = win32process.GetWindowThreadProcessId(chrome_hwnd)
                _, popup_pid = win32process.GetWindowThreadProcessId(hwnd)

                # 检查是否是Chrome相关窗口
                if popup_pid == chrome_pid:
                    # 检查窗口类型
                    if "Chrome_WidgetWin_1" in class_name:
                        # 检查是否是扩展程序相关窗口，放宽检测条件
                        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)

                        # 扩展窗口的特征
                        is_popup = (
                            "扩展程序" in title
                            or "插件" in title
                            or win32gui.GetParent(hwnd) == chrome_hwnd
                            or (style & win32con.WS_POPUP) != 0
                            or (style & win32con.WS_CHILD) != 0
                            or (ex_style & win32con.WS_EX_TOOLWINDOW) != 0
                        )

                        if is_popup:
                            popups.append(hwnd)

            except Exception as e:
                print(f"枚举窗口失败: {str(e)}")

        win32gui.EnumWindows(enum_windows_callback, None)
        return popups

    def monitor_popups(self):
        """监控插件窗口变化"""
        while self.is_syncing:
            try:
                # 更新插件窗口映射
                if (
                    self.master_window
                    and self.master_window in self.browser_processes.values()
                ):
                    self.popup_mappings[self.master_window] = self.get_chrome_popups(
                        self.master_window
                    )

                    for hwnd in self.sync_windows:
                        if hwnd in self.browser_processes.values():
                            self.popup_mappings[hwnd] = self.get_chrome_popups(hwnd)
            except Exception as e:
                print(f"监控插件窗口失败: {str(e)}")

            time.sleep(0.1)

    def message_loop(self):
        """消息循环"""
        while self.is_syncing:
            time.sleep(0.001)

    def on_mouse_event(self, event):
        """处理鼠标事件"""
        try:
            if self.is_syncing:
                current_window = win32gui.GetForegroundWindow()

                # 检查是否是主控窗口或其插件窗口
                is_master = current_window == self.master_window
                master_popups = self.get_chrome_popups(self.master_window)
                is_popup = current_window in master_popups

                if is_master or is_popup:
                    # 对于移动事件进行优化
                    if isinstance(event, mouse.MoveEvent):
                        # 检查移动距离和时间间隔
                        current_time = time.time()
                        if current_time - self.last_move_time < self.move_interval:
                            return

                        dx = abs(event.x - self.last_mouse_position[0])
                        dy = abs(event.y - self.last_mouse_position[1])
                        if dx < self.mouse_threshold and dy < self.mouse_threshold:
                            return

                        self.last_mouse_position = (event.x, event.y)
                        self.last_move_time = current_time

                    # 获取鼠标位置
                    x, y = mouse.get_position()

                    # 获取当前窗口的相对坐标
                    current_rect = win32gui.GetWindowRect(current_window)
                    rel_x = (x - current_rect[0]) / (current_rect[2] - current_rect[0])
                    rel_y = (y - current_rect[1]) / (current_rect[3] - current_rect[1])

                    # 同步到其他窗口
                    for hwnd in self.sync_windows:
                        try:
                            # 确定目标窗口
                            if is_master:
                                target_hwnd = hwnd
                            else:
                                # 查找对应的扩展程序窗口
                                target_popups = self.get_chrome_popups(hwnd)
                                # 按照相对位置匹配
                                best_match = None
                                min_diff = float("inf")
                                for popup in target_popups:
                                    popup_rect = win32gui.GetWindowRect(popup)
                                    master_rect = win32gui.GetWindowRect(current_window)
                                    # 计算相对位置差异
                                    master_rel_x = (
                                        master_rect[0]
                                        - win32gui.GetWindowRect(self.master_window)[0]
                                    )
                                    master_rel_y = (
                                        master_rect[1]
                                        - win32gui.GetWindowRect(self.master_window)[1]
                                    )
                                    popup_rel_x = (
                                        popup_rect[0] - win32gui.GetWindowRect(hwnd)[0]
                                    )
                                    popup_rel_y = (
                                        popup_rect[1] - win32gui.GetWindowRect(hwnd)[1]
                                    )

                                    diff = abs(master_rel_x - popup_rel_x) + abs(
                                        master_rel_y - popup_rel_y
                                    )
                                    if diff < min_diff:
                                        min_diff = diff
                                        best_match = popup
                                target_hwnd = best_match if best_match else hwnd

                            if not target_hwnd:
                                continue

                            # 获取目标窗口尺寸
                            target_rect = win32gui.GetWindowRect(target_hwnd)

                            # 计算目标坐标
                            client_x = int((target_rect[2] - target_rect[0]) * rel_x)
                            client_y = int((target_rect[3] - target_rect[1]) * rel_y)
                            lparam = win32api.MAKELONG(client_x, client_y)

                            # 处理滚轮事件
                            if isinstance(event, mouse.WheelEvent):
                                try:
                                    wheel_delta = int(event.delta)
                                    if keyboard.is_pressed("ctrl"):
                                        if wheel_delta > 0:
                                            win32gui.PostMessage(
                                                target_hwnd,
                                                win32con.WM_KEYDOWN,
                                                win32con.VK_CONTROL,
                                                0,
                                            )
                                            win32gui.PostMessage(
                                                target_hwnd,
                                                win32con.WM_KEYDOWN,
                                                0xBB,
                                                0,
                                            )  # VK_OEM_PLUS
                                            win32gui.PostMessage(
                                                target_hwnd, win32con.WM_KEYUP, 0xBB, 0
                                            )
                                            win32gui.PostMessage(
                                                target_hwnd,
                                                win32con.WM_KEYUP,
                                                win32con.VK_CONTROL,
                                                0,
                                            )
                                        else:
                                            win32gui.PostMessage(
                                                target_hwnd,
                                                win32con.WM_KEYDOWN,
                                                win32con.VK_CONTROL,
                                                0,
                                            )
                                            win32gui.PostMessage(
                                                target_hwnd,
                                                win32con.WM_KEYDOWN,
                                                0xBD,
                                                0,
                                            )  # VK_OEM_MINUS
                                            win32gui.PostMessage(
                                                target_hwnd, win32con.WM_KEYUP, 0xBD, 0
                                            )
                                            win32gui.PostMessage(
                                                target_hwnd,
                                                win32con.WM_KEYUP,
                                                win32con.VK_CONTROL,
                                                0,
                                            )
                                    else:
                                        vk_code = (
                                            win32con.VK_UP
                                            if wheel_delta > 0
                                            else win32con.VK_DOWN
                                        )
                                        repeat_count = min(abs(wheel_delta) * 3, 6)
                                        for _ in range(repeat_count):
                                            win32gui.PostMessage(
                                                target_hwnd,
                                                win32con.WM_KEYDOWN,
                                                vk_code,
                                                0,
                                            )
                                            win32gui.PostMessage(
                                                target_hwnd,
                                                win32con.WM_KEYUP,
                                                vk_code,
                                                0,
                                            )

                                except Exception as e:
                                    print(f"处理滚轮事件失败: {str(e)}")
                                    continue

                            # 处理鼠标点击
                            elif isinstance(event, mouse.ButtonEvent):
                                if event.event_type == mouse.DOWN:
                                    if event.button == mouse.LEFT:
                                        win32gui.PostMessage(
                                            target_hwnd,
                                            win32con.WM_LBUTTONDOWN,
                                            win32con.MK_LBUTTON,
                                            lparam,
                                        )
                                    elif event.button == mouse.RIGHT:
                                        win32gui.PostMessage(
                                            target_hwnd,
                                            win32con.WM_RBUTTONDOWN,
                                            win32con.MK_RBUTTON,
                                            lparam,
                                        )
                                elif event.event_type == mouse.UP:
                                    if event.button == mouse.LEFT:
                                        win32gui.PostMessage(
                                            target_hwnd,
                                            win32con.WM_LBUTTONUP,
                                            0,
                                            lparam,
                                        )
                                    elif event.button == mouse.RIGHT:
                                        win32gui.PostMessage(
                                            target_hwnd,
                                            win32con.WM_RBUTTONUP,
                                            0,
                                            lparam,
                                        )

                            # 处理鼠标移动
                            elif isinstance(event, mouse.MoveEvent):
                                win32gui.PostMessage(
                                    target_hwnd, win32con.WM_MOUSEMOVE, 0, lparam
                                )

                        except Exception as e:
                            print(f"同步到窗口 {target_hwnd} 失败: {str(e)}")
                            continue

        except Exception as e:
            print(f"处理鼠标事件失败: {str(e)}")

    def on_keyboard_event(self, event):
        """处理键盘事件"""
        try:
            if self.is_syncing:
                current_window = win32gui.GetForegroundWindow()

                # 检查是否是主控窗口或其插件窗口
                is_master = current_window == self.master_window
                master_popups = self.get_chrome_popups(self.master_window)
                is_popup = current_window in master_popups

                if is_master or is_popup:
                    # 获取实际的输入目标窗口

                    # 同步到其他窗口
                    for hwnd in self.sync_windows:
                        try:
                            # 确定目标窗口
                            if is_master:
                                target_hwnd = hwnd
                            else:
                                # 查找对应的扩展程序窗口
                                target_popups = self.get_chrome_popups(hwnd)
                                # 按照相对位置匹配
                                best_match = None
                                min_diff = float("inf")
                                for popup in target_popups:
                                    popup_rect = win32gui.GetWindowRect(popup)
                                    master_rect = win32gui.GetWindowRect(current_window)
                                    # 计算相对位置差异
                                    master_rel_x = (
                                        master_rect[0]
                                        - win32gui.GetWindowRect(self.master_window)[0]
                                    )
                                    master_rel_y = (
                                        master_rect[1]
                                        - win32gui.GetWindowRect(self.master_window)[1]
                                    )
                                    popup_rel_x = (
                                        popup_rect[0] - win32gui.GetWindowRect(hwnd)[0]
                                    )
                                    popup_rel_y = (
                                        popup_rect[1] - win32gui.GetWindowRect(hwnd)[1]
                                    )

                                    diff = abs(master_rel_x - popup_rel_x) + abs(
                                        master_rel_y - popup_rel_y
                                    )
                                    if diff < min_diff:
                                        min_diff = diff
                                        best_match = popup
                                target_hwnd = best_match if best_match else hwnd

                            if not target_hwnd:
                                continue

                            # 处理 Ctrl 组合键
                            if keyboard.is_pressed("ctrl"):
                                # 发送 Ctrl 按下
                                win32gui.PostMessage(
                                    target_hwnd,
                                    win32con.WM_KEYDOWN,
                                    win32con.VK_CONTROL,
                                    0,
                                )

                                # 处理常用组合键
                                if event.name in ["a", "c", "v", "x"]:
                                    vk_code = ord(event.name.upper())
                                    if event.event_type == keyboard.KEY_DOWN:
                                        win32gui.PostMessage(
                                            target_hwnd, win32con.WM_KEYDOWN, vk_code, 0
                                        )
                                        win32gui.PostMessage(
                                            target_hwnd, win32con.WM_KEYUP, vk_code, 0
                                        )
                                    win32gui.PostMessage(
                                        target_hwnd,
                                        win32con.WM_KEYUP,
                                        win32con.VK_CONTROL,
                                        0,
                                    )
                                    continue

                            # 处理普通按键
                            if event.name in [
                                "enter",
                                "backspace",
                                "tab",
                                "esc",
                                "space",
                                "up",
                                "down",
                                "left",
                                "right",
                                "home",
                                "end",
                                "page up",
                                "page down",
                                "delete",
                            ]:
                                vk_map = {
                                    "enter": win32con.VK_RETURN,
                                    "backspace": win32con.VK_BACK,
                                    "tab": win32con.VK_TAB,
                                    "esc": win32con.VK_ESCAPE,
                                    "space": win32con.VK_SPACE,
                                    "up": win32con.VK_UP,
                                    "down": win32con.VK_DOWN,
                                    "left": win32con.VK_LEFT,
                                    "right": win32con.VK_RIGHT,
                                    "home": win32con.VK_HOME,
                                    "end": win32con.VK_END,
                                    "page up": win32con.VK_PRIOR,
                                    "page down": win32con.VK_NEXT,
                                    "delete": win32con.VK_DELETE,
                                }
                                vk_code = vk_map[event.name]
                            else:
                                # 处理普通字符
                                if len(event.name) == 1:
                                    vk_code = win32api.VkKeyScan(event.name[0]) & 0xFF
                                    if event.event_type == keyboard.KEY_DOWN:
                                        # 发送字符消息
                                        win32gui.PostMessage(
                                            target_hwnd,
                                            win32con.WM_CHAR,
                                            ord(event.name[0]),
                                            0,
                                        )
                                    continue
                                else:
                                    continue

                            # 发送按键消息
                            if event.event_type == keyboard.KEY_DOWN:
                                win32gui.PostMessage(
                                    target_hwnd, win32con.WM_KEYDOWN, vk_code, 0
                                )
                            else:
                                win32gui.PostMessage(
                                    target_hwnd, win32con.WM_KEYUP, vk_code, 0
                                )

                            # 释放组合键
                            if keyboard.is_pressed("ctrl"):
                                win32gui.PostMessage(
                                    target_hwnd,
                                    win32con.WM_KEYUP,
                                    win32con.VK_CONTROL,
                                    0,
                                )

                        except Exception as e:
                            print(f"同步到窗口 {target_hwnd} 失败: {str(e)}")

        except Exception as e:
            print(f"处理键盘事件失败: {str(e)}")

    def update_screen_list(self):
        """更新屏幕列表"""
        try:
            screens = []

            def callback(hmonitor, hdc, lprect, lparam):
                try:
                    # 获取显示器信息
                    monitor_info = win32api.GetMonitorInfo(hmonitor)
                    screen_name = f"屏幕 {len(screens) + 1}"
                    if monitor_info["Flags"] & 1:  # MONITORINFOF_PRIMARY
                        screen_name += " (主屏幕)"
                    screens.append(
                        {
                            "name": screen_name,
                            "rect": monitor_info["Monitor"],
                            "work_rect": monitor_info["Work"],
                            "monitor": hmonitor,
                        }
                    )
                except Exception as e:
                    print(f"处理显示器信息失败: {str(e)}")
                return True

            # 定义回调函数类型
            MONITORENUMPROC = ctypes.WINFUNCTYPE(
                ctypes.c_bool,
                ctypes.c_ulong,
                ctypes.c_ulong,
                ctypes.POINTER(wintypes.RECT),
                ctypes.c_longlong,
            )

            # 创建回调函数
            callback_function = MONITORENUMPROC(callback)

            # 枚举显示器
            if (
                ctypes.windll.user32.EnumDisplayMonitors(0, 0, callback_function, 0)
                == 0
            ):
                # EnumDisplayMonitors 失败，尝试使用备用方法
                try:
                    # 获取虚拟屏幕范围
                    virtual_width = win32api.GetSystemMetrics(
                        win32con.SM_CXVIRTUALSCREEN
                    )
                    virtual_height = win32api.GetSystemMetrics(
                        win32con.SM_CYVIRTUALSCREEN
                    )
                    virtual_left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
                    virtual_top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)

                    # 获取主屏幕信息
                    primary_monitor = win32api.MonitorFromPoint(
                        (0, 0), win32con.MONITOR_DEFAULTTOPRIMARY
                    )
                    primary_info = win32api.GetMonitorInfo(primary_monitor)

                    # 添加主屏幕
                    screens.append(
                        {
                            "name": "屏幕 1 (主屏幕)",
                            "rect": primary_info["Monitor"],
                            "work_rect": primary_info["Work"],
                            "monitor": primary_monitor,
                        }
                    )

                    # 尝试获取第二个屏幕
                    try:
                        second_monitor = win32api.MonitorFromPoint(
                            (
                                virtual_left + virtual_width - 1,
                                virtual_top + virtual_height // 2,
                            ),
                            win32con.MONITOR_DEFAULTTONULL,
                        )
                        if second_monitor and second_monitor != primary_monitor:
                            second_info = win32api.GetMonitorInfo(second_monitor)
                            screens.append(
                                {
                                    "name": "屏幕 2",
                                    "rect": second_info["Monitor"],
                                    "work_rect": second_info["Work"],
                                    "monitor": second_monitor,
                                }
                            )
                    except Exception as e:
                        print(f"获取第二个屏幕失败: {str(e)}")
                        pass

                except Exception as e:
                    print(f"备用方法失败: {str(e)}")

            if not screens:
                # 如果仍然没有找到屏幕，使用基本方案
                screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
                screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
                screens.append(
                    {
                        "name": "屏幕 1 (主屏幕)",
                        "rect": (0, 0, screen_width, screen_height),
                        "work_rect": (0, 0, screen_width, screen_height),
                        "monitor": None,
                    }
                )

            # 按照屏幕位置排序（从左到右）
            screens.sort(key=lambda x: x["rect"][0])

            # 更新下拉框选项
            self.screen_combo.clear()
            for screen in screens:
                self.screen_combo.addItem(screen["name"])
            self.screens = screens  # 保存屏幕信息

            # 默认选择主屏幕
            for i, screen in enumerate(screens):
                if "主屏幕" in screen["name"]:
                    self.screen_combo.setCurrentIndex(i)
                    break
            else:
                self.screen_combo.setCurrentIndex(0)

            # 打印调试信息
            print("检测到的屏幕:")
            for screen in screens:
                print(f"名称: {screen['name']}")
                print(f"位置: {screen['rect']}")
                print(f"工作区: {screen['work_rect']}")
                print("---")

        except Exception as e:
            print(f"获取屏幕列表失败: {str(e)}")
            # 使用最基本的方案
            try:
                screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
                screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
                screens = [
                    {
                        "name": "主屏幕",
                        "rect": (0, 0, screen_width, screen_height),
                        "work_rect": (0, 0, screen_width, screen_height),
                        "monitor": None,
                    }
                ]
                self.screen_combo.clear()
                for screen in screens:
                    self.screen_combo.addItem(screen["name"])
                self.screens = screens
                self.screen_combo.setCurrentIndex(0)
            except Exception as e2:
                print(f"基本方案也失败了: {str(e2)}")

    def auto_arrange_windows(self):
        """自动排列窗口"""
        try:
            # 先停止同步
            was_syncing = self.is_syncing
            if was_syncing:
                self.stop_sync()

            # 获取选中的窗口并按编号排序
            selected = []
            for number in self.get_selected_profiles():
                if number in self.browser_processes:
                    hwnd = self.browser_processes[number]
                    selected.append((number, hwnd))

            if not selected:
                QMessageBox.information(self, "提示", "请先选择要排列的窗口！")
                return

            # 按编号正序排序
            selected.sort(key=lambda x: x[0])

            # 获取选中的屏幕信息
            screen_index = self.screen_combo.currentIndex()
            if screen_index < 0 or screen_index >= len(self.screens):
                QMessageBox.critical(self, "错误", "请选择有效的屏幕！")
                return

            screen = self.screens[screen_index]
            screen_rect = screen["work_rect"]  # 使用工作区而不是完整显示区

            # 计算屏幕尺寸
            screen_width = screen_rect[2] - screen_rect[0]
            screen_height = screen_rect[3] - screen_rect[1]

            # 计算最佳布局
            count = len(selected)
            cols = int(math.sqrt(count))
            if cols * cols < count:
                cols += 1
            rows = (count + cols - 1) // cols

            # 计算窗口大小
            width = screen_width // cols
            height = screen_height // rows

            # 创建位置映射（从左到右，从上到下）
            positions = []
            for i in range(count):
                row = i // cols
                col = i % cols
                x = screen_rect[0] + col * width
                y = screen_rect[1] + row * height
                positions.append((x, y))

            # 应用窗口位置
            for i, (_, hwnd) in enumerate(selected):
                try:
                    x, y = positions[i]
                    # 确保窗口可见并移动到指定位置
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    # 先设置窗口样式确保可以移动
                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                    style |= win32con.WS_SIZEBOX | win32con.WS_SYSMENU
                    win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
                    # 移动窗口
                    win32gui.MoveWindow(hwnd, x, y, width, height, True)
                    # 强制重绘
                    win32gui.UpdateWindow(hwnd)
                except Exception as e:
                    print(f"移动窗口 {hwnd} 失败: {str(e)}")
                    continue

            # 如果之前在同步，重新开启同步
            if was_syncing:
                self.start_sync([num for num, _ in selected])

        except Exception as e:
            QMessageBox.critical(self, "错误", f"自动排列失败: {str(e)}")

    def custom_arrange_windows(self):
        """自定义排列窗口"""
        try:
            # 先停止同步
            was_syncing = self.is_syncing
            if was_syncing:
                self.stop_sync()

            selected = []
            for number in self.get_selected_profiles():
                if number in self.browser_processes:
                    hwnd = self.browser_processes[number]
                    selected.append((number, hwnd))

            if not selected:
                QMessageBox.information(self, "提示", "请选择要排列的窗口！")
                return

            try:
                # 获取参数
                start_x = int(self.start_x.text())
                start_y = int(self.start_y.text())
                width = int(self.window_width.text())
                height = int(self.window_height.text())
                h_spacing = int(self.h_spacing.text())
                v_spacing = int(self.v_spacing.text())
                windows_per_row = int(self.windows_per_row.text())
                windows_per_column = int(self.windows_per_column.text())

                # 自动铺满屏幕计算
                if self.auto_fill_screen.isChecked():
                    # 获取屏幕尺寸
                    screen = QApplication.primaryScreen()
                    screen_geometry = screen.availableGeometry()
                    screen_width = screen_geometry.width()
                    screen_height = screen_geometry.height()

                    # 确定实际的窗口行数和列数
                    window_count = len(selected)

                    # 计算总行数和总列数
                    if windows_per_column > 0:
                        total_columns = windows_per_row
                        rows_in_block = windows_per_column

                    else:
                        total_columns = windows_per_row
                        total_rows = (
                            window_count + windows_per_row - 1
                        ) // windows_per_row
                        rows_in_block = total_rows

                    # 计算每个窗口的宽度和高度
                    h_spacing_total = h_spacing * (total_columns - 1)
                    v_spacing_total = v_spacing * (rows_in_block - 1)

                    available_width = screen_width - h_spacing_total - start_x * 2
                    available_height = screen_height - v_spacing_total - start_y * 2

                    # 计算单个窗口的宽度和高度
                    width = max(100, available_width // total_columns)
                    height = max(100, available_height // rows_in_block)

                # 排列窗口
                for i, (_, hwnd) in enumerate(selected):
                    try:
                        if windows_per_column > 0:
                            # 使用每行每列的限制来计算位置
                            # 计算完整的列数
                            cols = windows_per_row
                            # 计算行和列索引
                            major_index = i // (
                                windows_per_row * windows_per_column
                            )  # 大块索引
                            index_in_major = i % (
                                windows_per_row * windows_per_column
                            )  # 块内索引
                            row = index_in_major // windows_per_row  # 块内行
                            col = index_in_major % windows_per_row  # 块内列

                            # 计算实际坐标
                            major_offset_x = major_index * (width + h_spacing) * cols
                            x = start_x + major_offset_x + col * (width + h_spacing)
                            y = start_y + row * (height + v_spacing)
                        else:
                            # 原来的逻辑
                            row = i // windows_per_row
                            col = i % windows_per_row
                            x = start_x + col * (width + h_spacing)
                            y = start_y + row * (height + v_spacing)

                        # 确保窗口可见并移动到指定位置
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        win32gui.MoveWindow(hwnd, x, y, width, height, True)
                    except Exception as e:
                        print(f"移动窗口失败: {str(e)}")

                # 保存参数
                self.save_settings()

            except ValueError:
                QMessageBox.critical(self, "错误", "请输入有效的数字参数！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"排列窗口失败: {str(e)}")

            # 如果之前在同步，重新开启同步
            if was_syncing:
                self.start_sync([num for num, _ in selected])

        except Exception as e:
            QMessageBox.critical(self, "错误", f"排列窗口失败: {str(e)}")


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        icon_path = os.path.join(os.path.dirname(__file__), "app.ico")
        if os.path.exists(icon_path):
            from PyQt6.QtGui import QIcon

            app.setWindowIcon(QIcon(icon_path))
    except Exception as e:
        print(f"设置图标失败: {str(e)}")
    app.setStyle("Fusion")
    window = ChromeManager()
    window.show()
    sys.exit(app.exec())
