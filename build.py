import os
import sys
import shutil
import subprocess
from typing import List


def check_and_install_packages(packages: List[str]):
    """检查并安装所需的包"""
    print("检查并安装必要的包...")

    for package in packages:
        try:
            __import__(package)
            print(f"✓ {package} 已安装")
        except ImportError:
            print(f"正在安装 {package}...")
            subprocess.run(
                [sys.executable, "-m", "pip", "install", package], check=True
            )
            print(f"✓ {package} 安装成功")


def install_requirements():
    """安装所需的依赖包"""
    required_packages = [
        "pyinstaller",  # 用于打包
        "sv_ttk",  # 用于主题
        "keyboard",  # 用于键盘监听
        "mouse",  # 用于鼠标监听
        "pywin32",  # 用于Windows API
        "typing-extensions",  # 用于类型提示
    ]

    for package in required_packages:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except subprocess.CalledProcessError as e:
            print(f"安装 {package} 失败: {str(e)}")
            return False
    return True


def build():
    """打包程序"""
    # 需要的包列表
    required_packages = ["pyinstaller", "sv_ttk", "keyboard", "mouse", "pywin32"]

    # 检查并安装必要的包
    check_and_install_packages(required_packages)

    # 导入需要的模块（在安装后导入）
    import sv_ttk

    print("\n开始打包程序...")

    # 清理旧的构建文件
    if os.path.exists("build"):
        shutil.rmtree("build")
    if os.path.exists("dist"):
        shutil.rmtree("dist")

    # 获取 sv_ttk 路径
    sv_ttk_path = os.path.dirname(sv_ttk.__file__)

    # 创建 spec 文件内容
    spec_content = f"""# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['new_chrome_manager.py'],
    pathex=[],
    binaries=[],
    datas=[('chrome.png', '.'), ('app.ico', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Chrome多开管理器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app.ico'
)
"""

    # 写入 spec 文件
    with open("chrome_manager.spec", "w", encoding="utf-8") as f:
        f.write(spec_content)

    # 创建app.manifest文件
    with open("app.manifest", "w") as f:
        f.write("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
    <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
        <security>
            <requestedPrivileges>
                <requestedExecutionLevel level="requireAdministrator" uiAccess="false"/>
            </requestedPrivileges>
        </security>
    </trustInfo>
    </assembly>""")

    # 运行 PyInstaller
    subprocess.run(["pyinstaller", "chrome_manager.spec"])

    print("\n打包完成！程序文件在 dist 文件夹中。")


if __name__ == "__main__":
    try:
        build()
    except Exception as e:
        print(f"\n错误: {str(e)}")
        input("\n按回车键退出...")
