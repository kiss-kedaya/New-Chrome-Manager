<div align="center">

# Chrome 多窗口管理器（二创版）

[![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB.svg?style=flat&logo=python&logoColor=white)](https://www.python.org)
[![Windows](https://img.shields.io/badge/Windows-10%2B-0078D6.svg?style=flat&logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![Chrome](https://img.shields.io/badge/Chrome-Latest-4285F4.svg?style=flat&logo=google-chrome&logoColor=white)](https://www.google.com/chrome/)
[![License](https://img.shields.io/badge/License-GPL%20v3-blue.svg)](LICENSE)



  <strong>作者：KEDAYA</strong>  
  <span title="本项目为原版Chrome多窗口管理器的二次创作，原作者：Devilflasher"></span>

</div>

> [!IMPORTANT]
> ## ⚠️ 免责声明
> 
> 1. **本软件为开源项目，仅供学习交流使用，不得用于任何闭源商业用途**
> 2. **使用者应遵守当地法律法规，禁止用于任何非法用途**
> 3. **开发者不对因使用本软件导致的直接/间接损失承担任何责任**
> 4. **使用本软件即表示您已阅读并同意本免责声明**
> 5. **本项目为基于原版[Chrome多窗口管理器](https://github.com/Devilflasher/ChromeManager)的二次开发，遵循GPL-3.0协议，感谢原作者的贡献。**

## 工具介绍
本项目是对原版 Chrome 多窗口管理器的二次开发，保留并优化了多窗口批量管理、智能布局、同步控制等核心功能，适用于需要高效管理多个Chrome实例的用户。

## 功能特性

- `批量管理功能`：一键打开/关闭单个、多个Chrome实例
- `智能布局系统`：支持自动网格排列和自定义坐标布局
- `多窗口同步控制`：实时同步鼠标/键盘操作到所有选定窗口
- `批量打开网页`：支持批量相同网页打开
- `快捷方式图标替换`：支持一键替换多个快捷方式图标（项目根目录下有app.ico、chrome.png、chrome.svg等图标文件）
- `插件窗口同步`：支持弹出的插件窗口内的键盘和鼠标同步

## 环境要求

- Windows 10/11 (64-bit)
- Python 3.9+
- Chrome浏览器 最新

## 运行教程
### 方法一：打包成独立exe可执行文件（推荐）

1. **安装 Python 和依赖**
   ```bash
   # 安装 Python 3.9 或更高版本
   # 从 https://www.python.org/downloads/ 下载
   ```

2. **准备文件**
   - 确保目录里有以下文件：
     - new_chrome_manager.py（主程序）
     - build.py（打包脚本）
     - app.manifest（管理员权限配置）
     - app.ico（程序图标）
     - requirements.txt（依赖包列表）

3. **运行打包脚本**
   ```bash
   # 在程序目录下运行：
   python build.py
   ```

4. **查找生成文件**
   - 打包完成后，在 `dist` 目录下找到 `Chrome多开管理器.exe`
   - 双击运行即可打开程序

### 方法二：从源码运行

1. **安装 Python**
   ```bash
   # 下载并安装 Python 3.9 或更高版本
   # 从 https://www.python.org/downloads/ 下载
   ```

2. **安装依赖包**
   ```bash
   # 打开命令提示符（CMD）并运行：
   pip install -r requirements.txt
   ```

3. **运行程序**
   ```bash
   # 在程序目录下运行：
   python new_chrome_manager.py
   ```

## 使用说明

### 前期准备

- 在您存放 Chrome 多开快捷方式的文件夹下，快捷方式的文件名应按照 `1`、`2`、`3`... 的格式命名。
- 同一个文件夹下建立 `Data` 文件夹，`Data` 文件夹下存放每个浏览器独立的数据文件，文件夹名应按照 `1`、`2`、`3`... 的格式命名。

```目录结构示例：
                                 多开chrome的目录

                                ├── 1.lnk
                                ├── 2.lnk
                                ├── 3.lnk
                                └── Data
                                    ├── 1
                                    ├── 2
                                    └── 3
```
- 浏览器快捷方式的目标参数示例：（请根据您的浏览器安装路径修改）
```
"C:\Program Files\Google\Chrome\Application\chrome.exe" --user-data-dir="D:\chrom duo\Data\编号"
```

### 基本操作

1. **打开窗口**
   - 软件下方的"打开窗口"标签下，填入存放浏览器快捷方式的目录
   - "窗口编号"里填入想要打开的浏览器编号
   - 点击"打开窗口"按钮即可打开对应编号的chrome窗口

2. **导入窗口**
   - 点击"导入窗口"按钮导入当前打开的 Chrome 窗口
   - 在列表中选择要操作的窗口

3. **窗口排列**
   - 使用"自动排列"快速整理窗口
   - 或使用"自定义排列"设置详细的排列参数

4. **开启同步**
   - 选择一个主控窗口（点击"主控"列）
   - 选择要同步的从属窗口
   - 点击"开始同步"或使用设定的快捷键

## 注意事项

- 同步功能需要管理员权限
- 批量操作时注意系统资源占用
- 若遇杀毒软件拦截，请将本程序加入信任

## 常见问题

1. **无法开启同步**
   - win10和win11家庭版操作系统目前有兼容性上的问题
   - 如果你是chrome多账户多开，那无法使用同步功能
   - 如果仅是鼠标或者键盘无法同步，可能是打包的时候某个依赖出错，请重新打包程序文件

2. **窗口未正确导入**
   - 尝试重新点击"导入窗口"按钮

3. **滚动条同步幅度不同**
   - 目前的解决办法就是通过pageup和pagedown以及键盘上下左右键来调整同步幅度

## 更新日志

### v1.0（二创版）
- 基于原版功能，优化体验，修复部分兼容性问题
- 保留并增强窗口管理和同步功能

## 许可证

本项目采用 GPL-3.0 License，保留所有权利。使用本代码需明确标注来源，禁止闭源商业使用。

> 本项目为二次创作，原版项目地址：https://github.com/Devilflasher/ChromeManager

🔄 持续更新中

