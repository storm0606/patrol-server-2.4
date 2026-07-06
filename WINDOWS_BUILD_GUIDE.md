# Windows 编译指南

## 环境要求
- Windows 10/11
- Python 3.12（64 位）
- 管理员权限（用于安装依赖）

## 编译步骤

### 1. 安装 Python 3.12
从官网下载：https://www.python.org/downloads/release/python-3120/
安装时勾选 "Add Python to PATH"

### 2. 安装依赖
打开命令提示符（CMD）或 PowerShell：
```cmd
cd patrol-server-2.4-build-source
pip install -r requirements.txt
pip install pyinstaller
```

### 3. 编译 EXE
```cmd
pyinstaller --onedir --name patrol-server-2.4 --icon=patrol.ico --add-data "index.html;." --add-data "data\patrol.db;data" server.py
```

**注意**：`--add-data` 参数中，Windows 使用分号 `;` 分隔源路径和目标路径，Linux/macOS 使用冒号 `:`

### 4. 验证输出
编译完成后，检查 `dist/patrol-server-2.4/` 目录：
```cmd
dir dist\patrol-server-2.4
```

应包含：
- `patrol-server-2.4.exe`（主程序，约 2-3MB）
- `_internal/` 文件夹（依赖库）
- `index.html`
- `data/patrol.db`

### 5. 测试运行
```cmd
cd dist\patrol-server-2.4
patrol-server-2.4.exe
```

浏览器应自动打开 http://localhost:8000

## 常见问题

### Q: 编译时报错 "ModuleNotFoundError"
A: 确保已安装所有依赖：`pip install -r requirements.txt`

### Q: 运行时提示 "找不到 VCRUNTIME140.dll"
A: 安装 Visual C++ Redistributable：https://aka.ms/vs/17/release/vc_redist.x64.exe

### Q: PDF 导出失败
A: 新版本已使用 reportlab 库，无需额外安装软件。如仍有问题，检查是否安装了旧版本。

## 打包发布

编译成功后，可打包为 zip 文件分发：
```cmd
cd dist
tar -czf patrol-server-2.4_release_2026-07-04.tar.gz patrol-server-2.4
```

用户解压后运行 `patrol-server-2.4.exe` 即可。
