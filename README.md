# PDF 转可编辑 Word

一个面向 Windows 的本地桌面工具：将 PDF（包括从 WPS 导出的 PDF）转换为可编辑的 `.docx`。

它采用 Windows 自带的中文 OCR 识别文字，并通过 PDF 表格线重建 Word 表格与合并单元格。不会上传文档。

## 当前能力

- 拖入或选择多个 PDF，依次转换成 DOCX
- 支持简体中文 Windows OCR（系统已安装中文 OCR 语言包时）
- 识别带边框的表格，并还原大多数横向、纵向合并单元格
- 每个 PDF 单独生成一个 DOCX；保留原文件，不覆盖输入
- 可在软件中选择输出文件夹、查看进度和打开输出位置

## 已知边界

- 输出是可编辑内容，不是原 PDF 截图，因此字距、行距和跨页位置可能有差异。
- 没有表格线的复杂版式、印章、公式、手写文字和低清扫描件需要人工复核。
- Windows OCR 语言由系统提供。本机缺少 `zh-Hans-CN` 时，程序会说明如何安装，而不会产生乱码文档。

## 运行

在 PowerShell 中执行：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\setup.ps1
.\run.ps1
```

首次运行会在项目目录生成 `.venv` 并安装 Python 依赖。源代码入口为 `app\main.py`。

## 打包 EXE

```powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller
.\.venv\Scripts\pyinstaller.exe --noconfirm --windowed --name PDF转Word --add-data "app\windows_ocr.ps1;." --collect-all pypdfium2 --collect-all PIL app\main.py
```

生成文件位于 `dist\PDF转Word\PDF转Word.exe`。请在目标机器安装中文 OCR 语言包后再使用。

## 版本与调研

- 本项目使用 Git 保存每次本地修改。提交前可用 `git status` 查看变更、用 `git log --oneline` 查看历史。
- `docs/research-pdf-to-docx.md` 记录了 PDF 转 DOCX 技术路线及许可证取舍。
- `docs/reference-repository.md` 记录了用户提供参考仓库的当前状态；该仓库无源代码，未被复制进本项目。
