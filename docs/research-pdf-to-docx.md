# WPS 导出 PDF 转可编辑 DOCX：开源技术调研

> 调研日期：2026-09-11  
> 范围：Windows 离线桌面程序；输入是 WPS 导出的 `.pdf`，输出是可编辑 `.docx`。  
> 证据边界：只使用项目官方文档、官方代码库和官方许可证页面。本文不构成法律意见。

## 一、结论

**建议做“分流式”而不是只接一个转换库：**

1. **文本型 PDF 快速路径**：保留 PDF 原生文本与坐标，用版面规则重建 DOCX；`pdf2docx` 是最接近可直接复用的参考实现。它的官方技术文档明确是“PyMuPDF 取文本/图片/图形及坐标 → 按相对位置解析版面和表格 → python-docx 重建”。[官方技术文档](https://pdf2docx.readthedocs.io/en/latest/techdoc.html)
2. **扫描件/乱码回退路径**：优先评估 PaddleOCR `PP-StructureV3`，因为它同时做文档版面解析、OCR 和结构化输出，官方 API 可直接 `save_to_word()` 生成 `.docx`。[官方 PP-StructureV3 教程](https://www.paddleocr.ai/latest/en/version3.x/pipeline_usage/PP-StructureV3.html)
3. **不建议把 PyMuPDF4LLM 作为主转换器**：它的官方输出是 Markdown、JSON 和纯文本，定位是 RAG/LLM 数据提取，没有 DOCX 输出；可作版面 JSON 或自动发现乱码的参考，不适合直接追求 Word 版式还原。[官方仓库](https://github.com/pymupdf/pymupdf4llm)
4. **OCRmyPDF 不是 PDF→DOCX 引擎**：它的作用是向扫描 PDF 增加可搜索文本层，然后仍需要另一个布局/DOCX 重建阶段。[官方介绍](https://ocrmypdf.readthedocs.io/en/stable/introduction.html)
5. **如果要发行闭源免费版或商业版，PyMuPDF 路径先有许可问题**：`pdf2docx` 自身已改为 MIT，但它当前明确依赖 `PyMuPDF>=1.26.7`，而 PyMuPDF 官方仓库是 AGPL-3.0（另有商业授权）。不能因为上层项目是 MIT，就忽略被打包依赖的许可义务。[当前 requirements.txt](https://github.com/ArtifexSoftware/pdf2docx/blob/master/requirements.txt) · [PyMuPDF 官方仓库/许可](https://github.com/pymupdf/PyMuPDF)

**综合推荐：**把产品做成“原生文本提取 + 选择性 OCR + 表格结构重建 + 生成后检查”的流水线。技术验证阶段可快速用 `pdf2docx` 建立质量基线，但如产品必须闭源发行，应在立项时就二选一：购买 PyMuPDF/Artifex 商业授权，或将生产线替换为可兼容闭源分发的 PDF 解析组件；OCR 与复杂表格回退优先选 Apache-2.0 的 PaddleOCR/PaddlePaddle。

## 二、为什么必须先判断 PDF 类型

PDF 是页面描述格式，“看起来是表格”并不意味内部有 Word/Excel 式表格对象；PyMuPDF 官方文档明确说明，表格通常是普通文本和线条排成表格的视觉效果，结构识别可能很复杂。[官方文本/表格提取指南](https://pymupdf.readthedocs.io/en/latest/recipes-text.html)

| 输入类型 | 可观测特征 | 首选处理 | 主要风险 |
|---|---|---|---|
| WPS 正常导出、文本可选且复制正常 | 页面有字符层和坐标 | 保留原生文本，只做版面和表格推断 | 阅读顺序、分栏、表格边线、字体替代 |
| 有文本层，但复制出乱码/缺字 | 字形能显示，却缺失正确 Unicode 映射 | 按页或按区域 OCR，不覆盖其他正常文本 | OCR 误字、标点/数字混淆、格式丢失 |
| 纯扫描件/页面大图 | 无可选文本 | 全页 OCR + 版面分析 + 表格模型 | 速度、模型体积、低清扫描的准确率 |
| 混合页 | 数字文本、截图、扫描签章共存 | 区域级选择性 OCR | 重复文本、层叠错位 |

PyMuPDF 官方 FAQ 将乱码常见原因归为 PDF 使用自定义字体编码但没有正确字符映射（CMAP），字形可显示但无法反推回 Unicode，官方建议此时用 OCR 回退。[官方仓库 FAQ](https://github.com/pymupdf/PyMuPDF#text-extraction-returns-garbled-characters-or-empty-output-why)。因此，**中文乱码不一定是输出 Word 字体没选对**；如 PDF 本身缺少映射，仅将 Word 字体改为宋体或微软雅黑不能恢复原文。

## 三、候选方案核实

### 3.1 pdf2docx：文本型 PDF 的直接基线

- **能力**：官方文档的处理链就是 PDF 元素提取、规则版面解析、python-docx 重建，与“可编辑 Word”目标直接对应。[官方技术文档](https://pdf2docx.readthedocs.io/en/latest/techdoc.html)
- **适用性**：适合可正常提取文本的 WPS 导出 PDF，可作 MVP 与质量比较基线；不应把它当作扫描件 OCR 引擎。官方架构只描述提取 PDF 现有元素，且该项目的历史功能说明把 OCR 列为 TODO；实际支持边界必须用样例回归测试锁定。[官方技术文档](https://pdf2docx.readthedocs.io/en/latest/techdoc.html)
- **维护风险**：Artifex 已在官方仓库声明不再主动维护，只保留仓库并接受社区 PR。选它意味着要准备自己修 bug、锁定版本和维护分支。[官方仓库状态](https://github.com/ArtifexSoftware/pdf2docx)
- **许可风险**：`pdf2docx` 仓库已是 MIT，但它的当前依赖列表包含 PyMuPDF；PyMuPDF 为 AGPL-3.0/商业双许可。对“打包成 Windows EXE 并分发”的闭源产品，必须在发行前完成许可评估或购买商业授权。[官方 requirements.txt](https://github.com/ArtifexSoftware/pdf2docx/blob/master/requirements.txt) · [PyMuPDF COPYING](https://github.com/pymupdf/PyMuPDF/blob/main/COPYING)

### 3.2 PyMuPDF：低层提取和识别辅助，不是 DOCX 生成器

- **能力**：可提取文本、字体、颜色、坐标、图片和矢量线条，`Page.find_tables()` 可返回表格单元格边界框与内容；但它官方列出的输出不含 DOCX，仍需要 python-docx 或其他 OOXML 生成层。[官方仓库能力列表](https://github.com/pymupdf/PyMuPDF) · [`Page.find_tables()` 官方 API](https://pymupdf.readthedocs.io/en/latest/page.html#Page.find_tables)
- **阅读顺序**：PDF 文本的内部存储顺序可与视觉顺序不同，官方文档提供 `sort` 和布局保持提取来改善，但这不等于能对所有多栏/浮动对象无损还原。[官方文本提取指南](https://pymupdf.readthedocs.io/en/latest/recipes-text.html)
- **表格限制**：`find_tables()` 依据线条/矩形或文本策略推断单元格；对无边框、背景色代替边线、非常规结构，官方 FAQ 明确提示可能检测失败或需要自定义空间逻辑。[官方 FAQ](https://pymupdf.readthedocs.io/en/latest/faq/index.html#table-extraction)
- **离线性**：PyMuPDF 官方说明开源构建与商业构建均可在本地/隔离网络环境执行；OCR 时需可用的 Tesseract `tessdata`。[官方仓库 FAQ](https://github.com/pymupdf/PyMuPDF#can-i-use-pymupdf-pymupdf4llm-and-pymupdf-pro-without-sending-data-to-the-cloud)

### 3.3 PyMuPDF4LLM：适合结构化中间结果，不适合作终端转换器

- **能力**：官方仓库列出多栏阅读顺序、表格识别、图片、页眉页脚处理和选择性 OCR，并可输出带 bounding box/版面信息的 JSON。[官方仓库](https://github.com/pymupdf/pymupdf4llm)
- **不适用点**：官方只列出 Markdown、JSON 和纯文本输出；Markdown 表格不能完整表示 Word 中的页面几何、浮动对象、行高/列宽和复杂跨行跨列。即使先转 Markdown 再转 DOCX，也更像内容迁移，不是版式还原。[官方输出格式列表](https://github.com/pymupdf/pymupdf4llm#output-formats)
- **许可**：官方仓库为 AGPL-3.0，并依赖同样是 AGPL/商业双许可的 PyMuPDF；与闭源 EXE 绑定时有同类许可问题。[官方 LICENSE](https://github.com/pymupdf/pymupdf4llm/blob/main/LICENSE) · [PyMuPDF COPYING](https://github.com/pymupdf/PyMuPDF/blob/main/COPYING)

### 3.4 OCRmyPDF + Tesseract：可搜索 PDF 预处理链

- **定位**：OCRmyPDF 把 OCR 文本层加到扫描 PDF，使内容可搜索/复制；它不直接生成 DOCX，也不负责把单元格重建成 Word 表格。[官方介绍](https://ocrmypdf.readthedocs.io/en/stable/introduction.html)
- **中文**：官方文档支持 Tesseract 简体中文 `chi_sim`，混合语言可指定多个语言；Windows 上需将额外 `.traineddata` 放入 `tessdata` 目录。[官方语言包文档](https://ocrmypdf.readthedocs.io/en/stable/languages.html)
- **Tesseract 能力与许可**：Tesseract 官方声明支持 UTF-8 与 100 多种语言，代码为 Apache-2.0；官方 `tessdata_best` 仓库提供 `chi_sim`、`chi_tra` 及竖排模型。[官方 Tesseract 仓库](https://github.com/tesseract-ocr/tesseract) · [官方 tessdata_best](https://github.com/tesseract-ocr/tessdata_best)
- **Windows 分发复杂度**：OCRmyPDF 17.11 的官方 Windows 章节仍列出 64 位 Python、Tesseract、Ghostscript 且无单命令安装；但同一页的新版依赖说明又称 17.0 起 Ghostscript 对某些路径可选，可用 `pypdfium2` 栅格化，而 PDF/A 转换仍可需 Ghostscript 或 veraPDF。这两段官方说明存在不一致，不应在未做 Windows 打包 POC 前宣称“零外部依赖”。[官方安装/依赖文档](https://ocrmypdf.readthedocs.io/en/stable/installation.html)
- **许可**：OCRmyPDF 当前项目元数据是 MPL-2.0，Tesseract 是 Apache-2.0；但发行包还要逐项核对 pypdfium2/PDFium、pikepdf/qpdf、字体和可选 Ghostscript 等依赖的许可与 notice。[官方 OCRmyPDF pyproject.toml](https://github.com/ocrmypdf/OCRmyPDF/blob/main/pyproject.toml) · [官方 Tesseract LICENSE 说明](https://github.com/tesseract-ocr/tesseract#license)

### 3.5 PaddleOCR / PP-StructureV3：扫描件与复杂版面的主要回退候选

- **能力**：PP-StructureV3 官方教程提供文档版面结构化预测，结果可保存为 JSON、Markdown、Word `.docx`、表格 HTML 和 XLSX。相比“先做纯 OCR，再猜表格”，它的官方管线就包含版面与表格结构。[官方 PP-StructureV3 教程](https://www.paddleocr.ai/latest/en/version3.x/pipeline_usage/PP-StructureV3.html)
- **中文**：官方语言表明确列出简体中文 `ch` 和繁体中文 `chinese_cht`，更适合把中文表格作为一等用例。[官方 PP-StructureV3 语言表](https://www.paddleocr.ai/latest/en/version3.x/pipeline_usage/PP-StructureV3.html#supported-languages)
- **合并单元格**：PaddleOCR 的结构化输出为合并结构提供了比纯 OCR 文字框更合适的中间层，但“能输出 Word/HTML 表格”不等于“任意合并单元格都能正确还原”。行列跨度必须在用户真实样例上单独统计，错误时保留可视化编辑/人工修正入口。[官方结果保存 API](https://www.paddleocr.ai/latest/en/version3.x/pipeline_usage/PP-StructureV3.html#python-integration)
- **离线模型**：官方支持导出 PaddleX YAML 配置，用 `model_dir` 指向本地模型权重。离线安装包应在构建时预下载并固定所需模型，安装后统一指向本地目录，不让首次转换隐式访问网络。[官方本地模型配置](https://www.paddleocr.ai/latest/en/version3.x/pipeline_usage/PP-StructureV3.html#model-deployment)
- **Windows 打包**：PaddleOCR 官方已提供 PyInstaller 打包指南，需收集 PaddleX 数据、Paddle 二进制和包元数据；官方测试环境是 Windows 11，并明确表示当前不支持 Nuitka。[官方 PyInstaller 打包指南](https://github.com/PaddlePaddle/PaddleOCR/blob/main/docs/version3.x/inference_deployment/others/packaging.en.md)
- **体积/性能代价**：必须携带 Paddle 运行时、多个版面/OCR/表格模型及二进制，安装包和内存会明显大于纯规则快速路径。这是从官方打包命令需收集 `paddlex` 数据和 `paddle` 二进制所作的工程推断，不是官方性能承诺。[官方打包指南](https://github.com/PaddlePaddle/PaddleOCR/blob/main/docs/version3.x/inference_deployment/others/packaging.en.md)
- **许可**：PaddleOCR 和 PaddlePaddle 官方代码库均为 Apache-2.0，对闭源桌面软件通常比 AGPL 路径容易管理；仍要为预训练模型、间接依赖和字体逐项保留许可证/NOTICE。[官方 PaddleOCR LICENSE](https://github.com/PaddlePaddle/PaddleOCR/blob/main/LICENSE) · [PaddlePaddle 官方仓库](https://github.com/PaddlePaddle/Paddle)

## 四、表格与合并单元格的正确重建方式

1. **不能只把 OCR 文字用空格对齐。** 应保留每个单元格的矩形、行列网格、文字框、边线和置信度，再映射成 DOCX 表格。PDF 本身常常只有普通文本和线条，因此表格是推断结果。[官方 PyMuPDF 表格说明](https://pymupdf.readthedocs.io/en/latest/recipes-text.html#how-to-extract-table-content-from-documents)
2. **原生 PDF 表格走矢量路径。** 优先使用 PDF 的线条/矩形和原生文本坐标，不必先栅格化为图像；PyMuPDF `find_tables()` 就是这种实现，且可用 `strategy="text"` 处理某些无可见边线情况。[官方 FAQ](https://pymupdf.readthedocs.io/en/latest/faq/index.html#table-extraction)
3. **扫描表格走结构模型路径。** 让 PP-StructureV3 输出表格 HTML/结构化 JSON，再转为 DOCX，同时保留原始页面图用于后续对照。[官方输出 API](https://www.paddleocr.ai/latest/en/version3.x/pipeline_usage/PP-StructureV3.html#python-integration)
4. **DOCX 层能表示水平和垂直合并。** `python-docx` 官方文档的 `cell.merge()` 以对角单元格指定矩形合并区，底层对应 WordprocessingML 的 `gridSpan` 和 `vMerge`。[官方合并单元格文档](https://python-docx.readthedocs.io/en/latest/dev/analysis/features/table/cell-merge.html)
5. **难点在“识别合并区”，不在“写入 DOCX”。** 所以验收指标应分开：行列数、水平跨列、垂直跨行、单元格文字、表头层级和边框样式，不能只看“Word 打开后像不像”。这是根据 PDF 表格为视觉推断、DOCX 合并为显式结构的差异得出的工程结论。[官方 PyMuPDF 说明](https://pymupdf.readthedocs.io/en/latest/recipes-text.html#how-to-extract-table-content-from-documents) · [官方 python-docx 合并说明](https://python-docx.readthedocs.io/en/latest/dev/analysis/features/table/cell-merge.html)

## 五、建议的离线桌面架构（仅调研结论，不是实现）

```text
用户选 PDF
    ↓
按页检测：原生文本量 / Unicode 可用性 / 图像覆盖率 / 表格候选区
    ├─ 文本正常 → 原生文本+坐标+矢量线条重建
    ├─ 局部乱码/图像文字 → 只 OCR 有问题区域
    └─ 纯扫描/复杂表格 → PP-StructureV3 全页结构化
    ↓
统一中间模型：页/区块/段落/图片/表格/单元格跨度/置信度
    ↓
python-docx/OOXML 生成 DOCX
    ↓
自动检查：乱码、缺页、表格行列/合并异常、低置信页
    ↓
用户在 Word/WPS 中人工校对标记页
```

- **选择性 OCR 是关键**：PyMuPDF4LLM 官方也采用类似思路，仅对图像覆盖或不可读区域 OCR，避免把原本正确的数字文本再识别成错字。[官方混合 OCR 说明](https://github.com/pymupdf/pymupdf4llm#hybrid-ocr-strategy)
- **离线要在安装包层面完整闭环**：不仅 Python 代码要离线，OCR 语言包、Paddle 模型、DLL、字体策略、许可证和 NOTICE 都要在构建阶段冻结。PaddleOCR 支持本地 `model_dir`，Tesseract 则从本地 `tessdata` 读语言包。[官方 PaddleOCR 模型配置](https://www.paddleocr.ai/latest/en/version3.x/pipeline_usage/PP-StructureV3.html#model-deployment) · [官方 OCRmyPDF Windows 语言包说明](https://ocrmypdf.readthedocs.io/en/stable/languages.html#windows)
- **打包器本身不解决依赖许可**：PyInstaller 的官方例外允许商业应用打包和任意许可的输出，前提仍是遵守被打包依赖的许可；因此它不会“洗掉” PyMuPDF AGPL 或其他组件义务。[官方 PyInstaller 许可说明](https://pyinstaller.org/en/stable/license.html)
- **分发形态建议**：PaddleOCR 版本优先做 `onedir` 式安装目录再用标准 Windows 安装器封装，而非一开始追求单文件 EXE；官方打包程序本身就需要收集大量数据、二进制和元数据。这是基于官方打包流程的工程建议。[官方 PaddleOCR 打包指南](https://github.com/PaddlePaddle/PaddleOCR/blob/main/docs/version3.x/inference_deployment/others/packaging.en.md)

## 六、许可与分发决策表

| 组件 | 官方许可/状态 | 对 Windows 闭源分发的结论 |
|---|---|---|
| pdf2docx | MIT；Artifex 不再主动维护 | 上层可用，但当前依赖 PyMuPDF，不能只看 MIT。[官方仓库](https://github.com/ArtifexSoftware/pdf2docx) |
| PyMuPDF | AGPL-3.0，另有商业授权 | 闭源发行前必须做专项许可决策；最直接的商业路径是购授权。[官方 COPYING](https://github.com/pymupdf/PyMuPDF/blob/main/COPYING) |
| PyMuPDF4LLM | AGPL-3.0，且依赖 PyMuPDF | 与 PyMuPDF 同类；又因无 DOCX 直出，不建议为主路径。[官方 LICENSE](https://github.com/pymupdf/pymupdf4llm/blob/main/LICENSE) |
| python-docx | MIT | 适合作 DOCX 生成层，支持显式单元格合并。[官方仓库](https://github.com/python-openxml/python-docx) · [官方 LICENSE](https://github.com/python-openxml/python-docx/blob/master/LICENSE) |
| OCRmyPDF | MPL-2.0 | 可作独立预处理组件，但不输出 DOCX；修改它自身文件与分发时需按 MPL 要求处理，并审核间接依赖。[官方 pyproject.toml](https://github.com/ocrmypdf/OCRmyPDF/blob/main/pyproject.toml) |
| Tesseract / 官方语言数据 | Apache-2.0 | 可作离线 OCR，分发时保留许可/NOTICE 并同时审核 Leptonica 等依赖。[官方仓库](https://github.com/tesseract-ocr/tesseract) · [官方 tessdata_best](https://github.com/tesseract-ocr/tessdata_best) |
| PaddleOCR / PaddlePaddle | Apache-2.0 | 作闭源离线版的 OCR/结构化主要候选；仍需管理模型与所有间接依赖清单。[官方 PaddleOCR LICENSE](https://github.com/PaddlePaddle/PaddleOCR/blob/main/LICENSE) · [PaddlePaddle 官方仓库](https://github.com/PaddlePaddle/Paddle) |
| PyInstaller | GPL-2.0 + 商业打包例外，少量文件 Apache-2.0 | 可打包商业应用，但输出仍必须遵守每个依赖的许可。[官方许可页](https://pyinstaller.org/en/stable/license.html) |

## 七、建议的 POC 验收集

不应用一份简单 PDF 宣布方案成功。建议在开发前先准备 20–50 页可人工核对的真实样例，至少覆盖：

1. WPS 直接导出的简体中文 PDF，复制文本正常。
2. 视觉正常但复制乱码的中文子集字体 PDF。
3. 150/300/600 DPI 扫描页，包含中英数字混排。
4. 无边框表格、细线表格、背景色表格。
5. 水平合并、垂直合并、同时跨行跨列、多层表头。
6. 长表格跨页、重复表头、表格中换行与图片。
7. 多栏、页眉页脚、页码、浮动图片、章印。
8. 混合页：原生文本 + 扫描图片 + 乱码字体区域。

至少记录以下指标：字符错误率（中文/数字/英文分开）、段落顺序、表格行列正确率、合并区正确率、缺图/错图数、每页耗时、峰值内存、安装包大小和断网首次启动是否成功。

## 八、需要在开发前确定的产品决策

- 软件是完全开源/AGPL 发行，还是闭源免费/商业发行？这会直接决定能否无商业授权地绑定 PyMuPDF。
- “可编辑”是否优先于“像原 PDF”？复杂页面中，大量文本框能更像原稿，真正的段落/表格则更好编辑，两者不能总是同时最优。
- 是否接受数百 MB 甚至更大的离线模型包？如不接受，就必须收缩 OCR/版面模型范围或改用可选下载组件。
- 目标是 Windows 10 还是 Windows 11，x64 还需 x86/ARM64？PaddleOCR 官方 PyInstaller 证据只明确给出了 Win 11 的一组测试环境，其他目标需自己验证。[官方打包指南](https://github.com/PaddlePaddle/PaddleOCR/blob/main/docs/version3.x/inference_deployment/others/packaging.en.md)
- 合并单元格识别失败时，是否允许用户在导出前修正网格？如必须全自动，就要接受某些文档无法保证的事实并设置低置信提示。

## 九、最终选型建议

| 场景 | 建议 |
|---|---|
| 尽快做可运行的内部 POC，尚不对外分发 | 用 `pdf2docx` 跑文本型基线，同时用 PP-StructureV3 跑扫描/乱码/表格样例，按页比较。 |
| 闭源、免费或商业分发，不买 Artifex 授权 | 不把 PyMuPDF/pdf2docx/PyMuPDF4LLM 放入生产发行包；选可兼容闭源的 PDF 解析层 + PaddleOCR/PP-StructureV3 + python-docx，再做全量依赖许可审计。 |
| 闭源分发，愿意买 Artifex 授权 | 可用 PyMuPDF 做高性能原生文本/坐标/表格快速路径，PaddleOCR 做回退，python-docx 统一输出。 |
| 整个软件可按 AGPL 合规开源 | 可直接基于 pdf2docx/PyMuPDF 演进，但仍要承担 pdf2docx 已不再由 Artifex 主动维护的工程成本。 |

**不可诚实承诺的目标：**“所有 PDF 一键 100% 还原为可编辑 Word，中文、合并单元格、图文浮动位置均不出错”。可实现的产品目标应是：对目标样本集达到明确量化准确率，对低置信页显式标记，并允许用户在 WPS/Word 中继续编辑修正。
