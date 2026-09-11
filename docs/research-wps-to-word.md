# Windows 上将 `.wps` 转为 Word `.docx` 的可行性调研

> 调研日期：2026-09-11  
> 范围：仅调研，不实现软件。资料优先采用产品官方文档、上游官方代码库和项目自己的代码库。

## 1. 结论先行

这个软件可以做，但在写代码前必须先确认用户所说的 `.wps` 究竟是哪一种文件。`.wps` 至少有两种不同含义：

1. **Microsoft Works Word Processor 文档**。微软的格式参考明确把 `.wps` 定义为 Works 6–9 的默认文档格式；LibreOffice 当前源码也把扩展名 `.wps` 注册为 `Microsoft Works Document`，MIME 类型为 `application/vnd.ms-works`。[微软格式参考](https://learn.microsoft.com/en-us/office/compatibility/office-file-format-reference) · [LibreOffice 类型定义](https://github.com/LibreOffice/core/blob/master/filter/source/config/fragments/types/writer_MS_Works_Document.xcu)
2. **金山 WPS Office 自身的 WPS 文字文档**。WPS 官方技术规格把 `.wps` 列为 Writer 支持格式，官方界面命令表也分别列出了“WPS 文字文件（*.wps）”和“Word 文件（*.docx）”。这和 Microsoft Works 不是同一个可由扩展名唯一判断的格式概念。[WPS 技术规格](https://www.wps.com/tech-specs/) · [WPS 文字 idMso 参考](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/idmso-list/wps-idmso-reference)

因此推荐采用“双后端、先识别后转换”的产品设计：

- 对经内容检测确认的 **Microsoft Works `.wps`**，首选调用本机 LibreOffice 的无界面转换能力，导出 `.docx`。
- 对 **金山 WPS Office `.wps`**，优先调用用户已经合法安装的 WPS Office 自动化接口打开并 `SaveAs2` 为 `.docx`；不能仅凭文件扩展名交给 LibreOffice。
- 第一版若希望完全开源、部署简单，应明确只承诺“Microsoft Works `.wps` → `.docx`”。若要承诺“WPS Office `.wps` → `.docx`”，必须拿到真实样本做兼容性验证，并处理 WPS 客户端版本、授权和自动化部署条件。

## 2. LibreOffice headless / UNO

### 2.1 已确认的能力

LibreOffice 官方帮助文档提供 Windows 命令行入口 `soffice.com`，并正式支持：

- `--headless`：无界面运行，亦可由外部客户端通过 API 控制；
- `--convert-to OutputFileExtension[:OutputFilterName] [--outdir ...]`：按目标扩展名和导出过滤器转换；
- `--infilter=...`：必要时强制输入过滤器；
- `-env:UserInstallation=file:///...`：指定独立用户配置目录。

来源：[LibreOffice 官方命令行参数](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html?DbPAR=BASIC&System=WIN)

对 Microsoft Works 的支持不是只靠扩展名猜测。LibreOffice 上游源码中：

- `.wps` 类型绑定到 `com.sun.star.comp.Writer.MSWorksImportFilter`，名称为 `Microsoft Works Document`；
- 过滤器标记为只导入的第三方过滤器，并使用 `MSWorksImportFilter`；
- 实现直接包含 `libwps/libwps.h`，先用 `libwps::WPSDocument::isFileFormatSupported` 检测文件内容，再解析为 Writer 文档模型；
- 源码特别处理了 headless 模式下老 DOS 文档可能需要的字符编码参数。

来源：[LibreOffice `.wps` 类型配置](https://github.com/LibreOffice/core/blob/master/filter/source/config/fragments/types/writer_MS_Works_Document.xcu) · [LibreOffice Works 过滤器配置](https://github.com/LibreOffice/core/blob/master/filter/source/config/fragments/filters/MS_Works.xcu) · [LibreOffice `MSWorksImportFilter` 实现](https://github.com/LibreOffice/core/blob/master/writerperfect/source/writer/MSWorksImportFilter.cxx)

UNO 适合需要常驻转换进程、精细控制加载/保存属性或批量队列的版本。官方 `XStorable` API 区分 `storeAsURL` 和 `storeToURL`；纯导出过滤器应使用 `storeToURL`。对于普通桌面批量转换器，先调用 `soffice.com --headless --convert-to ...` 更简单，UNO 可放到后续优化阶段。[LibreOffice `XStorable` 接口](https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XStorable.html)

### 2.2 推荐调用形态

对已确认是 Microsoft Works 文本文档的输入，可按下列形态调用（这是基于官方参数和上游过滤器名组合出的实现建议，实际参数应由样本测试锁定）：

```text
soffice.com --headless --nologo --norestore \
  -env:UserInstallation=file:///C:/.../unique-profile \
  --convert-to "docx:Office Open XML Text" \
  --outdir "C:\...\output" "C:\...\input.wps"
```

每个并发任务应使用独立的 `UserInstallation` 目录，避免复用用户正在运行的 LibreOffice 实例或互相抢占配置锁。程序不能只相信进程退出码，还应确认目标文件存在、非空，并能作为 OOXML ZIP 包正常打开。

### 2.3 能力边界

- LibreOffice 的官方 `.wps` 过滤器明确命名为 **Microsoft Works**，不能据此推断它支持所有由金山 WPS Office 保存的 `.wps` 文件。
- Works 导入依赖逆向实现，复杂排版、嵌入对象、旧字符编码、缺失字体和宏都可能丢失或变化。`libwps` 官方项目也把支持表述为“retrieve almost all content”等非完全保证。[libwps 官方支持范围](https://sourceforge.net/p/libwps/wiki/Home/)
- 密码保护或需要交互选择编码的文件不适合静默批处理；首版应返回明确的“需要人工处理/不支持”状态，而不是生成疑似成功的空文档。

## 3. libwps

`libwps` 是专门用于 **Microsoft Works** 格式的 C++ 导入库，建立在 `librevenge` 之上。官方页面称其可导入约 1995 年以来的 Works 文字处理格式，并附带 `wps2html`、`wps2odt` 等控制台转换器。[libwps 官方项目页](https://sourceforge.net/projects/libwps/) · [libwps 官方 Wiki](https://sourceforge.net/p/libwps/wiki/Home/) · [libwps 官方源码](https://sourceforge.net/p/libwps/code/ci/master/tree/)

适用性判断：

- **适合**：把它作为 Microsoft Works 文件识别和导入的底层能力；LibreOffice 已经完成了这种集成。
- **不适合直接单独完成目标**：官方自带工具面向 HTML/ODT，并没有直接承诺输出 DOCX。若单独嵌入 libwps，仍需自己构建 DOCX 写出层或再经过 LibreOffice，工程量和格式风险都更高。
- **不应宣称支持金山 WPS 原生格式**：官方项目范围写的是 Microsoft Works，而不是 WPS Office。

许可证方面，官方项目同时列出 LGPLv2 与 MPL 2.0，源码树也包含两份许可证文本。若只是调用未修改的 LibreOffice 可执行文件，自己的 GUI/调度代码可保持独立许可证；若直接链接、修改或重新分发 libwps，则必须按所选择的许可证路径履行对应的源代码、声明及动态/静态链接义务。[libwps 项目许可证](https://sourceforge.net/projects/libwps/) · [libwps 源码许可证文件](https://sourceforge.net/p/libwps/code/ci/master/tree/)

## 4. Pandoc

Pandoc **不适合作为 `.wps` 的直接转换引擎**。官方用户指南列出的输入格式包含 DOCX、ODT、RTF 等，但没有 `.wps`/Microsoft Works；DOCX 虽然是支持的输出格式，但没有输入 reader 就无法直接转换。此外，Pandoc 官方明确提醒其中间文档模型比很多办公格式表达能力弱，复杂表格和页边距等细节可能丢失。[Pandoc 官方用户指南：输入/输出格式及保真边界](https://pandoc.org/MANUAL.html#specifying-formats)

Pandoc 只能作为“先用 libwps/LibreOffice 转成 ODT，再转 DOCX”的第二段工具，但这会增加一次有损中间转换且没有明显收益，故不建议纳入首版依赖。

## 5. WPS Office 自动化与命令行

### 5.1 官方可证实的自动化能力

WPS 官方开放平台说明，客户端支持 Windows 等平台，可通过加载项、宏以及 C++/Java/浏览器应用集成实现文档自动化。[WPS 客户端开发概述](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-client-dev-introduction)

文字对象模型提供了可用于转换的关键接口：

- `CreateObject("kwps.application")` 可取得 WPS 文字的 `Application` 对象并打开文档；
- `Document.SaveAs2(FileName, FileFormat, ...)` 可按新名称或格式保存；
- `WdSaveFormat` 中 `wdFormatXMLDocument = 12` 对应 XML 文档格式，即常用的 DOCX 保存目标；严格 OOXML 则为 `wdFormatStrictOpenXMLDocument = 24`。

来源：[WPS `Application` 对象](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Application/obj) · [WPS `SaveAs2`](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Document/member/SaveAs2) · [WPS `WdSaveFormat`](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/enum/WdSaveFormat)

因此，在用户已经安装兼容版本 WPS Office 的 Windows 电脑上，用 WPS 自己打开其原生 `.wps` 再保存为 `.docx`，从格式语义上比让 LibreOffice 猜测更可靠。

### 5.2 尚未证实的能力

本次查到的 WPS 官方资料**没有给出可与 LibreOffice `--headless --convert-to` 等价的、稳定公开的纯命令行批量转换开关**。官方证据指向客户端对象模型、JS 加载项及系统集成，而不是一个无需客户端上下文的 headless CLI。因此首版不能把网络文章中的私有启动参数当作稳定接口。

WPS 自动化方案还需在实施前用目标用户的具体版本验证：能否从外部进程创建对象、是否会显示首次启动/登录/升级/格式确认对话框、批量运行是否可靠、是否需要 WPS 365 或企业集成授权。

### 5.3 授权和分发限制

不能把 WPS Office 当作可自由捆绑的运行时。WPS 的国际版 EULA 将免费受限版限制为个人、非商业用途，并禁止未经书面允许转售、再许可、出租、分发，或通过网络、SaaS/服务局等方式向多用户提供功能；该协议还明确中国大陆版本适用另一套随软件附带的协议。[WPS Office EULA](https://www.wps.com/eula/)

产品上的稳妥做法是：

- 不在安装包内捆绑 WPS Office；
- 仅检测并调用用户自行合法安装、已激活且其许可证允许自动化使用的 WPS；
- 商业发布或企业批量部署前，向金山/WPS 获取书面授权与目标版本的 SDK/自动化支持说明；
- 不把本地客户端自动化包装成多人在线转换服务，除非另有明确商业许可。

## 6. 开源项目参考

### `AndersonBY/wps2office`

该项目是一个 MIT 许可的 Go 包装器，提供 CLI 和本地 Web UI，并通过单独的 LibreOffice 可执行文件调用 `docx:Office Open XML Text` 等过滤器。它在工程结构上可参考：引擎发现、每任务独立配置、超时、非空输出校验、覆盖策略、本地回环 Web 服务和第三方许可证清单。[项目 README](https://github.com/AndersonBY/wps2office) · [项目 LICENSE](https://github.com/AndersonBY/wps2office/blob/main/LICENSE) · [项目约定的过滤器映射](https://github.com/AndersonBY/wps2office/blob/main/AGENTS.md)

但不能直接把它的兼容性声明当作本项目验收依据：其 README 说明集成测试是先生成办公文件、转成旧容器，再**重命名为 WPS 扩展名**后回转，并非以真实金山 WPS `.wps/.et/.dps` 样本验证；而且其离线包目标是麒麟 Linux，不是 Windows。可借鉴外壳设计，不能借此证明 LibreOffice 支持金山 WPS 原生格式。[项目测试说明](https://github.com/AndersonBY/wps2office#testing)

## 7. 推荐的软件范围与架构

### 7.1 最小可交付版本（推荐）

Windows 桌面程序，功能限定为：选择一个或多个文件、选择输出目录、转换、显示逐文件成功/失败原因、打开输出目录。转换引擎先只支持本机 LibreOffice，产品名称和说明明确写“Microsoft Works `.wps` 转 Word `.docx`”。

建议模块：

1. **格式探测层**：不只看扩展名；先尝试由 LibreOffice/libwps 内容检测确认 Microsoft Works，无法确认则归为“未知或金山 WPS 格式”。
2. **LibreOffice 后端**：发现标准安装路径或让用户选择 `soffice.com`；每任务创建独立配置目录；设置超时；捕获标准输出、标准错误和退出码。
3. **结果校验层**：检查输出存在且非空，检查 DOCX ZIP 结构和 `[Content_Types].xml`，绝不因同名旧输出存在而误报成功。
4. **桌面 UI**：支持拖放、批量队列、取消、失败详情和“保留原文件”；默认不覆盖。
5. **隐私与安全**：全程本地处理；临时目录按任务隔离并在结束后清理；不要自动执行输入文档中的宏；对不受信任文档应使用普通用户权限、超时和资源上限。

### 7.2 第二阶段

增加可选 WPS Office 后端，仅当：

- 用户提供至少一组真实金山 WPS `.wps` 样本和期望 DOCX；
- 目标 WPS 版本的外部自动化调用已在干净 Windows 环境验证；
- 授权条件已确认；
- UI 能明确显示当前使用的是 LibreOffice 还是 WPS 引擎，以及失败是否由登录、授权、弹窗或格式不支持导致。

## 8. 分发建议

### 8.1 LibreOffice

LibreOffice 官方称其为自由软件，主体采用 MPL 2.0，同时每个发行版包含多种其他开源许可证的软件；官方要求以安装包内 `LICENSE` 和“License Information”为准。MPL 2.0 允许与独立程序组成 Larger Work，但重新分发 Covered Software 的可执行形式时，需要保留通知，并告知接收者如何获得相应源代码；商标权不随开源许可证自动授予。[LibreOffice 许可证页](https://www.libreoffice.org/licenses/)

首版最省风险的方式是让用户自行安装 LibreOffice，本软件只发现并调用它。若未来要捆绑未修改的 LibreOffice：

- 随安装包保留完整上游许可证和第三方声明；
- 固定版本并提供对应源代码获取方式；
- 明确 LibreOffice 是独立第三方组件，不暗示官方背书；
- 若修改 LibreOffice/libwps，提供相应修改过的 Covered Files 源代码并遵守商标政策；
- 发布前对实际捆绑版本的完整 `LICENSE` 做一次法律/合规复核。

### 8.2 本软件自身

若仅通过子进程调用用户安装的 LibreOffice/WPS，本软件可采用 MIT、Apache-2.0 或闭源商业许可证；第三方程序仍各自受其许可证约束。不要复制第三方源码或图标后仍假定只有自己的许可证生效。

## 9. 开发前仍缺少的信息

当前需求中的主要缺口不是界面，而是输入格式定义。进入实现前应向用户确认并收集：

1. `.wps` 文件由“Microsoft Works”还是“金山 WPS Office”生成；最好提供 3–10 个可公开用于测试的真实样本。
2. 只需要 `.docx`，还是也要旧 `.doc`。
3. 是否允许要求用户安装 LibreOffice 或 WPS Office；是否要求单文件离线安装包。
4. 是否用于个人本机、企业内部分发，还是公网转换服务。
5. 是否需要批量目录、递归扫描、覆盖策略、密码文档、宏、页眉页脚、图片、表格和批注等保真要求。

没有真实样本前，能负责任承诺的是“为 Microsoft Works `.wps` 提供基于 LibreOffice/libwps 的尽力转换”，不能笼统承诺“所有 WPS 文件均可无损转 Word”。
