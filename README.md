本地学习库（智能文档学习库）把 PDF、Word、Markdown 原始文档统一复制到可自行指定的资料库目录，建立原文索引，并通过 1-25 字关键词快速查找原文。所有数据仅在本机处理，不上传服务器，搜索结果只来自文档原始文字，不会进行 AI 概括、改写或联网查询。

主要用途：

将分散的 PDF、Word、Markdown 资料统一归档到本地资料库；

用 1-25 字关键词快速检索原文片段，并标黄关键词；

在软件内部阅读并批注 PDF（高亮、下划线、手写、备注等）；

从试卷中逐题检索知识库原文，生成答案与出处；

标记重点考点，并统计答案出处引用排名。

功能特点
完全离线运行，数据仅保存在本机，不上传服务器。

支持 .pdf、.docx、.md / .markdown 等格式导入与索引。

关键词检索（1-25 字），结果按关键词出现次数排序，卡片式展示原文片段。

内置 PDF 阅读器：标签页阅读、全文检索、精确逐字选字、高亮、下划线、手写、备注、批注清单与批注版另存。

试卷答案查找：逐题检索原文，生成答案 Markdown、答案 Word、出处分析 Markdown，并附跨卷引用排名。

重点标记清单：自定义标记类型与备注，绑定原文片段与来源，不修改原始资料。

文档与分类管理：新建分类、批量导入、标签、备注、移动、删除、排序与过滤。

备份与恢复：导出包含原始文档、解析文本、标签清单的 ZIP，并可从 ZIP 恢复。

扫描型 PDF 自动调用本机离线 OCR，建立可搜索索引。

项目源码结构清晰，便于阅读、修改和二次开发。

主要技术栈：Python、PowerShell。

支持通过 Git 进行版本管理和协作。

本项目为本地离线工具，AI 辅助分析开关默认关闭；未配置 AI 服务时不会伪造内容，只显示文档原文和相应提示。

环境要求
Git

Python 3.10 或更高版本

PowerShell（用于运行打包脚本 build.ps1）

安装与运行
克隆仓库：

bash
git clone https://github.com/<你的账号>/本地学习库.git
cd 本地学习库
创建虚拟环境并安装依赖：

powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt
从源码启动：

powershell
.\.venv\Scripts\python.exe main.py
重新打包（在 PowerShell 中运行）：

powershell
.\build.ps1
生成文件位于 dist\智能文档学习库3.6.4.exe。

直接使用（打包版）
双击 智能文档学习库.exe。

在顶部工具栏先选择“导入到分类”；需要新的分类时点击“新建分类”。

点击“添加文件”选择一个或多个资料，或者点击“批量添加文件夹”。

等待程序完成复制、解析和索引。

在顶部输入 1-25 字关键词，按回车或点击“搜索原文”。

左侧用自动换行的卡片完整显示原文片段；点选结果后，右侧显示完整原文并标黄关键词。

项目结构
build.ps1 —— 打包脚本，生成 Windows 可执行文件。

document_converter.py —— Word / Markdown 转可搜索 PDF 等文档转换逻辑。

exam_modules.py —— 试卷答案查找与导出相关模块。

library_core.py —— 资料库核心：导入、索引、检索、分类与数据库操作。

main.py —— 程序入口与主界面。

pdf_reader.py —— PDF 解析、文字提取与批注处理。

pdf_viewer.py —— 内置 PDF 阅读器界面。

question_research.py —— 文档问题检索整理与出处分析。

requirements.txt —— Python 依赖清单。

test_core.py —— 核心功能测试。

test_exam_modules.py —— 试卷模块测试。

test_exam_workflow.py —— 试卷流程测试。

test_new_features.py —— 新增功能测试。

使用说明.md —— 详细使用说明与版本更新日志。

assets/ —— 图标等资源文件。

关键文件
build.ps1 —— 打包入口，运行后生成 dist\智能文档学习库3.6.4.exe。

requirements.txt —— 依赖清单，安装环境时使用。

本地数据位置
默认保存在：

C:\Users\你的用户名\AppData\Local\智能文档学习库\

其中 资料文件 保存导入的原始文件副本，学习库.db 保存索引、完整解析文本、标签和元数据。所有数据仅在本机处理，不上传服务器。

贡献
欢迎提交 Issue 或 Pull Request。提交改动前，请先确认代码能够正常运行并且不包含密钥、Token、.env 等敏感信息。

许可证
当前项目暂未指定开源许可证。如需开放复制、修改和分发权限，请根据实际需要添加许可证。

本回答由 AI 生成，内容仅供参考，请仔细甄别
