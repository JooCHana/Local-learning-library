from __future__ import annotations

import ctypes
import json
import os
import queue
import re
import shutil
import sys
import threading
import uuid
import math
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText
from tkinter import font as tkfont

from library_core import LibraryDatabase, SearchHit, make_snippet, open_at_location, open_path, reveal_path
from library_core import collect_supported_files
from document_converter import convert_and_import
from question_research import (
    QuestionResearchResult,
    export_question_research_markdown,
    export_question_research_word,
    organize_question_from_library,
)
from pdf_reader import ZoteroPDFViewer as PDFViewer
from exam_modules import (
    ExamService,
    ExamTaskResult,
    export_batch_analysis_markdown,
    export_analysis_markdown,
    export_answer_markdown,
    export_answer_word,
    export_result_bundle,
    export_individual_answers,
    answer_source_text,
    answer_source_label,
    parse_exam,
    split_exam_content,
)


APP_NAME = "智能文档学习库"
APP_VERSION = "3.6.4"
CONFIG_DIR = Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / APP_NAME
CONFIG_FILE = CONFIG_DIR / "settings.json"


def load_library_root() -> Path:
    try:
        value = json.loads(CONFIG_FILE.read_text(encoding="utf-8")).get("library_root", "")
        if str(value).strip():
            return Path(str(value)).expanduser()
    except Exception:
        pass
    return CONFIG_DIR


def save_library_root(path: Path) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(
        json.dumps({"library_root": str(path.resolve())}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


APP_DIR = load_library_root()
DB_PATH = APP_DIR / "学习库.db"
STORAGE_DIR = APP_DIR / "资料文件"


def bundled_resource(relative_path: str) -> Path:
    """Return a resource path both in source mode and in a PyInstaller bundle."""
    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return bundle_root / relative_path


def center_window(window: tk.Toplevel, width: int, height: int) -> None:
    """Open a resizable child window in the visible center of the current screen."""
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    width = min(width, max(640, screen_width - 80))
    height = min(height, max(480, screen_height - 100))
    x = max(0, (screen_width - width) // 2)
    y = max(0, (screen_height - height) // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")
    window.resizable(True, True)


def enable_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class LearningLibraryApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        icon_path = bundled_resource("assets/app_icon.ico")
        if icon_path.exists():
            try:
                root.iconbitmap(default=str(icon_path))
            except tk.TclError:
                pass
        self.database = LibraryDatabase(DB_PATH, STORAGE_DIR)
        self.exam_service = ExamService(self.database)
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.hits: dict[str, SearchHit] = {}
        self.current_hit: SearchHit | None = None
        self.all_hits: list[SearchHit] = []
        self.current_page = 1
        self.page_size = 25
        self.busy = False
        self.exam_ui: dict[str, object] | None = None
        self.question_research_ui: dict[str, object] | None = None
        self.search_doc_id = 0
        self.selected_categories: set[str] = set()

        root.title(f"{APP_NAME} {APP_VERSION}")
        root.geometry("1180x780")
        root.minsize(940, 640)
        root.option_add("*Font", ("Microsoft YaHei UI", 10))

        self.search_text = tk.StringVar()
        self.status_text = tk.StringVar(value="准备就绪")
        self.stats_text = tk.StringVar()
        self.result_text = tk.StringVar(value="请输入关键词进行搜索")
        self.page_text = tk.StringVar(value="第 1 / 1 页")
        self.tag_filter = tk.StringVar(value="全部标签")
        self.category_filter = tk.StringVar(value="全部分类")
        self.import_category = tk.StringVar(value="未分类")
        self.search_scope_text = tk.StringVar(value="查询范围：全部文档")
        self.search_mode = tk.StringVar(value="普通文档检索")
        self.category_scope_text = tk.StringVar(value="全部分类")
        self.ai_analysis_enabled = tk.BooleanVar(value=False)
        self.location_text = tk.StringVar(value=f"存放位置：{APP_DIR}")

        self.configure_style()
        self.build_menu()
        self.build_ui()
        self.refresh_documents()
        root.after(100, self.poll_events)

    def configure_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 21, "bold"))
        style.configure("Section.TLabel", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 10, "bold"))
        # Double previous list row heights; also leave room for scaled Chinese fonts.
        line_height = tkfont.Font(family="Microsoft YaHei UI", size=10).metrics("linespace")
        self.list_rowheight = max(40, 2 * line_height)
        style.configure("Treeview", rowheight=self.list_rowheight, font=("Microsoft YaHei UI", 10))

    def build_menu(self) -> None:
        """Group top-level actions by file, library, retrieval, tools and help."""
        menu = tk.Menu(self.root)

        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="导入文件（Word/MD自动生成可搜索PDF）…", command=self.add_files)
        file_menu.add_command(label="批量添加文件夹…", command=self.add_folder)
        file_menu.add_separator()
        file_menu.add_command(label="导出当前文件…", command=self.export_current_file)
        file_menu.add_command(label="另存原版PDF…", command=lambda: self.open_current_pdf_for_export(False))
        file_menu.add_command(label="另存批注版PDF…", command=lambda: self.open_current_pdf_for_export(True))
        file_menu.add_separator()
        file_menu.add_command(label="重命名当前文件…", command=self.rename_current_document)
        file_menu.add_command(label="移动当前文件到分类…", command=self.move_current_document)
        file_menu.add_command(label="删除当前文件…", command=self.delete_current_document)
        file_menu.add_command(label="打开当前文件所在文件夹", command=self.reveal_selected_source)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.destroy)
        self.file_menu = file_menu
        file_menu.configure(postcommand=self.update_file_menu_state)
        menu.add_cascade(label="文件", menu=file_menu)

        library_menu = tk.Menu(menu, tearoff=False)
        library_menu.add_command(label="管理文档与分类…", command=self.manage_documents_v3)
        library_menu.add_command(label="新建分类…", command=self.new_category)
        library_menu.add_command(label="调整资料库存放位置…", command=self.change_library_location)
        library_menu.add_separator()
        library_menu.add_command(label="导出学习库备份…", command=self.export_backup)
        library_menu.add_command(label="恢复学习库备份…", command=self.restore_backup)
        library_menu.add_separator()
        library_menu.add_command(label="刷新资料列表", command=self.refresh_documents)
        library_menu.add_command(label="刷新全文索引", command=self.rebuild_all)
        menu.add_cascade(label="资料库", menu=library_menu)

        search_menu = tk.Menu(menu, tearoff=False)
        search_menu.add_command(label="文档问题检索整理…", command=self.open_question_research)
        search_menu.add_command(label="关键词原文检索", command=lambda: self.search_entry.focus_set())
        search_menu.add_separator()
        search_menu.add_command(label="选择单个文件查询…", command=self.choose_single_search_document)
        search_menu.add_command(label="恢复全部文档查询", command=self.clear_search_document)
        search_menu.add_command(label="资料/试卷文件列表…", command=self.find_documents_window)
        search_menu.add_separator()
        search_menu.add_command(label="试卷答案查找…", command=self.open_exam_module)
        menu.add_cascade(label="检索", menu=search_menu)

        tools_menu = tk.Menu(menu, tearoff=False)
        tools_menu.add_command(label="PDF阅读与批注管理…", command=self.open_pdf_annotation_manager)
        tools_menu.add_command(label="重点标记清单…", command=self.show_marker_list)
        tools_menu.add_command(label="重复考点分析…", command=self.open_exam_module)
        tools_menu.add_separator()
        tools_menu.add_command(label="当前Word/Markdown转换为可搜索PDF", command=self.convert_current_document)
        menu.add_cascade(label="工具", menu=tools_menu)

        help_menu = tk.Menu(menu, tearoff=False)
        help_menu.add_command(label="关于与版本更新日志…", command=self.show_about)
        menu.add_cascade(label="帮助", menu=help_menu)
        self.root.configure(menu=menu)

    def show_about(self) -> None:
        window = tk.Toplevel(self.root)
        window.title(f"关于 {APP_NAME}")
        center_window(window, 760, 620)
        window.minsize(620, 480)
        frame = ttk.Frame(window, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text=f"版本 {APP_VERSION}　｜　创作者：韩家臣", foreground="#245f8a").pack(
            anchor="w", pady=(4, 14)
        )
        ttk.Label(frame, text="版本更新日志", style="Section.TLabel").pack(anchor="w", pady=(0, 7))
        log = ScrolledText(
            frame, wrap="word", font=("Microsoft YaHei UI", 10), padx=12, pady=10,
            relief="solid", borderwidth=1,
        )
        log.pack(fill="both", expand=True)
        log.insert(
            "1.0",
            "3.6.4\n"
            "• 试卷答案只输出与题干直接对应的原文句子，不再复制整段或整章上下文。\n"
            "• 长检索片段会自动舍弃无关章节、条款和说明，保留可直接回答题目的部分。\n"
            "• 无可直接作答的有效原文时按未命中处理，显示“【无】”。\n"
            "• 出处统一为“【来源：文件名 · 段落/页码】”，未命中时显示“【来源：未检索到匹配文档】”。\n"
            "• 移除旧版冗长的答案标题，统一使用“答案”。\n\n"
            "3.6.3\n"
            "• 修复试卷答案 Word 导出遇到 NUL 或隐藏控制字符时失败的问题。\n"
            "• 试卷答案统一使用“题目—答案—出处”结构，内容只来自知识库原文。\n"
            "• 未检索到匹配文档时，答案明确显示“【无】”并标注未命中出处。\n"
            "• Markdown 和 Word 的单卷、多卷及自动生成结果保持一致。\n\n"
            "3.6.2\n"
            "• PDF 阅读改为浏览器式标签页，默认不再重复弹出多个窗口。\n"
            "• 右上角增加“新建窗口”，可按需将当前 PDF 独立打开。\n"
            "• PDF 增加内置目录面板和全文检索，支持跳转到页内匹配位置。\n"
            "• 文字选择改为逐字、跨行连续拖选，取消矩形框选效果。\n"
            "• 选中原文后右键可复制、高亮、下划线或添加文字备注。\n"
            "• 重新定义 PDF 默认缩放：原 300% 显示尺寸作为新的 100%。\n"
            "• 保留原有批注工具、颜色和存储格式，历史批注可继续编辑。\n\n"
            "3.6.1\n"
            "• 管理文档与分类支持安全删除分类，分类内资料自动移至“未分类”。\n"
            "• 放大新建分类窗口；管理窗口支持自由缩放和最小化。\n"
            "• 文档问题整理结果支持导出 Markdown 或 Word，并可选择本地保存位置。\n"
            "• 顶部菜单重新分组为文件、资料库、检索、工具、帮助。\n"
            "• PDF 支持字符级选中复制，右键可添加荧光高亮或下划线。\n"
            "• 工具窗口统一居中，修复部分窗口内容显示不全。\n"
            "• 软件图标更换为用户提供的 SVG 图标。\n\n"
            "3.6\n"
            "• 修复问题整理按钮位置及 PDF 精确选字、复制。\n\n"
            "3.5\n"
            "• 新增文档问题检索整理、原文去重、冲突标注及来源清单。",
        )
        log.configure(state="disabled")
        ttk.Button(frame, text="关闭", command=window.destroy).pack(anchor="e", pady=(10, 0), ipadx=12)

    def update_file_menu_state(self) -> None:
        has_document = self._current_document() is not None
        state = "normal" if has_document else "disabled"
        for label in (
            "导出当前文件…", "另存原版PDF…", "另存批注版PDF…", "重命名当前文件…",
            "当前Word/Markdown转换为可搜索PDF", "移动当前文件到分类…", "删除当前文件…", "打开当前文件所在文件夹",
        ):
            try:
                self.file_menu.entryconfigure(label, state=state)
            except tk.TclError:
                pass

    def build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x")
        ttk.Label(header, text=APP_NAME, style="Title.TLabel").pack(side="left")
        ttk.Label(header, textvariable=self.stats_text, foreground="#5d6670").pack(side="right", pady=(8, 0))
        ttk.Label(
            outer,
            text="PDF、Word、Markdown 私有学习资料库 · 全部文件和索引只保存在本机",
            foreground="#5d6670",
        ).pack(anchor="w", pady=(2, 14))

        toolbar = ttk.LabelFrame(outer, text="检索范围", padding=8)
        toolbar.pack(fill="x", pady=(0, 10))
        location_row = ttk.Frame(toolbar)
        location_row.pack(fill="x", pady=(0, 7))
        self.location_button = ttk.Button(location_row, text="调整资料库存放位置…", command=self.change_library_location)
        self.location_button.pack(side="left")
        ttk.Label(location_row, textvariable=self.location_text, foreground="#65707c").pack(
            side="left", padx=(10, 0)
        )
        category_row = ttk.Frame(toolbar)
        category_row.pack(fill="x")
        ttk.Label(category_row, text="导入到分类").pack(side="left")
        self.import_category_combo = ttk.Combobox(
            category_row, textvariable=self.import_category, state="readonly", width=18
        )
        self.import_category_combo.pack(side="left", padx=(5, 7))
        self.new_category_button = ttk.Button(category_row, text="新建分类文件夹", command=self.new_category)
        self.new_category_button.pack(side="left")
        ttk.Label(
            category_row,
            text="原始文件存入此分类文件夹",
            foreground="#65707c",
        ).pack(side="left", padx=(12, 0))

        action_row = ttk.Frame(toolbar)
        action_row.pack(fill="x", pady=(7, 0))
        self.add_files_button = ttk.Button(action_row, text="添加文件…", command=self.add_files)
        self.add_files_button.pack(side="left")
        self.add_folder_button = ttk.Button(action_row, text="批量添加文件夹…", command=self.add_folder)
        self.add_folder_button.pack(side="left", padx=7)
        self.manage_button = ttk.Button(action_row, text="管理文档与分类", command=self.manage_documents_v3)
        self.manage_button.pack(side="left")
        self.rebuild_button = ttk.Button(action_row, text="刷新索引", command=self.rebuild_all)
        self.rebuild_button.pack(side="left", padx=7)
        self.backup_button = ttk.Menubutton(action_row, text="备份与恢复")
        backup_menu = tk.Menu(self.backup_button, tearoff=False)
        backup_menu.add_command(label="导出备份…", command=self.export_backup)
        backup_menu.add_command(label="恢复备份…", command=self.restore_backup)
        self.backup_button.configure(menu=backup_menu)
        self.backup_button.pack(side="left")
        self.restore_button = self.backup_button
        # File operations live in the top File menu; keep these widgets only as
        # internal busy-state handles and remove the former home-page clutter.
        location_row.pack_forget()
        category_row.pack_forget()
        action_row.pack_forget()
        study_row = ttk.Frame(toolbar)
        study_row.pack(fill="x")
        self.single_file_button = ttk.Button(study_row, text="单个文件查询…", command=self.choose_single_search_document)
        self.single_file_button.pack(side="left")
        self.clear_scope_button = ttk.Button(study_row, text="恢复全部文档查询", command=self.clear_search_document)
        self.clear_scope_button.pack(side="left", padx=7)
        self.paper_find_button = ttk.Button(study_row, text="资料/试卷查找", command=self.find_documents_window)
        self.paper_find_button.pack(side="left")
        self.exam_button = ttk.Button(study_row, text="试卷答案查找", style="Accent.TButton", command=self.open_exam_module)
        self.marker_list_button = ttk.Button(study_row, text="重点标记清单", command=self.show_marker_list)
        ttk.Label(study_row, textvariable=self.search_scope_text, foreground="#145b91").pack(side="left", padx=(14, 0))

        search = ttk.Frame(outer)
        search.pack(fill="x", pady=(0, 12))
        ttk.Label(
            search, text="输入 1-25 字关键词检索原文", foreground="#53606c"
        ).pack(anchor="w", pady=(0, 4))
        search_row = ttk.Frame(search)
        search_row.pack(fill="x")
        self.search_mode_combo = ttk.Combobox(
            search_row, textvariable=self.search_mode, state="readonly", width=15,
            values=("普通文档检索", "试卷专项检索"),
        )
        self.search_mode_combo.pack(side="left", padx=(0, 8), ipady=5)
        self.search_mode_combo.bind("<<ComboboxSelected>>", self.on_search_mode_changed)
        self.search_entry = ttk.Entry(search_row, textvariable=self.search_text, font=("Microsoft YaHei UI", 13))
        self.search_entry.pack(side="left", fill="x", expand=True, ipady=7)
        self.search_entry.bind("<Return>", lambda _event: self.do_search())
        self.search_entry.bind("<KeyRelease>", self.on_search_key)
        self.search_button = ttk.Button(search_row, text="搜索原文", style="Accent.TButton", command=self.do_search)
        self.search_button.pack(side="left", padx=(10, 0), ipadx=12, ipady=6)

        filter_row = ttk.Frame(search)
        filter_row.pack(fill="x", pady=(6, 0))
        ttk.Label(filter_row, text="筛选标签", foreground="#53606c").pack(side="left", padx=(0, 5))
        self.tag_combo = ttk.Combobox(filter_row, textvariable=self.tag_filter, state="readonly", width=16)
        self.tag_combo["values"] = ("全部标签",)
        self.tag_combo.pack(side="left")
        self.tag_combo.bind(
            "<<ComboboxSelected>>",
            lambda _event: self.do_search() if self.search_text.get().strip() else None,
        )
        ttk.Label(filter_row, text="筛选分类", foreground="#53606c").pack(side="left", padx=(14, 5))
        self.category_combo = ttk.Combobox(
            filter_row, textvariable=self.category_filter, state="readonly", width=16
        )
        self.category_combo["values"] = ("全部分类",)
        self.category_scope_button = ttk.Button(
            filter_row, textvariable=self.category_scope_text, command=self.choose_search_categories
        )
        self.category_scope_button.pack(side="left")
        self.ai_analysis_check = ttk.Checkbutton(
            filter_row, text="允许AI辅助分析（默认关闭）", variable=self.ai_analysis_enabled,
            command=lambda: self.show_selected_hit() if self.current_hit else None,
            state="disabled",
        )
        self.ai_analysis_check.pack(side="left", padx=(14, 0))

        paned = ttk.Panedwindow(outer, orient="horizontal")
        paned.pack(fill="both", expand=True)
        left = ttk.Frame(paned, padding=(0, 0, 8, 0))
        right = ttk.Frame(paned, padding=(8, 0, 0, 0))
        paned.add(left, weight=3)
        paned.add(right, weight=4)

        result_header = ttk.Frame(left)
        result_header.pack(fill="x", pady=(0, 6))
        ttk.Label(result_header, text="搜索结果", style="Section.TLabel").pack(side="left")
        ttk.Label(result_header, textvariable=self.result_text, foreground="#65707c").pack(side="right")

        tree_frame = ttk.Frame(left)
        tree_frame.pack(fill="both", expand=True)
        self.results = ScrolledText(
            tree_frame, wrap="word", padx=10, pady=8, relief="solid", borderwidth=1,
            font=("Microsoft YaHei UI", 10), cursor="hand2", spacing1=2, spacing3=3,
        )
        self.results.pack(fill="both", expand=True)
        self.results.bind("<Key>", lambda _event: "break")
        self.results.tag_configure("result_match", background="#fff19c", foreground="#111111")

        pages = ttk.Frame(left)
        pages.pack(fill="x", pady=(8, 0))
        self.prev_button = ttk.Button(pages, text="上一页", command=lambda: self.change_page(-1), state="disabled")
        self.prev_button.pack(side="left")
        self.next_button = ttk.Button(pages, text="下一页", command=lambda: self.change_page(1), state="disabled")
        self.next_button.pack(side="left", padx=8)
        ttk.Label(pages, textvariable=self.page_text, foreground="#65707c").pack(side="left")

        detail_header = ttk.Frame(right)
        detail_header.pack(fill="x", pady=(0, 6))
        ttk.Label(detail_header, text="完整原文", style="Section.TLabel").pack(side="left")
        self.location_label = ttk.Label(detail_header, text="", foreground="#65707c")
        self.location_label.pack(side="right")
        self.detail = ScrolledText(
            right, wrap="word", padx=14, pady=12, relief="solid", borderwidth=1,
            font=("Microsoft YaHei UI", 11), spacing1=3, spacing3=5,
        )
        self.detail.pack(fill="both", expand=True)
        self.detail.configure(state="disabled")
        self.detail.tag_configure("match", background="#fff19c", foreground="#111111")

        detail_buttons = ttk.Frame(right)
        detail_buttons.pack(fill="x", pady=(8, 0))
        self.open_button = ttk.Button(detail_buttons, text="打开原始文档", command=self.open_selected_source, state="disabled")
        self.open_button.pack(side="left")
        self.reveal_button = ttk.Button(detail_buttons, text="打开所在文件夹", command=self.reveal_selected_source, state="disabled")
        self.reveal_button.pack(side="left", padx=8)
        self.copy_button = ttk.Button(detail_buttons, text="复制原文", command=self.copy_selected_text, state="disabled")
        self.copy_button.pack(side="left")
        self.mark_button = ttk.Button(
            detail_buttons, text="标记选中的原文", command=self.add_selected_marker, state="disabled"
        )
        self.mark_button.pack(side="left", padx=8)

        status_bar = ttk.Frame(outer)
        status_bar.pack(fill="x", pady=(8, 0))
        self.progress = ttk.Progressbar(status_bar, mode="determinate", length=180)
        self.progress.pack(side="right", padx=(10, 0))
        ttk.Label(status_bar, textvariable=self.status_text, foreground="#1261a0").pack(side="right")
        self.search_entry.focus_set()

    def on_search_key(self, _event=None) -> None:
        value = self.search_text.get()
        if self.search_mode.get() == "普通文档检索" and len(value) > 25:
            self.search_text.set(value[:25])
            messagebox.showwarning("关键词过长", "请输入 25 字以内关键词")

    def on_search_mode_changed(self, _event=None) -> None:
        exam_mode = self.search_mode.get() == "试卷专项检索"
        self.ai_analysis_check.configure(state="normal" if exam_mode else "disabled")
        if not exam_mode:
            self.ai_analysis_enabled.set(False)
        if self.search_text.get().strip():
            self.do_search()

    def do_search(self) -> None:
        keyword = self.search_text.get().strip()
        if not keyword:
            messagebox.showinfo("请输入关键词", "请输入 1-25 字关键词检索原文")
            return
        if self.search_mode.get() == "普通文档检索" and len(keyword) > 25:
            messagebox.showwarning("关键词过长", "请输入 25 字以内关键词")
            return
        try:
            tag = "" if self.tag_filter.get() == "全部标签" else self.tag_filter.get()
            categories = sorted(self.selected_categories)
            if self.search_mode.get() == "试卷专项检索":
                self.all_hits = self.exam_service.search_text_candidates(keyword, categories)
            else:
                self.all_hits = self.database.search(keyword, tag, categories, self.search_doc_id)
                for hit in self.all_hits:
                    if hit.content.rstrip().endswith(("：", ":")):
                        nearby = self.database.chunk_context(hit.chunk_id, following=3)
                        if len(nearby) > 1:
                            hit.content = "\n\n".join(str(row["content"]) for row in nearby)
                            hit.snippet = make_snippet(hit.content, keyword)
        except Exception as exc:
            messagebox.showerror("搜索失败", str(exc))
            return
        self.current_page = 1
        self.render_page()
        suffix = "候选题目" if self.search_mode.get() == "试卷专项检索" else "原文"
        self.result_text.set(f"找到 {len(self.all_hits)} 条{suffix}，已按匹配度排序")
        self.status_text.set("搜索完成")
        if not self.all_hits:
            messagebox.showinfo("未找到原文", "未在资料库找到该关键词相关原文，请更换关键词重试")

    def render_page(self) -> None:
        self.results.delete("1.0", "end")
        self.hits.clear()
        self.current_hit = None
        self.clear_detail()
        total_pages = max(1, (len(self.all_hits) + self.page_size - 1) // self.page_size)
        self.current_page = min(max(1, self.current_page), total_pages)
        start = (self.current_page - 1) * self.page_size
        for index, hit in enumerate(self.all_hits[start : start + self.page_size]):
            item_id = f"hit-{start + index}"
            self.hits[item_id] = hit
            block_start = self.results.index("end-1c")
            metric = f"匹配 {hit.score} 分" if self.search_mode.get() == "试卷专项检索" else f"命中 {hit.score} 次"
            self.results.insert("end", f"{hit.name}  ·  {hit.location}  ·  {metric}\n", (item_id, "result_header"))
            self.results.insert("end", hit.snippet + "\n\n", (item_id, "result_body"))
            block_end = self.results.index("end-1c")
            background = "#f4f7fb" if index % 2 == 0 else "#ffffff"
            self.results.tag_configure(
                item_id, background=background, lmargin1=8, lmargin2=8,
                rmargin=8, spacing1=6, spacing3=6,
            )
            self.results.tag_configure("result_header", foreground="#145b91", font=("Microsoft YaHei UI", 10, "bold"))
            self.results.tag_bind(item_id, "<Button-1>", lambda _event, key=item_id: self.select_result(key))
            self.results.tag_bind(item_id, "<Double-Button-1>", lambda _event, key=item_id: self.open_result(key))
            self._highlight_result_keyword(block_start, block_end)
        self.results.tag_raise("result_match")
        self.page_text.set(f"第 {self.current_page} / {total_pages} 页")
        self.prev_button.configure(state="normal" if self.current_page > 1 else "disabled")
        self.next_button.configure(state="normal" if self.current_page < total_pages else "disabled")
        first = next(iter(self.hits), None)
        if first:
            self.select_result(first)

    def _highlight_result_keyword(self, start: str, end: str) -> None:
        keyword = self.search_text.get().strip()
        position = start
        while keyword:
            found = self.results.search(keyword, position, stopindex=end, nocase=True)
            if not found:
                break
            found_end = f"{found}+{len(keyword)}c"
            self.results.tag_add("result_match", found, found_end)
            position = found_end

    def select_result(self, item_id: str) -> None:
        hit = self.hits.get(item_id)
        if not hit:
            return
        self.current_hit = hit
        self.show_selected_hit()

    def open_result(self, item_id: str) -> None:
        self.select_result(item_id)
        self.open_selected_source()

    def change_page(self, direction: int) -> None:
        self.current_page += direction
        self.render_page()

    def selected_hit(self) -> SearchHit | None:
        return self.current_hit

    def show_selected_hit(self, _event=None) -> None:
        hit = self.selected_hit()
        if not hit:
            return
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        if self.search_mode.get() == "试卷专项检索":
            parts = split_exam_content(hit.content)
            self.detail.configure(font=("Microsoft YaHei UI", 11))
            missing_analysis = (
                "尚未配置AI服务，未生成内容；当前仅展示文档原文。"
                if self.ai_analysis_enabled.get()
                else "当前文档解析有限；AI辅助已关闭，只展示文档原文。"
            )
            content = (
                f"题型：{parts['type']}\n\n"
                f"【题目】\n{parts['question'] or '未能单独拆分'}\n\n"
                f"【选项】\n{parts['options'] or '原文未识别到选项'}\n\n"
                f"【答案】\n{parts['answer'] or '原文未提供明确答案'}\n\n"
                f"【解析】\n{parts['analysis'] or missing_analysis}\n\n"
                f"【原文来源】\n{hit.name} · {hit.location}\n匹配度：{hit.score}分\n\n"
                f"【完整原文】\n{hit.content}"
            )
        elif hit.location.startswith("表格"):
            self.detail.configure(font=("Consolas", 10), tabs=(140, 280, 420, 560))
            content = hit.content
        else:
            self.detail.configure(font=("Microsoft YaHei UI", 11))
            content = hit.content
        self.detail.insert("1.0", content)
        self.highlight_keyword()
        self.detail.configure(state="disabled")
        self.location_label.configure(text=f"{hit.name} · {hit.location} · 命中 {hit.score} 次")
        for button in (self.open_button, self.reveal_button, self.copy_button, self.mark_button):
            button.configure(state="normal")

    def highlight_keyword(self) -> None:
        keyword = self.search_text.get().strip()
        start = "1.0"
        first_match = None
        while keyword:
            position = self.detail.search(keyword, start, stopindex="end", nocase=True)
            if not position:
                break
            end = f"{position}+{len(keyword)}c"
            self.detail.tag_add("match", position, end)
            first_match = first_match or position
            start = end
        if first_match:
            self.detail.see(first_match)

    def clear_detail(self) -> None:
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.configure(state="disabled")
        self.location_label.configure(text="")
        for button in (self.open_button, self.reveal_button, self.copy_button, self.mark_button):
            button.configure(state="disabled")

    def open_selected_source(self) -> None:
        hit = self.selected_hit()
        if hit:
            if not Path(hit.path).exists():
                messagebox.showwarning("文件不存在", "学习库中的原始文档副本已丢失，请恢复备份或重新导入。")
                return
            try:
                if Path(hit.path).suffix.lower() == ".pdf":
                    match = re.search(r"(\d+)", hit.location)
                    self.open_pdf_preview(hit.path, int(match.group(1)) if match else 1)
                else:
                    open_at_location(hit.path, hit.location)
            except Exception as exc:
                messagebox.showerror("无法定位原文", str(exc))

    def reveal_selected_source(self) -> None:
        hit = self.selected_hit()
        if hit:
            self._open_existing(hit.path, True)

    def _open_existing(self, path: str, reveal: bool) -> None:
        if not Path(path).exists():
            messagebox.showwarning("文件不存在", "学习库中的原始文档副本已丢失，请恢复备份或重新导入。")
            return
        try:
            if reveal:
                reveal_path(path)
            elif Path(path).suffix.lower() == ".pdf":
                row = self.database.get_document_by_path(path)
                if row:
                    PDFViewer(self.root, self.database, row)
                else:
                    open_path(path)
            else:
                open_path(path)
        except Exception as exc:
            messagebox.showerror("无法打开", str(exc))

    def open_pdf_preview(self, path: str, page: int = 1) -> None:
        row = self.database.get_document_by_path(path)
        if not row:
            messagebox.showwarning("无法预览", "该PDF不在当前学习库中，请先导入后再批注。")
            return
        PDFViewer(self.root, self.database, row, page)

    def copy_selected_text(self) -> None:
        hit = self.selected_hit()
        if hit:
            self.root.clipboard_clear()
            self.root.clipboard_append(hit.content)
            self.status_text.set("原文已复制")

    def add_selected_marker(self) -> None:
        """Store the selected source range separately; never edit the source document."""
        hit = self.selected_hit()
        if not hit:
            return
        try:
            start_index = self.detail.index("sel.first")
            selected = self.detail.get("sel.first", "sel.last").strip()
            start_offset = len(self.detail.get("1.0", start_index))
        except tk.TclError:
            messagebox.showinfo("请先选择原文", "请在右侧完整原文中拖动鼠标，选中要标记的文字。")
            return
        if not selected:
            return
        marker_type = simpledialog.askstring(
            "重点标记类型", "请输入标记类型（例如：考点、定义、公式）：",
            initialvalue="考点", parent=self.root,
        )
        if marker_type is None:
            return
        note = simpledialog.askstring(
            "标记备注", "可选：输入备注（可留空）：", initialvalue="", parent=self.root
        )
        if note is None:
            return
        try:
            self.exam_service.add_marker(hit, selected, marker_type, note, start_offset)
            self.status_text.set(f"已添加“{marker_type.strip() or '考点'}”标记")
            messagebox.showinfo("标记完成", "重点标记已独立保存，原始知识库文件未被修改。")
        except Exception as exc:
            messagebox.showerror("标记失败", str(exc))

    def open_question_research(self) -> None:
        if self.question_research_ui and self.question_research_ui.get("window"):
            window = self.question_research_ui["window"]
            try:
                window.lift()
                window.focus_force()
                return
            except tk.TclError:
                self.question_research_ui = None

        window = tk.Toplevel(self.root)
        window.title(f"文档问题检索整理 · {APP_VERSION}")
        center_window(window, 1120, 760)
        window.minsize(860, 620)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=14)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="文档问题检索整理", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            frame,
            text="从学习库原文提取答案；自动合并完全相同的表述，疑似冲突同时保留并标注。",
            foreground="#5d6670",
        ).pack(anchor="w", pady=(2, 10))

        question_frame = ttk.LabelFrame(frame, text="输入需要查询的问题", padding=9)
        question_frame.pack(fill="x")
        question = ScrolledText(
            question_frame, height=4, wrap="word", font=("Microsoft YaHei UI", 11),
            padx=8, pady=6,
        )
        question.pack(fill="x", expand=True)
        if self.search_text.get().strip():
            question.insert("1.0", self.search_text.get().strip())

        option_row = ttk.Frame(frame)
        option_row.pack(fill="x", pady=8)
        ttk.Label(option_row, text="检索分类").pack(side="left")
        category_var = tk.StringVar(value="全部分类")
        category = ttk.Combobox(
            option_row, textvariable=category_var, state="readonly", width=24,
            values=("全部分类", *self.database.list_categories()),
        )
        category.pack(side="left", padx=(6, 12))
        status = tk.StringVar(value="等待检索")
        ttk.Label(option_row, textvariable=status, foreground="#1261a0").pack(side="right")

        command_row = ttk.Frame(frame)
        command_row.pack(fill="x", pady=(0, 8))
        ttk.Label(command_row, text="提示：按 Ctrl+Enter 也可开始", foreground="#65707c").pack(side="left")

        result_frame = ttk.LabelFrame(frame, text="原文整理结果", padding=8)
        result_frame.pack(fill="both", expand=True, pady=(0, 8))
        output = ScrolledText(
            result_frame, wrap="word", font=("Microsoft YaHei UI", 11),
            padx=12, pady=10, spacing1=3, spacing3=6,
        )
        output.pack(fill="both", expand=True)
        output.insert("1.0", "输入问题后点击“检索并整理”。")
        output.configure(state="disabled")

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")

        def copy_result() -> None:
            text = output.get("1.0", "end-1c").strip()
            if text:
                self.root.clipboard_clear()
                self.root.clipboard_append(text)
                status.set("整理结果已复制")

        def export_markdown() -> None:
            result = self.question_research_ui.get("result") if self.question_research_ui else None
            if not isinstance(result, QuestionResearchResult):
                return
            target = filedialog.asksaveasfilename(
                title="导出文档问题整理结果（Markdown）",
                parent=window,
                defaultextension=".md",
                filetypes=[("Markdown", "*.md")],
                initialfile="文档问题检索整理.md",
            )
            if target:
                try:
                    export_question_research_markdown(result, target)
                    status.set(f"已导出 Markdown：{Path(target).name}")
                except Exception as exc:
                    messagebox.showerror("Markdown 导出失败", str(exc), parent=window)

        def export_word() -> None:
            result = self.question_research_ui.get("result") if self.question_research_ui else None
            if not isinstance(result, QuestionResearchResult):
                return
            target = filedialog.asksaveasfilename(
                title="导出文档问题整理结果（Word）",
                parent=window,
                defaultextension=".docx",
                filetypes=[("Word 文档", "*.docx")],
                initialfile="文档问题检索整理.docx",
            )
            if target:
                try:
                    export_question_research_word(result, target)
                    status.set(f"已导出 Word：{Path(target).name}")
                except Exception as exc:
                    messagebox.showerror("Word 导出失败", str(exc), parent=window)

        copy_button = ttk.Button(buttons, text="复制整理结果", command=copy_result, state="disabled")
        copy_button.pack(side="left")
        markdown_button = ttk.Button(buttons, text="导出 Markdown…", command=export_markdown, state="disabled")
        markdown_button.pack(side="left", padx=7)
        word_button = ttk.Button(buttons, text="导出 Word…", command=export_word, state="disabled")
        word_button.pack(side="left")

        def start() -> None:
            prompt = question.get("1.0", "end-1c").strip()
            if not prompt:
                messagebox.showinfo("请输入问题", "请输入需要从文档库检索整理的问题。", parent=window)
                return
            run_id = uuid.uuid4().hex
            self.question_research_ui["run_id"] = run_id
            start_button.configure(state="disabled", text="正在检索整理…")
            copy_button.configure(state="disabled")
            markdown_button.configure(state="disabled")
            word_button.configure(state="disabled")
            status.set("正在提取关键词并读取全部匹配原文…")
            output.configure(state="normal")
            output.delete("1.0", "end")
            output.insert("1.0", "正在检索，请稍候……")
            output.configure(state="disabled")
            selected_category = "" if category_var.get() == "全部分类" else category_var.get()

            def worker() -> None:
                try:
                    research = organize_question_from_library(
                        self.database, prompt, categories=selected_category
                    )
                    self.events.put(("question_research_done", (run_id, research)))
                except Exception as exc:
                    self.events.put(("question_research_error", (run_id, str(exc))))

            threading.Thread(target=worker, daemon=True).start()

        start_button = ttk.Button(command_row, text="检索并整理", style="Accent.TButton", command=start)
        start_button.pack(side="right", ipadx=14, ipady=5)
        ttk.Button(buttons, text="关闭", command=window.destroy).pack(side="right", padx=7)
        question.bind("<Control-Return>", lambda _event: start())

        def closed(_event=None) -> None:
            if self.question_research_ui and self.question_research_ui.get("window") is window:
                self.question_research_ui = None

        window.bind("<Destroy>", closed)
        self.question_research_ui = {
            "window": window,
            "run_id": "",
            "output": output,
            "status": status,
            "start": start_button,
            "copy": copy_button,
            "exports": (markdown_button, word_button),
            "result": None,
        }
        question.focus_set()

    def show_marker_list(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("知识库重点标记清单")
        center_window(window, 1060, 620)
        window.minsize(820, 480)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=12)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="知识库重点标记清单", style="Section.TLabel").pack(anchor="w", pady=(0, 8))
        columns = ("type", "text", "document", "location", "note", "created")
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        for column, title, width in (
            ("type", "类型", 90), ("text", "原文片段", 390), ("document", "来源文档", 180),
            ("location", "原文位置", 110), ("note", "备注", 160), ("created", "标记时间", 145),
        ):
            tree.heading(column, text=title)
            tree.column(column, width=width, minwidth=60)
        vertical = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        markers: dict[str, object] = {}

        def load() -> None:
            markers.clear()
            tree.delete(*tree.get_children())
            for marker in self.exam_service.list_markers():
                item = str(marker.id)
                markers[item] = marker
                preview = marker.selected_text.replace("\n", " ")
                if len(preview) > 80:
                    preview = preview[:80] + "…"
                tree.insert("", "end", iid=item, values=(
                    marker.marker_type, preview, marker.document_name, marker.location,
                    marker.note, marker.created_at,
                ))

        def open_document() -> None:
            selected = tree.selection()
            if selected:
                marker = markers[selected[0]]
                try:
                    open_at_location(marker.document_path, marker.location)
                except Exception as exc:
                    messagebox.showerror("无法定位原文", str(exc), parent=window)

        def copy_text() -> None:
            selected = tree.selection()
            if selected:
                marker = markers[selected[0]]
                self.root.clipboard_clear()
                self.root.clipboard_append(marker.selected_text)

        def delete() -> None:
            selected = tree.selection()
            if not selected:
                return
            if messagebox.askyesno("删除标记", "确定删除选中的重点标记吗？原始文档不会受影响。", parent=window):
                self.exam_service.delete_marker(int(selected[0]))
                load()

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="打开来源文档", command=open_document).pack(side="left")
        ttk.Button(buttons, text="复制原文", command=copy_text).pack(side="left", padx=7)
        ttk.Button(buttons, text="删除标记", command=delete).pack(side="left")
        ttk.Button(buttons, text="关闭", command=window.destroy).pack(side="right")
        tree.bind("<Double-1>", lambda _event: open_document())
        load()

    def open_exam_module(self) -> None:
        if self.exam_ui and self.exam_ui.get("window"):
            window = self.exam_ui["window"]
            try:
                window.lift()
                window.focus_force()
                return
            except tk.TclError:
                self.exam_ui = None

        window = tk.Toplevel(self.root)
        window.title(f"试卷答案查找与重复考点分析 · {APP_VERSION}")
        center_window(window, 1200, 760)
        window.minsize(920, 620)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=14)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(6, weight=1)
        source_var = tk.StringVar(value="尚未选择试卷（可一次选择多套）")
        status_var = tk.StringVar(value="请选择一套或多套 PDF、Word、Markdown 试卷")
        threshold_var = tk.DoubleVar(value=0.78)
        output_dir_var = tk.StringVar(value=str(Path.home() / "Desktop" / "试卷答案查找结果"))
        selected_paths: list[str] = []
        exam_categories: set[str] = set(self.selected_categories)
        exam_category_text = tk.StringVar(
            value="全部分类" if not exam_categories else f"已选 {len(exam_categories)} 个分类"
        )

        top = ttk.LabelFrame(frame, text="试卷任务", padding=10)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        def choose_exam() -> None:
            paths = filedialog.askopenfilenames(
                title="选择一套或多套试卷",
                filetypes=[
                    ("支持的试卷", "*.pdf *.docx *.doc *.md *.markdown"),
                    ("PDF", "*.pdf"), ("Word", "*.docx *.doc"),
                    ("Markdown", "*.md *.markdown"),
                ],
            )
            if paths:
                selected_paths[:] = paths
                names = "；".join(Path(path).name for path in paths[:3])
                if len(paths) > 3:
                    names += f"；另 {len(paths) - 3} 套"
                source_var.set(f"已选择 {len(paths)} 套：{names}")
                status_var.set(f"已选择 {len(paths)} 套试卷，请点击上方“开始搜索并生成”。")

        choose_button = ttk.Button(top, text="选择多套试卷…", command=choose_exam)
        choose_button.pack(side="left")
        start_button = ttk.Button(top, text="开始搜索并生成", style="Accent.TButton")
        start_button.pack(side="left", padx=12, ipady=5)
        ttk.Label(top, text="相似分组阈值").pack(side="left", padx=(8, 4))
        threshold = ttk.Spinbox(top, from_=0.50, to=1.00, increment=0.01, width=6, textvariable=threshold_var)
        threshold.pack(side="left")
        def choose_exam_categories() -> None:
            picker = tk.Toplevel(window)
            picker.title("选择试卷答案检索分类")
            center_window(picker, 480, 580)
            picker.transient(window)
            picker.grab_set()
            body = ttk.Frame(picker, padding=12)
            body.pack(fill="both", expand=True)
            ttk.Label(body, text="可多选；全部不选表示检索整个学习库。", wraplength=370).pack(anchor="w")
            variables: dict[str, tk.BooleanVar] = {}
            for category in self.database.list_categories():
                variable = tk.BooleanVar(value=category in exam_categories)
                variables[category] = variable
                ttk.Checkbutton(body, text=category, variable=variable).pack(anchor="w", pady=2)
            def apply_categories() -> None:
                exam_categories.clear()
                exam_categories.update(name for name, variable in variables.items() if variable.get())
                exam_category_text.set("全部分类" if not exam_categories else f"已选 {len(exam_categories)} 个分类")
                picker.destroy()
            actions = ttk.Frame(body)
            actions.pack(side="bottom", fill="x", pady=(10, 0))
            ttk.Button(actions, text="应用", command=apply_categories).pack(side="right")
            ttk.Button(actions, text="取消", command=picker.destroy).pack(side="right", padx=7)
        category_button = ttk.Button(top, text="选择答案检索分类…", command=choose_exam_categories)
        source_label = ttk.Label(frame, textvariable=source_var, wraplength=820)
        source_label.grid(row=1, column=0, sticky="ew", pady=(0, 5))

        output_row = ttk.Frame(frame)
        output_row.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        output_row.columnconfigure(1, weight=1)
        ttk.Label(output_row, text="自动生成到：").grid(row=0, column=0)
        output_entry = ttk.Entry(output_row, textvariable=output_dir_var)
        output_entry.grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Label(output_row, text="答案检索分类：").grid(row=1, column=0, pady=(7, 0), sticky="w")
        category_area = ttk.Frame(output_row)
        category_area.grid(row=1, column=1, columnspan=3, sticky="w", padx=6, pady=(7, 0))
        category_button.pack(side="left")
        ttk.Label(category_area, textvariable=exam_category_text, foreground="#145b91").pack(side="left", padx=(8, 0))

        def choose_output_dir() -> None:
            folder = filedialog.askdirectory(
                title="选择答案文档保存文件夹", initialdir=output_dir_var.get(), mustexist=False
            )
            if folder:
                output_dir_var.set(folder)

        folder_button = ttk.Button(output_row, text="选择文件夹…", command=choose_output_dir)
        folder_button.grid(row=0, column=2)

        def open_output_folder() -> None:
            folder = Path(self.exam_ui.get("last_output_dir") or output_dir_var.get())
            if not folder.is_dir():
                messagebox.showinfo("尚未生成", "请先开始搜索；完成后这里可以打开结果文件夹。", parent=window)
                return
            self._open_existing(str(folder), False)

        ttk.Button(output_row, text="打开结果文件夹", command=open_output_folder).grid(row=0, column=3, padx=(7, 0))
        state_row = ttk.Frame(frame)
        state_row.grid(row=3, column=0, sticky="ew", pady=4)
        state_row.columnconfigure(0, weight=1)
        state_label = ttk.Label(state_row, textvariable=status_var, foreground="#1261a0", wraplength=820)
        state_label.grid(row=0, column=0, sticky="ew")
        progress = ttk.Progressbar(state_row, mode="determinate")
        progress.grid(row=1, column=0, sticky="ew", pady=(5, 0))
        footer = ttk.Frame(frame)
        footer.grid(row=4, column=0, sticky="ew", pady=5)
        export_md_button = ttk.Button(footer, text="另存答案 MD", state="disabled")
        export_md_button.pack(side="left")
        export_word_button = ttk.Button(footer, text="另存答案 Word", state="disabled")
        export_word_button.pack(side="left", padx=7)
        export_analysis_button = ttk.Button(footer, text="另存考点分析", state="disabled")
        export_analysis_button.pack(side="left")
        def open_answer() -> None:
            outputs = self.exam_ui.get("last_outputs", ())
            answers = [path for path in outputs if path.suffix.lower() == '.docx']
            if len(answers) == 1:
                self._open_existing(str(answers[0]), False)
            elif answers:
                target = filedialog.askopenfilename(
                    title="选择要打开的单卷答案（每套试卷各一份）", parent=window,
                    initialdir=str(answers[0].parent), filetypes=[("试卷答案 Word", "*.docx")],
                )
                if target:
                    self._open_existing(target, False)
        answer_button = ttk.Button(footer, text="打开单卷答案 Word", command=open_answer, state="disabled")
        answer_button.pack(side="left", padx=10)
        generated_var = tk.StringVar(value="尚未生成文档。选卷后点击顶部按钮即可搜索并自动生成。")
        generated_label = ttk.Label(frame, textvariable=generated_var, wraplength=820, foreground="#53606c")
        generated_label.grid(row=5, column=0, sticky="ew", pady=(0, 5))
        def fit_labels(event) -> None:
            for label in (source_label, state_label, generated_label):
                label.configure(wraplength=max(320, event.width - 24))
        frame.bind("<Configure>", fit_labels)

        columns = ("paper", "number", "stem", "answer", "status", "source", "score")
        results_frame = ttk.Frame(frame)
        results_frame.grid(row=6, column=0, sticky="nsew")
        tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=2)
        for column, title, width in (
            ("paper", "试卷", 150), ("number", "题号", 65), ("stem", "题干", 330),
            ("answer", "答案原文（双击查看）", 350), ("status", "状态", 140),
            ("source", "知识库来源", 220), ("score", "匹配度", 80),
        ):
            tree.heading(column, text=title)
            tree.column(column, width=width, minwidth=55, anchor="w" if column in ("paper", "stem", "answer", "source") else "center")
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(results_frame, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=horizontal.set)
        horizontal.grid(row=1, column=0, sticky="ew")
        results_frame.rowconfigure(0, weight=1)
        results_frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="实时搜索日志", style="Section.TLabel").grid(row=7, column=0, sticky="w", pady=(6, 4))
        log = ScrolledText(frame, height=3, wrap="word", font=("Microsoft YaHei UI", 9))
        log.grid(row=8, column=0, sticky="ew")
        log.configure(state="disabled")

        self.exam_ui = {
            "window": window, "tree": tree, "log": log, "progress": progress,
            "status": status_var, "start": start_button, "choose": choose_button,
            "exports": (export_md_button, export_word_button, export_analysis_button),
            "results": [], "threshold": threshold_var, "output_dir": output_dir_var,
            "running": False, "run_id": None, "generated": generated_var, "answer": answer_button,
            "question_details": {},
            "inputs": (output_entry, folder_button, threshold, category_button),
        }

        def append_log(message: str) -> None:
            log.configure(state="normal")
            log.insert("end", message + "\n")
            log.see("end")
            log.configure(state="disabled")

        self.exam_ui["append_log"] = append_log

        def preview_answer(event) -> None:
            row = tree.identify_row(event.y)
            details = self.exam_ui["question_details"].get(row)
            if not details:
                return
            paper, question, match = details
            preview = tk.Toplevel(window)
            preview.title(f"{paper} · 第{question.number}题：原文答案")
            center_window(preview, 900, 680)
            original = answer_source_text(match) if match else ''
            def copy_original() -> None:
                preview.clipboard_clear()
                preview.clipboard_append(original)
            ttk.Button(preview, text='复制答案原文', command=copy_original,
                       state='normal' if match else 'disabled').pack(anchor='w', padx=12, pady=10)
            text = ScrolledText(preview, wrap='word', font=('Microsoft YaHei UI', 12), spacing1=6, spacing3=12)
            text.pack(fill='both', expand=True, padx=12, pady=(0, 12))
            text.insert('end', f'{question.number}. {question.stem}\n\n答案：\n{original if match else "【无】"}')
            text.insert('end', f'\n\n出处：\n{answer_source_label(match)}')
            text.configure(state='disabled')
        tree.bind('<Double-1>', preview_answer)

        def start() -> None:
            if self.exam_ui["running"]:
                return
            paths = list(selected_paths)
            if not paths:
                messagebox.showinfo("请选择试卷", "请先选择一套或多套 PDF、Word、Markdown 试卷。", parent=window)
                return
            try:
                chosen_threshold = float(threshold_var.get())
                if not math.isfinite(chosen_threshold) or not 0.5 <= chosen_threshold <= 1.0:
                    raise ValueError("重复阈值须在 0.50 到 1.00 之间。")
                if not output_dir_var.get().strip():
                    raise ValueError("请选择结果保存文件夹。")
                output_folder = Path(output_dir_var.get().strip()).resolve()
                output_folder.mkdir(parents=True, exist_ok=True)
            except (ValueError, OSError, tk.TclError) as exc:
                messagebox.showerror("无法开始", str(exc), parent=window)
                return
            run_id = uuid.uuid4().hex
            self.exam_ui.update(running=True, run_id=run_id, results=[], last_outputs=(), question_details={})
            generated_var.set("任务已启动：每套试卷单独生成答案 MD、Word 和分析，原文直接填到对应题目下方。")
            answer_button.configure(state="disabled")
            for widget in self.exam_ui["inputs"]:
                widget.configure(state="disabled")
            tree.delete(*tree.get_children())
            log.configure(state="normal")
            log.delete("1.0", "end")
            log.configure(state="disabled")
            progress.configure(mode="indeterminate")
            progress.start(60)
            status_var.set("正在解析试卷…")
            start_button.configure(state="disabled", text="搜索进行中…")
            choose_button.configure(state="disabled")
            for button in self.exam_ui["exports"]:
                button.configure(state="disabled")

            def emit(event: str, payload: object) -> None:
                # Tag every queued event so it cannot update a different task window.
                self.events.put((event, (run_id, payload)))

            def worker() -> None:
                results: list[ExamTaskResult] = []
                failures: list[tuple[str, str]] = []
                parsed: dict[str, list] = {}
                for index, path in enumerate(paths, 1):
                    emit("exam_stage", f"解析试卷 {index}/{len(paths)}：{Path(path).name}")
                    try:
                        parsed[path] = parse_exam(path)
                        emit("exam_log", f"{Path(path).name}：已解析 {len(parsed[path])} 道题")
                    except Exception as exc:
                        failures.append((Path(path).name, str(exc)))
                        emit("exam_file_error", (Path(path).name, str(exc)))
                total_questions = sum(len(items) for items in parsed.values())
                completed = 0
                for path in paths:
                    if path not in parsed:
                        continue
                    paper_name = Path(path).stem

                    def on_progress(current, _total, question, match, item_status, *, base=completed, paper=paper_name) -> None:
                        emit(
                            "exam_progress",
                            (base + current, total_questions, paper, question, match, item_status),
                        )

                    def on_activity(message, paper=paper_name):
                        emit("exam_stage", f"{paper}：{message}")

                    try:
                        results.append(self.exam_service.process_exam(
                            path, on_progress, activity=on_activity, questions=parsed[path],
                            categories=sorted(exam_categories),
                        ))
                    except Exception as exc:
                        failures.append((Path(path).name, str(exc)))
                        emit("exam_file_error", (Path(path).name, str(exc)))
                    completed += len(parsed[path])
                if results:
                    emit("exam_stage", "搜索结束，正在逐卷填写原文答案，分别生成 Word、Markdown 和分析…")
                    output_files = ()
                    export_error = ""
                    try:
                        # Export in this worker; Tk continues to repaint status and progress.
                        output_files = export_result_bundle(results, output_folder, chosen_threshold)
                    except Exception as exc:
                        export_error = str(exc)
                    emit("exam_done", (results, failures, output_folder, output_files, export_error))
                else:
                    emit("exam_error", "所选试卷均处理失败，请查看实时日志中的具体原因。")

            threading.Thread(target=worker, daemon=True).start()

        start_button.configure(command=start)

        def require_results() -> list[ExamTaskResult] | None:
            results = self.exam_ui.get("results") if self.exam_ui else None
            if not isinstance(results, list) or not results:
                messagebox.showinfo("尚无结果", "请先完成试卷检索。", parent=window)
                return None
            return results

        def export_md() -> None:
            results = require_results()
            if not results:
                return
            if len(results) > 1:
                export_separate_answers(results, 'md')
                return
            target = filedialog.asksaveasfilename(
                title="导出答案 Markdown", defaultextension=".md",
                initialfile=(f"{results[0].task_name}_知识库原文答案.md" if len(results) == 1 else f"{len(results)}套试卷_知识库原文答案.md"),
                filetypes=[("Markdown", "*.md")],
            )
            if target:
                try:
                    export_answer_markdown(results[0], target)
                    messagebox.showinfo("导出完成", target, parent=window)
                except Exception as exc:
                    messagebox.showerror("导出失败", str(exc), parent=window)

        def export_word() -> None:
            results = require_results()
            if not results:
                return
            if len(results) > 1:
                export_separate_answers(results, 'docx')
                return
            target = filedialog.asksaveasfilename(
                title="导出答案 Word", defaultextension=".docx",
                initialfile=(f"{results[0].task_name}_知识库原文答案.docx" if len(results) == 1 else f"{len(results)}套试卷_知识库原文答案.docx"),
                filetypes=[("Word", "*.docx")],
            )
            if target:
                try:
                    export_answer_word(results[0], target)
                    messagebox.showinfo("导出完成", target, parent=window)
                except Exception as exc:
                    messagebox.showerror("导出失败", str(exc), parent=window)

        def export_separate_answers(results: list[ExamTaskResult], file_format: str) -> None:
            target = filedialog.askdirectory(title="选择文件夹：每套试卷分别保存答案", parent=window)
            if not target:
                return
            try:
                paths = export_individual_answers(results, target, file_format)
                messagebox.showinfo("逐卷导出完成", f"已分别生成 {len(paths)} 份答案：\n{paths[0].parent}", parent=window)
            except Exception as exc:
                messagebox.showerror("逐卷导出失败", str(exc), parent=window)

        def export_analysis() -> None:
            results = require_results()
            if not results:
                return
            # A marker may have been added after the task completed; refresh it before analysis export.
            for result in results:
                for match in result.matches.values():
                    match.marker_types = list(dict.fromkeys(
                        marker.marker_type for marker in self.exam_service.task_markers_for_match(match)
                    ))
            target = filedialog.asksaveasfilename(
                title="导出重复考点分析", defaultextension=".md",
                initialfile=(f"{results[0].task_name}_重复考点分析.md" if len(results) == 1 else f"{len(results)}套试卷_综合重复考点分析.md"),
                filetypes=[("Markdown", "*.md")],
            )
            if target:
                if len(results) == 1:
                    export_analysis_markdown(results[0], target, float(threshold_var.get()))
                else:
                    export_batch_analysis_markdown(results, target, float(threshold_var.get()))
                messagebox.showinfo("导出完成", target, parent=window)

        export_md_button.configure(command=export_md)
        export_word_button.configure(command=export_word)
        export_analysis_button.configure(command=export_analysis)

        def closed() -> None:
            if self.exam_ui["running"]:
                messagebox.showinfo("任务正在运行", "请等待搜索和文档生成完成后关闭窗口。", parent=window)
                return
            self.exam_ui = None
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", closed)

    def choose_search_categories(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("选择检索分类（可多选）")
        center_window(window, 500, 600)
        window.transient(self.root)
        window.grab_set()
        frame = ttk.Frame(window, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="勾选一个或多个分类；全部不选表示搜索全部分类。", wraplength=390).pack(anchor="w")
        canvas = tk.Canvas(frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        choices = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=choices, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, pady=10)
        scrollbar.pack(side="right", fill="y", pady=10)
        values: dict[str, tk.BooleanVar] = {}
        for category in self.database.list_categories():
            variable = tk.BooleanVar(value=category in self.selected_categories)
            values[category] = variable
            ttk.Checkbutton(choices, text=category, variable=variable).pack(anchor="w", fill="x", pady=2)
        choices.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))

        def clear_all() -> None:
            for variable in values.values():
                variable.set(False)

        def apply() -> None:
            self.selected_categories = {name for name, variable in values.items() if variable.get()}
            if not self.selected_categories:
                self.category_scope_text.set("全部分类")
            elif len(self.selected_categories) <= 2:
                self.category_scope_text.set("、".join(sorted(self.selected_categories)))
            else:
                self.category_scope_text.set(f"已选 {len(self.selected_categories)} 个分类")
            window.destroy()
            if self.search_text.get().strip():
                self.do_search()

        buttons = ttk.Frame(window, padding=(14, 0, 14, 12))
        buttons.pack(fill="x")
        ttk.Button(buttons, text="清除选择（全部分类）", command=clear_all).pack(side="left")
        ttk.Button(buttons, text="应用", command=apply).pack(side="right")
        ttk.Button(buttons, text="取消", command=window.destroy).pack(side="right", padx=7)

    def _current_document(self):
        hit = self.selected_hit()
        return self.database.get_document(hit.doc_id) if hit else None

    def _require_current_document(self):
        row = self._current_document()
        if not row:
            messagebox.showinfo("请选择文件", "请先在搜索结果中选择一个文件，或在“管理文档与分类”中操作。")
        return row

    def export_current_file(self) -> None:
        row = self._require_current_document()
        if not row:
            return
        source = Path(row["stored_path"])
        target = filedialog.asksaveasfilename(title="导出当前文件", initialfile=row["name"])
        if target:
            shutil.copy2(source, target)
            messagebox.showinfo("导出完成", target)

    def convert_current_document(self) -> None:
        row = self._require_current_document()
        if row:
            self.convert_document_to_pdf(row, self.root)

    def open_current_pdf_for_export(self, annotated: bool) -> None:
        row = self._require_current_document()
        if not row:
            return
        if str(row["suffix"]).lower() != ".pdf":
            messagebox.showinfo("仅支持PDF", "此功能仅用于PDF。其他文件可使用“导出当前文件”。")
            return
        viewer = PDFViewer(self.root, self.database, row)
        self.root.after(100, viewer.save_annotated if annotated else viewer.save_original)

    def rename_current_document(self) -> None:
        row = self._require_current_document()
        if not row:
            return
        new_name = simpledialog.askstring("重命名", "输入新的文件名：", initialvalue=row["name"], parent=self.root)
        if new_name:
            try:
                self.database.rename_document(int(row["id"]), new_name)
                self.refresh_documents()
            except Exception as exc:
                messagebox.showerror("重命名失败", str(exc))

    def move_current_document(self) -> None:
        row = self._require_current_document()
        if not row:
            return
        category = simpledialog.askstring(
            "移动分类", "输入已有分类或新分类名称：", initialvalue=row["category"], parent=self.root
        )
        if category:
            try:
                self.database.move_document_category(int(row["id"]), category)
                self.refresh_documents()
            except Exception as exc:
                messagebox.showerror("移动失败", str(exc))

    def delete_current_document(self) -> None:
        row = self._require_current_document()
        if row and messagebox.askyesno("删除文件", f"确定删除“{row['name']}”及学习库内副本吗？"):
            self.database.delete_documents([int(row["id"])], delete_files=True)
            self.all_hits = [hit for hit in self.all_hits if hit.doc_id != int(row["id"])]
            self.render_page()
            self.refresh_documents()

    def open_pdf_annotation_manager(self) -> None:
        rows = [row for row in self.database.list_documents() if str(row["suffix"]).lower() == ".pdf"]
        if not rows:
            messagebox.showinfo("没有PDF", "学习库中还没有可批注的PDF文件。")
            return
        window = tk.Toplevel(self.root)
        window.title("PDF批注管理")
        center_window(window, 900, 640)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=12)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=("name", "category", "count"), show="headings")
        for key, title, width in (("name", "PDF标题", 480), ("category", "分类", 180), ("count", "批注数", 80)):
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor="w" if key != "count" else "center")
        for row in rows:
            count = len(self.database.list_pdf_annotations(int(row["id"])))
            tree.insert("", "end", iid=str(row["id"]), values=(row["name"], row["category"], count))
        tree.pack(fill="both", expand=True)
        def open_selected(_event=None):
            selected = tree.selection()
            if selected:
                document = self.database.get_document(int(selected[0]))
                if document:
                    PDFViewer(self.root, self.database, document)
        tree.bind("<Double-1>", open_selected)
        ttk.Button(frame, text="打开阅读与批注", command=open_selected).pack(anchor="e", pady=(8, 0), ipadx=8, ipady=4)

    def convert_document_to_pdf(self, row, parent=None) -> None:
        if not row:
            return
        if str(row["suffix"]).lower() not in {".doc", ".docx", ".md", ".markdown"}:
            messagebox.showinfo("无需转换", "请选择 Word 或 Markdown 文件。", parent=parent or self.root)
            return
        try:
            self.status_text.set(f"正在转换：{row['name']}")
            self.root.update_idletasks()
            status, _chunks, output_name = convert_and_import(self.database, row)
            self.refresh_documents()
            if status == "duplicate":
                messagebox.showinfo(
                    "转换完成", f"可搜索PDF已经存在：{output_name}", parent=parent or self.root
                )
            else:
                messagebox.showinfo(
                    "转换完成",
                    f"已保留原文件，并在同一分类生成：\n{output_name}\n\n新PDF包含可选择、可检索的文字层。",
                    parent=parent or self.root,
                )
            self.status_text.set("可搜索PDF转换完成")
        except Exception as exc:
            self.status_text.set("PDF转换失败")
            messagebox.showerror("转换失败", str(exc), parent=parent or self.root)

    def change_library_location(self) -> None:
        if self.busy:
            messagebox.showinfo("请稍候", "当前任务完成后再调整存放位置。")
            return
        selected = filedialog.askdirectory(
            title="选择资料库统一存放位置", initialdir=str(self.database.db_path.parent)
        )
        if not selected:
            return
        target_root = Path(selected).resolve()
        current_root = self.database.db_path.parent.resolve()
        if target_root == current_root:
            return
        if target_root.is_relative_to(current_root):
            messagebox.showwarning("位置不合适", "新位置不能放在当前资料库文件夹内部。")
            return
        target_db = target_root / "学习库.db"
        target_storage = target_root / "资料文件"
        try:
            target_root.mkdir(parents=True, exist_ok=True)
            if target_db.exists():
                if not messagebox.askyesno(
                    "切换资料库",
                    "所选位置已有一个学习库。是否切换到该学习库？\n\n当前资料不会被删除。",
                ):
                    return
            else:
                copy_current = messagebox.askyesnocancel(
                    "调整存放位置",
                    "是否把当前学习库复制到新位置？\n\n"
                    "选择“是”：复制数据库、分类文件和试卷记录后切换；\n"
                    "选择“否”：在新位置创建空学习库；\n"
                    "选择“取消”：不作修改。",
                )
                if copy_current is None:
                    return
                if copy_current:
                    self.database.checkpoint()
                    if self.database.db_path.exists():
                        shutil.copy2(self.database.db_path, target_db)
                    if self.database.storage_dir.exists():
                        shutil.copytree(self.database.storage_dir, target_storage, dirs_exist_ok=True)
            self.database = LibraryDatabase(target_db, target_storage)
            self.exam_service = ExamService(self.database)
            save_library_root(target_root)
            self.location_text.set(f"存放位置：{target_root}")
            self.clear_search_document()
            self.all_hits = []
            self.render_page()
            self.refresh_documents()
            messagebox.showinfo(
                "存放位置已调整",
                f"当前资料库位置：\n{target_root}\n\n旧位置中的文件会保留，可确认无误后自行删除。",
            )
        except Exception as exc:
            messagebox.showerror("调整存放位置失败", str(exc))

    def choose_single_search_document(self) -> None:
        self.find_documents_window(select_for_search=True)

    def clear_search_document(self) -> None:
        self.search_doc_id = 0
        self.search_scope_text.set("查询范围：全部文档")

    def find_documents_window(self, select_for_search: bool = False) -> None:
        window = tk.Toplevel(self.root)
        window.title("选择单个查询文件" if select_for_search else "资料/试卷查找")
        center_window(window, 1120, 720)
        window.minsize(820, 520)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=12)
        frame.pack(fill="both", expand=True)
        query = tk.StringVar()
        top = ttk.Frame(frame)
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="按标题、分类、标签或备注查找").pack(side="left")
        entry = ttk.Entry(top, textvariable=query)
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0), ipady=4)
        style = ttk.Style(window)
        style.configure("DocumentFinder.Treeview", rowheight=48, font=("Microsoft YaHei UI", 10))
        tree = ttk.Treeview(
            frame, columns=("name", "category", "type", "date"), show="headings",
            selectmode="browse", style="DocumentFinder.Treeview",
        )
        tree.heading("name", text="文件标题")
        tree.heading("category", text="分类")
        tree.heading("type", text="类型")
        tree.heading("date", text="添加日期")
        tree.column("name", width=650, minwidth=280, stretch=True)
        tree.column("category", width=170, minwidth=90, stretch=False)
        tree.column("type", width=75, minwidth=60, anchor="center", stretch=False)
        tree.column("date", width=150, minwidth=120, anchor="center", stretch=False)
        tree.tag_configure("even", background="#f1f5f9")
        tree.tag_configure("odd", background="#ffffff")
        tree.pack(fill="both", expand=True)

        def load(_event=None) -> None:
            tree.delete(*tree.get_children())
            for index, row in enumerate(self.database.find_documents(query.get().strip())):
                tree.insert(
                    "", "end", iid=str(row["id"]), tags=("even" if index % 2 == 0 else "odd",),
                    values=(row["name"], row["category"], str(row["suffix"]).lstrip(".").upper(), row["imported_at"]),
                )

        def selected_row():
            selected_items = tree.selection()
            return self.database.get_document(int(selected_items[0])) if selected_items else None

        def open_file() -> None:
            row = selected_row()
            if row:
                self._open_existing(str(row["stored_path"]), False)

        def reveal_file() -> None:
            row = selected_row()
            if row:
                self._open_existing(str(row["stored_path"]), True)

        def use_for_search() -> None:
            row = selected_row()
            if not row:
                messagebox.showinfo("请选择文件", "请先选择一个文件。", parent=window)
                return
            self.search_doc_id = int(row["id"])
            self.search_scope_text.set(f"单文件查询：{row['name']}")
            window.destroy()
            self.search_entry.focus_set()

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="预览文件", command=open_file).pack(side="left", ipadx=8, ipady=5)
        ttk.Button(buttons, text="打开所在文件夹", command=reveal_file).pack(side="left", padx=7, ipadx=8, ipady=5)
        ttk.Button(buttons, text="设为单文件查询范围", command=use_for_search).pack(side="left", ipadx=8, ipady=5)
        ttk.Button(buttons, text="关闭", command=window.destroy).pack(side="right")
        entry.bind("<KeyRelease>", load)
        tree.bind("<Double-1>", lambda _event: use_for_search() if select_for_search else open_file())
        tip = tk.Toplevel(window)
        tip.withdraw()
        tip.overrideredirect(True)
        tip_label = ttk.Label(tip, background="#fff8d8", relief="solid", padding=(6, 3))
        tip_label.pack()
        def hover(event) -> None:
            row_id = tree.identify_row(event.y)
            if row_id and tree.identify_column(event.x) == "#2":
                tip_label.configure(text=tree.item(row_id, "values")[1])
                tip.geometry(f"+{event.x_root + 12}+{event.y_root + 12}")
                tip.deiconify()
            else:
                tip.withdraw()
        tree.bind("<Motion>", hover)
        tree.bind("<Leave>", lambda _event: tip.withdraw())
        load()
        entry.focus_set()

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择学习资料",
            filetypes=[
                ("支持的学习资料", "*.pdf *.doc *.docx *.md *.markdown"),
                ("PDF 文件", "*.pdf"), ("Word 文件", "*.doc *.docx"),
                ("Markdown 文件", "*.md *.markdown"), ("所有文件", "*.*"),
            ],
        )
        if paths:
            self.prepare_import(list(paths), convert_after=True)

    def add_folder(self) -> None:
        path = filedialog.askdirectory(title="选择资料文件夹")
        if path:
            self.prepare_import([path], convert_after=True)

    def prepare_import(self, paths: list[str], convert_after: bool = False) -> None:
        if self.busy:
            return
        self.status_text.set("正在检查重复文件…")
        self.root.update_idletasks()
        try:
            duplicates = self.database.find_duplicates(paths)
        except Exception as exc:
            messagebox.showerror("检查文件失败", str(exc))
            return
        action = "skip"
        if duplicates:
            choice = messagebox.askyesnocancel(
                "发现重复文件",
                f"发现 {len(duplicates)} 个“文件名和内容均相同”的资料。\n\n"
                "选择“是”覆盖；选择“否”跳过；选择“取消”停止导入。",
            )
            if choice is None:
                self.status_text.set("已取消")
                return
            action = "overwrite" if choice else "skip"
        self.start_import(
            paths, duplicate_action=action, category=self.import_category.get(),
            convert_after=convert_after,
        )

    def new_category(self) -> None:
        name = self.prompt_category_name(self.root, "新建资料分类")
        if name is None:
            return
        try:
            category = self.database.create_category(name)
            self.refresh_documents()
            self.import_category.set(category)
            self.status_text.set(f"已新建分类：{category}")
        except Exception as exc:
            messagebox.showerror("无法新建分类", str(exc))

    def prompt_category_name(self, parent, title: str, initial: str = "") -> str | None:
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        center_window(dialog, 620, 320)
        dialog.minsize(520, 260)
        dialog.transient(parent)
        dialog.grab_set()
        body = ttk.Frame(dialog, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="分类名称", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            body, text="请输入新的分类文件夹名称，也可以选择已有分类作为参考。",
            foreground="#65707c", wraplength=560,
        ).pack(anchor="w", pady=(4, 12))
        value = tk.StringVar(value=initial)
        entry = ttk.Entry(body, textvariable=value, font=("Microsoft YaHei UI", 11))
        entry.pack(fill="x", ipady=6)
        chooser = ttk.Combobox(body, state="readonly", values=self.database.list_categories())
        chooser.pack(fill="x", pady=(12, 0), ipady=3)
        chooser.bind("<<ComboboxSelected>>", lambda _event: value.set(chooser.get()))
        result: list[str | None] = [None]

        def accept(_event=None) -> None:
            name = value.get().strip()
            if not name:
                messagebox.showinfo("请输入名称", "分类名称不能为空。", parent=dialog)
                return
            result[0] = name
            dialog.destroy()

        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(20, 0))
        ttk.Button(actions, text="确定", command=accept).pack(side="right", ipadx=12)
        ttk.Button(actions, text="取消", command=dialog.destroy).pack(side="right", padx=8)
        entry.bind("<Return>", accept)
        entry.focus_set()
        dialog.wait_window()
        return result[0]

    def rebuild_all(self) -> None:
        if not self.database.list_documents():
            messagebox.showinfo("学习库为空", "请先添加文件。")
            return
        self.start_import([], rebuild=True)

    def start_import(
        self,
        paths: list[str],
        rebuild: bool = False,
        duplicate_action: str = "skip",
        category: str = "未分类",
        convert_after: bool = False,
    ) -> None:
        if self.busy:
            return
        self.busy = True
        self.set_busy_ui(True)
        self.progress["value"] = 0
        self.status_text.set("正在解析并保存资料…")

        def progress(current, total, path, status, _detail) -> None:
            self.events.put(("progress", (current, total, path.name, status)))

        def worker() -> None:
            try:
                summary = (
                    self.database.rebuild_all(progress)
                    if rebuild
                    else self.database.import_paths(
                        paths, progress, duplicate_action=duplicate_action, category=category
                    )
                )
                summary["converted"] = 0
                summary["conversion_failed"] = []
                if convert_after and not rebuild:
                    convertible = [
                        path for path in collect_supported_files(paths)
                        if path.suffix.lower() in {".doc", ".docx", ".md", ".markdown"}
                    ]
                    for index, source in enumerate(convertible, 1):
                        try:
                            row = self.database.duplicate_for(source)
                            if not row:
                                raise ValueError("原文件导入后未找到对应记录")
                            result, _chunks, output_name = convert_and_import(self.database, row)
                            if result in {"added", "updated"}:
                                summary["converted"] = int(summary["converted"]) + 1
                            self.events.put(("progress", (index, len(convertible), output_name, "converted")))
                        except Exception as exc:
                            reason = (
                                "文件正在被 Word、PDF 预览窗口或其他程序占用，请关闭后重试"
                                if isinstance(exc, PermissionError)
                                else str(exc)
                            )
                            summary["conversion_failed"].append((str(source), reason))
                self.events.put(("import_done", summary))
            except Exception as exc:
                self.events.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def set_busy_ui(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        for button in (
            self.add_files_button, self.add_folder_button, self.manage_button,
            self.rebuild_button, self.backup_button, self.restore_button, self.new_category_button,
            self.location_button, self.single_file_button, self.clear_scope_button,
            self.paper_find_button, self.exam_button, self.marker_list_button,
        ):
            button.configure(state=state)

    def poll_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event.startswith("exam_"):
                    run_id, payload = payload
                    if not self.exam_ui or run_id != self.exam_ui.get("run_id"):
                        continue
                if event == "progress":
                    current, total, name, status = payload
                    self.progress["maximum"] = max(1, total)
                    self.progress["value"] = current
                    labels = {
                        "added": "已加入", "updated": "已更新", "skipped": "未变化",
                        "duplicate": "重复已跳过", "failed": "失败", "converted": "已生成可搜索PDF",
                    }
                    self.status_text.set(f"{current}/{total} {labels.get(status, status)}：{name}")
                elif event == "import_done":
                    self.busy = False
                    self.set_busy_ui(False)
                    self.refresh_documents()
                    failed = payload["failed"]
                    message = (
                        f"新增 {payload['added']} 个，更新 {payload['updated']} 个，"
                        f"重复跳过 {payload.get('duplicate', 0)} 个，未变化 {payload['skipped']} 个，"
                        f"失败 {len(failed)} 个。"
                    )
                    converted = int(payload.get("converted", 0))
                    conversion_failed = payload.get("conversion_failed", [])
                    if converted or conversion_failed:
                        message += f"\n已生成 {converted} 个可搜索PDF，转换失败 {len(conversion_failed)} 个。"
                    self.status_text.set("资料处理完成")
                    if failed or conversion_failed:
                        all_failures = [*failed, *conversion_failed]
                        details = "\n".join(f"• {Path(path).name}：{reason}" for path, reason in all_failures[:8])
                        messagebox.showwarning("处理完成（部分失败）", message + "\n\n" + details)
                    else:
                        messagebox.showinfo("处理完成", message)
                elif event == "error":
                    self.busy = False
                    self.set_busy_ui(False)
                    self.status_text.set("处理失败")
                    messagebox.showerror("处理失败", str(payload))
                elif event == "question_research_done":
                    run_id, research = payload
                    ui = self.question_research_ui
                    if not ui or run_id != ui.get("run_id"):
                        continue
                    assert isinstance(research, QuestionResearchResult)
                    output = ui["output"]
                    output.configure(state="normal")
                    output.delete("1.0", "end")
                    output.insert("1.0", research.render_text())
                    output.configure(state="disabled")
                    ui["result"] = research
                    ui["start"].configure(state="normal", text="检索并整理")
                    ui["copy"].configure(state="normal")
                    for button in ui["exports"]:
                        button.configure(state="normal")
                    ui["status"].set(
                        f"完成：{research.matched_chunks} 条匹配片段，{len(research.statements)} 条去重表述"
                    )
                elif event == "question_research_error":
                    run_id, reason = payload
                    ui = self.question_research_ui
                    if not ui or run_id != ui.get("run_id"):
                        continue
                    ui["start"].configure(state="normal", text="检索并整理")
                    ui["status"].set("检索整理失败")
                    messagebox.showerror("文档问题检索整理失败", str(reason), parent=ui["window"])
                elif event in ("exam_stage", "exam_log"):
                    self.exam_ui["append_log"](str(payload))
                    if event == "exam_stage":
                        self.exam_ui["status"].set(str(payload))
                elif event == "exam_progress":
                    if not self.exam_ui:
                        continue
                    current, total, paper_name, question, match, status = payload
                    tree = self.exam_ui["tree"]
                    progress = self.exam_ui["progress"]
                    status_var = self.exam_ui["status"]
                    progress.stop()
                    progress.configure(mode="determinate")
                    progress["maximum"] = max(1, total)
                    progress["value"] = current
                    preview = question.stem if len(question.stem) <= 90 else question.stem[:90] + "…"
                    source = match.source_name if match else ""
                    score = f"{match.similarity:.0%}" if match else ""
                    tree.insert(
                        "", "end", iid=str(question.id),
                        values=(paper_name, question.number, preview, answer_source_text(match) if match else '【无】', status, source, score),
                    )
                    self.exam_ui["question_details"][str(question.id)] = (paper_name, question, match)
                    tree.see(str(question.id))
                    self.exam_ui["append_log"](f"{paper_name} 第{question.number}题：{status}")
                    status_var.set(f"正在检索 {current}/{total}：{paper_name} 第 {question.number} 题")
                elif event == "exam_file_error":
                    if not self.exam_ui:
                        continue
                    name, reason = payload
                    self.exam_ui["append_log"](f"{name}：解析/处理失败：{reason}")
                elif event == "exam_done":
                    if not self.exam_ui:
                        continue
                    results, failures, output_folder, output_files, auto_export_error = payload
                    self.exam_ui["results"] = results
                    self.exam_ui["running"] = False
                    self.exam_ui["progress"].stop()
                    self.exam_ui["progress"].configure(mode="determinate", maximum=1, value=1)
                    self.exam_ui["start"].configure(state="normal", text="开始搜索并生成")
                    self.exam_ui["choose"].configure(state="normal")
                    for widget in self.exam_ui["inputs"]:
                        widget.configure(state="normal")
                    for button in self.exam_ui["exports"]:
                        button.configure(state="normal")
                    question_count = sum(len(result.questions) for result in results)
                    matched = sum(len(result.matches) for result in results)
                    append_log = self.exam_ui["append_log"]
                    if output_files:
                        output_folder = output_files[0].parent
                        self.exam_ui["last_outputs"] = output_files
                        self.exam_ui["last_output_dir"] = str(output_folder)
                        self.exam_ui["answer"].configure(state="normal")
                        self.exam_ui["generated"].set(f"已生成 {len(output_files)} 份文档：{len(results)} 套试卷的答案分别保存。点击“打开单卷答案 Word”选择试卷。")
                        append_log("===== 自动生成文档 =====")
                        for path in output_files:
                            append_log(str(path))
                        self.exam_ui["status"].set(
                            f"完成并已生成文档：{len(results)} 套、{question_count} 题、命中 {matched} 题"
                        )
                    else:
                        append_log(f"自动生成文档失败：{auto_export_error}")
                        self.exam_ui["generated"].set(f"生成失败：{auto_export_error}。可使用上方“另存”按钮重试。")
                        self.exam_ui["status"].set("检索完成，但自动生成文档失败")
                    messagebox.showinfo(
                        "批量试卷检索完成",
                        f"成功处理 {len(results)} 套、共 {question_count} 道题，命中知识库原文 {matched} 道。\n"
                        f"处理失败 {len(failures)} 套。\n"
                        + (
                            f"\n已逐卷生成 {len(output_files)} 份文档到：\n{output_folder}"
                            if output_files else
                            f"\n检索结果已保存，但自动生成文档失败：\n{auto_export_error}"
                        ),
                        parent=self.exam_ui["window"],
                    )
                elif event == "exam_error":
                    if not self.exam_ui:
                        continue
                    self.exam_ui["running"] = False
                    self.exam_ui["progress"].stop()
                    self.exam_ui["start"].configure(state="normal", text="开始搜索并生成")
                    self.exam_ui["choose"].configure(state="normal")
                    for widget in self.exam_ui["inputs"]:
                        widget.configure(state="normal")
                    self.exam_ui["generated"].set("本次未生成文档，请查看下方失败原因。")
                    self.exam_ui["status"].set("试卷处理失败")
                    messagebox.showerror("试卷处理失败", str(payload), parent=self.exam_ui["window"])
        except queue.Empty:
            pass
        self.root.after(100, self.poll_events)

    def refresh_documents(self) -> None:
        docs, chars = self.database.counts()
        self.stats_text.set(f"{docs} 个文档 · {chars:,} 个文字")
        values = ("全部标签", *self.database.all_tags())
        self.tag_combo["values"] = values
        if self.tag_filter.get() not in values:
            self.tag_filter.set("全部标签")
        categories = self.database.list_categories()
        self.import_category_combo["values"] = categories
        if self.import_category.get() not in categories:
            self.import_category.set(categories[0])
        category_values = ("全部分类", *categories)
        self.category_combo["values"] = category_values
        if self.category_filter.get() not in category_values:
            self.category_filter.set("全部分类")

    def manage_documents_v3(self) -> None:
        """Large, folder-oriented manager with a compact title/category list."""
        window = tk.Toplevel(self.root)
        window.title("管理文档与分类")
        center_window(window, 1360, 820)
        window.minsize(980, 600)
        # Keep normal window controls so users can freely resize and minimize.

        outer = ttk.Frame(window)
        outer.pack(fill="both", expand=True)
        toolbar = ttk.Frame(outer, padding=(12, 9))
        toolbar.pack(fill="x")
        ttk.Separator(outer, orient="horizontal").pack(fill="x")
        content = ttk.Panedwindow(outer, orient="horizontal")
        content.pack(fill="both", expand=True)
        sidebar = ttk.Frame(content, padding=(12, 12, 6, 12))
        panel = ttk.Frame(content, padding=(6, 12, 12, 12))
        content.add(sidebar, weight=1)
        content.add(panel, weight=5)

        ttk.Label(sidebar, text="分类目录", style="Section.TLabel").pack(anchor="w", pady=(0, 8))
        category_tree = ttk.Treeview(sidebar, show="tree", selectmode="browse", height=20)
        category_scroll = ttk.Scrollbar(sidebar, orient="vertical", command=category_tree.yview)
        category_tree.configure(yscrollcommand=category_scroll.set)
        category_tree.pack(side="left", fill="both", expand=True)
        category_scroll.pack(side="right", fill="y")

        header = ttk.Frame(panel)
        header.pack(fill="x", pady=(0, 8))
        title_var = tk.StringVar(value="全部文档")
        count_var = tk.StringVar()
        filter_var = tk.StringVar()
        sort_var = tk.StringVar(value="名称升序")
        ttk.Label(header, textvariable=title_var, style="Section.TLabel").pack(side="left")
        ttk.Label(header, textvariable=count_var, foreground="#65707c").pack(side="left", padx=10)
        ttk.Label(header, text="排序").pack(side="right", padx=(10, 4))
        sort_combo = ttk.Combobox(
            header,
            textvariable=sort_var,
            state="readonly",
            width=13,
            values=("名称升序", "名称降序", "添加日期最新", "添加日期最早"),
        )
        sort_combo.pack(side="right")
        ttk.Label(header, text="标题筛选").pack(side="right", padx=(10, 4))
        filter_entry = ttk.Entry(header, textvariable=filter_var, width=28)
        filter_entry.pack(side="right")

        style = ttk.Style(window)
        style.configure("CompactManager.Treeview", rowheight=48, font=("Microsoft YaHei UI", 11))
        style.configure("CompactManager.Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))
        tree_frame = ttk.Frame(panel)
        tree_frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(
            tree_frame,
            columns=("name", "category", "type", "date"),
            show="headings",
            selectmode="extended",
            style="CompactManager.Treeview",
        )
        tree.heading("name", text="文件标题")
        tree.heading("category", text="分类")
        tree.heading("type", text="类型")
        tree.heading("date", text="添加日期")
        tree.column("name", width=760, minwidth=280, anchor="w", stretch=True)
        tree.column("category", width=170, minwidth=90, anchor="w", stretch=False)
        tree.column("type", width=80, minwidth=65, anchor="center", stretch=False)
        tree.column("date", width=155, minwidth=125, anchor="center", stretch=False)
        vertical = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vertical.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        tree.tag_configure("even", background="#f3f6fa")
        tree.tag_configure("odd", background="#ffffff")

        sort_keys = {
            "名称升序": "name_asc", "名称降序": "name_desc",
            "添加日期最新": "date_desc", "添加日期最早": "date_asc",
        }

        def current_category() -> str:
            selected = category_tree.selection()
            if not selected or selected[0] == "all":
                return ""
            item = selected[0]
            return item[5:] if item.startswith("cat::") else ""

        def load_categories(select_category: str = "") -> None:
            category_tree.delete(*category_tree.get_children())
            category_tree.insert("", "end", iid="all", text="▣  全部文档", open=True)
            for category in self.database.list_categories():
                category_tree.insert("", "end", iid=f"cat::{category}", text=f"▤  {category}")
            category_tree.insert("", "end", iid="markers", text="★  重点标记清单")
            wanted = f"cat::{select_category}" if select_category else "all"
            if not category_tree.exists(wanted):
                wanted = "all"
            category_tree.selection_set(wanted)
            category_tree.focus(wanted)

        def load_rows(_event=None) -> None:
            category = current_category()
            title_var.set(category or "全部文档")
            query = filter_var.get().strip().casefold()
            rows = self.database.list_documents(category, sort_keys.get(sort_var.get(), "name_asc"))
            if query:
                rows = [row for row in rows if query in str(row["name"]).casefold()]
            tree.delete(*tree.get_children())
            for index, row in enumerate(rows):
                tree.insert(
                    "", "end", iid=str(row["id"]),
                    tags=("even" if index % 2 == 0 else "odd",),
                    values=(row["name"], row["category"], str(row["suffix"]).lstrip(".").upper(), row["imported_at"]),
                )
            count_var.set(f"{len(rows)} 个文件 · Ctrl+A 全选当前列表")

        def category_changed(_event=None) -> None:
            selected = category_tree.selection()
            if selected and selected[0] == "markers":
                self.show_marker_list()
                category_tree.selection_set("all")
            load_rows()

        def selected_ids(require_one: bool = False) -> list[int]:
            ids = [int(item) for item in tree.selection()]
            if not ids:
                messagebox.showinfo("请选择文件", "请先在列表中选择文件。", parent=window)
                return []
            if require_one and len(ids) != 1:
                messagebox.showinfo("请选择一个文件", "此操作一次只能选择一个文件。", parent=window)
                return []
            return ids

        def selected_document():
            ids = selected_ids(require_one=True)
            return self.database.get_document(ids[0]) if ids else None

        def open_selected() -> None:
            row = selected_document()
            if row:
                self._open_existing(str(row["stored_path"]), False)

        def reveal_selected() -> None:
            row = selected_document()
            if row:
                self._open_existing(str(row["stored_path"]), True)

        def open_category_folder() -> None:
            category = current_category()
            if not category:
                row = selected_document()
                category = str(row["category"]) if row else self.import_category.get()
            folder = self.database.storage_dir / self.database.clean_category(category)
            folder.mkdir(parents=True, exist_ok=True)
            open_path(str(folder))

        def set_single_scope() -> None:
            row = selected_document()
            if row:
                self.search_doc_id = int(row["id"])
                self.search_scope_text.set(f"单文件查询：{row['name']}")
                window.destroy()
                self.search_entry.focus_set()

        def convert_selected() -> None:
            row = selected_document()
            if row:
                self.convert_document_to_pdf(row, window)
                load_rows()

        def prompt_existing(title: str, prompt: str, current: str, existing: list[str], multi: bool) -> str | None:
            dialog = tk.Toplevel(window)
            dialog.title(title)
            center_window(dialog, 620, 320)
            dialog.transient(window)
            dialog.grab_set()
            body = ttk.Frame(dialog, padding=14)
            body.pack(fill="both", expand=True)
            ttk.Label(body, text=prompt).pack(anchor="w")
            value = tk.StringVar(value=current)
            entry = ttk.Entry(body, textvariable=value)
            entry.pack(fill="x", pady=(7, 10), ipady=4)
            pick = tk.StringVar(value=existing[0] if existing else "")
            chooser = ttk.Combobox(body, textvariable=pick, values=existing, state="readonly")
            chooser.pack(fill="x")
            result: list[str | None] = [None]

            def use_existing(_event=None) -> None:
                chosen = pick.get().strip()
                if not chosen:
                    return
                if multi:
                    parts = [part.strip() for part in value.get().replace("，", ",").split(",") if part.strip()]
                    if chosen not in parts:
                        parts.append(chosen)
                    value.set(",".join(parts))
                else:
                    value.set(chosen)

            def accept() -> None:
                result[0] = value.get().strip()
                dialog.destroy()

            chooser.bind("<<ComboboxSelected>>", use_existing)
            actions = ttk.Frame(body)
            actions.pack(fill="x", pady=(12, 0))
            ttk.Button(actions, text="使用所选项", command=use_existing).pack(side="left")
            ttk.Button(actions, text="确定", command=accept).pack(side="right")
            ttk.Button(actions, text="取消", command=dialog.destroy).pack(side="right", padx=7)
            entry.focus_set()
            dialog.wait_window()
            return result[0]

        def edit_tags() -> None:
            row = selected_document()
            if not row:
                return
            tags = prompt_existing(
                "编辑标签", "可直接输入；多个标签用逗号分隔，也可从已有标签中选择：",
                str(row["tags"]), self.database.all_tags(), True,
            )
            if tags is not None:
                self.database.update_tags(int(row["id"]), tags)
                self.refresh_documents()

        def edit_note() -> None:
            row = selected_document()
            if not row:
                return
            note = simpledialog.askstring(
                "文件备注", "请输入文件备注：", initialvalue=row["note"], parent=window
            )
            if note is not None:
                self.database.update_note(int(row["id"]), note)

        def move_category() -> None:
            ids = selected_ids()
            if not ids:
                return
            target = prompt_existing(
                "移动到分类", "可输入新分类，也可选择已有分类：", current_category(),
                self.database.list_categories(), False,
            )
            if target is None or not target.strip():
                return
            try:
                for doc_id in ids:
                    self.database.move_document_category(doc_id, target)
                self.refresh_documents()
                load_categories(self.database.clean_category(target))
                load_rows()
            except Exception as exc:
                messagebox.showerror("移动失败", str(exc), parent=window)

        def show_details() -> None:
            row = selected_document()
            if not row:
                return
            messagebox.showinfo(
                "文档详细信息",
                f"标题：{row['name']}\n分类：{row['category']}\n类型：{row['suffix']}\n"
                f"添加日期：{row['imported_at']}\n文字数：{int(row['char_count']):,}\n"
                f"标签：{row['tags'] or '无'}\n备注：{row['note'] or '无'}\n\n"
                f"存放位置：{row['stored_path']}",
                parent=window,
            )

        def remove_selected() -> None:
            ids = selected_ids()
            if not ids:
                return
            if not messagebox.askyesno(
                "删除文件", f"确定删除选中的 {len(ids)} 个文件及资料库内副本吗？", parent=window
            ):
                return
            self.database.delete_documents(ids, delete_files=True)
            self.refresh_documents()
            load_rows()

        def add_files_here() -> None:
            category = current_category()
            if category:
                self.import_category.set(category)
            self.add_files()

        def add_folder_here() -> None:
            category = current_category()
            if category:
                self.import_category.set(category)
            self.add_folder()

        def create_category() -> None:
            name = prompt_existing(
                "新建分类", "输入新分类名称，或选择已有分类：", "",
                self.database.list_categories(), False,
            )
            if name:
                category = self.database.create_category(name)
                self.refresh_documents()
                load_categories(category)
                load_rows()

        def delete_current_category() -> None:
            category = current_category()
            if not category:
                messagebox.showinfo("请选择分类", "请先在左侧选择需要删除的分类。", parent=window)
                return
            if category == "未分类":
                messagebox.showinfo("不能删除", "“未分类”是系统默认分类，不能删除。", parent=window)
                return
            count = len(self.database.list_documents(category))
            if not messagebox.askyesno(
                "删除分类",
                f"确定删除分类“{category}”吗？\n\n"
                f"分类中的 {count} 个文档不会被删除，将自动移动到“未分类”。",
                parent=window,
            ):
                return
            try:
                moved = self.database.delete_category(category)
                self.refresh_documents()
                load_categories()
                load_rows()
                messagebox.showinfo(
                    "分类已删除", f"分类“{category}”已删除，{moved} 个文档已移至“未分类”。", parent=window
                )
            except Exception as exc:
                messagebox.showerror("删除分类失败", str(exc), parent=window)

        def select_all_visible(_event=None):
            items = tree.get_children()
            if items:
                tree.selection_set(items)
            return "break"

        ttk.Button(toolbar, text="＋ 新建分类", command=create_category).pack(side="left", ipady=4)
        ttk.Button(toolbar, text="删除分类", command=delete_current_category).pack(side="left", padx=(8, 0), ipady=4)
        ttk.Button(toolbar, text="＋ 添加文件", command=add_files_here).pack(side="left", padx=(8, 0), ipady=4)
        ttk.Button(toolbar, text="批量添加文件夹", command=add_folder_here).pack(side="left", padx=8, ipady=4)
        ttk.Button(toolbar, text="打开当前分类目录", command=open_category_folder).pack(side="left", ipady=4)
        ttk.Button(toolbar, text="刷新", command=lambda: (load_categories(current_category()), load_rows())).pack(side="right")

        context = tk.Menu(tree, tearoff=False)
        context.add_command(label="打开文件", command=open_selected)
        context.add_command(label="打开文件所在文件夹", command=reveal_selected)
        context.add_command(label="打开所属分类目录", command=open_category_folder)
        context.add_command(label="只查询此文件", command=set_single_scope)
        context.add_command(label="转换为可搜索PDF", command=convert_selected)
        context.add_separator()
        context.add_command(label="查看详细信息", command=show_details)
        context.add_command(label="添加或修改备注", command=edit_note)
        context.add_command(label="编辑标签", command=edit_tags)
        context.add_command(label="移动到分类", command=move_category)
        context.add_separator()
        context.add_command(label="删除文件", command=remove_selected)

        category_context = tk.Menu(category_tree, tearoff=False)
        category_context.add_command(label="打开分类目录", command=open_category_folder)
        category_context.add_separator()
        category_context.add_command(label="删除分类（文档移至未分类）", command=delete_current_category)

        def show_category_context(event) -> None:
            row = category_tree.identify_row(event.y)
            if row:
                category_tree.selection_set(row)
                category_tree.focus(row)
                category_context.tk_popup(event.x_root, event.y_root)

        def show_context(event) -> None:
            row = tree.identify_row(event.y)
            if row:
                if row not in tree.selection():
                    tree.selection_set(row)
                tree.focus(row)
                context.tk_popup(event.x_root, event.y_root)

        hover_tip = tk.Toplevel(window)
        hover_tip.withdraw()
        hover_tip.overrideredirect(True)
        hover_label = ttk.Label(hover_tip, background="#fff8d8", relief="solid", padding=(6, 3))
        hover_label.pack()
        def show_category_tip(event) -> None:
            row_id = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            if row_id and column == "#2":
                category = str(tree.item(row_id, "values")[1])
                hover_label.configure(text=category)
                hover_tip.geometry(f"+{event.x_root + 12}+{event.y_root + 12}")
                hover_tip.deiconify()
            else:
                hover_tip.withdraw()
        tree.bind("<Motion>", show_category_tip)
        tree.bind("<Leave>", lambda _event: hover_tip.withdraw())

        category_tree.bind("<<TreeviewSelect>>", category_changed)
        category_tree.bind("<Double-1>", lambda _event: open_category_folder())
        category_tree.bind("<Button-3>", show_category_context)
        filter_entry.bind("<KeyRelease>", load_rows)
        sort_combo.bind("<<ComboboxSelected>>", load_rows)
        tree.bind("<Double-1>", lambda _event: open_selected())
        tree.bind("<Button-3>", show_context)
        tree.bind("<Control-a>", select_all_visible)
        tree.bind("<Control-A>", select_all_visible)
        window.bind("<Control-a>", select_all_visible)
        window.bind("<Control-A>", select_all_visible)
        load_categories()
        load_rows()
        tree.focus_set()

    def manage_documents_v2(self) -> None:
        """Reference-inspired manager: top actions, left folders, right document list."""
        window = tk.Toplevel(self.root)
        window.title("管理文档与分类")
        center_window(window, 1280, 720)
        window.minsize(980, 560)
        window.transient(self.root)

        outer = ttk.Frame(window)
        outer.pack(fill="both", expand=True)
        toolbar = ttk.Frame(outer, padding=(10, 8))
        toolbar.pack(fill="x")
        ttk.Separator(outer, orient="horizontal").pack(fill="x")

        content = ttk.Panedwindow(outer, orient="horizontal")
        content.pack(fill="both", expand=True)
        sidebar = ttk.Frame(content, padding=(10, 10, 6, 10))
        documents_panel = ttk.Frame(content, padding=(6, 10, 10, 10))
        content.add(sidebar, weight=1)
        content.add(documents_panel, weight=5)

        ttk.Label(sidebar, text="我的文库", style="Section.TLabel").pack(anchor="w", pady=(0, 8))
        category_tree = ttk.Treeview(sidebar, show="tree", selectmode="browse", height=18)
        category_scroll = ttk.Scrollbar(sidebar, orient="vertical", command=category_tree.yview)
        category_tree.configure(yscrollcommand=category_scroll.set)
        category_tree.pack(side="left", fill="both", expand=True)
        category_scroll.pack(side="right", fill="y")

        list_header = ttk.Frame(documents_panel)
        list_header.pack(fill="x", pady=(0, 7))
        title_var = tk.StringVar(value="全部文档")
        count_var = tk.StringVar(value="")
        filter_var = tk.StringVar()
        ttk.Label(list_header, textvariable=title_var, style="Section.TLabel").pack(side="left")
        ttk.Label(list_header, textvariable=count_var, foreground="#65707c").pack(side="left", padx=10)
        ttk.Label(list_header, text="搜索").pack(side="right", padx=(8, 4))
        filter_entry = ttk.Entry(list_header, textvariable=filter_var, width=28)
        filter_entry.pack(side="right")

        columns = ("name", "type", "category", "imported", "chars", "tags", "note", "path")
        table_frame = ttk.Frame(documents_panel)
        table_frame.pack(fill="both", expand=True)
        style = ttk.Style(window)
        style.configure("LibraryManager.Treeview", rowheight=max(68, self.list_rowheight), font=("Microsoft YaHei UI", 10))
        style.configure("LibraryManager.Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))
        tree = ttk.Treeview(
            table_frame, columns=columns, show="headings", selectmode="extended",
            style="LibraryManager.Treeview",
        )
        for column, title, width in (
            ("name", "标题", 330), ("type", "类型", 65), ("category", "分类", 115),
            ("imported", "添加日期", 145), ("chars", "文字数", 80), ("tags", "标签", 130),
            ("note", "备注", 190), ("path", "本地存储位置", 330),
        ):
            tree.heading(column, text=title)
            tree.column(column, width=width, minwidth=55, anchor="w")
        vertical = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        tree.tag_configure("even", background="#f5f7fa")
        tree.tag_configure("odd", background="#ffffff")

        selected_category = tk.StringVar(value="")

        def current_category() -> str:
            selected = category_tree.selection()
            if not selected or selected[0] == "all":
                return ""
            item = selected[0]
            return item[5:] if item.startswith("cat::") else ""

        def load_categories() -> None:
            current = current_category()
            category_tree.delete(*category_tree.get_children())
            category_tree.insert("", "end", iid="all", text="▣  全部文档", open=True)
            for category in self.database.list_categories():
                category_tree.insert("", "end", iid=f"cat::{category}", text=f"▤  {category}")
            category_tree.insert("", "end", iid="markers", text="★  重点标记清单")
            wanted = f"cat::{current}" if current and category_tree.exists(f"cat::{current}") else "all"
            category_tree.selection_set(wanted)
            category_tree.focus(wanted)

        def load_rows(_event=None) -> None:
            category = current_category()
            selected_category.set(category)
            title_var.set(category or "全部文档")
            query = filter_var.get().strip().casefold()
            rows = []
            for row in self.database.list_documents():
                if category and str(row["category"]) != category:
                    continue
                haystack = " ".join((str(row["name"]), str(row["tags"]), str(row["note"]))).casefold()
                if query and query not in haystack:
                    continue
                rows.append(row)
            tree.delete(*tree.get_children())
            for index, row in enumerate(rows):
                tree.insert("", "end", iid=str(row["id"]), tags=("even" if index % 2 == 0 else "odd",), values=(
                    row["name"], row["suffix"].lstrip(".").upper(), row["category"],
                    row["imported_at"], f"{int(row['char_count']):,}", row["tags"], row["note"], row["stored_path"],
                ))
            count_var.set(f"{len(rows)} 个条目")

        def category_changed(_event=None) -> None:
            selected = category_tree.selection()
            if selected and selected[0] == "markers":
                self.show_marker_list()
                category_tree.selection_set("all")
            load_rows()

        def selected_one(show=True) -> str | None:
            items = tree.selection()
            if len(items) != 1:
                if show:
                    messagebox.showinfo("请选择一个文档", "请在右侧列表中选择一个文档。", parent=window)
                return None
            return items[0]

        def open_selected() -> None:
            item = selected_one()
            if item:
                self._open_existing(str(tree.item(item, "values")[7]), False)

        def reveal_selected() -> None:
            item = selected_one()
            if item:
                self._open_existing(str(tree.item(item, "values")[7]), True)

        def open_category() -> None:
            category = current_category() or self.import_category.get()
            folder = self.database.storage_dir / self.database.clean_category(category)
            folder.mkdir(parents=True, exist_ok=True)
            open_path(str(folder))

        def add_files_here() -> None:
            category = current_category()
            if category:
                self.import_category.set(category)
            self.add_files()

        def add_folder_here() -> None:
            category = current_category()
            if category:
                self.import_category.set(category)
            self.add_folder()

        def create_category() -> None:
            name = simpledialog.askstring("新建分类", "请输入分类名称：", parent=window)
            if name is not None:
                category = self.database.create_category(name)
                self.refresh_documents()
                load_categories()
                item = f"cat::{category}"
                if category_tree.exists(item):
                    category_tree.selection_set(item)
                load_rows()

        def edit_tags() -> None:
            item = selected_one()
            if not item:
                return
            values = tree.item(item, "values")
            tags = simpledialog.askstring("编辑标签", "多个标签用逗号分隔：", initialvalue=values[5], parent=window)
            if tags is not None:
                self.database.update_tags(int(item), tags)
                self.refresh_documents()
                load_rows()

        def edit_note() -> None:
            item = selected_one()
            if not item:
                return
            values = tree.item(item, "values")
            note = simpledialog.askstring("文档备注", "请输入备注：", initialvalue=values[6], parent=window)
            if note is not None:
                self.database.update_note(int(item), note)
                load_rows()

        def move_category() -> None:
            items = tree.selection()
            if not items:
                return
            target = simpledialog.askstring("移动到分类", "输入目标分类名称：", parent=window)
            if target is None:
                return
            try:
                for item in items:
                    self.database.move_document_category(int(item), target)
                self.refresh_documents()
                load_categories()
                load_rows()
            except Exception as exc:
                messagebox.showerror("移动失败", str(exc), parent=window)

        def remove_selected() -> None:
            items = tree.selection()
            if not items:
                return
            if not messagebox.askyesno(
                "删除资料", f"确定删除选中的 {len(items)} 个文档及学习库内副本吗？", parent=window
            ):
                return
            self.database.delete_documents((int(item) for item in items), delete_files=True)
            self.refresh_documents()
            load_categories()
            load_rows()

        ttk.Button(toolbar, text="＋ 新建分类", command=create_category).pack(side="left")
        ttk.Button(toolbar, text="＋ 添加文件", command=add_files_here).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="批量添加文件夹", command=add_folder_here).pack(side="left", padx=8)
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=5)
        ttk.Button(toolbar, text="打开文档", command=open_selected).pack(side="left")
        ttk.Button(toolbar, text="打开所在位置", command=reveal_selected).pack(side="left", padx=8)
        ttk.Button(toolbar, text="标签", command=edit_tags).pack(side="left")
        ttk.Button(toolbar, text="备注", command=edit_note).pack(side="left", padx=8)
        ttk.Button(toolbar, text="移动分类", command=move_category).pack(side="left")
        ttk.Button(toolbar, text="删除", command=remove_selected).pack(side="left", padx=8)
        ttk.Button(toolbar, text="刷新", command=lambda: (load_categories(), load_rows())).pack(side="right")

        context = tk.Menu(tree, tearoff=False)
        context.add_command(label="打开文档", command=open_selected)
        context.add_command(label="打开所在位置", command=reveal_selected)
        context.add_separator()
        context.add_command(label="编辑标签", command=edit_tags)
        context.add_command(label="编辑备注", command=edit_note)
        context.add_command(label="移动到分类", command=move_category)
        context.add_separator()
        context.add_command(label="删除", command=remove_selected)

        def show_context(event) -> None:
            row = tree.identify_row(event.y)
            if row:
                if row not in tree.selection():
                    tree.selection_set(row)
                context.tk_popup(event.x_root, event.y_root)

        category_tree.bind("<<TreeviewSelect>>", category_changed)
        category_tree.bind("<Double-1>", lambda _event: open_category())
        filter_entry.bind("<KeyRelease>", load_rows)
        tree.bind("<Double-1>", lambda _event: open_selected())
        tree.bind("<Button-3>", show_context)
        load_categories()
        load_rows()

    def manage_documents(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("管理导入的文件、分类与备注")
        center_window(window, 1180, 700)
        window.minsize(900, 520)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=12)
        frame.pack(fill="both", expand=True)

        manage_style = ttk.Style(window)
        manage_style.configure("Manage.Treeview", rowheight=max(72, self.list_rowheight), font=("Microsoft YaHei UI", 10))
        manage_style.configure("Manage.Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))

        header = ttk.Frame(frame)
        header.pack(fill="x", pady=(0, 8))
        ttk.Label(header, text="已导入资料", style="Section.TLabel").pack(side="left")
        ttk.Label(header, text="打开分类目录").pack(side="right", padx=(10, 5))
        category_choice = tk.StringVar(value=self.import_category.get())
        category_picker = ttk.Combobox(
            header, textvariable=category_choice, state="readonly", width=16,
            values=self.database.list_categories(),
        )
        category_picker.pack(side="right")
        ttk.Button(
            header,
            text="打开目录",
            command=lambda: open_category(category_choice.get()),
        ).pack(side="right", padx=(0, 6))

        columns = ("name", "category", "type", "imported", "chars", "tags", "note", "path")
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings", selectmode="extended", style="Manage.Treeview"
        )
        for column, title, width in (
            ("name", "文件名", 210), ("category", "分类文件夹", 110), ("type", "类型", 55),
            ("imported", "导入时间", 145), ("chars", "文字数", 80),
            ("tags", "标签", 130), ("note", "文件备注", 220), ("path", "本地存储位置", 390),
        ):
            tree.heading(column, text=title)
            tree.column(column, width=width, minwidth=50)
        vertical = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        def load_rows() -> None:
            tree.delete(*tree.get_children())
            for row in self.database.list_documents():
                tree.insert("", "end", iid=str(row["id"]), values=(
                    row["name"], row["category"], row["suffix"], row["imported_at"],
                    row["char_count"], row["tags"], row["note"], row["stored_path"],
                ))

        def one_selected(show_message: bool = True) -> str | None:
            selected = tree.selection()
            if len(selected) != 1:
                if show_message:
                    messagebox.showinfo("选择一个文件", "请先选择一个文件。", parent=window)
                return None
            return selected[0]

        def open_selected_file() -> None:
            item = one_selected()
            if item:
                self._open_existing(str(tree.item(item, "values")[7]), False)

        def reveal_selected_file() -> None:
            item = one_selected()
            if item:
                self._open_existing(str(tree.item(item, "values")[7]), True)

        def open_category(category: str) -> None:
            folder = self.database.storage_dir / self.database.clean_category(category)
            try:
                folder.mkdir(parents=True, exist_ok=True)
                open_path(str(folder))
            except Exception as exc:
                messagebox.showerror("无法打开分类目录", str(exc), parent=window)

        def open_selected_category() -> None:
            item = one_selected(False)
            if item:
                open_category(str(tree.item(item, "values")[1]))
            else:
                open_category(category_choice.get())

        def remove_selected() -> None:
            selected = tree.selection()
            if not selected:
                return
            if not messagebox.askyesno(
                "删除资料",
                f"确定删除选中的 {len(selected)} 个文档吗？\n\n"
                "索引和学习库中的本地文件副本都会删除，此操作不可撤销。",
                parent=window,
            ):
                return
            self.database.delete_documents((int(item) for item in selected), delete_files=True)
            load_rows()
            self.refresh_documents()
            self.all_hits = []
            self.render_page()

        def edit_tags() -> None:
            item = one_selected()
            if not item:
                return
            values = tree.item(item, "values")
            tags = simpledialog.askstring(
                "编辑标签", "输入标签，多个标签用逗号分隔：",
                initialvalue=values[5] if len(values) > 5 else "", parent=window,
            )
            if tags is not None:
                self.database.update_tags(int(item), tags)
                load_rows()
                self.refresh_documents()

        def edit_note() -> None:
            item = one_selected()
            if not item:
                return
            values = tree.item(item, "values")
            note = simpledialog.askstring(
                "添加或修改文件备注",
                "请输入该文件的备注：",
                initialvalue=values[6] if len(values) > 6 else "",
                parent=window,
            )
            if note is not None:
                self.database.update_note(int(item), note)
                load_rows()

        def move_category() -> None:
            selected = tree.selection()
            if not selected:
                messagebox.showinfo("请选择文档", "请先选择要移动的文档。", parent=window)
                return
            category = simpledialog.askstring(
                "移动到分类",
                "输入目标分类文件夹名称（不存在时自动创建）：",
                parent=window,
            )
            if category is None:
                return
            try:
                for item in selected:
                    self.database.move_document_category(int(item), category)
                load_rows()
                self.refresh_documents()
                category_picker.configure(values=self.database.list_categories())
            except Exception as exc:
                messagebox.showerror("移动失败", str(exc), parent=window)

        def show_context_menu(event) -> None:
            row = tree.identify_row(event.y)
            if row:
                if row not in tree.selection():
                    tree.selection_set(row)
                tree.focus(row)
                context_menu.tk_popup(event.x_root, event.y_root)

        tree.bind("<Double-1>", lambda _event: open_selected_file())
        tree.bind("<Button-3>", show_context_menu)

        context_menu = tk.Menu(tree, tearoff=False)
        context_menu.add_command(label="打开文件", command=open_selected_file)
        context_menu.add_command(label="打开文件所在文件夹", command=reveal_selected_file)
        context_menu.add_command(label="打开所属分类目录", command=open_selected_category)
        context_menu.add_separator()
        context_menu.add_command(label="添加或修改文件备注", command=edit_note)
        context_menu.add_command(label="编辑标签", command=edit_tags)
        context_menu.add_command(label="移动到其他分类", command=move_category)
        context_menu.add_separator()
        context_menu.add_command(label="删除文件", command=remove_selected)

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="打开文件", command=open_selected_file).pack(side="left")
        ttk.Button(buttons, text="打开所在文件夹", command=reveal_selected_file).pack(side="left", padx=6)
        ttk.Button(buttons, text="打开分类目录", command=open_selected_category).pack(side="left")
        ttk.Button(buttons, text="文件备注", command=edit_note).pack(side="left", padx=6)
        ttk.Button(buttons, text="编辑标签", command=edit_tags).pack(side="left")
        ttk.Button(buttons, text="移动到分类", command=move_category).pack(side="left", padx=6)
        ttk.Button(buttons, text="删除文件", command=remove_selected).pack(side="left")
        ttk.Button(buttons, text="关闭", command=window.destroy).pack(side="right")
        category_picker.bind("<Double-1>", lambda _event: open_category(category_choice.get()))
        load_rows()

    def export_backup(self) -> None:
        if not self.database.list_documents():
            messagebox.showinfo("学习库为空", "当前没有可备份的资料。")
            return
        target = filedialog.asksaveasfilename(
            title="导出学习库备份", defaultextension=".zip",
            initialfile="智能文档学习库备份.zip", filetypes=[("ZIP 备份包", "*.zip")],
        )
        if not target:
            return
        try:
            self.database.export_backup(Path(target))
            messagebox.showinfo("备份完成", f"原始文件和全部解析文本已保存到：\n{target}")
        except Exception as exc:
            messagebox.showerror("备份失败", str(exc))

    def restore_backup(self) -> None:
        source = filedialog.askopenfilename(title="选择学习库备份", filetypes=[("ZIP 备份包", "*.zip")])
        if not source:
            return
        if not messagebox.askyesno("恢复备份", "遇到相同文件时会覆盖现有索引，是否继续？"):
            return
        if self.busy:
            return
        self.busy = True
        self.set_busy_ui(True)
        self.status_text.set("正在恢复备份…")

        def progress(current, total, path, status, _detail) -> None:
            self.events.put(("progress", (current, total, path.name, status)))

        def worker() -> None:
            try:
                summary = self.database.restore_backup(Path(source), progress)
                self.events.put(("import_done", summary))
            except Exception as exc:
                self.events.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()


def main() -> None:
    enable_dpi_awareness()
    root = tk.Tk()
    LearningLibraryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
