from __future__ import annotations

import json
import shutil
import tkinter as tk
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

import pymupdf as fitz
from PIL import Image, ImageTk

from pdf_viewer import export_annotated_pdf


COLORS = ("#ffd400", "#ff6666", "#55aa33", "#35a7d8", "#8f7bd9", "#d35bd3", "#f19a38", "#9e9e9e")
DEFAULT_RENDER_ZOOM = 3.0
TOOL_LABELS = {
    "select": "选择文字", "highlight": "高亮", "underline": "下划线",
    "ink": "手写", "note": "备注", "eraser": "擦除",
}


_SHARED_WORKSPACES: dict[int, "PDFWorkspace"] = {}


def center_window(window, width: int, height: int) -> None:
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    width = min(width, max(760, screen_width - 60))
    height = min(height, max(560, screen_height - 80))
    x = max(0, (screen_width - width) // 2)
    y = max(0, (screen_height - height) // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")
    window.resizable(True, True)


class PDFWorkspace:
    """One reader window containing browser-style document tabs."""

    def __init__(self, parent, registered_key: int | None = None) -> None:
        self.parent = parent.winfo_toplevel()
        self.registered_key = registered_key
        self.database = None
        self.views: dict[str, PDFDocumentView] = {}
        self.document_tabs: dict[tuple[str, int], str] = {}
        self.window = tk.Toplevel(self.parent)
        self.window.title("PDF阅读与批注")
        center_window(self.window, 1480, 920)
        self.window.minsize(980, 680)
        self.window.protocol("WM_DELETE_WINDOW", self.close)

        header = ttk.Frame(self.window, padding=(9, 6))
        header.pack(fill="x")
        ttk.Label(header, text="PDF 阅读工作区", font=("Microsoft YaHei UI", 11, "bold")).pack(side="left")
        ttk.Label(header, text="多份文档使用标签页切换", foreground="#68727d").pack(side="left", padx=(10, 0))
        ttk.Button(header, text="新建窗口", command=self.new_window).pack(side="right")
        ttk.Button(header, text="关闭当前标签", command=self.close_current_tab).pack(side="right", padx=(0, 7))

        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self.activate_current)

    @classmethod
    def shared(cls, parent) -> "PDFWorkspace":
        root = parent.winfo_toplevel()
        key = id(root)
        workspace = _SHARED_WORKSPACES.get(key)
        if workspace is None or not workspace.window.winfo_exists():
            workspace = cls(root, key)
            _SHARED_WORKSPACES[key] = workspace
        else:
            workspace.window.deiconify()
            workspace.window.lift()
        return workspace

    def open_document(self, database, document_row, initial_page: int = 1) -> "PDFDocumentView":
        key = (str(database.db_path), int(document_row["id"]))
        existing_tab = self.document_tabs.get(key)
        if existing_tab and existing_tab in self.views:
            view = self.views[existing_tab]
            self.notebook.select(existing_tab)
            view.jump_to_page(max(0, initial_page - 1))
            self.window.lift()
            return view

        frame = ttk.Frame(self.notebook)
        title = str(document_row["name"])
        display_title = title if len(title) <= 24 else f"{title[:22]}…"
        self.notebook.add(frame, text=display_title)
        tab_id = str(frame)
        view = PDFDocumentView(self, database, document_row, frame, initial_page)
        self.views[tab_id] = view
        self.document_tabs[key] = tab_id
        self.notebook.select(frame)
        self.window.title(f"PDF阅读与批注 · {title}")
        self.window.lift()
        return view

    def current_view(self) -> "PDFDocumentView | None":
        selected = self.notebook.select()
        return self.views.get(selected)

    def activate_current(self, _event=None) -> None:
        view = self.current_view()
        if not view:
            return
        self.window.title(f"PDF阅读与批注 · {view.row['name']}")
        self.window.bind("<Control-c>", view.copy_selection_event)
        self.window.bind("<Control-C>", view.copy_selection_event)
        self.window.bind("<Control-f>", view.focus_search)
        self.window.bind("<Control-F>", view.focus_search)

    def new_window(self) -> None:
        view = self.current_view()
        if not view:
            return
        workspace = PDFWorkspace(self.parent)
        workspace.open_document(view.database, view.row, view.current_page() + 1)

    def close_current_tab(self) -> None:
        view = self.current_view()
        if view:
            self.close_view(view)

    def close_view(self, view: "PDFDocumentView") -> None:
        tab_id = next((item for item, candidate in self.views.items() if candidate is view), "")
        if not tab_id:
            return
        view.dispose()
        self.views.pop(tab_id, None)
        for key, value in list(self.document_tabs.items()):
            if value == tab_id:
                self.document_tabs.pop(key, None)
        try:
            self.notebook.forget(tab_id)
        except tk.TclError:
            pass
        if not self.views:
            self.close()

    def close(self) -> None:
        for view in list(self.views.values()):
            view.dispose()
        self.views.clear()
        self.document_tabs.clear()
        if self.registered_key is not None:
            _SHARED_WORKSPACES.pop(self.registered_key, None)
        if self.window.winfo_exists():
            self.window.destroy()


class PDFDocumentView:
    """A centered, continuous PDF reader with Zotero-style text annotations."""

    def __init__(self, workspace: PDFWorkspace, database, document_row, container, initial_page: int = 1) -> None:
        self.workspace = workspace
        self.parent = workspace.window
        self.container = container
        self.database = database
        self.row = document_row
        self.doc_id = int(document_row["id"])
        self.path = Path(document_row["stored_path"])
        self.pdf = fitz.open(str(self.path))
        # The former 300% rendering size is the new logical 100% baseline.
        self.zoom = DEFAULT_RENDER_ZOOM
        self.page_gap = 22
        self.layouts: list[dict] = []
        self.photos: dict[int, ImageTk.PhotoImage] = {}
        self.image_items: dict[int, int] = {}
        self.mode = tk.StringVar(value="select")
        self.color = COLORS[0]
        self.drag_start: tuple[int, float, float] | None = None
        self.preview_item = None
        self.ink_points: list[tuple[float, float]] = []
        self.selected_page: int | None = None
        self.selected_rects: list[list[float]] = []
        self.selected_text = ""
        self.undo_stack: list[tuple[str, list]] = []
        self.annotation_bounds: dict[int, tuple[float, float, float, float]] = {}
        self.selection_popup = None
        self.layout_job = None
        self.character_cache: dict[int, list[tuple[int, int, int, int, str, fitz.Rect]]] = {}
        self.search_hits: dict[str, tuple[int, fitz.Rect]] = {}

        self.window = workspace.window
        self.page_text = tk.StringVar(value="")
        self.zoom_text = tk.StringVar(value="100%")
        self.status_text = tk.StringVar(value="批注自动保存；点击“保存”可确认当前状态")
        self._build_ui()
        self.window.after(80, lambda: self.rebuild_layout(initial_page - 1))

    def _build_ui(self) -> None:
        bar = ttk.Frame(self.container, padding=(10, 7))
        bar.pack(fill="x")
        nav_bar = ttk.Frame(bar)
        nav_bar.pack(fill="x")
        tool_bar = ttk.Frame(bar)
        tool_bar.pack(fill="x", pady=(6, 0))
        ttk.Button(nav_bar, text="◀", width=3, command=lambda: self.jump_relative(-1)).pack(side="left")
        ttk.Button(nav_bar, text="▶", width=3, command=lambda: self.jump_relative(1)).pack(side="left", padx=(3, 8))
        self.page_entry = ttk.Entry(nav_bar, width=6, justify="center")
        self.page_entry.pack(side="left")
        self.page_entry.bind("<Return>", self.go_page)
        ttk.Label(nav_bar, textvariable=self.page_text).pack(side="left", padx=(4, 14))
        ttk.Button(nav_bar, text="－", width=3, command=lambda: self.change_zoom(-0.30)).pack(side="left")
        ttk.Button(nav_bar, text="＋", width=3, command=lambda: self.change_zoom(0.30)).pack(side="left", padx=(3, 3))
        self.zoom_combo = ttk.Combobox(
            nav_bar, textvariable=self.zoom_text, values=("50%", "75%", "100%", "125%", "150%", "200%"),
            width=7, state="readonly",
        )
        self.zoom_combo.pack(side="left", padx=(1, 3))
        self.zoom_combo.bind("<<ComboboxSelected>>", self.set_zoom_percent)
        ttk.Button(nav_bar, text="适合宽度", command=self.fit_width).pack(side="left", padx=(3, 14))

        for value in ("select", "highlight", "underline", "ink", "note", "eraser"):
            ttk.Radiobutton(
                tool_bar, text=TOOL_LABELS[value], value=value, variable=self.mode,
                command=self.mode_changed,
            ).pack(side="left", padx=2)
        self.color_button = tk.Button(tool_bar, text="颜色", bg=self.color, relief="solid", command=self.choose_color, padx=8)
        self.color_button.pack(side="left", padx=7)
        ttk.Button(tool_bar, text="撤销", command=self.undo).pack(side="left", padx=(2, 4))
        ttk.Button(tool_bar, text="保存", command=self.save).pack(side="left")
        ttk.Button(tool_bar, text="复制选中文字", command=self.copy_selection).pack(side="left", padx=(6, 0))

        file_button = ttk.Menubutton(nav_bar, text="另存▼")
        file_menu = tk.Menu(file_button, tearoff=False)
        file_menu.add_command(label="另存原版PDF…", command=self.save_original)
        file_menu.add_command(label="另存批注版PDF…", command=self.save_annotated)
        file_menu.add_separator()
        file_menu.add_command(label="清除全部批注", command=self.clear_all)
        file_button.configure(menu=file_menu)
        file_button.pack(side="right")

        panes = ttk.Panedwindow(self.container, orient="horizontal")
        panes.pack(fill="both", expand=True)
        reader = ttk.Frame(panes)
        side = ttk.Frame(panes, padding=(10, 8))
        panes.add(reader, weight=5)
        panes.add(side, weight=1)
        self.canvas = tk.Canvas(reader, background="#e5e7ea", highlightthickness=0)
        ybar = ttk.Scrollbar(reader, orient="vertical", command=self._scroll_command)
        xbar = ttk.Scrollbar(reader, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        reader.rowconfigure(0, weight=1)
        reader.columnconfigure(0, weight=1)
        self.canvas.bind("<Configure>", self.on_canvas_resize)
        self.canvas.bind("<Button-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<Button-3>", self.annotation_context_menu)
        self.canvas.bind("<MouseWheel>", self.mouse_wheel)

        self.side_notebook = ttk.Notebook(side)
        self.side_notebook.pack(fill="both", expand=True)
        toc_panel = ttk.Frame(self.side_notebook, padding=7)
        search_panel = ttk.Frame(self.side_notebook, padding=7)
        annotation_panel = ttk.Frame(self.side_notebook, padding=7)
        self.side_notebook.add(toc_panel, text="目录")
        self.side_notebook.add(search_panel, text="全文检索")
        self.side_notebook.add(annotation_panel, text="批注")

        self.toc_tree = ttk.Treeview(toc_panel, show="tree", height=18)
        self.toc_tree.pack(fill="both", expand=True)
        self.toc_tree.bind("<Double-1>", self.jump_toc)
        self._populate_outline()

        search_row = ttk.Frame(search_panel)
        search_row.pack(fill="x", pady=(0, 7))
        self.pdf_search_text = tk.StringVar()
        self.pdf_search_entry = ttk.Entry(search_row, textvariable=self.pdf_search_text)
        self.pdf_search_entry.pack(side="left", fill="x", expand=True)
        self.pdf_search_entry.bind("<Return>", self.search_document)
        ttk.Button(search_row, text="查找", command=self.search_document).pack(side="left", padx=(5, 0))
        self.search_result_list = ttk.Treeview(search_panel, columns=("page", "text"), show="headings", height=18)
        self.search_result_list.heading("page", text="页码")
        self.search_result_list.heading("text", text="匹配原文")
        self.search_result_list.column("page", width=48, anchor="center", stretch=False)
        self.search_result_list.column("text", width=190, anchor="w")
        self.search_result_list.pack(fill="both", expand=True)
        self.search_result_list.bind("<Double-1>", self.jump_search_result)

        ttk.Label(annotation_panel, text="双击批注可跳转；右键可编辑或删除", foreground="#68727d", wraplength=230).pack(anchor="w", pady=(0, 7))
        self.annotation_list = ttk.Treeview(annotation_panel, columns=("page", "type"), show="headings", height=18)
        self.annotation_list.heading("page", text="页码")
        self.annotation_list.heading("type", text="类型/备注")
        self.annotation_list.column("page", width=55, anchor="center", stretch=False)
        self.annotation_list.column("type", width=175, anchor="w")
        self.annotation_list.pack(fill="both", expand=True)
        self.annotation_list.bind("<Double-1>", self.jump_annotation)
        self.annotation_list.bind("<Button-3>", self.list_context_menu)

        footer = ttk.Frame(self.container, padding=(10, 5))
        footer.pack(fill="x")
        ttk.Label(footer, textvariable=self.status_text, foreground="#245f8a").pack(side="left")
        ttk.Label(footer, text="逐字拖选·可跨行｜Ctrl+C 复制｜Ctrl+F 全文检索｜Ctrl+滚轮缩放").pack(side="right")

    def on_canvas_resize(self, _event=None) -> None:
        if self.layout_job:
            self.window.after_cancel(self.layout_job)
        self.layout_job = self.window.after(120, lambda: self.rebuild_layout(self.current_page()))

    def rebuild_layout(self, keep_page: int = 0) -> None:
        self.layout_job = None
        self.canvas.delete("all")
        self.photos.clear()
        self.image_items.clear()
        self.layouts.clear()
        viewport = max(640, self.canvas.winfo_width())
        max_width = 0.0
        y = self.page_gap
        for page_no, page in enumerate(self.pdf):
            width = page.rect.width * self.zoom
            height = page.rect.height * self.zoom
            content_width = max(viewport, width + 80)
            x = max(40, (content_width - width) / 2)
            self.layouts.append({"page": page_no, "x": x, "y": y, "w": width, "h": height})
            max_width = max(max_width, content_width)
            self.canvas.create_rectangle(x, y, x + width, y + height, fill="white", outline="#bec3c8", tags=("paper",))
            y += height + self.page_gap
        self.canvas.configure(scrollregion=(0, 0, max_width, y))
        self.redraw_annotations()
        self.refresh_annotation_list()
        self.jump_to_page(max(0, min(len(self.pdf) - 1, keep_page)))
        self.window.after_idle(self.render_visible)

    def render_visible(self) -> None:
        if not self.layouts:
            return
        top = self.canvas.canvasy(0)
        bottom = top + self.canvas.winfo_height()
        wanted = {
            item["page"] for item in self.layouts
            if item["y"] + item["h"] >= top - 500 and item["y"] <= bottom + 500
        }
        for page_no in list(self.photos):
            if page_no not in wanted:
                self.canvas.delete(self.image_items.pop(page_no))
                self.photos.pop(page_no, None)
        for page_no in wanted:
            if page_no in self.photos:
                continue
            layout = self.layouts[page_no]
            pixmap = self.pdf[page_no].get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom), alpha=False)
            image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
            photo = ImageTk.PhotoImage(image)
            item_id = self.canvas.create_image(layout["x"], layout["y"], image=photo, anchor="nw", tags=("pdf-page",))
            self.photos[page_no] = photo
            self.image_items[page_no] = item_id
        for overlay in ("annotation", "search-match", "selection"):
            if self.canvas.find_withtag(overlay):
                self.canvas.tag_raise(overlay)
        self.update_page_indicator()

    def redraw_annotations(self) -> None:
        self.canvas.delete("annotation")
        self.annotation_bounds.clear()
        for row in self.database.list_pdf_annotations(self.doc_id):
            page_no = int(row["page_no"])
            if page_no >= len(self.layouts):
                continue
            layout = self.layouts[page_no]
            try:
                geometry = json.loads(row["geometry"])
            except Exception:
                continue
            annotation_id = int(row["id"])
            color = str(row["color"])
            kind = str(row["kind"])
            tag = ("annotation", f"ann:{annotation_id}")
            if "rect" in geometry:
                px0, py0, px1, py1 = geometry["rect"]
                rect = (layout["x"] + px0 * self.zoom, layout["y"] + py0 * self.zoom,
                        layout["x"] + px1 * self.zoom, layout["y"] + py1 * self.zoom)
                self.annotation_bounds[annotation_id] = rect
                if kind == "highlight":
                    self.canvas.create_rectangle(*rect, fill=color, outline="", stipple="gray50", tags=tag)
                elif kind == "underline":
                    self.canvas.create_line(rect[0], rect[3], rect[2], rect[3], fill=color, width=3, tags=tag)
                elif kind == "note":
                    self.canvas.create_rectangle(*rect, fill="#fff1a6", outline="#bd8c00", width=2, tags=tag)
                    self.canvas.create_text(rect[0] + 5, rect[1] + 4, text=row["text"] or "备注", anchor="nw",
                                            width=max(80, rect[2] - rect[0] - 10), tags=tag)
                else:
                    self.canvas.create_rectangle(*rect, outline=color, width=3, tags=tag)
            elif "points" in geometry:
                points = [(layout["x"] + x * self.zoom, layout["y"] + y * self.zoom) for x, y in geometry["points"]]
                if len(points) > 1:
                    flat = [coordinate for point in points for coordinate in point]
                    self.canvas.create_line(*flat, fill=color, width=3, smooth=True, tags=tag)
                    xs, ys = [p[0] for p in points], [p[1] for p in points]
                    self.annotation_bounds[annotation_id] = (min(xs), min(ys), max(xs), max(ys))
        self.canvas.tag_raise("annotation")

    def refresh_annotation_list(self) -> None:
        self.annotation_list.delete(*self.annotation_list.get_children())
        labels = {"highlight": "高亮", "underline": "下划线", "ink": "手写", "note": "备注", "bold": "加粗", "color": "颜色"}
        for row in self.database.list_pdf_annotations(self.doc_id):
            description = str(row["text"]).strip() or labels.get(str(row["kind"]), str(row["kind"]))
            self.annotation_list.insert("", "end", iid=str(row["id"]), values=(int(row["page_no"]) + 1, description[:28]))

    def _populate_outline(self) -> None:
        self.toc_tree.delete(*self.toc_tree.get_children())
        outline = self.pdf.get_toc(simple=True)
        if not outline:
            self.toc_tree.insert("", "end", text="该 PDF 没有内置目录", values=(0,))
            return
        parents: dict[int, str] = {0: ""}
        for index, item in enumerate(outline):
            level, title, page_no = int(item[0]), str(item[1]), max(1, int(item[2]))
            parent = parents.get(level - 1, "")
            node = self.toc_tree.insert(parent, "end", iid=f"toc:{index}", text=title, values=(page_no,))
            parents[level] = node
            for deeper in [key for key in parents if key > level]:
                parents.pop(deeper, None)

    def jump_toc(self, _event=None) -> None:
        selected = self.toc_tree.selection()
        if not selected:
            return
        values = self.toc_tree.item(selected[0], "values")
        if values and int(values[0]) > 0:
            self.jump_to_page(int(values[0]) - 1)

    def focus_search(self, _event=None):
        self.side_notebook.select(1)
        self.pdf_search_entry.focus_set()
        self.pdf_search_entry.selection_range(0, "end")
        return "break"

    def search_document(self, _event=None) -> None:
        query = self.pdf_search_text.get().strip()
        self.search_result_list.delete(*self.search_result_list.get_children())
        self.search_hits.clear()
        self.canvas.delete("search-match")
        if not query:
            self.status_text.set("请输入 PDF 内的查找文字")
            return
        result_index = 0
        for page_no, page in enumerate(self.pdf):
            matches = page.search_for(query)
            if not matches:
                continue
            page_text = " ".join(page.get_text("text").split())
            position = page_text.lower().find(query.lower())
            if position >= 0:
                start = max(0, position - 28)
                end = min(len(page_text), position + len(query) + 44)
                snippet = page_text[start:end]
            else:
                snippet = query
            for rect in matches:
                item_id = f"match:{result_index}"
                self.search_result_list.insert("", "end", iid=item_id, values=(page_no + 1, snippet))
                self.search_hits[item_id] = (page_no, fitz.Rect(rect))
                result_index += 1
        self.status_text.set(f"找到 {result_index} 处匹配" if result_index else "未在该 PDF 中找到匹配文字")
        if result_index:
            first = self.search_result_list.get_children()[0]
            self.search_result_list.selection_set(first)
            self.jump_search_result()

    def jump_search_result(self, _event=None) -> None:
        selected = self.search_result_list.selection()
        if not selected or selected[0] not in self.search_hits:
            return
        page_no, rect = self.search_hits[selected[0]]
        if not self.layouts:
            self.rebuild_layout(page_no)
        self.jump_to_position(page_no, rect.y0)
        self.canvas.delete("search-match")
        layout = self.layouts[page_no]
        self.canvas.create_rectangle(
            layout["x"] + rect.x0 * self.zoom, layout["y"] + rect.y0 * self.zoom,
            layout["x"] + rect.x1 * self.zoom, layout["y"] + rect.y1 * self.zoom,
            outline="#d9480f", fill="#ffec99", stipple="gray50", width=2, tags=("search-match",),
        )
        self.canvas.tag_raise("search-match")

    def _scroll_command(self, *args) -> None:
        self.canvas.yview(*args)
        self.window.after_idle(self.render_visible)

    def mouse_wheel(self, event) -> str:
        if event.state & 0x0004:
            self.change_zoom(0.30 if event.delta > 0 else -0.30)
        else:
            self.canvas.yview_scroll(-4 if event.delta > 0 else 4, "units")
            self.window.after_idle(self.render_visible)
        return "break"

    def current_page(self) -> int:
        if not self.layouts:
            return 0
        center = self.canvas.canvasy(0) + self.canvas.winfo_height() / 2
        return min(range(len(self.layouts)), key=lambda i: abs(self.layouts[i]["y"] + self.layouts[i]["h"] / 2 - center))

    def update_page_indicator(self) -> None:
        page = self.current_page() + 1
        logical_percent = round(self.zoom / DEFAULT_RENDER_ZOOM * 100)
        self.zoom_text.set(f"{logical_percent}%")
        self.page_text.set(f"/ {len(self.pdf)} 页 · {logical_percent}%")
        self.page_entry.delete(0, "end")
        self.page_entry.insert(0, str(page))

    def jump_to_page(self, page_no: int) -> None:
        if not self.layouts:
            return
        total = float(self.canvas.cget("scrollregion").split()[3])
        self.canvas.yview_moveto(max(0.0, self.layouts[page_no]["y"] - 8) / max(total, 1))
        self.window.after_idle(self.render_visible)

    def jump_to_position(self, page_no: int, page_y: float) -> None:
        if not self.layouts:
            return
        page_no = max(0, min(len(self.layouts) - 1, page_no))
        total = float(self.canvas.cget("scrollregion").split()[3])
        target_y = self.layouts[page_no]["y"] + page_y * self.zoom - 70
        self.canvas.yview_moveto(max(0.0, target_y) / max(total, 1))
        self.window.after_idle(self.render_visible)

    def jump_relative(self, step: int) -> None:
        self.jump_to_page(max(0, min(len(self.pdf) - 1, self.current_page() + step)))

    def go_page(self, _event=None) -> None:
        try:
            target = int(self.page_entry.get()) - 1
        except ValueError:
            return
        self.jump_to_page(max(0, min(len(self.pdf) - 1, target)))

    def change_zoom(self, delta: float) -> None:
        page = self.current_page()
        self.zoom = max(0.75, min(6.0, self.zoom + delta))
        self.rebuild_layout(page)

    def set_zoom_percent(self, _event=None) -> None:
        try:
            percent = int(self.zoom_text.get().rstrip("%"))
        except ValueError:
            return
        page = self.current_page()
        self.zoom = max(0.75, min(6.0, DEFAULT_RENDER_ZOOM * percent / 100))
        self.rebuild_layout(page)

    def fit_width(self) -> None:
        page = self.current_page()
        available = max(500, self.canvas.winfo_width() - 100)
        self.zoom = max(0.75, min(6.0, available / self.pdf[page].rect.width))
        self.rebuild_layout(page)

    def mode_changed(self) -> None:
        self.clear_selection()
        cursor = "xterm" if self.mode.get() in {"select", "highlight", "underline"} else "crosshair"
        self.canvas.configure(cursor=cursor)
        self.status_text.set(f"当前工具：{TOOL_LABELS[self.mode.get()]}")

    def choose_color(self) -> None:
        color = colorchooser.askcolor(self.color, parent=self.window)[1]
        if color:
            self.color = color
            self.color_button.configure(bg=color)

    def page_point(self, event) -> tuple[int, float, float] | None:
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        for layout in self.layouts:
            if layout["x"] <= x <= layout["x"] + layout["w"] and layout["y"] <= y <= layout["y"] + layout["h"]:
                return layout["page"], (x - layout["x"]) / self.zoom, (y - layout["y"]) / self.zoom
        return None

    def on_press(self, event) -> None:
        self.close_selection_popup()
        point = self.page_point(event)
        if not point:
            return
        page_no, x, y = point
        mode = self.mode.get()
        if mode == "eraser":
            self.erase_at(event)
            return
        if mode == "note":
            text = simpledialog.askstring("添加备注", "请输入备注内容：", parent=self.window)
            if text:
                rect = [x, y, x + 210 / self.zoom, y + 90 / self.zoom]
                annotation_id = self.database.add_pdf_annotation(self.doc_id, page_no, "note", json.dumps({"rect": rect}), self.color, text)
                self.undo_stack.append(("add", [annotation_id]))
                self.refresh_after_edit("备注已保存")
            return
        if mode in {"select", "highlight", "underline"}:
            self.clear_selection()
        self.drag_start = (page_no, x, y)
        self.ink_points = [(x, y)] if mode == "ink" else []

    def on_drag(self, event) -> None:
        if not self.drag_start:
            return
        point = self.page_point(event)
        if not point or point[0] != self.drag_start[0]:
            return
        page_no, x, y = point
        layout = self.layouts[page_no]
        if self.mode.get() == "ink":
            px, py = self.ink_points[-1]
            self.ink_points.append((x, y))
            self.canvas.create_line(layout["x"] + px * self.zoom, layout["y"] + py * self.zoom,
                                    layout["x"] + x * self.zoom, layout["y"] + y * self.zoom,
                                    fill=self.color, width=3, tags=("selection",))
            return
        _, x0, y0 = self.drag_start
        if self.mode.get() in {"select", "highlight", "underline"}:
            # Render the actual character range while dragging. This avoids the
            # old rectangular marquee and behaves like a word processor.
            self.select_characters(page_no, (x0, y0), (x, y), announce=False)

    def on_release(self, event) -> None:
        if not self.drag_start:
            return
        point = self.page_point(event)
        page_no, x0, y0 = self.drag_start
        self.drag_start = None
        if not point or point[0] != page_no:
            self.clear_selection()
            return
        _, x1, y1 = point
        mode = self.mode.get()
        if mode == "ink":
            if len(self.ink_points) > 1:
                annotation_id = self.database.add_pdf_annotation(
                    self.doc_id, page_no, "ink", json.dumps({"points": self.ink_points}), self.color
                )
                self.undo_stack.append(("add", [annotation_id]))
                self.refresh_after_edit("手写批注已保存")
            return
        if (x1 - x0) ** 2 + (y1 - y0) ** 2 < 9:
            self.clear_selection()
            return
        self.select_characters(page_no, (x0, y0), (x1, y1))
        if mode in {"highlight", "underline"} and self.selected_rects:
            self.create_from_selection(mode)
        elif mode == "select" and self.selected_rects:
            self.show_selection_popup(event)

    def select_characters(
        self, page_no: int, start: tuple[float, float], end: tuple[float, float], announce: bool = True
    ) -> None:
        """Select a logical character range, including partial Chinese text lines."""
        characters = self._page_characters(page_no)

        self.selected_page = page_no
        self.selected_rects = []
        self.selected_text = ""
        if not characters:
            self.draw_selection()
            self.status_text.set("这个PDF页面没有可选择文字；图片扫描页需要先生成OCR文字层")
            return

        def nearest_index(point: tuple[float, float]) -> int:
            px, py = point
            best_index = 0
            best_distance = float("inf")
            for index, (*_metadata, rect) in enumerate(characters):
                dx = 0.0 if rect.x0 <= px <= rect.x1 else min(abs(px - rect.x0), abs(px - rect.x1))
                dy = 0.0 if rect.y0 <= py <= rect.y1 else min(abs(py - rect.y0), abs(py - rect.y1))
                # Prefer the correct text line before a horizontally closer
                # character on a neighbouring line.
                distance = dx * dx + (dy * 4.0) ** 2
                if distance < best_distance:
                    best_index, best_distance = index, distance
            return best_index

        start_index = nearest_index(start)
        end_index = nearest_index(end)
        first, last = sorted((start_index, end_index))
        selected = characters[first : last + 1]
        lines: list[str] = []
        current_key = None
        current_text: list[str] = []
        current_rect: fitz.Rect | None = None
        for block_no, line_no, _span_no, _char_no, value, rect in selected:
            key = (block_no, line_no)
            if current_key is not None and key != current_key:
                lines.append("".join(current_text))
                assert current_rect is not None
                self.selected_rects.append(
                    [current_rect.x0, current_rect.y0, current_rect.x1, current_rect.y1]
                )
                current_text = []
                current_rect = None
            current_key = key
            current_text.append(value)
            current_rect = fitz.Rect(rect) if current_rect is None else current_rect | rect
        if current_text:
            lines.append("".join(current_text))
            assert current_rect is not None
            self.selected_rects.append(
                [current_rect.x0, current_rect.y0, current_rect.x1, current_rect.y1]
            )
        self.selected_text = "\n".join(lines)
        self.draw_selection()
        if announce:
            self.status_text.set(f"已选择 {len(self.selected_text.replace(chr(10), ''))} 个字符；可按 Ctrl+C 复制")

    def _page_characters(self, page_no: int) -> list[tuple[int, int, int, int, str, fitz.Rect]]:
        """Return cached characters in the PDF's logical reading order."""
        cached = self.character_cache.get(page_no)
        if cached is not None:
            return cached
        characters: list[tuple[int, int, int, int, str, fitz.Rect]] = []
        raw = self.pdf[page_no].get_text("rawdict")
        for block_no, block in enumerate(raw.get("blocks", [])):
            if block.get("type") != 0:
                continue
            for line_no, line in enumerate(block.get("lines", [])):
                for span_no, span in enumerate(line.get("spans", [])):
                    for char_no, char in enumerate(span.get("chars", [])):
                        value = str(char.get("c", ""))
                        bbox = char.get("bbox")
                        if value and bbox:
                            characters.append((block_no, line_no, span_no, char_no, value, fitz.Rect(bbox)))
        self.character_cache[page_no] = characters
        return characters

    def select_words(self, page_no: int, selection: fitz.Rect) -> None:
        """Backward-compatible wrapper used by older integrations."""
        self.select_characters(page_no, (selection.x0, selection.y0), (selection.x1, selection.y1))

    @staticmethod
    def _group_words(words):
        import itertools
        return itertools.groupby(words, key=lambda word: (word[5], word[6]))

    def draw_selection(self) -> None:
        self.canvas.delete("selection")
        if self.selected_page is None or self.selected_page >= len(self.layouts):
            return
        layout = self.layouts[self.selected_page]
        for x0, y0, x1, y1 in self.selected_rects:
            self.canvas.create_rectangle(
                layout["x"] + x0 * self.zoom, layout["y"] + y0 * self.zoom,
                layout["x"] + x1 * self.zoom, layout["y"] + y1 * self.zoom,
                fill="#58a5ff", outline="", stipple="gray50", tags=("selection",),
            )

    def show_selection_popup(self, event) -> None:
        self.close_selection_popup()
        popup = tk.Toplevel(self.window)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.geometry(f"+{event.x_root + 8}+{event.y_root + 8}")
        panel = ttk.Frame(popup, padding=6, relief="solid")
        panel.pack()
        colors = ttk.Frame(panel)
        colors.pack(fill="x")
        for color in COLORS:
            tk.Button(colors, bg=color, width=2, command=lambda value=color: self.pick_popup_color(value)).pack(side="left", padx=2)
        actions = ttk.Frame(panel)
        actions.pack(fill="x", pady=(5, 0))
        ttk.Button(actions, text="高亮", command=lambda: self.create_from_selection("highlight")).pack(side="left")
        ttk.Button(actions, text="下划线", command=lambda: self.create_from_selection("underline")).pack(side="left", padx=4)
        ttk.Button(actions, text="备注", command=self.add_note_from_selection).pack(side="left")
        ttk.Button(actions, text="复制", command=self.copy_selection).pack(side="left")
        self.selection_popup = popup

    def pick_popup_color(self, color: str) -> None:
        self.color = color
        self.color_button.configure(bg=color)

    def create_from_selection(self, kind: str) -> None:
        if self.selected_page is None or not self.selected_rects:
            return
        ids = [
            self.database.add_pdf_annotation(
                self.doc_id, self.selected_page, kind, json.dumps({"rect": rect}), self.color,
                self.selected_text,
            )
            for rect in self.selected_rects
        ]
        self.undo_stack.append(("add", ids))
        self.clear_selection()
        self.refresh_after_edit(f"已保存{TOOL_LABELS[kind]}批注")

    def add_note_from_selection(self) -> None:
        if self.selected_page is None or not self.selected_rects:
            return
        note = simpledialog.askstring("添加文字备注", "请输入对选中原文的备注：", parent=self.window)
        if not note:
            return
        first = self.selected_rects[0]
        page = self.pdf[self.selected_page]
        width = min(210 / self.zoom, max(90 / self.zoom, page.rect.width - first[0] - 8))
        height = 90 / self.zoom
        x0 = min(first[0], max(0.0, page.rect.width - width))
        y0 = max(0.0, first[1] - height - 4 / self.zoom)
        rect = [x0, y0, x0 + width, min(page.rect.height, y0 + height)]
        annotation_id = self.database.add_pdf_annotation(
            self.doc_id, self.selected_page, "note", json.dumps({"rect": rect}), self.color, note,
        )
        self.undo_stack.append(("add", [annotation_id]))
        self.clear_selection()
        self.refresh_after_edit("文字备注已保存")

    def copy_selection(self) -> None:
        if self.selected_text:
            self.window.clipboard_clear()
            self.window.clipboard_append(self.selected_text)
            self.status_text.set("选中文字已复制")
        else:
            self.status_text.set("请先使用“选择文字”工具拖动选择几个字，再复制")
        self.close_selection_popup()

    def copy_selection_event(self, _event=None):
        self.copy_selection()
        return "break"

    def clear_selection(self) -> None:
        self.canvas.delete("selection")
        self.selected_page = None
        self.selected_rects = []
        self.selected_text = ""
        self.preview_item = None
        self.ink_points = []
        self.close_selection_popup()

    def close_selection_popup(self) -> None:
        if self.selection_popup and self.selection_popup.winfo_exists():
            self.selection_popup.destroy()
        self.selection_popup = None

    def erase_at(self, event) -> None:
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        annotation_id = self.annotation_at(x, y)
        if annotation_id is not None:
            rows = [dict(row) for row in self.database.list_pdf_annotations(self.doc_id) if int(row["id"]) == annotation_id]
            self.database.delete_pdf_annotation(annotation_id)
            self.undo_stack.append(("delete", rows))
            self.refresh_after_edit("批注已擦除；可点击撤销恢复")

    def annotation_at(self, x: float, y: float) -> int | None:
        for annotation_id, (x0, y0, x1, y1) in reversed(list(self.annotation_bounds.items())):
            if x0 - 7 <= x <= x1 + 7 and y0 - 7 <= y <= y1 + 7:
                return annotation_id
        return None

    def undo(self) -> None:
        if not self.undo_stack:
            self.status_text.set("没有可撤销的批注操作")
            return
        action, items = self.undo_stack.pop()
        if action == "add":
            for annotation_id in items:
                self.database.delete_pdf_annotation(annotation_id)
        elif action == "delete":
            for row in items:
                self.database.add_pdf_annotation(
                    int(row["doc_id"]), int(row["page_no"]), str(row["kind"]),
                    str(row["geometry"]), str(row["color"]), str(row["text"]),
                )
        self.refresh_after_edit("已撤销上一步批注操作")

    def save(self) -> None:
        self.database.checkpoint()
        self.status_text.set("全部批注已保存")

    def refresh_after_edit(self, message: str) -> None:
        self.redraw_annotations()
        self.refresh_annotation_list()
        self.status_text.set(message)

    def jump_annotation(self, _event=None) -> None:
        selected = self.annotation_list.selection()
        if selected:
            values = self.annotation_list.item(selected[0], "values")
            self.jump_to_page(int(values[0]) - 1)

    def list_context_menu(self, event) -> None:
        row = self.annotation_list.identify_row(event.y)
        if row:
            self.annotation_list.selection_set(row)
            self.show_annotation_menu(int(row), event.x_root, event.y_root)

    def annotation_context_menu(self, event) -> None:
        if self.selected_text and self.selected_rects:
            self.close_selection_popup()
            menu = tk.Menu(self.window, tearoff=False)
            menu.add_command(label="复制选中文字", command=self.copy_selection)
            menu.add_separator()
            menu.add_command(label="添加荧光高亮", command=lambda: self.create_from_selection("highlight"))
            menu.add_command(label="添加下划线", command=lambda: self.create_from_selection("underline"))
            menu.add_command(label="添加文字备注…", command=self.add_note_from_selection)
            menu.tk_popup(event.x_root, event.y_root)
            return
        annotation_id = self.annotation_at(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        if annotation_id is not None:
            self.show_annotation_menu(annotation_id, event.x_root, event.y_root)

    def show_annotation_menu(self, annotation_id: int, x: int, y: int) -> None:
        menu = tk.Menu(self.window, tearoff=False)
        menu.add_command(label="修改颜色", command=lambda: self.edit_annotation(annotation_id, False))
        menu.add_command(label="编辑备注", command=lambda: self.edit_annotation(annotation_id, True))
        menu.add_separator()
        menu.add_command(label="删除批注", command=lambda: self.delete_annotation(annotation_id))
        menu.tk_popup(x, y)

    def edit_annotation(self, annotation_id: int, edit_text: bool) -> None:
        rows = [row for row in self.database.list_pdf_annotations(self.doc_id) if int(row["id"]) == annotation_id]
        if not rows:
            return
        row = rows[0]
        color, text = str(row["color"]), str(row["text"])
        if edit_text:
            value = simpledialog.askstring("编辑备注", "备注内容：", initialvalue=text, parent=self.window)
            if value is None:
                return
            text = value
        else:
            value = colorchooser.askcolor(color, parent=self.window)[1]
            if not value:
                return
            color = value
        self.database.update_pdf_annotation(annotation_id, color, text)
        self.refresh_after_edit("批注已修改")

    def delete_annotation(self, annotation_id: int) -> None:
        rows = [dict(row) for row in self.database.list_pdf_annotations(self.doc_id) if int(row["id"]) == annotation_id]
        self.database.delete_pdf_annotation(annotation_id)
        self.undo_stack.append(("delete", rows))
        self.refresh_after_edit("批注已删除；可撤销")

    def clear_all(self) -> None:
        rows = [dict(row) for row in self.database.list_pdf_annotations(self.doc_id)]
        if rows and messagebox.askyesno("清除全部批注", "确定清除这份PDF的全部批注吗？", parent=self.window):
            self.database.clear_pdf_annotations(self.doc_id)
            self.undo_stack.append(("delete", rows))
            self.refresh_after_edit("全部批注已清除；可撤销")

    def save_original(self) -> None:
        target = filedialog.asksaveasfilename(
            title="另存原版PDF", defaultextension=".pdf", initialfile=self.path.name,
            filetypes=[("PDF", "*.pdf")], parent=self.window,
        )
        if target:
            shutil.copy2(self.path, target)
            messagebox.showinfo("保存完成", target, parent=self.window)

    def save_annotated(self) -> None:
        target = filedialog.asksaveasfilename(
            title="另存批注版PDF", defaultextension=".pdf", initialfile=f"{self.path.stem}_批注版.pdf",
            filetypes=[("PDF", "*.pdf")], parent=self.window,
        )
        if target:
            try:
                export_annotated_pdf(self.database, self.row, target)
                messagebox.showinfo("保存完成", target, parent=self.window)
            except Exception as exc:
                messagebox.showerror("导出失败", str(exc), parent=self.window)

    def dispose(self) -> None:
        self.close_selection_popup()
        if not self.pdf.is_closed:
            self.pdf.close()

    def close(self) -> None:
        self.workspace.close_view(self)


class ZoteroPDFViewer:
    """Compatibility factory that opens PDFs in the shared tabbed workspace."""

    def __new__(cls, parent, database, document_row, initial_page: int = 1):
        workspace = PDFWorkspace.shared(parent)
        return workspace.open_document(database, document_row, initial_page)
