from __future__ import annotations

import json
import shutil
import tkinter as tk
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

import pymupdf as fitz
from PIL import Image, ImageTk


TOOL_NAMES = {
    "view": "浏览",
    "highlight": "高亮",
    "underline": "下划线",
    "bold": "加粗标记",
    "color": "文字颜色标记",
    "ink": "手写",
    "note": "备注框",
}


def center_window(window, width: int, height: int) -> None:
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    width = min(width, max(760, screen_width - 60))
    height = min(height, max(560, screen_height - 80))
    window.geometry(
        f"{width}x{height}+{max(0, (screen_width - width) // 2)}+{max(0, (screen_height - height) // 2)}"
    )
    window.resizable(True, True)


def _rgb(hex_color: str) -> tuple[float, float, float]:
    value = hex_color.lstrip("#")
    if len(value) != 6:
        return 1.0, 0.83, 0.25
    return tuple(int(value[index:index + 2], 16) / 255 for index in (0, 2, 4))


def export_annotated_pdf(database, document_row, target: str | Path) -> Path:
    """Create a new annotated PDF and leave the stored original untouched."""
    source = Path(document_row["stored_path"])
    output = fitz.open(str(source))
    try:
        for row in database.list_pdf_annotations(int(document_row["id"])):
            page = output[int(row["page_no"])]
            geometry = json.loads(row["geometry"])
            kind = str(row["kind"])
            color = _rgb(str(row["color"]))
            annotation = None
            if "rect" in geometry:
                rect = fitz.Rect(*geometry["rect"])
                if kind == "highlight":
                    annotation = page.add_highlight_annot(rect)
                elif kind == "underline":
                    annotation = page.add_underline_annot(rect)
                elif kind == "note":
                    annotation = page.add_text_annot(rect.top_left, str(row["text"]) or "备注")
                else:
                    annotation = page.add_rect_annot(rect)
                    annotation.set_border(width=3 if kind == "bold" else 2)
            elif "points" in geometry:
                points = [fitz.Point(x, y) for x, y in geometry["points"]]
                if len(points) > 1:
                    annotation = page.add_ink_annot([points])
            if annotation:
                annotation.set_colors(stroke=color)
                if kind == "highlight":
                    annotation.set_opacity(0.35)
                annotation.update()
        target_path = Path(target)
        output.save(str(target_path), garbage=4, deflate=True)
        return target_path
    finally:
        output.close()


class PDFViewer:
    """Internal PDF reader whose annotations are stored separately in SQLite."""

    def __init__(self, parent, database, document_row, initial_page: int = 1) -> None:
        self.parent = parent
        self.database = database
        self.row = document_row
        self.doc_id = int(document_row["id"])
        self.path = Path(document_row["stored_path"])
        self.pdf = fitz.open(str(self.path))
        self.page_index = max(0, min(len(self.pdf) - 1, initial_page - 1))
        self.zoom = 1.25
        self.mode = tk.StringVar(value="view")
        self.color = "#ffd54f"
        self.photo = None
        self.drag_start: tuple[float, float] | None = None
        self.preview_item = None
        self.ink_points: list[list[float]] = []
        self.annotation_shapes: dict[int, tuple[float, float, float, float]] = {}

        self.window = tk.Toplevel(parent)
        self.window.title(f"PDF阅读与批注 · {document_row['name']}")
        center_window(self.window, 1280, 860)
        self.window.minsize(900, 620)
        self.window.protocol("WM_DELETE_WINDOW", self.close)

        self.page_text = tk.StringVar()
        self.status_text = tk.StringVar(value="原始PDF保持不变，批注已独立保存")
        self._build_ui()
        self.render()

    def _build_ui(self) -> None:
        toolbar = ttk.Frame(self.window, padding=(10, 8))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="上一页", command=lambda: self.change_page(-1)).pack(side="left")
        ttk.Button(toolbar, text="下一页", command=lambda: self.change_page(1)).pack(side="left", padx=5)
        self.page_entry = ttk.Entry(toolbar, width=7, justify="center")
        self.page_entry.pack(side="left", padx=(8, 3))
        self.page_entry.bind("<Return>", self.go_page)
        ttk.Label(toolbar, textvariable=self.page_text).pack(side="left", padx=(0, 12))
        ttk.Button(toolbar, text="－", width=3, command=lambda: self.change_zoom(-0.15)).pack(side="left")
        ttk.Button(toolbar, text="＋", width=3, command=lambda: self.change_zoom(0.15)).pack(side="left", padx=(3, 10))

        for value in ("view", "highlight", "underline", "bold", "color", "ink", "note"):
            ttk.Radiobutton(
                toolbar, text=TOOL_NAMES[value], value=value, variable=self.mode,
                command=self._mode_changed,
            ).pack(side="left", padx=2)
        self.color_button = tk.Button(
            toolbar, text="颜色", bg=self.color, relief="solid", borderwidth=1,
            command=self.choose_color, padx=8,
        )
        self.color_button.pack(side="left", padx=8)

        file_menu = ttk.Menubutton(toolbar, text="保存/另存")
        menu = tk.Menu(file_menu, tearoff=False)
        menu.add_command(label="另存原版PDF…", command=self.save_original)
        menu.add_command(label="另存批注版PDF…", command=self.save_annotated)
        menu.add_separator()
        menu.add_command(label="清除全部批注", command=self.clear_annotations)
        file_menu.configure(menu=menu)
        file_menu.pack(side="right")

        body = ttk.Frame(self.window)
        body.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(body, background="#5a5f66", highlightthickness=0, cursor="arrow")
        xbar = ttk.Scrollbar(body, orient="horizontal", command=self.canvas.xview)
        ybar = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self.canvas.bind("<Button-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<Button-3>", self.annotation_menu)
        self.canvas.bind("<MouseWheel>", self.mouse_wheel)

        status = ttk.Frame(self.window, padding=(10, 5))
        status.pack(fill="x")
        ttk.Label(status, textvariable=self.status_text, foreground="#235f8c").pack(side="left")
        ttk.Label(status, text="右键批注可修改颜色、编辑备注或删除").pack(side="right")

    def _mode_changed(self) -> None:
        self.canvas.configure(cursor="arrow" if self.mode.get() == "view" else "crosshair")
        self.status_text.set(f"当前工具：{TOOL_NAMES[self.mode.get()]} · 批注自动保存")

    def render(self) -> None:
        page = self.pdf[self.page_index]
        pixmap = page.get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom), alpha=False)
        image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        self.photo = ImageTk.PhotoImage(image)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.photo, anchor="nw", tags=("page",))
        self.canvas.configure(scrollregion=(0, 0, pixmap.width, pixmap.height))
        self.page_text.set(f"/ {len(self.pdf)} 页 · {round(self.zoom * 100)}%")
        self.page_entry.delete(0, "end")
        self.page_entry.insert(0, str(self.page_index + 1))
        self.draw_annotations()

    def draw_annotations(self) -> None:
        self.annotation_shapes.clear()
        for row in self.database.list_pdf_annotations(self.doc_id, self.page_index):
            try:
                geometry = json.loads(row["geometry"])
            except Exception:
                continue
            annotation_id = int(row["id"])
            kind = str(row["kind"])
            color = str(row["color"])
            if "rect" in geometry:
                rect = [value * self.zoom for value in geometry["rect"]]
                x0, y0, x1, y1 = rect
                self.annotation_shapes[annotation_id] = (x0, y0, x1, y1)
                if kind == "highlight":
                    self.canvas.create_rectangle(*rect, fill=color, outline=color, stipple="gray50", tags=(f"ann:{annotation_id}",))
                elif kind == "underline":
                    self.canvas.create_line(x0, y1, x1, y1, fill=color, width=3, tags=(f"ann:{annotation_id}",))
                elif kind == "bold":
                    self.canvas.create_rectangle(*rect, outline=color, width=4, tags=(f"ann:{annotation_id}",))
                elif kind == "color":
                    self.canvas.create_rectangle(*rect, outline=color, width=3, dash=(5, 2), tags=(f"ann:{annotation_id}",))
                elif kind == "note":
                    self.canvas.create_rectangle(*rect, fill="#fff3a6", outline="#c39300", width=2, tags=(f"ann:{annotation_id}",))
                    self.canvas.create_text(x0 + 5, y0 + 5, text=row["text"] or "备注", anchor="nw", width=max(80, x1 - x0 - 10), tags=(f"ann:{annotation_id}",))
            elif "points" in geometry:
                points = [[x * self.zoom, y * self.zoom] for x, y in geometry["points"]]
                flat = [coordinate for point in points for coordinate in point]
                if len(flat) >= 4:
                    self.canvas.create_line(*flat, fill=color, width=3, smooth=True, tags=(f"ann:{annotation_id}",))
                    xs = [point[0] for point in points]
                    ys = [point[1] for point in points]
                    self.annotation_shapes[annotation_id] = (min(xs), min(ys), max(xs), max(ys))

    def change_page(self, step: int) -> None:
        new_page = max(0, min(len(self.pdf) - 1, self.page_index + step))
        if new_page != self.page_index:
            self.page_index = new_page
            self.render()
            self.canvas.yview_moveto(0)

    def go_page(self, _event=None) -> None:
        try:
            page = int(self.page_entry.get())
        except ValueError:
            return
        self.page_index = max(0, min(len(self.pdf) - 1, page - 1))
        self.render()

    def change_zoom(self, amount: float) -> None:
        self.zoom = max(0.5, min(3.0, self.zoom + amount))
        self.render()

    def choose_color(self) -> None:
        chosen = colorchooser.askcolor(self.color, parent=self.window)[1]
        if chosen:
            self.color = chosen
            self.color_button.configure(bg=chosen)

    def _point(self, event) -> tuple[float, float]:
        return self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

    def on_press(self, event) -> None:
        mode = self.mode.get()
        if mode == "view":
            return
        x, y = self._point(event)
        if mode == "note":
            text = simpledialog.askstring("文本备注", "请输入备注内容：", parent=self.window)
            if text:
                rect = [x / self.zoom, y / self.zoom, (x + 220) / self.zoom, (y + 90) / self.zoom]
                self.database.add_pdf_annotation(
                    self.doc_id, self.page_index, "note", json.dumps({"rect": rect}), self.color, text,
                )
                self.draw_annotations()
            return
        self.drag_start = (x, y)
        if mode == "ink":
            self.ink_points = [[x, y]]

    def on_drag(self, event) -> None:
        if not self.drag_start:
            return
        x, y = self._point(event)
        if self.mode.get() == "ink":
            previous = self.ink_points[-1]
            self.ink_points.append([x, y])
            self.canvas.create_line(*previous, x, y, fill=self.color, width=3, tags=("preview",))
            return
        x0, y0 = self.drag_start
        if self.preview_item:
            self.canvas.delete(self.preview_item)
        self.preview_item = self.canvas.create_rectangle(x0, y0, x, y, outline=self.color, width=2, dash=(4, 2))

    def on_release(self, event) -> None:
        if not self.drag_start:
            return
        mode = self.mode.get()
        x, y = self._point(event)
        x0, y0 = self.drag_start
        if mode == "ink":
            if len(self.ink_points) > 1:
                points = [[px / self.zoom, py / self.zoom] for px, py in self.ink_points]
                self.database.add_pdf_annotation(
                    self.doc_id, self.page_index, "ink", json.dumps({"points": points}), self.color,
                )
        elif abs(x - x0) >= 5 and abs(y - y0) >= 3:
            rect = [min(x0, x) / self.zoom, min(y0, y) / self.zoom,
                    max(x0, x) / self.zoom, max(y0, y) / self.zoom]
            self.database.add_pdf_annotation(
                self.doc_id, self.page_index, mode, json.dumps({"rect": rect}), self.color,
            )
        self.drag_start = None
        self.ink_points = []
        if self.preview_item:
            self.canvas.delete(self.preview_item)
            self.preview_item = None
        self.render()

    def annotation_menu(self, event) -> None:
        x, y = self._point(event)
        selected = None
        for annotation_id, (x0, y0, x1, y1) in reversed(list(self.annotation_shapes.items())):
            if x0 - 8 <= x <= x1 + 8 and y0 - 8 <= y <= y1 + 8:
                selected = annotation_id
                break
        if selected is None:
            return
        menu = tk.Menu(self.window, tearoff=False)
        menu.add_command(label="修改颜色", command=lambda: self.edit_annotation(selected, False))
        menu.add_command(label="编辑备注文字", command=lambda: self.edit_annotation(selected, True))
        menu.add_separator()
        menu.add_command(label="删除此批注", command=lambda: self.delete_annotation(selected))
        menu.tk_popup(event.x_root, event.y_root)

    def edit_annotation(self, annotation_id: int, edit_text: bool) -> None:
        rows = [row for row in self.database.list_pdf_annotations(self.doc_id, self.page_index) if int(row["id"]) == annotation_id]
        if not rows:
            return
        row = rows[0]
        color = str(row["color"])
        text = str(row["text"])
        if edit_text:
            updated = simpledialog.askstring("编辑批注", "批注文字：", initialvalue=text, parent=self.window)
            if updated is None:
                return
            text = updated
        else:
            updated_color = colorchooser.askcolor(color, parent=self.window)[1]
            if not updated_color:
                return
            color = updated_color
        self.database.update_pdf_annotation(annotation_id, color, text)
        self.render()

    def delete_annotation(self, annotation_id: int) -> None:
        self.database.delete_pdf_annotation(annotation_id)
        self.render()

    def clear_annotations(self) -> None:
        if messagebox.askyesno("清除全部批注", "确定清除这份PDF的全部批注吗？原版PDF不会被删除。", parent=self.window):
            self.database.clear_pdf_annotations(self.doc_id)
            self.render()

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
            title="另存批注版PDF", defaultextension=".pdf",
            initialfile=f"{self.path.stem}_批注版.pdf", filetypes=[("PDF", "*.pdf")], parent=self.window,
        )
        if not target:
            return
        try:
            export_annotated_pdf(self.database, self.row, target)
            messagebox.showinfo("保存完成", target, parent=self.window)
        except Exception as exc:
            messagebox.showerror("导出批注版失败", str(exc), parent=self.window)

    def mouse_wheel(self, event) -> str:
        if event.state & 0x0004:
            self.change_zoom(0.15 if event.delta > 0 else -0.15)
        else:
            self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
        return "break"

    def close(self) -> None:
        self.pdf.close()
        self.window.destroy()
