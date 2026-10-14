"""Real Tk/worker/search/export regression; only dialogs and OS file opening are mocked."""
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import time
import tkinter as tk
import os

import main as app_module


def check_layout(ui):
    window = ui['window']
    for key in ('start', 'choose', 'progress', 'answer', 'log'):
        widget = ui[key]
        assert widget.winfo_ismapped(), key
        assert widget.winfo_rooty() + widget.winfo_height() <= window.winfo_rooty() + window.winfo_height(), key
        if key in ('start', 'choose', 'answer'):
            assert widget.winfo_width() >= widget.winfo_reqwidth(), key


def run(scale=2.0):
    with TemporaryDirectory() as temp:
        folder = Path(temp)
        knowledge = folder / '规章.md'
        knowledge.write_text('转场调车时，为保证调车车列的制动力，遇到情况能及时停车，带车数在11辆及以上必须按《车站行车工作细则》规定连结一定数量的软管。', encoding='utf-8')
        exams = [folder / '试卷一.md', folder / '试卷二.md']
        exams[0].write_text('1. 转场调车时，为保证调车车列的制动力，带车数达到多少必须连结软管？\n\n2. 月球背面的天气情况如何？', encoding='utf-8')
        exams[1].write_text('1. 转场调车时，为保证调车车列的制动力，遇到情况能及时停车，应执行什么规定？', encoding='utf-8')
        with patch.object(app_module, 'DB_PATH', folder / 'test.db'), patch.object(app_module, 'STORAGE_DIR', folder / 'storage'), patch.object(app_module.messagebox, 'showinfo'), patch.object(app_module.messagebox, 'showerror'), patch.object(app_module.filedialog, 'askopenfilenames', return_value=tuple(map(str, exams))), patch.object(app_module.LearningLibraryApp, '_open_existing') as opener:
            root = tk.Tk()
            root.tk.call('tk', 'scaling', scale)
            app = app_module.LearningLibraryApp(root)
            app.database.import_file(knowledge)
            app.open_exam_module()
            ui = app.exam_ui
            ui['window'].geometry('920x620')
            root.update()
            check_layout(ui)
            from tkinter import ttk
            assert int(ttk.Style().lookup('Treeview', 'rowheight')) >= 40
            ui['choose'].invoke()
            ui['output_dir'].set(str(folder / '用户选择的输出目录'))
            original_search = app.database.search
            def slow_search(*args, **kwargs):
                time.sleep(0.1)  # Ensure the event loop can display a live search stage.
                return original_search(*args, **kwargs)
            with patch.object(app.database, 'search', side_effect=slow_search):
                ui['start'].invoke()
                assert ui['running']
                live_log_seen = False
                deadline = time.monotonic() + 30
                while ui['running'] and time.monotonic() < deadline:
                    root.update()
                    live_log_seen |= ui['running'] and '正在搜索关键词' in ui['log'].get('1.0', 'end')
                    time.sleep(0.01)
                assert not ui['running'], 'Worker did not finish'
                assert live_log_seen, 'Search logs appeared only after completion'
            root.update()
            assert len(ui['results']) == 2
            assert len(ui['tree'].get_children()) == 3
            assert sum(len(result.matches) for result in ui['results']) == 2
            outputs = ui['last_outputs']
            assert len(outputs) == 7 and all(path.exists() for path in outputs)
            assert all(path.parent.parent == folder / '用户选择的输出目录' for path in outputs)
            assert '已生成 7 份文档' in ui['generated'].get()
            report = outputs[-1].read_text(encoding='utf-8')
            assert '答案出处排行榜' in report and '被 2 道题引用' in report
            with patch.object(app_module.filedialog, 'askopenfilename', return_value=str(outputs[4])):
                ui['answer'].invoke()
            opener.assert_called_with(str(outputs[4]), False)
            for result in ui['results']:
                for question in result.questions:
                    match = result.matches.get(question.id)
                    assert ui['tree'].set(str(question.id), 'answer') == (match.original_text if match else '【无】')
            with patch.object(app_module.filedialog, 'askdirectory', return_value=str(folder / '手动逐卷另存')):
                ui['exports'][1].invoke()
            assert len(list((folder / '手动逐卷另存').rglob('*.docx'))) == 2
            check_layout(ui)
            print('PASS: live search, 2 separate papers / 7 files, verbatim answer column, open chosen paper, manual split Word')
            if os.environ.get('EXAM_QA_CAPTURE'):
                import ctypes
                import win32gui
                import win32ui
                hwnd = win32gui.GetParent(ui['window'].winfo_id())
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                hdc = win32gui.GetWindowDC(hwnd)
                dc = win32ui.CreateDCFromHandle(hdc)
                memory = dc.CreateCompatibleDC()
                bitmap = win32ui.CreateBitmap()
                bitmap.CreateCompatibleBitmap(dc, right - left, bottom - top)
                memory.SelectObject(bitmap)
                assert ctypes.windll.user32.PrintWindow(hwnd, memory.GetSafeHdc(), 2)
                bitmap.SaveBitmapFile(memory, os.environ['EXAM_QA_CAPTURE'])
                win32gui.DeleteObject(bitmap.GetHandle())
                memory.DeleteDC()
                dc.DeleteDC()
                win32gui.ReleaseDC(hwnd, hdc)

            # Missing input must show a failure and restore controls, never pretend to generate.
            exams[0].unlink()
            exams[1].unlink()
            ui['start'].invoke()
            deadline = time.monotonic() + 10
            while ui['running'] and time.monotonic() < deadline:
                root.update()
                time.sleep(0.01)
            assert not ui['running'] and not ui['last_outputs']
            assert '本次未生成' in ui['generated'].get()
            assert str(ui['start']['state']) == 'normal'
            app.manage_documents_v2()
            root.update()
            assert int(ttk.Style().lookup('LibraryManager.Treeview', 'rowheight')) >= 68
            def descendants(widget):
                for child in widget.winfo_children():
                    yield child
                    yield from descendants(child)
            manager_tables = [widget for widget in descendants(root)
                              if isinstance(widget, ttk.Treeview) and str(widget['style']) == 'LibraryManager.Treeview']
            assert manager_tables
            first_row = manager_tables[-1].get_children()[0]
            assert manager_tables[-1].bbox(first_row)[3] >= 68
            for timer in root.tk.call('after', 'info'):
                root.after_cancel(timer)
            root.destroy()
            print('PASS: parse failure restores start button and reports no output')


if __name__ == '__main__':
    run()
