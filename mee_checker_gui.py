# -*- coding: utf-8 -*-
"""MEE 多国语数字校对工具 - 图形界面版 (v1.3)

用法: 双击运行(或 python mee_checker_gui.py)
需要: 英文红字指示稿(必填) + 多国语PDF文件夹(必填) + 客户锚定原稿(可选)
核心逻辑复用 mee_checker.run_job
"""
import os
import re
import sys
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import mee_checker

APP_TITLE = 'MEE 多国语数字校对工具 v1.3'


class GUIApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(APP_TITLE)
        root.geometry('720x560')
        root.minsize(640, 480)

        self.q = queue.Queue()
        self.busy = False
        self.out_paths = None   # 完成后 (report, report_html, snaps_dir, out_dir)

        pad = dict(padx=10, pady=6)
        frm = ttk.Frame(root, padding=10)
        frm.pack(fill='x')

        # 英文指示稿
        ttk.Label(frm, text='英文指示稿 (红字标注版):').grid(row=0, column=0, sticky='w', **pad)
        self.var_base = tk.StringVar()
        tk.Entry(frm, textvariable=self.var_base).grid(row=0, column=1, sticky='ew', **pad)
        ttk.Button(frm, text='浏览…', command=lambda: self.pick(self.var_base, 'pdf')).grid(row=0, column=2, **pad)

        # 锚定原稿(可选)
        ttk.Label(frm, text='客户锚定原稿 (可选):').grid(row=1, column=0, sticky='w', **pad)
        self.var_anchor = tk.StringVar()
        tk.Entry(frm, textvariable=self.var_anchor).grid(row=1, column=1, sticky='ew', **pad)
        ttk.Button(frm, text='浏览…', command=lambda: self.pick(self.var_anchor, 'pdf')).grid(row=1, column=2, **pad)

        # 多国语文件夹
        ttk.Label(frm, text='多国语PDF文件夹:').grid(row=2, column=0, sticky='w', **pad)
        self.var_dir = tk.StringVar()
        tk.Entry(frm, textvariable=self.var_dir).grid(row=2, column=1, sticky='ew', **pad)
        ttk.Button(frm, text='浏览…', command=lambda: self.pick(self.var_dir, 'dir')).grid(row=2, column=2, **pad)

        frm.columnconfigure(1, weight=1)

        # 操作区
        bar = ttk.Frame(root, padding=(10, 0))
        bar.pack(fill='x')
        self.chk_open = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text='完成后自动打开报告(HTML)', variable=self.chk_open).pack(side='left')
        self.btn_run = ttk.Button(bar, text='开始校对', command=self.on_run)
        self.btn_run.pack(side='right', padx=10)

        # 日志区
        lbl = ttk.Label(root, text='处理日志:', padding=(10, 6))
        lbl.pack(anchor='w')
        self.txt = tk.Text(root, height=16, wrap='word', state='disabled',
                           font=('Consolas', 9))
        self.txt.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        self.txt.tag_config('ok', foreground='#1a7f37')
        self.txt.tag_config('err', foreground='#cf222e')
        self.txt.tag_config('info', foreground='#0969da')
        self.txt.tag_config('warn', foreground='#e3b341')

        # 进度条
        self.prog = ttk.Progressbar(root, mode='indeterminate')
        self.prog.pack(fill='x', padx=10, pady=(0, 8))

        root.protocol('WM_DELETE_WINDOW', self.on_close)
        self.root.after(100, self.pump)

    # ---------- 控件 ----------
    def pick(self, var: tk.StringVar, kind: str):
        if kind == 'pdf':
            path = filedialog.askopenfilename(
                title='选择 PDF 文件', filetypes=[('PDF 文件', '*.pdf')])
        else:
            path = filedialog.askdirectory(title='选择多国语 PDF 文件夹')
        if path:
            var.set(path)

    def log(self, msg: str, tag: str = ''):
        self.q.put(('log', msg, tag))

    def done(self, ok: bool, paths=None):
        self.q.put(('done', ok, paths))

    # ---------- 运行 ----------
    def on_run(self):
        if self.busy:
            return
        base = self.var_base.get().strip()
        anchor = self.var_anchor.get().strip() or None
        data_dir = self.var_dir.get().strip()
        if not base or not os.path.isfile(base):
            messagebox.showerror(APP_TITLE, '请选择存在的英文指示稿 PDF')
            return
        if anchor and not os.path.isfile(anchor):
            messagebox.showerror(APP_TITLE, '锚定原稿不存在, 请重新选择')
            return
        if not data_dir or not os.path.isdir(data_dir):
            messagebox.showerror(APP_TITLE, '请选择存在的多国语 PDF 文件夹')
            return

        self.busy = True
        self.btn_run.config(state='disabled', text='处理中…')
        self.prog.start(12)
        self.set_log_text('')
        t = threading.Thread(target=self._worker, args=(base, anchor, data_dir), daemon=True)
        t.start()

    def _worker(self, base, anchor, data_dir):
        try:
            def logf(msg, end='\n'):
                self.log(str(msg), '')
            paths = mee_checker.run_job(base, anchor, data_dir, None, log=logf)
            self.done(True, paths)
        except Exception as e:
            import traceback
            self.done(False, None)
            self.log(f'发生异常: {e}', 'err')
            self.log(traceback.format_exc(), 'err')

    # ---------- 事件泵 ----------
    def pump(self):
        try:
            while True:
                item = self.q.get_nowait()
                if item[0] == 'log':
                    self.append_log(item[1], item[2])
                elif item[0] == 'done':
                    self.finish(item[1], item[2])
        except queue.Empty:
            pass
        self.root.after(100, self.pump)

    def append_log(self, msg: str, tag: str = ''):
        self.txt.config(state='normal')
        self.txt.insert('end', msg + '\n')
        if tag:
            self.txt.tag_add(tag, 'end-1c', 'end')
        self.txt.config(state='disabled')
        self.txt.see('end')
        self.root.update_idletasks()

    def set_log_text(self, s: str):
        self.txt.config(state='normal')
        self.txt.delete('1.0', 'end')
        self.txt.insert('1.0', s)
        self.txt.config(state='disabled')

    def finish(self, ok, paths):
        self.busy = False
        self.prog.stop()
        self.btn_run.config(state='normal', text='开始校对')
        if ok and paths:
            out_dir = paths[3]
            report_html = paths[1]
            self.log(f'\n✓ 校对完成! 输出目录: {out_dir}', 'ok')
            self.log(f'  报告: {paths[0]}', 'ok')
            self.log(f'  HTML: {report_html}', 'ok')
            self.log(f'  截图: {paths[2]}', 'info')
            self.out_paths = paths
            if self.chk_open.get() and os.path.exists(report_html):
                try:
                    import webbrowser
                    webbrowser.open('file:///' + os.path.abspath(report_html).replace('\\', '/'))
                except Exception:
                    pass
            messagebox.showinfo(APP_TITLE, f'校对完成!\n报告目录: {out_dir}')
        else:
            self.log('✗ 校对失败, 请查看上方错误信息', 'err')
            messagebox.showerror(APP_TITLE, '校对失败, 详见日志')

    def on_close(self):
        if self.busy:
            if not messagebox.askyesno(APP_TITLE, '正在处理中, 确认退出?'):
                return
        self.root.destroy()


def main():
    if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    root = tk.Tk()
    GUIApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
