# -*- coding: utf-8 -*-
"""MEE 多国语数字校对工具 - 图形界面版 (v1.7)

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

APP_TITLE = 'MEE 多国语数字校对工具 v1.7'


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

        # 校对模式
        ttk.Label(frm, text='校对模式:').grid(row=0, column=0, sticky='w', **pad)
        self.var_mode = tk.StringVar(value='red')
        self.cmb_mode = ttk.Combobox(frm, textvariable=self.var_mode, state='readonly', width=46,
                                     values=['red', 'highlight'])
        self.cmb_mode.grid(row=0, column=1, sticky='ew', **pad)
        self.cmb_mode.bind('<<ComboboxSelected>>', self.on_mode_change)
        self.lbl_mode = ttk.Label(frm, text='红字模式: 校全部红字数字 | 指示稿与译文页数需一致',
                                  foreground='#0969da')
        self.lbl_mode.grid(row=0, column=2, columnspan=1, sticky='w')

        # 英文指示稿
        self.lbl_base = ttk.Label(frm, text='英文指示稿 (红字标注版):')
        self.lbl_base.grid(row=1, column=0, sticky='w', **pad)
        self.var_base = tk.StringVar()
        tk.Entry(frm, textvariable=self.var_base).grid(row=1, column=1, sticky='ew', **pad)
        ttk.Button(frm, text='浏览…', command=lambda: self.pick(self.var_base, 'pdf')).grid(row=1, column=2, **pad)

        # 锚定原稿/客户指示稿
        self.lbl_anchor = ttk.Label(frm, text='客户锚定原稿 (可选):')
        self.lbl_anchor.grid(row=2, column=0, sticky='w', **pad)
        self.var_anchor = tk.StringVar()
        tk.Entry(frm, textvariable=self.var_anchor).grid(row=2, column=1, sticky='ew', **pad)
        ttk.Button(frm, text='浏览…', command=lambda: self.pick(self.var_anchor, 'pdf')).grid(row=2, column=2, **pad)

        # 多国语文件夹
        ttk.Label(frm, text='多国语PDF文件夹:').grid(row=3, column=0, sticky='w', **pad)
        self.var_dir = tk.StringVar()
        tk.Entry(frm, textvariable=self.var_dir).grid(row=3, column=1, sticky='ew', **pad)
        ttk.Button(frm, text='浏览…', command=lambda: self.pick(self.var_dir, 'dir')).grid(row=3, column=2, **pad)

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
    def on_mode_change(self, _evt=None):
        if self.var_mode.get() == 'highlight':
            self.lbl_base.config(text='英文指示稿 (高亮模式不需要):')
            self.lbl_anchor.config(text='客户指示稿 (红框+高亮, 必填):')
            self.lbl_mode.config(text='高亮模式: 只校红框内青色高亮内容 | 各文件页数需一致',
                                 foreground='#1a7f37')
        else:
            self.lbl_base.config(text='英文指示稿 (红字标注版):')
            self.lbl_anchor.config(text='客户锚定原稿 (可选):')
            self.lbl_mode.config(text='红字模式: 校全部红字数字 | 指示稿与译文页数需一致',
                                 foreground='#0969da')

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
        mode = self.var_mode.get()
        base = self.var_base.get().strip()
        anchor = self.var_anchor.get().strip() or None
        data_dir = self.var_dir.get().strip()
        if mode == 'highlight':
            if not anchor or not os.path.isfile(anchor):
                messagebox.showerror(APP_TITLE, '高亮模式请选择存在的客户指示稿(红框+高亮) PDF')
                return
        else:
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
        t = threading.Thread(target=self._worker, args=(mode, base, anchor, data_dir), daemon=True)
        t.start()

    def _worker(self, mode, base, anchor, data_dir):
        try:
            def logf(msg, end='\n'):
                self.log(str(msg), '')
            if mode == 'highlight':
                paths = mee_checker.run_highlight_job(anchor, data_dir, None, log=logf)
            else:
                paths = mee_checker.run_job(base, anchor, data_dir, None, log=logf)
            self.done(True, paths)
        except Exception as e:
            import traceback
            self.done(False, None)
            # 先给一句人可读的原因, 再附原始异常与堆栈供维护人排查
            txt = str(e)
            hint = ''
            if 'no such file' in txt.lower() or 'not found' in txt.lower() or '找不到' in txt:
                hint = '原因: 所选 PDF 文件不存在或已被移动, 请重新选择。'
            elif 'cannot open' in txt.lower() or 'encrypted' in txt.lower():
                hint = '原因: 文件损坏或加了口令, 请确认该 PDF 能正常打开。'
            elif 'page' in txt.lower() and 'not in document' in txt.lower():
                hint = '原因: 某份译文页数少于指示稿, 导致跳页越界; 请核对文件版本。'
            elif '全部被当作指示稿' in txt or '待校对文件: 0' in txt:
                hint = '原因: 多国语文件夹里没有可校对的译文 PDF(只有英文稿)。'
            self.log(f'✗ 校对失败: {hint or "见下方详细原因"}', 'err')
            self.log(f'详细信息: {txt}', 'err')
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
            if report_html:
                self.log(f'  HTML: {report_html}', 'ok')
            else:
                self.log('  (高亮模式: 打开 xlsx 与高亮清单.csv 核对提取范围)', 'info')
            self.log(f'  截图: {paths[2]}', 'info')
            self.out_paths = paths
            # 逐文件诊断(页数不一致/无红字/打不开)必须显眼, 不能混在日志里
            dfile = os.path.join(out_dir, '文件诊断.txt')
            if os.path.exists(dfile):
                try:
                    with open(dfile, encoding='utf-8') as fh:
                        txt = fh.read().strip()
                    n = txt.count('[')
                    self.log(f'\n⚠ 文件诊断: 发现 {n} 条问题:', 'warn')
                    for ln in txt.split('\n'):
                        self.log('  ' + ln, 'err' if ln.startswith('[错误]') else 'warn')
                    messagebox.showwarning(
                        APP_TITLE,
                        f'有 {n} 个文件需要确认(页数不一致/无红字/打不开等),\n'
                        f'详见日志与 {dfile}')
                except OSError:
                    pass
            if self.chk_open.get() and report_html and os.path.exists(report_html):
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
