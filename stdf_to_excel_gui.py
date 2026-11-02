"""GUI front-end for stdf_to_excel: pick one or more STDF files, get one combined Excel file."""
import os, threading, tkinter as tk
from tkinter import filedialog, messagebox, ttk

import stdf_to_excel as core


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('STDF Die Parser')
        self.geometry('720x480')
        self.files = []

        top = ttk.Frame(self, padding=8); top.pack(fill='both', expand=True)
        ttk.Label(top, text='STDF files (rows are appended in this order):').pack(anchor='w')
        mid = ttk.Frame(top); mid.pack(fill='both', expand=True, pady=4)
        self.lb = tk.Listbox(mid, selectmode='extended', height=8)
        self.lb.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(mid, command=self.lb.yview); sb.pack(side='left', fill='y')
        self.lb.config(yscrollcommand=sb.set)
        btns = ttk.Frame(mid); btns.pack(side='left', padx=6, anchor='n')
        for text, cmd in [('Add files...', self.add), ('Remove', self.remove), ('Up', lambda: self.move(-1)),
                          ('Down', lambda: self.move(1)), ('Clear', self.clear)]:
            ttk.Button(btns, text=text, command=cmd, width=12).pack(pady=2)

        out = ttk.Frame(top); out.pack(fill='x', pady=4)
        ttk.Label(out, text='Output xlsx:').pack(side='left')
        self.out = tk.StringVar()
        ttk.Entry(out, textvariable=self.out).pack(side='left', fill='x', expand=True, padx=4)
        ttk.Button(out, text='Browse...', command=self.browse).pack(side='left')

        self.run_btn = ttk.Button(top, text='Run', command=self.run); self.run_btn.pack(pady=4)
        self.bar = ttk.Progressbar(top, mode='indeterminate'); self.bar.pack(fill='x')
        self.log = tk.Text(top, height=8, state='disabled'); self.log.pack(fill='both', expand=True, pady=4)

    def say(self, msg):
        self.after(0, self._say, msg)

    def _say(self, msg):
        self.log.config(state='normal'); self.log.insert('end', msg + '\n'); self.log.see('end'); self.log.config(state='disabled')

    def refresh(self):
        self.lb.delete(0, 'end')
        for f in self.files:
            self.lb.insert('end', f)

    def add(self):
        for f in filedialog.askopenfilenames(title='Select STDF files',
                                             filetypes=[('STDF', '*.std *.stdf'), ('All files', '*.*')]):
            if f not in self.files:
                self.files.append(f)
        if self.files and not self.out.get():
            self.out.set(os.path.splitext(self.files[0])[0] + '_dies.xlsx')
        self.refresh()

    def remove(self):
        for i in reversed(self.lb.curselection()):
            del self.files[i]
        self.refresh()

    def move(self, d):
        sel = self.lb.curselection()
        if len(sel) != 1 or not 0 <= sel[0] + d < len(self.files):
            return
        i = sel[0]
        self.files[i], self.files[i + d] = self.files[i + d], self.files[i]
        self.refresh(); self.lb.selection_set(i + d)

    def clear(self):
        self.files = []; self.refresh()

    def browse(self):
        f = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if f:
            self.out.set(f)

    def run(self):
        if not self.files:
            return messagebox.showwarning('No input', 'Add at least one STDF file.')
        if not self.out.get():
            return messagebox.showwarning('No output', 'Choose an output xlsx path.')
        self.run_btn.config(state='disabled'); self.bar.start(10)
        threading.Thread(target=self.work, args=(list(self.files), self.out.get()), daemon=True).start()

    def work(self, files, dst):
        try:
            s = core.parse_files(files, self.say)
            n, m = core.write_xlsx(s, dst)
            self.say(f'Done: {n} rows, {m} test columns -> {dst}')
            self.after(0, messagebox.showinfo, 'Done', f'{n} rows, {m} test columns\n{dst}')
        except Exception as e:
            self.say(f'ERROR: {e!r}')
            self.after(0, messagebox.showerror, 'Error', str(e))
        finally:
            self.after(0, self.finish)

    def finish(self):
        self.bar.stop(); self.run_btn.config(state='normal')


if __name__ == '__main__':
    App().mainloop()
