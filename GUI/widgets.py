#! /opt/anaconda3/bin/python3

"""Widget riutilizzabili e costanti di stile per la GUI."""

import queue
import tkinter as tk
from tkinter import ttk

# ── Stile ──────────────────────────────────────────────────────────────────────
COLOR_BG = '#F5F5F5'
COLOR_PRIMARY = '#1565C0'
COLOR_ACCENT = '#1E88E5'
COLOR_SUCCESS = '#2E7D32'
COLOR_WARNING = '#EF6C00'
COLOR_ERROR = '#C62828'
COLOR_LABEL = '#37474F'
COLOR_MUTED = '#78909C'
COLOR_CONSOLE_BG = '#1E1E1E'
COLOR_CONSOLE_FG = '#D4D4D4'

# Sfondi righe per classe di rischio (coerenti con colora_classe_rischio)
COLOR_RISCHIO = {
    'BASSA': '#E8F5E9',
    'MEDIA': '#FFF8E1',
    'ALTA': '#FFEBEE',
}

FONT_TITLE = ('Helvetica', 15, 'bold')
FONT_SUBTITLE = ('Helvetica', 9)
FONT_HEADER = ('Helvetica', 11, 'bold')
FONT_LABEL = ('Helvetica', 10)
FONT_SMALL = ('Helvetica', 9)
FONT_MONO = ('Menlo', 10)


def setup_styles():
    style = ttk.Style()
    for theme in ('aqua', 'clam', 'alt', 'default'):
        try:
            style.theme_use(theme)
            break
        except Exception:
            continue
    style.configure('TNotebook.Tab', font=FONT_LABEL, padding=[16, 7])
    style.configure('TLabelframe.Label', font=('Helvetica', 10, 'bold'),
                    foreground=COLOR_PRIMARY)
    style.configure('Header.TLabel', font=FONT_HEADER, foreground=COLOR_PRIMARY)
    style.configure('Status.TLabel', font=FONT_SMALL, foreground=COLOR_LABEL)
    style.configure('Muted.TLabel', font=FONT_SMALL, foreground=COLOR_MUTED)
    style.configure('Run.TButton', font=('Helvetica', 11, 'bold'), padding=(10, 8))
    style.configure('Step.TButton', font=FONT_SMALL, padding=(6, 6))
    style.configure('Treeview.Heading', font=('Helvetica', 9, 'bold'))
    style.configure('Treeview', rowheight=22, font=FONT_SMALL)
    return style


# ── UiQueue ────────────────────────────────────────────────────────────────────
class UiQueue:
    """Marshalla lavoro dai thread di elaborazione al thread della GUI.

    Non si puo' chiamare root.after() da un thread secondario (Tkinter solleva
    "main thread is not in main loop"): i thread accodano qui delle callable e
    un poller in esecuzione sul thread principale le svuota.
    """

    def __init__(self, root, interval_ms=40):
        self._root = root
        self._queue = queue.Queue()
        self._interval = interval_ms
        self._root.after(self._interval, self._drain)

    def post(self, func):
        """Chiamabile da qualsiasi thread."""
        self._queue.put(func)

    def _drain(self):
        while True:
            try:
                func = self._queue.get_nowait()
            except queue.Empty:
                break
            try:
                func()
            except tk.TclError:
                pass
        self._root.after(self._interval, self._drain)


# ── LogView ────────────────────────────────────────────────────────────────────
class LogView(ttk.Frame):
    """Console di sola lettura. Thread-safe tramite UiQueue."""

    def __init__(self, parent, ui_queue, **kwargs):
        super().__init__(parent, **kwargs)
        self._ui = ui_queue

        toolbar = ttk.Frame(self)
        toolbar.pack(fill='x', padx=8, pady=(8, 0))
        ttk.Label(toolbar, text='Log di esecuzione', style='Header.TLabel').pack(side='left')
        ttk.Button(toolbar, text='Salva su file…', command=self._save).pack(side='right', padx=3)
        ttk.Button(toolbar, text='Pulisci', command=self.clear).pack(side='right', padx=3)

        body = ttk.Frame(self)
        body.pack(fill='both', expand=True, padx=8, pady=8)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        self._text = tk.Text(
            body, font=FONT_MONO, bg=COLOR_CONSOLE_BG, fg=COLOR_CONSOLE_FG,
            insertbackground='white', wrap=tk.WORD, relief='flat', bd=0,
            state='disabled', padx=8, pady=6,
        )
        vsb = ttk.Scrollbar(body, orient='vertical', command=self._text.yview)
        self._text.configure(yscrollcommand=vsb.set)
        self._text.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')

        self._text.tag_configure('error', foreground='#FF7043')
        self._text.tag_configure('warning', foreground='#FFCA28')
        self._text.tag_configure('success', foreground='#81C784')
        self._text.tag_configure('step', foreground='#64B5F6',
                                 font=('Menlo', 10, 'bold'))

    @staticmethod
    def _tag_for(chunk):
        low = chunk.lstrip().lower()
        if low.startswith(('errore', 'error', 'traceback')):
            return 'error'
        if low.startswith(('attenzione', 'warning', '⚠')):
            return 'warning'
        if low.startswith('▶'):
            return 'step'
        return None

    def append(self, chunk, tag=None):
        """Chiamabile da qualsiasi thread."""
        self._ui.post(lambda: self._append_now(chunk, tag))

    def _append_now(self, chunk, tag=None):
        try:
            self._text.configure(state='normal')
            self._text.insert(tk.END, chunk, tag or self._tag_for(chunk) or ())
            self._text.see(tk.END)
            self._text.configure(state='disabled')
        except tk.TclError:
            pass

    def clear(self):
        self._text.configure(state='normal')
        self._text.delete('1.0', tk.END)
        self._text.configure(state='disabled')

    def _save(self):
        from tkinter import filedialog, messagebox
        path = filedialog.asksaveasfilename(
            defaultextension='.log', title='Salva log',
            filetypes=[('Log', '*.log'), ('Testo', '*.txt'), ('Tutti', '*.*')])
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(self._text.get('1.0', 'end-1c'))
        except OSError as e:
            messagebox.showerror('Errore', f'Impossibile salvare il log:\n{e}')


class LogRedirector:
    """Sostituto di sys.stdout/sys.stderr che scrive su una LogView."""

    def __init__(self, log_view, tag=None):
        self._log = log_view
        self._tag = tag

    def write(self, string):
        if string:
            self._log.append(string, self._tag)
        return len(string)

    def flush(self):
        pass

    def isatty(self):
        return False


# ── DataFrameView ──────────────────────────────────────────────────────────────
class DataFrameView(ttk.Frame):
    """Treeview di sola lettura su un DataFrame pandas, ordinabile per colonna."""

    MAX_COL_CHARS = 40

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self._df = None
        self._sort_col = None
        self._sort_reverse = False

        self._tree = ttk.Treeview(self, show='headings', selectmode='browse')
        vsb = ttk.Scrollbar(self, orient='vertical', command=self._tree.yview)
        hsb = ttk.Scrollbar(self, orient='horizontal', command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._placeholder = ttk.Label(self, text='', style='Muted.TLabel',
                                      anchor='center')

        for classe, colore in COLOR_RISCHIO.items():
            self._tree.tag_configure(f'rischio_{classe}', background=colore)
        self._tree.tag_configure('dispari', background='#FAFAFA')

    def show_message(self, message):
        self._df = None
        self._tree.grid_remove()
        self._placeholder.configure(text=message)
        self._placeholder.grid(row=0, column=0, sticky='nsew', padx=20, pady=20)

    def set_dataframe(self, df):
        self._placeholder.grid_remove()
        self._tree.grid()
        self._df = df
        self._sort_col = None
        self._render()

    def _render(self):
        df = self._df
        self._tree.delete(*self._tree.get_children())
        cols = [str(c) for c in df.columns]
        self._tree['columns'] = cols

        for col in cols:
            self._tree.heading(col, text=col,
                               command=lambda c=col: self._sort_by(c))
            values = df[col].astype(str)
            widest = max([len(col)] + [len(v) for v in values.head(200)]) if len(df) else len(col)
            width = min(max(widest, 6), self.MAX_COL_CHARS) * 8 + 16
            self._tree.column(col, width=width, minwidth=50, anchor='center',
                              stretch=False)

        classe_col = next((c for c in cols if 'classe' in c.lower()), None)
        for i, (_, row) in enumerate(df.iterrows()):
            tags = []
            if classe_col is not None:
                classe = str(row[classe_col]).strip().upper()
                if classe in COLOR_RISCHIO:
                    tags.append(f'rischio_{classe}')
            if not tags and i % 2:
                tags.append('dispari')
            values = ['' if _is_missing(v) else str(v) for v in row]
            self._tree.insert('', 'end', values=values, tags=tuple(tags))

    def _sort_by(self, col):
        if self._df is None:
            return
        self._sort_reverse = (col == self._sort_col) and not self._sort_reverse
        self._sort_col = col
        try:
            self._df = self._df.sort_values(
                by=col, ascending=not self._sort_reverse, kind='stable')
        except TypeError:
            self._df = self._df.reindex(
                self._df[col].astype(str).sort_values(
                    ascending=not self._sort_reverse, kind='stable').index)
        self._render()


def _is_missing(value):
    return value is None or value != value  # NaN != NaN
