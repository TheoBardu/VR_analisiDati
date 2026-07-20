#! /opt/anaconda3/bin/python3

"""Tab "Documento": compilazione del contesto e generazione della relazione Word.

Logica invariata rispetto alla GUI precedente; il form viene ricavato leggendo
utility/write_docx.py con ast (mai eseguito).
"""

import ast
import json
import os
import threading
from collections import namedtuple
from enum import Enum, auto

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from GUI.widgets import (
    COLOR_ACCENT, COLOR_BG, COLOR_ERROR, COLOR_LABEL, COLOR_PRIMARY,
    COLOR_SUCCESS, FONT_LABEL,
)

WRITE_DOCX_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'utility', 'write_docx.py')

# Grouping of context keys into named sections; unlisted keys fall into "Altro"
FIELD_GROUPS = {
    'Azienda': [
        'nome_azienda', 'indirizzo_azienda', 'sede_legale',
        'ubicazione_unita_operativa', 'attivita_azienda', 'gruppo_appartenenza',
    ],
    'Documento': [
        'note_titolo', 'revisione', 'data_revisione', 'data_scadenza', 'motivo_revisione',
    ],
    'Figure responsabili': ['datore_di_lavoro', 'RSPP', 'medico_competente', 'RLS'],
    'Misurazioni': ['giornate', 'date_misurazione', 'strumentazione', 'condizioni_meteo'],
    'Rischi aggiuntivi': [
        'sostanze_ototossiche', 'misure_attuative_ototossiche',
        'interazione_vib_rum', 'misure_attuative_vib_rum',
        'effetti_indesiderati', 'misure_attuative_effetti_indesiderati',
    ],
    'Descrizione attività': ['descrizione_attivita_dettaglio'],
    'Logo aziendale': ['img_logo_azienda'],
}

TABLE_NAME_MAP = {
    'context_tabella_dpi':      'Tabella DPI',
    'context_tabella_orari':    'Tabella Orari',
    'context_tabella_mansioni': 'Tabella Mansioni',
    'context_tabella_heg':      'Tabella HEG',
}

TABLE_ORDER = list(TABLE_NAME_MAP.keys())


# ── FieldType ──────────────────────────────────────────────────────────────────
class FieldType(Enum):
    STRING           = auto()
    MULTILINE_STRING = auto()
    LIST_OF_STRING   = auto()
    INLINE_IMAGE     = auto()
    FILE_PATH        = auto()
    INTEGER          = auto()


ImageFieldWidgets = namedtuple('ImageFieldWidgets', ['path_entry', 'width_spinbox'])


# ── WriteDocxParser ────────────────────────────────────────────────────────────
class WriteDocxParser:
    """Parses write_docx.py with ast — never executes it."""

    def __init__(self, filepath=WRITE_DOCX_PATH):
        self.filepath = filepath

    def parse(self):
        """Returns (context_fields, table_contexts, scalars)."""
        with open(self.filepath, 'r', encoding='utf-8') as f:
            source = f.read()
        tree = ast.parse(source)

        scalars = {}
        context_fields = []
        table_contexts = []

        for node in tree.body:
            if not isinstance(node, ast.Assign):
                continue
            if len(node.targets) != 1 or not isinstance(node.targets[0], ast.Name):
                continue
            name = node.targets[0].id
            val = node.value

            if isinstance(val, ast.Constant):
                scalars[name] = val.value
                continue

            if name == 'context' and isinstance(val, ast.Dict):
                for k_node, v_node in zip(val.keys, val.values):
                    try:
                        key = ast.literal_eval(k_node)
                    except Exception:
                        continue
                    field = self._classify_value(key, v_node, scalars)
                    if field:
                        context_fields.append(field)
                continue

            if name.startswith('context_tabella_') and isinstance(val, ast.Dict):
                tc = self._extract_table(name, val)
                if tc:
                    table_contexts.append(tc)

        table_contexts.sort(
            key=lambda t: TABLE_ORDER.index(t['var_name'])
            if t['var_name'] in TABLE_ORDER else 999
        )
        return context_fields, table_contexts, scalars

    def _classify_value(self, key, node, scalars):
        if isinstance(node, ast.Constant):
            val = node.value if node.value is not None else ''
            is_long = len(str(val)) > 80 or '\n' in str(val)
            ftype = FieldType.MULTILINE_STRING if is_long else FieldType.STRING
            return {'key': key, 'type': ftype, 'default': val}

        if isinstance(node, ast.List):
            try:
                items = [ast.literal_eval(e) for e in node.elts]
                if all(isinstance(i, str) for i in items):
                    return {'key': key, 'type': FieldType.LIST_OF_STRING, 'default': items}
            except Exception:
                pass
            return None

        if isinstance(node, ast.Call):
            func_name = (node.func.id if isinstance(node.func, ast.Name)
                         else node.func.attr if isinstance(node.func, ast.Attribute)
                         else '')
            if func_name == 'InlineImage':
                default_path = ''
                default_width = 50
                if len(node.args) >= 2:
                    arg = node.args[1]
                    default_path = (arg.value if isinstance(arg, ast.Constant)
                                    else scalars.get(arg.id, '') if isinstance(arg, ast.Name)
                                    else '')
                for kw in node.keywords:
                    if kw.arg == 'width':
                        w_node = kw.value
                        if isinstance(w_node, ast.Call) and w_node.args:
                            inner = w_node.args[0]
                            if isinstance(inner, ast.Constant):
                                default_width = int(inner.value)
                            elif isinstance(inner, ast.Name):
                                default_width = int(scalars.get(inner.id, 50))
                return {
                    'key': key, 'type': FieldType.INLINE_IMAGE,
                    'default_path': default_path, 'default_width': default_width,
                }
        return None

    def _extract_table(self, var_name, node):
        if len(node.keys) != 1:
            return None
        try:
            table_key = ast.literal_eval(node.keys[0])
        except Exception:
            return None
        list_node = node.values[0]
        if not isinstance(list_node, ast.List) or not list_node.elts:
            return None

        columns = []
        initial_rows = []
        for row_node in list_node.elts:
            if not isinstance(row_node, ast.Dict):
                continue
            row = {}
            for k_node, v_node in zip(row_node.keys, row_node.values):
                try:
                    col_key = ast.literal_eval(k_node)
                    col_val = ast.literal_eval(v_node) if isinstance(v_node, ast.Constant) else ''
                    row[col_key] = str(col_val) if col_val is not None else ''
                except Exception:
                    continue
            if not columns and row:
                columns = [{'key': k, 'default': row.get(k, '')} for k in row]
            if row:
                initial_rows.append(row)

        if not columns:
            return None

        label = TABLE_NAME_MAP.get(
            var_name,
            var_name.replace('context_tabella_', 'Tabella ').replace('_', ' ').title()
        )
        return {
            'var_name': var_name,
            'table_key': table_key,
            'columns': columns,
            'initial_rows': initial_rows,
            'label': label,
        }


# ── TableEditor ─────────────────────────────────────────────────────────────
class TableEditor(ttk.Frame):
    MIN_COL_W = 10

    def __init__(self, parent, columns, initial_rows=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.columns = columns
        self.rows = []
        self._build_chrome()
        for row_data in (initial_rows or [{}]):
            self.add_row(values=row_data)

    def _build_chrome(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(fill='x', padx=6, pady=(6, 0))
        ttk.Button(toolbar, text='+ Aggiungi riga', command=self.add_row).pack(side='left', padx=2)
        self._btn_remove = ttk.Button(toolbar, text='- Rimuovi ultima', command=self.remove_last_row)
        self._btn_remove.pack(side='left', padx=2)

        container = ttk.Frame(self)
        container.pack(fill='both', expand=True, padx=6, pady=6)

        self._canvas = tk.Canvas(container, bg=COLOR_BG, highlightthickness=0)
        hsb = ttk.Scrollbar(container, orient='horizontal', command=self._canvas.xview)
        vsb = ttk.Scrollbar(container, orient='vertical', command=self._canvas.yview)
        self._canvas.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set)
        hsb.pack(side='bottom', fill='x')
        vsb.pack(side='right', fill='y')
        self._canvas.pack(side='left', fill='both', expand=True)

        self._inner = ttk.Frame(self._canvas)
        self._win_id = self._canvas.create_window((0, 0), window=self._inner, anchor='nw')
        self._inner.bind('<Configure>', lambda e: self._canvas.configure(
            scrollregion=self._canvas.bbox('all')))
        self._canvas.bind('<Configure>', lambda e: self._canvas.itemconfig(
            self._win_id, width=max(e.width, self._inner.winfo_reqwidth())))
        self._canvas.bind('<Enter>', lambda e: self._canvas.bind_all(
            '<MouseWheel>', lambda ev: self._canvas.yview_scroll(int(-1 * (ev.delta / 120)), 'units')))
        self._canvas.bind('<Leave>', lambda e: self._canvas.unbind_all('<MouseWheel>'))

        # Header
        header = ttk.Frame(self._inner)
        header.grid(row=0, column=0, sticky='ew', padx=2, pady=(2, 0))
        for i, col in enumerate(self.columns):
            w = max(self.MIN_COL_W, len(col['key']) + 2)
            lbl = ttk.Label(header, text=col['key'], font=('Helvetica', 9, 'bold'),
                            foreground=COLOR_PRIMARY, width=w, anchor='center',
                            relief='groove', padding=(2, 2))
            lbl.grid(row=0, column=i, padx=1, sticky='ew')
            header.columnconfigure(i, minsize=w * 8)

        self._rows_frame = ttk.Frame(self._inner)
        self._rows_frame.grid(row=1, column=0, sticky='ew')
        self._update_buttons()

    def add_row(self, values=None):
        values = values or {}
        rf = ttk.Frame(self._rows_frame)
        rf.pack(fill='x', pady=1)
        row_widgets = {'_frame': rf}
        for i, col in enumerate(self.columns):
            w = max(self.MIN_COL_W, len(col['key']) + 2)
            e = ttk.Entry(rf, width=w, font=FONT_LABEL)
            e.insert(0, values.get(col['key'], col.get('default', '')))
            e.grid(row=0, column=i, padx=1, sticky='ew')
            rf.columnconfigure(i, minsize=w * 8)
            row_widgets[col['key']] = e
        self.rows.append(row_widgets)
        self._update_buttons()
        self._canvas.after_idle(lambda: self._canvas.configure(
            scrollregion=self._canvas.bbox('all')))

    def remove_last_row(self):
        if not self.rows:
            return
        self.rows.pop()['_frame'].destroy()
        self._update_buttons()

    def _update_buttons(self):
        self._btn_remove.configure(state='normal' if self.rows else 'disabled')

    def get_rows(self):
        return [
            {k: w.get() for k, w in row.items() if k != '_frame'}
            for row in self.rows
        ]

    def reset(self, initial_rows=None):
        for row in self.rows:
            row['_frame'].destroy()
        self.rows.clear()
        for row_data in (initial_rows or [{}]):
            self.add_row(values=row_data)




# ── DocumentoTab ───────────────────────────────────────────────────────────────
class DocumentoTab(ttk.Frame):
    """Form per la relazione Word, ricavato da write_docx.py."""

    def __init__(self, parent, root, **kwargs):
        super().__init__(parent, **kwargs)
        self.root = root
        self._widget_refs = {}
        self._table_editors = {}
        self._context_fields = []
        self._table_contexts = []
        self._scalars = {}
        self._entry_template = None
        self._entry_output = None

        self._load_data(warn=True)
        self._build()

    def _load_data(self, warn=False):
        try:
            self._context_fields, self._table_contexts, self._scalars = \
                WriteDocxParser().parse()
            return True
        except Exception as e:
            if warn:
                print(f'ATTENZIONE: impossibile leggere write_docx.py: {e}')
            return False

    # ── Costruzione ────────────────────────────────────────────────────────────
    def _build(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(fill='x', padx=10, pady=(10, 2))
        ttk.Label(toolbar, text='Dati per la relazione Word',
                  style='Header.TLabel').pack(side='left')
        ttk.Button(toolbar, text='Genera Documento',
                   command=self._generate_document).pack(side='right', padx=3)
        ttk.Button(toolbar, text='Salva JSON', command=self._save_json).pack(side='right', padx=3)
        ttk.Button(toolbar, text='Carica JSON', command=self._load_json).pack(side='right', padx=3)
        ttk.Button(toolbar, text='↻ Ricarica',
                   command=self._reload_write_docx).pack(side='right', padx=3)

        # Percorsi template / output
        paths_lf = ttk.LabelFrame(self, text='Template e output', padding=8)
        paths_lf.pack(fill='x', padx=10, pady=(6, 2))
        paths_lf.columnconfigure(1, weight=1)

        ttk.Label(paths_lf, text='Template (.docx)', font=FONT_LABEL,
                  foreground=COLOR_LABEL).grid(row=0, column=0, sticky='e', padx=(0, 8))
        self._entry_template = ttk.Entry(paths_lf, font=FONT_LABEL)
        self._entry_template.insert(0, self._scalars.get('documento_word_template', ''))
        self._entry_template.grid(row=0, column=1, sticky='ew', pady=3)
        ttk.Button(paths_lf, text='Sfoglia…',
                   command=lambda: self._browse_open(
                       self._entry_template, [('Word', '*.docx'), ('Tutti', '*.*')])
                   ).grid(row=0, column=2, padx=(6, 0))

        ttk.Label(paths_lf, text='Output (.docx)', font=FONT_LABEL,
                  foreground=COLOR_LABEL).grid(row=1, column=0, sticky='e', padx=(0, 8))
        self._entry_output = ttk.Entry(paths_lf, font=FONT_LABEL)
        self._entry_output.insert(0, self._scalars.get('OUTPUT_DOCUMENT', ''))
        self._entry_output.grid(row=1, column=1, sticky='ew', pady=3)
        ttk.Button(paths_lf, text='Salva come…',
                   command=lambda: self._browse_save(self._entry_output)
                   ).grid(row=1, column=2, padx=(6, 0))

        self._lbl_status = ttk.Label(self, text='', style='Status.TLabel')
        self._lbl_status.pack(fill='x', padx=10, pady=(2, 2))

        inner_nb = ttk.Notebook(self)
        inner_nb.pack(fill='both', expand=True, padx=10, pady=(2, 10))

        tab_gen = ttk.Frame(inner_nb)
        inner_nb.add(tab_gen, text='  Dati Generali  ')
        self._build_dati_generali(tab_gen)

        for tc in self._table_contexts:
            tab_t = ttk.Frame(inner_nb)
            inner_nb.add(tab_t, text=f'  {tc["label"]}  ')
            te = TableEditor(tab_t, tc['columns'], initial_rows=tc['initial_rows'])
            te.pack(fill='both', expand=True)
            self._table_editors[tc['var_name']] = te

    def _build_dati_generali(self, parent):
        canvas = tk.Canvas(parent, bg=COLOR_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        scrollable = ttk.Frame(canvas)
        win_id = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        scrollable.columnconfigure(0, weight=1)

        scrollable.bind('<Configure>',
                        lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.bind('<Enter>', lambda e: canvas.bind_all(
            '<MouseWheel>',
            lambda ev: canvas.yview_scroll(int(-1 * (ev.delta / 120)), 'units')))
        canvas.bind('<Leave>', lambda e: canvas.unbind_all('<MouseWheel>'))

        if not self._context_fields:
            ttk.Label(scrollable,
                      text='Nessun campo trovato in utility/write_docx.py.',
                      style='Muted.TLabel').grid(row=0, column=0, padx=20, pady=20)
            return

        key_to_field = {f['key']: f for f in self._context_fields}
        all_grouped = {k for keys in FIELD_GROUPS.values() for k in keys}

        grid_row = 0
        for group_name, group_keys in FIELD_GROUPS.items():
            fields = [key_to_field[k] for k in group_keys if k in key_to_field]
            if not fields:
                continue
            lf = ttk.LabelFrame(scrollable, text=group_name, padding=10)
            lf.grid(row=grid_row, column=0, sticky='ew', padx=10, pady=5)
            lf.columnconfigure(1, weight=1)
            grid_row += 1
            for i, f in enumerate(fields):
                self._render_field(lf, f, i)

        extras = [f for f in self._context_fields if f['key'] not in all_grouped]
        if extras:
            lf = ttk.LabelFrame(scrollable, text='Altro', padding=10)
            lf.grid(row=grid_row, column=0, sticky='ew', padx=10, pady=5)
            lf.columnconfigure(1, weight=1)
            for i, f in enumerate(extras):
                self._render_field(lf, f, i)

    def _render_field(self, parent, field, row):
        ttk.Label(parent, text=field['key'], font=FONT_LABEL,
                  foreground=COLOR_LABEL, anchor='e').grid(
            row=row, column=0, sticky='ne', padx=(0, 8), pady=3)
        self._widget_refs[field['key']] = self._build_field_widget(parent, field, row)

    def _build_field_widget(self, parent, field, row):
        ft = field['type']

        if ft == FieldType.STRING:
            e = ttk.Entry(parent, font=FONT_LABEL)
            e.insert(0, field.get('default', ''))
            e.grid(row=row, column=1, sticky='ew', pady=3)
            return e

        if ft == FieldType.MULTILINE_STRING:
            t = tk.Text(parent, font=FONT_LABEL, height=4, wrap=tk.WORD,
                        relief='solid', bd=1, bg='white')
            t.insert('1.0', field.get('default', ''))
            t.grid(row=row, column=1, sticky='ew', pady=3)
            return t

        if ft == FieldType.LIST_OF_STRING:
            frame = ttk.Frame(parent)
            frame.grid(row=row, column=1, sticky='ew', pady=3)
            frame.columnconfigure(1, weight=1)
            entries = []
            for i, val in enumerate(field.get('default', [])):
                ttk.Label(frame, text=f'[{i}]', font=FONT_LABEL,
                          foreground=COLOR_LABEL).grid(row=i, column=0, sticky='e', padx=(0, 4))
                e = ttk.Entry(frame, font=FONT_LABEL)
                e.insert(0, val)
                e.grid(row=i, column=1, sticky='ew', pady=1)
                entries.append(e)
            return entries

        if ft == FieldType.INLINE_IMAGE:
            frame = ttk.Frame(parent)
            frame.grid(row=row, column=1, sticky='ew', pady=3)
            frame.columnconfigure(0, weight=1)
            path_e = ttk.Entry(frame, font=FONT_LABEL)
            path_e.insert(0, field.get('default_path', ''))
            path_e.grid(row=0, column=0, sticky='ew', padx=(0, 4))
            ttk.Button(frame, text='Sfoglia',
                       command=lambda e=path_e: self._browse_open(
                           e, [('Immagini', '*.jpg *.jpeg *.png *.bmp'), ('Tutti', '*.*')])
                       ).grid(row=0, column=1)
            ttk.Label(frame, text='Larghezza (mm):', font=FONT_LABEL).grid(
                row=1, column=0, sticky='w', pady=(4, 0))
            spin = ttk.Spinbox(frame, from_=10, to=300, increment=5, width=8, font=FONT_LABEL)
            spin.set(field.get('default_width', 50))
            spin.grid(row=1, column=1, sticky='w', pady=(4, 0))
            return ImageFieldWidgets(path_e, spin)

        return None

    # ── Browse ─────────────────────────────────────────────────────────────────
    def _browse_open(self, entry, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    def _browse_save(self, entry):
        path = filedialog.asksaveasfilename(
            defaultextension='.docx', filetypes=[('Word', '*.docx'), ('Tutti', '*.*')])
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    # ── Valori dei campi ───────────────────────────────────────────────────────
    def _get_field_value(self, field_type, ref):
        if ref is None:
            return ''
        if field_type == FieldType.STRING:
            return ref.get()
        if field_type == FieldType.MULTILINE_STRING:
            return ref.get('1.0', 'end-1c')
        if field_type == FieldType.LIST_OF_STRING:
            return [e.get() for e in ref]
        if field_type == FieldType.INLINE_IMAGE:
            try:
                width = int(ref.width_spinbox.get())
            except ValueError:
                width = 50
            return {'path': ref.path_entry.get(), 'width': width}
        if field_type == FieldType.INTEGER:
            try:
                return int(ref.get())
            except ValueError:
                return 0
        return ''

    def _set_field_value(self, field_type, ref, value):
        if ref is None:
            return
        if field_type == FieldType.STRING:
            ref.delete(0, tk.END)
            ref.insert(0, str(value))
        elif field_type == FieldType.MULTILINE_STRING:
            ref.delete('1.0', tk.END)
            ref.insert('1.0', str(value))
        elif field_type == FieldType.LIST_OF_STRING:
            for i, entry in enumerate(ref):
                entry.delete(0, tk.END)
                if isinstance(value, list) and i < len(value):
                    entry.insert(0, str(value[i]))
        elif field_type == FieldType.INLINE_IMAGE:
            if isinstance(value, dict):
                ref.path_entry.delete(0, tk.END)
                ref.path_entry.insert(0, value.get('path', ''))
                ref.width_spinbox.set(value.get('width', 50))
        elif field_type == FieldType.INTEGER:
            ref.set(str(value))

    # ── Ricarica ───────────────────────────────────────────────────────────────
    def _reload_write_docx(self):
        try:
            new_fields, new_tables, new_scalars = WriteDocxParser().parse()
        except Exception as e:
            messagebox.showerror('Errore', f'Impossibile ricaricare write_docx.py:\n{e}')
            return

        new_field_map = {f['key']: f for f in new_fields}
        for field in self._context_fields:
            key = field['key']
            ref = self._widget_refs.get(key)
            if ref is None:
                continue
            if key in new_field_map:
                nf = new_field_map[key]
                if field['type'] == FieldType.INLINE_IMAGE:
                    value = {'path': nf.get('default_path', ''),
                             'width': nf.get('default_width', 50)}
                else:
                    value = nf.get('default', '')
                self._set_field_value(field['type'], ref, value)
            else:
                print(f'[Ricarica] Chiave rimossa da write_docx.py: {key}')

        for key in new_field_map:
            if key not in self._widget_refs:
                print(f'[Ricarica] Nuovo campo: {key!r}. Riavvia l\'app per vederlo nel form.')

        for tc in new_tables:
            te = self._table_editors.get(tc['var_name'])
            if te:
                te.reset(tc['initial_rows'])

        self._scalars = new_scalars
        self._entry_template.delete(0, tk.END)
        self._entry_template.insert(0, new_scalars.get('documento_word_template', ''))
        self._entry_output.delete(0, tk.END)
        self._entry_output.insert(0, new_scalars.get('OUTPUT_DOCUMENT', ''))
        self._lbl_status.configure(text='Form ricaricato da write_docx.py',
                                   foreground=COLOR_SUCCESS)

    # ── Generazione ────────────────────────────────────────────────────────────
    def _generate_document(self):
        template_path = self._entry_template.get().strip()
        output_path = self._entry_output.get().strip()

        if not template_path or not os.path.isfile(template_path):
            messagebox.showerror('Errore', f'File template non trovato:\n{template_path}')
            return
        if not output_path:
            messagebox.showerror('Errore', 'Specifica il percorso di output.')
            return

        self._lbl_status.configure(text='Generando documento…', foreground=COLOR_ACCENT)
        threading.Thread(target=self._run_generate,
                         args=(template_path, output_path), daemon=True).start()

    def _run_generate(self, template_path, output_path):
        try:
            from docxtpl import DocxTemplate, InlineImage
            from docx.shared import Mm

            doc = DocxTemplate(template_path)
            context = {}

            for field in self._context_fields:
                ref = self._widget_refs.get(field['key'])
                if ref is None:
                    continue
                raw = self._get_field_value(field['type'], ref)
                if field['type'] == FieldType.INLINE_IMAGE:
                    img_path = raw.get('path', '')
                    context[field['key']] = (
                        InlineImage(doc, img_path, width=Mm(raw.get('width', 50)))
                        if img_path and os.path.isfile(img_path) else '')
                else:
                    context[field['key']] = raw

            for tc in self._table_contexts:
                te = self._table_editors.get(tc['var_name'])
                if te:
                    context[tc['table_key']] = te.get_rows()

            doc.render(context)
            doc.save(output_path)
            self.root.after(0, lambda p=output_path: self._lbl_status.configure(
                text=f'Salvato: {p}', foreground=COLOR_SUCCESS))
        except Exception as e:
            msg = str(e)
            def on_error(m=msg):
                self._lbl_status.configure(text=f'Errore: {m}', foreground=COLOR_ERROR)
                messagebox.showerror('Errore generazione', m)
            self.root.after(0, on_error)

    # ── JSON ───────────────────────────────────────────────────────────────────
    def _collect_form_data(self):
        data = {'context': {}, 'tables': {}}
        for field in self._context_fields:
            ref = self._widget_refs.get(field['key'])
            if ref is not None:
                data['context'][field['key']] = self._get_field_value(field['type'], ref)
        for tc in self._table_contexts:
            te = self._table_editors.get(tc['var_name'])
            if te:
                data['tables'][tc['table_key']] = te.get_rows()
        data['template_path'] = self._entry_template.get()
        data['output_path'] = self._entry_output.get()
        return data

    def _save_json(self):
        path = filedialog.asksaveasfilename(
            defaultextension='.json', title='Salva preset',
            filetypes=[('JSON', '*.json'), ('Tutti', '*.*')])
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self._collect_form_data(), f, ensure_ascii=False, indent=2)
            self._lbl_status.configure(text=f'Preset salvato: {path}',
                                       foreground=COLOR_SUCCESS)
        except Exception as e:
            messagebox.showerror('Errore', f'Impossibile salvare:\n{e}')

    def _load_json(self):
        path = filedialog.askopenfilename(
            title='Carica preset', filetypes=[('JSON', '*.json'), ('Tutti', '*.*')])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror('Errore', f'Impossibile leggere il file:\n{e}')
            return

        field_map = {f['key']: f for f in self._context_fields}
        for key, value in data.get('context', {}).items():
            if key in self._widget_refs and key in field_map:
                self._set_field_value(field_map[key]['type'], self._widget_refs[key], value)
            else:
                print(f'[Carica JSON] Chiave non presente nel form: {key!r}')

        for tc in self._table_contexts:
            rows = data.get('tables', {}).get(tc['table_key'])
            if rows is not None:
                te = self._table_editors.get(tc['var_name'])
                if te:
                    te.reset(rows)

        if 'template_path' in data:
            self._entry_template.delete(0, tk.END)
            self._entry_template.insert(0, data['template_path'])
        if 'output_path' in data:
            self._entry_output.delete(0, tk.END)
            self._entry_output.insert(0, data['output_path'])

        self._lbl_status.configure(text=f'Preset caricato: {path}',
                                   foreground=COLOR_SUCCESS)
