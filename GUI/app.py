#! /opt/anaconda3/bin/python3

"""GUI per l'analisi di Valutazione Rischio Rumore."""

import os
import sys
import threading
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from config import NOME_VR8h_riepilogo, SCHEDA_DPI, SCHEDA_MANSIONI
from GUI import pipeline as pl
from GUI.widgets import (
    COLOR_ACCENT, COLOR_BG, COLOR_ERROR, COLOR_LABEL, COLOR_MUTED,
    COLOR_PRIMARY, COLOR_SUCCESS, COLOR_WARNING,
    FONT_LABEL, FONT_SMALL, FONT_SUBTITLE, FONT_TITLE,
    DataFrameView, LogRedirector, LogView, UiQueue, setup_styles,
)

# Step singoli: (etichetta bottone, funzione, ricarica_riepilogo)
STEP_BUTTONS = [
    ('Leggi misure',            pl.step_lettura_misure,      False),
    ('Calcola medie',           pl.step_medie,               False),
    ('Analisi 8h',              pl.step_analisi_8h,          True),
    ('Applica DPI',             pl.step_dpi,                 True),
    ('Rilievi Fonometrici',     pl.step_rilievi_fonometrici, False),
    ('Esporta XLSX → PDF',      pl.step_export_pdf,          False),
]


class App:
    def __init__(self, root):
        self.root = root
        self.root.title('Analisi Valutazione Rischio Rumore')
        self.root.geometry('1040x780')
        self.root.minsize(880, 620)
        self.root.configure(bg=COLOR_BG)

        self._worker = None
        self._resume = threading.Event()
        self._abort = False
        self._run_buttons = []
        self._paused = False
        self._ui = UiQueue(root)

        setup_styles()
        self._build_header()
        self._build_notebook()
        self._build_statusbar()

        sys.stdout = LogRedirector(self._log)
        sys.stderr = LogRedirector(self._log, tag='error')

    # ── Percorsi ───────────────────────────────────────────────────────────────
    def _paths(self):
        return pl.Paths(
            main_directory=self._var_main.get().strip(),
            misure=self._var_misure.get().strip() or pl.DEFAULT_MISURE,
            risultati=self._var_risultati.get().strip() or pl.DEFAULT_RISULTATI,
            dpi=self._var_dpi.get().strip() or pl.DEFAULT_DPI,
        )

    def _valid_main_directory(self):
        main_dir = self._var_main.get().strip()
        if not main_dir:
            messagebox.showerror('Errore', 'Seleziona la cartella di lavoro.')
            return False
        if not os.path.isdir(main_dir):
            messagebox.showerror('Errore', f'Cartella non trovata:\n{main_dir}')
            return False
        return True

    # ── Header ─────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self.root, bg=COLOR_PRIMARY, pady=14)
        hdr.pack(fill='x')
        tk.Label(hdr, text='Analisi Valutazione Rischio Rumore',
                 font=FONT_TITLE, bg=COLOR_PRIMARY, fg='white').pack()
        tk.Label(hdr, text='VRR Analysis Tool',
                 font=FONT_SUBTITLE, bg=COLOR_PRIMARY, fg='#90CAF9').pack()

    def _build_notebook(self):
        outer = tk.Frame(self.root, bg=COLOR_BG)
        outer.pack(fill='both', expand=True)
        nb = ttk.Notebook(outer)
        nb.pack(fill='both', expand=True, padx=8, pady=8)

        tab_exec = ttk.Frame(nb)
        tab_scheda = ttk.Frame(nb)
        tab_riepilogo = ttk.Frame(nb)
        tab_documento = ttk.Frame(nb)
        tab_log = ttk.Frame(nb)

        nb.add(tab_exec, text='  Esecuzione  ')
        nb.add(tab_scheda, text='  Scheda Gruppi DPI  ')
        nb.add(tab_riepilogo, text='  Riepilogo  ')
        nb.add(tab_documento, text='  Documento  ')
        nb.add(tab_log, text='  Log  ')

        self._notebook = nb
        self._tab_log = tab_log

        # Il log va costruito per primo: gli altri tab ci scrivono.
        self._log = LogView(tab_log, self._ui)
        self._log.pack(fill='both', expand=True)

        self._build_tab_exec(tab_exec)
        self._build_tab_scheda(tab_scheda)
        self._build_tab_riepilogo(tab_riepilogo)

        from GUI.documento import DocumentoTab
        DocumentoTab(tab_documento, self.root).pack(fill='both', expand=True)

    # ── Tab Esecuzione ─────────────────────────────────────────────────────────
    def _build_tab_exec(self, parent):
        # Cartella di lavoro
        lf_dir = ttk.LabelFrame(parent, text='Cartella di lavoro', padding=10)
        lf_dir.pack(fill='x', padx=14, pady=(14, 6))
        lf_dir.columnconfigure(0, weight=1)

        self._var_main = tk.StringVar()
        self._var_main.trace_add('write', lambda *_: self._refresh_path_hint())
        entry = ttk.Entry(lf_dir, textvariable=self._var_main, font=FONT_LABEL)
        entry.grid(row=0, column=0, sticky='ew', padx=(0, 8))
        ttk.Button(lf_dir, text='Sfoglia…', command=self._choose_main).grid(row=0, column=1)

        self._lbl_hint = ttk.Label(lf_dir, text='Nessuna cartella selezionata.',
                                   style='Muted.TLabel')
        self._lbl_hint.grid(row=1, column=0, columnspan=2, sticky='w', pady=(6, 0))

        # Percorsi avanzati (collassabile)
        self._adv_open = tk.BooleanVar(value=False)
        self._btn_adv = ttk.Button(lf_dir, text='▸ Percorsi avanzati',
                                   command=self._toggle_advanced, style='Step.TButton')
        self._btn_adv.grid(row=2, column=0, sticky='w', pady=(8, 0))

        self._frm_adv = ttk.Frame(lf_dir)
        self._frm_adv.columnconfigure(1, weight=1)
        self._var_misure = tk.StringVar(value=pl.DEFAULT_MISURE)
        self._var_risultati = tk.StringVar(value=pl.DEFAULT_RISULTATI)
        self._var_dpi = tk.StringVar(value=pl.DEFAULT_DPI)
        for i, (label, var) in enumerate([
            ('Misure', self._var_misure),
            ('Risultati', self._var_risultati),
            ('DPI', self._var_dpi),
        ]):
            ttk.Label(self._frm_adv, text=label, font=FONT_SMALL,
                      foreground=COLOR_LABEL, width=10, anchor='e').grid(
                row=i, column=0, sticky='e', padx=(0, 8), pady=2)
            e = ttk.Entry(self._frm_adv, textvariable=var, font=FONT_SMALL)
            e.grid(row=i, column=1, sticky='ew', pady=2)
            var.trace_add('write', lambda *_: self._refresh_path_hint())
        ttk.Label(self._frm_adv,
                  text='Percorsi relativi alla cartella di lavoro (es. /misure).',
                  style='Muted.TLabel').grid(row=3, column=1, sticky='w', pady=(4, 0))

        # Esecuzione completa
        lf_run = ttk.LabelFrame(parent, text='Esecuzione completa', padding=10)
        lf_run.pack(fill='x', padx=14, pady=6)

        row = ttk.Frame(lf_run)
        row.pack(fill='x')
        self._btn_auto = ttk.Button(row, text='▶  Modalità automatica',
                                    style='Run.TButton', command=self._run_auto)
        self._btn_auto.pack(side='left', padx=(0, 8))
        self._btn_manual = ttk.Button(row, text='⏸  Modalità manuale',
                                      style='Run.TButton', command=self._run_manual)
        self._btn_manual.pack(side='left')
        self._run_buttons += [self._btn_auto, self._btn_manual]

        ttk.Label(lf_run,
                  text='Automatica: esegue tutto senza fermarsi.   '
                       'Manuale: si ferma prima di ogni step.',
                  style='Muted.TLabel').pack(anchor='w', pady=(8, 0))

        # Pannello di pausa (nascosto finche' non serve)
        self._frm_pause = ttk.Frame(lf_run)
        self._lbl_pause = ttk.Label(self._frm_pause, text='', font=FONT_LABEL,
                                    foreground=COLOR_WARNING, wraplength=700,
                                    justify='left')
        self._lbl_pause.pack(side='top', anchor='w', pady=(0, 6))
        btns = ttk.Frame(self._frm_pause)
        btns.pack(anchor='w')
        ttk.Button(btns, text='Continua ▶', command=self._continue).pack(side='left', padx=(0, 6))
        ttk.Button(btns, text='Interrompi ✕', command=self._stop).pack(side='left')

        # Step singoli
        lf_steps = ttk.LabelFrame(parent, text='Step singoli', padding=10)
        lf_steps.pack(fill='x', padx=14, pady=6)
        grid = ttk.Frame(lf_steps)
        grid.pack(fill='x')
        for i, (label, func, reload_riep) in enumerate(STEP_BUTTONS):
            b = ttk.Button(grid, text=label, style='Step.TButton', width=22,
                           command=lambda f=func, l=label, r=reload_riep:
                           self._run_single(f, l, r))
            b.grid(row=i // 3, column=i % 3, padx=4, pady=4, sticky='ew')
            self._run_buttons.append(b)
        for c in range(3):
            grid.columnconfigure(c, weight=1)

        ttk.Label(parent,
                  text='Gli step singoli ricalcolano al volo i dati intermedi '
                       'di cui hanno bisogno.',
                  style='Muted.TLabel').pack(anchor='w', padx=16, pady=(0, 8))

    def _toggle_advanced(self):
        if self._adv_open.get():
            self._frm_adv.grid_remove()
            self._btn_adv.configure(text='▸ Percorsi avanzati')
            self._adv_open.set(False)
        else:
            self._frm_adv.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(6, 0))
            self._btn_adv.configure(text='▾ Percorsi avanzati')
            self._adv_open.set(True)

    def _choose_main(self):
        folder = filedialog.askdirectory(title='Seleziona la cartella di lavoro')
        if folder:
            self._var_main.set(folder)

    def _refresh_path_hint(self):
        main_dir = self._var_main.get().strip()
        if not main_dir:
            self._lbl_hint.configure(text='Nessuna cartella selezionata.',
                                     foreground=COLOR_MUTED)
            return
        if not os.path.isdir(main_dir):
            self._lbl_hint.configure(text='✕  Cartella inesistente.',
                                     foreground=COLOR_ERROR)
            return
        paths = self._paths()
        mancanti = [n for n, p in (('misure', paths.misure_dir),
                                   ('scheda_gruppi_dpi.xlsx', paths.scheda_dpi_xlsx))
                    if not os.path.exists(p)]
        if mancanti:
            self._lbl_hint.configure(text='⚠  Non trovati: ' + ', '.join(mancanti),
                                     foreground=COLOR_WARNING)
        else:
            self._lbl_hint.configure(text='✓  Cartella valida.',
                                     foreground=COLOR_SUCCESS)

    # ── Tab Scheda Gruppi DPI ──────────────────────────────────────────────────
    def _build_tab_scheda(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill='x', padx=10, pady=(10, 4))
        ttk.Label(bar, text='scheda_gruppi_dpi.xlsx', style='Header.TLabel').pack(side='left')
        ttk.Button(bar, text='↻ Ricarica', command=self._load_scheda).pack(side='right')
        self._var_sheet = tk.StringVar(value=SCHEDA_MANSIONI)
        combo = ttk.Combobox(bar, textvariable=self._var_sheet, state='readonly',
                             width=20, values=[SCHEDA_MANSIONI, SCHEDA_DPI],
                             font=FONT_SMALL)
        combo.pack(side='right', padx=8)
        combo.bind('<<ComboboxSelected>>', lambda e: self._load_scheda())

        self._view_scheda = DataFrameView(parent)
        self._view_scheda.pack(fill='both', expand=True, padx=10, pady=(4, 10))
        self._view_scheda.show_message(
            'Seleziona una cartella di lavoro, poi premi Ricarica.')

    def _load_scheda(self):
        import pandas as pd
        path = self._paths().scheda_dpi_xlsx
        if not self._var_main.get().strip():
            self._view_scheda.show_message('Nessuna cartella di lavoro selezionata.')
            return
        if not os.path.exists(path):
            self._view_scheda.show_message(f'File non trovato:\n{path}')
            return
        try:
            # header=1: le intestazioni reali stanno sulla seconda riga
            df = pd.read_excel(path, sheet_name=self._var_sheet.get(), header=1)
            self._view_scheda.set_dataframe(df)
        except Exception as e:
            self._view_scheda.show_message(f'Impossibile leggere il foglio:\n{e}')

    # ── Tab Riepilogo ──────────────────────────────────────────────────────────
    def _build_tab_riepilogo(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill='x', padx=10, pady=(10, 4))
        ttk.Label(bar, text=f'{NOME_VR8h_riepilogo} — foglio Riepilogo',
                  style='Header.TLabel').pack(side='left')
        ttk.Button(bar, text='↻ Ricarica', command=self._load_riepilogo).pack(side='right')

        self._view_riep = DataFrameView(parent)
        self._view_riep.pack(fill='both', expand=True, padx=10, pady=(4, 4))
        self._view_riep.show_message(
            'Esegui l\'analisi 8h, oppure premi Ricarica se i risultati esistono già.')

        legend = ttk.Frame(parent)
        legend.pack(fill='x', padx=10, pady=(0, 10))
        for classe, colore in (('BASSA', '#E8F5E9'), ('MEDIA', '#FFF8E1'), ('ALTA', '#FFEBEE')):
            box = ttk.Frame(legend)
            box.pack(side='left', padx=(0, 14))
            tk.Label(box, text='   ', bg=colore, relief='solid', bd=1).pack(side='left')
            ttk.Label(box, text=f' Classe {classe}', style='Muted.TLabel').pack(side='left')

    def _load_riepilogo(self):
        import pandas as pd
        if not self._var_main.get().strip():
            self._view_riep.show_message('Nessuna cartella di lavoro selezionata.')
            return
        paths = self._paths()
        try:
            if os.path.exists(paths.vr8h_riepilogo):
                df = pd.read_excel(paths.vr8h_riepilogo, sheet_name='Riepilogo')
            elif os.path.exists(paths.vr8h_aggiornato):
                # Il riepilogo e' il primo foglio del file aggiornato
                df = pd.read_excel(paths.vr8h_aggiornato, sheet_name=0)
            else:
                self._view_riep.show_message(
                    f'Nessun risultato in:\n{paths.risultati_dir}\n\n'
                    'Esegui l\'analisi 8h.')
                return
            self._view_riep.set_dataframe(df)
        except Exception as e:
            self._view_riep.show_message(f'Impossibile leggere il riepilogo:\n{e}')

    # ── Barra di stato ─────────────────────────────────────────────────────────
    def _build_statusbar(self):
        bar = tk.Frame(self.root, bg='#ECEFF1')
        bar.pack(fill='x', side='bottom')
        self._lbl_status = tk.Label(bar, text='Pronto.', font=FONT_SMALL,
                                    bg='#ECEFF1', fg=COLOR_LABEL, anchor='w')
        self._lbl_status.pack(side='left', padx=12, pady=6)
        self._progress = ttk.Progressbar(bar, mode='indeterminate', length=160)
        self._progress.pack(side='right', padx=12, pady=6)

    def _status(self, text, color=COLOR_LABEL):
        self._ui.post(lambda: self._lbl_status.configure(text=text, fg=color))

    # ── Esecuzione ─────────────────────────────────────────────────────────────
    def _set_running(self, running):
        state = 'disabled' if running else 'normal'
        for b in self._run_buttons:
            b.configure(state=state)
        if running:
            self._progress.start(12)
        else:
            self._progress.stop()
            self._hide_pause()

    def _start(self, target, args=()):
        if self._worker and self._worker.is_alive():
            messagebox.showwarning('Attenzione', 'Un\'elaborazione è già in corso.')
            return False
        if not self._valid_main_directory():
            return False
        self._abort = False
        self._resume.clear()
        self._set_running(True)
        self._notebook.select(self._tab_log)
        self._worker = threading.Thread(target=target, args=args, daemon=True)
        self._worker.start()
        return True

    def _finish(self, text, color):
        self._ui.post(lambda: self._set_running(False))
        self._status(text, color)

    def _report_error(self, e):
        print(f'\nErrore: {e}\n')
        print(traceback.format_exc())
        self._ui.post(lambda: messagebox.showerror('Errore di esecuzione', str(e)))

    # Modalita' automatica
    def _run_auto(self):
        paths = self._paths()
        self._start(self._work_pipeline, (paths, False))

    # Modalita' manuale
    def _run_manual(self):
        paths = self._paths()
        self._start(self._work_pipeline, (paths, True))

    def _work_pipeline(self, paths, manual):
        try:
            pl.run_pipeline(paths, on_step=lambda *a: self._on_step(manual, *a))
            self._reload_views()
            self._finish('Pipeline completata.', COLOR_SUCCESS)
        except pl.PipelineAbort:
            print('\nEsecuzione interrotta dall\'utente.\n')
            self._finish('Interrotto.', COLOR_WARNING)
        except Exception as e:
            self._report_error(e)
            self._finish(f'Errore: {e}', COLOR_ERROR)

    def _on_step(self, manual, chiave, etichetta, indice, totale):
        """Callback chiamato dal worker prima di ogni step."""
        if self._abort:
            raise pl.PipelineAbort()
        print(f'\n▶ [{indice + 1}/{totale}] {etichetta}\n')
        self._status(f'[{indice + 1}/{totale}] {etichetta}', COLOR_ACCENT)

        if not manual:
            return

        msg = f'In pausa prima di: {etichetta}'
        if chiave == 'controllo_ti_grom':
            msg = f'{etichetta}\n{pl.MSG_CONTROLLO_TI_GROM}'
        self._resume.clear()
        self._ui.post(lambda: self._show_pause(msg))
        self._resume.wait()
        if self._abort:
            raise pl.PipelineAbort()
        self._ui.post(lambda: self._hide_pause(resume_progress=True))

    def _show_pause(self, msg):
        self._paused = True
        self._lbl_pause.configure(text=msg)
        self._frm_pause.pack(fill='x', pady=(10, 0))
        self._progress.stop()

    def _hide_pause(self, resume_progress=False):
        if self._paused:
            self._paused = False
            self._frm_pause.pack_forget()
        if resume_progress:
            self._progress.start(12)

    def _continue(self):
        self._resume.set()

    def _stop(self):
        self._abort = True
        self._resume.set()

    # Step singoli
    def _run_single(self, func, label, reload_riepilogo):
        paths = self._paths()
        self._start(self._work_single, (func, label, paths, reload_riepilogo))

    def _work_single(self, func, label, paths, reload_riepilogo):
        try:
            print(f'\n▶ {label}\n')
            self._status(label, COLOR_ACCENT)
            func(paths)
            if reload_riepilogo:
                self._reload_views()
            self._finish(f'{label}: completato.', COLOR_SUCCESS)
        except Exception as e:
            self._report_error(e)
            self._finish(f'Errore: {e}', COLOR_ERROR)

    def _reload_views(self):
        self._ui.post(self._load_riepilogo)
        self._ui.post(self._load_scheda)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
