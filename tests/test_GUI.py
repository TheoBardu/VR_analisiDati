import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import subprocess
import threading
import sys
import os

class RedirectedStdout:
    """Classe per reindirizzare stdout/stderr al widget Text"""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)  # autoscroll

    def flush(self):
        pass  # richiesto per compatibilità

def run_script(main_directory, output_widget):
    """Esegui il tuo codice Python"""
    try:
        # QUI va importato o eseguito il tuo codice, usando main_directory
        # esempio: una funzione principale che prende main_directory
        import main  # supponiamo tu abbia il tuo codice in 'tuo_script.py'
        main.main(main_directory)  # chiamata alla tua funzione principale
    except Exception as e:
        print(f"Errore: {e}")

def choose_directory(entry_widget):
    folder = filedialog.askdirectory()
    if folder:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, folder)

def start_process(entry_widget, output_widget):
    main_directory = entry_widget.get().strip()
    if not main_directory or not os.path.isdir(main_directory):
        messagebox.showerror("Errore", "Seleziona una cartella valida")
        return
    
    # pulisco output
    output_widget.delete(1.0, tk.END)

    # avvio in thread separato per non bloccare la GUI
    thread = threading.Thread(target=run_script, args=(main_directory, output_widget))
    thread.start()

def main():
    root = tk.Tk()
    root.title("Analisi Valutazione Rischio")
    root.geometry("700x500")

    # Selettore cartella
    frame = tk.Frame(root)
    frame.pack(pady=10, fill="x")

    entry_dir = tk.Entry(frame, width=60)
    entry_dir.pack(side="left", padx=5)

    btn_browse = tk.Button(frame, text="Sfoglia", command=lambda: choose_directory(entry_dir))
    btn_browse.pack(side="left", padx=5)

    btn_start = tk.Button(root, text="Avvia Script", command=lambda: start_process(entry_dir, txt_output))
    btn_start.pack(pady=5)

    # Area di output
    txt_output = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20)
    txt_output.pack(padx=10, pady=10, fill="both", expand=True)

    # Reindirizza stdout e stderr al widget
    sys.stdout = RedirectedStdout(txt_output)
    sys.stderr = RedirectedStdout(txt_output)

    root.mainloop()

if __name__ == "__main__":
    main()
