import pandas as pd
import pyautogui
import subprocess
import time
import os
import configparser
import tkinter as tk
import sys
from tkinter import filedialog
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QPushButton, QTextEdit
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import Qt, QSize
import re




class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')   
        self.setWindowTitle("Lavorazione Lamiere Lazio - Multitool")
        self.setGeometry(100, 100, 800, 600)  # x, y, width, height

        # Layout principale
        main_layout = QVBoxLayout()

        # Widget GIF
        self.gif_label = QLabel(self)
        self.movie = QMovie("AI_Machine.gif")
        size = self.movie.scaledSize()
        self.setGeometry(100, 100, size.width(), size.height())
        
        self.gif_label.setMovie(self.movie)
        self.movie.start()  # Avvia l'animazione della GIF
        main_layout.addWidget(self.gif_label, 40)  # 40% della GUI

        # Widget Console Log
        self.console_log = QTextEdit(self)
        self.console_log.setReadOnly(True)
        main_layout.addWidget(self.console_log, 20)  # 20% della GUI

        # Widget Pulsante Punzonatrice
        self.button_widget = QWidget(self)
        button_layout = QVBoxLayout(self.button_widget)
        self.punzonatrice_button = QPushButton("Punzonatrice", self)
        self.punzonatrice_button.clicked.connect(self.avvia_script)
        self.punzonatrice_button.setFixedSize(200, 200)  # Imposta una dimensione fissa per il pulsante
        button_layout.addWidget(self.punzonatrice_button)
        main_layout.addWidget(self.button_widget, 1)  # 40% della GUI

        # Imposta il layout principale
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    def avvia_script(self):
       
        main(self,self.config)

def valida_excel(df):
    funzione_pattern = re.compile(r'^f[1-9][0-2]?$')  # Corrisponde a F1, F2, ..., F12
    numero_pattern = re.compile(r'^\d+$')  # Corrisponde a stringhe di soli numeri
    comandi_ammessi = ['su', 'giu', 'destra', 'sinistra', 'invio', 'del']

    for index, riga in df.iterrows():
        for i, cella in enumerate(riga):
            cella_str = str(cella).strip().lower()

            # Salta la validazione per celle vuote o con un solo carattere
            if pd.isna(cella) or len(cella_str) == 1:
                continue

            # Esegue la validazione per celle con più di un carattere
            if numero_pattern.match(cella_str) or cella_str in comandi_ammessi or funzione_pattern.match(cella_str):
                continue  # Accetta numeri o comandi ammessi

            # Se la cella non corrisponde a nessuna delle condizioni valide
            raise ValueError(f"Errore a riga {index+1}, colonna {i+1}. '{cella_str}' non è un valore valido.")

    return True  # Se tutto è valido






def scegli_file_excel():
    root = tk.Tk()
    root.withdraw()  # Nascondi la finestra principale di Tkinter
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls*")])
    if file_path == "":  # Nessun file selezionato
        return None
    return file_path


def main(self,config):
    # Path dell'applicazione

    app_directory = config['Settings']['app_directory']
    excel_sheet_name = config['Settings']['excel_sheet_name']
    delay = float(config['Settings']['delay'])
    app_executable = "EUROWIN.EXE"
    file_excel = scegli_file_excel()
    if file_excel is None:
        self.console_log.append("Nessun file Excel selezionato.")
      
        return  # Termina la funzione main se non viene selezionato alcun file

    # Cambia la directory di lavoro
  
    os.chdir(app_directory)

    # Caricare i dati da Excel (NO HEADER)
    df = pd.read_excel(file_excel,sheet_name=excel_sheet_name, header=None, dtype=str)
    self.console_log.append("Caricamento Excel OK")
    # Utilizza questa funzione per validare il DataFrame prima di procedere con l'esecuzione dello script
    try:
        valida_excel(df)
    except ValueError as e:
        self.console_log.append(f"Validazione fallita: {e}")
        print(f"Validazione fallita: {e}")
        sys.exit(1)  # Chiude il programma con un codice di errore    
    # Avvia l'applicazione
    self.console_log.append("Validazione Excel OK")
    subprocess.Popen(app_executable)
    time.sleep(2)  # Attendere 5 secondi per l'avvio dell'applicazione
    self.console_log.append("Script in avvio")
  
    
    esegui_procedura(df,delay,config)
    self.console_log.append("Programma terminato!")

def esegui_procedura(df,delay,config):
    env_setting = config['Settings']['env']
    for index, riga in df.iterrows():
        
        # Controlla se la riga è vuota (processo completato)
        if riga.isnull().all():
            print("null: ")
            break

        for cella in riga:
            # Se la cella è vuota, passa alla cella successiva
           
            if pd.isna(cella):
                continue

            cella_str = str(cella).strip().lower()
            print("cella: " +cella_str)
            # Gestione dei comandi speciali
            if cella_str == "del":
                pyautogui.press('delete')
            elif cella_str == "invio":
                pyautogui.press('enter')
            elif cella_str == "destra":
                pyautogui.press('right')
            elif cella_str == "sinistra":
                pyautogui.press('left')
            elif cella_str == "giu":
                pyautogui.press('down')
                print("giu")
            elif cella_str == "#":
                if env_setting == "omar":
                    # Comportamento specifico per "omar"
                    print("Comportamento per 'omar'")
                    # Placeholder per la logica specifica di "omar"
                else:
                    # Comportamento alternativo
                    print("Comportamento alternativo")
                    pyautogui.press('enter')
                    pyautogui.press('t')
                    pyautogui.press('left')
                    pyautogui.press('enter')
                    pyautogui.press('enter')
                    pyautogui.press('enter')
                    # Placeholder per la logica alternativa
                # Se trova "#", termina il processo corrente e passa alla riga successiva
                break
            # Gestione dei tasti funzione
            elif len(cella_str) == 2 and cella_str[0].lower() == 'f' and cella_str[1].isdigit():
                pyautogui.press('f' + cella_str[1])
                
            elif len(cella_str) == 1:
                # Se la cella contiene un solo carattere, lo simula
                pyautogui.press(cella_str)
                
            else:
                # Se la cella contiene più di un carattere, la separa e simula ogni tasto individualmenteù
                
                for carattere in cella_str:
                    pyautogui.press(carattere)
                    print("carattere: " +carattere)
                  
                    time.sleep(delay)  # Pausa di 1 secondo tra ogni azione

            time.sleep(delay)


# Avvio dell'applicazione
app = QApplication(sys.argv)
main_window = MainWindow()
main_window.show()
sys.exit(app.exec_())
