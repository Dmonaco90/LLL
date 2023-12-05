import pandas as pd
import pyautogui
import subprocess
import time
import os
import configparser
import tkinter as tk
import sys
from tkinter import filedialog
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QPushButton, QTextEdit,QHBoxLayout, QSizePolicy,QGridLayout,QSpacerItem,QFileDialog,QComboBox
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtGui import QMovie,QIcon
from PyQt5.QtCore import Qt, QSize,QUrl,QThread, pyqtSignal, QTimer,QTime
import re
import threading

# Definisci una classe per il thread del timer
class TimerThread(QThread):
    # Definisci un segnale che verrà emesso quando il timer viene aggiornato
    timer_updated = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.running = True

    def run(self):
        seconds = 0
        while self.running:
            time_string = f"{seconds // 3600:02d}:{(seconds // 60) % 60:02d}:{seconds % 60:02d}"
            self.timer_updated.emit(time_string)  # Invia solo il tempo
            seconds += 1
            self.sleep(1)

    def stop(self):
        # Aggiungi un metodo 'stop' personalizzato per terminare il thread in modo sicuro
        self.running = False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')
        self.delay = float(self.config['Settings']['delay'])
        # ...
        
            # Imposta la finestra per aprirsi massimizzata
         # Crea un'istanza del thread del timer
        self.timer_thread = TimerThread()
        # Connetti il segnale del thread al metodo per aggiornare il timer label
        self.timer_thread.timer_updated.connect(self.update_timer_label)
        # Avvia il thread del timer
        #self.timer_thread.start()
        self.setWindowTitle("Lavorazione Lamiere Lazio")
        self.setGeometry(100, 100, 800, 600)  # x, y, width, height
        self.logo_label = QLabel(self)
        self.setWindowIcon(QIcon('logoLLL.png')) 
        self.execution_time = QTime(0, 0)
       
         # Creazione della QComboBox per le velocità
        self.speed_combo = QComboBox(self)
        self.speed_combo.addItem("0")
        self.speed_combo.addItem("0.5")
        self.speed_combo.addItem("1")
        self.speed_combo.currentIndexChanged.connect(self.update_delay)
        self.update_delay()  # Chiama la funzione per inizializzare il valore di delay  
       
        # Layout principale
        
       

        main_layout = QVBoxLayout()
        
        
        
        # Crea un QLabel per la firma
        self.firma_label = QLabel("Creato da: Daniele Monaco", self)
        
        # Stile opzionale per il QLabel
        self.firma_label.setStyleSheet("color: gray; font-style: italic; font-size: 10px;")
        
        # Posiziona il QLabel in basso a destra (o dove preferisci)
        self.firma_label.setAlignment(Qt.AlignBottom | Qt.AlignRight)
        self.titolo_label = QLabel("Artificial Metal Automation", self)
        self.titolo_label.setStyleSheet("background-color: black; color: white; font-size: 40px;")
        self.titolo_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.titolo_label)
        
      

      
# Layout superiore per GIF e widget nero
        top_layout = QHBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)  # Rimuovi i margini
        top_layout.setSpacing(0)  # Rimuovi lo spazio tra i widget

        # Widget 1: GIF
        self.gif_label = QLabel(self)
        self.movie = QMovie("AI_Machine.gif")
        self.gif_label.setMovie(self.movie)
        self.movie.start()
        self.gif_label.setScaledContents(True)  # Assicurati che la GIF riempia la QLabel

        # Widget 2: Nuovo widget con sfondo bianco
        self.black_widget = QWidget(self)
        self.black_widget.setStyleSheet("background-color: black; border: none;")
        black_widget_layout = QGridLayout(self.black_widget)
        black_widget_layout.setContentsMargins(0, 0, 0, 0)  # Imposta i margini a zero

        # Crea il QTextEdit per il log e applica lo stile
        self.log_text_edit = QTextEdit(self.black_widget)
        self.log_text_edit.setStyleSheet(
            "QTextEdit {"
            "background-color: white;"  # Imposta lo sfondo dell'intero widget
            "color: black;"  # Imposta il colore del testo
            "border: none;"  # Rimuove i bordi
            "font-size: 14pt;"  # Aumenta la dimensione del font (modifica il valore in base alle tue esigenze)
            "}"
        )
        self.log_text_edit.setReadOnly(True)

        # Aggiungi il QTextEdit al layout del black_widget
        black_widget_layout.addWidget(self.log_text_edit, 0, 0)

        # Crea uno spaziatore per occupare l'altra metà dello spazio
        spacer = QSpacerItem(20, 40, QSizePolicy.Expanding, QSizePolicy.Minimum)
        black_widget_layout.addItem(spacer, 0, 1)

        # Aggiungi la GIF e il widget nero al layout superiore
        top_layout.addWidget(self.gif_label, 0)  
        top_layout.addWidget(self.black_widget, 1)

        # Assicurati che non ci siano margini che separano i widget
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)  # Nessuno spazio tra i widget

        # Aggiungi il layout superiore al layout principale
        main_layout.addLayout(top_layout, 1)  # Assegna il doppio dello spazio ai widget superiori
        
        
         # Timer and Timer QLabel
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_timer)
        self.timer_label = QLabel("Tempo di esecuzione: 00:00:00", self.black_widget)
        self.timer_label.setStyleSheet("color: white; font-size: 40px;")
        self.timer_label.setAlignment(Qt.AlignCenter)

        # Add the timer QLabel to the black_widget layout
        black_widget_layout.addWidget(self.timer_label, 0, 1, Qt.AlignCenter)

        
       # Creazione del layout per il titolo e il widget combo
        delay_layout = QHBoxLayout()
        # Aggiungi uno spaziatore vuoto a sinistra per spostare il titolo a destra
        delay_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        # Aggiungi un QLabel per il titolo
        delay_label = QLabel("Delay:", self)
        
        delay_label.setStyleSheet("font-size: 16px;")  # Imposta la dimensione del testo
        delay_label.setFixedHeight(20)  # Imposta l'altezza fissa

        # Aggiungi il widget self.speed_combo al layout
        self.speed_combo.setFixedHeight(20)  # Imposta l'altezza fissa per la combo box

        delay_layout.addWidget(delay_label)

        # Aggiungi il widget self.speed_combo al layout
        delay_layout.addWidget(self.speed_combo)

        # Aggiungi il layout delay_layout al layout principale
        main_layout.addLayout(delay_layout)
        
        
        
        main_layout.addWidget(self.firma_label)
      
        main_layout.addWidget(self.speed_combo)

        # Widget 4: Pulsanti
        button_layout = QGridLayout()
        numero_pulsanti = 3
        self.punzonatrice_button = QPushButton("Smart-Punching Machine (Punzonatrice)", self)
        self.punzonatrice_button.clicked.connect(self.avvia_thread_punzonatrice)

        self.secondo_pulsante = QPushButton("Funzione 2", self)
        self.secondo_pulsante.clicked.connect(self.avvia_script_secondo)

        self.terzo_pulsante = QPushButton("Funzione 3", self)
        self.terzo_pulsante.clicked.connect(self.avvia_script_terzo)

        # Calcola la larghezza desiderata per ogni pulsante
        larghezza_pulsante = int(self.width() / numero_pulsanti)
        altezza_pulsante =  int(self.height() / numero_pulsanti)
        self.punzonatrice_button.setFixedHeight(altezza_pulsante)
        self.secondo_pulsante.setFixedHeight(altezza_pulsante)
        self.terzo_pulsante.setFixedHeight(altezza_pulsante)
        # Imposta la larghezza e aggiungi i pulsanti al layout a griglia
        self.punzonatrice_button.setFixedWidth(larghezza_pulsante)
        button_layout.addWidget(self.punzonatrice_button, 0, 0)

        self.secondo_pulsante.setFixedWidth(larghezza_pulsante)
        button_layout.addWidget(self.secondo_pulsante, 0, 1)

        self.terzo_pulsante.setFixedWidth(larghezza_pulsante)
        button_layout.addWidget(self.terzo_pulsante, 0, 2)

        # Aggiungi il layout dei pulsanti al layout principale
        main_layout.addLayout(button_layout)

       # Imposta il layout principale
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        
        
        
        
             # Imposta il lettore multimediale
        self.player = QMediaPlayer()
        self.player.setMedia(QMediaContent(QUrl.fromLocalFile("music.mp3")))
        self.player.setVolume(50)  # Imposta il volume iniziale

          # Pulsante per attivare/disattivare l'audio
        self.toggle_audio_button = QPushButton(self)
        self.toggle_audio_button.clicked.connect(self.toggle_audio)
        self.toggle_audio_button.setFixedSize(50, 50)  # Dimensione del pulsante
        self.toggle_audio_button.setStyleSheet("QPushButton {border-radius: 25;}")

        # Imposta il lettore multimediale
        self.player = QMediaPlayer()
        self.player.setMedia(QMediaContent(QUrl.fromLocalFile("music.mp3")))
        self.player.setVolume(50)  # Imposta il volume iniziale

        # Inizia a suonare la musica all'avvio
        

        # Aggiorna l'icona del pulsante per riflettere lo stato iniziale del lettore
        self.update_button_icon()
    
        # Aggiungi il pulsante al layout principale
        main_layout.addWidget(self.toggle_audio_button)
        
        
    def update_delay(self):
        selected_speed = self.speed_combo.currentText()
        if selected_speed == "0":
            self.delay = 0
        elif selected_speed == "0.5":
            self.delay = 0.5
        elif selected_speed == "1":
            self.delay = 1

    
    def update_timer(self):
        # Questo metodo verrà chiamato ogni volta che scade il timer
        # Aggiorna il timer label con il tempo di esecuzione
        current_time = QTime.currentTime()
        time_string = current_time.toString("hh:mm:ss")
        self.timer_label.setText(f"Tempo di esecuzione: {time_string}")

  
        
    def update_timer_label(self, time_string):
        # Converti la stringa del tempo in un oggetto QTime
        time_data = time_string.split(":")
        hours = int(time_data[0])
        minutes = int(time_data[1])
        seconds = int(time_data[2])
        new_time = QTime(hours, minutes, seconds)
        
        # Calcola il tempo trascorso dall'ultima emissione del segnale
        elapsed_time = self.execution_time.msecsTo(new_time)
        
        # Aggiorna l'etichetta con il tempo trascorso
        self.timer_label.setText(f"Tempo di esecuzione: {time_string}")
        
        # Aggiorna il tempo di esecuzione memorizzato
        self.execution_time = new_time

    def start_timer(self):
        self.timer_thread.start()
        self.player.play()
    def stop_timer(self):
        self.timer_thread.stop()
        self.player.pause()
       
    def toggle_audio(self):
        if self.player.state() == QMediaPlayer.PlayingState:
            self.player.pause()  # Metti in pausa se sta suonando
            
        else:
            self.player.play()  # Riproduci se è in pausa
        self.update_button_icon()

    def update_button_icon(self):
        icon_path = 'icon_audio_off.png' if self.player.state() == QMediaPlayer.PlayingState else 'icon_audio_on.png'
        
        icon = QIcon(icon_path)
        if icon.isNull():
            print(f"Errore nel caricamento dell'icona: {icon_path}")
            return

        self.toggle_audio_button.setIcon(icon)
        self.toggle_audio_button.setIconSize(QSize(40, 40))
        
    def avvia_thread_punzonatrice(self):
        # Crea un thread per l'esecuzione del programma di punzonatura
        punzonatrice_thread = threading.Thread(target=self.avvia_script_punzonatrice)
        
        # Avvia il thread
        punzonatrice_thread.start()
        
    def avvia_script_punzonatrice(self):
        
        main(self, self.config)
        # Termina il thread del timer quando il programma di punzonatura è terminato
        self.timer_thread.running = False
        
    def avvia_script_secondo(self):
        self.log_text_edit.append("Funzione non presente.")
    def avvia_script_terzo(self):
        self.log_text_edit.append("Funzione non presente.")
     
     
    def esegui_procedura(self,df,config):
        env_setting = config['Settings']['env']
        self.start_timer()
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
                #print("cella: " +cella_str)
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
                    #print("giu")
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
                
                elif len(cella_str) == 3 and cella_str[0].lower() == 'f' and cella_str[1].isdigit() and cella_str[2].isdigit():
                    pyautogui.press('f' + cella_str[1] + cella_str[2])
                
                elif len(cella_str) == 1:
                    # Se la cella contiene un solo carattere, lo simula
                    pyautogui.press(cella_str)
                    
                else:
                    # Se la cella contiene più di un carattere, la separa e simula ogni tasto individualmenteù
                    
                    for carattere in cella_str:
                        pyautogui.press(carattere)
                        #print("carattere: " +carattere)
                    
                        time.sleep(self.delay)  # Pausa di 1 secondo tra ogni azione

                time.sleep(self.delay)
        self.stop_timer()   

def valida_excel(df):
    funzione_pattern = re.compile(r'^f[1-9]\d*$', re.IGNORECASE)  # Corrisponde a F1, F2, ..., F12, F13, ...
    numero_pattern = re.compile(r'^-?\d+(\.\d+)?$')  # Corrisponde a numeri interi e decimali, anche negativi
    comandi_ammessi = ['su', 'giu', 'destra', 'sinistra', 'invio', 'del']
    
    for index, riga in df.iterrows():
        for i, cella in enumerate(riga):
            # Converti la cella in stringa e normalizza
            if ',' in cella  or '.' in cella:
                try:
                    cella = round(float(cella), 1)  # Arrotonda a 1 cifra decimale
                    df.at[index, i] = cella  # Assegna il nuovo valore alla cella nel DataFrame
                except ValueError:
                    pass  # Gestisci il caso in cui il valore non sia un numero valido
            cella_str = str(cella).strip().lower().replace(',', '.')
            
            # Salta le celle vuote
            if pd.isna(cella) or cella_str == '' or len(cella_str) == 1:
                continue

            # Interrompi la ricerca nella riga corrente se trovi "#"
            if cella_str == "#":
                break

            # Controlla se la cella contiene un valore non valido
            if not any([
                numero_pattern.match(cella_str),
                cella_str in comandi_ammessi,
                funzione_pattern.match(cella_str)
            ]):
                errore = f"Errore a riga {index+1}, colonna {i+1}: '{cella}' non è un valore valido."
                raise ValueError(errore)

    return True, "Validazione completata con successo."









def scegli_file_excel():
    options = QFileDialog.Options()
    options |= QFileDialog.ReadOnly  # Opzione per aprire il file solo in lettura
    file_path, _ = QFileDialog.getOpenFileName(None, "Seleziona il file Excel", "", "Excel files (*.xls *.xlsx);;All Files (*)", options=options)
    
    if file_path == "":
        return None
    return file_path



def main(self,config):
    # Path dell'applicazione

    app_directory = config['Settings']['app_directory']
    excel_sheet_name = config['Settings']['excel_sheet_name']
    #delay = float(config['Settings']['delay'])
    app_executable = "EUROWIN.EXE"
    file_excel = scegli_file_excel()
    if file_excel is None:
        self.log_text_edit.append("Nessun file Excel selezionato.")
      
        return  # Termina la funzione main se non viene selezionato alcun file

    # Cambia la directory di lavoro
  
    os.chdir(app_directory)

    # Caricare i dati da Excel (NO HEADER)
    df = pd.read_excel(file_excel,sheet_name=excel_sheet_name, header=None, dtype=str)
    self.log_text_edit.append("Caricamento Excel OK")
    # Utilizza questa funzione per validare il DataFrame prima di procedere con l'esecuzione dello script
    try:
        valido, messaggio = valida_excel(df)
        if not valido:
            self.log_text_edit.append(messaggio)
            return
    except ValueError as e:
        self.log_text_edit.append(f"Validazione fallita: {e}")
        return
    # Avvia l'applicazione
    self.log_text_edit.append("Validazione Excel OK")
    subprocess.Popen(app_executable)
    time.sleep(2)  # Attendere 5 secondi per l'avvio dell'applicazione
    self.log_text_edit.append("Script in avvio")
  
    
    self.esegui_procedura(df,config)
    self.log_text_edit.append("Programma terminato!")
    
def converti_virgola_in_punto(valore):
    if isinstance(valore, str):
        valore = valore.replace(',', '.')
        try:
            valore = round(float(valore), 1)  # Arrotonda a 1 cifra decimale
        except ValueError:
            pass  # Gestisci il caso in cui il valore non sia un numero valido
    return valore




# Avvio dell'applicazione
app = QApplication(sys.argv)
main_window = MainWindow()
# Imposta la finestra per aprirsi massimizzata
main_window.showMaximized()
sys.exit(app.exec_())


