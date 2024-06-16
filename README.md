
# Bauzeitenplan Creator / Construction Schedule Creator

## Deutsch

Der Bauzeitenplan Creator ist eine Anwendung zur Erstellung von Bauzeitenplänen. Mit Hilfe einer benutzerfreundlichen grafischen Benutzeroberfläche (GUI) können Benutzer Bauzeitenpläne eingeben, speichern und in Excel öffnen.

### Funktionen

- Eingabe von Bauherreninformationen, Projektname und Projektnummer
- Auswahl des Start- und Endjahres sowie des Start- und Endmonats
- Option zur Berücksichtigung von Wochentagen
- Speichern der Eingaben als Excel-Datei
- Öffnen der gespeicherten Excel-Datei direkt aus der Anwendung
- Anpassbare Darstellung (Light, Dark, System)
- Anpassbare UI-Skalierung

### Installation

1. Klone das Repository:
    ```sh
    git clone https://github.com/BlackHack13/Bauzeitenplan.git
    ```
2. Navigiere in das Projektverzeichnis:
    ```sh
    cd Bauzeitenplan
    ```
3. Installiere die benötigten Abhängigkeiten:
    ```sh
    pip install -r requirements.txt
    ```

### Verwendung

1. Starte die Anwendung:
    ```sh
    python Bauzeitenplan.py
    ```
2. Fülle die Felder Bauherr, Projektname und Projektnummer aus.
3. Wähle das Start- und Endjahr sowie den Start- und Endmonat aus.
4. Optional: Aktivieren der Checkbox für Wochentage.
5. Klicke auf "Speichern", um die Eingaben als Excel-Datei zu speichern.
6. Klicke auf "Öffnen in Excel", um die gespeicherte Datei in Excel zu öffnen.

### Abhängigkeiten

- Python 3.x
- tkinter
- customtkinter
- openpyxl
- Pillow
- requests
- keyboard
- holidays

Diese können über die `requirements.txt` installiert werden.

### Projektstruktur

```
Bauzeitenplan/
│
├── IMG/                   # Bilderverzeichnis
│   ├── icon.ico
│   └── logo.png
│
├── Bauzeitenplan.py       # Hauptskript
├── requirements.txt       # Abhängigkeiten
└── README.md              # Diese Datei
```

### Entwickler

- J. Winkler

### Lizenz

Dieses Projekt steht unter der MIT-Lizenz. Weitere Informationen findest du in der LICENSE-Datei.

### Versionskontrolle

Diese Anwendung überprüft automatisch, ob eine neue Version verfügbar ist, und informiert den Benutzer entsprechend.

### Support

Für Support, Fragen oder Anmerkungen, erstelle bitte ein Issue im [GitHub Repository](https://github.com/BlackHack13/Bauzeitenplan/issues).

---

## English

The Construction Schedule Creator is an application for creating construction schedules. With the help of a user-friendly graphical user interface (GUI), users can enter, save, and open construction schedules in Excel.

### Features

- Input of client information, project name, and project number
- Selection of start and end year as well as start and end month
- Option to consider weekdays
- Save entries as an Excel file
- Open the saved Excel file directly from the application
- Customizable display (Light, Dark, System)
- Customizable UI scaling

### Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/BlackHack13/ConstructionScheduleCreator.git
    ```
2. Navigate to the project directory:
    ```sh
    cd ConstructionScheduleCreator
    ```
3. Install the required dependencies:
    ```sh
    pip install -r requirements.txt
    ```

### Usage

1. Start the application:
    ```sh
    python ConstructionScheduleCreator.py
    ```
2. Fill in the fields for client, project name, and project number.
3. Select the start and end year as well as the start and end month.
4. Optional: Activate the checkbox for weekdays.
5. Click "Save" to save the entries as an Excel file.
6. Click "Open in Excel" to open the saved file in Excel.

### Dependencies

- Python 3.x
- tkinter
- customtkinter
- openpyxl
- Pillow
- requests
- keyboard
- holidays

These can be installed via the `requirements.txt`.

### Project Structure

```
ConstructionScheduleCreator/
│
├── IMG/                   # Image directory
│   ├── icon.ico
│   └── logo.png
│
├── ConstructionScheduleCreator.py  # Main script
├── requirements.txt       # Dependencies
└── README.md              # This file
```

### Developer

- J. Winkler

### License

This project is licensed under the MIT License. For more information, see the LICENSE file.

### Version Control

This application automatically checks for a new version and informs the user accordingly.

### Support

For support, questions, or comments, please create an issue in the [GitHub repository](https://github.com/BlackHack13/ConstructionScheduleCreator/issues).
