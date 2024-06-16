
# Bauzeitenplan Creator / Construction Schedule Creator

[![Deutsch](https://img.shields.io/badge/Deutsch-Deutsch-blue)](#deutsch)
[![English](https://img.shields.io/badge/English-English-green)](#english)

---

## Deutsch

Der Bauzeitenplan Creator ist eine Anwendung zur Erstellung von Bauzeitenplänen und Ausgabe als .xlsx Datei. Mit Hilfe einer benutzerfreundlichen grafischen Benutzeroberfläche können Benutzer einfach Bauzeitenpläne eingeben, speichern und in Excel öffnen.

[Zurück](#bauzeitenplan-creator--construction-schedule-creator)

### Funktionen

- Eingabe von Bauherreninformationen, Projektname und Projektnummer
- Auswahl des Start- und Endjahres sowie des Start- und Endmonats
- Option zur Berücksichtigung von Wochentagen
- Speichern der Eingaben als Excel-Datei
- Öffnen der gespeicherten Excel-Datei direkt aus der Anwendung
- Anpassbare Darstellung (Light, Dark, System)
- Anpassbare UI-Skalierung

[Zurück](#bauzeitenplan-creator--construction-schedule-creator)

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

[Zurück](#bauzeitenplan-creator--construction-schedule-creator)

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

[Zurück](#bauzeitenplan-creator--construction-schedule-creator)

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

[Zurück](#bauzeitenplan-creator--construction-schedule-creator)

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

[Zurück](#bauzeitenplan-creator--construction-schedule-creator)

### Entwickler

- J. Winkler

[Zurück](#bauzeitenplan-creator--construction-schedule-creator)

### Lizenz

Dieses Projekt ist unter der MIT-Lizenz lizenziert. Weitere Informationen finden Sie in der LICENSE-Datei.

[Zurück](#bauzeitenplan-creator--construction-schedule-creator)

### Versionskontrolle

Diese Anwendung überprüft automatisch auf eine neue Version und informiert den Benutzer entsprechend.

[Zurück](#bauzeitenplan-creator--construction-schedule-creator)

### Support

Für Support, Fragen oder Kommentare erstellen Sie bitte ein Issue im [GitHub-Repository](https://github.com/BlackHack13/Bauzeitenplan/issues).

[Zurück](#bauzeitenplan-creator--construction-schedule-creator)

## English

The Construction Schedule Creator is an application for creating construction schedules and exporting them as .xlsx files. With a user-friendly graphical interface, users can easily input, save, and open construction schedules in Excel.

[Back to top](#bauzeitenplan-creator--construction-schedule-creator)

### Features

- Input of client information, project name, and project number
- Selection of start and end year as well as start and end month
- Option to consider weekdays
- Save entries as an Excel file
- Open the saved Excel file directly from the application
- Customizable display (Light, Dark, System)
- Customizable UI scaling

[Back to top](#bauzeitenplan-creator--construction-schedule-creator)

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

[Back to top](#bauzeitenplan-creator--construction-schedule-creator)

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

[Back to top](#bauzeitenplan-creator--construction-schedule-creator)

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

[Back to top](#bauzeitenplan-creator--construction-schedule-creator)

### Project Structure

```
ConstructionScheduleCreator/
│
├── IMG/                   # Image directory
│   ├── icon.ico
│   └── logo.png
│
├── Bauzeitenplan.py       # Main script
├── requirements.txt       # Dependencies
└── README.md              # This file
```

[Back to top](#bauzeitenplan-creator--construction-schedule-creator)

### Developer

- J. Winkler

[Back to top](#bauzeitenplan-creator--construction-schedule-creator)

### License

This project is licensed under the MIT License. For more information, see the LICENSE file.

[Back to top](#bauzeitenplan-creator--construction-schedule-creator)

### Version Control

This application automatically checks for a new version and informs the user accordingly.

[Back to top](#bauzeitenplan-creator--construction-schedule-creator)

### Support

For support, questions, or comments, please create an issue in the [GitHub repository](https://github.com/BlackHack13/ConstructionScheduleCreator/issues).

[Back to top](#bauzeitenplan-creator--construction-schedule-creator)
