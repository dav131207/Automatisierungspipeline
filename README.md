# Automatisierungspipeline (Vollautomatisches Daten-Sourcing & Eintrags-System)

## Business Case
Im öffentlichen Sektor müssen Datensätze immer wieder aus eingescannten PDF-Stammdatenblättern manuell validiert, strukturiert und in starre, behördliche Web-Datenbanken eingepflegt werden. Dieser Prozess ist fehleranfällig, repetitiv und zeitintensiv.

## Die Lösung
Diese Python-Pipeline automatisiert den gesamten Workflow end-to-end:
1. **File-Sourcing:** Scannt Laufwerke nach neuen Dokumenten (Stammdatenblätter/Nachweise) und strukturiert diese in personenspezifischen Ordnern.
2. **Robustes OCR:** Extrahiert mit `PyMuPDF` und `Tesseract-OCR` (inkl. Kontrastoptimierung via `Pillow`) persönliche Daten, SV-Nummern, Adressen und komplexe Checkbox-Muster (Ja/Nein/Keine Angabe) – selbst bei variierenden Formular-Layouts.
3. **Data Quality Gate:** Überprüft die Daten vorab auf logische Widersprüche (z. B. eingetragene SV-Nummer vs. angekreuztes "Keine Angabe") und protokolliert Fehler direkt in einem Excel-Dashboard (`openpyxl`).
4. **Web-Automatisierung (RPA):** Steuert via `Playwright` im Hintergrund den Browser, loggt sich im behördlichen Portal ein, füllt komplexe Formular-Hierarchien aus, generiert die behördliche Kennung und spiegelt die generierte ID in Echtzeit zurück in das Excel-Reporting.

## 💻 Tech-Stack
* **Core:** Python 3
* **Data & Excel:** openpyxl
* **OCR & Image Processing:** PyMuPDF (fitz), pytesseract, Pillow (PIL)
* **Browser-Automatisierung (RPA):** Playwright (Python sync-api)
* **Quality Assurance:** Logging, ConfigParser, automatisches Screenshot-Testing bei UI-Abstürzen

## Impact
* **Zeitersparnis:** Reduktion der Bearbeitungszeit pro Fall von mehreren Minuten auf wenige Sekunden.
* **Fehlerquote:** Nahezu 0 % Übertragungsfehler durch das integrierte *Quality Gate*.
* **Datenschutz:** Sichere, lokale Vorverarbeitung vor der verschlüsselten Web-Übertragung.
