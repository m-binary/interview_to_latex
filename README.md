# Interview to LaTeX

Streamlit-App zum Konvertieren einer Excel-Datei (2 Spalten: Teilnehmer, Inhalt) in eine LaTeX-Datei nach dem Vorbild `interview.tex`.

Installation (lokal):

1. Python-Umgebung einrichten (empfohlen: venv)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. App starten:

```bash
streamlit run main.py
```

Benutzung:

- Fülle die Metadaten im Expander aus (Name, Interviewpartner, Funktion, etc.).
- Lade eine Excel-Datei (.xlsx) hoch ohne Kopfzeile; die erste Spalte ist der Sprecher, die zweite der Text.
- Klicke auf "Kompilieren" und lade die generierte `interview.tex` herunter.
