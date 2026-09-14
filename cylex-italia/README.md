# Cylex Scraper

Scraper con interfaccia grafica (Tkinter) per Cylex Italia.

## Funzionalità
- Selezione geografica Regione → Provincia → Città (o intera provincia)
- Scarica tutte le pagine di risultati per ogni città
- Estrae anche l'email dalla pagina di dettaglio (gestisce l'offuscamento Cloudflare Email Protection)
- Esporta i risultati in Excel (.xlsx), ordinati alfabeticamente
- Pulsante Stop per interrompere la ricerca in corso
- Download automatico, al primo avvio, del database dei comuni italiani

## Installazione
```bash
pip install -r requirements.txt
```

## Uso
```bash
python cylex_scraper.py
```

Al primo avvio viene scaricato automaticamente il database dei comuni italiani (`gi_comuni_cap.csv`).

## File generati automaticamente
Durante l'uso, il programma crea alcuni file/cartelle in locale che **non fanno parte del codice sorgente** e sono esclusi tramite `.gitignore`:

- **`gi_comuni_cap.csv`** — database dei comuni italiani, scaricato automaticamente al primo avvio nella cartella dell'eseguibile/script.
- **`Scaricati/`** — cartella di destinazione degli export Excel, raggiungibile dall'interfaccia.

Questi file vengono rigenerati automaticamente ad ogni utilizzo e non devono essere versionati.

## Creare l'eseguibile Windows (.exe)
Il progetto può essere compilato in un unico eseguibile con [PyInstaller](https://pyinstaller.org/):

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=risorse/icona.ico cylex_scraper.py
```

L'eseguibile verrà generato in `dist/cylex_scraper.exe`. Non è incluso nel repository per motivi di dimensione: se ti serve una build già pronta, pubblicala come Release della repo invece di versionarla nel codice sorgente.

## Struttura cartella
```
cylex-italia/
├── cylex_scraper.py     # codice sorgente
├── cylex_scraper.exe    # eseguibile (unsigned e rilevato come falso positivo in AVG Antivirus, ecc..)
├── gi_comuni_cap.csv    # database comuni italiani
├── requirements.txt     # dipendenze Python
├── risorse/
│   └── icona.ico         # icona per l'eseguibile
├── Scaricati/
└── README.md
```
