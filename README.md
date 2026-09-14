# Scrapers

Raccolta di scraper Python per l'estrazione di dati da diverse directory/siti web (aziende, contatti, email, ecc.), ciascuno organizzato nella propria sottocartella.

## Struttura del repository

```
scrapers/
└── cylex-italia/      # Scraper per Cylex Italia (aziende, contatti, email)
```

Ogni sottocartella contiene il proprio codice, `requirements.txt` e `README.md` con le istruzioni specifiche.

## Progetti disponibili

### [cylex-italia](./cylex-italia)
Scraper con interfaccia grafica (Tkinter) per Cylex Italia: ricerca aziende per regione/provincia/città, estrae anche l'email dalla pagina di dettaglio (gestendo l'offuscamento Cloudflare Email Protection) ed esporta tutto in Excel.

## Requisiti generali
- Python 3.9+
- Le dipendenze specifiche sono elencate nel `requirements.txt` di ciascun progetto.

## Licenza
MIT (vedi [LICENSE](./LICENSE)).
