# Gestione Compensi - Studio Serra

Applicazione Desktop in Python (Tkinter + ttkbootstrap) per la gestione dei preventivi e il calcolo dei compensi per Studio Serra.

## Funzionalità principali
- **Calcolo Compensi**: Calcolo automatico in base ad unità abitative, box auto, spese fisse, trasferte, e condizioni speciali.
- **Riepilogo in Tempo Reale**: Calcolo dell'imponibile, cassa previdenza e IVA aggiornati in tempo reale.
- **Generazione Documenti**: Esportazione automatica dei preventivi in formato Word (.docx) e PDF tramite modelli Word personalizzati.
- **Storico e Ricerca**: Archivio locale dei preventivi generati con ricerca rapida e gestione file.
- **Scadenzario Bilanci**: Tracciamento delle date di chiusura dei bilanci condominiali con avvisi cromatici e stampa report.

## Avvio
Per avviare l'applicazione:
1. Eseguire `installa_librerie.bat` per installare le dipendenze necessarie (`ttkbootstrap`, `docxtpl`, `pywin32`, `geopy`, `certifi`).
2. Avviare `Avvia Preventivi.bat` o eseguire `pythonw preventivi.pyw`.
