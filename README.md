# Excel Combiner

- Unisce tutti i file .xlsx contenuti nella input folder
- Usa il primo foglio di ogni workbook come sorgente dati
- Valida la coerenza degli header tra i file
- Aggiunge una colonna `source_file` per tenere traccia della provenienza
- Esporta l’output in formato `combined.xlsx` e `combined.csv` (delimiter `;`)

## Come usarlo

- Configura i path in `config.ini`
- Avvia lo script con Python (data reconciliation automatico)

## Output esempio

Trovati 3 file Excel:

- spreadsheet1.xlsx
- spreadsheet2.xlsx
- spreadsheet3.xlsx
  Righe combinate (header incluso): 30
  File generati: data/output/combined.xlsx, data/output/combined.csv

## Per recruiter

- Dimostra automazione Python su dati aziendali reali
- Attenzione alla data quality (validazione header + audit trail)
- Struttura funzionale e configurabile
