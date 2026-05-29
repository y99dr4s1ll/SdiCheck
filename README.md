# Controllo Fatture COESA

Confronta le fatture emesse (gestionale) con il file SDI e segnala le discrepanze.

## Requisiti

- Python 3.x
- Librerie: `pip install pandas openpyxl`

(`tkinter` è già incluso in Python.)

## Avvio

Doppio clic sul file, oppure da terminale:

```
python controllo_fatture.py
```

Si apre una finestra. Non serve usare la riga di comando per i parametri.

## Uso

1. **FattureEmesse (.xlsx)** — il file del gestionale.
2. **Clienti SDI (.csv)** — l'export SDI.
3. **Output (.xlsx)** — dove salvare il risultato.
4. **Data filtro** — vengono prese solo le fatture con data *successiva* a questa (es. `2026-03-31` → solo aprile in poi).
5. **Codice Sz.** — uno o più valori separati da virgola (es. `2` oppure `0,1`).
6. Premi **Avvia elaborazione**. A fine lavoro, **Apri output** apre il file Excel.

## Cosa fa

- Da FattureEmesse esclude le righe con `Cod. documento` = `C-ORDINE-PA`.
- Filtra per `Sz.` e per data, costruisce il numero fattura nel formato `'000/00'`.
- Le note di credito (`C-NOTACR-PA` nel gestionale, `Tipo documento` = "Nota di credito" nel CSV) vengono conteggiate con segno negativo.
- Incrocia i due file sul numero fattura (full outer join).
- Se nel CSV lo stesso numero fattura compare più volte (es. rifiutata + riemessa), tiene una sola riga, preferendo quella accettata.

## Output Excel

**Foglio "Controllo"** — dettaglio fattura per fattura:
- Verde = presente in entrambi i file.
- Arancio = presente solo in uno dei due.
- Colonne affiancate: imponibile/IVA da FattureEmesse, imponibile/IVA da CSV, `Delta` (differenza imponibile, arrotondata a 0,001), stato di consegna SDI.

**Foglio "Riepilogo"** — totali imponibile/IVA dei due file e relativi delta.

## Ogni mese

Cambi solo i percorsi dei file e i parametri (data filtro, Sz.) nella finestra. Il codice non va modificato.