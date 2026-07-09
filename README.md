# Posta Massiva - Electron

Applicazione desktop Electron per:

- leggere un file Excel di invii;
- classificare/riordinare i record per `BACINO DESTINAZIONE` e `DESTINAZIONE TARIFFARIA` usando il manuale in `docs/postamassiva-elenco-bacini-destinazione.pdf`;
- generare un file dedicato alla realizzazione dei **chiudi scatola**.

## Requisiti

- Node.js 20+
- Cartella `docs` presente alla radice progetto con i PDF manuale.

## Avvio

```bash
npm install
npm start
```

## Uso

1. Seleziona il file Excel input.
2. Scegli foglio e colonna CAP.
3. Se disponibile, seleziona la colonna **Numero Elementi (1-7)** per calcolare in automatico peso e spessore busta.
4. Imposta **altezza utile scatola** e **coefficiente reazione**: il limite per scatola viene calcolato con `floor((altezza_mm * coefficiente) / spessore_busta_mm)`.
5. Imposta il **limite massimo invii per scatola** (cap superiore).
6. Imposta i parametri etichetta/chiudi scatola (azienda, centro, codice spedizione, ecc.).
7. Premi `Elabora file`.

Output generati:

- `*_record_ordinati_YYYYMMDD_HHMMSS.xlsx`
- `*_chiudi_scatola_YYYYMMDD_HHMMSS.xlsx`

## Fogli output

### Record ordinati

Contiene i record originali con colonne aggiuntive:

- `CAP_NORMALIZZATO`
- `BACINO_DESTINAZIONE`
- `DESTINAZIONE_TARIFFARIA`
- `PROVINCIA_BACINO`
- `STATO_ELABORAZIONE`

Se ci sono CAP non classificati viene aggiunto il foglio `Scarti_CAP`.

### Chiudi scatola

Righe pronte per etichette scatola con campi coerenti con il PDF `postamassiva-etichette-scatole-pallet.pdf`.

Se presenti, gli scarti CAP vengono comunque inseriti nel file `Chiudi_Scatola` in fondo al flusso, con:

- `Tipo chiusura = CHIUDI PLICO MIX BACINI DESTINAZIONI VARIE`
- `BACINO DESTINAZIONE = MIX BACINI DESTINAZIONI VARIE`
- `Destinazione tariffaria = MIX`

## Note implementative

- Le regole CAP sono lette/parzialmente parse dal PDF bacini.
- Per i CAP AM di Roma/Milano indicati come `VEDI NOTA`, il tool applica una classificazione AM su intervalli municipali (`001xx` e `201xx`) per supportare l'ordinamento e la produzione chiudi scatola.
