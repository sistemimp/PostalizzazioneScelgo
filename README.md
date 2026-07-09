# Posta Massiva - Electron

Applicazione desktop Electron per:

- leggere un file Excel di invii;
- classificare/riordinare i record per `BACINO DESTINAZIONE` e `DESTINAZIONE TARIFFARIA` usando il manuale in `docs/postamassiva-elenco-bacini-destinazione.pdf`;
- generare un file dedicato alla realizzazione dei **plichi**;
- generare un file dedicato alla composizione dei **bancali**.

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
3. Imposta il **limite massimo invii per plico**.
4. Imposta i parametri etichetta plichi (azienda, centro, codice spedizione, ecc.).
5. Inserisci il **peso unitario di default (grammi)**: viene applicato a tutti i record.
6. Premi `Elabora file`.

Output generati:

- `*_record_ordinati_YYYYMMDD_HHMMSS.xlsx`
- `*_plichi_YYYYMMDD_HHMMSS.xlsx`
- `*_bancali_YYYYMMDD_HHMMSS.xlsx`

## Fogli output

### Record ordinati

Contiene i record originali con colonne aggiuntive:

- `CAP_NORMALIZZATO`
- `BACINO_DESTINAZIONE`
- `DESTINAZIONE_TARIFFARIA`
- `PROVINCIA_BACINO`
- `STATO_ELABORAZIONE`

Se ci sono CAP non classificati viene aggiunto il foglio `Scarti_CAP`.

### Plichi

Righe pronte per etichette plichi con peso unitario assegnato da interfaccia.

Se presenti, gli scarti CAP vengono comunque inseriti nel file `Plichi` in fondo al flusso, con:

- `Tipo chiusura = CHIUDI PLICO MIX BACINI DESTINAZIONI VARIE`
- `BACINO DESTINAZIONE = MIX BACINI DESTINAZIONI VARIE`
- `Destinazione tariffaria = MIX`

### Bancali

I bancali vengono creati aggregando i plichi secondo lo schema cliente:

- `ANCONA -> ANCONA + PESARO + MACERATA`
- `ANCONA -> ASCOLI PICENO + FERMO`
- `PESCARA -> TERAMO`
- `PESCARA -> PESCARA`
- `PESCARA -> CHIETI + L'AQUILA + CAMPOBASSO + ISERNIA`
- residui sotto soglia in `MIX BACINI DESTINAZIONI VARIE`

Vincoli applicati:

- peso massimo bancale configurabile da interfaccia
- peso minimo bancale configurabile da interfaccia
- i gruppi sotto minimo vengono riversati nel bancale `MIX`

## Note implementative

- Le regole CAP sono lette/parzialmente parse dal PDF bacini.
- Per i CAP AM di Roma/Milano indicati come `VEDI NOTA`, il tool applica una classificazione AM su intervalli municipali (`001xx` e `201xx`) per supportare l'ordinamento e la produzione chiudi scatola.
