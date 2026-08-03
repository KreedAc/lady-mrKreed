# Piano alimentare — Giovanni & Rosalia

Sito statico generato automaticamente da `data/piano_alimentare.xlsx`.
Due sezioni separate (Giovanni e Rosalia), ogni pasto del piano è cliccabile e
porta alla ricetta corrispondente, più ricettario cercabile, lista della spesa
con le spunte e le ricette da provare.

## Come si aggiorna

1. Sostituisci il file **`data/piano_alimentare.xlsx`** con la versione nuova (stesso nome).
2. Fai commit e push.
3. In circa un minuto GitHub Actions rigenera i dati e ripubblica il sito.

Non serve toccare altro: i giorni, i pasti, le ricette e la spesa vengono riletti
dal foglio ogni volta.

### Prima volta: attivare GitHub Pages

Su GitHub → **Settings → Pages → Build and deployment → Source: GitHub Actions**.
Da quel momento il workflow `.github/workflows/deploy.yml` pubblica da solo a
ogni push.

## Come funziona

```
data/piano_alimentare.xlsx     il file sorgente, l'unica cosa da aggiornare
scripts/build_data.py          legge l'xlsx e scrive site/data/plan.json
site/                          il sito (HTML + CSS + JS, nessuna dipendenza)
.github/workflows/deploy.yml   rigenera e pubblica a ogni push
```

Per rigenerare i dati in locale:

```bash
pip install -r requirements.txt
python scripts/build_data.py
```

Per vedere il sito in locale (serve un server, il browser non legge il JSON da `file://`):

```bash
python -m http.server 8000 --directory site
# poi apri http://localhost:8000
```

Per una pagina unica da tenere offline o mandare via chat (dati inclusi dentro l'HTML):

```bash
python scripts/build_data.py --standalone dist/piano.html
```

## Cosa si aspetta lo script dal foglio Excel

I fogli vengono cercati per parola chiave nel nome, quindi piccole variazioni non
danno problemi. La struttura attesa è quella attuale:

| Foglio | Struttura |
|---|---|
| `Piano Settimanale` | Giorno in maiuscolo in colonna A, poi righe `Pasto / Alimento / Quantità / Kcal 100g / Kcal`. Il piatto del pasto è marcato con `▸`. |
| `Giovanni - Ricette` | `Piatto · Quando · Preparazione` |
| `Rosalia - Piano 21gg` | `SETTIMANA n`, giorno, poi righe `Pasto · Piatto · Kcal` |
| `Rosalia - Ricette` | Categoria in maiuscolo, poi `Piatto · Kcal · Ingredienti` |
| `Ricette da provare` | Come sopra |
| `Lista della Spesa` | Titolo gruppo in maiuscolo, categoria, poi `Voce · Quantità` |

Note utili:

- **Il collegamento pasto → ricetta si basa sul nome del piatto.** Se un piatto nel
  piano si chiama esattamente come nel foglio ricette, il link si crea da solo.
  Se non trova la ricetta, il sito mostra il piatto senza link e lo script lo
  segnala in output (`ricetta mancante per: ...`).
- **Puoi aggiungere una colonna `Preparazione`** (colonna D) al foglio
  `Rosalia - Ricette` o a `Ricette da provare`: viene pubblicata in automatico.
- Aggiungere giorni, pasti, ricette o voci della spesa non richiede modifiche al
  codice. Rinominare i fogli o cambiare l'ordine delle colonne sì.

## Contenuto attuale

- Giovanni: 7 giorni, 28 pasti, 14 ricette, tutti i piatti collegati
- Rosalia: 21 giorni (3 settimane), 105 pasti, 49 ricette, tutti i piatti collegati
- 6 ricette da provare · 4 liste della spesa
