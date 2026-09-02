# Analisi esame Blackboard

Piccola app Streamlit per analizzare gli export Excel di esami/test scaricati da Blackboard.

## Avvio locale

Da questa cartella:

```bash
python3 -m streamlit run app.py --server.port 8501 --server.address localhost
```

Poi apri:

```text
http://localhost:8501
```

## Uso

- Carica dalla sidebar un file Excel `.xlsx` esportato da Blackboard.
- In alternativa, sul Mac locale puoi indicare un percorso file già presente sul disco.
- I fogli che contengono colonne `Question ID n` vengono rilevati automaticamente.
- Puoi scegliere dalla sidebar quali fogli includere, ad esempio `Attending` e/o `Non attending`.
- Se sostituisci un file con una nuova versione e Streamlit mostra ancora dati vecchi, usa `Svuota cache e ricarica file`.
- Le analisi possono essere filtrate per `CLASSE`.

## Cosa fa

- riepilogo con medie, mediane, distribuzione dei voti e analisi per domanda;
- selezione dei fogli esame da analizzare nello stesso file Excel;
- dettaglio del singolo studente;
- evidenziazione di domande errate, parziali o senza voto;
- gestione dei `Manual Score`;
- generazione schede visione compiti in Word;
- esportazione Excel corretta;
- compilazione del file ufficiale esiti `.xls`.

## Privacy e file dati

I file Excel degli esami, i file `.docx` esportati e `data/manual_scores.csv` non devono essere caricati su GitHub perche possono contenere dati degli studenti. Sono esclusi tramite `.gitignore`.

Su Streamlit Community Cloud l'app non puo leggere percorsi locali del Mac o OneDrive: carica i file Excel dalla sidebar dell'app.
