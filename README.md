# Settant'anni di popolazione, comune per comune

Dashboard sulla popolazione dei comuni italiani ai censimenti dal 1951 al 2021, letta attraverso la classificazione delle Aree Interne (SNAI 2021–2027).

**Pagina:** https://tinalovat.github.io/dashboard-aree-interne/

Lavoro di prova di Martina Lovat, sviluppato con l'aiuto di [Claude](https://claude.ai) (Anthropic) per la preparazione dei dati e la scrittura del codice. Impostazione, scelte di analisi e verifica dei risultati sono dell'autrice.

## Contenuto

| File | Cosa contiene |
|---|---|
| `index.html` | La dashboard: unico file con struttura, stile e codice |
| `data/data.js` | Dati comunali usati dalla dashboard (generato) |
| `data/geo.js` | Confini comunali e regionali semplificati (generato) |
| `data/comuni_2021.csv` | Dati comunali ai confini 2021, scaricabili |
| `data/corrispondenze_geo.csv` | Per ogni comune 2021, il poligono usato sulle mappe |
| `data/ricostruzione_log.txt` | Registro delle operazioni di ricostruzione |
| `scripts/prepara_dati.py` | Dal foglio di partenza a `data.js` e `comuni_2021.csv` |
| `scripts/prepara_geo.py` | Dai confini openpolis a `geo.js` |

La pagina si può aprire anche in locale con un doppio clic su `index.html` (serve la connessione a internet per caricare le librerie ECharts e topojson-client).

## Fonti

- **Popolazione:** ISTAT, Censimenti della popolazione, tramite [Esploradati](https://esploradati.istat.it/databrowser/#/it/censpop/categories/IT1,Z1200CPA,1.0/CPA_POP). Dati estratti il 15 ottobre 2026. Fino al 2011 censimenti generali (date di riferimento in autunno); 2021 Censimento permanente (31 dicembre).
- **Classificazione:** Aree Interne SNAI 2021–2027. La classe attuale è applicata a tutti gli anni.
- **Confini:** ISTAT, tramite [openpolis/geojson-italy](https://github.com/openpolis/geojson-italy) (CC BY 4.0), file `limits_IT_all.topo.json` con confini 2025–2026.

## Ricostruzione ai confini 2021

La serie 1951–2011 di ISTAT è già ricostruita ai confini di qualche anno fa, il 2021 è ai confini del 2021. Per avere una serie omogenea:

- **Fusioni 2020:** Brez, Cagnò, Cloz, Revò e Romallo → Novella; Castelfondo, Fondo e Malosco → Borgo d'Anaunia; Carano, Daiano e Varena → Ville di Fiemme. La storia del nuovo comune è la somma dei comuni d'origine; la popolazione 2021 viene da ISTAT. Classe SNAI: E – Periferico per tutti e tre.
- **Incorporazioni 2020:** Vendrogno → Bellano, Faedo → San Michele all'Adige, Monteciccardo → Pesaro.
- **Scorporo 2021:** Misiliscemi, nato da una parte di Trapani, è riunito a Trapani ("Trapani (incl. Misiliscemi)", classe B). I suoi 8.415 abitanti 2021 sono quindi attribuiti alla classe di Trapani.
- **Cambio di regione:** Montecopiolo e Sassofeltrio passano dalle Marche (Pesaro e Urbino) all'Emilia-Romagna (Rimini).

Risultato: 7.903 comuni; i totali coincidono con quelli ISTAT per tutti gli anni (2021: 59.030.133) e per tutte le regioni nel 2021.

Si assume che, prima della fusione, i confini dei comuni d'origine coincidessero con quelli del comune attuale.

## Confini delle mappe

I confini disponibili sono più recenti dei dati (2025–2026). Per collegarli:

- 377 comuni sardi e Sovizzo hanno cambiato codice ISTAT ma non confini: sono collegati per nome;
- 16 comuni fusi dopo il 2021 condividono sulla mappa il poligono del nuovo comune, colorato con la somma dei loro dati (es. Moransengo e Tonengo → Moransengo-Tonengo); nella scheda del comune restano separati;
- il poligono di Misiliscemi è unito a quello di Trapani.

## Rigenerare i dati

```
python scripts/prepara_dati.py dati-dashboard-popolazione.xlsx data
python scripts/prepara_geo.py limits_IT_all.topo.json data/comuni_2021.csv data
```

Il primo script richiede Python 3 con `openpyxl`; il secondo anche Node.js con `mapshaper` (`npm install mapshaper`). Entrambi controllano codici e nomi dei comuni e si fermano senza scrivere nulla se qualcosa non corrisponde.
