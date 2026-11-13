"""
Prepara i dati della dashboard a partire dal foglio "dati-dashboard-popolazione".

Uso:
    python prepara_dati.py <file_di_input> <cartella_di_output>

Il file di input è il foglio "data" scaricato da Google Sheets
(File → Scarica → Microsoft Excel .xlsx oppure CSV del solo foglio "data").

Lo script:
  1. legge i dati comunali (popolazione ai censimenti 1951–2021);
  2. ricostruisce la serie ai confini comunali del 2021 (vedi RICOSTRUZIONE);
  3. controlla i totali con i valori ISTAT di riferimento;
  4. scrive:
       - data.js          → dati per la dashboard (caricabili anche aprendo
                            index.html con doppio clic, senza server)
       - comuni_2021.csv  → stessi dati in formato tabellare, per lo scaricamento
       - ricostruzione_log.txt → registro delle operazioni svolte

Fonte: ISTAT, Esploradati – Censimenti della popolazione.
Classificazione: Aree Interne SNAI 2021–2027 (applicata a tutti gli anni).
"""

import csv
import json
import sys
from pathlib import Path

ANNI = [1951, 1961, 1971, 1981, 1991, 2001, 2011, 2021]
COLONNE_POP = [f"{a}-popolazione" for a in ANNI]

# Totali censuari ISTAT usati come controllo (Italia)
TOTALE_ISTAT_2021 = 59_030_133
TOTALE_ISTAT_1951 = 47_515_537

# ---------------------------------------------------------------------------
# RICOSTRUZIONE AI CONFINI 2021
# La serie 1951–2011 di ISTAT è già ricostruita ai confini di qualche anno fa;
# il 2021 è ai confini del 2021. Qui si allinea tutto ai confini 2021.
# ---------------------------------------------------------------------------

# 1) Fusioni (2020): i comuni soppressi confluiscono in un nuovo comune.
#    La popolazione 2021 del nuovo comune viene da ISTAT (Censimento permanente).
FUSIONI = [
    {
        "codice": "022253", "comune": "Novella", "classe": "E - Periferico",
        "pop_2021": 3581,
        "soppressi": {"022027": "Brez", "022030": "Cagnò", "022063": "Cloz", "022152": "Revò", "022154": "Romallo"},
    },
    {
        "codice": "022252", "comune": "Borgo d'Anaunia", "classe": "E - Periferico",
        "pop_2021": 2504,
        "soppressi": {"022046": "Castelfondo", "022088": "Fondo", "022111": "Malosco"},
    },
    {
        "codice": "022254", "comune": "Ville di Fiemme", "classe": "E - Periferico",
        "pop_2021": 2637,
        "soppressi": {"022041": "Carano", "022070": "Daiano", "022211": "Varena"},
    },
]

# 2) Incorporazioni (2020): la storia del comune soppresso si somma a quella
#    del comune che lo ha incorporato (il cui 2021 la include già).
# codice soppresso: (nome soppresso, codice incorporante, nome incorporante)
INCORPORAZIONI = {
    "097085": ("Vendrogno", "097008", "Bellano"),
    "022080": ("Faedo", "022167", "San Michele all'Adige"),
    "041032": ("Monteciccardo", "041044", "Pesaro"),
}

# 3) Scorporo (2021): Misiliscemi nasce da una parte di Trapani e non ha una
#    storia propria. Per avere una serie omogenea si riunisce a Trapani.
SCORPORI = [
    {"codice_madre": "081021", "nome_madre": "Trapani", "nuovo_nome": "Trapani (incl. Misiliscemi)",
     "nome_figlio": "Misiliscemi", "pop_2021_figlio": 8415},
]

# 4) Cambio di regione (2021): da Pesaro e Urbino (Marche) a Rimini (Emilia-Romagna).
CAMBI_REGIONE = {
    "041033": {"nome": "Montecopiolo", "codice": "099030", "provincia": "Rimini", "regione": "Emilia-Romagna", "macroregione": "Nord-est"},
    "041060": {"nome": "Sassofeltrio", "codice": "099031", "provincia": "Rimini", "regione": "Emilia-Romagna", "macroregione": "Nord-est"},
}

CLASSI = ["A - Polo", "B - Polo intercomunale", "C - Cintura",
          "D - Intermedio", "E - Periferico", "F - Ultraperiferico"]


def numero(valore):
    """Converte un valore del foglio in intero; '-' o vuoto → None.
    Accetta sia numeri veri sia testo in formato italiano ('1.157')."""
    if valore is None:
        return None
    if isinstance(valore, (int, float)):
        return int(round(valore))
    testo = str(valore).strip()
    if testo in ("", "-", "-%"):
        return None
    testo = testo.replace(".", "").replace(",", ".")
    return int(round(float(testo)))


def leggi_input(percorso):
    percorso = Path(percorso)
    if percorso.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(percorso, read_only=True, data_only=True)
        ws = wb["data"]
        righe = list(ws.iter_rows(values_only=True))
        intestazione = [str(c).strip() if c is not None else "" for c in righe[0]]
        dati = [dict(zip(intestazione, r)) for r in righe[1:] if r and r[0]]
    else:
        with open(percorso, encoding="utf-8-sig", newline="") as f:
            dati = [r for r in csv.DictReader(f) if r.get("codice_istat")]
    comuni = {}
    for r in dati:
        codice = str(r["codice_istat"]).strip().zfill(6)
        classe = str(r["classificazione-ai"]).strip()
        comuni[codice] = {
            "codice": codice,
            "comune": str(r["comune"]).strip(),
            "provincia": str(r["provincia"]).strip(),
            "regione": str(r["regione"]).strip(),
            "macroregione": str(r["macroregione"]).strip(),
            "classe": None if classe in ("-", "") else classe,
            "pop": [numero(r[c]) for c in COLONNE_POP],
        }
    return comuni


def somma(a, b):
    """Somma due valori di popolazione; None conta come 0 solo se l'altro esiste."""
    if a is None and b is None:
        return None
    return (a or 0) + (b or 0)


def verifica_codici(comuni):
    """Controlla che ogni codice usato nella ricostruzione corrisponda al nome atteso."""
    attesi = {}
    for f in FUSIONI:
        attesi.update(f["soppressi"])
    for cs, (ns, cd, nd) in INCORPORAZIONI.items():
        attesi[cs] = ns
        attesi[cd] = nd
    for sc in SCORPORI:
        attesi[sc["codice_madre"]] = sc["nome_madre"]
    for cod, nuovo in CAMBI_REGIONE.items():
        attesi[cod] = nuovo["nome"]
    errori = []
    for cod, nome in attesi.items():
        trovato = comuni.get(cod, {}).get("comune")
        if trovato != nome:
            errori.append(f"Codice {cod}: atteso '{nome}', trovato '{trovato}'")
    return errori


def ricostruisci(comuni, log):
    # Fusioni
    for f in FUSIONI:
        primo = comuni[next(iter(f["soppressi"]))]
        storia = [None] * len(ANNI)
        nomi = []
        for cod in f["soppressi"]:
            c = comuni.pop(cod)
            nomi.append(c["comune"])
            for i in range(len(ANNI) - 1):
                storia[i] = somma(storia[i], c["pop"][i])
        storia[-1] = f["pop_2021"]
        comuni[f["codice"]] = {
            "codice": f["codice"], "comune": f["comune"],
            "provincia": primo["provincia"], "regione": primo["regione"],
            "macroregione": primo["macroregione"], "classe": f["classe"], "pop": storia,
        }
        log.append(f"FUSIONE: {', '.join(nomi)} → {f['comune']} ({f['codice']}), "
                   f"classe {f['classe']}, 1951 = {storia[0]}, 2021 = {storia[-1]}")

    # Incorporazioni
    for cod_sopp, (_, cod_dest, _) in INCORPORAZIONI.items():
        s = comuni.pop(cod_sopp)
        d = comuni[cod_dest]
        prima = d["pop"][0]
        for i in range(len(ANNI) - 1):
            d["pop"][i] = somma(d["pop"][i], s["pop"][i])
        log.append(f"INCORPORAZIONE: {s['comune']} → {d['comune']} ({cod_dest}); "
                   f"1951 da {prima} a {d['pop'][0]}")

    # Scorpori
    for sc in SCORPORI:
        d = comuni[sc["codice_madre"]]
        prima = d["pop"][-1]
        d["pop"][-1] = prima + sc["pop_2021_figlio"]
        log.append(f"SCORPORO: {sc['nome_figlio']} riunito a {d['comune']}; "
                   f"2021 da {prima} a {d['pop'][-1]}; nome → {sc['nuovo_nome']}")
        d["comune"] = sc["nuovo_nome"]

    # Cambi di regione
    for cod_vecchio, nuovo in CAMBI_REGIONE.items():
        c = comuni.pop(cod_vecchio)
        nuovo = {k: v for k, v in nuovo.items() if k != "nome"}
        log.append(f"CAMBIO REGIONE: {c['comune']} {cod_vecchio} ({c['provincia']}, {c['regione']}) → "
                   f"{nuovo['codice']} ({nuovo['provincia']}, {nuovo['regione']})")
        c.update(nuovo)
        comuni[c["codice"]] = c
    return comuni


def controlla(comuni, log):
    errori = []
    tot = [sum(c["pop"][i] or 0 for c in comuni.values()) for i in range(len(ANNI))]
    log.append("TOTALI: " + ", ".join(f"{a}: {t:,}".replace(",", ".") for a, t in zip(ANNI, tot)))
    if tot[0] != TOTALE_ISTAT_1951:
        errori.append(f"Totale 1951 {tot[0]} ≠ ISTAT {TOTALE_ISTAT_1951}")
    if tot[-1] != TOTALE_ISTAT_2021:
        errori.append(f"Totale 2021 {tot[-1]} ≠ ISTAT {TOTALE_ISTAT_2021}")
    for c in comuni.values():
        if c["classe"] not in CLASSI:
            errori.append(f"Classe mancante o non valida: {c['codice']} {c['comune']} ({c['classe']})")
        if c["pop"][-1] is None:
            errori.append(f"Popolazione 2021 mancante: {c['codice']} {c['comune']}")
    log.append(f"COMUNI: {len(comuni)}")
    return errori


def scrivi(comuni, cartella, log):
    cartella = Path(cartella)
    cartella.mkdir(parents=True, exist_ok=True)
    ordinati = sorted(comuni.values(), key=lambda c: c["codice"])

    # Tabelle di decodifica per tenere il file leggero
    province = sorted({c["provincia"] for c in ordinati})
    regioni = sorted({c["regione"] for c in ordinati})
    macro = ["Nord-ovest", "Nord-est", "Centro", "Sud", "Isole"]
    reg_macro = {c["regione"]: c["macroregione"] for c in ordinati}
    prov_reg = {c["provincia"]: c["regione"] for c in ordinati}

    dati = {
        "fonte": "ISTAT, Censimenti della popolazione (Esploradati). Elaborazione ai confini comunali 2021.",
        "anni": ANNI,
        "classi": CLASSI,
        "macroregioni": macro,
        "regioni": [{"nome": r, "macro": macro.index(reg_macro[r])} for r in regioni],
        "province": [{"nome": p, "regione": regioni.index(prov_reg[p])} for p in province],
        # ogni comune: [codice, nome, indice provincia, indice classe, pop 1951…2021]
        "campi": ["codice", "comune", "provincia", "classe"] + [str(a) for a in ANNI],
        "comuni": [[c["codice"], c["comune"], province.index(c["provincia"]),
                    CLASSI.index(c["classe"])] + c["pop"] for c in ordinati],
    }
    testo = json.dumps(dati, ensure_ascii=False, separators=(",", ":"))
    (cartella / "data.js").write_text(
        "// Generato da scripts/prepara_dati.py – non modificare a mano\n"
        f"window.DATI = {testo};\n", encoding="utf-8")

    with open(cartella / "comuni_2021.csv", "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["codice_istat", "comune", "provincia", "regione", "macroregione",
                    "classificazione_ai"] + [f"pop_{a}" for a in ANNI] + ["var_perc_1951_2021"])
        for c in ordinati:
            p = c["pop"]
            var = "" if not p[0] or p[-1] is None else round((p[-1] - p[0]) / p[0] * 100, 1)
            w.writerow([c["codice"], c["comune"], c["provincia"], c["regione"], c["macroregione"],
                        c["classe"]] + ["" if x is None else x for x in p] + [var])

    (cartella / "ricostruzione_log.txt").write_text("\n".join(log) + "\n", encoding="utf-8")


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    log = []
    comuni = leggi_input(sys.argv[1])
    log.append(f"LETTI: {len(comuni)} comuni da {Path(sys.argv[1]).name}")
    errori = verifica_codici(comuni)
    if errori:
        print("ERRORI nei codici della ricostruzione – file non scritti:")
        for e in errori:
            print("  -", e)
        sys.exit(2)
    comuni = ricostruisci(comuni, log)
    errori = controlla(comuni, log)
    if errori:
        print("ERRORI – file non scritti:")
        for e in errori:
            print("  -", e)
        sys.exit(2)
    scrivi(comuni, sys.argv[2], log)
    print("\n".join(log))
    print(f"\nOK: file scritti in {sys.argv[2]}")


if __name__ == "__main__":
    main()
