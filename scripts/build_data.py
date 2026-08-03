#!/usr/bin/env python3
"""Legge il piano alimentare in formato .xlsx e produce il JSON che alimenta il sito.

Uso:
    python scripts/build_data.py                       # data/piano_alimentare.xlsx -> site/data/plan.json
    python scripts/build_data.py --xlsx altro.xlsx
    python scripts/build_data.py --standalone dist/piano.html   # pagina singola con i dati dentro

Il parser e' tollerante: cerca i fogli per parole chiave, ignora le righe vuote e
accetta colonne aggiuntive (per esempio una colonna "Preparazione" aggiunta alle
ricette di Rosalia viene letta e pubblicata automaticamente).
"""

from __future__ import annotations

import argparse
import base64
import json
import re
import sys
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("Manca openpyxl: pip install -r requirements.txt")

ROOT = Path(__file__).resolve().parent.parent

GIORNI = ["lunedi", "martedi", "mercoledi", "giovedi", "venerdi", "sabato", "domenica"]
PASTI = ["colazione", "pranzo", "merenda", "cena", "dopo cena", "spuntino"]
MARCATORE_RICETTA = "▸"  # il ▸ che nel foglio segna il piatto di un pasto


# ---------------------------------------------------------------- utilities


def norm(value) -> str:
    """Minuscolo, senza accenti, spazi compattati. Serve per confrontare i nomi."""
    if value is None:
        return ""
    text = str(value).replace(MARCATORE_RICETTA, " ")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip().lower()


def slugify(value: str) -> str:
    text = norm(value)
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-") or "voce"


def clean(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def num(value):
    """Converte in numero quando ha senso, altrimenti restituisce None."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    match = re.search(r"-?\d+(?:[.,]\d+)?", str(value))
    if not match:
        return None
    return round(float(match.group(0).replace(",", ".")), 2)


def is_heading(text: str) -> bool:
    """True per le righe-titolo scritte tutte in maiuscolo (SETTIMANA 1, COLAZIONI, ...)."""
    letters = [ch for ch in text if ch.isalpha()]
    if len(letters) < 3:
        return False
    return sum(1 for ch in letters if ch.isupper()) / len(letters) > 0.8


def find_sheet(wb, *keywords):
    for sheet in wb.worksheets:
        title = norm(sheet.title)
        if all(k in title for k in keywords):
            return sheet
    return None


def rows_of(sheet, width: int = 5):
    """Righe come liste di lunghezza fissa, cosi' l'unpacking non esplode mai."""
    for row in sheet.iter_rows(values_only=True):
        cells = list(row) + [None] * width
        yield cells[:width]


def meta_from_title(text: str) -> dict:
    meta = {}
    kcal = re.search(r"([\d.]+)\s*kcal", text or "", re.I)
    if kcal:
        meta["kcal_target"] = num(kcal.group(1))
    pasti = re.search(r"(\d+)\s*pasti", text or "", re.I)
    if pasti:
        meta["meals_per_day"] = int(pasti.group(1))
    return meta


# ---------------------------------------------------------------- Giovanni


def parse_piano_giovanni(sheet) -> dict:
    righe = list(rows_of(sheet))
    titolo = clean(righe[0][0]) if righe else ""
    sottotitolo = clean(righe[1][0]) if len(righe) > 1 else ""

    giorni, note = [], []
    giorno = pasto = None

    for a, b, c, d, e in righe[2:]:
        ca, cb, cd = clean(a), clean(b), clean(d)

        if ca and norm(ca) in GIORNI and not cb:
            giorno = {"name": ca.capitalize(), "total": num(e), "meals": []}
            giorni.append(giorno)
            pasto = None
            continue

        if giorno is None:
            # coda del foglio: blocchi di note (titolo maiuscolo + corpo)
            if ca and not cb:
                if is_heading(ca) or not note:
                    note.append({"title": ca, "body": ""})
                else:
                    sep = " " if note[-1]["body"] else ""
                    note[-1]["body"] += sep + ca
            continue

        if norm(ca) == "pasto" and norm(cb) == "alimento":
            continue

        if norm(cd).startswith("totale"):
            giorno["total"] = num(e) or giorno["total"]
            giorno = pasto = None
            continue

        if ca and norm(ca) in PASTI:
            pasto = {"name": ca, "recipe": None, "items": []}
            giorno["meals"].append(pasto)

        if pasto is None or not cb:
            continue

        if str(b).lstrip().startswith(MARCATORE_RICETTA):
            pasto["recipe"] = clean(str(b).replace(MARCATORE_RICETTA, ""))
        else:
            pasto["items"].append(
                {
                    "food": cb,
                    "qty": num(c),
                    "kcal_100": num(d),
                    "kcal": num(e),
                }
            )

    for g in giorni:
        for m in g["meals"]:
            m["kcal"] = round(sum(i["kcal"] or 0 for i in m["items"]), 1) or None
        if g["total"] is None:
            g["total"] = round(sum(m["kcal"] or 0 for m in g["meals"]), 1)

    return {
        "title": titolo,
        "subtitle": sottotitolo,
        "weeks": [{"name": "Settimana tipo", "days": giorni}],
        "notes": note,
        **meta_from_title(titolo),
    }


def parse_ricette_giovanni(sheet) -> tuple[list, list]:
    ricette, extra = [], []
    intestazione_vista = False

    for a, b, c, *_ in rows_of(sheet):
        ca, cb, cc = clean(a), clean(b), clean(c)
        if not ca:
            continue
        if norm(ca) == "piatto":
            intestazione_vista = True
            continue
        if not intestazione_vista or not cc:
            continue
        if not cb:
            # senza la colonna "Quando" non e' un piatto del piano ma un blocco libero
            extra.append({"title": ca, "body": cc})
            continue
        ricette.append(
            {
                "slug": slugify(ca),
                "name": ca,
                "when": cb,
                "prep": cc,
                "ingredients": "",
                "kcal": None,
                "category": "Pranzi e cene",
            }
        )
    return ricette, extra


# ---------------------------------------------------------------- Rosalia


def parse_piano_rosalia(sheet) -> dict:
    righe = list(rows_of(sheet, 3))
    titolo = clean(righe[0][0]) if righe else ""
    sottotitolo = clean(righe[1][0]) if len(righe) > 1 else ""

    settimane, media = [], None
    settimana = giorno = None

    for a, b, c in righe[2:]:
        ca, cb = clean(a), clean(b)
        if not ca:
            continue

        if norm(ca).startswith("settimana"):
            settimana = {"name": ca.capitalize(), "days": []}
            settimane.append(settimana)
            continue

        if "media giornaliera" in norm(ca):
            media = num(c)
            continue

        if norm(ca) in GIORNI and not cb:
            if settimana is None:
                settimana = {"name": "Settimana 1", "days": []}
                settimane.append(settimana)
            giorno = {"name": ca.capitalize(), "total": num(c), "meals": []}
            settimana["days"].append(giorno)
            continue

        if giorno is not None and norm(ca) in PASTI and cb:
            giorno["meals"].append(
                {"name": ca, "recipe": cb, "items": [], "kcal": num(c)}
            )

    for s in settimane:
        for g in s["days"]:
            if g["total"] is None:
                g["total"] = round(sum(m["kcal"] or 0 for m in g["meals"]), 1)

    dati = {
        "title": titolo,
        "subtitle": sottotitolo,
        "weeks": settimane,
        "notes": [],
        **meta_from_title(titolo),
    }
    if media:
        dati["daily_average"] = media
    return dati


def parse_ricettario(sheet) -> tuple[list, str, str, list]:
    """Fogli 'Rosalia - Ricette' e 'Ricette da provare': categorie maiuscole + righe piatto."""
    righe = list(rows_of(sheet, 4))
    titolo = clean(righe[0][0]) if righe else ""
    sottotitolo = ""
    ricette, note = [], []
    categoria = "Ricette"

    for indice, (a, b, c, d) in enumerate(righe[1:], start=1):
        ca, cb, cc, cd = clean(a), clean(b), clean(c), clean(d)
        if not ca:
            continue
        if norm(ca) == "piatto":
            continue
        if not cb and not cc:
            if is_heading(ca):
                categoria = ca.capitalize()
            elif not sottotitolo and indice <= 3:
                sottotitolo = ca
            else:
                note.append(ca)
            continue
        if not cc:
            continue
        ricette.append(
            {
                "slug": slugify(ca),
                "name": ca,
                "when": "",
                "prep": cd,
                "ingredients": cc,
                "kcal": num(cb),
                "kcal_label": cb,
                "category": categoria,
            }
        )
    return ricette, titolo, sottotitolo, note


# ---------------------------------------------------------------- spesa


def parse_spesa(sheet) -> dict:
    righe = list(rows_of(sheet, 2))
    titolo = clean(righe[0][0]) if righe else ""
    sottotitolo = clean(righe[1][0]) if len(righe) > 1 else ""

    gruppi = []
    gruppo = categoria = None

    for a, b in righe[2:]:
        ca, cb = clean(a), clean(b)
        if not ca:
            continue

        if cb:
            if gruppo is None:
                gruppo = {"name": "Lista", "person": "", "label": "", "categories": []}
                gruppi.append(gruppo)
            if categoria is None:
                categoria = {"name": "Varie", "items": []}
                gruppo["categories"].append(categoria)
            categoria["items"].append({"name": ca, "qty": cb, "id": slugify(f"{gruppo['name']}-{ca}")})
            continue

        if is_heading(ca) or "—" in ca:
            persona = ""
            for candidato in ("giovanni", "rosalia"):
                if candidato in norm(ca):
                    persona = candidato
            etichetta = ca.split("—", 1)[1].strip() if "—" in ca else ca
            gruppo = {
                "name": ca,
                "person": persona,
                "label": etichetta,
                "slug": slugify(ca),
                "categories": [],
            }
            gruppi.append(gruppo)
            categoria = None
        else:
            if gruppo is None:
                gruppo = {"name": "Lista", "person": "", "label": "", "slug": "lista", "categories": []}
                gruppi.append(gruppo)
            categoria = {"name": ca, "items": []}
            gruppo["categories"].append(categoria)

    return {"title": titolo, "subtitle": sottotitolo, "groups": gruppi}


# ---------------------------------------------------------------- collegamenti


def collega_ricette(persona: dict) -> None:
    """Trasforma il nome del piatto nel piano in un link alla ricetta corrispondente."""
    indice = {norm(r["name"]): r for r in persona["recipes"]}
    for r in persona["recipes"]:
        r.setdefault("used_in", [])

    for settimana in persona["plan"]["weeks"]:
        for giorno in settimana["days"]:
            for pasto in giorno["meals"]:
                titolo = pasto.get("recipe")
                if not titolo:
                    continue
                ricetta = indice.get(norm(titolo))
                pasto["recipe"] = {
                    "title": titolo,
                    "slug": ricetta["slug"] if ricetta else None,
                }
                if ricetta:
                    ricetta["used_in"].append(
                        {
                            "week": settimana["name"],
                            "day": giorno["name"],
                            "meal": pasto["name"],
                        }
                    )

    # per Giovanni le kcal della ricetta si ricavano dal pasto in cui compare
    for settimana in persona["plan"]["weeks"]:
        for giorno in settimana["days"]:
            for pasto in giorno["meals"]:
                slug = (pasto.get("recipe") or {}).get("slug")
                if not slug:
                    continue
                ricetta = next((r for r in persona["recipes"] if r["slug"] == slug), None)
                if ricetta and ricetta.get("kcal") is None and pasto.get("kcal"):
                    ricetta["kcal"] = pasto["kcal"]
                if ricetta and pasto["items"] and not ricetta.get("plan_items"):
                    ricetta["plan_items"] = pasto["items"]


def statistiche(persona: dict) -> dict:
    giorni = [g for s in persona["plan"]["weeks"] for g in s["days"]]
    totali = [g["total"] for g in giorni if g["total"]]
    con_piatto = [
        m
        for s in persona["plan"]["weeks"]
        for g in s["days"]
        for m in g["meals"]
        if (m.get("recipe") or {}).get("title")
    ]
    collegate = [m for m in con_piatto if m["recipe"].get("slug")]
    return {
        "days": len(giorni),
        "meals": sum(len(g["meals"]) for g in giorni),
        "meals_with_dish": len(con_piatto),
        "linked_meals": len(collegate),
        "orphans": sorted({m["recipe"]["title"] for m in con_piatto if not m["recipe"].get("slug")}),
        "recipes": len(persona["recipes"]),
        "avg_kcal": round(sum(totali) / len(totali)) if totali else None,
    }


# ---------------------------------------------------------------- build


def build(xlsx: Path) -> dict:
    wb = openpyxl.load_workbook(xlsx, data_only=True)

    foglio_g = find_sheet(wb, "piano", "settimanal") or find_sheet(wb, "piano")
    foglio_gr = find_sheet(wb, "giovanni", "ricett")
    foglio_r = find_sheet(wb, "rosalia", "piano")
    foglio_rr = find_sheet(wb, "rosalia", "ricett")
    foglio_prova = find_sheet(wb, "provare")
    foglio_spesa = find_sheet(wb, "spesa")

    mancanti = [
        nome
        for nome, foglio in [
            ("piano Giovanni", foglio_g),
            ("ricette Giovanni", foglio_gr),
            ("piano Rosalia", foglio_r),
            ("ricette Rosalia", foglio_rr),
        ]
        if foglio is None
    ]
    if mancanti:
        sys.exit("Fogli non trovati nel file: " + ", ".join(mancanti))

    ricette_g, extra_g = parse_ricette_giovanni(foglio_gr)
    giovanni = {
        "key": "giovanni",
        "name": "Giovanni",
        "accent": "teal",
        "plan": parse_piano_giovanni(foglio_g),
        "recipes": ricette_g,
        "extras": extra_g,
        "detail": "items",  # il piano mostra i singoli alimenti con grammi e kcal
    }
    giovanni["plan"]["notes"] += [
        {"title": e["title"], "body": e["body"]} for e in extra_g if not e["body"].startswith("·")
    ]

    ricette_r, titolo_r, _, note_r = parse_ricettario(foglio_rr)
    rosalia = {
        "key": "rosalia",
        "name": "Rosalia",
        "accent": "rose",
        "plan": parse_piano_rosalia(foglio_r),
        "recipes": ricette_r,
        "extras": [],
        "detail": "dish",  # il piano mostra il piatto, il dettaglio sta nella ricetta
    }
    rosalia["plan"]["notes"] += [{"title": "Note", "body": n} for n in note_r]
    rosalia["recipes_title"] = titolo_r

    for persona in (giovanni, rosalia):
        collega_ricette(persona)
        persona["stats"] = statistiche(persona)

    da_provare = {"title": "", "subtitle": "", "recipes": []}
    if foglio_prova is not None:
        ricette_p, titolo_p, sottotitolo_p, _ = parse_ricettario(foglio_prova)
        da_provare = {"title": titolo_p, "subtitle": sottotitolo_p, "recipes": ricette_p}

    spesa = parse_spesa(foglio_spesa) if foglio_spesa is not None else {"groups": []}

    return {
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "source": xlsx.name,
        "people": [giovanni, rosalia],
        "to_try": da_provare,
        "shopping": spesa,
    }


def scrivi_standalone(dati: dict, sorgente: Path, destinazione: Path) -> None:
    """Pagina singola con CSS, JS e dati incorporati: funziona anche offline."""
    html = (sorgente / "index.html").read_text(encoding="utf-8")
    css = (sorgente / "assets" / "style.css").read_text(encoding="utf-8")
    js = (sorgente / "assets" / "app.js").read_text(encoding="utf-8")

    payload = base64.b64encode(json.dumps(dati, ensure_ascii=False).encode("utf-8")).decode("ascii")
    inline = (
        f"<style>\n{css}\n</style>\n"
        f'<script>window.__PLAN_DATA__ = JSON.parse(new TextDecoder().decode('
        f'Uint8Array.from(atob("{payload}"), c => c.charCodeAt(0))));</script>\n'
    )
    html = html.replace('<link rel="stylesheet" href="assets/style.css" />', inline)
    html = html.replace('<script src="assets/app.js"></script>', f"<script>\n{js}\n</script>")

    destinazione.parent.mkdir(parents=True, exist_ok=True)
    destinazione.write_text(html, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera i dati del sito dal piano alimentare .xlsx")
    parser.add_argument("--xlsx", default=str(ROOT / "data" / "piano_alimentare.xlsx"))
    parser.add_argument("--out", default=str(ROOT / "site" / "data" / "plan.json"))
    parser.add_argument("--standalone", help="scrive anche una pagina HTML unica con i dati dentro")
    args = parser.parse_args()

    xlsx = Path(args.xlsx)
    if not xlsx.exists():
        sys.exit(f"File non trovato: {xlsx}")

    dati = build(xlsx)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(dati, ensure_ascii=False, indent=1), encoding="utf-8")

    for persona in dati["people"]:
        s = persona["stats"]
        print(
            f"  {persona['name']:9s} {s['days']:3d} giorni · {s['meals']:3d} pasti · "
            f"{s['recipes']:3d} ricette · {s['linked_meals']}/{s['meals_with_dish']} piatti collegati"
        )
        if s["orphans"]:
            print("    ricetta mancante per: " + ", ".join(s["orphans"]))
    print(f"  Da provare {len(dati['to_try']['recipes'])} ricette · Spesa {len(dati['shopping']['groups'])} liste")
    print(f"OK -> {out}")

    if args.standalone:
        scrivi_standalone(dati, ROOT / "site", Path(args.standalone))
        print(f"OK -> {args.standalone}")


if __name__ == "__main__":
    main()
