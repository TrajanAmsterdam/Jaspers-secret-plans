import sys
import csv
import re
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Toplevel, MULTIPLE, END
from pathlib import Path
from datetime import datetime, timedelta
from openpyxl import load_workbook, Workbook


# Voertuigtype-labels in de uitvoer.
# Fietsdata (weektabellen) krijgt "fiets"; telslang/moto-data (dagtabellen) dit label.
FIETS_VOERTUIGTYPE = "fiets"
TELSLANG_VOERTUIGTYPE = "motorvoertuig"


# Standaard dagdeel-indeling. Wordt bij de eerste run weggeschreven naar
# "dagdelen.csv" naast het script/de exe. Gebruikers kunnen dat bestand
# aanpassen (tijden en namen) zonder de exe opnieuw te bouwen.
# "van" inclusief, "tot" exclusief. Loopt "van" voorbij "tot" (bv. 19:00-07:00),
# dan loopt het dagdeel over middernacht heen.
DEFAULT_DAGDELEN = [
    ("07:00", "10:00", "ochtendspits"),
    ("10:00", "16:00", "overdag"),
    ("16:00", "19:00", "avondspits"),
    ("19:00", "07:00", "nacht"),
]


# Mapping van kolomnaam naar weekdag-offset (maandag=0) voor fiets-weektabellen
DAG_KOLOMMEN = {
    "Maandag": 0,
    "Dinsdag": 1,
    "Woensdag": 2,
    "Donderdag": 3,
    "Vrijdag": 4,
    "Zaterdag": 5,
    "Zondag": 6,
}

# Kop van een fiets-weektabel:  "16-02-2026 - 22-02-2026 - <richting>"
WEEK_HEADER = re.compile(
    r"^(\d{2}-\d{2}-\d{4})\s*-\s*(\d{2}-\d{2}-\d{4})\s*-\s*(.+)$"
)
# Kop van een telslang-dagtabel: "16-02-2026 - <richting>"
DAG_HEADER = re.compile(r"^(\d{2}-\d{2}-\d{4})\s*-\s*(.+)$")


# ---------------------------------------------------------------------------
# Hulpfuncties: dagdelen-config, dagsoort, dagdeel
# ---------------------------------------------------------------------------

def basis_dir():
    """Map waarin het script of (na bouwen) de exe staat."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def _tijd_naar_minuten(waarde):
    """'07:00' -> 420 minuten. Geeft None bij geen geldige tijd."""
    m = re.match(r"^\s*(\d{1,2}):(\d{2})\s*$", str(waarde))
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


def lees_dagdelen_strings(base_dir):
    """Lees opgeslagen dagdelen (als tekst-tuples van/tot/naam) uit dagdelen.csv.
    Bestaat het bestand niet, dan de ingebouwde standaard."""
    pad = base_dir / "dagdelen.csv"
    rijen = []
    try:
        with open(pad, "r", encoding="utf-8-sig") as f:
            for row in csv.reader(f, delimiter=";"):
                if not row or not str(row[0]).strip() or str(row[0]).strip().startswith("#"):
                    continue
                if len(row) < 3:
                    continue
                van, tot, naam = str(row[0]).strip(), str(row[1]).strip(), str(row[2]).strip()
                if _tijd_naar_minuten(van) is None or _tijd_naar_minuten(tot) is None or not naam:
                    continue  # slaat ook de kopregel 'van;tot;dagdeel' over
                rijen.append((van, tot, naam))
    except OSError:
        pass
    return rijen if rijen else [tuple(r) for r in DEFAULT_DAGDELEN]


def schrijf_dagdelen_strings(base_dir, rijen):
    """Sla dagdelen op zodat de keuze de volgende keer onthouden wordt."""
    try:
        with open(base_dir / "dagdelen.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["van", "tot", "dagdeel"])
            for van, tot, naam in rijen:
                w.writerow([van, tot, naam])
    except OSError:
        pass


def strings_naar_dagdelen(rijen):
    """Zet tekst-tuples om naar (van_minuten, tot_minuten, naam)."""
    out = []
    for van, tot, naam in rijen:
        vm, tm = _tijd_naar_minuten(van), _tijd_naar_minuten(tot)
        if vm is not None and tm is not None and naam:
            out.append((vm, tm, naam))
    return out


def load_dagdelen(base_dir):
    """Niet-grafisch laden (voor tests/headless): geeft (van_min, tot_min, naam)."""
    return strings_naar_dagdelen(lees_dagdelen_strings(base_dir))


def vraag_dagdelen_dialog(base_dir):
    """Toon een bewerkbaar invul-venster met de dagdelen. Onthoudt de keuze
    in dagdelen.csv. Geeft een lijst (van_min, tot_min, naam)."""
    prefill = lees_dagdelen_strings(base_dir)

    dialog = Toplevel()
    dialog.title("Dagdelen instellen")
    dialog.geometry("440x380")
    dialog.minsize(420, 320)
    dialog.grab_set()

    tk.Label(
        dialog,
        text=("Stel de dagdelen in. Tijden als uu:mm (bv. 07:00).\n"
              "'van' inclusief, 'tot' exclusief. Loopt 'tot' voorbij 'van'\n"
              "(bv. 19:00 - 07:00), dan loopt het dagdeel over middernacht."),
        justify="left",
    ).pack(padx=10, pady=(10, 5), anchor="w")

    head = tk.Frame(dialog)
    head.pack(padx=10, fill="x")
    tk.Label(head, text="Van", width=8, anchor="w").grid(row=0, column=0)
    tk.Label(head, text="Tot", width=8, anchor="w").grid(row=0, column=1)
    tk.Label(head, text="Dagdeel", width=22, anchor="w").grid(row=0, column=2)

    rows_frame = tk.Frame(dialog)
    rows_frame.pack(padx=10, fill="both", expand=True)
    row_widgets = []

    def add_row(van="", tot="", naam=""):
        rf = tk.Frame(rows_frame)
        rf.pack(fill="x", pady=1)
        e_van = tk.Entry(rf, width=8)
        e_van.insert(0, van)
        e_van.grid(row=0, column=0, padx=1)
        e_tot = tk.Entry(rf, width=8)
        e_tot.insert(0, tot)
        e_tot.grid(row=0, column=1, padx=1)
        e_naam = tk.Entry(rf, width=22)
        e_naam.insert(0, naam)
        e_naam.grid(row=0, column=2, padx=1)
        entry = (e_van, e_tot, e_naam)

        def remove():
            if entry in row_widgets:
                row_widgets.remove(entry)
            rf.destroy()

        tk.Button(rf, text="x", width=2, command=remove).grid(row=0, column=3, padx=3)
        row_widgets.append(entry)

    def herbouw(rijen):
        for child in rows_frame.winfo_children():
            child.destroy()
        row_widgets.clear()
        for van, tot, naam in rijen:
            add_row(van, tot, naam)

    herbouw(prefill)

    knoppen = tk.Frame(dialog)
    knoppen.pack(padx=10, pady=6, fill="x")
    tk.Button(knoppen, text="+ Regel toevoegen", command=lambda: add_row()).pack(side="left")
    tk.Button(knoppen, text="Standaard", command=lambda: herbouw([tuple(r) for r in DEFAULT_DAGDELEN])).pack(side="left", padx=5)

    resultaat = {"rijen": None}

    def on_ok():
        rijen = []
        for e_van, e_tot, e_naam in row_widgets:
            van, tot, naam = e_van.get().strip(), e_tot.get().strip(), e_naam.get().strip()
            if not (van or tot or naam):
                continue
            if _tijd_naar_minuten(van) is None or _tijd_naar_minuten(tot) is None or not naam:
                messagebox.showerror(
                    "Ongeldige invoer",
                    f"Controleer deze regel:\n van='{van}'  tot='{tot}'  dagdeel='{naam}'\n\n"
                    "Gebruik tijden als uu:mm en vul een naam in.",
                )
                return
            rijen.append((van, tot, naam))
        if not rijen:
            messagebox.showerror("Leeg", "Voeg minstens één dagdeel toe.")
            return
        resultaat["rijen"] = rijen
        dialog.destroy()

    ok_frame = tk.Frame(dialog)
    ok_frame.pack(padx=10, pady=(0, 10), fill="x")
    tk.Button(ok_frame, text="OK", width=12, command=on_ok).pack(side="right", padx=3)
    tk.Button(ok_frame, text="Annuleren", width=12, command=dialog.destroy).pack(side="right", padx=3)

    dialog.wait_window()

    if resultaat["rijen"] is None:
        # geannuleerd: gebruik de vorige/standaard instelling zonder op te slaan
        return strings_naar_dagdelen(prefill)
    schrijf_dagdelen_strings(base_dir, resultaat["rijen"])
    return strings_naar_dagdelen(resultaat["rijen"])


def bepaal_dagdeel(tijd_str, dagdelen):
    t = _tijd_naar_minuten(tijd_str)
    if t is None:
        return ""
    for van, tot, naam in dagdelen:
        if van < tot:
            if van <= t < tot:
                return naam
        elif van > tot:  # loopt over middernacht
            if t >= van or t < tot:
                return naam
        else:  # van == tot -> hele dag
            return naam
    return ""


def bepaal_dagsoort(datum):
    """weekdag (ma-vr) of weekenddag (za-zo)."""
    return "weekenddag" if datum.weekday() >= 5 else "weekdag"


# ---------------------------------------------------------------------------
# Inlezen
# ---------------------------------------------------------------------------

def read_csv_rows(file_path):
    with open(file_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=";")
        return list(reader)


def read_xlsx_rows(file_path, sheet_name):
    wb = load_workbook(file_path, read_only=True, data_only=True)
    ws = wb[sheet_name]
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append([cell if cell is not None else "" for cell in row])
    wb.close()
    return rows


def select_sheets(file_path):
    """Open een dialoog om tabbladen te kiezen uit een xlsx bestand."""
    wb = load_workbook(file_path, read_only=True)
    sheet_names = wb.sheetnames
    wb.close()

    if len(sheet_names) == 1:
        return sheet_names

    dialog = Toplevel()
    dialog.title(f"Tabbladen - {Path(file_path).name}")
    dialog.geometry("450x400")
    dialog.minsize(350, 250)
    dialog.grab_set()

    tk.Label(dialog, text="Selecteer tabbladen om samen te voegen (Ctrl+klik voor meerdere):").pack(pady=5)
    listbox = Listbox(dialog, selectmode=MULTIPLE, width=50, height=15)
    listbox.pack(padx=10, pady=5, fill="both", expand=True)
    for name in sheet_names:
        listbox.insert(END, name)

    selected = []

    def on_ok():
        for idx in listbox.curselection():
            selected.append(sheet_names[idx])
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="OK", command=on_ok, width=15).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Annuleren", command=dialog.destroy, width=15).pack(side="left", padx=5)
    dialog.wait_window()
    return selected


def to_number(value, empty=""):
    """Converteer een waarde naar een getal. Lege waarde -> 'empty'."""
    if isinstance(value, (int, float)):
        return value
    if not isinstance(value, str):
        return value
    value = value.strip()
    if not value:
        return empty
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value.replace(",", "."))
    except ValueError:
        return value


def _num(value):
    """Waarde als getal voor optelling; niet-getallen tellen als 0."""
    return value if isinstance(value, (int, float)) else 0


def _cell0(row):
    """Eerste cel van een rij als gestripte string ('' als leeg)."""
    if row and len(row) > 0 and row[0] is not None:
        return str(row[0]).strip()
    return ""


# ---------------------------------------------------------------------------
# Parsen
# ---------------------------------------------------------------------------

def parse_rows(rows):
    """Parse één tabel (CSV of xlsx-sheet) en herken per blok automatisch
    of het een fiets-weektabel of een telslang-dagtabel is.

    Geeft terug: (lijst van rij-dicts, telslang_klassekolommen of None).
    """
    data = []
    telslang_headers = None
    i = 0
    n = len(rows)

    while i < n:
        cell0 = _cell0(rows[i])
        if cell0:
            # ---- FIETS: weektabel (eerst proberen; week-bereik is specifieker) ----
            wm = WEEK_HEADER.match(cell0)
            if wm:
                start_date = datetime.strptime(wm.group(1), "%d-%m-%Y")
                richting = wm.group(3).strip()

                col_header = rows[i + 1] if i + 1 < n else []
                dag_indices = {}
                for j, col in enumerate(col_header):
                    name = str(col).strip()
                    if name in DAG_KOLOMMEN:
                        dag_indices[name] = j

                i += 2  # skip tabelkop + kolom-headerrij
                uur_rijen = []
                while i < n and _cell0(rows[i]):
                    uur_rijen.append(rows[i])
                    i += 1

                for dag_naam, offset in DAG_KOLOMMEN.items():
                    if dag_naam in dag_indices:
                        datum = start_date + timedelta(days=offset)
                        idx = dag_indices[dag_naam]
                        for dr in uur_rijen:
                            tijd = str(dr[0]).strip()
                            aantal = to_number(dr[idx], empty=0) if idx < len(dr) else 0
                            data.append({
                                "Voertuigtype": FIETS_VOERTUIGTYPE,
                                "Datum": datum,
                                "Tijd": tijd,
                                "Richting": richting,
                                "classes": {},
                                "V85": "",
                                "Gem.": "",
                                "Totaal": aantal,
                            })
                continue

            # ---- TELSLANG/MOTO: dagtabel ----
            dm = DAG_HEADER.match(cell0)
            if dm:
                datum = datetime.strptime(dm.group(1), "%d-%m-%Y")
                richting = dm.group(2).strip()

                header1 = rows[i + 1] if i + 1 < n else []  # lengteklasse
                header2 = rows[i + 2] if i + 2 < n else []  # snelheidsklasse
                block_headers = []
                for j in range(1, 25):
                    lc = str(header1[j]).strip() if j < len(header1) else ""
                    sc = str(header2[j]).strip() if j < len(header2) else ""
                    block_headers.append(f"{lc};{sc}")
                if telslang_headers is None:
                    telslang_headers = block_headers

                i += 3  # skip tabelkop + 2 headerrijen
                while i < n and _cell0(rows[i]):
                    dr = rows[i]
                    tijd = str(dr[0]).strip()
                    classes = {}
                    for k, h in enumerate(block_headers):
                        j = k + 1
                        classes[h] = to_number(dr[j]) if j < len(dr) else ""
                    v85 = to_number(dr[27]) if len(dr) > 27 else ""
                    gem = to_number(dr[29]) if len(dr) > 29 else ""
                    totaal = to_number(dr[30]) if len(dr) > 30 else ""
                    data.append({
                        "Voertuigtype": TELSLANG_VOERTUIGTYPE,
                        "Datum": datum,
                        "Tijd": tijd,
                        "Richting": richting,
                        "classes": classes,
                        "V85": v85,
                        "Gem.": gem,
                        "Totaal": totaal,
                    })
                    i += 1
                continue
        i += 1

    return data, telslang_headers


# ---------------------------------------------------------------------------
# Uitvoer opbouwen
# ---------------------------------------------------------------------------

def groepeer_klassen(telslang_headers):
    """Groepeer klassekolommen op lengteklasse en op snelheidsklasse.
    Geeft twee geordende dicts: {lengte: [headers]}, {snelheid: [headers]}."""
    lengte_groepen = {}
    snelheid_groepen = {}
    for h in telslang_headers:
        parts = h.split(";")
        lengte = parts[0] if parts else h
        snelheid = parts[1] if len(parts) > 1 else ""
        lengte_groepen.setdefault(lengte, []).append(h)
        snelheid_groepen.setdefault(snelheid, []).append(h)
    return lengte_groepen, snelheid_groepen


def build_output(all_rows, telslang_headers, dagdelen):
    """Bouw de kopregel en de uitvoerrijen. Eén tabblad met alle data."""
    lengte_groepen, snelheid_groepen = ({}, {})
    if telslang_headers:
        lengte_groepen, snelheid_groepen = groepeer_klassen(telslang_headers)

    columns = ["Locatie", "Voertuigtype", "Datum", "Dagsoort", "Tijd", "Dagdeel", "Richting"]
    if telslang_headers:
        columns += list(telslang_headers)
        columns += [f"Totaal lengte {l}" for l in lengte_groepen]
        columns += [f"Totaal snelheid {s}" for s in snelheid_groepen]
        columns += ["V85", "Gem."]
    columns += ["Totaal"]

    out_rows = []
    for d in all_rows:
        dagsoort = bepaal_dagsoort(d["Datum"])
        dagdeel = bepaal_dagdeel(d["Tijd"], dagdelen)
        row = [d["Locatie"], d["Voertuigtype"], d["Datum"], dagsoort,
               d["Tijd"], dagdeel, d["Richting"]]

        if telslang_headers:
            is_tel = bool(d["classes"])
            for h in telslang_headers:
                row.append(d["classes"].get(h, ""))
            # totalen per lengteklasse
            for l, hdrs in lengte_groepen.items():
                row.append(sum(_num(d["classes"].get(h, "")) for h in hdrs) if is_tel else "")
            # totalen per snelheidsklasse
            for s, hdrs in snelheid_groepen.items():
                row.append(sum(_num(d["classes"].get(h, "")) for h in hdrs) if is_tel else "")
            row.append(d["V85"])
            row.append(d["Gem."])

        row.append(d["Totaal"])
        out_rows.append(row)

    return columns, out_rows


# ---------------------------------------------------------------------------
# Hoofdprogramma
# ---------------------------------------------------------------------------

def main():
    root = tk.Tk()
    root.withdraw()

    dagdelen = vraag_dagdelen_dialog(basis_dir())

    print("Selecteer invoerbestand(en)...")
    input_files = filedialog.askopenfilenames(
        title="Selecteer invoerbestanden (fiets en/of telslang/moto)",
        filetypes=[
            ("CSV en Excel bestanden", "*.csv *.xlsx"),
            ("CSV bestanden", "*.csv"),
            ("Excel bestanden", "*.xlsx"),
            ("Alle bestanden", "*.*"),
        ],
    )
    if not input_files:
        print("Geen bestanden geselecteerd, script gestopt.")
        return

    all_rows = []
    telslang_headers = None

    for file_path in input_files:
        path = Path(file_path)
        locatie = path.stem
        print(f"Verwerken: {path.name}")

        if path.suffix.lower() == ".xlsx":
            sheets = select_sheets(file_path)
            if not sheets:
                print(f"  Geen tabbladen geselecteerd voor {path.name}, overgeslagen.")
                continue
            for sheet in sheets:
                print(f"  Tabblad: {sheet}")
                rows = read_xlsx_rows(file_path, sheet)
                data, headers = parse_rows(rows)
                if headers and telslang_headers is None:
                    telslang_headers = headers
                for d in data:
                    d["Locatie"] = locatie
                    all_rows.append(d)
        else:
            rows = read_csv_rows(file_path)
            data, headers = parse_rows(rows)
            if headers and telslang_headers is None:
                telslang_headers = headers
            for d in data:
                d["Locatie"] = locatie
                all_rows.append(d)

    if not all_rows:
        print("Geen data gevonden in de geselecteerde bestanden.")
        return

    out_header, out_rows = build_output(all_rows, telslang_headers, dagdelen)

    n_fiets = sum(1 for d in all_rows if d["Voertuigtype"] == FIETS_VOERTUIGTYPE)
    n_telslang = len(all_rows) - n_fiets
    print(f"{len(all_rows)} rijen verwerkt ({n_fiets} fiets, {n_telslang} telslang/moto).")

    print("Selecteer opslaglocatie voor het uitvoerbestand...")
    output_file = filedialog.asksaveasfilename(
        title="Opslaan als",
        initialdir=str(Path(input_files[0]).parent),
        initialfile="samengevoegd.xlsx",
        defaultextension=".xlsx",
        filetypes=[("Excel bestanden", "*.xlsx"), ("CSV bestanden", "*.csv"), ("Alle bestanden", "*.*")],
    )
    if not output_file:
        print("Geen opslaglocatie geselecteerd, script gestopt.")
        return

    datum_kolom = 3  # kolom C (Locatie, Voertuigtype, Datum, ...)

    if output_file.lower().endswith(".xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Samengevoegd"
        ws.append(out_header)
        for row in out_rows:
            ws.append(row)
        for row_idx in range(2, len(out_rows) + 2):
            ws.cell(row=row_idx, column=datum_kolom).number_format = "DD-MM-YYYY"
        wb.save(output_file)
    else:
        with open(output_file, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(out_header)
            for row in out_rows:
                csv_row = list(row)
                csv_row[datum_kolom - 1] = row[datum_kolom - 1].strftime("%d-%m-%Y")
                writer.writerow(csv_row)

    print(f"Klaar! {len(out_rows)} rijen geschreven naar {output_file}")


if __name__ == "__main__":
    main()
