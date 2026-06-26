import sys
import csv
import re
from pathlib import Path
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook

# Optionele onderdelen: slepen-en-neerzetten en kaart. Ontbreken ze, dan
# blijft de tool werken (zonder die functie).
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except Exception:
    HAS_DND = False

try:
    from tkintermapview import TkinterMapView
    HAS_MAP = True
except Exception:
    HAS_MAP = False


# Voertuigtype-labels in de uitvoer.
FIETS_VOERTUIGTYPE = "fiets"
TELSLANG_VOERTUIGTYPE = "motorvoertuig"

# Standaard dagdeel-indeling (van inclusief, tot exclusief; van>tot loopt over middernacht)
DEFAULT_DAGDELEN = [
    ("07:00", "10:00", "ochtendspits"),
    ("10:00", "16:00", "overdag"),
    ("16:00", "19:00", "avondspits"),
    ("19:00", "07:00", "nacht"),
]

DAG_KOLOMMEN = {
    "Maandag": 0, "Dinsdag": 1, "Woensdag": 2, "Donderdag": 3,
    "Vrijdag": 4, "Zaterdag": 5, "Zondag": 6,
}

WEEK_HEADER = re.compile(r"^(\d{2}-\d{2}-\d{4})\s*-\s*(\d{2}-\d{2}-\d{4})\s*-\s*(.+)$")
DAG_HEADER = re.compile(r"^(\d{2}-\d{2}-\d{4})\s*-\s*(.+)$")


# ---------------------------------------------------------------------------
# Dagdelen / dagsoort
# ---------------------------------------------------------------------------

def basis_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def _tijd_naar_minuten(waarde):
    m = re.match(r"^\s*(\d{1,2}):(\d{2})\s*$", str(waarde))
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


def lees_dagdelen_strings(base_dir):
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
                    continue
                rijen.append((van, tot, naam))
    except OSError:
        pass
    return rijen if rijen else [tuple(r) for r in DEFAULT_DAGDELEN]


def schrijf_dagdelen_strings(base_dir, rijen):
    try:
        with open(base_dir / "dagdelen.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["van", "tot", "dagdeel"])
            for van, tot, naam in rijen:
                w.writerow([van, tot, naam])
    except OSError:
        pass


def strings_naar_dagdelen(rijen):
    out = []
    for van, tot, naam in rijen:
        vm, tm = _tijd_naar_minuten(van), _tijd_naar_minuten(tot)
        if vm is not None and tm is not None and naam:
            out.append((vm, tm, naam))
    return out


def load_dagdelen(base_dir):
    return strings_naar_dagdelen(lees_dagdelen_strings(base_dir))


def bepaal_dagdeel(tijd_str, dagdelen):
    t = _tijd_naar_minuten(tijd_str)
    if t is None:
        return ""
    for van, tot, naam in dagdelen:
        if van < tot:
            if van <= t < tot:
                return naam
        elif van > tot:
            if t >= van or t < tot:
                return naam
        else:
            return naam
    return ""


def bepaal_dagsoort(datum):
    return "weekenddag" if datum.weekday() >= 5 else "weekdag"


# ---------------------------------------------------------------------------
# Inlezen / parsen
# ---------------------------------------------------------------------------

def read_csv_rows(file_path):
    with open(file_path, "r", encoding="utf-8-sig") as f:
        return list(csv.reader(f, delimiter=";"))


def read_xlsx_rows(file_path, sheet_name):
    wb = load_workbook(file_path, read_only=True, data_only=True)
    ws = wb[sheet_name]
    rows = [[c if c is not None else "" for c in row] for row in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def to_number(value, empty=""):
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
    return value if isinstance(value, (int, float)) else 0


def _cell0(row):
    if row and len(row) > 0 and row[0] is not None:
        return str(row[0]).strip()
    return ""


def parse_rows(rows):
    data = []
    telslang_headers = None
    i, n = 0, len(rows)
    while i < n:
        cell0 = _cell0(rows[i])
        if cell0:
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
                i += 2
                uur_rijen = []
                while i < n and _cell0(rows[i]):
                    uur_rijen.append(rows[i]); i += 1
                for dag_naam, offset in DAG_KOLOMMEN.items():
                    if dag_naam in dag_indices:
                        datum = start_date + timedelta(days=offset)
                        idx = dag_indices[dag_naam]
                        for dr in uur_rijen:
                            tijd = str(dr[0]).strip()
                            aantal = to_number(dr[idx], empty=0) if idx < len(dr) else 0
                            data.append({"Voertuigtype": FIETS_VOERTUIGTYPE, "Datum": datum,
                                         "Tijd": tijd, "Richting": richting, "classes": {},
                                         "V85": "", "Gem.": "", "Totaal": aantal})
                continue
            dm = DAG_HEADER.match(cell0)
            if dm:
                datum = datetime.strptime(dm.group(1), "%d-%m-%Y")
                richting = dm.group(2).strip()
                header1 = rows[i + 1] if i + 1 < n else []
                header2 = rows[i + 2] if i + 2 < n else []
                block_headers = []
                for j in range(1, 25):
                    lc = str(header1[j]).strip() if j < len(header1) else ""
                    sc = str(header2[j]).strip() if j < len(header2) else ""
                    block_headers.append(f"{lc};{sc}")
                if telslang_headers is None:
                    telslang_headers = block_headers
                i += 3
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
                    data.append({"Voertuigtype": TELSLANG_VOERTUIGTYPE, "Datum": datum,
                                 "Tijd": tijd, "Richting": richting, "classes": classes,
                                 "V85": v85, "Gem.": gem, "Totaal": totaal})
                    i += 1
                continue
        i += 1
    return data, telslang_headers


def lees_info_coordinaten(path):
    """Lees locatienaam/plaats en lat/lon uit het Info-blad van een xlsx."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return None
    sheet = next((s for s in wb.sheetnames if s.strip().lower() == "info"), wb.sheetnames[0])
    rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
    wb.close()

    def f(v):
        try:
            return float(str(v).replace(",", "."))
        except (TypeError, ValueError):
            return None

    naam = plaats = lat = lon = None
    for row in rows:
        cells = [("" if c is None else str(c)).strip() for c in row]
        if not cells:
            continue
        if cells[0].lower() == "naam" and naam is None:
            naam = next((c for c in cells[1:] if c), None)
        if cells[0].lower() == "plaats" and plaats is None:
            plaats = next((c for c in cells[1:] if c), None)
        for i, c in enumerate(cells):
            if c.lower() == "locatie" and i + 2 < len(row):
                a, b = f(row[i + 1]), f(row[i + 2])
                if a is not None and b is not None:
                    lat, lon = a, b
    if lat is None or lon is None:
        return None
    return {"naam": naam, "plaats": plaats, "lat": lat, "lon": lon}


def verwerk_bestand(path, verwacht_label):
    """Lees een bestand en geef de rijen van het verwachte type terug.
    Voor xlsx wordt automatisch het juiste data-tabblad gekozen."""
    p = Path(path)
    locatie = p.stem
    if p.suffix.lower() == ".csv":
        data, headers = parse_rows(read_csv_rows(path))
        typed = [d for d in data if d["Voertuigtype"] == verwacht_label]
        for d in typed:
            d["Locatie"] = locatie
        return typed, (headers if verwacht_label == TELSLANG_VOERTUIGTYPE else None)

    try:
        wb = load_workbook(path, read_only=True)
        sheetnames = wb.sheetnames
        wb.close()
    except Exception:
        return [], None
    for s in sheetnames:
        if s.strip().lower() == "info":
            continue
        data, headers = parse_rows(read_xlsx_rows(path, s))
        typed = [d for d in data if d["Voertuigtype"] == verwacht_label]
        if typed:
            for d in typed:
                d["Locatie"] = locatie
            return typed, headers
    return [], None


# ---------------------------------------------------------------------------
# Uitvoer opbouwen
# ---------------------------------------------------------------------------

def groepeer_klassen(telslang_headers):
    lengte_groepen, snelheid_groepen = {}, {}
    for h in telslang_headers:
        parts = h.split(";")
        lengte = parts[0] if parts else h
        snelheid = parts[1] if len(parts) > 1 else ""
        lengte_groepen.setdefault(lengte, []).append(h)
        snelheid_groepen.setdefault(snelheid, []).append(h)
    return lengte_groepen, snelheid_groepen


def build_output(all_rows, telslang_headers, dagdelen):
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
            for l, hdrs in lengte_groepen.items():
                row.append(sum(_num(d["classes"].get(h, "")) for h in hdrs) if is_tel else "")
            for s, hdrs in snelheid_groepen.items():
                row.append(sum(_num(d["classes"].get(h, "")) for h in hdrs) if is_tel else "")
            row.append(d["V85"]); row.append(d["Gem."])
        row.append(d["Totaal"])
        out_rows.append(row)
    return columns, out_rows


def schrijf_uitvoer(output_file, out_header, out_rows):
    datum_kolom = 3
    if output_file.lower().endswith(".xlsx"):
        wb = Workbook(); ws = wb.active; ws.title = "Samengevoegd"
        ws.append(out_header)
        for row in out_rows:
            ws.append(row)
        for ri in range(2, len(out_rows) + 2):
            ws.cell(row=ri, column=datum_kolom).number_format = "DD-MM-YYYY"
        wb.save(output_file)
    else:
        with open(output_file, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(out_header)
            for row in out_rows:
                csv_row = list(row)
                csv_row[datum_kolom - 1] = row[datum_kolom - 1].strftime("%d-%m-%Y")
                writer.writerow(csv_row)


# ---------------------------------------------------------------------------
# Grafische applicatie
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        root.title("Verkeerstellingen samenvoegen")
        root.geometry("1180x700")
        root.minsize(980, 600)

        self.bu_files = []
        self.fiets_files = []
        self.coords = {}   # path -> {naam, plaats, lat, lon}
        self.dd_rows = []

        root.rowconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)

        links = tk.Frame(root)
        links.grid(row=0, column=0, sticky="ns", padx=8, pady=8)

        self._bouw_dagdelen(links)
        self.bu_listbox = self._bouw_dropzone(
            links, "Gemotoriseerd vervoer",
            "dit zijn meestal de BU-bestanden", self.bu_files, TELSLANG_VOERTUIGTYPE)
        self.fiets_listbox = self._bouw_dropzone(
            links, "Fietsen",
            "dit zijn meestal de TU-xxx-fiets bestanden", self.fiets_files, FIETS_VOERTUIGTYPE)

        # Kaart
        kaart_frame = tk.LabelFrame(root, text="Locaties (controle op compleetheid)")
        kaart_frame.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        kaart_frame.rowconfigure(0, weight=1)
        kaart_frame.columnconfigure(0, weight=1)
        if HAS_MAP:
            self.map_widget = TkinterMapView(kaart_frame, corner_radius=0)
            self.map_widget.grid(row=0, column=0, sticky="nsew")
            self.map_widget.set_position(52.2, 5.3)  # midden NL
            self.map_widget.set_zoom(7)
        else:
            self.map_widget = None
            tk.Label(kaart_frame, text="(kaart niet beschikbaar)").grid(row=0, column=0)

        # Onderbalk
        onder = tk.Frame(root)
        onder.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        self.status = tk.Label(onder, text="Sleep bestanden in de velden of gebruik 'Bladeren'.", anchor="w")
        self.status.pack(side="left", fill="x", expand=True)
        tk.Button(onder, text="Verwerken en opslaan", width=22,
                  command=self.verwerken).pack(side="right")

    # ----- dagdelen -----
    def _bouw_dagdelen(self, parent):
        frame = tk.LabelFrame(parent, text="Dagdelen")
        frame.pack(fill="x", pady=(0, 8))
        kop = tk.Frame(frame); kop.pack(fill="x", padx=5, pady=(4, 0))
        tk.Label(kop, text="Van", width=6).grid(row=0, column=0)
        tk.Label(kop, text="Tot", width=6).grid(row=0, column=1)
        tk.Label(kop, text="Dagdeel", width=16).grid(row=0, column=2)
        self.dd_container = tk.Frame(frame); self.dd_container.pack(fill="x", padx=5)
        knoppen = tk.Frame(frame); knoppen.pack(fill="x", padx=5, pady=4)
        tk.Button(knoppen, text="+", width=2, command=lambda: self._dd_add()).pack(side="left")
        tk.Button(knoppen, text="Standaard", command=self._dd_standaard).pack(side="left", padx=4)
        self._dd_herbouw(lees_dagdelen_strings(basis_dir()))

    def _dd_add(self, van="", tot="", naam=""):
        rf = tk.Frame(self.dd_container); rf.pack(fill="x", pady=1)
        e_van = tk.Entry(rf, width=6); e_van.insert(0, van); e_van.grid(row=0, column=0, padx=1)
        e_tot = tk.Entry(rf, width=6); e_tot.insert(0, tot); e_tot.grid(row=0, column=1, padx=1)
        e_naam = tk.Entry(rf, width=16); e_naam.insert(0, naam); e_naam.grid(row=0, column=2, padx=1)
        entry = (e_van, e_tot, e_naam)

        def remove():
            if entry in self.dd_rows:
                self.dd_rows.remove(entry)
            rf.destroy()
        tk.Button(rf, text="x", width=2, command=remove).grid(row=0, column=3, padx=2)
        self.dd_rows.append(entry)

    def _dd_herbouw(self, rijen):
        for child in self.dd_container.winfo_children():
            child.destroy()
        self.dd_rows.clear()
        for van, tot, naam in rijen:
            self._dd_add(van, tot, naam)

    def _dd_standaard(self):
        self._dd_herbouw([tuple(r) for r in DEFAULT_DAGDELEN])

    def _dd_lees(self):
        rijen = []
        for e_van, e_tot, e_naam in self.dd_rows:
            van, tot, naam = e_van.get().strip(), e_tot.get().strip(), e_naam.get().strip()
            if not (van or tot or naam):
                continue
            if _tijd_naar_minuten(van) is None or _tijd_naar_minuten(tot) is None or not naam:
                messagebox.showerror("Ongeldige dagdeel-invoer",
                                     f"Controleer: van='{van}' tot='{tot}' dagdeel='{naam}'")
                return None
            rijen.append((van, tot, naam))
        if not rijen:
            messagebox.showerror("Dagdelen leeg", "Voeg minstens één dagdeel toe.")
            return None
        return rijen

    # ----- dropzones -----
    def _bouw_dropzone(self, parent, titel, hint, file_list, categorie):
        frame = tk.LabelFrame(parent, text=titel)
        frame.pack(fill="both", expand=True, pady=(0, 8))
        tk.Label(frame, text=hint, fg="#555").pack(anchor="w", padx=6, pady=(2, 0))

        listbox = tk.Listbox(frame, width=40, height=6, selectmode="extended")
        listbox.pack(fill="both", expand=True, padx=6, pady=4)

        if HAS_DND:
            listbox.drop_target_register(DND_FILES)
            listbox.dnd_bind("<<Drop>>",
                             lambda e, fl=file_list, lb=listbox: self._op_drop(e, fl, lb))

        btns = tk.Frame(frame); btns.pack(fill="x", padx=6, pady=(0, 4))
        tk.Button(btns, text="Bladeren...",
                  command=lambda fl=file_list, lb=listbox: self._bladeren(fl, lb)).pack(side="left")
        tk.Button(btns, text="Wissen",
                  command=lambda fl=file_list, lb=listbox: self._wissen(fl, lb)).pack(side="left", padx=4)
        return listbox

    def _op_drop(self, event, file_list, listbox):
        paden = self.root.tk.splitlist(event.data)
        self._voeg_toe(paden, file_list, listbox)

    def _bladeren(self, file_list, listbox):
        paden = filedialog.askopenfilenames(
            title="Selecteer bestanden",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("Alle bestanden", "*.*")])
        self._voeg_toe(paden, file_list, listbox)

    def _voeg_toe(self, paden, file_list, listbox):
        toegevoegd = 0
        for p in paden:
            p = str(p)
            if not (p.lower().endswith(".xlsx") or p.lower().endswith(".csv")):
                continue
            if p in file_list:
                continue
            file_list.append(p)
            listbox.insert("end", Path(p).name)
            if p.lower().endswith(".xlsx"):
                info = lees_info_coordinaten(p)
                if info:
                    self.coords[p] = info
            toegevoegd += 1
        if toegevoegd:
            self._ververs_kaart()
            self.status.config(
                text=f"{len(self.bu_files)} gemotoriseerd, {len(self.fiets_files)} fiets — "
                     f"{len(self.coords)} locatie(s) op de kaart.")

    def _wissen(self, file_list, listbox):
        for p in list(file_list):
            self.coords.pop(p, None)
        file_list.clear()
        listbox.delete(0, "end")
        self._ververs_kaart()

    def _ververs_kaart(self):
        if not self.map_widget:
            return
        self.map_widget.delete_all_marker()
        actief = set(self.bu_files) | set(self.fiets_files)
        punten = []
        for p, info in self.coords.items():
            if p not in actief:
                continue
            label = info.get("naam") or Path(p).stem
            if info.get("plaats"):
                label = f"{label} ({info['plaats']})"
            self.map_widget.set_marker(info["lat"], info["lon"], text=label)
            punten.append((info["lat"], info["lon"]))
        if len(punten) == 1:
            self.map_widget.set_position(punten[0][0], punten[0][1]); self.map_widget.set_zoom(13)
        elif len(punten) > 1:
            lats = [a for a, _ in punten]; lons = [b for _, b in punten]
            self.map_widget.fit_bounding_box((max(lats), min(lons)), (min(lats), max(lons)))

    # ----- verwerken -----
    def verwerken(self):
        dd_strings = self._dd_lees()
        if dd_strings is None:
            return
        if not self.bu_files and not self.fiets_files:
            messagebox.showerror("Geen bestanden", "Voeg eerst bestanden toe.")
            return
        schrijf_dagdelen_strings(basis_dir(), dd_strings)
        dagdelen = strings_naar_dagdelen(dd_strings)

        self.status.config(text="Bezig met verwerken..."); self.root.update()

        all_rows = []
        telslang_headers = None
        problemen = []
        for path in self.bu_files:
            typed, headers = verwerk_bestand(path, TELSLANG_VOERTUIGTYPE)
            if headers and telslang_headers is None:
                telslang_headers = headers
            if not typed:
                problemen.append(f"- {Path(path).name} (geen gemotoriseerde data gevonden)")
            all_rows.extend(typed)
        for path in self.fiets_files:
            typed, _ = verwerk_bestand(path, FIETS_VOERTUIGTYPE)
            if not typed:
                problemen.append(f"- {Path(path).name} (geen fietsdata gevonden)")
            all_rows.extend(typed)

        if not all_rows:
            self.status.config(text="Geen data gevonden.")
            messagebox.showerror("Geen data", "In de gekozen bestanden is geen herkenbare data gevonden.")
            return

        out_header, out_rows = build_output(all_rows, telslang_headers, dagdelen)
        n_fiets = sum(1 for d in all_rows if d["Voertuigtype"] == FIETS_VOERTUIGTYPE)
        n_tel = len(all_rows) - n_fiets

        output_file = filedialog.asksaveasfilename(
            title="Opslaan als", initialfile="samengevoegd.xlsx", defaultextension=".xlsx",
            filetypes=[("Excel bestanden", "*.xlsx"), ("CSV bestanden", "*.csv")])
        if not output_file:
            self.status.config(text="Opslaan geannuleerd.")
            return
        try:
            schrijf_uitvoer(output_file, out_header, out_rows)
        except PermissionError:
            messagebox.showerror("Bestand in gebruik",
                                 "Het doelbestand staat open (bv. in Excel). Sluit het en probeer opnieuw.")
            return

        self.status.config(text=f"Klaar: {len(out_rows)} rijen opgeslagen.")
        bericht = (f"{len(out_rows)} rijen geschreven naar:\n{output_file}\n\n"
                   f"({n_fiets} fiets, {n_tel} gemotoriseerd)")
        if problemen:
            bericht += "\n\nLet op, overgeslagen:\n" + "\n".join(problemen)
        messagebox.showinfo("Klaar", bericht)


def main():
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        import traceback
        traceback.print_exc()
        try:
            messagebox.showerror("Fout", f"Er ging iets mis:\n\n{exc}")
        except Exception:
            pass
