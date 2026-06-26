import csv
import re
import tkinter as tk
from tkinter import filedialog, Listbox, Toplevel, MULTIPLE, END
from pathlib import Path
from datetime import datetime, timedelta
from openpyxl import load_workbook, Workbook


# Voertuigtype-labels in de uitvoer.
# Fietsdata (weektabellen) krijgt "fiets"; telslang/moto-data (dagtabellen) dit label.
# Pas TELSLANG_VOERTUIGTYPE gerust aan als je een andere benaming wilt.
FIETS_VOERTUIGTYPE = "fiets"
TELSLANG_VOERTUIGTYPE = "motorvoertuig"


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


def _cell0(row):
    """Eerste cel van een rij als gestripte string ('' als leeg)."""
    if row and len(row) > 0 and row[0] is not None:
        return str(row[0]).strip()
    return ""


def parse_rows(rows):
    """Parse één tabel (CSV of xlsx-sheet) en herken per blok automatisch
    of het een fiets-weektabel of een telslang-dagtabel is.

    Geeft terug: (lijst van rij-dicts, telslang_klassekolommen of None).
    Elke rij-dict bevat: Voertuigtype, Datum, Tijd, Richting, classes, V85, Gem., Totaal.
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

                # kolom-headerrij van dit blok staat direct onder de tabelkop
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

                # Uitklappen: per dag, dan per uur
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


def build_output(all_rows, telslang_headers):
    """Bouw de kopregel en de uitvoerrijen. Eén tabblad met alle data."""
    columns = ["Locatie", "Voertuigtype", "Datum", "Tijd", "Richting"]
    if telslang_headers:
        columns += telslang_headers + ["V85", "Gem."]
    columns += ["Totaal"]

    out_rows = []
    for d in all_rows:
        row = [d["Locatie"], d["Voertuigtype"], d["Datum"], d["Tijd"], d["Richting"]]
        if telslang_headers:
            for h in telslang_headers:
                row.append(d["classes"].get(h, ""))
            row.append(d["V85"])
            row.append(d["Gem."])
        row.append(d["Totaal"])
        out_rows.append(row)

    return columns, out_rows


def main():
    root = tk.Tk()
    root.withdraw()

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

    out_header, out_rows = build_output(all_rows, telslang_headers)

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
        # Formatteer datumkolom als datum
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
