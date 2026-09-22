# -*- coding: utf-8 -*-
"""
Vergelijk de Intune-app-inventaris voor de periodieke IMS-check.

Legt drie lijsten naast elkaar en levert een werkbestand op waarin alleen nog
staat wat beoordeeld moet worden:

  1. Gedetecteerde apps (Intune > Apps > Monitor > Gedetecteerde apps > Exporteren)
     Dit is wat er feitelijk op de toestellen staat, inclusief wat medewerkers
     zelf hebben geinstalleerd.
  2. De vorige export van diezelfde lijst (optioneel). Hiermee wordt zichtbaar
     wat er sinds de vorige ronde is bijgekomen of verdwenen; dat verschil is
     het eigenlijke controlewerk.
  3. De lijst beheerde apps (Intune > Apps > Alle apps > Exporteren, optioneel).
     Alles wat daar niet op staat is niet door ons uitgerold.

De uitvoer is een xlsx met een tabblad 'Beoordelen' waarin de apps staan die
aandacht vragen: niet van Microsoft, niet beheerd, of nieuw sinds de vorige
ronde. Gesorteerd op aantal apparaten, want software op een enkel toestel is
interessanter dan software die overal staat.

Flow:
  1. Vraag via bestandsdialogen om de drie bestanden (of geef ze mee als argument).
  2. Koppel de lijsten op genormaliseerde applicatienaam.
  3. Schrijf het werkbestand weg voor bij de Monday-taak.
"""
import argparse
import csv
import os
import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import filedialog, messagebox

# ---------------------------------------------------------------- config
# Intune wisselt nogal eens van kolomnaam; de eerste die past wordt gebruikt.
KOL_NAAM = ["ApplicationName", "Application name", "AppName", "App name", "Name",
            "Applicatienaam", "Naam"]
KOL_VERSIE = ["ApplicationVersion", "ApplicationShortVersion", "Version", "AppVersion", "Versie"]
KOL_UITGEVER = ["ApplicationPublisher", "Publisher", "AppPublisher", "Uitgever"]
KOL_APPARATEN = ["Device count", "DeviceCount", "Devices", "Aantal apparaten"]
KOL_PLATFORM = ["Platform", "OS"]
# Staat er een kolom met de apparaatnaam in, dan is het de ruwe export met een
# regel per apparaat-app-combinatie (AppInvRawData) en wordt er geaggregeerd.
KOL_APPARAAT = ["DeviceName", "Device name", "Apparaatnaam"]
KOL_GEBRUIKER = ["UserName", "User name", "Gebruiker", "Gebruikersnaam"]

# Boven dit aantal apparaten zeggen losse namen niets meer; dan volstaat het aantal.
MAX_APPARAATNAMEN = 8

MICROSOFT_PATROON = re.compile(r"microsoft|windows", re.IGNORECASE)

KOP_VULLING = PatternFill("solid", fgColor="DDEBF7")
LET_OP_VULLING = PatternFill("solid", fgColor="FCE4D6")
KOP_FONT = Font(bold=True)


# ---------------------------------------------------------------- inlezen

def vind_kolom(velden, kandidaten):
    """Geef de eerste kolomnaam uit kandidaten die in velden voorkomt."""
    genormaliseerd = {v.strip().lower(): v for v in velden if v}
    for kandidaat in kandidaten:
        if kandidaat.strip().lower() in genormaliseerd:
            return genormaliseerd[kandidaat.strip().lower()]
    return None


def lees_csv(pad):
    """Lees een Intune-export; scheidingsteken en BOM worden automatisch bepaald."""
    with open(pad, "r", encoding="utf-8-sig", newline="") as bestand:
        monster = bestand.read(8192)
        bestand.seek(0)
        try:
            dialect = csv.Sniffer().sniff(monster, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        return list(csv.DictReader(bestand, dialect=dialect))


def normaliseer(naam):
    """Maak een applicatienaam vergelijkbaar tussen de verschillende exports."""
    if not naam:
        return ""
    tekst = naam.lower()
    tekst = re.sub(r"\(x64\)|\(x86\)|\(64-bit\)|\(32-bit\)", " ", tekst)
    tekst = re.sub(r"\b\d+(\.\d+)+\b", " ", tekst)   # losse versienummers
    tekst = re.sub(r"[^a-z0-9]+", " ", tekst)
    return " ".join(tekst.split())


def aggregeer_ruwe_export(rijen, kol_naam, kol_versie, kol_uitgever, kol_platform,
                          kol_apparaat, kol_gebruiker):
    """Vat de ruwe export (een regel per apparaat) samen tot een regel per app."""
    verzameld = {}
    for rij in rijen:
        naam = (rij.get(kol_naam) or "").strip()
        if not naam:
            continue
        versie = (rij.get(kol_versie) or "").strip() if kol_versie else ""
        sleutel = (normaliseer(naam), versie)

        app = verzameld.get(sleutel)
        if app is None:
            app = {
                "naam": naam,
                "sleutel": sleutel[0],
                "versie": versie,
                "uitgever": (rij.get(kol_uitgever) or "").strip() if kol_uitgever else "",
                "platform": (rij.get(kol_platform) or "").strip() if kol_platform else "",
                "apparaten": 0,
                "_apparaatnamen": set(),
                "_gebruikers": set(),
            }
            verzameld[sleutel] = app

        apparaat = (rij.get(kol_apparaat) or "").strip()
        if apparaat:
            app["_apparaatnamen"].add(apparaat)
        if kol_gebruiker:
            gebruiker = (rij.get(kol_gebruiker) or "").strip()
            if gebruiker:
                app["_gebruikers"].add(gebruiker)

    apps = []
    for app in verzameld.values():
        namen = sorted(app.pop("_apparaatnamen"))
        gebruikers = sorted(app.pop("_gebruikers"))
        app["apparaten"] = len(namen)
        app["apparatenlijst"] = (
            ", ".join(namen) if len(namen) <= MAX_APPARAATNAMEN
            else "{} apparaten".format(len(namen))
        )
        app["gebruikerslijst"] = (
            ", ".join(gebruikers) if len(gebruikers) <= MAX_APPARAATNAMEN
            else "{} gebruikers".format(len(gebruikers))
        )
        apps.append(app)
    return apps


def lees_apps(pad, label):
    """Lees een app-export en geef een lijst dicts met vaste sleutels terug."""
    rijen = lees_csv(pad)
    if not rijen:
        raise ValueError("{} bevat geen rijen: {}".format(label, pad))

    velden = list(rijen[0].keys())
    kol_naam = vind_kolom(velden, KOL_NAAM)
    if not kol_naam:
        raise ValueError(
            "Geen kolom met de applicatienaam gevonden in {}.\n"
            "Aanwezige kolommen: {}".format(label, ", ".join(velden))
        )
    kol_versie = vind_kolom(velden, KOL_VERSIE)
    kol_uitgever = vind_kolom(velden, KOL_UITGEVER)
    kol_apparaten = vind_kolom(velden, KOL_APPARATEN)
    kol_platform = vind_kolom(velden, KOL_PLATFORM)
    kol_apparaat = vind_kolom(velden, KOL_APPARAAT)
    kol_gebruiker = vind_kolom(velden, KOL_GEBRUIKER)

    if kol_apparaat:
        return aggregeer_ruwe_export(rijen, kol_naam, kol_versie, kol_uitgever,
                                     kol_platform, kol_apparaat, kol_gebruiker)

    apps = []
    for rij in rijen:
        naam = (rij.get(kol_naam) or "").strip()
        if not naam:
            continue
        try:
            apparaten = int(float(rij.get(kol_apparaten) or 0)) if kol_apparaten else 0
        except ValueError:
            apparaten = 0
        apps.append({
            "naam": naam,
            "sleutel": normaliseer(naam),
            "versie": (rij.get(kol_versie) or "").strip() if kol_versie else "",
            "uitgever": (rij.get(kol_uitgever) or "").strip() if kol_uitgever else "",
            "platform": (rij.get(kol_platform) or "").strip() if kol_platform else "",
            "apparaten": apparaten,
        })
    return apps


# ---------------------------------------------------------------- vergelijken

def is_microsoft(app):
    """Bepaal of de app van Microsoft is; uitgever gaat voor op de naam."""
    if app["uitgever"]:
        return bool(MICROSOFT_PATROON.search(app["uitgever"]))
    return bool(MICROSOFT_PATROON.search(app["naam"]))


def is_beheerd(app, beheerde_sleutels):
    """Koppel losjes: exacte sleutel, of de beheerde naam zit in de gevonden naam."""
    if not beheerde_sleutels:
        return None
    if app["sleutel"] in beheerde_sleutels:
        return True
    for sleutel in beheerde_sleutels:
        if len(sleutel) >= 4 and (sleutel in app["sleutel"] or app["sleutel"] in sleutel):
            return True
    return False


def bepaal_status(app, vorige_namen, vorige_versies):
    """Nieuw, nieuwe versie of bekend ten opzichte van de vorige ronde."""
    if not vorige_namen:
        return ""
    if app["sleutel"] not in vorige_namen:
        return "nieuwe app"
    if (app["sleutel"], app["versie"]) not in vorige_versies:
        return "nieuwe versie"
    return "bekend"


def verrijk(huidig, vorige, beheerd):
    """Voeg per app de oordeelsvelden toe."""
    beheerde_sleutels = {app["sleutel"] for app in beheerd} if beheerd else set()
    vorige_namen = {app["sleutel"] for app in vorige} if vorige else set()
    vorige_versies = {(app["sleutel"], app["versie"]) for app in vorige} if vorige else set()

    for app in huidig:
        app["microsoft"] = is_microsoft(app)
        app["beheerd"] = is_beheerd(app, beheerde_sleutels)
        app["status"] = bepaal_status(app, vorige_namen, vorige_versies)
        # Twee filters: alles met Microsoft als uitgever eruit, en alles wat
        # via Intune is uitgerold eruit. Wat overblijft is door de gebruiker
        # zelf geinstalleerde software van derden. Nieuwe Microsoft- of
        # beheerde apps blijven zichtbaar op het tabblad 'Nieuw'.
        app["beoordelen"] = not app["microsoft"] and app["beheerd"] is not True
    return huidig


def verdwenen_apps(huidig, vorige):
    """Apps die in de vorige export stonden en nu niet meer."""
    if not vorige:
        return []
    huidige_sleutels = {app["sleutel"] for app in huidig}
    gezien = set()
    verdwenen = []
    for app in vorige:
        if app["sleutel"] not in huidige_sleutels and app["sleutel"] not in gezien:
            gezien.add(app["sleutel"])
            app["microsoft"] = is_microsoft(app)
            app["beheerd"] = None      # niet meer aangetroffen, dus niet te bepalen
            app["status"] = "verdwenen"
            verdwenen.append(app)
    return verdwenen


# ---------------------------------------------------------------- uitvoer

KOLOMMEN = [
    ("Applicatie", "naam", 46),
    ("Versie", "versie", 16),
    ("Uitgever", "uitgever", 30),
    ("Apparaten", "apparaten", 11),
    ("Apparaatnamen", "apparatenlijst", 42),
    ("Gebruiker(s)", "gebruikerslijst", 34),
    ("Microsoft", "microsoft", 11),
    ("Beheerd", "beheerd", 10),
    ("T.o.v. vorige ronde", "status", 20),
]


def waarde_voor_cel(app, sleutel):
    waarde = app.get(sleutel, "")
    if sleutel == "microsoft":
        return "ja" if waarde else "nee"
    if sleutel == "beheerd":
        if waarde is None:
            return ""
        return "ja" if waarde else "nee"
    return waarde


def schrijf_blad(werkboek, titel, apps, toelichting=""):
    blad = werkboek.create_sheet(titel)
    rij_nr = 1

    if toelichting:
        blad.cell(row=1, column=1, value=toelichting).font = Font(italic=True)
        rij_nr = 3

    for kolom_nr, (kop, _, breedte) in enumerate(KOLOMMEN, start=1):
        cel = blad.cell(row=rij_nr, column=kolom_nr, value=kop)
        cel.font = KOP_FONT
        cel.fill = KOP_VULLING
        cel.alignment = Alignment(vertical="center")
        blad.column_dimensions[get_column_letter(kolom_nr)].width = breedte

    kop_rij = rij_nr
    for app in apps:
        rij_nr += 1
        for kolom_nr, (_, sleutel, _) in enumerate(KOLOMMEN, start=1):
            cel = blad.cell(row=rij_nr, column=kolom_nr, value=waarde_voor_cel(app, sleutel))
            if app.get("status") == "nieuwe app" or app.get("beheerd") is False:
                cel.fill = LET_OP_VULLING

    blad.freeze_panes = blad.cell(row=kop_rij + 1, column=1)
    blad.auto_filter.ref = "A{}:{}{}".format(
        kop_rij, get_column_letter(len(KOLOMMEN)), max(rij_nr, kop_rij)
    )
    return blad


def schrijf_samenvatting(werkboek, apps, verdwenen, bronnen):
    blad = werkboek.create_sheet("Samenvatting", 0)
    te_beoordelen = [a for a in apps if a["beoordelen"]]

    regels = [
        ("Vergelijking Intune-app-inventaris", ""),
        ("", ""),
        ("Datum", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("", ""),
        ("Gedetecteerde apps (huidig)", os.path.basename(bronnen["huidig"])),
        ("Vorige export", os.path.basename(bronnen["vorige"]) if bronnen["vorige"] else "niet meegegeven"),
        ("Beheerde apps", os.path.basename(bronnen["beheerd"]) if bronnen["beheerd"] else "niet meegegeven"),
        ("", ""),
        ("Totaal aangetroffen apps", len(apps)),
        ("Afgevallen: Microsoft als uitgever", sum(1 for a in apps if a["microsoft"])),
        ("Afgevallen: via Intune uitgerold", sum(1 for a in apps if not a["microsoft"] and a["beheerd"] is True)),
        ("Nieuw sinds vorige ronde", sum(1 for a in apps if a["status"] == "nieuwe app")),
        ("Nieuwe versie van bekende app", sum(1 for a in apps if a["status"] == "nieuwe versie")),
        ("Verdwenen sinds vorige ronde", len(verdwenen)),
        ("", ""),
        ("Te beoordelen op tabblad 'Beoordelen'", len(te_beoordelen)),
        ("", ""),
        ("Let op: de koppeling met beheerde apps is een naamvergelijking en", ""),
        ("daarmee indicatief. Controleer 'nee' in de kolom Beheerd handmatig.", ""),
    ]

    for rij_nr, (label, waarde) in enumerate(regels, start=1):
        cel = blad.cell(row=rij_nr, column=1, value=label)
        if rij_nr == 1:
            cel.font = Font(bold=True, size=13)
        elif waarde != "":
            cel.font = KOP_FONT
        blad.cell(row=rij_nr, column=2, value=waarde)

    blad.column_dimensions["A"].width = 46
    blad.column_dimensions["B"].width = 40
    return blad


def bouw_werkbestand(apps, verdwenen, bronnen, doel):
    werkboek = Workbook()
    werkboek.remove(werkboek.active)

    op_apparaten = sorted(apps, key=lambda a: (a["apparaten"], a["naam"].lower()))

    schrijf_blad(
        werkboek, "Beoordelen",
        [a for a in op_apparaten if a["beoordelen"]],
        "Software van derden die niet via Intune is uitgerold: Microsoft als uitgever "
        "en de beheerde apps zijn eruit gefilterd. Gesorteerd op aantal apparaten, "
        "dus bovenaan staat wat op de minste toestellen voorkomt.",
    )
    schrijf_blad(
        werkboek, "Nieuw",
        [a for a in op_apparaten if a["status"] in ("nieuwe app", "nieuwe versie")],
        "Bijgekomen sinds de vorige export.",
    )
    schrijf_blad(
        werkboek, "Verdwenen", verdwenen,
        "Stond in de vorige export en is nu niet meer aangetroffen.",
    )
    schrijf_blad(
        werkboek, "Alle apps", op_apparaten,
        "De volledige huidige inventaris.",
    )
    schrijf_samenvatting(werkboek, apps, verdwenen, bronnen)

    werkboek.save(doel)
    return doel


# ---------------------------------------------------------------- aansturing

def vraag_bestanden():
    """Vraag de drie exports op via bestandsdialogen."""
    venster = tk.Tk()
    venster.withdraw()

    huidig = filedialog.askopenfilename(
        title="Gedetecteerde apps - huidige export",
        filetypes=[("CSV-bestanden", "*.csv"), ("Alle bestanden", "*.*")],
    )
    if not huidig:
        return None, None, None

    messagebox.showinfo(
        "Vorige export",
        "Kies nu de vorige export van Gedetecteerde apps.\n\n"
        "Annuleren mag: dan wordt er geen vergelijking met de vorige ronde gemaakt.",
    )
    vorige = filedialog.askopenfilename(
        title="Gedetecteerde apps - vorige export (optioneel)",
        filetypes=[("CSV-bestanden", "*.csv"), ("Alle bestanden", "*.*")],
    )

    messagebox.showinfo(
        "Beheerde apps",
        "Kies nu de export van Alle apps (de door ons beheerde uitrol).\n\n"
        "Annuleren mag: dan blijft de kolom Beheerd leeg.",
    )
    beheerd = filedialog.askopenfilename(
        title="Alle apps - beheerde uitrol (optioneel)",
        filetypes=[("CSV-bestanden", "*.csv"), ("Alle bestanden", "*.*")],
    )

    venster.destroy()
    return huidig, vorige or None, beheerd or None


def main():
    parser = argparse.ArgumentParser(
        description="Vergelijk Intune-app-exports voor de periodieke IMS-check."
    )
    parser.add_argument("--huidig", help="CSV: huidige export Gedetecteerde apps")
    parser.add_argument("--vorige", help="CSV: vorige export Gedetecteerde apps")
    parser.add_argument("--beheerd", help="CSV: export Alle apps (beheerde uitrol)")
    parser.add_argument("--uitvoer", help="Pad voor het xlsx-werkbestand")
    argumenten = parser.parse_args()

    huidig_pad = argumenten.huidig
    vorige_pad = argumenten.vorige
    beheerd_pad = argumenten.beheerd

    if not huidig_pad:
        huidig_pad, vorige_pad, beheerd_pad = vraag_bestanden()
        if not huidig_pad:
            print("Geen bestand gekozen; gestopt.")
            return

    huidig = lees_apps(huidig_pad, "de huidige export")
    vorige = lees_apps(vorige_pad, "de vorige export") if vorige_pad else []
    beheerd = lees_apps(beheerd_pad, "de lijst beheerde apps") if beheerd_pad else []

    apps = verrijk(huidig, vorige, beheerd)
    verdwenen = verdwenen_apps(huidig, vorige)

    doel = argumenten.uitvoer
    if not doel:
        doel = os.path.join(
            os.path.dirname(os.path.abspath(huidig_pad)),
            "Appvergelijking_{}.xlsx".format(datetime.now().strftime("%Y-%m-%d")),
        )

    bouw_werkbestand(apps, verdwenen, {
        "huidig": huidig_pad,
        "vorige": vorige_pad or "",
        "beheerd": beheerd_pad or "",
    }, doel)

    te_beoordelen = sum(1 for a in apps if a["beoordelen"])
    print("Apps aangetroffen      : {}".format(len(apps)))
    print("Niet van Microsoft     : {}".format(sum(1 for a in apps if not a["microsoft"])))
    print("Niet beheerd           : {}".format(sum(1 for a in apps if a["beheerd"] is False)))
    print("Nieuw sinds vorige     : {}".format(sum(1 for a in apps if a["status"] == "nieuwe app")))
    print("Verdwenen sinds vorige : {}".format(len(verdwenen)))
    print("Te beoordelen          : {}".format(te_beoordelen))
    print("\nWerkbestand: {}".format(doel))


if __name__ == "__main__":
    main()
