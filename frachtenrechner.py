import streamlit as st
import pandas as pd

# Excel-Daten laden
def load_data():
    xls = pd.ExcelFile("0493_Frachtenrechner_KEP_DATA.xlsx")
    country_codes = pd.read_excel(xls, "COUNTRY_CODES")
    zonen_import = pd.read_excel(xls, "Zonen_Import")
    zonen_export = pd.read_excel(xls, "Zonen_Export")
    gewichtsklassen = pd.read_excel(xls, "Gewichtsklassen")
    frachtraten = pd.read_excel(xls, "Frachtraten")
    adds = pd.read_excel(xls, "adds")
    return country_codes, zonen_import, zonen_export, gewichtsklassen, frachtraten, adds

# Gewichtsklasse finden
def finde_gewichtsklasse(gewicht, tarif, gewichtsklassen):
    tarif_col = [col for col in gewichtsklassen.columns if col.lower() == "tarif"]
    if not tarif_col:
        st.error("❌ Spalte 'Tarif' in 'Gewichtsklassen' nicht gefunden. Bitte Excel prüfen.")
        st.stop()
    tarif_col = tarif_col[0]
    df = gewichtsklassen[gewichtsklassen[tarif_col] == tarif]
    zeile = df[(df["von"] <= gewicht) & (df["bis"] >= gewicht)]
    if not zeile.empty:
        return zeile.iloc[0]["GK"]
    return None

# Zone finden
def finde_zone(land, tarif, zonen):
    df = zonen[zonen["LAND"] == land]
    if not df.empty and tarif in df.columns:
        return df.iloc[0][tarif]
    return None

# Rate finden
def finde_rate(tarif, gk, zone, frachtraten):
    df = frachtraten[(frachtraten["TARIF"] == tarif) & (frachtraten["GK"] == gk)]
    if not df.empty and zone in df.columns:
        return df.iloc[0][zone]
    return None

# Zuschlag finden
def finde_zuschlag(tarif, adds):
    df = adds[adds["TARIF"] == tarif]
    if not df.empty:
        return df.iloc[0]["FUELSURCHARGE"]
    return 0.0

# Fracht berechnen (Import/Export)
def berechne_fracht(gewicht, land, tarif, zonen, gewichtsklassen, frachtraten, adds):
    zone = finde_zone(land, tarif, zonen)
    gk = finde_gewichtsklasse(gewicht, tarif, gewichtsklassen)
    rate = finde_rate(tarif, gk, zone, frachtraten)
    zuschlag = finde_zuschlag(tarif, adds)

    if gk and rate is not None:
        kosten = rate * gewicht if gewicht > 20 else rate
        diesel = kosten * zuschlag
        return kosten, diesel, kosten + diesel
    return 0.0, 0.0, 0.0

# Streamlit UI
st.title("KEP Frachtenrechner")

country_codes, zonen_import, zonen_export, gewichtsklassen, frachtraten, adds = load_data()
land_map = dict(zip(country_codes["COUNTRY"], country_codes["LAND"]))

# Block 1: Import
st.header("Importkosten")
import_country = st.selectbox("Import aus:", country_codes["COUNTRY"].unique())
import_gewicht = st.number_input("Gewicht in kg Import", min_value=0.0, step=0.1)
import_tarif = st.selectbox("Serviceart Import", [
    "UPS_I_Express_Plus_pkg1", "UPS_I_Express_env", "UPS_I_Express_doc1", "UPS_I_Express_pgk2",
    "UPS_I_Svr_env", "UPS_I_Svr_doc1", "UPS_I_Svr_pgk2", "UPS_I_Std_Single2",
    "UPS_I_Std_Multi2", "UPS_I_Expideted1"])

import_land = land_map.get(import_country)
import_kosten, import_diesel, import_total = berechne_fracht(
    import_gewicht, import_land, import_tarif,
    zonen_import, gewichtsklassen, frachtraten, adds)

st.write(f"**Frachtkosten Import:** {import_kosten:.2f} €")
st.write(f"**Dieselzuschlag Import:** {import_diesel:.2f} €")
st.write(f"**Total Import:** {import_total:.2f} €")

# Export
st.header("Exportkosten")
export_country = st.selectbox("Export nach:", country_codes["COUNTRY"].unique())
export_gewicht = st.number_input("Gewicht in kg Export", min_value=0.0, step=0.1)
export_tarif = st.selectbox("Serviceart Export", [
    "UPS_E_Express_Plus_env", "UPS_E_Express_Plus_doc1", "UPS_E_Express_Plus_pgk2", "UPS_E_Express_env",
    "UPS_E_Express_doc1", "UPS_E_Express_pgk2", "UPS_E_Express_Noon_env", "UPS_E_Express_Noon_pgk1",
    "UPS_E_Express_Svr_env", "UPS_E_Express_Svr_doc1", "UPS_E_Express_Svr_pgk1",
    "UPS_E_Std_Single2", "UPS_E_Std_Multi2", "UPS_E_Expedited1"])

export_land = land_map.get(export_country)
export_kosten, export_diesel, export_total = berechne_fracht(
    export_gewicht, export_land, export_tarif,
    zonen_export, gewichtsklassen, frachtraten, adds)

st.write(f"**Frachtkosten Export:** {export_kosten:.2f} €")
st.write(f"**Dieselzuschlag Export:** {export_diesel:.2f} €")
st.write(f"**Total Export:** {export_total:.2f} €")

# Block 3: Jobkosten
st.header("Jobkosten")
if import_gewicht > 0:
    anteil_import = import_total * (export_gewicht / import_gewicht)
else:
    anteil_import = 0.0

total_job = anteil_import + export_total
st.write(f"**Anteilige Importkosten:** {anteil_import:.2f} €")
st.write(f"**Exportkosten:** {export_total:.2f} €")
st.write(f"**Jobkosten Total:** {total_job:.2f} €")
