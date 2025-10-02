import streamlit as st
import pandas as pd
import io
import re

# App-Konfiguration
st.set_page_config(page_title="ISE Importer", layout="centered", page_icon="📊")

# Konstanten
COLUMN_NAMES = [
    "MACAddress","EndPointPolicy","IdentityGroup","PortalUser.GuestType","Description",
    "PortalUser.Location","PortalUser.GuestStatus","StaticAssignment","User-Name",
    "DeviceRegistrationStatus","PortalUser.CreationType","AUPAccepted","PortalUser.EmailAddress",
    "PortalUser.PhoneNumber","FirstName","ip","Device Type","host-name","StaticGroupAssignment",
    "MDMEnrolled","MDMOSVersion","PortalUser.LastName","PortalUser.GuestSponsor","EmailAddress",
    "PortalUser","PortalUser.FirstName","BYODRegistration","MDMServerName","LastName",
    "MDMServerID","Location"
]  # Länge = 31

def read_excel_safe(file):
    try:
        df = pd.read_excel(file, header=None, dtype=str, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Excel konnte nicht gelesen werden: {e}")
    if df.empty:
        raise ValueError("Die hochgeladene Datei ist leer.")
    if df.shape[1] < 4:
        raise ValueError("Zu wenige Spalten erkannt. Benötigt werden mindestens 4 Spalten: MAC, ISE MAC Gruppe, Beschreibung, Standort.")
    return df.iloc[:, :4]

def validate_mac_column(df):
    for idx, cell in enumerate(df.iloc[:, 0], start=1):
        val = str(cell).strip() if pd.notna(cell) else ""
        if val == "" or val.lower() == "nan":
            raise ValueError(f"Fehler in Zeile {idx}: MAC-Adresse ist leer.")

def count_commas(df):
    return df.iloc[:, :4].astype(str).apply(lambda col: col.str.count(",")).sum().sum()

def remove_commas(df):
    return df.iloc[:, :4].astype(str).applymap(lambda v: v.replace(",", ""))

def build_csv_text(df, include_description, remove_commas_flag):
    if remove_commas_flag:
        df = remove_commas(df)
    rows = []
    for _, row in df.iterrows():
        mac = row[0].strip()
        group = row[1].strip() if pd.notna(row[1]) else ""
        desc = row[2].strip() if pd.notna(row[2]) else ""
        loc = row[3].strip() if pd.notna(row[3]) else ""
        # CSV-Struktur: 31 Spalten
        vals = [mac, "", ""]        # 1–3
        vals.append(group)          # 4
        vals += ["", ""]            # 5–6
        vals.append(desc if include_description else "")  # 7
        vals += [""] * 23           # 8–30 (23 leere Felder)
        vals.append(loc)            # 31
        rows.append(vals)
    buf = io.StringIO()
    pd.DataFrame(rows, columns=COLUMN_NAMES).to_csv(buf, index=False)
    return buf.getvalue()

def check_text_mac_in_csv(csv_text):
    lines = csv_text.splitlines()
    for i, line in enumerate(lines[1:], start=2):
        if re.search(r"\bMAC\b", line, re.IGNORECASE):
            return i
    return None

def main():
    st.title("🔧 ISE Importer")
    st.markdown(
        "**Bitte Excel-Tabelle hochladen:**  \n"
        "`MAC Adresse | ISE MAC Gruppe | Beschreibung | Standort`  \n"
        "Leere MAC-Adressen (Spalte A) werden validiert und blockiert."
    )

    uploaded = st.file_uploader("Excel-Datei hochladen (.xlsx, .xls)", type=["xlsx", "xls"])
    if not uploaded:
        st.info("👆 Bitte eine Excel-Datei auswählen.")
        return

    try:
        df = read_excel_safe(uploaded)
        validate_mac_column(df)

        comma_count = count_commas(df)
        remove_commas_flag = False
        if comma_count > 0:
            st.warning(f"⚠️ {int(comma_count)} Kommas in den Daten entdeckt.")
            remove_commas_flag = st.checkbox("Kommas automatisch entfernen", value=True)

        include_description = st.checkbox("Beschreibung in CSV-Export einschließen", value=True)

        csv_text = build_csv_text(df, include_description, remove_commas_flag)

        bad_line = check_text_mac_in_csv(csv_text)
        if bad_line:
            st.error(f"❌ Wort 'MAC' in Zeile {bad_line} der CSV entdeckt. Bitte korrigieren.")

        st.success("✅ Datei erfolgreich verarbeitet!")
        st.subheader("📋 CSV-Vorschau (erste 20 Zeilen)")
        preview = "\n".join(csv_text.splitlines()[:21])
        st.text(preview)

        st.download_button(
            "📥 CSV herunterladen",
            data=csv_text,
            file_name="ise_import.csv",
            mime="text/csv"
        )

    except ValueError as ve:
        st.error(f"❌ {ve}")
    except Exception as ex:
        st.error(f"💥 Unerwarteter Fehler: {ex}")

if __name__ == "__main__":
    main()
