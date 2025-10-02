import streamlit as st
import pandas as pd
import io
import re

# App-Konfiguration
st.set_page_config(page_title="ISE Importer", layout="centered", page_icon="📊")

# Spaltenüberschriften (31 Spalten)
COLUMN_NAMES = [
    "MACAddress","EndPointPolicy","IdentityGroup","PortalUser.GuestType","Description",
    "PortalUser.Location","PortalUser.GuestStatus","StaticAssignment","User-Name",
    "DeviceRegistrationStatus","PortalUser.CreationType","AUPAccepted","PortalUser.EmailAddress",
    "PortalUser.PhoneNumber","FirstName","ip","Device Type","host-name","StaticGroupAssignment",
    "MDMEnrolled","MDMOSVersion","PortalUser.LastName","PortalUser.GuestSponsor","EmailAddress",
    "PortalUser","PortalUser.FirstName","BYODRegistration","MDMServerName","LastName",
    "MDMServerID","Location"
]

def read_excel_safe(file):
    try:
        df = pd.read_excel(file, header=None, dtype=str, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Excel konnte nicht gelesen werden: {e}")
    if df.empty:
        raise ValueError("Die hochgeladene Datei ist leer.")
    if df.shape[1] < 4:
        raise ValueError("Zu wenige Spalten erkannt. Mindestens MAC, ISE MAC Gruppe, Beschreibung, Standort erforderlich.")
    return df.iloc[:, :4]

def validate_mac_column(df):
    for idx, mac in enumerate(df.iloc[:,0], start=1):
        val = str(mac).strip() if pd.notna(mac) else ""
        if not val or val.lower()=="nan":
            raise ValueError(f"Fehler in Zeile {idx}: MAC-Adresse ist leer.")

def count_commas(df):
    return df.iloc[:, :4].astype(str).apply(lambda col: col.str.count(",")).sum().sum()

def remove_commas(df):
    return df.iloc[:, :4].astype(str).applymap(lambda v: v.replace(",", ""))

def build_csv_text(df, include_desc, remove_commas_flag):
    if remove_commas_flag:
        df = remove_commas(df)
    lines = []
    for _, row in df.iterrows():
        mac  = (row[0] or "").strip()
        func = (row[1] or "").strip()
        desc = (row[2] or "").strip()
        loc  = (row[3] or "").strip()
        # Neues Layout:
        # MAC, "", Function, "", Description, 26x "", Location
        parts = [mac, ""]            # MAC und eine leere
        parts.append(func)           # Function
        parts.append("")             # eine leere
        parts.append(desc if include_desc else "")  # Description oder leer
        parts += [""] * 25           # 26 leere Felder
        parts.append(loc)            # Standort
        lines.append(",".join(parts))
    header = ",".join(COLUMN_NAMES)
    return header + "\n" + "\n".join(lines)

def find_text_mac(csv_text):
    for i, line in enumerate(csv_text.splitlines()[1:], start=2):
        if re.search(r"\bMAC\b", line, re.IGNORECASE):
            return i
    return None

def main():
    st.title("🔧 ISE Importer")
    st.markdown(
        "**Excel hochladen:**  \n"
        "`MAC Adresse | ISE MAC Gruppe | Beschreibung | Standort`  \n"
        "Leere MAC-Adressen (Spalte A) werden blockiert."
    )
    uploaded = st.file_uploader("Excel-Datei (.xlsx, .xls)", type=["xlsx","xls"])
    if not uploaded:
        st.info("👆 Datei auswählen")
        return

    try:
        df = read_excel_safe(uploaded)
        validate_mac_column(df)

        comma_count = count_commas(df)
        remove_commas_flag = False
        if comma_count>0:
            st.warning(f"⚠️ {int(comma_count)} Kommas gefunden.")
            remove_commas_flag = st.checkbox("Kommas automatisch entfernen", value=True)

        include_desc = st.checkbox("Beschreibung einfügen", value=True)

        csv_text = build_csv_text(df, include_desc, remove_commas_flag)

        bad = find_text_mac(csv_text)
        if bad:
            st.error(f"❌ 'MAC' in Zeile {bad} gefunden.")

        st.success("✅ Verarbeitet!")
        st.subheader("📋 Vorschau (20 Zeilen)")
        preview = "\n".join(csv_text.splitlines()[:21])
        st.text(preview)

        st.download_button("📥 CSV herunterladen", data=csv_text,
                           file_name="ise_import.csv", mime="text/csv")

    except ValueError as ve:
        st.error(f"❌ {ve}")
    except Exception as ex:
        st.error(f"💥 Fehler: {ex}")

if __name__=="__main__":
    main()
