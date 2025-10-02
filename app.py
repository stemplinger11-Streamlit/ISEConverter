import streamlit as st
import pandas as pd
import io

# Konfiguration
st.set_page_config(
    page_title="ISE Importer", 
    layout="centered",
    page_icon="📊"
)

# Konstanten
COLUMN_NAMES = [
    "MACAddress","EndPointPolicy","IdentityGroup","PortalUser.GuestType","Description",
    "PortalUser.Location","PortalUser.GuestStatus","StaticAssignment","User-Name",
    "DeviceRegistrationStatus","PortalUser.CreationType","AUPAccepted","PortalUser.EmailAddress",
    "PortalUser.PhoneNumber","FirstName","ip","Device Type","host-name","StaticGroupAssignment",
    "MDMEnrolled","MDMOSVersion","PortalUser.LastName","PortalUser.GuestSponsor","EmailAddress",
    "PortalUser","PortalUser.FirstName","BYODRegistration","MDMServerName","LastName",
    "MDMServerID","Location"
]

def escape_commas_if_needed(value, comma_handling):
    """Fügt Anführungszeichen um Werte mit Kommas hinzu"""
    if comma_handling and isinstance(value, str) and ',' in value:
        return f'"{value}"'
    return value

def validate_excel_data(df):
    """Validiert die Excel-Daten"""
    if df.empty:
        raise ValueError("Die hochgeladene Datei ist leer.")
    
    if df.shape < 4:
        raise ValueError("Zu wenige Spalten erkannt. Benötigt werden mindestens 4 Spalten: MAC, ISE MAC Gruppe, Beschreibung, Standort.")
    
    # Prüfe MAC-Adressen
    for idx, row in df.iterrows():
        mac = str(row).strip() if pd.notna(row) else ""
        if mac == "" or mac.lower() == "nan":
            raise ValueError(f"Fehler in Zeile {idx+1}: MAC-Adresse ist leer.")

def count_commas_in_data(df, comma_handling):
    """Zählt Zellen mit Kommas"""
    if not comma_handling:
        return 0
    
    comma_count = 0
    for col in df.columns[:4]:  # Nur die ersten 4 Spalten prüfen
        comma_count += df[col].astype(str).str.contains(',', na=False).sum()
    return comma_count

def build_csv_rows(df, include_description, comma_handling):
    """Erstellt CSV-Zeilen nach ISE-Spezifikation"""
    rows = []
    
    for idx, row in df.iterrows():
        mac = str(row).strip() if pd.notna(row) else ""
        group = str(row).strip() if pd.notna(row) else ""
        desc = str(row).strip() if pd.notna(row) else ""
        loc = str(row).strip() if pd.notna(row) else ""
        
        # Escape Kommas falls aktiviert
        if comma_handling:
            mac = escape_commas_if_needed(mac, True)
            group = escape_commas_if_needed(group, True)
            desc = escape_commas_if_needed(desc, True)
            loc = escape_commas_if_needed(loc, True)
        
        # CSV-Struktur: MAC, 2 leere, ISE-Gruppe, 2 leere, Beschreibung (optional), 26 leere, Standort
        values = [mac, "", ""]  # Position 1-3
        values.append(group)    # Position 4
        values.extend(["", ""])  # Position 5-6
        
        if include_description:
            values.append(desc)  # Position 7
        else:
            values.append("")    # Position 7 leer
            
        values.extend([""] * 26)  # Position 8-33 (26 leere Spalten)
        values.append(loc)        # Position 34
        
        rows.append(values)
    
    return rows

def read_excel_safe(file):
    """Liest Excel-Datei sicher ein"""
    try:
        df = pd.read_excel(file, header=None, dtype=str)
        return df.iloc[:, :4]  # Nur erste 4 Spalten
    except Exception as e:
        raise ValueError(f"Excel konnte nicht gelesen werden: {str(e)}")

def main():
    # Header
    st.title("🔧 ISE Importer")
    st.markdown("""
    **Bitte laden Sie eine Excel-Tabelle mit Spalten in dieser Reihenfolge hoch:**  
    `MAC Adresse | ISE MAC Gruppe | Beschreibung | Standort`
    
    ⚠️ Nur leere MAC-Adressen (Spalte A) werden validiert und blockiert.
    """)
    
    # Upload
    uploaded_file = st.file_uploader(
        "Excel-Datei hochladen",
        type=["xlsx", "xls"],
        help="Unterstützte Formate: .xlsx, .xls"
    )
    
    if uploaded_file is not None:
        # Optionen
        col1, col2 = st.columns(2)
        with col1:
            include_description = st.checkbox(
                "✅ Beschreibung in CSV-Export einschließen", 
                value=True
            )
        with col2:
            comma_handling = st.checkbox(
                "🔤 Erweiterte Komma-Behandlung aktivieren", 
                value=True
            )
        
        try:
            # Excel einlesen und validieren
            df = read_excel_safe(uploaded_file)
            validate_excel_data(df)
            
            # Kommas zählen
            comma_count = count_commas_in_data(df, comma_handling)
            
            # CSV generieren
            csv_rows = build_csv_rows(df, include_description, comma_handling)
            csv_df = pd.DataFrame(csv_rows, columns=COLUMN_NAMES)
            
            # CSV zu String
            csv_buffer = io.StringIO()
            csv_df.to_csv(csv_buffer, index=False)
            csv_content = csv_buffer.getvalue()
            
            # Erfolg anzeigen
            st.success("✅ Datei erfolgreich verarbeitet!")
            
            # Komma-Warnung
            if comma_count > 0:
                st.info(f"ℹ️ {comma_count} Zellen mit Kommas erkannt und automatisch mit Anführungszeichen versehen.")
            
            # Voransicht
            st.subheader("📋 CSV-Voransicht (erste 20 Zeilen)")
            preview_lines = csv_content.split('\n')[:21]  # Header + 20 Zeilen
            st.text('\n'.join(preview_lines))
            
            # Download
            st.download_button(
                label="📥 CSV herunterladen",
                data=csv_content,
                file_name="ise_import.csv",
                mime="text/csv",
                use_container_width=True
            )
            
        except ValueError as ve:
            st.error(f"❌ {str(ve)}")
        except Exception as ex:
            st.error(f"💥 Unerwarteter Fehler: {str(ex)}")
    else:
        st.info("👆 Bitte wählen Sie eine Excel-Datei aus.")

if __name__ == "__main__":
    main()
