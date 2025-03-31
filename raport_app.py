import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
from io import BytesIO
import re

# Function to extract data from one file
def extract_event_info(file_path):
    result = {
        "Eveniment complet": None, "ID Eveniment": None, "Eveniment": None,
        "Artiști": None, "Oraș": None, "Locație": None, "Dată": None, "Total de plată (RON)": 0.0
    }
    try:
        df = pd.read_excel(file_path, header=None)
        for _, row in df.iterrows():
            val = str(row[0]) if not pd.isna(row[0]) else ""

            if val.lower().startswith("eveniment:"):
                event_full = val.split("Eveniment:")[-1].strip()
                result["Eveniment complet"] = event_full

                # Extract event ID (assuming it's the first number in the title)
                event_id_match = re.search(r'\b\d{5,}\b', event_full)
                result["ID Eveniment"] = event_id_match.group(0) if event_id_match else "Fără ID"

                # Extract city (exclude numeric values)
                city_candidate = event_full.split(":")[0].strip()
                result["Oraș"] = city_candidate if not any(char.isdigit() for char in city_candidate) else None

                # Extract artist names based on known patterns like "cu X, Y și Z"
                artist_match = re.search(r'cu (.+)', event_full, re.IGNORECASE)
                if artist_match:
                    artists = artist_match.group(1).strip()
                    # Remove IDs and trailing patterns like (ID:) or "- "
                    artists_clean = re.split(r'(\(ID:.*?\)|-\s)', artists)[0].strip(" -")
                    artists_clean = re.sub(r'\b\d{5,}\b', '', artists_clean).strip(" -")
                    result["Artiști"] = artists_clean

                # Clean event name without ID or artists
                cleaned_title = re.sub(r'\b\d{5,}\b', '', event_full)  # remove event id
                cleaned_title = re.sub(r'cu .+', '', cleaned_title, flags=re.IGNORECASE)  # remove artists
                result["Eveniment"] = cleaned_title.strip(" -:.").strip()

            elif val.lower().startswith("locatie / data eveniment:"):
                loc_data = val.split("Locatie / Data eveniment:")[-1].strip()
                if " / " in loc_data:
                    location, date = map(str.strip, loc_data.split(" / ", 1))
                    result["Locație"] = location
                    result["Dată"] = date
                else:
                    result["Locație"] = loc_data

            if "Total de plată cf. raport (RON) (=1-2-3)" in val:
                if len(row) > 3 and isinstance(row[3], (int, float)):
                    result["Total de plată (RON)"] = float(row[3])
    except Exception as e:
        st.warning(f"❌ Eroare în fișierul {os.path.basename(file_path)}: {e}")
    return result

# UI
st.title("📊 Raport Evenimente - Procesare Fișiere Excel")

st.info("📂 Încarcă un fișier .zip care conține rapoarte Excel (.xlsx) sau încarcă fișiere .xlsx direct. Fișierele .rar nu sunt acceptate.")

uploaded_files = st.file_uploader(
    "Selectează fișierul .zip sau fișierele .xlsx",
    type=["zip", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    with tempfile.TemporaryDirectory() as tmp_dir:
        xlsx_files = []

        for uploaded_file in uploaded_files:
            filename = uploaded_file.name
            file_path = os.path.join(tmp_dir, filename)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.read())

            if filename.endswith(".zip"):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(tmp_dir)

            elif filename.endswith(".xlsx"):
                xlsx_files.append(file_path)

        # Collect all .xlsx files from extraction folder
        for root, _, files in os.walk(tmp_dir):
            for file in files:
                if file.endswith(".xlsx") and os.path.join(root, file) not in xlsx_files:
                    xlsx_files.append(os.path.join(root, file))

        results = [extract_event_info(file) for file in xlsx_files]

        df = pd.DataFrame(results)
        df_clean = df[df["Eveniment"].notna() & df["Total de plată (RON)"].apply(lambda x: isinstance(x, (int, float)))]

        if not df_clean.empty:
            df_clean["Dată"] = pd.to_datetime(df_clean["Dată"], errors="coerce", dayfirst=True)
            df_grouped = df_clean.groupby(["ID Eveniment", "Eveniment", "Artiști", "Oraș", "Locație", "Dată"], as_index=False)["Total de plată (RON)"].sum()
            df_grouped["Dată"] = df_grouped["Dată"].dt.strftime("%d.%m.%Y")
            df_sorted = df_grouped[["Dată", "ID Eveniment", "Eveniment", "Artiști", "Oraș", "Locație", "Total de plată (RON)"]].sort_values("Dată")

            st.success("✅ Procesare completă!")

            selected_city = st.selectbox("Filtrează după oraș (opțional):", ["Toate"] + sorted(df_sorted["Oraș"].dropna().unique()))
            filtered_df = df_sorted if selected_city == "Toate" else df_sorted[df_sorted["Oraș"] == selected_city]

            selected_id = st.selectbox("Filtrează după ID Eveniment (opțional):", ["Toate"] + sorted(filtered_df["ID Eveniment"].dropna().unique()))
            filtered_df = filtered_df if selected_id == "Toate" else filtered_df[filtered_df["ID Eveniment"] == selected_id]

            selected_artist = st.selectbox("Filtrează după artist (opțional):", ["Toți"] + sorted(filtered_df["Artiști"].dropna().unique()))
            filtered_df = filtered_df if selected_artist == "Toți" else filtered_df[filtered_df["Artiști"] == selected_artist]

            st.dataframe(filtered_df)

            excel_buffer = BytesIO()
            filtered_df.to_excel(excel_buffer, index=False)
            st.download_button("📥 Descarcă fișierul Excel", excel_buffer.getvalue(), "raport_evenimente.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            csv_data = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button("📥 Descarcă fișierul CSV", csv_data, "raport_evenimente.csv", mime="text/csv")

        else:
            st.warning("⚠️ Nicio informație validă găsită în fișierele Excel.")
