import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
from io import BytesIO

# 📊 Extract logic
def extract_event_info(file_path):
    result = {
        "Eveniment": None, "Oraș": None, "Locație": None,
        "Dată": None, "Total de plată (RON)": 0.0
    }
    try:
        df = pd.read_excel(file_path, header=None)
        for _, row in df.iterrows():
            val = str(row[0]) if not pd.isna(row[0]) else ""

            if val.lower().startswith("eveniment:"):
                event_name = val.split("Eveniment:")[-1].strip()
                result["Eveniment"] = event_name
                if ":" in event_name:
                    result["Oraș"] = event_name.split(":")[0].strip()

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

# 🖼️ Streamlit UI
st.title("📊 Raport Evenimente - Procesare Rapoarte Excel")

uploaded_file = st.file_uploader("Încarcă un fișier .zip cu rapoarte Excel", type=["zip"])

if uploaded_file:
    if not uploaded_file.name.endswith(".zip"):
        st.error("❌ Te rog să încarci un fișier .zip.")
    else:
        with tempfile.TemporaryDirectory() as tmp_dir:
            zip_path = os.path.join(tmp_dir, "uploaded.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_file.read())

            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(tmp_dir)

            results = []
            for root, _, files in os.walk(tmp_dir):
                for file in files:
                    if file.endswith(".xlsx"):
                        full_path = os.path.join(root, file)
                        result = extract_event_info(full_path)
                        results.append(result)

            df = pd.DataFrame(results)
            df_clean = df[df["Eveniment"].notna() & df["Total de plată (RON)"].apply(lambda x: isinstance(x, (int, float)))]

            if not df_clean.empty:
                df_clean["Dată"] = pd.to_datetime(df_clean["Dată"], errors="coerce", dayfirst=True)
                df_grouped = df_clean.groupby(["Eveniment", "Oraș", "Locație", "Dată"], as_index=False)["Total de plată (RON)"].sum()
                df_grouped["Dată"] = df_grouped["Dată"].dt.strftime("%d.%m.%Y")
                df_sorted = df_grouped[["Dată", "Eveniment", "Oraș", "Locație", "Total de plată (RON)"]].sort_values("Dată")

                # Dropdown filters
                st.success("✅ Procesare completă!")

                selected_city = st.selectbox("Filtrează după oraș (opțional):", ["Toate"] + sorted(df_sorted["Oraș"].dropna().unique()))
                filtered_df = df_sorted if selected_city == "Toate" else df_sorted[df_sorted["Oraș"] == selected_city]

                selected_event = st.selectbox("Filtrează după eveniment (opțional):", ["Toate"] + sorted(filtered_df["Eveniment"].dropna().unique()))
                filtered_df = filtered_df if selected_event == "Toate" else filtered_df[filtered_df["Eveniment"] == selected_event]

                st.dataframe(filtered_df)

                # Download buttons
                excel_buffer = BytesIO()
                filtered_df.to_excel(excel_buffer, index=False)
                st.download_button(
                    label="📥 Descarcă fișierul Excel",
                    data=excel_buffer.getvalue(),
                    file_name="raport_evenimente.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                csv_data = filtered_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="📥 Descarcă fișierul CSV",
                    data=csv_data,
                    file_name="raport_evenimente.csv",
                    mime="text/csv"
                )

            else:
                st.warning("⚠️ Nicio informație validă găsită în fișierele Excel.")
else:
    st.info("📁 Încarcă un fișier .zip care conține fișiere .xlsx exportate din rapoarte.")
