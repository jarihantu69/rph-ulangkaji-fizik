import streamlit as st
import google.generativeai as genai
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import datetime

# ==========================================
# 1. SETUP & KONFIGURASI 
# ==========================================
st.set_page_config(page_title="RPH Fizik Pro", page_icon="🚀", layout="wide")

GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel('gemini-2.5-flash')

# ==========================================
# 2. PANGKALAN DATA SILIBUS FIZIK KSSM
# ==========================================
silibus = {
    "Tingkatan 4": {
        "F4 Bab 1: Pengukuran": ["1.1 Kuantiti Fizik", "1.2 Penyiasatan Saintifik"],
        "F4 Bab 2: Daya dan Gerakan I": ["2.1 Gerakan Linear", "2.2 Graf Gerakan Linear", "2.3 Gerakan Jatuh Bebas", "2.4 Inersia", "2.5 Momentum", "2.6 Daya", "2.7 Impuls dan Daya Impuls", "2.8 Berat"],
        "F4 Bab 3: Kegravitian": ["3.1 Hukum Kegravitian Semesta Newton", "3.2 Hukum Kepler", "3.3 Satelit Buatan Manusia"],
        "F4 Bab 4: Haba": ["4.1 Keseimbangan Terma", "4.2 Muatan Haba Tentu", "4.3 Haba Pendam Tentu", "4.4 Hukum Gas"],
        "F4 Bab 5: Gelombang": ["5.1 Asas Gelombang", "5.2 Pelembapan dan Resonans", "5.3 Pantulan Gelombang", "5.4 Pembiasan Gelombang", "5.5 Pembelauan Gelombang", "5.6 Interferens Gelombang", "5.7 Gelombang Elektromagnet"],
        "F4 Bab 6: Cahaya dan Optik": ["6.1 Pembiasan Cahaya", "6.2 Pantulan Dalam Penuh", "6.3 Pembentukan Imej oleh Kanta", "6.4 Formula Kanta Nipis", "6.5 Peralatan Optik", "6.6 Pembentukan Imej oleh Cermin Sfera"]
    },
    "Tingkatan 5": {
        "F5 Bab 1: Daya dan Gerakan II": ["1.1 Daya Paduan", "1.2 Leraian Daya", "1.3 Keseimbangan Daya", "1.4 Kekenyalan"],
        "F5 Bab 2: Tekanan": ["2.1 Tekanan Cecair", "2.2 Tekanan Atmosfera", "2.3 Tekanan Gas", "2.4 Prinsip Pascal", "2.5 Prinsip Archimedes", "2.6 Prinsip Bernoulli"],
        "F5 Bab 3: Keelektrikan": ["3.1 Arus dan Beza Keupayaan", "3.2 Rintangan", "3.3 Daya Gerak Elektrik (d.g.e) dan Rintangan Dalam", "3.4 Tenaga dan Kuasa Elektrik"],
        "F5 Bab 4: Keelektromagnetan": ["4.1 Daya ke atas Konduktor Pembawa Arus dalam Medan Magnet", "4.2 Aruhan Elektromagnet", "4.3 Transformer"],
        "F5 Bab 5: Elektronik": ["5.1 Elektron", "5.2 Diod Semikonduktor", "5.3 Transistor"],
        "F5 Bab 6: Fizik Nuklear": ["6.1 Reputan Radioaktif", "6.2 Tenaga Nuklear"],
        "F5 Bab 7: Fizik Kuantum": ["7.1 Teori Kuantum Cahaya", "7.2 Kesan Fotoelektrik", "7.3 Teori Fotoelektrik Einstein"]
    }
}

# ==========================================
# 3. SIDEBAR (MENU NAVIGASI UTAMA)
# ==========================================
st.sidebar.title("⚙️ Modul RPH Fizik")
st.sidebar.markdown("Pilih jenis RPH yang ingin dijana:")
mod_rph = st.sidebar.radio("Pilihan Modul:", [
    "📘 PdP Biasa (Model 5E)", 
    "📝 Ulangkaji SPM"
])

st.sidebar.divider()
st.sidebar.info("Dibangunkan dengan ❤️ menggunakan AI & Streamlit")

# ==========================================
# 4. BAHAGIAN 1: MODUL PDP BIASA (5E)
# ==========================================
if mod_rph == "📘 PdP Biasa (Model 5E)":
    st.title("🚀 Penjana RPH Fizik (Model 5E)")
    st.markdown("Sistem janaan RPH berpandukan model pembelajaran 5E untuk silibus biasa.")

    kolom_kiri, kolom_kanan = st.columns(2)
    with kolom_kiri:
        st.subheader("📅 Maklumat Kelas")
        tarikh_pilihan = st.date_input("Tarikh Kelas", key="d1")
        hari_dict = {0: "Isnin", 1: "Selasa", 2: "Rabu", 3: "Khamis", 4: "Jumaat", 5: "Sabtu", 6: "Ahad"}
        hari_auto = hari_dict[tarikh_pilihan.weekday()]
        
        col_m1, col_m2 = st.columns(2)
        with col_m1: masa_mula = st.time_input("Masa Mula", value=datetime.time(7, 40), step=600, key="m1")
        with col_m2: masa_tamat = st.time_input("Masa Tamat", value=datetime.time(9, 40), step=600, key="m2")
        
        t1 = datetime.datetime.combine(datetime.date.today(), masa_mula)
        t2 = datetime.datetime.combine(datetime.date.today(), masa_tamat)
        tempoh_minit = int((t2 - t1).total_seconds() / 60)
        if tempoh_minit < 0: tempoh_minit += 1440
        
        col_k1, col_k2 = st.columns(2)
        with col_k1: nama_kelas = st.text_input("Nama Kelas", "4 Sains 1", key="k1")
        with col_k2: bil_murid = st.number_input("Bil. Murid", min_value=1, max_value=50, value=30, key="b1")
        st.info(f"📌 {hari_auto} | Tempoh: {tempoh_minit} Minit")

    with kolom_kanan:
        st.subheader("📚 Maklumat Topik")
        tingkatan = st.selectbox("Pilih Tingkatan", ["Tingkatan 4", "Tingkatan 5"], key="t1")
        senarai_bab = list(silibus[tingkatan].keys())
        bab = st.selectbox("Pilih Bab Utama", senarai_bab, key="bab1")
        senarai_sub = silibus[tingkatan][bab]
        sub_bab = st.multiselect("Pilih Sub-Topik", senarai_sub, key="sub1")
        topik_penuh = f"{bab} ({', '.join(sub_bab)})" if sub_bab else bab

    st.divider()
    st.subheader("⏳ Agihan Masa 5E (Mesti = Masa Kelas)")
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1: m_engage = st.number_input("Engage", value=5, min_value=1)
    with col2: m_explore = st.number_input("Explore", value=min(20, tempoh_minit//4), min_value=1)
    with col3: m_explain = st.number_input("Explain", value=min(20, tempoh_minit//4), min_value=1)
    with col4: m_elaborate = st.number_input("Elaborate", value=min(10, tempoh_minit//5), min_value=1)
    with col5: m_evaluate = st.number_input("Evaluate", value=5, min_value=1)
    jumlah_masa_5e = m_engage + m_explore + m_explain + m_elaborate + m_evaluate
    st.caption(f"**Jumlah Masa Disusun:** {jumlah_masa_5e} Minit / {tempoh_minit} Minit")

    if st.button("🚀 Jana RPH 5E Sekarang", use_container_width=True, type="primary"):
        if not sub_bab: st.warning("Sila pilih sekurang-kurangnya satu Sub-Topik.")
        elif jumlah_masa_5e != tempoh_minit: st.error("Jumlah masa fasa 5E tidak sama dengan tempoh kelas.")
        else:
            with st.spinner("🧠 AI sedang menyusun RPH 5E..."):
                prompt = f"""Bertindak sebagai Guru Cemerlang Fizik KSSM. Bina RPH (Model 5E) untuk topik: {topik_penuh} ({tingkatan}). Bilangan murid: {bil_murid} orang.
                Format WAJIB:
                [OBJEKTIF] (2 objektif) --- [KRITERIA KEJAYAAN] (2-3 kriteria) --- [ABM] (Senarai bahan) --- [EMK] (Elemen merentas kurikulum) --- [KBAT] (Kemahiran berfikir) --- [ENGAGE] (Aktiviti induksi) --- [EXPLORE] (Aktiviti murid) --- [EXPLAIN] (Penerangan guru) --- [ELABORATE] (Aplikasi) --- [EVALUATE] (Penilaian) --- [RUMUSAN] (Penutup) --- [REFLEKSI] (Refleksi guru realistik selepas kelas)."""
                
                hasil = model.generate_content(prompt).text
                
                teks_bersih = hasil.replace("**", "")
                def ekstrak(tag, teks):
                    if tag in teks:
                        mula = teks.split(tag)[1]
                        if "---" in mula: return mula.split("---")[0].strip()
                        else: return mula.split("[")[0].strip()
                    return ""
                
                wb = openpyxl.Workbook()
                ws = wb.active; ws.title = "RPH_5E"
                fill_title = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
                fill_sec = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
                fill_light = PatternFill(start_color="E6F0FA", end_color="E6F0FA", fill_type="solid")
                f_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
                f_sec = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
                f_b = Font(name="Segoe UI", size=11, bold=True)
                f_norm = Font(name="Segoe UI", size=11)
                border = Border(left=Side(style='thin', color='A0A0A0'), right=Side(style='thin', color='A0A0A0'), top=Side(style='thin', color='A0A0A0'), bottom=Side(style='thin', color='A0A0A0'))
                
                def gaya(r, c, f, a, fill=None):
                    cell = ws.cell(row=r, column=c)
                    cell.font = f; cell.alignment = a; cell.border = border
                    if fill: cell.fill = fill

                # Header
                ws.merge_cells("A1:D1"); ws["A1"] = "RANCANGAN PENGAJARAN HARIAN (RPH)"
                gaya(1,1,f_title, Alignment(horizontal="center", vertical="center"), fill_title); ws.row_dimensions[1].height = 25
                ws.merge_cells("A2:D2"); ws["A2"] = f"FIZIK | KSSM | {tingkatan.upper()}"
                gaya(2,1,f_sec, Alignment(horizontal="center", vertical="center"), fill_title)
                
                # Maklumat Kelas
                ws.merge_cells("A4:D4"); ws["A4"] = " MAKLUMAT KELAS"
                gaya(4,1,f_sec, Alignment(horizontal="left", vertical="center"), fill_sec)
                ws["A5"], ws["B5"] = "Hari & Tarikh:", f"{hari_auto}, {tarikh_pilihan.strftime('%d/%m/%Y')}"
                ws["C5"], ws["D5"] = "Masa & Tempoh:", f"{masa_mula.strftime('%I:%M %p')} - {masa_tamat.strftime('%I:%M %p')} ({tempoh_minit} minit)"
                ws["A6"], ws["B6"] = "Kelas:", nama_kelas
                ws["C6"], ws["D6"] = "Kehadiran:", f"_____ / {bil_murid} orang"
                
                for r in [5,6]:
                    gaya(r,1,f_b, Alignment(horizontal="left", vertical="center"), fill_light); gaya(r,2,f_norm, Alignment(horizontal="left", vertical="center"))
                    gaya(r,3,f_b, Alignment(horizontal="left", vertical="center"), fill_light); gaya(r,4,f_norm, Alignment(horizontal="left", vertical="center"))

                # Maklumat PdP
                ws.merge_cells("A8:D8"); ws["A8"] = " MAKLUMAT PENGAJARAN & PEMBELAJARAN"
                gaya(8,1,f_sec, Alignment(horizontal="left", vertical="center"), fill_sec)
                
                pdp_items = [("Topik", topik_penuh), ("Objektif", ekstrak("[OBJEKTIF]", teks_bersih)), ("Kriteria Kejayaan", ekstrak("[KRITERIA KEJAYAAN]", teks_bersih)), ("BBM", ekstrak("[ABM]", teks_bersih)), ("EMK", ekstrak("[EMK]", teks_bersih)), ("KBAT", ekstrak("[KBAT]", teks_bersih))]
                r = 9
                for lbl, val in pdp_items:
                    ws[f"A{r}"] = lbl; ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"] = val
                    gaya(r,1,f_b, Alignment(horizontal="left", vertical="top", wrap_text=True), fill_light)
                    for c in [2,3,4]: gaya(r,c,f_norm, Alignment(horizontal="left", vertical="top", wrap_text=True))
                    ws.row_dimensions[r].height = max(25, (len(str(val))//80 + 1)*18)
                    r+=1

                # Aktiviti 5E
                r+=1; ws.merge_cells(f"A{r}:D{r}"); ws[f"A{r}"] = " AKTIVITI PENGAJARAN (MODEL 5E)"
                gaya(r,1,f_sec, Alignment(horizontal="left", vertical="center"), fill_sec); r+=1
                
                aktiviti = [(f"ENGAGE\n{m_engage} Min", ekstrak("[ENGAGE]", teks_bersih)), (f"EXPLORE\n{m_explore} Min", ekstrak("[EXPLORE]", teks_bersih)), (f"EXPLAIN\n{m_explain} Min", ekstrak("[EXPLAIN]", teks_bersih)), (f"ELABORATE\n{m_elaborate} Min", ekstrak("[ELABORATE]", teks_bersih)), (f"EVALUATE\n{m_evaluate} Min", ekstrak("[EVALUATE]", teks_bersih))]
                for lbl, val in aktiviti:
                    ws[f"A{r}"] = lbl; ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"] = val
                    gaya(r,1,f_b, Alignment(horizontal="center", vertical="top", wrap_text=True), fill_light)
                    for c in [2,3,4]: gaya(r,c,f_norm, Alignment(horizontal="left", vertical="top", wrap_text=True))
                    ws.row_dimensions[r].height = max(45, (len(str(val))//80 + 2)*18)
                    r+=1

                # Penutup
                r+=1; ws.merge_cells(f"A{r}:D{r}"); ws[f"A{r}"] = " PENUTUP & REFLEKSI"
                gaya(r,1,f_sec, Alignment(horizontal="left", vertical="center"), fill_sec); r+=1
                
                ws[f"A{r}"] = "Rumusan"; ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"] = ekstrak("[RUMUSAN]", teks_bersih)
                gaya(r,1,f_b, Alignment(horizontal="left", vertical="top"), fill_light)
                for c in [2,3,4]: gaya(r,c,f_norm, Alignment(horizontal="left", vertical="top", wrap_text=True))
                ws.row_dimensions[r].height = 45; r+=1
                
                ws[f"A{r}"] = "Refleksi Guru"; ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"] = ekstrak("[REFLEKSI]", teks_bersih)
                gaya(r,1,f_b, Alignment(horizontal="left", vertical="top"), fill_light)
                for c in [2,3,4]: gaya(r,c,f_norm, Alignment(horizontal="left", vertical="top", wrap_text=True))
                ws.row_dimensions[r].height = 80

                ws.column_dimensions['A'].width = 20; ws.column_dimensions['B'].width = 25; ws.column_dimensions['C'].width = 25; ws.column_dimensions['D'].width = 30
                
                out = io.BytesIO(); wb.save(out)
                st.success("🎉 RPH 5E Berjaya Dijana!")
                st.download_button("📥 Muat Turun RPH 5E (Excel)", data=out.getvalue(), file_name=f"RPH_5E_{nama_kelas}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

# ==========================================
# 5. BAHAGIAN 2: MODUL ULANGKAJI SPM
# ==========================================
elif mod_rph == "📝 Ulangkaji SPM":
    st.title("📝 Penjana RPH Ulangkaji SPM")
    st.markdown("Gabungkan topik Tingkatan 4 & 5 untuk kelas pecutan dan latih tubi peperiksaan.")

    kolom_kiri, kolom_kanan = st.columns(2)
    with kolom_kiri:
        st.subheader("📅 Maklumat Kelas")
        tarikh_pilihan = st.date_input("Tarikh Kelas", key="d2")
        hari_dict = {0: "Isnin", 1: "Selasa", 2: "Rabu", 3: "Khamis", 4: "Jumaat", 5: "Sabtu", 6: "Ahad"}
        hari_auto = hari_dict[tarikh_pilihan.weekday()]
        
        col_m1, col_m2 = st.columns(2)
        with col_m1: masa_mula = st.time_input("Masa Mula", value=datetime.time(7, 40), step=600, key="m3")
        with col_m2: masa_tamat = st.time_input("Masa Tamat", value=datetime.time(9, 40), step=600, key="m4")
        
        t1 = datetime.datetime.combine(datetime.date.today(), masa_mula)
