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

try:
    GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
    genai.configure(api_key=GOOGLE_API_KEY)
    model = genai.GenerativeModel('gemini-2.0-flash') 
except:
    st.error("⚠️ API Key tidak dijumpai. Pastikan anda telah masukkan GOOGLE_API_KEY di Advanced Settings > Secrets.")

# ==========================================
# 2. PANGKALAN DATA SILIBUS FIZIK KSSM
# ==========================================
silibus = {
    "Tingkatan 4": {
        "F4 Bab 1: Pengukuran": ["1.1 Kuantiti Fizik", "1.2 Penyiasatan Saintifik"],
        "F4 Bab 2: Daya dan Gerakan I": ["2.1 Gerakan Linear", "2.2 Graf Gerakan Linear", "2.3 Inersia", "2.4 Momentum", "2.5 Daya", "2.6 Impuls dan Daya Impuls", "2.7 Berat"],
        "F4 Bab 3: Kegravitian": ["3.1 Hukum Kegravitian Semesta Newton", "3.2 Hukum Kepler", "3.3 Satelit Buatan Manusia"],
        "F4 Bab 4: Haba": ["4.1 Keseimbangan Terma", "4.2 Muatan Haba Tentu", "4.3 Haba Pendam Tentu", "4.4 Hukum Gas"],
        "F4 Bab 5: Gelombang": ["5.1 Asas Gelombang", "5.2 Pelembapan dan Resonans", "5.3 Pantulan Gelombang", "5.4 Pembiasan Gelombang", "5.5 Pembelauan Gelombang", "5.6 Interferens Gelombang", "5.7 Gelombang Elektromagnet"],
        "F4 Bab 6: Cahaya dan Optik": ["6.1 Pembiasan Cahaya", "6.2 Pantulan Dalam Penuh", "6.3 Pembentukan Imej oleh Kanta", "6.4 Formula Kanta Nipis", "6.5 Peralatan Optik", "6.6 Pembentukan Imej oleh Cermin Sfera"]
    },
    "Tingkatan 5": {
        "F5 Bab 1: Daya dan Gerakan II": ["1.1 Paduan Daya", "1.2 Leraian Daya", "1.3 Keseimbangan Daya", "1.4 Kekenyalan"],
        "F5 Bab 2: Tekanan": ["2.1 Tekanan Cecair", "2.2 Tekanan Atmosfera", "2.3 Tekanan Gas", "2.4 Prinsip Pascal", "2.5 Prinsip Archimedes", "2.6 Prinsip Bernoulli"],
        "F5 Bab 3: Elektrik": ["3.1 Arus dan Beda Keupayaan", "3.2 Rintangan", "3.3 Daya Gerak Elektrik (d.g.e) dan Rintangan Dalam", "3.4 Tenaga dan Kuasa Elektrik"],
        "F5 Bab 4: Keelektromagnetan": ["4.1 Daya ke atas Konduktor Pembawa Arus dalam Medan Magnet", "4.2 Aruhan Elektromagnet", "4.3 Transformer"],
        "F5 Bab 5: Elektronik": ["5.1 Elektron", "5.2 Diod Semikonduktor", "5.3 Transistor"],
        "F5 Bab 6: Fizik Nuklear": ["6.1 Reputan Radioaktif", "6.2 Tenaga Nuklear"],
        "F5 Bab 7: Fizik Kuantum": ["7.1 Teori Kuantum Cahaya", "7.2 Kesan Fotoelektrik", "7.3 Teori Fotoelektrik Einstein"]
    }
}

# ==========================================
# 3. SIDEBAR (MENU NAVIGASI UTAMA)
# ==========================================
st.sidebar.title("⚙️ Modul RPH")
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
        with col_m2: masa_tamat = st.time_input("Masa Tamat", value=datetime.time(9, 00), step=600, key="m2")
        
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
    st.subheader("⏳ Agihan Masa 5E")
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1: m_engage = st.number_input("Engage", value=5, min_value=1)
    with col2: m_explore = st.number_input("Explore", value=min(20, tempoh_minit//4), min_value=1)
    with col3: m_explain = st.number_input("Explain", value=min(20, tempoh_minit//4), min_value=1)
    with col4: m_elaborate = st.number_input("Elaborate", value=min(10, tempoh_minit//5), min_value=1)
    with col5: m_evaluate = st.number_input("Evaluate", value=5, min_value=1)
    jumlah_masa_5e = m_engage + m_explore + m_explain + m_elaborate + m_evaluate
    st.caption(f"**Jumlah Masa:** {jumlah_masa_5e} Minit / {tempoh_minit} Minit")

    if st.button("🚀 Jana RPH 5E Sekarang", use_container_width=True, type="primary"):
        if not sub_bab: st.warning("Sila pilih sekurang-kurangnya satu Sub-Topik.")
        elif jumlah_masa_5e != tempoh_minit: st.error("Jumlah masa fasa 5E tidak sama dengan tempoh kelas.")
        else:
            with st.spinner("🧠 AI sedang menyusun RPH 5E..."):
                prompt = f"""Bertindak sebagai Guru Cemerlang Fizik KSSM. Bina RPH (Model 5E) untuk topik: {topik_penuh} ({tingkatan}). Bilangan murid: {bil_murid} orang.
                Format WAJIB:
                [OBJEKTIF] (2 objektif) --- [KRITERIA KEJAYAAN] (2-3 kriteria) --- [ABM] (Senarai bahan) --- [EMK] (Elemen merentas kurikulum) --- [KBAT] (Kemahiran berfikir) --- [ENGAGE] (Aktiviti induksi) --- [EXPLORE] (Aktiviti murid) --- [EXPLAIN] (Penerangan guru) --- [ELABORATE] (Aplikasi) --- [EVALUATE] (Penilaian) --- [RUMUSAN] (Penutup) --- [REFLEKSI] (Refleksi guru)."""
                
                hasil = model.generate_content(prompt).text
                
                # Fungsi Ekstrak Teks
                teks_bersih = hasil.replace("**", "")
                def ekstrak(tag, teks):
                    if tag in teks: return teks.split(tag)[1].split("---")[0].split("[")[0].strip()
                    return ""
                
                # Setup Excel
                wb = openpyxl.Workbook()
                ws = wb.active; ws.title = "RPH_5E"
                fill_title = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
                fill_sec = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
                fill_light = PatternFill(start_color="E6F0FA", end_color="E6F0FA", fill_type="solid")
                f_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
                f_sec = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
                f_b = Font(name="Segoe UI", size=11, bold=True)
                border = Border(left=Side(style='thin', color='A0A0A0'), right=Side(style='thin', color='A0A0A0'), top=Side(style='thin', color='A0A0A0'), bottom=Side(style='thin', color='A0A0A0'))
                
                def gaya(r, c, f, a, fill=None):
                    ws.cell(r,c).font = f; ws.cell(r,c).alignment = a; ws.cell(r,c).border = border; 
                    if fill: ws.cell(r,c).fill = fill

                # Header
                ws.merge_cells("A1:D1"); ws["A1"] = "RPH FIZIK (MODEL 5E)"; gaya(1,1,f_title, Alignment(horizontal="center"), fill_title)
                ws.merge_cells("A2:D2"); ws["A2"] = f"{tingkatan.upper()} | {topik_penuh}"; gaya(2,1,f_sec, Alignment(horizontal="center"), fill_title)
                
                # Info
                ws.merge_cells("A4:D4"); ws["A4"] = "MAKLUMAT KELAS & PDP"; gaya(4,1,f_sec, Alignment(horizontal="left"), fill_sec)
                info = [("Tarikh", f"{hari_auto}, {tarikh_pilihan.strftime('%d/%m/%Y')}"), ("Masa", f"{masa_mula.strftime('%I:%M %p')} - {masa_tamat.strftime('%I:%M %p')}"), ("Kelas", nama_kelas), ("Objektif", ekstrak("[OBJEKTIF]", teks_bersih)), ("Kriteria", ekstrak("[KRITERIA KEJAYAAN]", teks_bersih))]
                r = 5
                for k,v in info:
                    ws[f"A{r}"] = k; ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"] = v
                    gaya(r,1,f_b, Alignment(vertical="top"), fill_light); 
                    for c in [2,3,4]: gaya(r,c,Font(name="Segoe UI", size=11), Alignment(wrap_text=True))
                    ws.row_dimensions[r].height = max(20, (len(str(v))//80 + 1)*18)
                    r+=1
                
                # Aktiviti
                ws.merge_cells(f"A{r}:D{r}"); ws[f"A{r}"] = "AKTIVITI 5E"; gaya(r,1,f_sec, Alignment(horizontal="left"), fill_sec); r+=1
                akt = [("ENGAGE", ekstrak("[ENGAGE]", teks_bersih)), ("EXPLORE", ekstrak("[EXPLORE]", teks_bersih)), ("EXPLAIN", ekstrak("[EXPLAIN]", teks_bersih)), ("ELABORATE", ekstrak("[ELABORATE]", teks_bersih)), ("EVALUATE", ekstrak("[EVALUATE]", teks_bersih))]
                for k,v in akt:
                    ws[f"A{r}"] = k; ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"] = v
                    gaya(r,1,f_b, Alignment(vertical="top"), fill_light); 
                    for c in [2,3,4]: gaya(r,c,Font(name="Segoe UI", size=11), Alignment(wrap_text=True))
                    ws.row_dimensions[r].height = max(40, (len(str(v))//80 + 2)*18)
                    r+=1

                ws.column_dimensions['A'].width = 15; ws.column_dimensions['B'].width = 25; ws.column_dimensions['C'].width = 25; ws.column_dimensions['D'].width = 25
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
        with col_m2: masa_tamat = st.time_input("Masa Tamat", value=datetime.time(9, 00), step=600, key="m4")
        
        t1 = datetime.datetime.combine(datetime.date.today(), masa_mula)
        t2 = datetime.datetime.combine(datetime.date.today(), masa_tamat)
        tempoh_minit = int((t2 - t1).total_seconds() / 60)
        if tempoh_minit < 0: tempoh_minit += 1440
        
        col_k1, col_k2 = st.columns(2)
        with col_k1: nama_kelas = st.text_input("Nama Kelas", "5 Sains 1", key="k2")
        with col_k2: bil_murid = st.number_input("Bil. Murid", min_value=1, max_value=50, value=30, key="b2")
        st.info(f"📌 {hari_auto} | Tempoh: {tempoh_minit} Minit")

    with kolom_kanan:
        st.subheader("🎯 Fokus Topik Ulangkaji")
        format_kertas = st.selectbox("Fokus Kertas", ["Kertas 1 (Objektif)", "Kertas 2 - Bahagian A", "Kertas 2 - Esei", "Kertas 3 (Amali)"])
        bahan_bbm = st.text_input("Bahan Ulangkaji", "Modul Latih Tubi SPM")
        tingkatan_pilihan = st.multiselect("Pilih Tingkatan", ["Tingkatan 4", "Tingkatan 5"], default=["Tingkatan 5"])
        
        senarai_bab_gabungan = []
        for t in tingkatan_pilihan: senarai_bab_gabungan.extend(list(silibus[t].keys()))
        
        pilih_semua_bab = st.checkbox("☑️ Pilih Semua Bab")
        bab_dipilih = st.multiselect("Pilih Bab", options=senarai_bab_gabungan, default=senarai_bab_gabungan if pilih_semua_bab else None)
        
        senarai_sub_gabungan = []
        for t in tingkatan_pilihan:
            for b in bab_dipilih:
                if b in silibus[t]: senarai_sub_gabungan.extend(silibus[t][b])
        
        sub_bab_dipilih = st.multiselect("Sub-topik Spesifik (Pilihan)", options=senarai_sub_gabungan)
        txt_tingkatan = " & ".join(tingkatan_pilihan)
        topik_penuh = "Semua Bab" if pilih_semua_bab else ", ".join(bab_dipilih)

    st.divider()
    st.subheader("⏳ Agihan Masa Ulangkaji")
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1: m_recall = st.number_input("Recall", value=10)
    with col2: m_teknik = st.number_input("Teknik", value=20)
    with col3: m_latih = st.number_input("Latih Tubi", value=30)
    with col4: m_bincang = st.number_input("Bincang", value=20)
    with col5: m_tutup = st.number_input("Penutup", value=10)
    jumlah_masa_fasa = m_recall + m_teknik + m_latih + m_bincang + m_tutup
    st.caption(f"**Jumlah Masa:** {jumlah_masa_fasa} Minit / {tempoh_minit} Minit")

    if st.button("🚀 Jana RPH Ulangkaji Sekarang", use_container_width=True, type="primary"):
        if not bab_dipilih: st.warning("Sila pilih sekurang-kurangnya satu Bab.")
        elif jumlah_masa_fasa != tempoh_minit: st.error("Jumlah masa fasa tidak sama dengan tempoh kelas.")
        else:
            with st.spinner("🧠 AI sedang menyusun strategi ulangkaji..."):
                prompt = f"""Bertindak sebagai Guru Cemerlang Fizik SPM. Bina RPH Ulangkaji untuk topik: {topik_penuh} ({txt_tingkatan}). Fokus: Latih Tubi {format_kertas}. Bahan: {bahan_bbm}. Bilangan murid: {bil_murid} orang.
                Format WAJIB:
                [OBJEKTIF] (2 objektif) --- [KRITERIA KEJAYAAN] (2-3 kriteria) --- [ABM] (Senarai bahan: {bahan_bbm}) --- [EMK] (Elemen) --- [KBAT] (Kata tugas SPM) --- [RECALL] (Aktiviti mengingat) --- [TEKNIK] (Penerangan teknik) --- [LATIH_TUBI] (Aktiviti menjawab) --- [PERBINCANGAN] (Menyemak skema) --- [PENUTUP] (Rumusan) --- [REFLEKSI] (Refleksi guru)."""
                
                hasil = model.generate_content(prompt).text
                
                teks_bersih = hasil.replace("**", "")
                def ekstrak(tag, teks):
                    if tag in teks: return teks.split(tag)[1].split("---")[0].split("[")[0].strip()
                    return ""
                
                # Setup Excel
                wb = openpyxl.Workbook()
                ws = wb.active; ws.title = "RPH_Ulangkaji"
                fill_title = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
                fill_sec = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
                fill_light = PatternFill(start_color="E6F0FA", end_color="E6F0FA", fill_type="solid")
                f_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
                f_sec = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
                f_b = Font(name="Segoe UI", size=11, bold=True)
                border = Border(left=Side(style='thin', color='A0A0A0'), right=Side(style='thin', color='A0A0A0'), top=Side(style='thin', color='A0A0A0'), bottom=Side(style='thin', color='A0A0A0'))
                
                def gaya(r, c, f, a, fill=None):
                    ws.cell(r,c).font = f; ws.cell(r,c).alignment = a; ws.cell(r,c).border = border; 
                    if fill: ws.cell(r,c).fill = fill

                # Header
                ws.merge_cells("A1:D1"); ws["A1"] = "RPH ULANGKAJI SPM"; gaya(1,1,f_title, Alignment(horizontal="center"), fill_title)
                ws.merge_cells("A2:D2"); ws["A2"] = f"FIZIK KSSM | {txt_tingkatan}"; gaya(2,1,f_sec, Alignment(horizontal="center"), fill_title)
                
                # Info
                ws.merge_cells("A4:D4"); ws["A4"] = "MAKLUMAT KELAS & FOKUS"; gaya(4,1,f_sec, Alignment(horizontal="left"), fill_sec)
                info = [("Tarikh", f"{hari_auto}, {tarikh_pilihan.strftime('%d/%m/%Y')}"), ("Topik", topik_penuh), ("Fokus Kertas", format_kertas), ("Objektif", ekstrak("[OBJEKTIF]", teks_bersih)), ("Kriteria", ekstrak("[KRITERIA KEJAYAAN]", teks_bersih))]
                r = 5
                for k,v in info:
                    ws[f"A{r}"] = k; ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"] = v
                    gaya(r,1,f_b, Alignment(vertical="top"), fill_light); 
                    for c in [2,3,4]: gaya(r,c,Font(name="Segoe UI", size=11), Alignment(wrap_text=True))
                    ws.row_dimensions[r].height = max(20, (len(str(v))//80 + 1)*18)
                    r+=1
                
                # Aktiviti
                ws.merge_cells(f"A{r}:D{r}"); ws[f"A{r}"] = "STRATEGI ULANGKAJI"; gaya(r,1,f_sec, Alignment(horizontal="left"), fill_sec); r+=1
                akt = [("RECALL", ekstrak("[RECALL]", teks_bersih)), ("TEKNIK", ekstrak("[TEKNIK]", teks_bersih)), ("LATIH TUBI", ekstrak("[LATIH_TUBI]", teks_bersih)), ("BINCANG", ekstrak("[PERBINCANGAN]", teks_bersih)), ("PENUTUP", ekstrak("[PENUTUP]", teks_bersih))]
                for k,v in akt:
                    ws[f"A{r}"] = k; ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"] = v
                    gaya(r,1,f_b, Alignment(vertical="top"), fill_light); 
                    for c in [2,3,4]: gaya(r,c,Font(name="Segoe UI", size=11), Alignment(wrap_text=True))
                    ws.row_dimensions[r].height = max(40, (len(str(v))//80 + 2)*18)
                    r+=1

                ws.column_dimensions['A'].width = 15; ws.column_dimensions['B'].width = 25; ws.column_dimensions['C'].width = 25; ws.column_dimensions['D'].width = 25
                out = io.BytesIO(); wb.save(out)
                
                st.success("🎉 RPH Ulangkaji Berjaya Dijana!")
                st.download_button("📥 Muat Turun RPH Ulangkaji (Excel)", data=out.getvalue(), file_name=f"RPH_Ulangkaji_{nama_kelas}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
