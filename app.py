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
st.set_page_config(page_title="RPH Ulangkaji Fizik F4-F5", page_icon="📝", layout="wide")

# LETAK API KEY HANG KAT SINI
GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel('gemini-2.5-flash')

# ==========================================
# 2. PANGKALAN DATA SILIBUS FIZIK KSSM (T4 & T5)
# ==========================================
silibus_fizik_kssm = {
    "Tingkatan 4": {
        "1. Pengukuran": ["1.1 Kuantiti Fizik", "1.2 Penyiasatan Saintifik"],
        "2. Daya dan Gerakan I": ["2.1 Gerakan Linear", "2.2 Graf Gerakan Linear", "2.3 Inersia", "2.4 Momentum", "2.5 Daya", "2.6 Impuls dan Daya Impuls", "2.7 Berat"],
        "3. Kegravitian": ["3.1 Hukum Kegravitian Semesta Newton", "3.2 Hukum Kepler", "3.3 Satelit Buatan Manusia"],
        "4. Haba": ["4.1 Keseimbangan Terma", "4.2 Muatan Haba Tentu", "4.3 Haba Pendam Tentu", "4.4 Hukum Gas"],
        "5. Gelombang": ["5.1 Asas Gelombang", "5.2 Pelembapan dan Resonans", "5.3 Pantulan Gelombang", "5.4 Pembiasan Gelombang", "5.5 Pembelauan Gelombang", "5.6 Interferens Gelombang", "5.7 Gelombang Elektromagnet"],
        "6. Cahaya dan Optik": ["6.1 Pembiasan Cahaya", "6.2 Pantulan Dalam Penuh", "6.3 Pembentukan Imej oleh Kanta", "6.4 Formula Kanta Nipis", "6.5 Peralatan Optik", "6.6 Pembentukan Imej oleh Cermin Sfera"]
    },
    "Tingkatan 5": {
        "1. Daya dan Gerakan II": ["1.1 Paduan Daya", "1.2 Leraian Daya", "1.3 Keseimbangan Daya", "1.4 Kekenyalan"],
        "2. Tekanan": ["2.1 Tekanan Cecair", "2.2 Tekanan Atmosfera", "2.3 Tekanan Gas", "2.4 Prinsip Pascal", "2.5 Prinsip Archimedes", "2.6 Prinsip Bernoulli"],
        "3. Elektrik": ["3.1 Arus dan Beda Keupayaan", "3.2 Rintangan", "3.3 Daya Gerak Elektrik (d.g.e) dan Rintangan Dalam", "3.4 Tenaga dan Kuasa Elektrik"],
        "4. Keelektromagnetan": ["4.1 Daya ke atas Konduktor Pembawa Arus dalam Medan Magnet", "4.2 Aruhan Elektromagnet", "4.3 Transformer"],
        "5. Elektronik": ["5.1 Elektron", "5.2 Diod Semikonduktor", "5.3 Transistor"],
        "6. Fizik Nuklear": ["6.1 Reputan Radioaktif", "6.2 Tenaga Nuklear"],
        "7. Fizik Kuantum": ["7.1 Teori Kuantum Cahaya", "7.2 Kesan Fotoelektrik", "7.3 Teori Fotoelektrik Einstein"]
    }
}

# ==========================================
# 3. ANTARA MUKA (UI) WEB APPS
# ==========================================
st.title("📝 Penjana RPH Ulangkaji Fizik SPM Cg Azaril")
st.markdown("Isi maklumat di bawah khusus untuk kelas persediaan SPM. AI akan jana RPH berfokuskan latih tubi dan teknik menjawab.")

kolom_kiri, kolom_kanan = st.columns(2)

with kolom_kiri:
    st.subheader("📅 Maklumat Kelas")
    tarikh_pilihan = st.date_input("Tarikh Kelas")
    hari_dict = {0: "Isnin", 1: "Selasa", 2: "Rabu", 3: "Khamis", 4: "Jumaat", 5: "Sabtu", 6: "Ahad"}
    hari_auto = hari_dict[tarikh_pilihan.weekday()]
    
    col_m1, col_m2 = st.columns(2)
    with col_m1: masa_mula = st.time_input("Masa Mula", value=datetime.time(7, 40), step=600)
    with col_m2: masa_tamat = st.time_input("Masa Tamat", value=datetime.time(9, 00), step=600)

    t1 = datetime.datetime.combine(datetime.date.today(), masa_mula)
    t2 = datetime.datetime.combine(datetime.date.today(), masa_tamat)
    tempoh_minit = int((t2 - t1).total_seconds() / 60)
    if tempoh_minit < 0: tempoh_minit += 1440 
    
    col_k1, col_k2 = st.columns(2)
    with col_k1: nama_kelas = st.text_input("Nama Kelas", "5 Sains 1")
    with col_k2: bil_murid = st.number_input("Bilangan Murid", min_value=1, max_value=50, value=30)
    
    st.info(f"📌 **Rumusan:** Hari {hari_auto}, {tarikh_pilihan.strftime('%d/%m/%Y')} | Tempoh: {tempoh_minit} Minit")

with kolom_kanan:
    st.subheader("🎯 Fokus Ulangkaji & Topik")
    format_kertas = st.selectbox("Fokus Kertas SPM", [
        "Kertas 1 (Objektif)", 
        "Kertas 2 - Bahagian A (Struktur)", 
        "Kertas 2 - Bahagian B & C (Esei)", 
        "Kertas 3 (Ujian Amali Bersepadu)"
    ])
    bahan_bbm = st.text_input("Bahan Ulangkaji / Modul", placeholder="Cth: Modul Pintar, Percubaan SBP 2023...")

    tingkatan = st.selectbox("Pilih Tingkatan", ["Tingkatan 4", "Tingkatan 5"])
    senarai_bab = list(silibus_fizik_kssm[tingkatan].keys())
    bab = st.multiselect("Pilih Bab (Boleh lebih dari 1 untuk ulangkaji)", senarai_bab)
    
    # Ambil subtopik berdasarkan bab yang dipilih
    senarai_sub = []
    for b in bab:
        senarai_sub.extend(silibus_fizik_kssm[tingkatan][b])
        
    sub_bab = st.multiselect("Pilih Sub-topik Fokus", senarai_sub)
    
    topik_penuh = ", ".join(bab) if not sub_bab else f"{', '.join(bab)} ({', '.join(sub_bab)})"

st.divider()

st.subheader("⏳ Agihan Masa Fasa Ulangkaji (Pastikan Jumlah = Masa Kelas)")
col1, col2, col3, col4, col5 = st.columns(5)
with col1: m_recall = st.number_input("Recall/Induksi", value=10, min_value=1)
with col2: m_teknik = st.number_input("Teknik Menjawab", value=min(20, tempoh_minit//4), min_value=1)
with col3: m_latih = st.number_input("Latih Tubi", value=min(30, tempoh_minit//3), min_value=1)
with col4: m_bincang = st.number_input("Perbincangan", value=min(15, tempoh_minit//4), min_value=1)
with col5: m_tutup = st.number_input("Rumusan/Penutup", value=5, min_value=1)

jumlah_masa_fasa = m_recall + m_teknik + m_latih + m_bincang + m_tutup
st.caption(f"**Jumlah Masa Disusun:** {jumlah_masa_fasa} Minit daripada {tempoh_minit} Minit")

# ==========================================
# 4. FUNGSI JANA AI
# ==========================================
def jana_rph_ai(topik, ting, bil_murid, format_spm, bbm):
    prompt = f"""
    Bertindak sebagai Guru Cemerlang Fizik SPM.
    Bina Rancangan Pengajaran Harian (RPH) Ulangkaji untuk topik: {topik} ({ting}).
    Fokus Kelas: Latih Tubi {format_spm}.
    Bahan digunakan: {bbm}.
    Bilangan murid: {bil_murid} orang.

    Format WAJIB (Jangan letak simbol asteris atau bold pada tag, tulis sebijik macam ni):
    [OBJEKTIF] (Tulis 2 objektif pembelajaran berfokuskan kemahiran menjawab soalan SPM)
    ---
    [KRITERIA KEJAYAAN] (Tulis 2 atau 3 kriteria kejayaan mengikut format pemarkahan SPM)
    ---
    [ABM] (Senarai bahan: {bbm} dan alat lain yang sesuai)
    ---
    [EMK] (Elemen merentas kurikulum)
    ---
    [KBAT] (Nyatakan kata tugas soalan cth: Menilai, Mencipta, Menganalisis)
    ---
    [RECALL] (Aktiviti mengingat semula formula/konsep asas Fizik yang berkaitan)
    ---
    [TEKNIK] (Penerangan guru tentang teknik menjawab, penggunaan rumus, atau cara dapat markah penuh untuk {format_spm})
    ---
    [LATIH_TUBI] (Aktiviti murid menjawab soalan daripada bahan {bbm} secara kendiri/kumpulan)
    ---
    [PERBINCANGAN] (Aktiviti menyemak jawapan, mengenal pasti miskonsepsi pelajar, dan pendedahan skema pemarkahan)
    ---
    [PENUTUP] (Rumusan markah/prestasi kelas dan take-away points untuk elak kecuaian exam)
    ---
    [REFLEKSI] (Reka satu perenggan refleksi guru yang realistik seolah-olah kelas {bil_murid} pelajar ini telah selesai menjalani ujian formatif).
    """
    response = model.generate_content(prompt)
    return response.text

# ==========================================
# 5. FUNGSI EKSPORT KE EXCEL 
# ==========================================
def bina_excel(data_ai, data_input):
    teks_bersih = data_ai.replace("**", "")
    
    def ekstrak(tag, teks):
        if tag in teks:
            mula = teks.split(tag)[1]
            if "---" in mula:
                return mula.split("---")[0].strip()
            else:
                return mula.split("[")[0].strip()
        return ""

    excel_data = {
        'objektif': ekstrak("[OBJEKTIF]", teks_bersih), 
        'kk': ekstrak("[KRITERIA KEJAYAAN]", teks_bersih),
        'abm': ekstrak("[ABM]", teks_bersih), 
        'emk': ekstrak("[EMK]", teks_bersih), 
        'kbat': ekstrak("[KBAT]", teks_bersih),
        'recall': ekstrak("[RECALL]", teks_bersih), 
        'teknik': ekstrak("[TEKNIK]", teks_bersih), 
        'latih': ekstrak("[LATIH_TUBI]", teks_bersih),
        'bincang': ekstrak("[PERBINCANGAN]", teks_bersih), 
        'tutup': ekstrak("[PENUTUP]", teks_bersih),
        'refleksi': ekstrak("[REFLEKSI]", teks_bersih)
    }

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RPH_Ulangkaji"

    navy_dark, blue_accent = "1B365D", "E6F0FA"
    fill_title = PatternFill(start_color=navy_dark, end_color=navy_dark, fill_type="solid")
    fill_section = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
    fill_light = PatternFill(start_color=blue_accent, end_color=blue_accent, fill_type="solid")
    
    f_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    f_sec = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    f_b = Font(name="Segoe UI", size=11, bold=True)
    f_norm = Font(name="Segoe UI", size=11)
    
    t_side = Side(style='thin', color='A0A0A0')
    border = Border(left=t_side, right=t_side, top=t_side, bottom=t_side)
    
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_lt = Alignment(horizontal="left", vertical="top", wrap_text=True)

    def gayakan(r, c, font, align, bord=border, fill=None):
        cell = ws.cell(row=r, column=c)
        cell.font, cell.alignment = font, align
        if bord: cell.border = bord
        if fill: cell.fill = fill

    # --- TAJUK UTAMA ---
    ws.merge_cells("A1:D1"); ws["A1"] = "RANCANGAN PENGAJARAN HARIAN (ULANGKAJI SPM)"
    gayakan(1, 1, f_title, al_c, None, fill_title); ws.row_dimensions[1].height = 25
    ws.merge_cells("A2:D2"); ws["A2"] = f"FIZIK KSSM | {data_input['tingkatan'].upper()}"
    gayakan(2, 1, Font(name="Segoe UI", size=11, bold=True, color="FFFFFF"), al_c, None, fill_title)
    
    # --- MAKLUMAT KELAS ---
    ws.merge_cells("A4:D4"); ws["A4"] = " MAKLUMAT KELAS"
    gayakan(4, 1, f_sec, al_l, border, fill_section)
    ws["A5"], ws["B5"] = "Hari & Tarikh:", f"{data_input['hari']}, {data_input['tarikh']}"
    ws["C5"], ws["D5"] = "Masa & Tempoh:", f"{data_input['masa']} ({data_input['tempoh']} minit)"
    ws["A6"], ws["B6"] = "Kelas:", data_input['kelas']
    ws["C6"], ws["D6"] = "Kehadiran:", f"_____ / {data_input['bil_murid']} orang"
    
    for r in [5,6]:
        gayakan(r, 1, f_b, al_l, border, fill_light); gayakan(r, 2, f_norm, al_l, border)
        gayakan(r, 3, f_b, al_l, border, fill_light); gayakan(r, 4, f_norm, al_l, border)

    # --- MAKLUMAT PDP ---
    ws.merge_cells("A8:D8"); ws["A8"] = " MAKLUMAT FOKUS ULANGKAJI"
    gayakan(8, 1, f_sec, al_l, border, fill_section)
    
    pdp_items = [("Topik / Bab", data_input['topik']), ("Fokus Kertas", data_input['format_spm']),
                 ("Bahan/Modul", data_input['bbm']),
                 ("Objektif", excel_data['objektif']), ("Kriteria Kejayaan", excel_data['kk']), 
                 ("KBAT / Fokus", excel_data['kbat'])]
    r = 9
    for lbl, val in pdp_items:
        ws[f"A{r}"] = lbl; ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"] = val
        gayakan(r, 1, f_b, al_lt, border, fill_light)
        for c in [2,3,4]: gayakan(r, c, f_norm, al_lt, border)
        ws.row_dimensions[r].height = max(25, (len(str(val)) // 80 + 1) * 18)
        r+=1

    # --- AKTIVITI PENGAJARAN ---
    r+=1; ws.merge_cells(f"A{r}:D{r}"); ws[f"A{r}"] = " STRATEGI ULANGKAJI & LATIH TUBI"
    gayakan(r, 1, f_sec, al_l, border, fill_section); r+=1
    
    aktiviti = [
        (f"RECALL / INDUKSI\n{data_input['m_recall']} Min", excel_data['recall']),
        (f"TEKNIK MENJAWAB\n{data_input['m_teknik']} Min", excel_data['teknik']),
        (f"LATIH TUBI\n{data_input['m_latih']} Min", excel_data['latih']),
        (f"PERBINCANGAN\n{data_input['m_bincang']} Min", excel_data['bincang']),
        (f"PENUTUP\n{data_input['m_tutup']} Min", excel_data['tutup'])
    ]
    for lbl, val in aktiviti:
        ws[f"A{r}"] = lbl; ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"] = val
        gayakan(r, 1, f_b, Alignment(horizontal="center", vertical="top", wrap_text=True), border, fill_light)
        for c in [2,3,4]: gayakan(r, c, f_norm, al_lt, border)
        ws.row_dimensions[r].height = max(45, (len(str(val)) // 80 + 2) * 18)
        r+=1

    # --- REFLEKSI ---
    r+=1; ws.merge_cells(f"A{r}:D{r}"); ws[f"A{r}"] = " REFLEKSI GURU"
    gayakan(r, 1, f_sec, al_l, border, fill_section); r+=1
    
    ws.merge_cells(f"A{r}:D{r}"); ws[f"A{r}"] = excel_data['refleksi']
    gayakan(r, 1, f_norm, al_lt, border)
    for c in [2,3,4]: gayakan(r, c, f_norm, al_lt, border)
    ws.row_dimensions[r].height = max(60, (len(str(excel_data['refleksi'])) // 80 + 2) * 18)
    
    # --- LAYOUT AKHIR ---
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 30
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ==========================================
# 6. BUTANG PELAKSANAAN UTAMA
# ==========================================
if st.button("🚀 Jana RPH Ulangkaji SPM Sekarang", use_container_width=True, type="primary"):
    if not bab:
        st.warning("Sila pilih sekurang-kurangnya satu Bab untuk diulangkaji.")
    elif jumlah_masa_fasa != tempoh_minit:
        st.error(f"Jumlah masa fasa ulangkaji ({jumlah_masa_fasa} minit) tidak sama dengan tempoh kelas ({tempoh_minit} minit). Sila betulkan.")
    else:
        with st.spinner("🧠 AI sedang menyusun strategi teknik menjawab SPM..."):
            hasil = jana_rph_ai(topik_penuh, tingkatan, bil_murid, format_kertas, bahan_bbm)
            
            data_input = {
                "tingkatan": tingkatan, "topik": topik_penuh, "bil_murid": bil_murid,
                "hari": hari_auto, "tarikh": tarikh_pilihan.strftime('%d/%m/%Y'),
                "masa": f"{masa_mula.strftime('%I:%M %p')} - {masa_tamat.strftime('%I:%M %p')}",
                "tempoh": tempoh_minit, "kelas": nama_kelas, "format_spm": format_kertas, "bbm": bahan_bbm,
                "m_recall": m_recall, "m_teknik": m_teknik, "m_latih": m_latih,
                "m_bincang": m_bincang, "m_tutup": m_tutup
            }
            
            file_excel = bina_excel(hasil, data_input)
            
            st.success("🎉 RPH Ulangkaji Berjaya Dijana!")
            st.download_button(
                label="📥 Muat Turun RPH Ulangkaji (Excel)",
                data=file_excel,
                file_name=f"RPH_Ulangkaji_Fizik_{nama_kelas}_{tarikh_pilihan}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
