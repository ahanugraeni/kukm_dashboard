import streamlit as st
import pandas as pd
import plotly.express as px
import traceback

st.set_page_config(page_title="Dashboard Analisis KUKM Prov Jatim", layout="wide")
st.title("📊 Dashboard Peserta Pelatihan Provinsi Jatim (2021-2024)")

def clean_rupiah(val):
    try:
        if pd.isna(val): return 0
        return int(str(val).replace("Rp.", "").replace("Rp", "").replace(",", "").replace(".", "").replace("-", "0").strip())
    except:
        return 0

try:
    # baca file Excel
    xls = pd.ExcelFile("2024 DATA ALUMNI UPT.xlsx")
    sheet_names = xls.sheet_names
    selected_sheet = st.selectbox("📁 Pilih sheet yang ingin dianalisis:", sheet_names)

    df = pd.read_excel(xls, sheet_name=selected_sheet)

    with st.expander("📄 Tampilkan Data Mentah"):
        st.dataframe(df)

    # bersihkan kolom rupiah
    rupiah_cols = [
        "OMZET KOPERASI",
        "SHU KOPERASI TAHUN BERJALAN/31 DESEMBER",
        "MODAL USAHA KOPERASI (MODAL SENDIRI)",
        "NILAI OMZET USAHA PER TAHUN",
    ]
    for col in rupiah_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_rupiah)

    # sheet Peserta UMKM
    if selected_sheet.strip().lower() == "peserta umkm":
        st.markdown("### 📈 Ringkasan Finansial (Peserta UMKM)")
        col1, col2 = st.columns(2)

        omzet_umkm_total = df["NILAI OMZET USAHA PER TAHUN"].sum() if "NILAI OMZET USAHA PER TAHUN" in df.columns else 0
        total_data_umkm = len(df)

        col1.metric("Total Omzet Peserta UMKM", f"Rp {int(omzet_umkm_total):,}")
        col2.metric("Total Data", f"{total_data_umkm}")

        #melihat persentase penyandang disability yang aktif di bidang UMKM
        if "APAKAH ANDA PENYANDANG DISABILITAS" in df.columns:
            st.markdown("### ♿ Informasi Peserta Penyandang Disabilitas")
            dis_chart = df["APAKAH ANDA PENYANDANG DISABILITAS"].value_counts().reset_index()
            dis_chart.columns = ["Status", "Jumlah"]
            st.plotly_chart(px.pie(dis_chart, names="Status", values="Jumlah"), use_container_width=True)

        #melihat persentasi berdasarkan gender karyawan yang aktif di bidang UMKM
        if "JUMLAH KARYAWAN LAKI-LAKI" in df.columns and "JUMLAH KARYAWAN PEREMPUAN" in df.columns:
            st.markdown("### 👥 Komposisi Karyawan")
            laki = pd.to_numeric(df["JUMLAH KARYAWAN LAKI-LAKI"], errors='coerce').sum()
            perempuan = pd.to_numeric(df["JUMLAH KARYAWAN PEREMPUAN"], errors='coerce').sum()
            st.plotly_chart(px.pie(names=["Laki-laki", "Perempuan"], values=[laki, perempuan]), use_container_width=True)

        #menetukan omzet UMKM pertahun di setiap kabupaten 
        if "KABUPATEN USAHA" in df.columns and "NILAI OMZET USAHA PER TAHUN" in df.columns:
            st.markdown("### 💼 Total Omzet per Kabupaten (Peserta UMKM)")
            omzet_kab = df.groupby("KABUPATEN USAHA")["NILAI OMZET USAHA PER TAHUN"].sum().reset_index()
            omzet_kab.columns = ["Kabupaten", "Total Omzet"]
            omzet_kab = omzet_kab.sort_values(by="Total Omzet", ascending=False)
            omzet_kab["Total Omzet"] = omzet_kab["Total Omzet"].apply(lambda x: f"Rp {x:,}")
            st.dataframe(omzet_kab)
        
        #menerapkan filter by sertifikasi merk
        if "NAMA USAHA" in df.columns and "SERTIFIKASI USAHA" in df.columns:
            st.markdown("### 🏷️ UMKM dengan Sertifikasi Merk")
            merk_df = df[df["SERTIFIKASI USAHA"].str.contains("MERK", case=False, na=False)]

            nama_filter = st.text_input("🔍 Cari berdasarkan Nama Peserta:")
            if nama_filter:
                merk_df = merk_df[merk_df["NAMA PESERTA"].str.contains(nama_filter, case=False, na=False)]

            merk_df = merk_df.sort_values(by="SERTIFIKASI USAHA")
            st.dataframe(merk_df[["NAMA PESERTA", "NAMA USAHA", "SERTIFIKASI USAHA"]])

        #mengetahui jumlah yang pemasaran nya sudah ekspor/masih dilingkup regional
        if "WILAYAH PEMASARAN" in df.columns:
            st.markdown("### 📜 Persebaran Wilayah Pemasaran")
            wilayah = df["WILAYAH PEMASARAN"].value_counts().reset_index()
            wilayah.columns = ["Wilayah Pemasaran", "Jumlah"]
            st.plotly_chart(px.bar(wilayah, x="Wilayah Pemasaran", y="Jumlah"), use_container_width=True)
            
        #group by omzet dan nib yang bisa digunakan untuk keputusan rekomendasi usaha yang perlu di fasilitasi izin ekspor 
        if "NILAI OMZET USAHA PER TAHUN" in df.columns and "NO. NIB" in df.columns:
            st.markdown("### 🕵️ Pelaku UMKM dengan Omzet > Rp100 Juta dan Memiliki NO. NIB")
            filtered_df = df[
                (df["NILAI OMZET USAHA PER TAHUN"] > 100_000_000) &
                (df["NO. NIB"].notna()) &
                (df["NO. NIB"].astype(str).str.strip().str.match(r'^\d{13}$'))
            ]
            filtered_df = filtered_df.sort_values(by="NILAI OMZET USAHA PER TAHUN", ascending=False)
            filtered_df["NILAI OMZET USAHA PER TAHUN"] = filtered_df["NILAI OMZET USAHA PER TAHUN"].apply(lambda x: f"Rp {x:,}")
            st.dataframe(filtered_df[["NAMA PESERTA", "NAMA USAHA", "NO. NIB", "NILAI OMZET USAHA PER TAHUN"]])
        
        #untuk mengetahui status jabatan yang datang ke pelatihan 
        if "JABATAN PESERTA DI USAHA" in df.columns:
            st.markdown("### 🧑‍💼 Jabatan Peserta di Usaha")
            jabatan = df["JABATAN PESERTA DI USAHA"].value_counts().reset_index()
            jabatan.columns = ["Jabatan", "Jumlah"]
            st.plotly_chart(px.pie(jabatan, names="Jabatan", values="Jumlah"), use_container_width=True)

        #informasi jumlah pesertas disetiap bidang usaha

        if "BIDANG USAHA" in df.columns:
            st.markdown("### 📟 Jumlah per Bidang Usaha")
            bidang = df["BIDANG USAHA"].str.strip().str.upper().value_counts().reset_index()
            bidang.columns = ["Bidang Usaha", "Jumlah"]
            st.plotly_chart(px.bar(bidang, x="Bidang Usaha", y="Jumlah"), use_container_width=True)

        #mengetahui jumlah permasalah yang sama,  bisa untuk mengambil keputusan tentang pelatihan/solusi kedepannya
        if "PERMSALAHAN YANG DIHADAPI" in df.columns:
            st.markdown("### ⚠️ Permasalahan yang Dihadapi")
            masalah = df["PERMSALAHAN YANG DIHADAPI"].str.strip().str.upper().value_counts().head(5).reset_index()
            masalah.columns = ["Masalah", "Jumlah"]
            st.plotly_chart(px.bar(masalah, x="Masalah", y="Jumlah"), use_container_width=True)

        #mengetahui pelatihan yang banyak dikunjungi,
        if "KEBUTUHAN DIKLAT/PELATIHAN" in df.columns:
            st.markdown("### 🎯 Top 10 Kebutuhan Pelatihan")
            chart = df["KEBUTUHAN DIKLAT/PELATIHAN"].str.strip().str.upper().value_counts().head(10).reset_index()
            chart.columns = ["Jenis Pelatihan", "Jumlah"]
            st.plotly_chart(px.bar(chart, x="Jenis Pelatihan", y="Jumlah"), use_container_width=True)

    # sheet Peserta Koperasi
    if selected_sheet.strip().lower() == "peserta koperasi":
        st.markdown("### 📈 Ringkasan Finansial")
        col1, col2, col3 = st.columns(3)
        omzet_total = df["OMZET KOPERASI"].sum() if "OMZET KOPERASI" in df.columns else 0
        shu_total = df["SHU KOPERASI TAHUN BERJALAN/31 DESEMBER"].sum() if "SHU KOPERASI TAHUN BERJALAN/31 DESEMBER" in df.columns else 0
        col1.metric("Total Omzet", f"Rp {int(omzet_total):,}")
        col2.metric("Total SHU", f"Rp {int(shu_total):,}")
        col3.metric("Total Data", f"{len(df)}")
        
        #informasi gender dari bidang koperasi yang mengikuti pelatihan
        if "JENIS KELAMIN" in df.columns:
            st.markdown("### 🧝 Informasi Jenis Kelamin")
            chart = df["JENIS KELAMIN"].value_counts().reset_index()
            chart.columns = ["Jenis Kelamin", "Jumlah"]
            st.plotly_chart(px.pie(chart, names="Jenis Kelamin", values="Jumlah"), use_container_width=True)

        #informasi jenis koperasi yang aktif
        if "JENIS KOPERASI" in df.columns:
            st.markdown("### 🏢 Jenis Koperasi")
            chart = df["JENIS KOPERASI"].value_counts().reset_index()
            chart.columns = ["Jenis Koperasi", "Jumlah"]
            st.plotly_chart(px.pie(chart, names="Jenis Koperasi", values="Jumlah"), use_container_width=True)

        #menentukan jumlah koperasi yang terdata di setiap kabupaten
        if "KABUPATEN" in df.columns:
            st.markdown("### 🚩 Jumlah Koperasi per Kabupaten")
            kab = df["KABUPATEN"].value_counts().reset_index()
            kab.columns = ["Kabupaten", "Jumlah"]
            st.plotly_chart(px.bar(kab, x="Kabupaten", y="Jumlah"), use_container_width=True)
        
        #mengambil data SHU dan data OMZET yang dlaporkan Koperasi periode 2021-2024
        if "OMZET KOPERASI" in df.columns and "SHU KOPERASI TAHUN BERJALAN/31 DESEMBER" in df.columns and "NAMA KOPERASI" in df.columns:
            st.markdown("### 🏆 Koperasi dengan SHU Tertinggi")
            top_shu = df[["NAMA KOPERASI", "OMZET KOPERASI", "SHU KOPERASI TAHUN BERJALAN/31 DESEMBER"]].copy()
            top_shu["NAMA KOPERASI"] = top_shu["NAMA KOPERASI"].str.strip().str.upper()
            top_shu = top_shu.drop_duplicates(subset="NAMA KOPERASI")
            top_shu = top_shu.sort_values(by="SHU KOPERASI TAHUN BERJALAN/31 DESEMBER", ascending=False).head(10)
            top_shu["OMZET KOPERASI"] = top_shu["OMZET KOPERASI"].apply(lambda x: f"Rp {x:,}")
            top_shu["SHU KOPERASI TAHUN BERJALAN/31 DESEMBER"] = top_shu["SHU KOPERASI TAHUN BERJALAN/31 DESEMBER"].apply(lambda x: f"Rp {x:,}")
            st.dataframe(top_shu)

       #mengambil data SHU tapi modal kecil untuk pengambilan keputusan fasilitas permodalan
        if "MODAL USAHA KOPERASI (MODAL SENDIRI)" in df.columns:
            st.markdown("### 💰 Koperasi Efisien: SHU Tinggi, Modal Kecil")
            efisien = df[["NAMA KOPERASI", "MODAL USAHA KOPERASI (MODAL SENDIRI)", "SHU KOPERASI TAHUN BERJALAN/31 DESEMBER"]]
            efisien = efisien.drop_duplicates(subset="NAMA KOPERASI")
            efisien = efisien[(efisien["MODAL USAHA KOPERASI (MODAL SENDIRI)"] < 50000000) & (efisien["SHU KOPERASI TAHUN BERJALAN/31 DESEMBER"] > 0)]
            efisien["MODAL USAHA KOPERASI (MODAL SENDIRI)"] = efisien["MODAL USAHA KOPERASI (MODAL SENDIRI)"].apply(lambda x: f"Rp {x:,}")
            efisien["SHU KOPERASI TAHUN BERJALAN/31 DESEMBER"] = efisien["SHU KOPERASI TAHUN BERJALAN/31 DESEMBER"].apply(lambda x: f"Rp {x:,}")
            st.dataframe(efisien.sort_values(by="SHU KOPERASI TAHUN BERJALAN/31 DESEMBER", ascending=False).head(10))

        #mengetahui jumlah permasalahan yang banyak terjadi di bidang koperasi
        if "PERMSALAHAN YANG DIHADAPI" in df.columns:
            st.markdown("### ⚠️ Permasalahan yang Dihadapi")
            masalah = df["PERMSALAHAN YANG DIHADAPI"].value_counts().head(5).reset_index()
            masalah.columns = ["Masalah", "Jumlah"]
            st.plotly_chart(px.bar(masalah, x="Masalah", y="Jumlah"), use_container_width=True)
        
        #mengetahui jumlah pelatihan yang banyak diikuti bidang koperasi
        if "KEBUTUHAN DIKLAT/PELATIHAN" in df.columns:
            st.markdown("### 🎯 Top 10 Kebutuhan Pelatihan")
            chart = df["KEBUTUHAN DIKLAT/PELATIHAN"].value_counts().head(10).reset_index()
            chart.columns = ["Jenis Pelatihan", "Jumlah"]
            st.plotly_chart(px.bar(chart, x="Jenis Pelatihan", y="Jumlah"), use_container_width=True)

except Exception as e:
    st.error(f"❌ Terjadi error saat membaca atau memproses data: {e}")
    st.code(traceback.format_exc())
