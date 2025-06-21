import pandas as pd
import re
from openpyxl import Workbook
import os

# ===============================
# 1. Fungsi Normalisasi Nama Dosen
# ===============================
def normalisasi_nama_dosen(nama):
    nama = nama.split("//")[0].strip()
    nama = re.sub(r"\s+", " ", nama)
    nama = nama.replace(" ,", ",").replace(" .", ".").replace(",", ", ")

    gelar_map = {
        "s.sit": "S.Si.T.", "s.si.t": "S.Si.T.", "s.kom": "S.Kom.",
        "m.kom": "M.Kom.", "m.t": "M.T.", "mt": "M.T.", "m.stat": "M.Stat.",
        "m.mat": "M.Mat.", "ph.d": "Ph.D.", "drs.": "Drs.", "dr.": "Dr.",
        "st.": "S.T.", "s.t": "S.T.", "sp.": "Sp.", "mm": "M.M.", "m.m": "M.M."
    }

    nama_lower = nama.lower()
    for k, v in gelar_map.items():
        nama_lower = re.sub(rf"\b{k}\b", v.lower(), nama_lower)

    nama_parts = nama_lower.split()
    nama_bersih = [word.upper() if word.upper().endswith('.') else word.capitalize() for word in nama_parts]

    return ' '.join(nama_bersih)

# ===============================
# 2. Load Data Jadwal dari Excel
# ===============================
def load_kelas_dari_excel(file_excel):
    xls = pd.ExcelFile(file_excel)
    sheets = ['angkatan 2022', 'angkatan 2023', 'angkatan 2024']
    semua_kelas = {}

    def extract_kelas_dan_jadwal(df):
        kelas_dict = {}
        kelas_nama = None
        for _, row in df.iterrows():
            row = row.tolist()
            if isinstance(row[0], str) and "Kelas" in row[0]:
                kelas_nama = row[0].split(":")[1].strip().split(" ")[0]
                kelas_dict.setdefault(kelas_nama, [])
            elif len(row) >= 8 and isinstance(row[2], str) and kelas_nama:
                kelas_dict[kelas_nama].append({
                    "mata_kuliah": row[2],
                    "hari": row[4] if pd.notna(row[4]) else "",
                    "jam": row[5] if pd.notna(row[5]) else "",
                    "dosen": row[7] if pd.notna(row[7]) else ""
                })
        return kelas_dict

    for sheet in sheets:
        df = xls.parse(sheet)
        kelas_data = extract_kelas_dan_jadwal(df)
        semua_kelas.update(kelas_data)

    return semua_kelas, sheets

# ===============================
# 3. Daftar Ruangan & Jadwal
# ===============================
daftar_ruangan = [
    {"gedung": "A", "lantai": 4, "ruangan": "Lab Software"},
    {"gedung": "A", "lantai": 4, "ruangan": "Lab Hardware"},
    *[{"gedung": "B", "lantai": 4, "ruangan": f"B4{huruf}"} for huruf in "ABCDEFGH"]
]
jadwal_terisi = []

# ===============================
# 4. Validasi Jam Istirahat
# ===============================
def validasi_jam_istirahat(jam):
    waktu_istirahat = [("12:00", "13:00"), ("16:00", "18:00")]
    if " - " not in jam:
        print("Format jam tidak valid. Gunakan format 'HH:MM - HH:MM'")
        return False
    try:
        mulai, selesai = [s.strip() for s in jam.split("-")]
    except ValueError:
        print("Jam tidak dapat dipisahkan.")
        return False
    for istirahat_mulai, istirahat_selesai in waktu_istirahat:
        if mulai < istirahat_selesai and selesai > istirahat_mulai:
            print(f"Jadwal bentrok dengan waktu istirahat: {istirahat_mulai} - {istirahat_selesai}")
            return False
    return True

# ===============================
# 5. Cek Ruangan & Dosen Tersedia
# ===============================
def ruangan_tersedia(hari, jam):
    return [r for r in daftar_ruangan if not any(
        j for j in jadwal_terisi if j["gedung"] == r["gedung"] and j["lantai"] == r["lantai"]
        and j["ruangan"] == r["ruangan"] and j["hari"] == hari and j["jam"] == jam)]

def dosen_bentrok(dosen, hari, jam):
    return any(j for j in jadwal_terisi if j["dosen"] == dosen and j["hari"] == hari and j["jam"] == jam)

# ===============================
# 6. Booking Jadwal Kuliah
# ===============================
def input_booking(jadwal_kelas_excel):
    print("\n=== Booking Jadwal Kuliah Dosen ===")
    angkatan_map = {}
    for kls in jadwal_kelas_excel:
        angkatan = kls[:4]
        angkatan_map.setdefault(angkatan, []).append(kls)

    angkatan_list = sorted(angkatan_map)
    for i, ang in enumerate(angkatan_list, 1): print(f"{i}. {ang}")
    angkatan_pilih = angkatan_list[int(input("Pilih nomor angkatan: ")) - 1]
    kelas_list = sorted(angkatan_map[angkatan_pilih])
    for i, kls in enumerate(kelas_list, 1): print(f"{i}. {kls}")
    kelas = kelas_list[int(input("Pilih nomor kelas: ")) - 1]

    mk_list = sorted(set(mk['mata_kuliah'] for mk in jadwal_kelas_excel[kelas] if mk['mata_kuliah'].lower() != "mata kuliah"))
    for i, mk in enumerate(mk_list, 1): print(f"{i}. {mk}")
    mata_kuliah = mk_list[int(input("Pilih nomor mata kuliah: ")) - 1]
    hari = input("Hari (Senin - Minggu): ").strip()
    jam = input("Jam (contoh: 08:00 - 21:00): ").strip()
    if not validasi_jam_istirahat(jam): return

    dosen_set = {normalisasi_nama_dosen(mk['dosen']) for kelas_data in jadwal_kelas_excel.values()
                 for mk in kelas_data if mk['mata_kuliah'] and mata_kuliah.lower() in mk['mata_kuliah'].lower() and mk['dosen']}
    dosen_list = sorted(dosen_set)

    if not dosen_list:
        print("Tidak ditemukan dosen pengampu dari data Excel.")
        dosen = input("Masukkan nama dosen pengampu secara manual: ").strip()
    elif len(dosen_list) == 1:
        dosen = dosen_list[0]
        print(f"Dosen otomatis dipilih: {dosen}")
    else:
        for i, d in enumerate(dosen_list, 1): print(f"{i}. {d}")
        dosen = dosen_list[int(input("Pilih dosen (nomor): ")) - 1]

    if dosen_bentrok(dosen, hari, jam):
        print(f"Dosen {dosen} sudah memiliki jadwal di waktu tersebut.")
        return

    ruangan_opsi = ruangan_tersedia(hari, jam)
    if not ruangan_opsi:
        print("Tidak ada ruangan tersedia.")
        return
    for i, r in enumerate(ruangan_opsi, 1): print(f"{i}. Gedung {r['gedung']} - Lantai {r['lantai']} - Ruangan {r['ruangan']}")
    ruangan_pilih = ruangan_opsi[int(input("Pilih nomor ruangan: ")) - 1]

    jadwal_terisi.append({
        "kelas": kelas, "mata_kuliah": mata_kuliah, "dosen": dosen,
        "gedung": ruangan_pilih["gedung"], "lantai": ruangan_pilih["lantai"],
        "ruangan": ruangan_pilih["ruangan"], "hari": hari, "jam": jam
    })
    print("\nJadwal berhasil dibooking!")

# ===============================
# 7. Tampilkan & Export Jadwal
# ===============================
def tampilkan_jadwal():
    print("\n=== Daftar Jadwal Kuliah Terbooking ===")
    if not jadwal_terisi:
        print("(Kosong)")
    for i, j in enumerate(jadwal_terisi, 1):
        print(f"{i}. {j['dosen']} - {j['mata_kuliah']} ({j['kelas']}) di Gedung {j['gedung']} Lt.{j['lantai']} R.{j['ruangan']}, {j['hari']} {j['jam']}")

def export_jadwal_dengan_filter(nama_file="jadwal_terisi.xlsx"):
    if not jadwal_terisi:
        print("Belum ada data jadwal untuk diekspor.")
        return

    df = pd.DataFrame(jadwal_terisi).drop_duplicates()

    # Buat atau timpa file Excel dengan semua data ke satu sheet
    with pd.ExcelWriter(nama_file, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, sheet_name="JadwalTerisi", index=False)

    print(f"Semua data berhasil diekspor ke file {nama_file}.")




# ===============================
# 8. Menu Utama
# ===============================
def menu():
    file_excel = "Mapping Jadwal Mengajar Prodi Teknik Informatika.xlsx"
    try:
        jadwal_kelas_excel, _ = load_kelas_dari_excel(file_excel)
    except Exception as e:
        print(f"Gagal memuat file Excel: {e}")
        return

    while True:
        print("\n=== MENU UTAMA ===")
        print("1. Booking Jadwal Kuliah")
        print("2. Lihat Jadwal Yang Sudah Dibooking")
        print("3. Export Jadwal ke Excel")
        print("4. Keluar")
        pilihan = input("Pilih menu (1-4): ")

        if pilihan == "1":
            input_booking(jadwal_kelas_excel)
        elif pilihan == "2":
            tampilkan_jadwal()
        elif pilihan == "3":
            export_jadwal_dengan_filter()
        elif pilihan == "4":
            print("Keluar dari aplikasi.")
            break
        else:
            print("Pilihan tidak valid.")

# ===============================
# 9. Jalankan Program
# ===============================
if __name__ == "__main__":
    menu()
