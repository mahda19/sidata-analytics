SIDATA Analytics V2 — ekstrak ZIP, buka folder di VS Code.
Jupyter: jalankan %cd "alamat folder SIDATA_Analytics_V2" lalu %pip install -r requirements.txt.
Cell berikut: import sys; kemudian !{sys.executable} -m streamlit run app.py --server.port 8502
Buka http://localhost:8502. Biarkan server berjalan.
Unggah file per Bab. Tahun dibaca dari A3 tiap tabel; satu versi per tahun/bab (file pertama diterima). Filter bekerja pada tabel dan wilayah yang berhasil dipetakan. Tidak ada data bawaan atau akun login. File asli tidak diubah. Modul lama memerlukan tiga file 2026. Data disimpan hanya dalam sesi aplikasi. Jangan unggah data internal pada hosting publik.
Untuk hosting Streamlit Cloud: unggah source code saja termasuk requirements.txt, config dan modul ke GitHub, pilih app.py sebagai main file. URL diberikan setelah deployment. Gunakan data publik/simulasi; aplikasi belum memiliki autentikasi.
