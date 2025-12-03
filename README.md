# Dokumentasi Proyek: Excel Data Processor & Odoo Integration

## Deskripsi Proyek
Aplikasi ini adalah alat otomatisasi desktop yang dirancang untuk memproses data pengujian beton (Benda Uji), mengintegrasikannya dengan file Excel, database PostgreSQL, dan sistem Odoo (via web automation). Aplikasi ini memiliki antarmuka grafis (GUI) modern berbasis `customtkinter` yang memudahkan pengguna untuk mengelola konfigurasi, menjalankan proses data, dan memantau log aktivitas.

### Fitur Utama
1.  **Pemrosesan Data Excel**:
    *   Membaca dan memproses file Excel "Rencana" dan "Pengujian".
    *   Menghasilkan nilai beban acak (randomized load generation) berdasarkan Mutu Beton, Umur, dan Docket, dengan logika spesifik untuk berbagai kelas mutu (misal: B-2, K-400, Fc-30).
    *   Mendeteksi duplikasi data.
2.  **Integrasi Database**:
    *   Koneksi ke database PostgreSQL (`alatujidb`).
    *   Kemampuan untuk mengunggah data hasil proses ke tabel database.
3.  **Otomatisasi Web (Selenium)**:
    *   Fitur "Input Rencana Benda Uji" yang mengotomatiskan input data ke sistem web (`https://rmc.adhimix.web.id/benda_uji/`).
    *   Berjalan di background thread agar GUI tetap responsif.
4.  **Sinkronisasi Odoo**:
    *   Fitur sinkronisasi data dengan sistem Odoo.
5.  **Grid View**:
    *   Modul terpisah untuk menampilkan data dalam bentuk Grid (Tabel) untuk pemantauan yang lebih mudah.

## Requirements (Persyaratan Sistem)

### Lingkungan
*   **Sistem Operasi**: Windows (disarankan, karena penggunaan path dan driver).
*   **Python**: Versi 3.8 atau lebih baru.
*   **Database**: PostgreSQL.
*   **Browser**: Google Chrome (untuk fitur otomatisasi Selenium).

### Pustaka Python (Dependencies)
Pastikan pustaka berikut terinstall:
*   `tkinter` (bawaan Python)
*   `customtkinter` (UI modern)
*   `pandas` (Manipulasi data)
*   `openpyxl` (Baca/Tulis Excel)
*   `xlwings` (Interaksi Excel tingkat lanjut)
*   `sqlalchemy` (ORM Database)
*   `psycopg2` atau `psycopg2-binary` (Driver PostgreSQL)
*   `selenium` (Otomatisasi Web)
*   `requests` (HTTP Requests)
*   `webdriver_manager` (Manajemen driver Chrome otomatis - opsional tapi disarankan)

## Instalasi dan Konfigurasi

1.  **Clone/Copy Repository**: Pastikan seluruh folder proyek tersimpan di lokal.
2.  **Install Dependencies**:
    ```bash
    pip install customtkinter pandas openpyxl xlwings sqlalchemy psycopg2-binary selenium requests webdriver_manager
    ```
3.  **Konfigurasi Database**:
    *   Pastikan PostgreSQL berjalan.
    *   Buat database `alatujidb` (atau sesuaikan di config).
4.  **Konfigurasi Aplikasi (`config.cnf`)**:
    File `config.cnf` menyimpan pengaturan dasar. Sesuaikan isinya:
    ```ini
    [data]
    database = 'alatujidb'
    host = 'localhost'
    user = 'postgres'
    password = 'password_anda'
    
    [webser]
    webser_bendaUji = https://rmc.adhimix.web.id/benda_uji/?doc_no=
    ```

## Panduan Penggunaan

Jalankan aplikasi dengan perintah:
```bash
python main.py
```

### Tab Menu
1.  **📁 File Configuration**:
    *   Pilih file Excel sumber (Rencana) dan target (Pengujian).
    *   Tentukan folder output.
    *   Konfigurasi nama sheet yang akan diproses.
2.  **🗄️ Database Configure**:
    *   Masukkan kredensial database (User, Password, Host, Port, DB Name).
    *   Gunakan tombol "Test Connection" untuk memverifikasi koneksi.
3.  **⚙️ Process Control**:
    *   **Start Full Process**: Menjalankan seluruh alur pemrosesan data.
    *   **Copy Data**: Hanya menyalin data dari sumber ke target.
    *   **Generate CSV**: Membuat file CSV dari data yang diproses.
    *   **Upload to DB**: Mengunggah data ke database PostgreSQL.
    *   **Open Grid Benda Uji**: Membuka jendela terpisah untuk melihat data dalam bentuk tabel.
4.  **🧾 Input Rencana Benda Uji**:
    *   Klik "Start Process" untuk memulai otomatisasi input data ke web RMC.
    *   Proses berjalan di latar belakang, klik "Stop" untuk membatalkan.
5.  **📝 Process Logs**:
    *   Memantau log aktivitas aplikasi secara real-time.

## Struktur Proyek

```
odooCompressionTest/
├── main.py                     # Entry point aplikasi
├── config.cnf                  # File konfigurasi
├── modules/
│   ├── excel_handler.py        # Logika pemrosesan Excel & perhitungan beban
│   ├── input_rencana_benda_uji.py # Script otomatisasi Selenium
│   ├── grid_benda_uji.py       # Tampilan Grid/Tabel data
│   ├── db_controller.py        # Kontroler database
│   ├── daemon_sync.py          # Logika sinkronisasi background
│   ├── selenium_helpers.py     # Helper untuk Selenium
│   └── ui/
│       └── main_window.py      # Kode utama antarmuka GUI (ExcelProcessorGUI)
├── Data Input Pengujian/       # Folder data input
├── Output/                     # Folder output
└── logs/                       # (Opsional) Folder log aplikasi
```

## Catatan Pengembang
*   **Logika Beban**: Perhitungan beban (Load) terdapat di `modules/excel_handler.py` class `ExcelBebanProcessor`. Logika ini sangat spesifik berdasarkan jenis mutu beton dan umur (7 vs 28 hari).
*   **Threading**: Operasi berat seperti input web dan pemrosesan data besar dijalankan di thread terpisah untuk mencegah GUI membeku (Not Responding).
