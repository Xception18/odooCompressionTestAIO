# Dokumentasi Proyek: Excel Data Processor & Odoo Integration

## Deskripsi Proyek
Aplikasi desktop otomatisasi untuk memproses data pengujian beton (Benda Uji), mengintegrasikannya dengan file Excel, database PostgreSQL, dan sistem Odoo via web automation. GUI modern berbasis `customtkinter`.

### Fitur Utama
1.  **Pemrosesan Data Excel**:
    *   Membaca file Excel "Rencana" dan "Pengujian".
    *   Menghasilkan nilai beban (randomized load generation) berdasarkan Mutu Beton, Umur, dan Docket — logika spesifik untuk kelas mutu B-2, K-400, Fc-30, dll.
    *   Mendeteksi duplikasi data.
2.  **Integrasi Database**:
    *   Koneksi PostgreSQL (`alatujidb`) via **connection pooling** (SQLAlchemy `QueuePool`).
    *   Upload data hasil proses ke tabel database.
3.  **Otomatisasi Web**:
    *   **Selenium**: Input data ke `https://rmc.adhimix.web.id/benda_uji/`.
    *   **Playwright**: Alternatif helper otomatisasi web (`playwright_helpers.py`).
    *   Berjalan di background thread agar GUI tetap responsif.
4.  **Sinkronisasi Odoo**:
    *   Sinkronisasi data background via `daemon_sync.py`.
5.  **Graceful Shutdown**:
    *   `shutdown_manager.py` memastikan semua thread dan koneksi DB ditutup bersih saat aplikasi keluar.
6.  **Grid View**:
    *   Tampilan data dalam bentuk tabel (`grid_benda_uji.py`).

## Requirements (Persyaratan Sistem)

### Lingkungan
*   **OS**: Windows (disarankan).
*   **Python**: 3.8+
*   **Database**: PostgreSQL.
*   **Browser**: Google Chrome (Selenium/Playwright).

### Pustaka Python (Dependencies)
Lihat [`requirements.txt`](requirements.txt) untuk versi lengkap.

| Kategori | Paket |
|---|---|
| GUI | `customtkinter` |
| Data | `pandas`, `openpyxl`, `xlwings` |
| Database | `sqlalchemy`, `psycopg2-binary` |
| Web Automation | `selenium`, `webdriver-manager` |
| HTTP | `requests` |
| Config | `python-dotenv` |
| Testing | `pytest`, `pytest-cov`, `pytest-mock`, `pytest-asyncio` |

## Instalasi dan Konfigurasi

1.  **Clone/Copy Repository**.
2.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
3.  **Konfigurasi Environment Variables (`.env`)**:
    Salin `.env` dan sesuaikan nilainya:
    ```ini
    # Database
    DB_HOST=localhost
    DB_PORT=5432
    DB_NAME=alatujidb
    DB_USER=postgres
    DB_PASSWORD=password_anda

    # Web Automation
    SELENIUM_URL=https://rmc.adhimix.web.id/web/login
    SELENIUM_USER=email_anda@example.com
    SELENIUM_PASSWORD=password_anda

    # Opsional
    LOG_LEVEL=INFO
    ENVIRONMENT=production
    DB_POOL_SIZE=5
    DB_POOL_OVERFLOW=10
    DB_CONNECT_TIMEOUT=10
    DB_STATEMENT_TIMEOUT=30
    ```
    > **⚠️ Keamanan**: Jangan commit `.env` ke version control (sudah di `.gitignore`).

4.  **Setup Database**:
    ```bash
    psql -U postgres -c "CREATE DATABASE alatujidb;"
    psql -U postgres -d alatujidb -f database_schema.sql
    ```

## Panduan Penggunaan

```bash
python main.py
```

### Tab Menu
| Tab | Fungsi |
|---|---|
| 📁 **File Configuration** | Pilih file Excel sumber/target, folder output, nama sheet |
| 🗄️ **Database Configure** | Masukkan kredensial DB, uji koneksi |
| ⚙️ **Process Control** | Start Full Process, Copy Data, Generate CSV, Upload to DB, buka Grid |
| 🧾 **Input Rencana Benda Uji** | Otomatisasi input data ke web RMC (start/stop) |
| 📝 **Process Logs** | Monitor log aktivitas real-time |

## Struktur Proyek

```
odooCompressionTest/
├── main.py                          # Entry point
├── .env                             # Konfigurasi environment (jangan dicommit)
├── config.cnf                       # Config lama (deprecated, gunakan .env)
├── requirements.txt                 # Daftar dependencies
├── database_schema.sql              # Schema PostgreSQL
├── pytest.ini                       # Konfigurasi pytest
├── IMPROVEMENTS_ROADMAP.md          # Roadmap pengembangan
├── modules/
│   ├── config.py                    # Manajemen konfigurasi (baca .env)
│   ├── logging_config.py            # Setup logging terpusat
│   ├── exceptions.py                # Custom exceptions
│   ├── utils.py                     # Utility functions
│   ├── excel_handler.py             # Logika pemrosesan Excel & perhitungan beban
│   ├── data_generators.py           # Generator data (random load, dll.)
│   ├── db_controller.py             # Kontroler database (CRUD)
│   ├── db_pool.py                   # Connection pool SQLAlchemy
│   ├── daemon_sync.py               # Sinkronisasi background Odoo
│   ├── shutdown_manager.py          # Graceful shutdown manager
│   ├── input_rencana_benda_uji.py   # Otomatisasi Selenium
│   ├── selenium_helpers.py          # Helper Selenium
│   ├── playwright_helpers.py        # Helper Playwright
│   ├── grid_benda_uji.py            # Tampilan Grid/Tabel
│   └── ui/
│       └── main_window.py           # Antarmuka GUI utama (ExcelProcessorGUI)
├── Data Input Pengujian/            # Folder data input Excel
├── Output/                          # Folder output hasil proses
├── Rencana/                         # Folder file Rencana
└── Pengujian/                       # Folder file Pengujian
```

## Arsitektur Teknis

### Database Connection Pooling
`db_pool.py` menggunakan SQLAlchemy `QueuePool`:
- Pool size: 5 (via `DB_POOL_SIZE`)
- Max overflow: 10 (via `DB_POOL_OVERFLOW`)
- Connection recycle: 3600 detik

### Threading Model
- **Main thread**: GUI (`tkinter`/`customtkinter`)
- **Worker threads**: Selenium, pemrosesan Excel, upload DB
- **Daemon thread**: Sinkronisasi Odoo (`daemon_sync.py`)
- **Shutdown**: `shutdown_manager.py` join semua thread + tutup pool saat `atexit`

### Logging
Terpusat via `modules/logging_config.py`. Level dikontrol via `LOG_LEVEL` di `.env`.

## Testing

```bash
# Semua test
pytest

# Dengan coverage
pytest --cov=modules --cov-report=term-missing
```

> **Status**: 15/18 tests pass (83%). Lihat [`IMPROVEMENTS_ROADMAP.md`](IMPROVEMENTS_ROADMAP.md) untuk rencana peningkatan coverage.

## Catatan Pengembang
*   **Logika Beban**: Di `modules/excel_handler.py` class `ExcelBebanProcessor`. Sangat spesifik per jenis mutu beton dan umur (7 vs 28 hari).
*   **Playwright vs Selenium**: Keduanya tersedia. `playwright_helpers.py` sebagai alternatif untuk halaman dinamis.
*   **`config.cnf`**: Deprecated. Semua config kini via `.env` + `modules/config.py`.
*   **`chromedriver.exe`**: Di-bundle di root. Update manual jika versi Chrome berubah, atau aktifkan `webdriver-manager` untuk auto-manage.
*   **Threading**: Operasi berat dijalankan di thread terpisah untuk mencegah GUI Not Responding.
