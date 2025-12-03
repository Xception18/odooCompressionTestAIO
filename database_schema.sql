-- Script untuk membuat struktur database alatujidb
-- Database: alatujidb

-- 1. Buat Database (jika belum ada)
-- CREATE DATABASE alatujidb;

-- 2. Gunakan Database
-- \c alatujidb;

-- 3. Hapus tabel jika sudah ada (opsional, hati-hati menggunakan ini)
-- DROP TABLE IF EXISTS pengujian;

-- 4. Buat Tabel 'pengujian'
CREATE TABLE IF NOT EXISTS pengujian (
    idpengujian SERIAL PRIMARY KEY,           -- Auto-incrementing ID
    tgluji DATE,                              -- Tanggal Pengujian
    idalat CHAR(100),                         -- ID Alat
    kodebendauji CHAR(200),                   -- Kode Benda Uji
    nodocket CHAR(100),                       -- Nomor Docket
    nourutbenda CHAR(2),                      -- Nomor Urut Benda
    nilaikn NUMERIC(10, 2),                   -- Nilai kN
    beratbenda NUMERIC(10, 2),                -- Berat Benda
    tiperetak CHAR(1),                        -- Tipe Retak
    sinkron CHAR(1) DEFAULT 'B',              -- Status Sinkronisasi (B/Belum, S/Sudah)
    idbendauji CHAR(100),                     -- ID Benda Uji (dari sistem lain/Odoo)
    tglrencanauji DATE,                       -- Tanggal Rencana Uji
    bujnama CHAR(20),                         -- Nama/Jenis Benda Uji
    kuattekan NUMERIC(10, 2),                 -- Kuat Tekan
    bebanmpa NUMERIC(10, 2),                  -- Beban MPa
    umur INTEGER,                             -- Umur Beton (hari)
    tglbendauji DATE                          -- Tanggal Benda Uji
);

-- 5. Tambahkan Index (Opsional, untuk performa pencarian)
CREATE INDEX IF NOT EXISTS idx_pengujian_tgluji ON pengujian(tgluji);
CREATE INDEX IF NOT EXISTS idx_pengujian_nodocket ON pengujian(nodocket);
CREATE INDEX IF NOT EXISTS idx_pengujian_sinkron ON pengujian(sinkron);

-- Keterangan Kolom berdasarkan source code (modules/db_controller.py):
-- idpengujian  : Primary Key (Serial)
-- tgluji       : Tanggal pengujian dilakukan
-- idalat       : Identitas alat uji
-- kodebendauji : Kode unik benda uji
-- nodocket     : Nomor docket pengiriman
-- nourutbenda  : Nomor urut benda uji dalam satu sampel
-- nilaikn      : Hasil pembacaan beban (kN)
-- beratbenda   : Berat benda uji
-- tiperetak    : Kode tipe keretakan
-- sinkron      : Flag sinkronisasi ke Odoo (misal: 'N' belum, 'Y' sudah)
-- idbendauji   : ID referensi benda uji
-- tglrencanauji: Tanggal rencana pengujian
-- bujnama      : Jenis benda uji
-- kuattekan    : Nilai kuat tekan
-- bebanmpa     : Nilai beban dalam MPa
-- umur         : Umur beton saat diuji
-- tglbendauji  : Tanggal pembuatan benda uji
