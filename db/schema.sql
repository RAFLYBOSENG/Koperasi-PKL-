CREATE TABLE IF NOT EXISTS anggota (
    id_anggota TEXT PRIMARY KEY,
    no_anggota TEXT NOT NULL UNIQUE,
    nik TEXT,
    nama TEXT NOT NULL,
    alamat TEXT,
    no_telp TEXT,
    tgl_bergabung DATE,
    no_rekening TEXT,
    nama_bank TEXT,
    penghasilan_bersih NUMERIC(18,2) DEFAULT 0,
    cicilan_lain NUMERIC(18,2) DEFAULT 0,
    simpanan_pokok NUMERIC(18,2) DEFAULT 0
);

CREATE TABLE IF NOT EXISTS users (
    id_user TEXT PRIMARY KEY,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('admin', 'user')),
    id_anggota TEXT,
    created_at DATE,
    CONSTRAINT fk_users_anggota FOREIGN KEY (id_anggota) REFERENCES anggota (id_anggota) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS simpanan (
    id_anggota TEXT PRIMARY KEY,
    total_simpanan NUMERIC(18,2) NOT NULL DEFAULT 0,
    CONSTRAINT fk_simpanan_anggota FOREIGN KEY (id_anggota) REFERENCES anggota (id_anggota) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS simpanan_transaksi (
    id_transaksi TEXT PRIMARY KEY,
    id_anggota TEXT NOT NULL,
    no_anggota TEXT,
    nama_anggota TEXT,
    tanggal TIMESTAMP,
    jenis_simpanan TEXT,
    jumlah NUMERIC(18,2) NOT NULL DEFAULT 0,
    keterangan TEXT,
    diajukan_oleh TEXT,
    CONSTRAINT fk_simpanan_transaksi_anggota FOREIGN KEY (id_anggota) REFERENCES anggota (id_anggota) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS simpanan_pengajuan (
    id_pengajuan TEXT PRIMARY KEY,
    id_anggota TEXT NOT NULL,
    no_anggota TEXT,
    nama_anggota TEXT,
    tanggal_pengajuan TIMESTAMP,
    jenis_simpanan TEXT,
    jumlah NUMERIC(18,2) NOT NULL DEFAULT 0,
    keterangan TEXT,
    status TEXT,
    tanggal_konfirmasi TIMESTAMP,
    dikonfirmasi_oleh TEXT,
    diajukan_oleh TEXT,
    CONSTRAINT fk_simpanan_pengajuan_anggota FOREIGN KEY (id_anggota) REFERENCES anggota (id_anggota) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS iuran_sosial (
    id_iuran TEXT PRIMARY KEY,
    id_anggota TEXT NOT NULL,
    no_anggota TEXT,
    nama_anggota TEXT,
    tanggal TIMESTAMP,
    jumlah NUMERIC(18,2) NOT NULL DEFAULT 0,
    keterangan TEXT,
    diajukan_oleh TEXT,
    CONSTRAINT fk_iuran_sosial_anggota FOREIGN KEY (id_anggota) REFERENCES anggota (id_anggota) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS pinjaman (
    id_pinjaman TEXT PRIMARY KEY,
    id_anggota TEXT NOT NULL,
    nama_anggota TEXT,
    no_anggota TEXT,
    jenis_pinjaman TEXT,
    jenis_simpanan TEXT,
    plafon NUMERIC(18,2) NOT NULL DEFAULT 0,
    tenor_awal INTEGER,
    tenor_bulan INTEGER,
    bunga_persen NUMERIC(10,4) DEFAULT 0,
    total_bayar NUMERIC(18,2) DEFAULT 0,
    cicilan_per_bulan NUMERIC(18,2) DEFAULT 0,
    sisa_pinjaman NUMERIC(18,2) DEFAULT 0,
    tanggal_pengajuan TIMESTAMP,
    status TEXT,
    tanggal_lunas DATE,
    CONSTRAINT fk_pinjaman_anggota FOREIGN KEY (id_anggota) REFERENCES anggota (id_anggota) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS pinjaman_cicilan (
    id_cicilan TEXT PRIMARY KEY,
    id_pinjaman TEXT NOT NULL,
    id_anggota TEXT NOT NULL,
    no_anggota TEXT,
    nama_anggota TEXT,
    jumlah NUMERIC(18,2) NOT NULL DEFAULT 0,
    tanggal_pengajuan TIMESTAMP,
    status TEXT,
    tanggal_konfirmasi TIMESTAMP,
    dikonfirmasi_oleh TEXT,
    diajukan_oleh TEXT,
    keterangan TEXT,
    metode_pembayaran TEXT,
    detail_pembayaran TEXT,
    status_transaksi TEXT,
    va_number TEXT,
    idempotency_key TEXT,
    periode_tagihan TEXT,
    expires_at TIMESTAMP,
    CONSTRAINT fk_pinjaman_cicilan_pinjaman FOREIGN KEY (id_pinjaman) REFERENCES pinjaman (id_pinjaman) ON DELETE CASCADE,
    CONSTRAINT fk_pinjaman_cicilan_anggota FOREIGN KEY (id_anggota) REFERENCES anggota (id_anggota) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS pendaftaran_anggota (
    id_pengajuan TEXT PRIMARY KEY,
    nama TEXT NOT NULL,
    alamat TEXT,
    no_telp TEXT,
    penghasilan_bersih NUMERIC(18,2) DEFAULT 0,
    cicilan_lain NUMERIC(18,2) DEFAULT 0,
    simpanan_pokok NUMERIC(18,2) DEFAULT 0,
    status TEXT,
    tanggal_pengajuan DATE,
    catatan_admin TEXT,
    id_anggota_dibuat TEXT,
    no_anggota_dibuat TEXT
);

CREATE TABLE IF NOT EXISTS import_log (
    id BIGSERIAL PRIMARY KEY,
    waktu TIMESTAMP NOT NULL DEFAULT NOW(),
    "user" TEXT,
    mode TEXT,
    nama_file TEXT,
    berhasil INTEGER DEFAULT 0,
    gagal INTEGER DEFAULT 0,
    catatan TEXT
);
