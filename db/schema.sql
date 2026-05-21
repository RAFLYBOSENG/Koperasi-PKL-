/*CREATE TABLE IF NOT EXISTS anggota (
    id_anggota TEXT PRIMARY KEY,
    no_anggota TEXT NOT NULL UNIQUE,
    nik TEXT,
    nama_lengkap TEXT NOT NULL,
    email TEXT,
    no_hp TEXT,
    alamat TEXT,
    kategori_anggota TEXT,
    tgl_bergabung DATE,
    status_anggota TEXT DEFAULT 'Aktif',
    status_kredit TEXT DEFAULT 'Lancar',
    catatan_kredit TEXT,
    no_rekening TEXT,
    nama_bank TEXT,
    penghasilan_bersih NUMERIC(18,2) DEFAULT 0,
    cicilan_lain NUMERIC(18,2) DEFAULT 0,
    simpanan_pokok NUMERIC(18,2) DEFAULT 0,
    foto_ktp TEXT,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);*/

-- Master: Roles
CREATE TABLE IF NOT EXISTS roles (
    id_role TEXT PRIMARY KEY,
    role_name TEXT NOT NULL UNIQUE,
    deskripsi TEXT,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- Master: Permissions
CREATE TABLE IF NOT EXISTS permissions (
    id_permission TEXT PRIMARY KEY,
    permission_name TEXT NOT NULL UNIQUE,
    deskripsi TEXT,
    created_at TIMESTAMP DEFAULT NOW()
);

-- Relasi: Role -> Permissions
CREATE TABLE IF NOT EXISTS role_permissions (
    id_role TEXT NOT NULL,
    id_permission TEXT NOT NULL,
    assigned_at TIMESTAMP DEFAULT NOW(),
    PRIMARY KEY (id_role, id_permission),
    CONSTRAINT fk_role_permissions_role FOREIGN KEY (id_role) REFERENCES roles (id_role) ON DELETE CASCADE,
    CONSTRAINT fk_role_permissions_permission FOREIGN KEY (id_permission) REFERENCES permissions (id_permission) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS users (
    id_user TEXT PRIMARY KEY,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL,
    id_anggota TEXT,
    created_at DATE,
    updated_at TIMESTAMP DEFAULT NOW(),
    is_active BOOLEAN DEFAULT TRUE,
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
    tanggal_pencairan DATE,
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

-- Master: Jenis Simpanan
CREATE TABLE IF NOT EXISTS jenis_simpanan (
    id_jenis_simpanan TEXT PRIMARY KEY,
    nama_jenis TEXT NOT NULL UNIQUE,
    nominal_tetap NUMERIC(18,2),
    min_nominal NUMERIC(18,2),
    max_nominal NUMERIC(18,2),
    bisa_ditarik BOOLEAN DEFAULT TRUE,
    masuk_saldo BOOLEAN DEFAULT TRUE,
    tipe_pembayaran TEXT,
    frekuensi TEXT,
    urutan_display INTEGER DEFAULT 0,
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- Master: Jenis Pinjaman
CREATE TABLE IF NOT EXISTS jenis_pinjaman (
    id_jenis_pinjaman TEXT PRIMARY KEY,
    nama_jenis TEXT NOT NULL UNIQUE,
    min_tenor INTEGER,
    max_tenor INTEGER,
    bunga_persen NUMERIC(10,4),
    sistem_bunga TEXT DEFAULT 'Flat',
    ada_provisi BOOLEAN DEFAULT FALSE,
    persentase_provisi NUMERIC(10,4),
    min_tenor_provisi INTEGER,
    urutan_display INTEGER DEFAULT 0,
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- Operasional: Jadwal Angsuran (Cicilan)
CREATE TABLE IF NOT EXISTS jadwal_angsuran (
    id_jadwal TEXT PRIMARY KEY,
    id_pinjaman TEXT NOT NULL,
    id_anggota TEXT NOT NULL,
    angsuran_ke INTEGER,
    tanggal_jatuh_tempo DATE,
    jumlah_pokok NUMERIC(18,2),
    jumlah_bunga NUMERIC(18,2),
    jumlah_total NUMERIC(18,2),
    status TEXT DEFAULT 'Belum Jatuh Tempo',
    catatan TEXT,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW(),
    CONSTRAINT fk_jadwal_angsuran_pinjaman FOREIGN KEY (id_pinjaman) REFERENCES pinjaman (id_pinjaman) ON DELETE CASCADE,
    CONSTRAINT fk_jadwal_angsuran_anggota FOREIGN KEY (id_anggota) REFERENCES anggota (id_anggota) ON DELETE CASCADE
);

-- Operasional: Pembayaran Angsuran
CREATE TABLE IF NOT EXISTS pembayaran_angsuran (
    id_pembayaran TEXT PRIMARY KEY,
    id_jadwal TEXT NOT NULL,
    id_pinjaman TEXT NOT NULL,
    id_anggota TEXT NOT NULL,
    tanggal_pembayaran TIMESTAMP,
    jumlah_pembayaran NUMERIC(18,2),
    metode_pembayaran TEXT,
    bukti_transfer TEXT,
    status TEXT DEFAULT 'Menunggu Validasi',
    divalidasi_oleh TEXT,
    tanggal_validasi TIMESTAMP,
    catatan TEXT,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW(),
    CONSTRAINT fk_pembayaran_angsuran_jadwal FOREIGN KEY (id_jadwal) REFERENCES jadwal_angsuran (id_jadwal) ON DELETE CASCADE,
    CONSTRAINT fk_pembayaran_angsuran_pinjaman FOREIGN KEY (id_pinjaman) REFERENCES pinjaman (id_pinjaman) ON DELETE CASCADE,
    CONSTRAINT fk_pembayaran_angsuran_anggota FOREIGN KEY (id_anggota) REFERENCES anggota (id_anggota) ON DELETE CASCADE
);

-- Operasional: Kas Koperasi
CREATE TABLE IF NOT EXISTS kas_koperasi (
    id_kas TEXT PRIMARY KEY,
    tanggal_transaksi TIMESTAMP,
    jenis_transaksi TEXT,
    kategori_kas TEXT,
    deskripsi TEXT,
    nominal NUMERIC(18,2) NOT NULL DEFAULT 0,
    saldo_setelah NUMERIC(18,2),
    referensi_id TEXT,
    referensi_tabel TEXT,
    diinput_oleh TEXT,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- SHU: Perkiraan SHU Tahunan
CREATE TABLE IF NOT EXISTS shu_tahunan (
    id_shu TEXT PRIMARY KEY,
    tahun INTEGER NOT NULL UNIQUE,
    tanggal_input TIMESTAMP DEFAULT NOW(),
    total_shu NUMERIC(18,2),
    cadangan_umum NUMERIC(18,2),
    shu_pasif_total NUMERIC(18,2),
    shu_aktif_total NUMERIC(18,2),
    dana_kesejahteraan NUMERIC(18,2),
    dana_pendidikan NUMERIC(18,2),
    dana_sosial NUMERIC(18,2),
    dana_pembangunan NUMERIC(18,2),
    dana_pengurus NUMERIC(18,2),
    dana_risiko NUMERIC(18,2),
    status TEXT DEFAULT 'Draft',
    dikonfirmasi_oleh TEXT,
    tanggal_konfirmasi TIMESTAMP,
    catatan TEXT,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- SHU: Alokasi Komponen
CREATE TABLE IF NOT EXISTS shu_alokasi (
    id_alokasi TEXT PRIMARY KEY,
    id_shu TEXT NOT NULL,
    nama_komponen TEXT NOT NULL,
    persentase NUMERIC(10,4) NOT NULL,
    nominal NUMERIC(18,2),
    urutan INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT NOW(),
    CONSTRAINT fk_shu_alokasi_shu FOREIGN KEY (id_shu) REFERENCES shu_tahunan (id_shu) ON DELETE CASCADE
);

-- SHU: Hak Anggota
CREATE TABLE IF NOT EXISTS shu_anggota (
    id_shu_anggota TEXT PRIMARY KEY,
    id_shu TEXT NOT NULL,
    id_anggota TEXT NOT NULL,
    no_anggota TEXT,
    nama_anggota TEXT,
    jasa_anggota NUMERIC(18,2),
    nilai_jasa NUMERIC(18,2),
    shu_pasif NUMERIC(18,2),
    shu_aktif NUMERIC(18,2),
    shu_total NUMERIC(18,2),
    status_ambil TEXT DEFAULT 'Belum Diambil',
    tanggal_ambil TIMESTAMP,
    catatan TEXT,
    created_at TIMESTAMP DEFAULT NOW(),
    CONSTRAINT fk_shu_anggota_shu FOREIGN KEY (id_shu) REFERENCES shu_tahunan (id_shu) ON DELETE CASCADE,
    CONSTRAINT fk_shu_anggota_anggota FOREIGN KEY (id_anggota) REFERENCES anggota (id_anggota) ON DELETE CASCADE
);

ALTER TABLE IF EXISTS shu_tahunan
    ADD COLUMN IF NOT EXISTS tanggal_input TIMESTAMP DEFAULT NOW(),
    ADD COLUMN IF NOT EXISTS cadangan_umum NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS shu_pasif_total NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS shu_aktif_total NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS dana_kesejahteraan NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS dana_pendidikan NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS dana_sosial NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS dana_pembangunan NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS dana_pengurus NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS dana_risiko NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS catatan TEXT;

ALTER TABLE IF EXISTS shu_anggota
    ADD COLUMN IF NOT EXISTS no_anggota TEXT,
    ADD COLUMN IF NOT EXISTS nama_anggota TEXT,
    ADD COLUMN IF NOT EXISTS jasa_anggota NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS nilai_jasa NUMERIC(18,2),
    ADD COLUMN IF NOT EXISTS catatan TEXT;

CREATE INDEX IF NOT EXISTS idx_shu_tahunan_tahun ON shu_tahunan (tahun);
CREATE INDEX IF NOT EXISTS idx_shu_alokasi_id_shu ON shu_alokasi (id_shu);
CREATE INDEX IF NOT EXISTS idx_shu_anggota_id_shu_anggota ON shu_anggota (id_shu, id_anggota);

-- Notifikasi
CREATE TABLE IF NOT EXISTS notifikasi (
    id_notifikasi TEXT PRIMARY KEY,
    id_user TEXT,
    id_anggota TEXT,
    jenis_notifikasi TEXT,
    subjek TEXT,
    pesan TEXT,
    tipe_medium TEXT,
    status_pengiriman TEXT DEFAULT 'Draft',
    tanggal_dijadwalkan TIMESTAMP,
    tanggal_terkirim TIMESTAMP,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- Upload: Dokumen
CREATE TABLE IF NOT EXISTS dokumen_upload (
    id_dokumen TEXT PRIMARY KEY,
    tipe_dokumen TEXT,
    referensi_id TEXT,
    referensi_tabel TEXT,
    nama_file TEXT,
    path_file TEXT,
    ukuran_file INTEGER,
    tipe_mime TEXT,
    diupload_oleh TEXT,
    created_at TIMESTAMP DEFAULT NOW()
);

-- Audit: Activity Log
CREATE TABLE IF NOT EXISTS audit_logs (
    id_log BIGSERIAL PRIMARY KEY,
    id_user TEXT,
    username TEXT,
    modul TEXT,
    aksi TEXT,
    deskripsi TEXT,
    tabel_terdampak TEXT,
    id_record TEXT,
    nilai_lama TEXT,
    nilai_baru TEXT,
    ip_address TEXT,
    user_agent TEXT,
    created_at TIMESTAMP DEFAULT NOW()
);

-- Settings: Konfigurasi Sistem
CREATE TABLE IF NOT EXISTS settings (
    id_setting TEXT PRIMARY KEY,
    kunci TEXT NOT NULL UNIQUE,
    nilai TEXT,
    tipe_data TEXT DEFAULT 'TEXT',
    deskripsi TEXT,
    readonly BOOLEAN DEFAULT FALSE,
    updated_at TIMESTAMP DEFAULT NOW()
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
