import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)

DATABASE_URL = os.getenv('DATABASE_URL', '').strip()
DB_POOL_SIZE = int(os.getenv('DB_POOL_SIZE', '5'))
DB_MAX_OVERFLOW = int(os.getenv('DB_MAX_OVERFLOW', '10'))
DB_POOL_TIMEOUT = int(os.getenv('DB_POOL_TIMEOUT', '30'))
DB_ECHO = os.getenv('DB_ECHO', '0') == '1'
DB_USE_SSL = os.getenv('DB_USE_SSL', '1') == '1'

FILE_ANGGOTA = os.path.join(DATA_DIR, 'anggota.xlsx')
FILE_SIMPANAN = os.path.join(DATA_DIR, 'simpanan.xlsx')
FILE_SIMPANAN_TRANSAKSI = os.path.join(DATA_DIR, 'simpanan_transaksi.xlsx')
FILE_SIMPANAN_PENGAJUAN = os.path.join(DATA_DIR, 'simpanan_pengajuan.xlsx')
FILE_IURAN_SOSIAL = os.path.join(DATA_DIR, 'iuran_sosial.xlsx')
FILE_PINJAMAN = os.path.join(DATA_DIR, 'pinjaman.xlsx')
FILE_PINJAMAN_CICILAN = os.path.join(DATA_DIR, 'pinjaman_cicilan.xlsx')
FILE_USERS = os.path.join(DATA_DIR, 'users.xlsx')
FILE_PENDAFTARAN_ANGGOTA = os.path.join(DATA_DIR, 'pendaftaran_anggota.xlsx')
FILE_IMPORT_LOG = os.path.join(DATA_DIR, 'import_log.xlsx')
FILE_BERITA = os.path.join(DATA_DIR, 'berita.json')
IMPORT_PREVIEW_DIR = os.path.join(DATA_DIR, 'import_preview')
BACKUP_DIR = os.path.join(DATA_DIR, 'backup')

DSR_DEFAULT = 0.35
JENIS_SIMPANAN_IMPORT = 'Manasuka'
DEFAULT_TENOR_IMPORT_PINJAMAN = 12
JENIS_IMPORT_CSV = 'Import CSV'
PROVISI_RATE_LONG_TENOR = 0.02
PROVISI_MIN_TENOR_BULAN = 13

METODE_BAYAR_CHOICES = (
    'Indomaret',
    'Alfamart',
    'Dana',
    'QRIS',
    'Transfer Bank',
    'Lainnya',
)

# Info rekening tujuan koperasi (untuk transfer). Silakan isi sesuai kebutuhan.
KOPERASI_REKENING_BANK = {
    'nama_bank': 'DANA',
    'no_rekening': '0881023452481',
    'atas_nama': 'Admin Koperasi',
}

ADMIN_NOTIFICATION_EMAIL = 'shidiqper@gmail.com'

CICILAN_FIELDNAMES = [
    'id_cicilan', 'id_pinjaman', 'id_anggota', 'no_anggota',
    'nama_anggota', 'jumlah', 'tanggal_pengajuan', 'status',
    'tanggal_konfirmasi', 'dikonfirmasi_oleh', 'diajukan_oleh',
    'keterangan', 'metode_pembayaran', 'detail_pembayaran',
    'status_transaksi', 'va_number', 'idempotency_key', 'periode_tagihan', 'expires_at',
]

SIMPANAN_FIELDNAMES = ['id_anggota', 'total_simpanan']

SIMPANAN_TRANSAKSI_FIELDNAMES = [
    'id_transaksi', 'id_anggota', 'no_anggota', 'nama_anggota',
    'tanggal', 'jenis_simpanan', 'jumlah', 'keterangan', 'diajukan_oleh',
]

SIMPANAN_PENGAJUAN_FIELDNAMES = [
    'id_pengajuan', 'id_anggota', 'no_anggota', 'nama_anggota',
    'tanggal_pengajuan', 'jenis_simpanan', 'jumlah', 'keterangan',
    'status', 'tanggal_konfirmasi', 'dikonfirmasi_oleh', 'diajukan_oleh',
]

IURAN_SOSIAL_FIELDNAMES = [
    'id_iuran', 'id_anggota', 'no_anggota', 'nama_anggota',
    'tanggal', 'jumlah', 'keterangan', 'diajukan_oleh',
]

PINJAMAN_FIELDNAMES = [
    'id_pinjaman', 'id_anggota', 'nama_anggota', 'no_anggota',
    'jenis_pinjaman', 'jenis_simpanan', 'plafon', 'tenor_awal', 'tenor_bulan', 'bunga_persen',
    'total_bayar', 'cicilan_per_bulan', 'sisa_pinjaman',
    'tanggal_pengajuan', 'status', 'tanggal_lunas',
]

ANGGOTA_FIELDNAMES = [
    'id_anggota', 'no_anggota', 'nik', 'nama', 'alamat', 'no_telp', 'tgl_bergabung',
    'no_rekening', 'nama_bank',
    'penghasilan_bersih', 'cicilan_lain', 'simpanan_pokok',
]

PENDAFTARAN_FIELDNAMES = [
    'id_pengajuan', 'nama', 'alamat', 'no_telp', 'penghasilan_bersih', 'cicilan_lain', 'simpanan_pokok',
    'status', 'tanggal_pengajuan', 'catatan_admin', 'id_anggota_dibuat', 'no_anggota_dibuat',
]

JENIS_PINJAMAN = {
    'Jangka Panjang': {'bunga': 0.8, 'tenor': 24},
    'Jangka Pendek': {'bunga': 1.5, 'tenor': 12},
    'Solusi Cepat': {'bunga': 2.0, 'tenor': 2},
    'Modal Usaha': {'bunga': 0.5, 'tenor': 160},
}

JENIS_PINJAMAN_CHOICES = [
    'Solusi Cepat',
    'Jangka Pendek',
    'Jangka Panjang',
    'Modal Usaha',
]

JENIS_SIMPANAN = ['Manasuka', 'Hari Raya', 'Pendidikan']
