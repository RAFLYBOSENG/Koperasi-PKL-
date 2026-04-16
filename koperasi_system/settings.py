import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)

FILE_ANGGOTA = os.path.join(DATA_DIR, 'anggota.xlsx')
FILE_SIMPANAN = os.path.join(DATA_DIR, 'simpanan.xlsx')
FILE_SIMPANAN_TRANSAKSI = os.path.join(DATA_DIR, 'simpanan_transaksi.xlsx')
FILE_PINJAMAN = os.path.join(DATA_DIR, 'pinjaman.xlsx')
FILE_PINJAMAN_CICILAN = os.path.join(DATA_DIR, 'pinjaman_cicilan.xlsx')
FILE_USERS = os.path.join(DATA_DIR, 'users.xlsx')
FILE_PENDAFTARAN_ANGGOTA = os.path.join(DATA_DIR, 'pendaftaran_anggota.xlsx')
FILE_IMPORT_LOG = os.path.join(DATA_DIR, 'import_log.xlsx')
IMPORT_PREVIEW_DIR = os.path.join(DATA_DIR, 'import_preview')

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
    'Transfer Bank',
    'Lainnya',
)

CICILAN_FIELDNAMES = [
    'id_cicilan', 'id_pinjaman', 'id_anggota', 'no_anggota',
    'nama_anggota', 'jumlah', 'tanggal_pengajuan', 'status',
    'tanggal_konfirmasi', 'dikonfirmasi_oleh', 'diajukan_oleh',
    'keterangan', 'metode_pembayaran', 'detail_pembayaran',
]

SIMPANAN_FIELDNAMES = ['id_anggota', 'total_simpanan']

SIMPANAN_TRANSAKSI_FIELDNAMES = [
    'id_transaksi', 'id_anggota', 'no_anggota', 'nama_anggota',
    'tanggal', 'jenis_simpanan', 'jumlah', 'keterangan', 'diajukan_oleh',
]

PINJAMAN_FIELDNAMES = [
    'id_pinjaman', 'id_anggota', 'nama_anggota', 'no_anggota',
    'jenis_pinjaman', 'jenis_simpanan', 'plafon', 'tenor_awal', 'tenor_bulan', 'bunga_persen',
    'total_bayar', 'cicilan_per_bulan', 'sisa_pinjaman',
    'tanggal_pengajuan', 'status', 'tanggal_lunas',
]

ANGGOTA_FIELDNAMES = [
    'id_anggota', 'no_anggota', 'nik', 'nama', 'alamat', 'no_telp', 'tgl_bergabung',
    'penghasilan_bersih', 'cicilan_lain',
]

PENDAFTARAN_FIELDNAMES = [
    'id_pengajuan', 'nama', 'alamat', 'no_telp', 'penghasilan_bersih', 'cicilan_lain',
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
