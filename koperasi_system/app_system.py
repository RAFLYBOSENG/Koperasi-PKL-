import os
import csv
import calendar
import re
import json
import uuid
from datetime import datetime, timedelta
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file, session, flash, abort
import io
import sys
from secrets import token_hex
try:
    from openpyxl import Workbook, load_workbook
except ModuleNotFoundError:
    Workbook = None
    load_workbook = None
from werkzeug.security import generate_password_hash, check_password_hash
from koperasi_system.settings import (
    BASE_DIR,
    FILE_ANGGOTA,
    FILE_SIMPANAN,
    FILE_SIMPANAN_TRANSAKSI,
    FILE_PINJAMAN,
    FILE_PINJAMAN_CICILAN,
    FILE_USERS,
    FILE_PENDAFTARAN_ANGGOTA,
    FILE_IMPORT_LOG,
    DSR_DEFAULT,
    PROVISI_RATE_LONG_TENOR,
    PROVISI_MIN_TENOR_BULAN,
    JENIS_SIMPANAN_IMPORT,
    DEFAULT_TENOR_IMPORT_PINJAMAN,
    METODE_BAYAR_CHOICES,
    CICILAN_FIELDNAMES,
    SIMPANAN_FIELDNAMES,
    SIMPANAN_TRANSAKSI_FIELDNAMES,
    PINJAMAN_FIELDNAMES,
    JENIS_IMPORT_CSV,
    IMPORT_PREVIEW_DIR,
    ANGGOTA_FIELDNAMES,
    PENDAFTARAN_FIELDNAMES,
    JENIS_PINJAMAN,
    JENIS_PINJAMAN_CHOICES,
    JENIS_SIMPANAN,
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static'),
)
app.secret_key = os.getenv('KOPERASI_SECRET_KEY') or token_hex(32)

# ──────────────────────────────────────────────
#  HELPER: Baca & Tulis Excel/CSV
# ──────────────────────────────────────────────
def baca_csv(filepath):
    """Baca file Excel (.xlsx) atau CSV (.csv) dan kembalikan list of dict."""
    if not os.path.exists(filepath):
        return []
    
    file_ext = os.path.splitext(filepath)[1].lower()
    
    # Jika file Excel
    if file_ext == '.xlsx':
        try:
            wb = load_workbook(filepath, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return []
            
            headers = [str(h or '').strip() for h in rows[0]]
            result = []
            for row in rows[1:]:
                row_dict = {}
                for i, header in enumerate(headers):
                    if header:
                        value = row[i] if i < len(row) else None
                        row_dict[header] = '' if value is None else str(value).strip()
                # Hanya tambah jika ada minimal satu field yang terisi
                if any(row_dict.values()):
                    result.append(row_dict)
            return result
        except Exception as e:
            print(f"Error membaca {filepath}: {e}")
            return []
    
    # Fallback untuk CSV (backward compatibility)
    elif file_ext == '.csv':
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                return list(reader) if reader else []
        except Exception as e:
            print(f"Error membaca CSV {filepath}: {e}")
            return []
    
    return []


def tulis_csv(filepath, data, fieldnames):
    """Tulis list of dict ke file Excel (.xlsx)."""
    try:
        # Create new workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        
        # Tulis header
        ws.append(fieldnames)
        
        # Tulis data dengan tipe yang tepat
        for row_dict in data:
            row_values = []
            for field in fieldnames:
                value = row_dict.get(field, '')
                # Coba convert ke angka jika memungkinkan
                if value and str(value).replace(',', '').replace('.', '').isdigit():
                    try:
                        if ',' in str(value):
                            row_values.append(float(str(value).replace(',', '')))
                        else:
                            row_values.append(float(value) if '.' in str(value) else int(value))
                    except (ValueError, TypeError):
                        row_values.append(value)
                else:
                    row_values.append(value)
            ws.append(row_values)
        
        # Format header (bold)
        from openpyxl.styles import Font
        for cell in ws[1]:
            cell.font = Font(bold=True)
        
        # Auto-adjust column width
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(filepath)
    except Exception as e:
        print(f"Error menulis {filepath}: {e}")


def migrate_csv_to_excel():
    """Konversi file CSV lama ke format Excel (jika ada)."""
    DATA_DIR = os.path.dirname(FILE_ANGGOTA)
    csv_files = {
        'anggota.csv': (FILE_ANGGOTA, ANGGOTA_FIELDNAMES),
        'simpanan.csv': (FILE_SIMPANAN, SIMPANAN_FIELDNAMES),
        'simpanan_transaksi.csv': (FILE_SIMPANAN_TRANSAKSI, SIMPANAN_TRANSAKSI_FIELDNAMES),
        'pinjaman.csv': (FILE_PINJAMAN, PINJAMAN_FIELDNAMES),
        'pinjaman_cicilan.csv': (FILE_PINJAMAN_CICILAN, CICILAN_FIELDNAMES),
        'users.csv': (FILE_USERS, ['id_user', 'username', 'password_hash', 'role', 'id_anggota', 'created_at']),
        'pendaftaran_anggota.csv': (FILE_PENDAFTARAN_ANGGOTA, PENDAFTARAN_FIELDNAMES),
        'import_log.csv': (FILE_IMPORT_LOG, ['waktu', 'user', 'mode', 'nama_file', 'berhasil', 'gagal', 'catatan']),
    }
    
    for csv_filename, (excel_path, fieldnames) in csv_files.items():
        csv_path = os.path.join(DATA_DIR, csv_filename)
        if os.path.exists(csv_path) and not os.path.exists(excel_path):
            try:
                # Baca CSV lama
                with open(csv_path, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    rows = list(reader) if reader else []
                
                # Tulis ke Excel baru
                tulis_csv(excel_path, rows, fieldnames)
                
                # Hapus file CSV lama
                os.remove(csv_path)
                print(f"Migrated {csv_filename} -> {csv_filename.replace('.csv', '.xlsx')}")
            except Exception as e:
                print(f"Error migrating {csv_filename}: {e}")


def init_csv():
    """Inisialisasi file Excel jika belum ada."""
    if not os.path.exists(FILE_ANGGOTA):
        tulis_csv(FILE_ANGGOTA, [], ANGGOTA_FIELDNAMES)
    if not os.path.exists(FILE_SIMPANAN):
        tulis_csv(FILE_SIMPANAN, [], SIMPANAN_FIELDNAMES)
    if not os.path.exists(FILE_SIMPANAN_TRANSAKSI):
        tulis_csv(FILE_SIMPANAN_TRANSAKSI, [], SIMPANAN_TRANSAKSI_FIELDNAMES)
    if not os.path.exists(FILE_PINJAMAN):
        tulis_csv(FILE_PINJAMAN, [], PINJAMAN_FIELDNAMES)
    if not os.path.exists(FILE_PINJAMAN_CICILAN):
        tulis_csv(FILE_PINJAMAN_CICILAN, [], CICILAN_FIELDNAMES)
    if not os.path.exists(FILE_USERS):
        default_users = [
            {
                'id_user': str(uuid.uuid4()),
                'username': 'admin',
                'password_hash': generate_password_hash('Admin@123'),
                'role': 'admin',
                'id_anggota': '',
                'created_at': datetime.now().strftime('%Y-%m-%d')
            },
            {
                'id_user': str(uuid.uuid4()),
                'username': 'user',
                'password_hash': generate_password_hash('User@1234'),
                'role': 'user',
                'id_anggota': '',
                'created_at': datetime.now().strftime('%Y-%m-%d')
            }
        ]
        tulis_csv(FILE_USERS, default_users, ['id_user', 'username', 'password_hash', 'role', 'id_anggota', 'created_at'])
    if not os.path.exists(FILE_PENDAFTARAN_ANGGOTA):
        tulis_csv(FILE_PENDAFTARAN_ANGGOTA, [], PENDAFTARAN_FIELDNAMES)


def kategori_pinjaman_dari_tenor(tenor_bulan: int) -> str:
    """Kelompokkan pinjaman impor menjadi Jangka Pendek atau Jangka Panjang."""
    tenor = max(int(tenor_bulan or 0), 0)
    if tenor >= PROVISI_MIN_TENOR_BULAN:
        return 'Jangka Panjang'
    return 'Jangka Pendek'


# Migrasi CSV lama ke Excel
migrate_csv_to_excel()

# Inisialisasi file Excel
init_csv()


def migrate_simpanan_ke_saldo():
    """Migrasi skema lama (transaksi) ke saldo per id_anggota."""
    rows = baca_csv(FILE_SIMPANAN)
    if not rows:
        return
    if 'total_simpanan' in rows[0] and 'id_anggota' in rows[0]:
        merged = _dedupe_rows_simpanan(rows)
        tulis_csv(FILE_SIMPANAN, merged, SIMPANAN_FIELDNAMES)
        return
    agg = {}
    for s in rows:
        id_a = (s.get('id_anggota') or '').strip()
        if not id_a:
            continue
        if (s.get('status') or 'Disetujui') != 'Disetujui':
            continue
        try:
            agg[id_a] = agg.get(id_a, 0.0) + float(s.get('jumlah') or 0)
        except (TypeError, ValueError):
            continue
    out = [{'id_anggota': k, 'total_simpanan': str(round(v, 2))} for k, v in sorted(agg.items())]
    tulis_csv(FILE_SIMPANAN, out, SIMPANAN_FIELDNAMES)


def migrate_pinjaman_ke_saldo():
    """Migrasi berbagai skema lama ke format multi-baris + jenis pinjaman."""
    rows = baca_csv(FILE_PINJAMAN)
    if not rows:
        return
    sample = rows[0]

    if 'id_pinjaman' in sample:
        normalized = []
        for r in rows:
            row = {k: (r.get(k) if r.get(k) is not None else '') for k in PINJAMAN_FIELDNAMES}
            if not row.get('id_pinjaman'):
                row['id_pinjaman'] = str(uuid.uuid4())
            if not row.get('tenor_awal'):
                row['tenor_awal'] = row.get('tenor_bulan') or ''
            jenis_raw = (row.get('jenis_pinjaman') or '').strip()
            if not jenis_raw or jenis_raw == JENIS_IMPORT_CSV:
                try:
                    tenor_norm = int(float(row.get('tenor_bulan') or 0))
                except (TypeError, ValueError):
                    tenor_norm = 0
                row['jenis_pinjaman'] = kategori_pinjaman_dari_tenor(tenor_norm)
            normalized.append(row)
        tulis_csv(FILE_PINJAMAN, normalized, PINJAMAN_FIELDNAMES)
        return

    # Skema saldo 3 kolom (id_anggota, total_pinjaman, tenor)
    if 'total_pinjaman' in sample and 'tenor' in sample and 'plafon' not in sample:
        amap = {a.get('id_anggota'): a for a in baca_csv(FILE_ANGGOTA)}
        out = []
        for r in rows:
            id_a = (r.get('id_anggota') or '').strip()
            if not id_a:
                continue
            try:
                plaf = float(r.get('total_pinjaman') or 0)
                ten = int(float(r.get('tenor') or 0))
            except (TypeError, ValueError):
                continue
            ag = amap.get(id_a) or {}
            bunga_p = bunga_dari_tenor(ten) if ten > 0 else 0.0
            total_bayar, _ = hitung_total_bayar_flat(plaf, bunga_p, ten, kategori_pinjaman_dari_tenor(ten))
            # Hitung cicilan bulanan: cicilan = (p / b) + (p × j%)
            cic = hitung_cicilan_bulanan(plaf, bunga_p, ten)
            out.append({
                'id_pinjaman': str(uuid.uuid4()),
                'id_anggota': id_a,
                'nama_anggota': ag.get('nama', ''),
                'no_anggota': ag.get('no_anggota', ''),
                'jenis_pinjaman': kategori_pinjaman_dari_tenor(ten),
                'plafon': str(round(plaf, 2)),
                'tenor_awal': str(ten),
                'tenor_bulan': str(ten),
                'bunga_persen': str(bunga_p),
                'total_bayar': str(round(total_bayar, 2)),
                'cicilan_per_bulan': str(round(cic, 2)),
                'sisa_pinjaman': str(round(plaf, 2)),
                'tanggal_pengajuan': '-',
                'status': 'Disetujui',
                'tanggal_lunas': '',
            })
        tulis_csv(FILE_PINJAMAN, out, PINJAMAN_FIELDNAMES)
        return

    # Skema transaksi lama (plafon, tenor_bulan, ...)
    agg_tot = {}
    agg_tenor = {}
    for p in rows:
        id_a = (p.get('id_anggota') or '').strip()
        if not id_a:
            continue
        st = p.get('status', '')
        if st not in ('Disetujui', 'Lunas', 'Menunggu'):
            continue
        try:
            plaf = float(p.get('plafon') or 0)
            t = int(float(p.get('tenor_bulan') or 0))
        except (TypeError, ValueError):
            continue
        agg_tot[id_a] = agg_tot.get(id_a, 0.0) + plaf
        agg_tenor[id_a] = max(agg_tenor.get(id_a, 0), t)
    amap = {a.get('id_anggota'): a for a in baca_csv(FILE_ANGGOTA)}
    out = []
    for id_a in sorted(agg_tot.keys()):
        ag = amap.get(id_a) or {}
        plaf = agg_tot[id_a]
        ten = max(agg_tenor.get(id_a, 0), 0)
        bunga_p = bunga_dari_tenor(ten) if ten > 0 else 0.0
        total_bayar, _ = hitung_total_bayar_flat(plaf, bunga_p, ten, kategori_pinjaman_dari_tenor(ten))
        # Hitung cicilan bulanan: cicilan = (p / b) + (p × j%)
        cic = hitung_cicilan_bulanan(plaf, bunga_p, ten)
        out.append({
            'id_pinjaman': str(uuid.uuid4()),
            'id_anggota': id_a,
            'nama_anggota': ag.get('nama', ''),
            'no_anggota': ag.get('no_anggota', ''),
            'jenis_pinjaman': kategori_pinjaman_dari_tenor(ten),
            'plafon': str(round(plaf, 2)),
            'tenor_awal': str(ten),
            'tenor_bulan': str(ten),
            'bunga_persen': str(bunga_p),
            'total_bayar': str(round(total_bayar, 2)),
            'cicilan_per_bulan': str(round(cic, 2)),
            'sisa_pinjaman': str(round(plaf, 2)),
            'tanggal_pengajuan': '-',
            'status': 'Disetujui',
            'tanggal_lunas': '',
        })
    tulis_csv(FILE_PINJAMAN, out, PINJAMAN_FIELDNAMES)


def _dedupe_rows_simpanan(rows):
    seen = {}
    for r in rows:
        id_a = (r.get('id_anggota') or '').strip()
        if not id_a:
            continue
        try:
            v = float(r.get('total_simpanan') or 0)
        except (TypeError, ValueError):
            v = 0.0
        seen[id_a] = seen.get(id_a, 0.0) + v
    return [{'id_anggota': k, 'total_simpanan': str(round(v, 2))} for k, v in sorted(seen.items())]


def _dedupe_rows_pinjaman(rows):
    """Satukan baris dengan (id_anggota, jenis_pinjaman) sama — untuk impor CSV."""
    if not rows:
        return []
    if 'jenis_pinjaman' not in rows[0]:
        return rows
    merged = {}
    order = []
    for r in rows:
        id_a = (r.get('id_anggota') or '').strip()
        tenor_raw = r.get('tenor_bulan') or r.get('tenor') or 0
        try:
            tenor_norm = int(float(tenor_raw))
        except (TypeError, ValueError):
            tenor_norm = 0
        jn = (r.get('jenis_pinjaman') or '').strip()
        if not jn or jn == JENIS_IMPORT_CSV:
            jn = kategori_pinjaman_dari_tenor(tenor_norm)
        if not id_a:
            continue
        key = (id_a, jn)
        if key not in merged:
            merged[key] = dict(r)
            order.append(key)
            continue
        m = merged[key]
        try:
            plaf = float(m.get('plafon') or 0) + float(r.get('plafon') or 0)
            ten = max(int(float(m.get('tenor_bulan') or 0)), int(float(r.get('tenor_bulan') or 0)))
        except (TypeError, ValueError):
            continue
        bp = float(m.get('bunga_persen') or 0)
        tb, _ = hitung_total_bayar_flat(plaf, bp, ten, jn)
        cic = hitung_cicilan_bulanan(plaf, bp, ten)
        m['plafon'] = str(round(plaf, 2))
        m['tenor_awal'] = str(ten)
        m['tenor_bulan'] = str(ten)
        m['total_bayar'] = str(round(tb, 2))
        m['cicilan_per_bulan'] = str(round(cic, 2))
        if m.get('status') == 'Disetujui':
            m['sisa_pinjaman'] = str(round(plaf, 2))
    return [merged[k] for k in order]


def merge_akumulasi(simpanan_rows: list, id_anggota: str, tambah: float) -> None:
    """Akumulasi simpanan per id_anggota (satu baris per anggota)."""
    tambah = max(float(tambah or 0), 0.0)
    for r in simpanan_rows:
        if r.get('id_anggota') == id_anggota:
            cur = float(r.get('total_simpanan') or 0)
            r['total_simpanan'] = str(round(cur + tambah, 2))
            return
    simpanan_rows.append({'id_anggota': id_anggota, 'total_simpanan': str(round(tambah, 2))})


def merge_pinjaman_akumulasi(
    pinjaman_rows: list,
    id_anggota: str,
    tambah_pinjaman: float,
    tenor_baru: int,
    jenis: str = None,
) -> None:
    """Akumulasi pinjaman per (anggota, jenis). Jangka pendek & panjang tidak digabung."""
    jenis_key = (jenis or kategori_pinjaman_dari_tenor(tenor_baru)).strip()
    tambah_pinjaman = max(float(tambah_pinjaman or 0), 0.0)
    tenor_baru = max(int(tenor_baru or 0), 0)
    for r in pinjaman_rows:
        if r.get('id_anggota') != id_anggota:
            continue
        if (r.get('jenis_pinjaman') or '').strip() != jenis_key:
            continue
        plaf = float(r.get('plafon') or 0) + tambah_pinjaman
        ten = max(int(float(r.get('tenor_bulan') or 0)), tenor_baru)
        bp = float(r.get('bunga_persen') or 0)
        tb, _ = hitung_total_bayar_flat(plaf, bp, ten, jenis_key)
        cic = hitung_cicilan_bulanan(plaf, bp, ten)
        r['plafon'] = str(round(plaf, 2))
        r['tenor_bulan'] = str(ten)
        r['total_bayar'] = str(round(tb, 2))
        r['cicilan_per_bulan'] = str(round(cic, 2))
        if r.get('status') == 'Disetujui':
            r['sisa_pinjaman'] = str(round(plaf, 2))
        return
    bp = bunga_dari_tenor(tenor_baru) if tenor_baru > 0 else 0.0
    plaf = tambah_pinjaman
    tb, _ = hitung_total_bayar_flat(plaf, bp, tenor_baru, jenis_key)
    cic = hitung_cicilan_bulanan(plaf, bp, tenor_baru)
    pinjaman_rows.append({
        'id_pinjaman': str(uuid.uuid4()),
        'id_anggota': id_anggota,
        'nama_anggota': '',
        'no_anggota': '',
        'jenis_pinjaman': jenis_key,
        'plafon': str(round(plaf, 2)),
        'tenor_awal': str(tenor_baru),
        'tenor_bulan': str(tenor_baru),
        'bunga_persen': str(bp),
        'total_bayar': str(round(tb, 2)),
        'cicilan_per_bulan': str(round(cic, 2)),
        'sisa_pinjaman': str(round(plaf, 2)),
        'tanggal_pengajuan': datetime.now().strftime('%Y-%m-%d'),
        'status': 'Disetujui',
        'tanggal_lunas': '',
    })


def merge_pinjaman_sama_anggota(pinjaman_rows: list, id_pinjaman_baru: str, id_anggota: str) -> None:
    """Merge pinjaman yang baru disetujui dengan pinjaman existing dari anggota sama.
    Gabungkan plafon dengan pinjaman baru, lalu pakai tenor dan bunga dari pinjaman terbaru."""
    pinjaman_baru = None
    pinjaman_existing = None
    idx_existing = -1

    # Cari pinjaman baru yang baru dikonfirmasi
    for p in pinjaman_rows:
        if p.get('id_pinjaman') == id_pinjaman_baru:
            pinjaman_baru = p
            break

    if not pinjaman_baru or pinjaman_baru.get('status') != 'Disetujui':
        return

    # Cari pinjaman approved existing dari anggota sama dengan jenis sama
    jenis_baru = (pinjaman_baru.get('jenis_pinjaman') or '').strip()
    for i, p in enumerate(pinjaman_rows):
        jenis_existing = (p.get('jenis_pinjaman') or '').strip()
        if (p.get('id_anggota') == id_anggota and 
            p.get('status') == 'Disetujui' and 
            p.get('id_pinjaman') != id_pinjaman_baru and
            jenis_existing == jenis_baru):
            pinjaman_existing = p
            idx_existing = i
            break

    # Jika ada pinjaman existing, merge keduanya
    if pinjaman_existing and idx_existing >= 0:
        try:
            plaf_baru = float(pinjaman_baru.get('plafon') or 0)
            plaf_existing = float(pinjaman_existing.get('plafon') or 0)
            cic_baru = float(pinjaman_baru.get('cicilan_per_bulan') or 0)
            cic_existing = float(pinjaman_existing.get('cicilan_per_bulan') or 0)
            ten_baru = int(float(pinjaman_baru.get('tenor_bulan') or 0))
            bp = float(pinjaman_baru.get('bunga_persen') or 0)

            # Gabung plafon
            plaf_merged = plaf_baru + plaf_existing

            # Gunakan tenor terbaru dari pinjaman baru.
            ten_merged = ten_baru

            # Total kewajiban mengikuti tenor terbaru, bukan tenor gabungan.
            tb_merged = hitung_total_bayar_tanpa_provisi(plaf_merged, bp, ten_merged, jenis_baru)

            # Hitung cicilan merged
            cic_merged = hitung_cicilan_bulanan(plaf_merged, bp, ten_merged)

            # Update pinjaman existing dengan data merged
            pinjaman_existing['plafon'] = str(round(plaf_merged, 2))
            pinjaman_existing['tenor_bulan'] = str(ten_merged)
            pinjaman_existing['total_bayar'] = str(round(tb_merged, 2))
            pinjaman_existing['cicilan_per_bulan'] = str(round(cic_merged, 2))
            sisa_existing = saldo_pinjaman_aktual(pinjaman_existing)
            pinjaman_existing['sisa_pinjaman'] = str(round(sisa_existing + plaf_baru, 2))

            # Hapus pinjaman baru (karena sudah digabung ke existing)
            pinjaman_rows.pop(pinjaman_rows.index(pinjaman_baru))
        except (TypeError, ValueError):
            pass


def get_pinjaman_saldo(pinjaman_rows: list, id_anggota: str):
    for p in pinjaman_rows:
        if p.get('id_anggota') == id_anggota:
            return p
    return None


def cicilan_per_bulan_saldo(p: dict) -> float:
    if p.get('cicilan_per_bulan') not in (None, ''):
        try:
            return float(p['cicilan_per_bulan'])
        except (TypeError, ValueError):
            pass
    tot = saldo_pinjaman_aktual(p)
    ten = int(float(p.get('tenor_bulan') or p.get('tenor') or 0))
    if ten <= 0 or tot <= 0:
        return 0.0
    return tot / ten


def tenor_sisa_pinjaman_aktual(p: dict) -> int:
    """Jumlah tenor tersisa yang tersimpan di data pinjaman."""
    try:
        tenor = int(float(p.get('tenor_bulan') or p.get('tenor_awal') or p.get('tenor') or 0))
    except (TypeError, ValueError):
        tenor = 0
    return max(tenor, 0)


def enrich_simpanan_untuk_tampilan(simpanan_rows: list, anggota_rows: list) -> list:
    amap = {a.get('id_anggota'): a for a in anggota_rows}
    out = []
    for s in simpanan_rows:
        a = amap.get(s.get('id_anggota'), {})
        out.append({
            **s,
            'no_anggota': a.get('no_anggota', ''),
            'nama_anggota': a.get('nama', ''),
            'jumlah': s.get('total_simpanan', '0'),
            'tanggal': '-',
            'status': 'Disetujui',
            'jenis_simpanan': 'Saldo',
            'keterangan': 'Total saldo per anggota',
            'id_simpanan': s.get('id_anggota', ''),
        })
    return out


def enrich_pinjaman_untuk_tampilan(pinjaman_rows: list, anggota_rows: list) -> list:
    amap = {a.get('id_anggota'): a for a in anggota_rows}
    riwayat_jenis_map = {}

    for p in pinjaman_rows:
        id_a = p.get('id_anggota', '')
        tenor_raw_hist = p.get('tenor_bulan', p.get('tenor', '0'))
        try:
            tenor_norm_hist = int(float(tenor_raw_hist or 0))
        except (TypeError, ValueError):
            tenor_norm_hist = 0
        jenis_hist = (p.get('jenis_pinjaman') or '').strip()
        if not jenis_hist or jenis_hist == JENIS_IMPORT_CSV:
            jenis_hist = kategori_pinjaman_dari_tenor(tenor_norm_hist)
        if not id_a or not jenis_hist:
            continue
        riwayat_jenis_map.setdefault(id_a, [])
        if jenis_hist not in riwayat_jenis_map[id_a]:
            riwayat_jenis_map[id_a].append(jenis_hist)

    out = []
    for p in pinjaman_rows:
        a = amap.get(p.get('id_anggota'), {})
        tenor_raw = p.get('tenor_bulan', p.get('tenor', '0'))
        try:
            tenor_norm = int(float(tenor_raw or 0))
        except (TypeError, ValueError):
            tenor_norm = 0
        jenis_raw = (p.get('jenis_pinjaman') or '').strip()
        if not jenis_raw or jenis_raw == JENIS_IMPORT_CSV:
            jenis_raw = kategori_pinjaman_dari_tenor(tenor_norm)
        if p.get('plafon') is not None or p.get('id_pinjaman'):
            plaf = float(p.get('plafon') or 0)
            sisa = saldo_pinjaman_aktual(p)
            cic = float(p.get('cicilan_per_bulan') or 0) or cicilan_per_bulan_saldo(p)
            st = p.get('status') or 'Menunggu'
            jatuh_tempo = tanggal_jatuh_tempo_pinjaman(p)
            tenor_tampil = tenor_sisa_pinjaman_aktual(p)
            tenor_provisi = int(float(p.get('tenor_awal') or p.get('tenor_bulan') or 0))
            provisi_nominal = provisi_nominal_pinjaman(jenis_raw, plaf, tenor_provisi)
            out.append({
                **p,
                'no_anggota': p.get('no_anggota') or a.get('no_anggota', ''),
                'nama_anggota': p.get('nama_anggota') or a.get('nama', ''),
                'plafon': str(plaf),
                'tenor_bulan': str(max(int(tenor_tampil), 0)),
                'cicilan_per_bulan': str(round(cic, 2)),
                'sisa_pinjaman': str(round(sisa, 2)),
                'status': st,
                'jatuh_tempo': jatuh_tempo.strftime('%Y-%m-%d') if jatuh_tempo else '',
                'telat_bayar': pinjaman_telat_bayar_aktual(p),
                'tanggal_pengajuan': p.get('tanggal_pengajuan') or '-',
                'id_pinjaman': p.get('id_pinjaman', ''),
                'id_anggota': p.get('id_anggota', ''),
                'jenis_pinjaman': jenis_raw,
                'riwayat_jenis_pinjaman': ', '.join(riwayat_jenis_map.get(p.get('id_anggota', ''), [jenis_raw])),
                'jenis_simpanan': (p.get('jenis_simpanan') or 'Manasuka').strip() or 'Manasuka',
                'provisi_nominal': str(round(provisi_nominal, 2)),
            })
            continue
        tot = float(p.get('total_pinjaman') or 0)
        ten = tenor_norm
        cic = cicilan_per_bulan_saldo(p)
        st = 'Disetujui' if tot > 0 else 'Lunas'
        out.append({
            **p,
            'no_anggota': a.get('no_anggota', ''),
            'nama_anggota': a.get('nama', ''),
            'plafon': p.get('total_pinjaman', '0'),
            'tenor_bulan': p.get('tenor', '0'),
            'cicilan_per_bulan': str(round(cic, 2)),
            'sisa_pinjaman': str(round(tot, 2)),
            'status': st,
            'jatuh_tempo': '',
            'telat_bayar': False,
            'tanggal_pengajuan': '-',
            'jenis_pinjaman': jenis_raw or 'Saldo',
            'riwayat_jenis_pinjaman': ', '.join(riwayat_jenis_map.get(p.get('id_anggota', ''), [jenis_raw or 'Saldo'])),
            'jenis_simpanan': (p.get('jenis_simpanan') or 'Manasuka').strip() or 'Manasuka',
            'provisi_nominal': '0',
            'id_pinjaman': p.get('id_anggota', ''),
        })
    return out


def ensure_anggota_schema():
    """Pastikan data anggota punya kolom nik, penghasilan, dan cicilan lain."""
    rows = baca_csv(FILE_ANGGOTA)
    if not rows:
        return
    if 'nik' in rows[0] and 'penghasilan_bersih' in rows[0] and 'cicilan_lain' in rows[0]:
        return
    for r in rows:
        r['nik'] = r.get('nik', '')
        r['penghasilan_bersih'] = r.get('penghasilan_bersih', '0')
        r['cicilan_lain'] = r.get('cicilan_lain', '0')
    tulis_csv(FILE_ANGGOTA, rows, ANGGOTA_FIELDNAMES)


def ensure_pinjaman_plafon_schema():
    """Kompatibilitas: migrasi skema lama ke saldo per id_anggota."""
    migrate_pinjaman_ke_saldo()


def ensure_pinjaman_cicilan_schema():
    """Pastikan file pinjaman_cicilan.csv memiliki fieldnames terbaru."""
    if not os.path.exists(FILE_PINJAMAN_CICILAN):
        tulis_csv(FILE_PINJAMAN_CICILAN, [], CICILAN_FIELDNAMES)
        return
    rows = baca_csv(FILE_PINJAMAN_CICILAN)
    normalized = []
    for r in rows:
        normalized.append({k: (r.get(k) or '') for k in CICILAN_FIELDNAMES})
    tulis_csv(FILE_PINJAMAN_CICILAN, normalized, CICILAN_FIELDNAMES)


def ensure_import_log_schema():
    if not os.path.exists(FILE_IMPORT_LOG):
        tulis_csv(
            FILE_IMPORT_LOG,
            [],
            ['waktu', 'user', 'mode', 'nama_file', 'berhasil', 'gagal', 'catatan'],
        )


def ensure_import_preview_dir():
    os.makedirs(IMPORT_PREVIEW_DIR, exist_ok=True)


def migrate_sisa_pinjaman_aktif_ke_total():
    """Migrasi data lama: pinjaman aktif memakai sisa pokok (plafon)."""
    rows = baca_csv(FILE_PINJAMAN)
    if not rows:
        return
    changed = False
    for r in rows:
        status = (r.get('status') or '').strip()
        if status not in ('Menunggu', 'Disetujui'):
            continue
        try:
            plaf = float(r.get('plafon') or r.get('total_pinjaman') or 0)
        except (TypeError, ValueError):
            plaf = 0.0
        if plaf > 0:
            try:
                sisa_now = float(r.get('sisa_pinjaman') or 0)
            except (TypeError, ValueError):
                sisa_now = 0.0
            if sisa_now > plaf:
                r['sisa_pinjaman'] = str(round(plaf, 2))
                changed = True
                continue
        try:
            sisa = float(r.get('sisa_pinjaman') or 0)
        except (TypeError, ValueError):
            sisa = 0.0
        if sisa > 0:
            continue
        if plaf > 0:
            r['sisa_pinjaman'] = str(round(plaf, 2))
            changed = True
    if changed:
        tulis_csv(FILE_PINJAMAN, rows, PINJAMAN_FIELDNAMES)


def ensure_pendaftaran_schema():
    """Migrasi kolom pengajuan anggota (penghasilan, cicilan, id anggota dibuat, dll.)."""
    if not os.path.exists(FILE_PENDAFTARAN_ANGGOTA):
        return
    rows = baca_csv(FILE_PENDAFTARAN_ANGGOTA)
    if rows:
        sample = rows[0]
        if all(k in sample for k in PENDAFTARAN_FIELDNAMES):
            return
        for r in rows:
            for k in PENDAFTARAN_FIELDNAMES:
                if k not in r:
                    r[k] = '0' if k in ('penghasilan_bersih', 'cicilan_lain') else ''
        tulis_csv(FILE_PENDAFTARAN_ANGGOTA, rows, PENDAFTARAN_FIELDNAMES)
        return
    # File kosong: tulis ulang dengan header schema terbaru.
    tulis_csv(FILE_PENDAFTARAN_ANGGOTA, [], PENDAFTARAN_FIELDNAMES)


ensure_pendaftaran_schema()

PASSWORD_POLICY_MSG = (
    'Password minimal 8 karakter dan wajib memuat huruf besar (A-Z), huruf kecil (a-z), '
    'angka (0-9), dan simbol khusus (contoh: !@#$).'
)


def validate_password_policy(password: str):
    """Minimal 8 karakter, kombinasi huruf besar/kecil, angka, dan simbol khusus."""
    if not password or len(password) < 8:
        return False, PASSWORD_POLICY_MSG
    if not re.search(r'[A-Z]', password):
        return False, PASSWORD_POLICY_MSG
    if not re.search(r'[a-z]', password):
        return False, PASSWORD_POLICY_MSG
    if not re.search(r'\d', password):
        return False, PASSWORD_POLICY_MSG
    simbol = set('!@#$%^&*()_+-=[]{}|;:,.<>?/`~')
    if not any(c in password for c in simbol):
        return False, PASSWORD_POLICY_MSG
    return True, ''


def get_anggota_by_id(id_anggota: str):
    for a in baca_csv(FILE_ANGGOTA):
        if a.get('id_anggota') == id_anggota:
            return a
    return None


def generate_no_anggota_berikutnya(data_anggota):
    """Generate nomor AGT-XXXX berbasis nomor tertinggi yang sudah ada."""
    max_no = 0
    for a in data_anggota:
        no = (a.get('no_anggota') or '').strip()
        if no.startswith('AGT-'):
            bagian = no.replace('AGT-', '')
            if bagian.isdigit():
                max_no = max(max_no, int(bagian))
    return f"AGT-{(max_no + 1):04d}"


def ensure_simpanan_schema():
    """Kompatibilitas: migrasi skema lama ke saldo per id_anggota."""
    migrate_simpanan_ke_saldo()


def ensure_simpanan_transaksi_schema():
    """Pastikan file transaksi simpanan tersedia dengan skema terbaru."""
    if not os.path.exists(FILE_SIMPANAN_TRANSAKSI):
        tulis_csv(FILE_SIMPANAN_TRANSAKSI, [], SIMPANAN_TRANSAKSI_FIELDNAMES)
        return
    rows = baca_csv(FILE_SIMPANAN_TRANSAKSI)
    normalized = []
    for r in rows:
        normalized.append({k: (r.get(k) or '') for k in SIMPANAN_TRANSAKSI_FIELDNAMES})
    tulis_csv(FILE_SIMPANAN_TRANSAKSI, normalized, SIMPANAN_TRANSAKSI_FIELDNAMES)


def bunga_dari_tenor(tenor_bulan: int) -> float:
    """Fallback bunga per bulan jika jenis pinjaman tidak ditentukan."""
    if tenor_bulan <= 0:
        return 0.0
    if 1 <= tenor_bulan <= 2:
        return 2.0
    if 3 <= tenor_bulan <= 12:
        return 1.5
    if 13 <= tenor_bulan <= 24:
        return 0.80
    if 36 <= tenor_bulan <= 120:
        return 0.75
    raise ValueError('Tenor valid: 1-2, 3-12, 13-24, atau 36-120 bulan.')


def provisi_persen_dari_tenor(tenor_bulan: int) -> float:
    """Provisi one-time 2% dari pokok jika tenor > 12 bulan."""
    return PROVISI_RATE_LONG_TENOR * 100 if tenor_bulan >= PROVISI_MIN_TENOR_BULAN else 0.0


def provisi_persen_dari_pinjaman(jenis_pinjaman: str, tenor_bulan: int) -> float:
    """Provisi per jenis pinjaman.

    - Solusi Cepat: selalu 2%
    - Modal Usaha: selalu 2%
    - Jangka Panjang: 2% jika tenor >= 13 bulan
    - Jenis lain: 0%
    """
    jenis = (jenis_pinjaman or '').strip().lower()
    if jenis == 'solusi cepat':
        return PROVISI_RATE_LONG_TENOR * 100
    if jenis == 'modal usaha':
        return PROVISI_RATE_LONG_TENOR * 100
    if jenis == 'jangka panjang':
        return provisi_persen_dari_tenor(tenor_bulan)
    return 0.0


def provisi_nominal_pinjaman(jenis_pinjaman: str, plafon: float, tenor_bulan: int) -> float:
    """Hitung provisi sesuai jenis pinjaman dan tenor."""
    plafon = max(float(plafon or 0), 0.0)
    return plafon * (provisi_persen_dari_pinjaman(jenis_pinjaman, tenor_bulan) / 100.0)


def hitung_total_bayar_flat(plafon: float, bunga_persen_bulanan: float, tenor_bulan: int, jenis_pinjaman: str = '') -> tuple:
    """Hitung total bayar skema flat bulanan + provisi tenor panjang."""
    if tenor_bulan <= 0:
        return max(float(plafon or 0), 0.0), 0.0
    plafon = max(float(plafon or 0), 0.0)
    # Bunga/jasa hanya dihitung pada bulan 1 s.d. n-1; bulan terakhir bebas bunga.
    bulan_berbunga = max(tenor_bulan - 1, 0)
    bunga_nominal = plafon * (bunga_persen_bulanan / 100.0) * bulan_berbunga
    provisi_nominal = plafon * (provisi_persen_dari_pinjaman(jenis_pinjaman, tenor_bulan) / 100.0)
    return plafon + bunga_nominal + provisi_nominal, provisi_nominal


def hitung_total_bayar_tanpa_provisi(plafon: float, bunga_persen_bulanan: float, tenor_bulan: int, jenis_pinjaman: str = '') -> float:
    """Hitung total bayar skema flat bulanan TANPA provisi (untuk pengajuan awal)."""
    total_bayar, _ = hitung_total_bayar_flat(plafon, bunga_persen_bulanan, tenor_bulan, jenis_pinjaman)
    return total_bayar


def apply_provisi_setelah_konfirmasi(plafon: float, tenor_bulan: int, jenis_pinjaman: str = '') -> float:
    """Hitung provisi yang diterapkan saat admin mengkonfirmasi pinjaman."""
    provisi_nominal = plafon * (provisi_persen_dari_pinjaman(jenis_pinjaman, tenor_bulan) / 100.0)
    return provisi_nominal


def hitung_cicilan_bulanan(plafon: float, bunga_persen_bulanan: float, tenor_bulan: int) -> float:
    """
    Hitung cicilan bulanan per formula: cicilan = (p / b) + (p × j%)
    
    Keterangan:
    - p = plafon (pokok pinjaman)
    - b = tenor (bulan)
    - j = bunga_persen_bulanan (% per bulan)
    
    Contoh: p=15.000.000, b=10, j=1.5%
    cicilan = (15.000.000 / 10) + (15.000.000 × 1.5%) = 1.500.000 + 225.000 = 1.725.000
    """
    if tenor_bulan <= 0:
        return 0.0
    plafon = max(float(plafon or 0), 0.0)
    bunga_persen_bulanan = max(float(bunga_persen_bulanan or 0), 0.0)
    
    # cicilan = (p / b) + (p × j%)
    cicilan_pokok = plafon / tenor_bulan
    total_bunga = plafon * (bunga_persen_bulanan / 100.0)
    cicilan = cicilan_pokok + total_bunga
    
    return cicilan


def nominal_cicilan_aktual(pinjaman: dict) -> float:
    """Nominal yang benar-benar dibayarkan pada transaksi cicilan berjalan."""
    sisa = saldo_pinjaman_aktual(pinjaman)
    jenis = (pinjaman.get('jenis_pinjaman') or '').strip()
    # Solusi Cepat dibayar sekali melalui admin (bulan ke-1 atau ke-2).
    if jenis == 'Solusi Cepat':
        return max(sisa, 0.0)
    cic = float(pinjaman.get('cicilan_per_bulan') or 0) or cicilan_per_bulan_saldo(pinjaman)
    if sisa <= 0 or cic <= 0:
        return 0.0
    return min(sisa, cic)


def saldo_pinjaman_aktual(pinjaman: dict) -> float:
    """Saldo pinjaman yang dipakai UI dan proses bayar.

    Jika data lama masih menyimpan sisa 0, fallback ke plafon (pokok pinjaman).
    """
    try:
        sisa = max(float(pinjaman.get('sisa_pinjaman') or 0), 0.0)
    except (TypeError, ValueError):
        sisa = 0.0
    if sisa > 0:
        return sisa
    status = (pinjaman.get('status') or '').strip()
    if status in ('Lunas', 'Ditolak'):
        return 0.0
    try:
        total = max(float(pinjaman.get('plafon') or pinjaman.get('total_pinjaman') or 0), 0.0)
    except (TypeError, ValueError):
        total = 0.0
    return total


def _parse_tanggal_iso(value: str):
    value = (value or '').strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, '%Y-%m-%d').date()
    except ValueError:
        return None


def _tambah_bulan_ke_tanggal(tanggal_awal, jumlah_bulan: int):
    if not tanggal_awal:
        return None
    jumlah_bulan = max(int(jumlah_bulan or 0), 0)
    if jumlah_bulan <= 0:
        return tanggal_awal
    total_bulan = tanggal_awal.month - 1 + jumlah_bulan
    year = tanggal_awal.year + (total_bulan // 12)
    month = (total_bulan % 12) + 1
    day = min(tanggal_awal.day, calendar.monthrange(year, month)[1])
    return tanggal_awal.replace(year=year, month=month, day=day)


def tanggal_jatuh_tempo_pinjaman(pinjaman: dict):
    tanggal_pengajuan = _parse_tanggal_iso(pinjaman.get('tanggal_pengajuan'))
    tenor = int(float(pinjaman.get('tenor_awal') or pinjaman.get('tenor_bulan') or pinjaman.get('tenor') or 0))
    return _tambah_bulan_ke_tanggal(tanggal_pengajuan, tenor)


def pinjaman_telat_bayar_aktual(pinjaman: dict) -> bool:
    if (pinjaman.get('status') or '').strip() != 'Disetujui':
        return False
    if saldo_pinjaman_aktual(pinjaman) <= 0:
        return False
    jatuh_tempo = tanggal_jatuh_tempo_pinjaman(pinjaman)
    if not jatuh_tempo:
        return False
    return datetime.now().date() > jatuh_tempo


def ada_pinjaman_solusi_cepat_aktif(pinjaman_rows: list, id_anggota: str) -> bool:
    """Cek apakah anggota masih punya pinjaman Solusi Cepat yang belum lunas."""
    for p in pinjaman_rows:
        if p.get('id_anggota') != id_anggota:
            continue
        if (p.get('jenis_pinjaman') or '').strip() != 'Solusi Cepat':
            continue
        status = (p.get('status') or '').strip()
        if status in ('Menunggu', 'Disetujui'):
            return True
        if saldo_pinjaman_aktual(p) > 0 and status != 'Ditolak':
            return True
    return False


migrate_simpanan_ke_saldo()
migrate_pinjaman_ke_saldo()
migrate_sisa_pinjaman_aktif_ke_total()
ensure_pinjaman_cicilan_schema()
ensure_anggota_schema()
ensure_import_log_schema()
ensure_import_preview_dir()
ensure_simpanan_transaksi_schema()


def hitung_kapasitas_cicilan(penghasilan_bersih: float, cicilan_lain: float, dsr: float) -> float:
    return max((penghasilan_bersih * dsr) - cicilan_lain, 0.0)


def dsr_otomatis_dari_penghasilan(penghasilan_bersih: float) -> float:
    """DSR 30–40% mengikuti penghasilan bersih (otomatis, tanpa pilihan manual)."""
    if penghasilan_bersih <= 0:
        return DSR_DEFAULT
    if penghasilan_bersih <= 5_000_000:
        return 0.40
    if penghasilan_bersih <= 15_000_000:
        return 0.35
    return 0.30


def hitung_plafon_maks_anuitas(cicilan_bulanan: float, bunga_persen_bulanan: float, tenor_bulan: int, jenis_pinjaman: str = '') -> float:
    n = int(tenor_bulan or 0)
    if n <= 0:
        raise ValueError('Tenor harus lebih dari 0.')
    cicilan_bulanan = max(float(cicilan_bulanan or 0), 0.0)
    i = max(float(bunga_persen_bulanan or 0), 0.0) / 100.0
    provisi_rate = provisi_persen_dari_pinjaman(jenis_pinjaman, n) / 100.0
    faktor = (1 + (i * n) + provisi_rate)
    if faktor <= 0:
        return 0.0
    return (cicilan_bulanan * n) / faktor


def info_pinjaman_dsr_anggota(anggota_row: dict) -> dict:
    """Ringkasan DSR otomatis dan kapasitas cicilan untuk satu baris anggota (tampilan admin)."""
    ph = max(float(anggota_row.get('penghasilan_bersih') or 0), 0.0)
    cl = max(float(anggota_row.get('cicilan_lain') or 0), 0.0)
    dsr = dsr_otomatis_dari_penghasilan(ph)
    kap = hitung_kapasitas_cicilan(ph, cl, dsr)
    return {
        'dsr_persen': int(round(dsr * 100)),
        'kapasitas_cicilan': kap,
    }


def parse_rupiah_to_float(value):
    return float(str(value or '0').replace(',', '').replace('.', ''))


def normalize_nik(value: str) -> str:
    """Normalisasi NIK: hanya digit, tanpa spasi/simbol."""
    return re.sub(r'\D', '', str(value or '').strip())


def is_valid_nik(value: str) -> bool:
    """Valid jika NIK tepat 16 digit angka."""
    return bool(re.fullmatch(r'\d{16}', str(value or '')))


def get_user_by_username(username: str):
    users = baca_csv(FILE_USERS)
    for u in users:
        if u.get('username', '').lower() == (username or '').lower():
            return u
    return None


def get_user_by_id(id_user: str):
    users = baca_csv(FILE_USERS)
    for u in users:
        if u.get('id_user') == id_user:
            return u
    return None


def get_current_user_id_anggota():
    return session.get('id_anggota') or ''


def is_current_user_admin():
    return session.get('role') == 'admin'


def ensure_users_schema():
    """Pastikan users.csv punya kolom id_anggota (upgrade schema lama)."""
    users = baca_csv(FILE_USERS)
    if not users:
        return
    if 'id_anggota' in users[0]:
        return
    for u in users:
        u['id_anggota'] = ''
    tulis_csv(FILE_USERS, users, ['id_user', 'username', 'password_hash', 'role', 'id_anggota', 'created_at'])


def login_required(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if not session.get('user'):
            return redirect(url_for('login', next=request.path))
        return view_func(*args, **kwargs)
    return wrapper


def admin_required(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if not session.get('user'):
            return redirect(url_for('login', next=request.path))
        if session.get('role') != 'admin':
            abort(403)
        return view_func(*args, **kwargs)
    return wrapper


def restrict_id_anggota_or_forbid(id_anggota: str):
    """User biasa hanya boleh akses data id_anggota dirinya. Admin boleh semua."""
    if is_current_user_admin():
        return
    current = get_current_user_id_anggota()
    if not current or current != id_anggota:
        abort(403)


@app.after_request
def add_no_cache_headers(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0, private'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    response.headers['Referrer-Policy'] = 'no-referrer'
    response.headers['X-Robots-Tag'] = 'noindex, nofollow, noarchive, nosnippet'
    return response


@app.context_processor
def inject_globals():
    return {
        'now': datetime.now().strftime('%d %B %Y'),
        'current_user': session.get('user'),
        'current_role': session.get('role'),
        'is_admin': is_current_user_admin(),
        'current_id_anggota': get_current_user_id_anggota()
    }


# ══════════════════════════════════════════════
#  ROUTE: LANDING & AUTH
# ══════════════════════════════════════════════
@app.route('/')
def landing():
    if session.get('user'):
        return redirect(url_for('dashboard'))
    ensure_simpanan_schema()
    ensure_pinjaman_plafon_schema()
    anggota = baca_csv(FILE_ANGGOTA)
    simpanan = baca_csv(FILE_SIMPANAN)
    pinjaman = baca_csv(FILE_PINJAMAN)
    n_anggota = len(anggota)
    total_simpanan_disetujui = sum(float(s.get('total_simpanan') or 0) for s in simpanan)
    total_pinjaman_disalurkan = sum(
        float(p.get('plafon') or p.get('total_pinjaman') or 0)
        for p in pinjaman
        if p.get('status') in ('Disetujui', 'Lunas')
    )
    # Estimasi ilustratif (bukan pembukuan resmi SHU): proporsi dari total simpanan disetujui
    shu_estimasi_ilustratif = int(total_simpanan_disetujui * 0.08)
    return render_template(
        'landing.html',
        n_anggota=n_anggota,
        total_pinjaman_disalurkan=total_pinjaman_disalurkan,
        shu_estimasi_ilustratif=shu_estimasi_ilustratif,
        total_simpanan_disetujui=total_simpanan_disetujui,
    )


@app.route('/pendaftaran-anggota', methods=['POST'])
def pengajuan_anggota_baru():
    """Pengajuan anggota baru dari landing page."""
    ensure_pendaftaran_schema()
    nama = (request.form.get('nama') or '').strip()
    alamat = (request.form.get('alamat') or '').strip()
    no_telp = (request.form.get('no_telp') or '').strip()
    penghasilan = parse_rupiah_to_float(request.form.get('penghasilan_bersih', '0'))
    cicilan_lain = parse_rupiah_to_float(request.form.get('cicilan_lain', '0'))

    if not nama or not alamat:
        flash('Nama dan alamat wajib diisi untuk pengajuan anggota.', 'danger')
        return redirect(url_for('landing'))
    if penghasilan <= 0:
        flash('Penghasilan bersih per bulan wajib diisi (lebih dari Rp 0) untuk pengajuan anggota.', 'danger')
        return redirect(url_for('landing'))

    data = baca_csv(FILE_PENDAFTARAN_ANGGOTA)
    data.append({
        'id_pengajuan': str(uuid.uuid4()),
        'nama': nama,
        'alamat': alamat,
        'no_telp': no_telp,
        'penghasilan_bersih': str(int(penghasilan)),
        'cicilan_lain': str(int(cicilan_lain)),
        'status': 'Menunggu',
        'tanggal_pengajuan': datetime.now().strftime('%Y-%m-%d'),
        'catatan_admin': '',
        'id_anggota_dibuat': '',
        'no_anggota_dibuat': ''
    })
    tulis_csv(FILE_PENDAFTARAN_ANGGOTA, data, PENDAFTARAN_FIELDNAMES)
    flash('Pengajuan anggota berhasil dikirim. Mohon tunggu konfirmasi admin.', 'success')
    return redirect(url_for('landing'))


@app.route('/robots.txt')
def robots_txt():
    return 'User-agent: *\nDisallow: /\n', 200, {'Content-Type': 'text/plain; charset=utf-8'}


@app.route('/login', methods=['GET', 'POST'])
def login():
    ensure_users_schema()
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        next_url = request.form.get('next') or url_for('dashboard')

        user = get_user_by_username(username)
        if not user or not check_password_hash(user.get('password_hash', ''), password):
            flash('Username atau password salah.', 'danger')
            return render_template('login.html', next=next_url), 401

        session['user'] = user.get('username')
        session['role'] = user.get('role')
        session['id_user'] = user.get('id_user')
        session['id_anggota'] = user.get('id_anggota', '')
        if session.get('role') == 'user' and not session.get('id_anggota'):
            session.clear()
            flash('Akun ini belum dihubungkan ke anggota. Hubungi admin.', 'danger')
            return render_template('login.html', next=next_url), 403
        flash(f"Selamat datang, {session['user']}!", 'success')
        return redirect(next_url)

    if session.get('user'):
        return redirect(url_for('dashboard'))
    return render_template('login.html', next=request.args.get('next', url_for('dashboard')))


@app.route('/lupa-password', methods=['GET', 'POST'])
def lupa_password():
    """Reset password mandiri untuk akun user (verifikasi nomor anggota yang terhubung)."""
    ensure_users_schema()
    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        no_anggota = (request.form.get('no_anggota') or '').strip()
        pwd1 = request.form.get('password') or ''
        pwd2 = request.form.get('password_confirm') or ''

        user = get_user_by_username(username)
        if not user or user.get('role') != 'user' or not user.get('id_anggota'):
            flash('Data tidak cocok atau akun tidak memenuhi syarat reset mandiri. Hubungi admin.', 'danger')
            return render_template('lupa_password.html'), 400

        anggota = get_anggota_by_id(user.get('id_anggota'))
        if (not anggota or
                (anggota.get('no_anggota') or '').strip().upper() != no_anggota.strip().upper()):
            flash('Data tidak cocok atau akun tidak memenuhi syarat reset mandiri. Hubungi admin.', 'danger')
            return render_template('lupa_password.html'), 400

        if pwd1 != pwd2:
            flash('Konfirmasi password tidak sama.', 'danger')
            return render_template('lupa_password.html'), 400

        ok, msg = validate_password_policy(pwd1)
        if not ok:
            flash(msg, 'danger')
            return render_template('lupa_password.html'), 400

        users = baca_csv(FILE_USERS)
        for u in users:
            if u.get('id_user') == user.get('id_user'):
                u['password_hash'] = generate_password_hash(pwd1)
                break
        tulis_csv(FILE_USERS, users, ['id_user', 'username', 'password_hash', 'role', 'id_anggota', 'created_at'])
        flash('Password berhasil diubah. Silakan login.', 'success')
        return redirect(url_for('login'))

    return render_template('lupa_password.html')


@app.route('/logout')
@login_required
def logout():
    session.clear()
    flash('Anda berhasil logout.', 'success')
    return redirect(url_for('landing'))


# ══════════════════════════════════════════════
#  ROUTE: MANAJEMEN USER (ADMIN)
# ══════════════════════════════════════════════
@app.route('/users')
@admin_required
def users_index():
    ensure_users_schema()
    users = baca_csv(FILE_USERS)
    anggota = baca_csv(FILE_ANGGOTA)
    anggota_map = {a.get('id_anggota'): a for a in anggota}
    users.sort(key=lambda u: (u.get('role') != 'admin', (u.get('username') or '').lower()))
    return render_template('users.html', users=users, anggota=anggota, anggota_map=anggota_map)


@app.route('/users/tambah', methods=['POST'])
@admin_required
def users_tambah():
    ensure_users_schema()
    username = (request.form.get('username') or '').strip()
    password = request.form.get('password') or ''
    role = request.form.get('role') or 'user'
    id_anggota = request.form.get('id_anggota') or ''

    if not username or not password:
        flash('Username dan password wajib diisi.', 'danger')
        return redirect(url_for('users_index'))
    ok_pol, msg_pol = validate_password_policy(password)
    if not ok_pol:
        flash(msg_pol, 'danger')
        return redirect(url_for('users_index'))
    if role not in ('admin', 'user'):
        flash('Role tidak valid.', 'danger')
        return redirect(url_for('users_index'))
    if get_user_by_username(username):
        flash('Username sudah dipakai. Gunakan username lain.', 'danger')
        return redirect(url_for('users_index'))

    if role == 'user' and not id_anggota:
        flash('Untuk role user, wajib pilih anggota.', 'danger')
        return redirect(url_for('users_index'))

    if id_anggota:
        anggota = baca_csv(FILE_ANGGOTA)
        if not any(a.get('id_anggota') == id_anggota for a in anggota):
            flash('Anggota tidak ditemukan.', 'danger')
            return redirect(url_for('users_index'))

    users = baca_csv(FILE_USERS)
    users.append({
        'id_user': str(uuid.uuid4()),
        'username': username,
        'password_hash': generate_password_hash(password),
        'role': role,
        'id_anggota': id_anggota if role == 'user' else '',
        'created_at': datetime.now().strftime('%Y-%m-%d')
    })
    tulis_csv(FILE_USERS, users, ['id_user', 'username', 'password_hash', 'role', 'id_anggota', 'created_at'])
    flash('Akun berhasil dibuat.', 'success')
    return redirect(url_for('users_index'))


@app.route('/users/reset_password/<id_user>', methods=['POST'])
@admin_required
def users_reset_password(id_user):
    ensure_users_schema()
    new_password = request.form.get('new_password') or ''
    if not new_password:
        flash('Password baru wajib diisi.', 'danger')
        return redirect(url_for('users_index'))
    ok_pol, msg_pol = validate_password_policy(new_password)
    if not ok_pol:
        flash(msg_pol, 'danger')
        return redirect(url_for('users_index'))

    users = baca_csv(FILE_USERS)
    found = False
    for u in users:
        if u.get('id_user') == id_user:
            u['password_hash'] = generate_password_hash(new_password)
            found = True
            break
    if not found:
        flash('User tidak ditemukan.', 'danger')
        return redirect(url_for('users_index'))

    tulis_csv(FILE_USERS, users, ['id_user', 'username', 'password_hash', 'role', 'id_anggota', 'created_at'])
    flash('Password berhasil direset.', 'success')
    return redirect(url_for('users_index'))


@app.route('/users/hapus/<id_user>')
@admin_required
def users_hapus(id_user):
    ensure_users_schema()
    current_id = session.get('id_user')
    if current_id and current_id == id_user:
        flash('Tidak bisa menghapus akun yang sedang login.', 'danger')
        return redirect(url_for('users_index'))

    users = baca_csv(FILE_USERS)
    before = len(users)
    users = [u for u in users if u.get('id_user') != id_user]
    if len(users) == before:
        flash('User tidak ditemukan.', 'danger')
        return redirect(url_for('users_index'))
    tulis_csv(FILE_USERS, users, ['id_user', 'username', 'password_hash', 'role', 'id_anggota', 'created_at'])
    flash('User berhasil dihapus.', 'success')
    return redirect(url_for('users_index'))


def bunga_untuk_jenis_pinjaman(jenis: str, tenor_bulan: int) -> float:
    """Bunga bulanan per produk + validasi tenor tiap jenis pinjaman."""
    j = (jenis or '').strip()
    if j == 'Solusi Cepat':
        if 1 <= tenor_bulan <= 2:
            return 2.0
        raise ValueError('Tenor Solusi Cepat wajib 1-2 bulan.')
    if j == 'Jangka Pendek':
        if 1 <= tenor_bulan <= 12:
            return 1.5
        raise ValueError('Tenor Jangka Pendek wajib 1-12 bulan.')
    if j == 'Jangka Panjang':
        if 13 <= tenor_bulan <= 24:
            return 0.80
        if 25 <= tenor_bulan <= 120:
            return 0.75
        raise ValueError('Tenor Jangka Panjang wajib 13-24 bulan atau 25-120 bulan.')
    if j == 'Modal Usaha':
        if 1 <= tenor_bulan <= 120:
            return 0.50
        raise ValueError('Tenor Modal Usaha wajib 1-120 bulan.')
    if j in JENIS_PINJAMAN:
        return float(JENIS_PINJAMAN[j]['bunga'])
    return bunga_dari_tenor(tenor_bulan)


# ══════════════════════════════════════════════
#  ROUTE: DASHBOARD
# ══════════════════════════════════════════════
@app.route('/dashboard')
@login_required
def dashboard():
    ensure_simpanan_schema()
    ensure_simpanan_transaksi_schema()
    anggota = baca_csv(FILE_ANGGOTA)
    simpanan = baca_csv(FILE_SIMPANAN)
    simpanan_transaksi = baca_csv(FILE_SIMPANAN_TRANSAKSI)
    pinjaman = baca_csv(FILE_PINJAMAN)

    if not is_current_user_admin():
        id_anggota = get_current_user_id_anggota()
        anggota = [a for a in anggota if a.get('id_anggota') == id_anggota]
        simpanan = [s for s in simpanan if s.get('id_anggota') == id_anggota]
        simpanan_transaksi = [t for t in simpanan_transaksi if t.get('id_anggota') == id_anggota]
        pinjaman = [p for p in pinjaman if p.get('id_anggota') == id_anggota]

    total_anggota = len(anggota)
    total_simpanan = sum(float(s.get('total_simpanan') or 0) for s in simpanan)
    pv_aktif = [p for p in pinjaman if p.get('status') == 'Disetujui']
    total_pinjaman = sum(float(p.get('plafon') or 0) for p in pv_aktif)
    total_pinjaman_beredar = sum(saldo_pinjaman_aktual(p) for p in pv_aktif)

    # Saldo simpanan per jenis dihitung dari transaksi, dengan fallback untuk data lama.
    simpanan_per_jenis = {j: 0.0 for j in JENIS_SIMPANAN}
    total_transaksi_simpanan = 0.0
    for t in simpanan_transaksi:
        jenis = (t.get('jenis_simpanan') or 'Manasuka').strip() or 'Manasuka'
        if jenis not in simpanan_per_jenis:
            jenis = 'Manasuka'
        try:
            nominal = float(t.get('jumlah') or 0)
        except (TypeError, ValueError):
            nominal = 0.0
        simpanan_per_jenis[jenis] += nominal
        total_transaksi_simpanan += nominal
    if total_simpanan > total_transaksi_simpanan:
        simpanan_per_jenis['Manasuka'] += (total_simpanan - total_transaksi_simpanan)

    # Ringkasan aktivitas: gabungan semua transaksi simpanan, pinjaman, dan cicilan.
    transaksi_terakhir = []

    def append_transaksi(tipe: str, tanggal: str, nama: str, keterangan: str, status: str = ''):
        transaksi_terakhir.append({
            'tipe': tipe,
            'tanggal': tanggal or '-',
            'nama': nama or '',
            'keterangan': keterangan,
            'status': status or '-',
        })

    for t in simpanan_transaksi:
        append_transaksi(
            'Simpanan',
            t.get('tanggal', '-'),
            t.get('nama_anggota', ''),
            f"Simpanan {t.get('jenis_simpanan', 'Manasuka')} — Rp {float(t.get('jumlah') or 0):,.0f}",
            'Disetujui',
        )

    for p in baca_csv(FILE_PINJAMAN):
        tanggal_evt = p.get('tanggal_lunas') or p.get('tanggal_pengajuan') or '-'
        status_evt = p.get('status') or '-'
        jenis_evt = p.get('jenis_pinjaman') or 'Pinjaman'
        keterangan_evt = f"{jenis_evt} — Rp {float(p.get('plafon') or 0):,.0f}"
        if status_evt == 'Menunggu':
            keterangan_evt = f"Pengajuan {keterangan_evt}"
        elif status_evt == 'Disetujui':
            keterangan_evt = f"Pinjaman disetujui {keterangan_evt}"
        elif status_evt == 'Lunas':
            keterangan_evt = f"Pinjaman lunas {keterangan_evt}"
        elif status_evt == 'Ditolak':
            keterangan_evt = f"Pengajuan ditolak {keterangan_evt}"
        append_transaksi('Pinjaman', tanggal_evt, p.get('nama_anggota', ''), keterangan_evt, status_evt)

    for c in baca_csv(FILE_PINJAMAN_CICILAN):
        status_evt = c.get('status') or '-'
        tanggal_evt = c.get('tanggal_konfirmasi') or c.get('tanggal_pengajuan') or '-'
        if status_evt == 'Menunggu':
            label_evt = 'Pengajuan cicilan'
        elif status_evt == 'Disetujui':
            label_evt = 'Cicilan diterima'
        elif status_evt == 'Gagal Bayar':
            label_evt = 'Cicilan gagal bayar'
        elif status_evt == 'Ditolak':
            label_evt = 'Cicilan ditolak'
        else:
            label_evt = 'Cicilan'
        keterangan_evt = f"{label_evt} — Rp {float(c.get('jumlah') or 0):,.0f}"
        if c.get('metode_pembayaran'):
            keterangan_evt += f" ({c.get('metode_pembayaran')})"
        status_tampil = 'Cicilan diterima' if status_evt == 'Disetujui' else status_evt
        append_transaksi('Cicilan', tanggal_evt, c.get('nama_anggota', ''), keterangan_evt, status_tampil)

    transaksi_terakhir.sort(key=lambda x: x.get('tanggal', ''), reverse=True)

    pengajuan_anggota = []
    jumlah_pengajuan_menunggu = 0
    if is_current_user_admin():
        pengajuan_anggota = baca_csv(FILE_PENDAFTARAN_ANGGOTA)
        pengajuan_anggota.sort(key=lambda x: x.get('tanggal_pengajuan', ''), reverse=True)
        jumlah_pengajuan_menunggu = sum(1 for p in pengajuan_anggota if p.get('status') == 'Menunggu')
        pengajuan_anggota = pengajuan_anggota[:8]

    return render_template('dashboard.html',
                           total_anggota=total_anggota,
                           total_simpanan=total_simpanan,
                           total_pinjaman=total_pinjaman,
                           total_pinjaman_beredar=total_pinjaman_beredar,
                           simpanan_per_jenis=simpanan_per_jenis,
                           transaksi_terakhir=transaksi_terakhir,
                           pengajuan_anggota=pengajuan_anggota,
                           jumlah_pengajuan_menunggu=jumlah_pengajuan_menunggu)


@app.route('/pengajuan-anggota/konfirmasi/<id_pengajuan>')
@admin_required
def konfirmasi_pengajuan_anggota(id_pengajuan):
    ensure_pendaftaran_schema()
    pengajuan = baca_csv(FILE_PENDAFTARAN_ANGGOTA)
    found = False
    for p in pengajuan:
        if p.get('id_pengajuan') == id_pengajuan:
            # Auto-create anggota hanya sekali saat pengajuan pertama kali disetujui.
            if not p.get('id_anggota_dibuat'):
                anggota = baca_csv(FILE_ANGGOTA)
                id_anggota = str(uuid.uuid4())
                no_anggota = generate_no_anggota_berikutnya(anggota)
                ph = parse_rupiah_to_float(p.get('penghasilan_bersih', '0'))
                cl = parse_rupiah_to_float(p.get('cicilan_lain', '0'))
                anggota.append({
                    'id_anggota': id_anggota,
                    'no_anggota': no_anggota,
                    'nik': '',
                    'nama': p.get('nama', ''),
                    'alamat': p.get('alamat', ''),
                    'no_telp': p.get('no_telp', ''),
                    'tgl_bergabung': datetime.now().strftime('%Y-%m-%d'),
                    'penghasilan_bersih': str(int(ph)),
                    'cicilan_lain': str(int(cl))
                })
                tulis_csv(FILE_ANGGOTA, anggota, ANGGOTA_FIELDNAMES)
                p['id_anggota_dibuat'] = id_anggota
                p['no_anggota_dibuat'] = no_anggota
            p['status'] = 'Disetujui'
            p['catatan_admin'] = (
                f"Dikonfirmasi {session.get('user', 'admin')} pada {datetime.now().strftime('%Y-%m-%d')} "
                f"(anggota: {p.get('no_anggota_dibuat', '-')})"
            )
            found = True
            break
    if not found:
        flash('Pengajuan anggota tidak ditemukan.', 'danger')
        return redirect(url_for('dashboard'))
    tulis_csv(FILE_PENDAFTARAN_ANGGOTA, pengajuan, PENDAFTARAN_FIELDNAMES)
    flash('Pengajuan anggota disetujui.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/pengajuan-anggota/tolak/<id_pengajuan>')
@admin_required
def tolak_pengajuan_anggota(id_pengajuan):
    ensure_pendaftaran_schema()
    pengajuan = baca_csv(FILE_PENDAFTARAN_ANGGOTA)
    found = False
    for p in pengajuan:
        if p.get('id_pengajuan') == id_pengajuan:
            p['status'] = 'Ditolak'
            p['catatan_admin'] = f"Ditolak {session.get('user', 'admin')} pada {datetime.now().strftime('%Y-%m-%d')}"
            found = True
            break
    if not found:
        flash('Pengajuan anggota tidak ditemukan.', 'danger')
        return redirect(url_for('dashboard'))
    tulis_csv(FILE_PENDAFTARAN_ANGGOTA, pengajuan, PENDAFTARAN_FIELDNAMES)
    flash('Pengajuan anggota ditolak.', 'warning')
    return redirect(url_for('dashboard'))


# ══════════════════════════════════════════════
#  ROUTE: ANGGOTA
# ══════════════════════════════════════════════
@app.route('/anggota')
@login_required
def halaman_anggota():
    ensure_anggota_schema()
    anggota = baca_csv(FILE_ANGGOTA)
    if not is_current_user_admin():
        id_anggota = get_current_user_id_anggota()
        anggota = [a for a in anggota if a.get('id_anggota') == id_anggota]
    else:
        for a in anggota:
            info = info_pinjaman_dsr_anggota(a)
            a['_dsr_persen'] = info['dsr_persen']
            a['_kap_cicilan'] = info['kapasitas_cicilan']
    return render_template('anggota.html', anggota=anggota)


@app.route('/anggota/tambah', methods=['POST'])
@admin_required
def tambah_anggota():
    ensure_anggota_schema()
    data = baca_csv(FILE_ANGGOTA)
    nik = normalize_nik(request.form.get('nik'))
    if nik and not is_valid_nik(nik):
        flash('NIK harus terdiri dari tepat 16 digit angka.', 'danger')
        return redirect('/anggota')
    nama = (request.form.get('nama') or '').strip()
    alamat = (request.form.get('alamat') or '').strip()
    no_telp = (request.form.get('no_telp') or '').strip()
    penghasilan_bersih = str(parse_rupiah_to_float(request.form.get('penghasilan_bersih', '0')))
    cicilan_lain = str(parse_rupiah_to_float(request.form.get('cicilan_lain', '0')))
    no_anggota = (request.form.get('no_anggota') or '').strip()
    if not no_anggota:
        no_anggota = generate_no_anggota_berikutnya(data)
    target = None
    if nik:
        target = next((a for a in data if (a.get('nik') or '').strip().upper() == nik.upper()), None)
    if target is None and no_anggota:
        target = next((a for a in data if (a.get('no_anggota') or '').strip().upper() == no_anggota.upper()), None)

    if target:
        target['no_anggota'] = no_anggota or target.get('no_anggota', '')
        target['nik'] = nik or target.get('nik', '')
        target['nama'] = nama or target.get('nama', '')
        target['alamat'] = alamat or target.get('alamat', '')
        target['no_telp'] = no_telp or target.get('no_telp', '')
        target['penghasilan_bersih'] = penghasilan_bersih
        target['cicilan_lain'] = cicilan_lain
        flash('Data anggota berhasil diperbarui tanpa membuat data ganda.', 'success')
    else:
        data.append({
            'id_anggota': str(uuid.uuid4()),
            'no_anggota': no_anggota,
            'nik': nik,
            'nama': nama,
            'alamat': alamat,
            'no_telp': no_telp,
            'tgl_bergabung': datetime.now().strftime('%Y-%m-%d'),
            'penghasilan_bersih': penghasilan_bersih,
            'cicilan_lain': cicilan_lain,
        })
        flash('Data anggota berhasil ditambahkan.', 'success')
    tulis_csv(FILE_ANGGOTA, data, ANGGOTA_FIELDNAMES)
    return redirect('/anggota')


@app.route('/anggota/hapus/<id_anggota>')
@admin_required
def hapus_anggota(id_anggota):
    data = baca_csv(FILE_ANGGOTA)
    data = [a for a in data if a['id_anggota'] != id_anggota]
    tulis_csv(FILE_ANGGOTA, data, ANGGOTA_FIELDNAMES)
    return redirect('/anggota')


@app.route('/anggota/edit/<id_anggota>', methods=['GET', 'POST'])
@admin_required
def edit_anggota(id_anggota):
    ensure_anggota_schema()
    data = baca_csv(FILE_ANGGOTA)
    idx = next((i for i, a in enumerate(data) if a.get('id_anggota') == id_anggota), None)
    if idx is None:
        flash('Anggota tidak ditemukan.', 'danger')
        return redirect(url_for('halaman_anggota'))

    if request.method == 'POST':
        nik = normalize_nik(request.form.get('nik'))
        if nik and not is_valid_nik(nik):
            flash('NIK harus terdiri dari tepat 16 digit angka.', 'danger')
            return redirect(url_for('edit_anggota', id_anggota=id_anggota))
        data[idx]['nik'] = nik
        data[idx]['nama'] = (request.form.get('nama') or '').strip()
        data[idx]['alamat'] = (request.form.get('alamat') or '').strip()
        data[idx]['no_telp'] = (request.form.get('no_telp') or '').strip()
        data[idx]['penghasilan_bersih'] = str(int(parse_rupiah_to_float(request.form.get('penghasilan_bersih', '0'))))
        data[idx]['cicilan_lain'] = str(int(parse_rupiah_to_float(request.form.get('cicilan_lain', '0'))))
        tulis_csv(FILE_ANGGOTA, data, ANGGOTA_FIELDNAMES)
        flash('Data anggota berhasil diperbarui.', 'success')
        return redirect(url_for('halaman_anggota'))

    return render_template('anggota_edit.html', anggota=data[idx])


@app.route('/anggota/import-excel', methods=['POST'])
@admin_required
def import_anggota_excel():
    """Import anggota dari Excel dengan smart upsert berdasarkan NIK, fallback No Anggota."""
    ensure_anggota_schema()
    if load_workbook is None:
        flash('Fitur import Excel membutuhkan openpyxl.', 'danger')
        return redirect(url_for('halaman_anggota'))

    file = request.files.get('file_excel')
    if not file or not file.filename:
        flash('File Excel belum dipilih.', 'danger')
        return redirect(url_for('halaman_anggota'))
    if not file.filename.lower().endswith(('.xlsx', '.xlsm')):
        flash('Format file tidak didukung. Gunakan .xlsx atau .xlsm', 'danger')
        return redirect(url_for('halaman_anggota'))

    try:
        wb = load_workbook(file, data_only=True)
    except Exception as e:
        flash(f'Gagal membaca file Excel: {e}', 'danger')
        return redirect(url_for('halaman_anggota'))

    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    headers_raw = next(rows_iter, None)
    if not headers_raw:
        flash('File Excel kosong.', 'danger')
        return redirect(url_for('halaman_anggota'))

    def norm(h):
        return str(h or '').strip().lower().replace(' ', '_')

    header_map = {norm(h): i for i, h in enumerate(headers_raw)}

    def get_val(row, *aliases):
        for a in aliases:
            idx = header_map.get(a)
            if idx is not None and idx < len(row):
                v = row[idx]
                return '' if v is None else str(v).strip()
        return ''

    anggota = baca_csv(FILE_ANGGOTA)
    by_nik = {
        normalize_nik(a.get('nik')): a
        for a in anggota
        if normalize_nik(a.get('nik'))
    }
    by_no = {(a.get('no_anggota') or '').strip(): a for a in anggota if (a.get('no_anggota') or '').strip()}

    added = 0
    updated = 0
    max_no = 0
    for a in anggota:
        no = (a.get('no_anggota') or '').strip()
        if no.startswith('AGT-'):
            num = no.replace('AGT-', '')
            if num.isdigit():
                max_no = max(max_no, int(num))

    invalid_nik_rows = 0
    for row in rows_iter:
        nik = normalize_nik(get_val(row, 'nik'))
        no_anggota = get_val(row, 'no_anggota')
        nama = get_val(row, 'nama')
        alamat = get_val(row, 'alamat')
        no_telp = get_val(row, 'no_telp', 'no_telepon', 'telepon')
        penghasilan = get_val(row, 'penghasilan_bersih')
        cicilan = get_val(row, 'cicilan_lain')
        if nik and not is_valid_nik(nik):
            invalid_nik_rows += 1
            continue
        if not (nik or no_anggota or nama):
            continue

        target = None
        if nik and nik in by_nik:
            target = by_nik[nik]
        elif no_anggota and no_anggota in by_no:
            target = by_no[no_anggota]

        if target:
            target['nik'] = nik or target.get('nik', '')
            target['nama'] = nama or target.get('nama', '')
            target['alamat'] = alamat or target.get('alamat', '')
            target['no_telp'] = no_telp or target.get('no_telp', '')
            target['penghasilan_bersih'] = str(parse_rupiah_to_float(penghasilan or target.get('penghasilan_bersih', '0')))
            target['cicilan_lain'] = str(parse_rupiah_to_float(cicilan or target.get('cicilan_lain', '0')))
            updated += 1
            continue

        if not no_anggota:
            max_no += 1
            no_anggota = f"AGT-{max_no:04d}"

        new_row = {
            'id_anggota': str(uuid.uuid4()),
            'no_anggota': no_anggota,
            'nik': nik,
            'nama': nama,
            'alamat': alamat,
            'no_telp': no_telp,
            'tgl_bergabung': datetime.now().strftime('%Y-%m-%d'),
            'penghasilan_bersih': str(parse_rupiah_to_float(penghasilan or '0')),
            'cicilan_lain': str(parse_rupiah_to_float(cicilan or '0')),
        }
        anggota.append(new_row)
        if nik:
            by_nik[nik] = new_row
        if no_anggota:
            by_no[no_anggota] = new_row
        added += 1

    tulis_csv(FILE_ANGGOTA, anggota, ANGGOTA_FIELDNAMES)
    flash(f'Import selesai: {added} anggota baru ditambahkan, {updated} data anggota diperbarui.', 'success')
    if invalid_nik_rows:
        flash(f'{invalid_nik_rows} baris dilewati karena NIK tidak valid (harus 16 digit angka).', 'warning')
    return redirect(url_for('halaman_anggota'))


def _norm_csv_header(h):
    return str(h or '').strip().lower().replace(' ', '_').replace('-', '_')


def _map_import_csv_headers(fieldnames):
    """Map normalized header -> canonical key."""
    if not fieldnames:
        return None
    norm = {_norm_csv_header(h): h for h in fieldnames}
    aliases = {
        'nama': ('nama', 'name'),
        'nik': ('nik',),
        'no_anggota': ('no_anggota', 'no anggota'),
        'no_hp': ('no_hp', 'nohp', 'no_telp', 'telepon', 'telp', 'hp'),
        'alamat': ('alamat', 'address'),
        'simpanan': ('simpanan', 'tabungan'),
        'pinjaman': ('pinjaman', 'hutang'),
        'tenor_bulan': ('tenor_bulan', 'tenor', 'jangka_waktu', 'lama_bulan'),
    }
    out = {}
    for key, alist in aliases.items():
        for a in alist:
            if a in norm:
                out[key] = norm[a]
                break
    if 'nik' not in out:
        return None
    # Hindari file "satu kolom".
    if len(out.keys()) < 2:
        return None
    return out


def _parse_opt_float_cell(raw):
    s = (raw or '').strip()
    if not s:
        return None
    return parse_rupiah_to_float(s)


def parse_anggota_csv_upload(file_storage):
    """Validasi & parse file Excel ringkasan anggota. Mengembalikan dict preview atau raise ValueError."""
    if not file_storage or not file_storage.filename:
        raise ValueError('File belum dipilih.')
    ext = os.path.splitext(file_storage.filename.lower())[1]
    if ext not in ('.xlsx', '.xlsm'):
        raise ValueError('Gunakan file Excel (.xlsx atau .xlsm).')
    file_storage.stream.seek(0)
    raw = file_storage.read()
    if len(raw) > 2 * 1024 * 1024:
        raise ValueError('Ukuran file maksimal 2 MB.')
    rows_src = []
    detected_columns = []
    if load_workbook is None:
        raise ValueError('Import Excel membutuhkan openpyxl.')
    wb = load_workbook(io.BytesIO(raw), data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        raise ValueError('File kosong.')
    headers = [str(h or '').strip() for h in all_rows[0]]
    header_non_empty = [h for h in headers if h]
    if len(header_non_empty) < 2:
        raise ValueError('Format Excel tidak valid: header harus terpisah di beberapa kolom, bukan satu kolom gabungan.')
    fieldnames = headers
    for rr in all_rows[1:]:
        row_map = {}
        for i, h in enumerate(headers):
            v = rr[i] if i < len(rr) else ''
            row_map[h] = '' if v is None else str(v).strip()
        # Cegah pola data gabungan dalam satu kolom (mis: "Nama,NIK,No HP,...")
        isi_terisi = [v for v in row_map.values() if str(v or '').strip()]
        if len(isi_terisi) == 1 and any(sep in isi_terisi[0] for sep in (',', ';', '|', '\t')):
            raise ValueError('Format Excel tidak valid: data terdeteksi gabungan dalam satu kolom. Pisahkan per kolom.')
        rows_src.append(row_map)
    header_map = _map_import_csv_headers(fieldnames)
    if not header_map:
        raise ValueError(
            'Header tidak valid. Wajib ada kolom nik; disarankan: nama,nik,no hp,simpanan,pinjaman,tenor'
        )
    detected_columns = [k for k in ('nama', 'nik', 'no_anggota', 'no_hp', 'alamat', 'simpanan', 'pinjaman', 'tenor_bulan') if k in header_map]
    rows_out = []
    line_no = 1
    for row in rows_src:
        line_no += 1
        nama = (row.get(header_map['nama']) or '').strip() if 'nama' in header_map else ''
        nik = normalize_nik(row.get(header_map['nik']))
        no_anggota = (row.get(header_map['no_anggota']) or '').strip() if 'no_anggota' in header_map else ''
        no_hp = (row.get(header_map['no_hp']) or '').strip() if 'no_hp' in header_map else ''
        alamat = (row.get(header_map['alamat']) or '').strip() if 'alamat' in header_map else ''
        simp_raw = row.get(header_map['simpanan']) if 'simpanan' in header_map else ''
        pin_raw = row.get(header_map['pinjaman']) if 'pinjaman' in header_map else ''
        tenor_raw = row.get(header_map['tenor_bulan']) if 'tenor_bulan' in header_map else ''
        err = []
        if not nik:
            err.append('NIK wajib diisi')
        elif not is_valid_nik(nik):
            err.append('NIK harus 16 digit angka')
        simpanan = None
        pinjaman = None
        tenor_bulan = None
        try:
            if (simp_raw or '').strip():
                simpanan = _parse_opt_float_cell(simp_raw)
        except Exception:
            err.append('simpanan bukan angka valid')
        try:
            if (pin_raw or '').strip():
                pinjaman = _parse_opt_float_cell(pin_raw)
        except Exception:
            err.append('pinjaman bukan angka valid')
        try:
            if (tenor_raw or '').strip():
                tenor_bulan = int(float(str(tenor_raw).strip()))
                if tenor_bulan < 1 or tenor_bulan > 120:
                    err.append('tenor harus 1-120 bulan')
        except Exception:
            err.append('tenor bukan angka valid')
        rows_out.append({
            'line': line_no,
            'nama': nama,
            'nik': nik,
            'no_anggota': no_anggota,
            'no_hp': no_hp,
            'alamat': alamat,
            'simpanan': simpanan,
            'pinjaman': pinjaman,
            'tenor_bulan': tenor_bulan,
            'errors': err,
        })
    if not rows_out:
        raise ValueError('Tidak ada baris data (selain header).')
    if len(rows_out) > 500:
        raise ValueError('Maksimal 500 baris data per unggah.')
    return {'rows': rows_out, 'filename': file_storage.filename, 'detected_columns': detected_columns}


def _append_import_log(berhasil: int, gagal: int, mode: str, nama_file: str, catatan: str):
    ensure_import_log_schema()
    rows = baca_csv(FILE_IMPORT_LOG)
    rows.append({
        'waktu': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'user': session.get('user') or '',
        'mode': mode,
        'nama_file': nama_file,
        'berhasil': str(berhasil),
        'gagal': str(gagal),
        'catatan': catatan[:2000],
    })
    tulis_csv(
        FILE_IMPORT_LOG,
        rows,
        ['waktu', 'user', 'mode', 'nama_file', 'berhasil', 'gagal', 'catatan'],
    )


def jalankan_import_csv_ringkasan(preview_rows: list, mode: str) -> tuple:
    """Legacy preview import (dinonaktifkan di UI)."""
    return 0, len(preview_rows or []), 'Import ringkasan tidak lagi digunakan.'


def _float_impor_csv(val) -> float:
    if val is None or str(val).strip() == '':
        return 0.0
    return float(str(val).strip().replace(',', '').replace(' ', ''))


def upsert_anggota_dari_baris_impor(anggota_list: list, id_a: str, nama: str, nik: str, no_hp: str, alamat: str, no_anggota: str = '') -> None:
    no_anggota = (no_anggota or '').strip()
    for a in anggota_list:
        if a.get('id_anggota') == id_a:
            if nama:
                a['nama'] = nama
            if nik:
                a['nik'] = nik
            if no_hp:
                a['no_telp'] = no_hp
            if alamat:
                a['alamat'] = alamat
            if no_anggota:
                bentrok = next(
                    (
                        x for x in anggota_list
                        if x.get('id_anggota') != id_a
                        and (x.get('no_anggota') or '').strip().upper() == no_anggota.upper()
                    ),
                    None,
                )
                if not bentrok:
                    a['no_anggota'] = no_anggota
            return
    no_baru = no_anggota or generate_no_anggota_berikutnya(anggota_list)
    bentrok_baru = next(
        (x for x in anggota_list if (x.get('no_anggota') or '').strip().upper() == no_baru.upper()),
        None,
    )
    if bentrok_baru:
        no_baru = generate_no_anggota_berikutnya(anggota_list)
    anggota_list.append({
        'id_anggota': id_a,
        'no_anggota': no_baru,
        'nik': nik or '',
        'nama': nama or 'Anggota',
        'alamat': alamat or '',
        'no_telp': no_hp or '',
        'tgl_bergabung': datetime.now().strftime('%Y-%m-%d'),
        'penghasilan_bersih': '0',
        'cicilan_lain': '0',
    })


IMPORT_CSV_FIELDNAMES = [
    'nik', 'no_hp',
]


def _pick_import_value(row: dict, *keys: str) -> str:
    """Ambil nilai kolom dengan dukungan alias header (spasi/underscore/case-insensitive)."""
    if not row:
        return ''
    for k in keys:
        v = row.get(k)
        if v is not None and str(v).strip() != '':
            return str(v).strip()
    norm_map = {str(k or '').strip().lower().replace(' ', '_'): k for k in row.keys()}
    for k in keys:
        nk = str(k or '').strip().lower().replace(' ', '_')
        rk = norm_map.get(nk)
        if rk is None:
            continue
        v = row.get(rk)
        if v is not None and str(v).strip() != '':
            return str(v).strip()
    return ''


def _resolve_id_anggota_by_nik_nohp(anggota_list: list, nik: str, no_hp: str) -> str:
    """Key import anggota berdasarkan kombinasi NIK + No HP."""
    nik_key = (nik or '').strip()
    hp_key = (no_hp or '').strip()
    if not nik_key or not hp_key:
        return ''
    # Kecocokan utama: NIK + No HP
    for a in anggota_list:
        if (a.get('nik') or '').strip() == nik_key and (a.get('no_telp') or '').strip() == hp_key:
            return (a.get('id_anggota') or '').strip()
    # Fallback aman: NIK sama dan nomor lama kosong -> isi nomor dari import
    for a in anggota_list:
        if (a.get('nik') or '').strip() == nik_key and not (a.get('no_telp') or '').strip():
            a['no_telp'] = hp_key
            return (a.get('id_anggota') or '').strip()
    return str(uuid.uuid4())


def _resolve_id_anggota_by_identity(anggota_list: list, nik: str, no_hp: str, no_anggota: str = '') -> str:
    """Resolve anggota by no_anggota first, then by NIK + No HP."""
    no_anggota_key = (no_anggota or '').strip()
    if no_anggota_key:
        for a in anggota_list:
            if (a.get('no_anggota') or '').strip() == no_anggota_key:
                return (a.get('id_anggota') or '').strip()
    return _resolve_id_anggota_by_nik_nohp(anggota_list, nik, no_hp)


def kategori_pinjaman_dari_tenor(tenor_bulan: int) -> str:
    """Kelompokkan pinjaman impor menjadi Jangka Pendek atau Jangka Panjang."""
    tenor = max(int(tenor_bulan or 0), 0)
    if tenor >= PROVISI_MIN_TENOR_BULAN:
        return 'Jangka Panjang'
    return 'Jangka Pendek'


@app.route('/import_csv', methods=['POST'])
@admin_required
def import_csv_unified():
    """Unggah satu file (CSV/Excel): merge anggota, akumulasi simpanan & pinjaman (tenor max)."""
    redir = request.referrer or url_for('halaman_anggota')
    try:
        if 'file' not in request.files:
            flash('Tidak ada berkas yang diunggah.', 'danger')
            return redirect(redir)
        f = request.files['file']
        if not f.filename or not str(f.filename).lower().endswith(('.csv', '.xlsx', '.xlsm')):
            flash('Unggah file dengan ekstensi .csv atau .xlsx.', 'danger')
            return redirect(redir)
        
        # Parse file (CSV atau Excel)
        ext = os.path.splitext(f.filename.lower())[1]
        f.stream.seek(0)
        raw = f.read()
        rows_data = []
        fieldnames = []
        
        if ext == '.csv':
            text = raw.decode('utf-8-sig')
            reader = csv.DictReader(io.StringIO(text))
            rows_data = list(reader) if reader else []
            fieldnames = reader.fieldnames if reader else []
        else:
            # Excel format
            if load_workbook is None:
                flash('Fitur Excel membutuhkan openpyxl.', 'danger')
                return redirect(redir)
            try:
                wb = load_workbook(io.BytesIO(raw), data_only=True)
                ws = wb.active
                all_rows = list(ws.iter_rows(values_only=True))
                if not all_rows:
                    flash('File Excel kosong.', 'danger')
                    return redirect(redir)
                headers = [str(h or '').strip() for h in all_rows[0]]
                fieldnames = headers
                for rr in all_rows[1:]:
                    row_map = {}
                    for i, h in enumerate(headers):
                        v = rr[i] if i < len(rr) else ''
                        row_map[h] = '' if v is None else str(v).strip()
                    rows_data.append(row_map)
            except Exception as e:
                flash(f'Gagal membaca Excel: {str(e)}', 'danger')
                return redirect(redir)
        
        if not fieldnames:
            flash('File tidak memiliki header.', 'danger')
            return redirect(redir)
        
        header_norm = {str(h or '').strip().lower().replace(' ', '_') for h in fieldnames if h}
        has_nik = 'nik' in header_norm
        has_nohp = ('no_hp' in header_norm) or ('no_telp' in header_norm)
        if not has_nik or not has_nohp:
            flash('Kolom wajib tidak lengkap: nik dan no_hp/no_telp.', 'danger')
            return redirect(redir)

        ensure_anggota_schema()
        ensure_simpanan_schema()
        ensure_pinjaman_plafon_schema()
        anggota = baca_csv(FILE_ANGGOTA)
        simpanan = baca_csv(FILE_SIMPANAN)
        pinjaman = baca_csv(FILE_PINJAMAN)

        diproses = 0
        dilewati = 0
        
        for row in rows_data:
            nik = _pick_import_value(row, 'nik')
            no_hp = _pick_import_value(row, 'no_hp', 'no_telp', 'no hp', 'no telp')
            if not nik or not no_hp:
                dilewati += 1
                continue
            no_anggota = _pick_import_value(row, 'no_anggota', 'no anggota')
            id_a = _resolve_id_anggota_by_identity(anggota, nik, no_hp, no_anggota)
            if not id_a:
                dilewati += 1
                continue
            try:
                v_simp = _float_impor_csv(_pick_import_value(row, 'simpanan'))
                v_pin = _float_impor_csv(_pick_import_value(row, 'pinjaman'))
                v_ten = int(float(str(_pick_import_value(row, 'tenor', 'tenor_bulan') or '0').strip()))
            except (TypeError, ValueError):
                dilewati += 1
                continue
            if v_simp < 0 or v_pin < 0 or v_ten < 0:
                dilewati += 1
                continue

            nama = _pick_import_value(row, 'nama')
            alamat = _pick_import_value(row, 'alamat')
            upsert_anggota_dari_baris_impor(anggota, id_a, nama, nik, no_hp, alamat, no_anggota)
            if v_simp > 0:
                merge_akumulasi(simpanan, id_a, v_simp)
            if v_pin > 0 or v_ten > 0:
                merge_pinjaman_akumulasi(
                    pinjaman,
                    id_a,
                    v_pin,
                    v_ten,
                    kategori_pinjaman_dari_tenor(v_ten),
                )
            diproses += 1

        simpanan[:] = _dedupe_rows_simpanan(simpanan)
        pinjaman[:] = _dedupe_rows_pinjaman(pinjaman)
        tulis_csv(FILE_ANGGOTA, anggota, ANGGOTA_FIELDNAMES)
        tulis_csv(FILE_SIMPANAN, simpanan, SIMPANAN_FIELDNAMES)
        amap_fin = {a.get('id_anggota'): a for a in anggota}
        for p in pinjaman:
            ag = amap_fin.get(p.get('id_anggota'))
            if ag:
                p['nama_anggota'] = ag.get('nama', '')
                p['no_anggota'] = ag.get('no_anggota', '')
        tulis_csv(FILE_PINJAMAN, pinjaman, PINJAMAN_FIELDNAMES)
        flash(
            f'Import selesai: {diproses} baris diproses, {dilewati} baris dilewati (tidak valid).',
            'success',
        )
    except Exception as ex:
        flash(f'Import gagal: {ex}', 'danger')
    return redirect(redir)


@app.route('/anggota/import-csv', methods=['GET'])
@admin_required
def halaman_import_csv_anggota():
    ensure_import_log_schema()
    logs = baca_csv(FILE_IMPORT_LOG)
    logs = sorted(logs, key=lambda x: x.get('waktu', ''), reverse=True)[:50]
    return render_template('import_csv_anggota.html', import_log=logs)


@app.route('/anggota/import-csv/sample')
@admin_required
def download_sample_csv_anggota():
    if Workbook is None:
        flash('Fitur unduh contoh Excel membutuhkan openpyxl.', 'danger')
        return redirect(url_for('halaman_import_csv_anggota'))

    wb = Workbook()
    ws = wb.active
    ws.title = 'Template Import'
    ws.append(['nama', 'nik', 'no_anggota', 'no_hp', 'alamat', 'simpanan', 'pinjaman', 'tenor_bulan'])
    ws.append(['Budi Santoso', '3174xxxxxxxxxxxx', 'AG0001', '081234567890', 'Jakarta', '1500000', '40000000', '40'])
    ws.append(['Siti Aulia', '3273xxxxxxxxxxxx', 'AG0002', '081298765432', 'Bandung', '500000', '0', ''])

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return send_file(
        bio,
        as_attachment=True,
        download_name='template_import_gabungan.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@app.route('/anggota/import-csv/preview', methods=['POST'])
@admin_required
def preview_import_csv_anggota():
    try:
        if 'file_csv' not in request.files:
            flash('Tidak ada file yang dipilih.', 'danger')
            return redirect(url_for('halaman_import_csv_anggota'))
        payload = parse_anggota_csv_upload(request.files['file_csv'])
        token = token_hex(16)
        preview_path = os.path.join(IMPORT_PREVIEW_DIR, f'{token}.json')
        with open(preview_path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False)
        return render_template(
            'import_csv_preview.html',
            token=token,
            filename=payload.get('filename', ''),
            rows=payload.get('rows', []),
            detected_columns=payload.get('detected_columns', []),
        )
    except ValueError as ex:
        flash(str(ex), 'danger')
        return redirect(url_for('halaman_import_csv_anggota'))
    except Exception as ex:
        flash(f'Gagal membuat preview import: {ex}', 'danger')
        return redirect(url_for('halaman_import_csv_anggota'))


@app.route('/anggota/import-csv/preview', methods=['GET'])
@admin_required
def tampil_preview_import_csv():
    token = (request.args.get('token') or '').strip()
    if not token:
        flash('Token preview tidak ditemukan.', 'warning')
        return redirect(url_for('halaman_import_csv_anggota'))
    preview_path = os.path.join(IMPORT_PREVIEW_DIR, f'{token}.json')
    if not os.path.exists(preview_path):
        flash('Preview tidak ditemukan atau sudah kedaluwarsa.', 'warning')
        return redirect(url_for('halaman_import_csv_anggota'))
    with open(preview_path, 'r', encoding='utf-8') as f:
        payload = json.load(f)
    return render_template(
        'import_csv_preview.html',
        token=token,
        filename=payload.get('filename', ''),
        rows=payload.get('rows', []),
        detected_columns=payload.get('detected_columns', []),
    )


@app.route('/anggota/import-csv/execute', methods=['POST'])
@admin_required
def execute_import_csv_anggota():
    token = (request.form.get('token') or '').strip()
    mode = (request.form.get('mode') or 'append').strip().lower()
    if mode not in ('append', 'overwrite'):
        mode = 'append'
    if not token:
        flash('Token preview tidak valid.', 'danger')
        return redirect(url_for('halaman_import_csv_anggota'))

    preview_path = os.path.join(IMPORT_PREVIEW_DIR, f'{token}.json')
    if not os.path.exists(preview_path):
        flash('Preview tidak ditemukan atau sudah kedaluwarsa.', 'warning')
        return redirect(url_for('halaman_import_csv_anggota'))

    try:
        with open(preview_path, 'r', encoding='utf-8') as f:
            payload = json.load(f)
        rows = payload.get('rows', []) or []
        filename = payload.get('filename', 'import.xlsx')

        ensure_anggota_schema()
        ensure_simpanan_schema()
        ensure_pinjaman_plafon_schema()

        anggota = baca_csv(FILE_ANGGOTA)
        simpanan = baca_csv(FILE_SIMPANAN)
        pinjaman = baca_csv(FILE_PINJAMAN)

        berhasil = 0
        gagal = 0

        for r in rows:
            if r.get('errors'):
                gagal += 1
                continue

            nik = (r.get('nik') or '').strip()
            no_hp = (r.get('no_hp') or '').strip()
            no_anggota = (r.get('no_anggota') or '').strip()
            nama = (r.get('nama') or '').strip()
            alamat = (r.get('alamat') or '').strip()

            id_a = _resolve_id_anggota_by_identity(anggota, nik, no_hp, no_anggota)
            if not id_a and nik:
                found_by_nik = next((a for a in anggota if (a.get('nik') or '').strip() == nik), None)
                if found_by_nik:
                    id_a = (found_by_nik.get('id_anggota') or '').strip()
            if not id_a:
                id_a = str(uuid.uuid4())

            upsert_anggota_dari_baris_impor(anggota, id_a, nama, nik, no_hp, alamat, no_anggota)

            try:
                v_simp = float(r.get('simpanan') or 0)
            except (TypeError, ValueError):
                v_simp = 0.0
            try:
                v_pin = float(r.get('pinjaman') or 0)
            except (TypeError, ValueError):
                v_pin = 0.0
            try:
                v_ten = int(float(r.get('tenor_bulan') or 0))
            except (TypeError, ValueError):
                v_ten = 0

            if v_simp < 0 or v_pin < 0 or v_ten < 0:
                gagal += 1
                continue

            if mode == 'overwrite':
                if v_simp > 0:
                    target = next((s for s in simpanan if s.get('id_anggota') == id_a), None)
                    if target:
                        target['total_simpanan'] = str(round(v_simp, 2))
                    else:
                        simpanan.append({'id_anggota': id_a, 'total_simpanan': str(round(v_simp, 2))})
                if v_pin > 0 or v_ten > 0:
                    tenor_input = v_ten if v_ten > 0 else DEFAULT_TENOR_IMPORT_PINJAMAN
                    jenis = kategori_pinjaman_dari_tenor(tenor_input)
                    p_exist = next(
                        (
                            p for p in pinjaman
                            if p.get('id_anggota') == id_a
                            and (p.get('jenis_pinjaman') or '').strip() == jenis
                            and (p.get('status') or '').strip() == 'Disetujui'
                        ),
                        None,
                    )
                    if p_exist:
                        try:
                            old_plaf = float(p_exist.get('plafon') or 0)
                            old_sisa = saldo_pinjaman_aktual(p_exist)
                        except (TypeError, ValueError):
                            old_plaf = 0.0
                            old_sisa = 0.0
                        paid_nominal = max(old_plaf - old_sisa, 0.0)
                        tenor_data = int(float(p_exist.get('tenor_bulan') or p_exist.get('tenor_awal') or tenor_input))
                        tenor_data = max(tenor_data, 1)
                        bunga = float(p_exist.get('bunga_persen') or bunga_untuk_jenis_pinjaman(jenis, tenor_data))
                        total_bayar = hitung_total_bayar_tanpa_provisi(v_pin, bunga, tenor_data, jenis)
                        cic = hitung_cicilan_bulanan(v_pin, bunga, tenor_data)
                        p_exist['plafon'] = str(round(v_pin, 2))
                        p_exist['tenor_awal'] = p_exist.get('tenor_awal') or str(tenor_data)
                        p_exist['tenor_bulan'] = str(tenor_data)
                        p_exist['total_bayar'] = str(round(total_bayar, 2))
                        p_exist['cicilan_per_bulan'] = str(round(cic, 2))
                        p_exist['sisa_pinjaman'] = str(round(max(v_pin - paid_nominal, 0.0), 2))
                    else:
                        merge_pinjaman_akumulasi(pinjaman, id_a, v_pin, tenor_input, jenis)
            else:
                if v_simp > 0:
                    merge_akumulasi(simpanan, id_a, v_simp)
                if v_pin > 0 or v_ten > 0:
                    tenor_input = v_ten if v_ten > 0 else DEFAULT_TENOR_IMPORT_PINJAMAN
                    merge_pinjaman_akumulasi(
                        pinjaman,
                        id_a,
                        v_pin,
                        tenor_input,
                        kategori_pinjaman_dari_tenor(tenor_input),
                    )

            berhasil += 1

        simpanan[:] = _dedupe_rows_simpanan(simpanan)
        pinjaman[:] = _dedupe_rows_pinjaman(pinjaman)
        tulis_csv(FILE_ANGGOTA, anggota, ANGGOTA_FIELDNAMES)
        tulis_csv(FILE_SIMPANAN, simpanan, SIMPANAN_FIELDNAMES)

        amap_fin = {a.get('id_anggota'): a for a in anggota}
        for p in pinjaman:
            ag = amap_fin.get(p.get('id_anggota'))
            if ag:
                p['nama_anggota'] = ag.get('nama', '')
                p['no_anggota'] = ag.get('no_anggota', '')
        tulis_csv(FILE_PINJAMAN, pinjaman, PINJAMAN_FIELDNAMES)

        _append_import_log(
            berhasil=berhasil,
            gagal=gagal,
            mode=f'preview-{mode}',
            nama_file=filename,
            catatan='Import gabungan via preview Excel',
        )

        try:
            os.remove(preview_path)
        except OSError:
            pass

        flash(f'Import selesai: {berhasil} baris berhasil, {gagal} baris gagal/dilewati.', 'success')
        return redirect(url_for('halaman_anggota'))
    except Exception as ex:
        flash(f'Eksekusi import gagal: {ex}', 'danger')
        return redirect(url_for('halaman_import_csv_anggota'))


@app.route('/anggota/cari_nama', methods=['GET'])
@login_required
def cari_nama():
    """Autocomplete nama anggota berdasarkan No Anggota."""
    no_anggota = request.args.get('no_anggota', '')
    data = baca_csv(FILE_ANGGOTA)
    for a in data:
        if a['no_anggota'] == no_anggota:
            return jsonify({'nama': a['nama'], 'id_anggota': a['id_anggota']})
    return jsonify({'nama': '', 'id_anggota': ''})


@app.route('/anggota/cari_anggota', methods=['GET'])
@login_required
def cari_anggota_autocomplete():
    """Cari anggota berdasarkan nama (autocomplete)."""
    query = request.args.get('q', '').lower()
    data = baca_csv(FILE_ANGGOTA)
    hasil = [a for a in data if query in a['nama'].lower()]
    return jsonify(hasil)


# ══════════════════════════════════════════════
#  ROUTE: SIMPANAN
# ══════════════════════════════════════════════
@app.route('/simpanan')
@login_required
def halaman_simpanan():
    ensure_simpanan_schema()
    ensure_simpanan_transaksi_schema()
    simpanan = baca_csv(FILE_SIMPANAN)
    simpanan_transaksi = baca_csv(FILE_SIMPANAN_TRANSAKSI)
    anggota = baca_csv(FILE_ANGGOTA)
    if not is_current_user_admin():
        id_anggota = get_current_user_id_anggota()
        simpanan = [s for s in simpanan if s.get('id_anggota') == id_anggota]
        simpanan_transaksi = [t for t in simpanan_transaksi if t.get('id_anggota') == id_anggota]
        anggota = [a for a in anggota if a.get('id_anggota') == id_anggota]

    simpanan_tampil = enrich_simpanan_untuk_tampilan(simpanan, anggota)
    simpanan_tampil.reverse()

    saldo_per_jenis_map = {}
    for t in simpanan_transaksi:
        id_a = (t.get('id_anggota') or '').strip()
        if not id_a:
            continue
        jenis = (t.get('jenis_simpanan') or 'Manasuka').strip() or 'Manasuka'
        if jenis not in JENIS_SIMPANAN:
            jenis = 'Manasuka'
        try:
            nominal = float(t.get('jumlah') or 0)
        except (TypeError, ValueError):
            nominal = 0.0
        if id_a not in saldo_per_jenis_map:
            saldo_per_jenis_map[id_a] = {k: 0.0 for k in JENIS_SIMPANAN}
        saldo_per_jenis_map[id_a][jenis] += nominal

    for s in simpanan_tampil:
        id_a = s.get('id_anggota', '')
        by_jenis = saldo_per_jenis_map.get(id_a, {k: 0.0 for k in JENIS_SIMPANAN})
        saldo_manasuka = float(by_jenis.get('Manasuka', 0.0))
        saldo_hari_raya = float(by_jenis.get('Hari Raya', 0.0))
        saldo_pendidikan = float(by_jenis.get('Pendidikan', 0.0))
        # Jika ada data lama (total tanpa rincian jenis), alokasikan selisih ke Manasuka.
        try:
            total_saldo = float(s.get('jumlah') or 0)
        except (TypeError, ValueError):
            total_saldo = 0.0
        subtotal_jenis = saldo_manasuka + saldo_hari_raya + saldo_pendidikan
        if total_saldo > subtotal_jenis:
            saldo_manasuka += (total_saldo - subtotal_jenis)
        s['saldo_manasuka'] = round(saldo_manasuka, 2)
        s['saldo_hari_raya'] = round(saldo_hari_raya, 2)
        s['saldo_pendidikan'] = round(saldo_pendidikan, 2)

    saldo_anggota = {}
    for a in anggota:
        saldo_anggota[a['id_anggota']] = 0.0
    for s in simpanan:
        if s['id_anggota'] in saldo_anggota:
            saldo_anggota[s['id_anggota']] += float(s.get('total_simpanan') or 0)

    transaksi_tampil = sorted(simpanan_transaksi, key=lambda x: x.get('tanggal', ''), reverse=True)

    return render_template(
        'simpanan.html',
        simpanan=simpanan_tampil,
        simpanan_transaksi=transaksi_tampil,
        anggota=anggota,
        saldo_anggota=saldo_anggota,
        jenis_simpanan_choices=JENIS_SIMPANAN,
    )


@app.route('/simpanan/tambah', methods=['POST'])
@login_required
def tambah_simpanan():
    ensure_simpanan_schema()
    ensure_simpanan_transaksi_schema()
    simpanan = baca_csv(FILE_SIMPANAN)
    simpanan_transaksi = baca_csv(FILE_SIMPANAN_TRANSAKSI)
    if is_current_user_admin():
        id_anggota = request.form['id_anggota']
    else:
        id_anggota = get_current_user_id_anggota()
        if not id_anggota:
            abort(403)
    anggota = baca_csv(FILE_ANGGOTA)
    anggota_data = next((a for a in anggota if a['id_anggota'] == id_anggota), None)

    if not anggota_data:
        return "Anggota tidak ditemukan", 400

    try:
        jumlah = float(request.form['jumlah'].replace(',', '').replace('.', ''))
    except ValueError:
        flash('Jumlah simpanan tidak valid.', 'danger')
        return redirect('/simpanan')
    if jumlah < 0:
        flash('Jumlah simpanan tidak boleh negatif.', 'danger')
        return redirect('/simpanan')

    merge_akumulasi(simpanan, id_anggota, jumlah)
    tulis_csv(FILE_SIMPANAN, simpanan, SIMPANAN_FIELDNAMES)

    jenis_simpanan = (request.form.get('jenis_simpanan') or 'Manasuka').strip() or 'Manasuka'
    if jenis_simpanan not in JENIS_SIMPANAN:
        jenis_simpanan = 'Manasuka'

    simpanan_transaksi.append({
        'id_transaksi': str(uuid.uuid4()),
        'id_anggota': id_anggota,
        'no_anggota': anggota_data.get('no_anggota', ''),
        'nama_anggota': anggota_data.get('nama', ''),
        'tanggal': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'jenis_simpanan': jenis_simpanan,
        'jumlah': str(round(jumlah, 2)),
        'keterangan': 'Simpanan',
        'diajukan_oleh': session.get('user') or '',
    })
    tulis_csv(FILE_SIMPANAN_TRANSAKSI, simpanan_transaksi, SIMPANAN_TRANSAKSI_FIELDNAMES)
    flash('Simpanan berhasil ditambahkan ke saldo.', 'success')
    return redirect('/simpanan')


@app.route('/simpanan/konfirmasi/<id_simpanan>')
@admin_required
def konfirmasi_simpanan(id_simpanan):
    flash('Skema simpanan saldo tidak memerlukan konfirmasi baris.', 'info')
    return redirect('/simpanan')


@app.route('/simpanan/tolak/<id_simpanan>')
@admin_required
def tolak_simpanan(id_simpanan):
    flash('Skema simpanan saldo tidak memerlukan penolakan baris.', 'info')
    return redirect('/simpanan')


@app.route('/simpanan/hapus/<id_anggota>')
@admin_required
def hapus_simpanan(id_anggota):
    data = [s for s in baca_csv(FILE_SIMPANAN) if s.get('id_anggota') != id_anggota]
    tulis_csv(FILE_SIMPANAN, data, SIMPANAN_FIELDNAMES)
    flash('Saldo simpanan anggota dihapus dari daftar.', 'success')
    return redirect('/simpanan')


@app.route('/simpanan/saldo/<id_anggota>')
@login_required
def cek_saldo(id_anggota):
    """Cek total saldo simpanan anggota."""
    restrict_id_anggota_or_forbid(id_anggota)
    ensure_simpanan_schema()
    ensure_simpanan_transaksi_schema()
    simpanan = baca_csv(FILE_SIMPANAN)
    simpanan_transaksi = baca_csv(FILE_SIMPANAN_TRANSAKSI)
    total = 0.0
    for s in simpanan:
        if s.get('id_anggota') == id_anggota:
            total = float(s.get('total_simpanan') or 0)
            break
    per_jenis = {k: 0.0 for k in JENIS_SIMPANAN}
    for t in simpanan_transaksi:
        if t.get('id_anggota') != id_anggota:
            continue
        jenis = (t.get('jenis_simpanan') or 'Manasuka').strip() or 'Manasuka'
        if jenis not in per_jenis:
            jenis = 'Manasuka'
        try:
            nominal = float(t.get('jumlah') or 0)
        except (TypeError, ValueError):
            nominal = 0.0
        per_jenis[jenis] += nominal
    subtotal = sum(per_jenis.values())
    if total > subtotal:
        per_jenis['Manasuka'] += (total - subtotal)

    jenis_req = (request.args.get('jenis_simpanan') or 'Manasuka').strip() or 'Manasuka'
    if jenis_req not in per_jenis:
        jenis_req = 'Manasuka'

    return jsonify({
        'saldo': total,
        'saldo_per_jenis': {k: round(v, 2) for k, v in per_jenis.items()},
        'jenis_simpanan': jenis_req,
        'saldo_jenis': round(per_jenis.get(jenis_req, 0.0), 2),
    })


# ══════════════════════════════════════════════
#  ROUTE: PINJAMAN
# ══════════════════════════════════════════════
@app.route('/pinjaman')
@login_required
def halaman_pinjaman():
    ensure_pinjaman_plafon_schema()
    pinjaman = baca_csv(FILE_PINJAMAN)
    anggota_full = baca_csv(FILE_ANGGOTA)
    cicilan_menunggu = []
    if not is_current_user_admin():
        id_anggota = get_current_user_id_anggota()
        pinjaman = [p for p in pinjaman if p.get('id_anggota') == id_anggota]
        anggota = [a for a in anggota_full if a.get('id_anggota') == id_anggota]
    else:
        anggota = anggota_full
        cicilan = baca_csv(FILE_PINJAMAN_CICILAN)
        cicilan_menunggu = [c for c in cicilan if c.get('status') == 'Menunggu']

    riwayat_map = {}
    for p in pinjaman:
        id_a = p.get('id_anggota', '')
        if not id_a:
            continue
        tenor_raw = p.get('tenor_bulan', p.get('tenor', '0'))
        try:
            tenor_norm = int(float(tenor_raw or 0))
        except (TypeError, ValueError):
            tenor_norm = 0
        jenis = (p.get('jenis_pinjaman') or '').strip()
        if not jenis or jenis == JENIS_IMPORT_CSV:
            jenis = kategori_pinjaman_dari_tenor(tenor_norm)
        if not jenis:
            continue

        row = riwayat_map.setdefault(id_a, {
            'id_anggota': id_a,
            'no_anggota': p.get('no_anggota', ''),
            'nama_anggota': p.get('nama_anggota', ''),
            'total_pengajuan': 0,
            'total_aktif': 0,
            'jenis_list': [],
        })

        if not row['no_anggota']:
            row['no_anggota'] = p.get('no_anggota', '')
        if not row['nama_anggota']:
            row['nama_anggota'] = p.get('nama_anggota', '')

        row['total_pengajuan'] += 1
        if (p.get('status') or '').strip() == 'Disetujui' and saldo_pinjaman_aktual(p) > 0:
            row['total_aktif'] += 1
        if jenis not in row['jenis_list']:
            row['jenis_list'].append(jenis)

    anggota_map = {a.get('id_anggota'): a for a in anggota_full}
    riwayat_pinjaman_anggota = []
    for item in riwayat_map.values():
        a = anggota_map.get(item['id_anggota'], {})
        no_anggota = item['no_anggota'] or a.get('no_anggota', '-')
        nama_anggota = item['nama_anggota'] or a.get('nama', '-')
        jenis_items = item['jenis_list']
        riwayat_pinjaman_anggota.append({
            'id_anggota': item['id_anggota'],
            'no_anggota': no_anggota,
            'nama_anggota': nama_anggota,
            'total_pengajuan': item['total_pengajuan'],
            'total_aktif': item['total_aktif'],
            'jenis_items': jenis_items,
            'jenis_text': ', '.join(jenis_items),
        })

    riwayat_pinjaman_anggota.sort(key=lambda x: (x.get('nama_anggota') or '').lower())

    pinjaman_tampil = enrich_pinjaman_untuk_tampilan(pinjaman, anggota_full)
    pinjaman_tampil.reverse()
    return render_template(
        'pinjaman.html',
        pinjaman=pinjaman_tampil,
        riwayat_pinjaman_anggota=riwayat_pinjaman_anggota,
        anggota=anggota,
        cicilan_menunggu=cicilan_menunggu,
        metode_bayar_choices=METODE_BAYAR_CHOICES,
        jenis_pinjaman_choices=[j for j in JENIS_PINJAMAN_CHOICES if j in JENIS_PINJAMAN],
    )


@app.route('/pinjaman/hitung', methods=['GET'])
def hitung_pinjaman():
    """Hitung estimasi pinjaman berdasarkan plafon + tenor fleksibel."""
    try:
        raw_plafon = request.args.get('plafon') or request.args.get('Pinjaman', 0)
        raw_tenor = request.args.get('tenor')
        id_anggota = request.args.get('id_anggota')
        jenis_pm = (request.args.get('jenis_pinjaman') or '').strip()
        plafon = float(raw_plafon)
        if not raw_tenor:
            return jsonify({'error': 'Tenor wajib diisi'})
        tenor = int(raw_tenor)
        if tenor <= 0:
            return jsonify({'error': 'Tenor harus lebih dari 0'}), 400
        bunga_persen = bunga_untuk_jenis_pinjaman(jenis_pm, tenor)
        total_bayar = hitung_total_bayar_tanpa_provisi(plafon, bunga_persen, tenor, jenis_pm)
        provisi_nominal = provisi_nominal_pinjaman(jenis_pm, plafon, tenor)
        provisi_persen = provisi_persen_dari_pinjaman(jenis_pm, tenor)
        # Hitung cicilan bulanan: cicilan = (p / b) + (p × j%)
        cicilan = hitung_cicilan_bulanan(plafon, bunga_persen, tenor)

        kapasitas_cicilan = None
        plafon_maks = None
        dsr = None
        dsr_persen = None
        if id_anggota:
            ensure_anggota_schema()
            anggota = get_anggota_by_id(id_anggota)
            if anggota:
                penghasilan = float(anggota.get('penghasilan_bersih') or 0)
                cicilan_lain = float(anggota.get('cicilan_lain') or 0)
                dsr = dsr_otomatis_dari_penghasilan(penghasilan)
                dsr_persen = round(dsr * 100)
                kapasitas_cicilan = hitung_kapasitas_cicilan(penghasilan, cicilan_lain, dsr)
                plafon_maks = hitung_plafon_maks_anuitas(kapasitas_cicilan, bunga_persen, tenor, jenis_pm)
                if jenis_pm == 'Solusi Cepat':
                    plafon_maks = min(plafon_maks, 3_000_000.0)
        return jsonify({
            'bunga_persen': bunga_persen,
            'tenor': tenor,
            'total_bayar': round(total_bayar, 2),
            'cicilan_per_bulan': round(cicilan, 2),
            'provisi_nominal': round(provisi_nominal, 2),
            'provisi_persen': round(provisi_persen, 2),
            'kapasitas_cicilan': round(kapasitas_cicilan, 2) if kapasitas_cicilan is not None else None,
            'plafon_maks': round(plafon_maks, 2) if plafon_maks is not None else None,
            'dsr_persen': dsr_persen,
        })
    except ValueError as e:
        return jsonify({'error': str(e)}), 400


@app.route('/pinjaman/cek_tunggakan/<id_anggota>')
@login_required
def cek_tunggakan(id_anggota):
    """Cek apakah anggota masih punya tunggakan."""
    restrict_id_anggota_or_forbid(id_anggota)
    pinjaman = baca_csv(FILE_PINJAMAN)
    tunggakan = [
        p for p in pinjaman
        if p.get('id_anggota') == id_anggota
        and p.get('status') == 'Disetujui'
        and saldo_pinjaman_aktual(p) > 0
    ]
    ada_solusi_cepat_aktif = ada_pinjaman_solusi_cepat_aktif(pinjaman, id_anggota)
    if tunggakan:
        return jsonify({
            'ada_tunggakan': True,
            'pinjaman_aktif': len(tunggakan),
            'ada_solusi_cepat_aktif': ada_solusi_cepat_aktif,
        })
    return jsonify({'ada_tunggakan': False, 'pinjaman_aktif': 0, 'ada_solusi_cepat_aktif': ada_solusi_cepat_aktif})


@app.route('/pinjaman/tambah', methods=['POST'])
@login_required
def tambah_pinjaman():
    pinjaman = baca_csv(FILE_PINJAMAN)
    if is_current_user_admin():
        id_anggota = request.form['id_anggota']
    else:
        id_anggota = get_current_user_id_anggota()
        if not id_anggota:
            abort(403)
    ensure_anggota_schema()
    anggota = baca_csv(FILE_ANGGOTA)
    anggota_data = next((a for a in anggota if a['id_anggota'] == id_anggota), None)

    if not anggota_data:
        return "Anggota tidak ditemukan", 400

    if not is_current_user_admin():
        pinjaman_telat = [
            p for p in pinjaman
            if p.get('id_anggota') == id_anggota
            and p.get('status') == 'Disetujui'
            and pinjaman_telat_bayar_aktual(p)
        ]
        if pinjaman_telat:
            flash('Pengajuan ditolak: masih ada pinjaman yang telat bayar melewati tanggal jatuh tempo.', 'danger')
            return redirect('/pinjaman')

    jenis = (request.form.get('jenis_pinjaman') or 'Jangka Pendek').strip()
    try:
        plafon = float(request.form['plafon'].replace(',', '').replace('.', ''))
        tenor = int(request.form['tenor_bulan'])
        bunga_persen = bunga_untuk_jenis_pinjaman(jenis, tenor)
    except ValueError as e:
        flash(str(e), 'danger')
        return redirect('/pinjaman')

    if jenis == 'Solusi Cepat' and ada_pinjaman_solusi_cepat_aktif(pinjaman, id_anggota):
        flash('Pengajuan Solusi Cepat ditolak: Solusi Cepat sebelumnya belum lunas. Jenis pinjaman lain tetap bisa diajukan.', 'danger')
        return redirect('/pinjaman')

    jenis_simpanan = (request.form.get('jenis_simpanan') or 'Manasuka').strip() or 'Manasuka'
    if jenis_simpanan not in JENIS_SIMPANAN:
        jenis_simpanan = 'Manasuka'

    if jenis == 'Solusi Cepat' and plafon > 3_000_000:
        flash('Pengajuan ditolak: plafon Solusi Cepat maksimal Rp 3.000.000.', 'danger')
        return redirect('/pinjaman')

    # Hitung total_bayar TANPA provisi untuk pengajuan awal
    total_bayar = hitung_total_bayar_tanpa_provisi(plafon, bunga_persen, tenor, jenis)
    cicilan = hitung_cicilan_bulanan(plafon, bunga_persen, tenor)
    
    # Hitung provisi untuk informasi user (khusus Jangka Panjang, dipotong saat pencairan)
    provisi_nominal = provisi_nominal_pinjaman(jenis, plafon, tenor)

    penghasilan = float(anggota_data.get('penghasilan_bersih') or 0)
    cicilan_lain = float(anggota_data.get('cicilan_lain') or 0)
    dsr = dsr_otomatis_dari_penghasilan(penghasilan)
    kapasitas_cicilan = hitung_kapasitas_cicilan(penghasilan, cicilan_lain, dsr)
    plafon_maks = hitung_plafon_maks_anuitas(kapasitas_cicilan, bunga_persen, tenor, jenis)
    if jenis == 'Solusi Cepat':
        plafon_maks = min(plafon_maks, 3_000_000.0)
    if kapasitas_cicilan <= 0:
        flash('Pengajuan ditolak: kapasitas cicilan anggota tidak mencukupi.', 'danger')
        return redirect('/pinjaman')
    if cicilan > kapasitas_cicilan:
        flash(
            f"Pengajuan ditolak: cicilan bulanan Rp {cicilan:,.0f} melebihi kapasitas Rp {kapasitas_cicilan:,.0f}.",
            'danger'
        )
        return redirect('/pinjaman')
    if plafon > plafon_maks:
        flash(
            f"Pengajuan melebihi plafon maksimal anggota. Maksimal Rp {plafon_maks:,.0f} "
            f"(kapasitas cicilan Rp {kapasitas_cicilan:,.0f}/bln).",
            'danger'
        )
        return redirect('/pinjaman')

    tgl = datetime.now().strftime('%Y-%m-%d')
    pid = str(uuid.uuid4())
    pinjaman.append({
        'id_pinjaman': pid,
        'id_anggota': id_anggota,
        'nama_anggota': anggota_data.get('nama', ''),
        'no_anggota': anggota_data.get('no_anggota', ''),
        'jenis_pinjaman': jenis,
        'jenis_simpanan': jenis_simpanan,
        'plafon': str(round(plafon, 2)),
        'tenor_awal': str(tenor),
        'tenor_bulan': str(tenor),
        'bunga_persen': str(bunga_persen),
        'total_bayar': str(round(total_bayar, 2)),
        'cicilan_per_bulan': str(round(cicilan, 2)),
        'sisa_pinjaman': str(round(plafon, 2)),
        'tanggal_pengajuan': tgl,
        'status': 'Menunggu',
        'tanggal_lunas': '',
    })
    tulis_csv(FILE_PINJAMAN, pinjaman, PINJAMAN_FIELDNAMES)
    provisi_info = (
        f' Provisi Rp {provisi_nominal:,.0f} akan dipotong dari dana pencairan saat disetujui admin.'
        if provisi_nominal > 0 else ''
    )
    flash(f'Pengajuan pinjaman tercatat dan menunggu konfirmasi admin.{provisi_info}', 'success')
    return redirect('/pinjaman')


@app.route('/pinjaman/konfirmasi/<id_pinjaman>')
@admin_required
def konfirmasi_pinjaman(id_pinjaman):
    pinjaman = baca_csv(FILE_PINJAMAN)
    found = False
    provisi_nominal = 0.0
    dana_cair = 0.0
    for i, p in enumerate(pinjaman):
        if p.get('id_pinjaman') == id_pinjaman and p.get('status') == 'Menunggu':
            # Ubah status menjadi Disetujui
            p['status'] = 'Disetujui'
            
            # Provisi dipotong dari dana cair, bukan ditambahkan ke cicilan.
            try:
                jenis_pinjaman = (p.get('jenis_pinjaman') or '').strip()
                plafon = float(p.get('plafon') or 0)
                tenor = int(float(p.get('tenor_bulan') or 0))
                bunga_persen = float(p.get('bunga_persen') or 0)
                
                # Kewajiban cicilan tetap tanpa provisi (provisi dipotong saat pencairan)
                total_bayar = hitung_total_bayar_tanpa_provisi(plafon, bunga_persen, tenor, jenis_pinjaman)
                cicilan = hitung_cicilan_bulanan(plafon, bunga_persen, tenor)
                provisi_nominal = provisi_nominal_pinjaman(jenis_pinjaman, plafon, tenor)
                dana_cair = max(plafon - provisi_nominal, 0.0)
                
                p['total_bayar'] = str(round(total_bayar, 2))
                p['cicilan_per_bulan'] = str(round(cicilan, 2))
                p['tenor_awal'] = str(tenor)
                p['sisa_pinjaman'] = str(round(plafon, 2))
            except (TypeError, ValueError):
                # Jika gagal hitung, gunakan nilai existing
                plaf = float(p.get('plafon') or 0)
                p['tenor_awal'] = str(int(float(p.get('tenor_bulan') or 0)))
                p['sisa_pinjaman'] = str(round(plaf, 2))
            
            found = True
            break
    
    if not found:
        flash('Pengajuan tidak ditemukan atau sudah diproses.', 'warning')
        return redirect('/pinjaman')
    
    tulis_csv(FILE_PINJAMAN, pinjaman, PINJAMAN_FIELDNAMES)
    
    # Cari id_anggota dari pinjaman yang baru dikonfirmasi
    id_anggota = None
    for p in pinjaman:
        if p.get('id_pinjaman') == id_pinjaman:
            id_anggota = p.get('id_anggota')
            break
    
    # Merge dengan pinjaman approved lain dari anggota sama (jika ada)
    if id_anggota:
        merge_pinjaman_sama_anggota(pinjaman, id_pinjaman, id_anggota)
        tulis_csv(FILE_PINJAMAN, pinjaman, PINJAMAN_FIELDNAMES)
    
    if provisi_nominal > 0:
        flash(
            f'Pinjaman disetujui. Provisi Rp {provisi_nominal:,.0f} dipotong saat pencairan. '
            f'Dana cair bersih: Rp {dana_cair:,.0f}.',
            'success'
        )
    else:
        flash('Pinjaman disetujui.', 'success')
    return redirect('/pinjaman')


@app.route('/pinjaman/tolak/<id_pinjaman>')
@admin_required
def tolak_pinjaman(id_pinjaman):
    pinjaman = baca_csv(FILE_PINJAMAN)
    found = False
    for p in pinjaman:
        if p.get('id_pinjaman') == id_pinjaman and p.get('status') == 'Menunggu':
            p['status'] = 'Ditolak'
            found = True
            break
    if not found:
        flash('Pengajuan tidak ditemukan atau sudah diproses.', 'warning')
        return redirect('/pinjaman')
    tulis_csv(FILE_PINJAMAN, pinjaman, PINJAMAN_FIELDNAMES)
    flash('Pengajuan pinjaman ditolak.', 'warning')
    return redirect('/pinjaman')


def _proses_angsur_pinjaman(id_pinjaman: str):
    """Kurangi sisa kewajiban satu kali cicilan dan turunkan tenor tersisa 1 bulan."""
    pinjaman = baca_csv(FILE_PINJAMAN)
    for p in pinjaman:
        if p.get('id_pinjaman') != id_pinjaman:
            continue
        if p.get('status') != 'Disetujui':
            return
        sisa = saldo_pinjaman_aktual(p)
        cic = nominal_cicilan_aktual(p)
        baru = max(sisa - cic, 0)
        p['sisa_pinjaman'] = str(round(baru, 2))
        try:
            tenor_now = int(float(p.get('tenor_bulan') or p.get('tenor_awal') or 0))
        except (TypeError, ValueError):
            tenor_now = 0
        p['tenor_bulan'] = str(max(tenor_now - 1, 0))
        if baru <= 0:
            p['status'] = 'Lunas'
            p['tanggal_lunas'] = datetime.now().strftime('%Y-%m-%d')
            p['tenor_bulan'] = '0'
        break
    tulis_csv(FILE_PINJAMAN, pinjaman, PINJAMAN_FIELDNAMES)


def _pinjaman_row_dengan_nama(id_pinjaman: str):
    pinjaman = baca_csv(FILE_PINJAMAN)
    target = next((p for p in pinjaman if p.get('id_pinjaman') == id_pinjaman), None)
    if not target:
        return None, None
    ag = get_anggota_by_id(target.get('id_anggota', '')) or {}
    cic = float(target.get('cicilan_per_bulan') or 0) or cicilan_per_bulan_saldo(target)
    sisa = saldo_pinjaman_aktual(target)
    return target, {
        'no_anggota': target.get('no_anggota') or ag.get('no_anggota', ''),
        'nama_anggota': target.get('nama_anggota') or ag.get('nama', ''),
        'cicilan_per_bulan': str(round(cic, 2)),
        'saldo_pinjaman': str(round(sisa, 2)),
    }


@app.route('/pinjaman/angsur/<id_pinjaman>', methods=['POST'])
@admin_required
def angsur_pinjaman(id_pinjaman):
    """Bayar cicilan langsung oleh admin; mencatat metode Admin di riwayat cicilan."""
    ensure_pinjaman_cicilan_schema()
    ket = (request.form.get('keterangan') or '').strip() or 'Pembayaran cicilan melalui admin (kas koperasi)'
    target, meta = _pinjaman_row_dengan_nama(id_pinjaman)
    if not target:
        flash('Data pinjaman tidak ditemukan.', 'danger')
        return redirect('/pinjaman')
    if target.get('status') != 'Disetujui' or saldo_pinjaman_aktual(target) <= 0:
        flash('Pinjaman ini tidak memiliki cicilan yang dapat dibayar.', 'warning')
        return redirect('/pinjaman')

    id_anggota = target.get('id_anggota', '')
    jumlah_bayar = nominal_cicilan_aktual(target)
    cicilan = baca_csv(FILE_PINJAMAN_CICILAN)
    cicilan.append({
        'id_cicilan': str(uuid.uuid4()),
        'id_pinjaman': id_pinjaman,
        'id_anggota': id_anggota,
        'no_anggota': meta['no_anggota'],
        'nama_anggota': meta['nama_anggota'],
        'jumlah': str(round(jumlah_bayar, 2)),
        'tanggal_pengajuan': datetime.now().strftime('%Y-%m-%d'),
        'status': 'Disetujui',
        'tanggal_konfirmasi': datetime.now().strftime('%Y-%m-%d'),
        'dikonfirmasi_oleh': session.get('user') or '',
        'diajukan_oleh': session.get('user') or '',
        'keterangan': ket,
        'metode_pembayaran': 'Admin',
        'detail_pembayaran': '',
    })
    tulis_csv(FILE_PINJAMAN_CICILAN, cicilan, CICILAN_FIELDNAMES)
    _proses_angsur_pinjaman(id_pinjaman)
    flash('Cicilan dicatat (bayar melalui admin). Saldo pinjaman diperbarui.', 'success')
    return redirect('/pinjaman')


@app.route('/pinjaman/cicilan/gagal-bayar/<id_pinjaman>', methods=['POST'])
@admin_required
def gagal_bayar_pinjaman(id_pinjaman):
    """Catat cicilan gagal bayar dan akumulasikan ke sisa pinjaman."""
    ensure_pinjaman_cicilan_schema()
    ket = (request.form.get('keterangan') or '').strip() or 'Gagal bayar cicilan (terlambat / tidak tepat waktu)'
    target, meta = _pinjaman_row_dengan_nama(id_pinjaman)
    if not target:
        flash('Data pinjaman tidak ditemukan.', 'danger')
        return redirect('/pinjaman')
    if target.get('status') != 'Disetujui' or saldo_pinjaman_aktual(target) <= 0:
        flash('Pinjaman ini tidak memiliki cicilan aktif yang bisa ditandai gagal bayar.', 'warning')
        return redirect('/pinjaman')

    jumlah_bayar = nominal_cicilan_aktual(target)
    cicilan = baca_csv(FILE_PINJAMAN_CICILAN)
    cicilan.append({
        'id_cicilan': str(uuid.uuid4()),
        'id_pinjaman': id_pinjaman,
        'id_anggota': target.get('id_anggota', ''),
        'no_anggota': meta['no_anggota'],
        'nama_anggota': meta['nama_anggota'],
        'jumlah': str(round(jumlah_bayar, 2)),
        'tanggal_pengajuan': datetime.now().strftime('%Y-%m-%d'),
        'status': 'Gagal Bayar',
        'tanggal_konfirmasi': datetime.now().strftime('%Y-%m-%d'),
        'dikonfirmasi_oleh': session.get('user') or '',
        'diajukan_oleh': session.get('user') or '',
        'keterangan': ket,
        'metode_pembayaran': 'Admin',
        'detail_pembayaran': '',
    })
    tulis_csv(FILE_PINJAMAN_CICILAN, cicilan, CICILAN_FIELDNAMES)

    # Akumulasi tunggakan: cicilan gagal bayar ditambahkan ke sisa pinjaman.
    pinjaman_rows = baca_csv(FILE_PINJAMAN)
    for p in pinjaman_rows:
        if p.get('id_pinjaman') != id_pinjaman:
            continue
        sisa_sekarang = saldo_pinjaman_aktual(p)
        sisa_baru = max(sisa_sekarang + jumlah_bayar, 0.0)
        p['sisa_pinjaman'] = str(round(sisa_baru, 2))
        if p.get('status') == 'Lunas':
            p['status'] = 'Disetujui'
            p['tanggal_lunas'] = ''
        break
    tulis_csv(FILE_PINJAMAN, pinjaman_rows, PINJAMAN_FIELDNAMES)

    flash('Cicilan ditandai gagal bayar. Sisa cicilan telah diakumulasi.', 'warning')
    return redirect('/pinjaman')


@app.route('/pinjaman/ajukan-cicilan/<id_pinjaman>', methods=['POST'])
@login_required
def ajukan_cicilan(id_pinjaman):
    """Pengajuan bayar cicilan oleh user; perlu dikonfirmasi admin sebelum mengurangi sisa pinjaman."""
    ensure_pinjaman_cicilan_schema()
    target, meta = _pinjaman_row_dengan_nama(id_pinjaman)
    if not target:
        flash('Data pinjaman tidak ditemukan.', 'danger')
        return redirect('/pinjaman')

    id_anggota = target.get('id_anggota', '')
    if not is_current_user_admin():
        id_anggota_user = get_current_user_id_anggota()
        if not id_anggota_user or id_anggota != id_anggota_user:
            abort(403)

    if target.get('status') != 'Disetujui' or saldo_pinjaman_aktual(target) <= 0:
        flash('Pinjaman ini tidak memiliki cicilan yang perlu dibayar.', 'warning')
        return redirect('/pinjaman')

    if (target.get('jenis_pinjaman') or '').strip() == 'Solusi Cepat':
        flash('Solusi Cepat hanya bisa dibayar melalui admin.', 'warning')
        return redirect('/pinjaman')

    metode = (request.form.get('metode_pembayaran') or 'Lainnya').strip()
    if metode not in METODE_BAYAR_CHOICES:
        metode = 'Lainnya'
    detail = (request.form.get('detail_pembayaran') or '').strip()
    ket_parts = [f'Pengajuan pembayaran via {metode}']
    if detail:
        ket_parts.append(detail)
    keterangan = ' — '.join(ket_parts)

    cicilan = baca_csv(FILE_PINJAMAN_CICILAN)
    jumlah_bayar = nominal_cicilan_aktual(target)
    cicilan.append({
        'id_cicilan': str(uuid.uuid4()),
        'id_pinjaman': id_pinjaman,
        'id_anggota': id_anggota,
        'no_anggota': meta['no_anggota'],
        'nama_anggota': meta['nama_anggota'],
        'jumlah': str(round(jumlah_bayar, 2)),
        'tanggal_pengajuan': datetime.now().strftime('%Y-%m-%d'),
        'status': 'Menunggu',
        'tanggal_konfirmasi': '',
        'dikonfirmasi_oleh': '',
        'diajukan_oleh': session.get('user') or '',
        'keterangan': keterangan,
        'metode_pembayaran': metode,
        'detail_pembayaran': detail,
    })
    tulis_csv(FILE_PINJAMAN_CICILAN, cicilan, CICILAN_FIELDNAMES)
    flash('Pengajuan bayar cicilan berhasil dikirim. Menunggu konfirmasi admin.', 'success')
    return redirect('/pinjaman')


@app.route('/pinjaman/cicilan/konfirmasi/<id_cicilan>')
@admin_required
def konfirmasi_cicilan(id_cicilan):
    """Admin menyetujui pengajuan cicilan dan langsung mengurangi sisa pinjaman."""
    cicilan = baca_csv(FILE_PINJAMAN_CICILAN)
    target = None
    for c in cicilan:
        if c.get('id_cicilan') == id_cicilan:
            target = c
            break
    if not target:
        flash('Data pengajuan cicilan tidak ditemukan.', 'danger')
        return redirect('/pinjaman')
    if target.get('status') != 'Menunggu':
        flash('Pengajuan cicilan ini sudah diproses sebelumnya.', 'warning')
        return redirect('/pinjaman')

    pinjaman, _ = _pinjaman_row_dengan_nama(target['id_pinjaman'])
    if not pinjaman:
        flash('Data pinjaman tidak ditemukan.', 'danger')
        return redirect('/pinjaman')

    jumlah_bayar = nominal_cicilan_aktual(pinjaman)

    target['status'] = 'Disetujui'
    target['tanggal_konfirmasi'] = datetime.now().strftime('%Y-%m-%d')
    target['dikonfirmasi_oleh'] = session.get('user') or ''
    target['jumlah'] = str(round(jumlah_bayar, 2))
    tulis_csv(FILE_PINJAMAN_CICILAN, cicilan, CICILAN_FIELDNAMES)

    _proses_angsur_pinjaman(target['id_pinjaman'])
    flash('Pengajuan bayar cicilan diterima dan sisa pinjaman telah diperbarui.', 'success')
    return redirect('/pinjaman')


@app.route('/pinjaman/cicilan/menunggu', methods=['GET'])
@admin_required
def cicilan_menunggu_json():
    """Data pengajuan cicilan menunggu untuk refresh otomatis di halaman pinjaman."""
    ensure_pinjaman_cicilan_schema()
    rows = [c for c in baca_csv(FILE_PINJAMAN_CICILAN) if (c.get('status') or '').strip() == 'Menunggu']
    rows.sort(key=lambda x: x.get('tanggal_pengajuan', ''), reverse=True)

    items = []
    for c in rows:
        try:
            jumlah = float(c.get('jumlah') or 0)
        except (TypeError, ValueError):
            jumlah = 0.0
        items.append({
            'id_cicilan': c.get('id_cicilan', ''),
            'tanggal_pengajuan': c.get('tanggal_pengajuan', '-'),
            'no_anggota': c.get('no_anggota', '-') or '-',
            'nama_anggota': c.get('nama_anggota', '-') or '-',
            'jumlah': round(jumlah, 2),
            'metode_pembayaran': c.get('metode_pembayaran', '-') or '-',
            'keterangan': c.get('keterangan', '-') or '-',
        })

    return jsonify({'count': len(items), 'items': items})


@app.route('/pinjaman/cicilan/tolak/<id_cicilan>')
@admin_required
def tolak_cicilan(id_cicilan):
    """Admin menolak pengajuan cicilan (tanpa mengubah sisa pinjaman)."""
    cicilan = baca_csv(FILE_PINJAMAN_CICILAN)
    found = False
    for c in cicilan:
        if c.get('id_cicilan') == id_cicilan:
            c['status'] = 'Ditolak'
            c['tanggal_konfirmasi'] = datetime.now().strftime('%Y-%m-%d')
            c['dikonfirmasi_oleh'] = session.get('user') or ''
            found = True
            break
    if not found:
        flash('Data pengajuan cicilan tidak ditemukan.', 'danger')
        return redirect('/pinjaman')
    tulis_csv(FILE_PINJAMAN_CICILAN, cicilan, CICILAN_FIELDNAMES)
    flash('Pengajuan bayar cicilan ditolak.', 'warning')
    return redirect('/pinjaman')


@app.route('/pinjaman/hapus/<id_pinjaman>')
@admin_required
def hapus_pinjaman(id_pinjaman):
    data = [p for p in baca_csv(FILE_PINJAMAN) if p.get('id_pinjaman') != id_pinjaman]
    tulis_csv(FILE_PINJAMAN, data, PINJAMAN_FIELDNAMES)
    flash('Data pinjaman dihapus.', 'success')
    return redirect('/pinjaman')


# ══════════════════════════════════════════════
#  ROUTE: LAPORAN & EXPORT
# ══════════════════════════════════════════════

def upsert_anggota_from_riwayat(pinjaman_rows: list, simpanan_rows: list) -> dict:
    """Sinkronkan anggota dari baris saldo pinjaman/simpanan (berbasis id_anggota)."""
    ensure_anggota_schema()
    anggota = baca_csv(FILE_ANGGOTA)
    by_id = {a.get('id_anggota'): a for a in anggota}

    added = 0

    def ensure_id(id_a: str):
        nonlocal added
        id_a = (id_a or '').strip()
        if not id_a:
            return
        if id_a in by_id:
            return
        no_baru = generate_no_anggota_berikutnya(anggota)
        anggota.append({
            'id_anggota': id_a,
            'no_anggota': no_baru,
            'nik': '',
            'nama': '',
            'alamat': '',
            'no_telp': '',
            'tgl_bergabung': datetime.now().strftime('%Y-%m-%d'),
            'penghasilan_bersih': '0',
            'cicilan_lain': '0',
        })
        by_id[id_a] = anggota[-1]
        added += 1

    for p in pinjaman_rows:
        ensure_id(p.get('id_anggota'))
    for s in simpanan_rows:
        ensure_id(s.get('id_anggota'))

    tulis_csv(FILE_ANGGOTA, anggota, ANGGOTA_FIELDNAMES)
    return {'added': added, 'updated': 0}


def hitung_shu_pinjaman() -> dict:
    """SHU perkiraan: proporsi dari nominal cicilan yang sudah disetujui (saldo pokok)."""
    ensure_pinjaman_cicilan_schema()
    cicilan_rows = [c for c in baca_csv(FILE_PINJAMAN_CICILAN) if c.get('status') == 'Disetujui']
    anggota_map = {a.get('id_anggota'): a for a in baca_csv(FILE_ANGGOTA)}

    bunga_per_anggota = {}
    total_bunga_dibayar = 0.0
    for c in cicilan_rows:
        id_anggota = (c.get('id_anggota') or '').strip()
        j = float(c.get('jumlah') or 0)
        if j <= 0:
            continue
        total_bunga_dibayar += j
        ag = anggota_map.get(id_anggota, {})
        if id_anggota not in bunga_per_anggota:
            bunga_per_anggota[id_anggota] = {
                'id_anggota': id_anggota,
                'no_anggota': ag.get('no_anggota', ''),
                'nama_anggota': ag.get('nama', ''),
                'bunga_dibayar': 0.0,
            }
        bunga_per_anggota[id_anggota]['bunga_dibayar'] += j

    shu_kotor = total_bunga_dibayar
    shu_anggota_65 = shu_kotor * 0.65

    alokasi = []
    for row in bunga_per_anggota.values():
        prop = (row['bunga_dibayar'] / shu_kotor) if shu_kotor > 0 else 0.0
        shu_member = shu_anggota_65 * prop
        alokasi.append({
            'no_anggota': row['no_anggota'],
            'nama_anggota': row['nama_anggota'],
            'bunga_dibayar': row['bunga_dibayar'],
            'shu_anggota': shu_member,
        })

    alokasi.sort(key=lambda x: x['shu_anggota'], reverse=True)
    return {
        'shu_kotor': shu_kotor,
        'shu_anggota_65': shu_anggota_65,
        'shu_lain': shu_kotor - shu_anggota_65,
        'alokasi': alokasi
    }


def generate_excel_laporan_terpadu():
    """Buat workbook Excel berisi 2 tabel: pinjaman + simpanan, plus tabel SHU pinjaman."""
    ensure_pinjaman_plafon_schema()
    ensure_simpanan_schema()
    ensure_pinjaman_cicilan_schema()

    wb_cls = Workbook
    if wb_cls is None:
        try:
            from openpyxl import Workbook as wb_cls  # type: ignore
        except ModuleNotFoundError:
            raise RuntimeError("Fitur export membutuhkan openpyxl.")

    anggota_full = baca_csv(FILE_ANGGOTA)
    pinjaman_t = enrich_pinjaman_untuk_tampilan(baca_csv(FILE_PINJAMAN), anggota_full)
    pinjaman_t.reverse()

    wb = wb_cls()
    ws = wb.active
    ws.title = 'Laporan Terpadu'

    # -------------------
    # Tabel: PINJAMAN
    # -------------------
    ws.append(['DATA PINJAMAN'])
    ws.append(['No', 'Jenis', 'No Anggota', 'Nama', 'Plafon', 'Tenor', 'Bunga %', 'Cicilan/bln', 'Sisa', 'Status'])
    for i, p in enumerate(pinjaman_t, 1):
        ws.append([
            i,
            p.get('jenis_pinjaman', ''),
            p.get('no_anggota', ''),
            p.get('nama_anggota', ''),
            float(p.get('plafon') or p.get('total_pinjaman') or 0),
            int(float(p.get('tenor_bulan') or p.get('tenor') or 0)),
            float(p.get('bunga_persen') or 0),
            float(p.get('cicilan_per_bulan') or 0),
            saldo_pinjaman_aktual(p),
            p.get('status', ''),
        ])

    # -------------------
    # SHU (65% untuk anggota)
    # -------------------
    shu = hitung_shu_pinjaman()
    ws.append([])
    ws.append(['SHU PINJAMAN'])
    ws.append(['SHU Kotor (bunga dibayar)', float(shu['shu_kotor'])])
    ws.append(['SHU Anggota (65%)', float(shu['shu_anggota_65'])])
    ws.append(['SHU Lainnya (35%)', float(shu['shu_lain'])])
    ws.append([])

    ws.append(['Alokasi SHU Anggota 65% (proporsional bunga dibayar)'])
    ws.append(['No', 'No Anggota', 'Nama', 'Bunga dibayar', 'SHU anggota'])
    for i, a in enumerate(shu['alokasi'], 1):
        ws.append([i, a['no_anggota'], a['nama_anggota'], float(a['bunga_dibayar']), float(a['shu_anggota'])])

    # -------------------
    # Tabel: SIMPANAN
    # -------------------
    simpanan_t = enrich_simpanan_untuk_tampilan(baca_csv(FILE_SIMPANAN), anggota_full)
    simpanan_t.reverse()
    ws.append([])
    ws.append(['DATA SIMPANAN (saldo per anggota)'])
    ws.append(['No', 'No Anggota', 'Nama', 'Total Simpanan'])
    for i, s in enumerate(simpanan_t, 1):
        ws.append([
            i,
            s.get('no_anggota', ''),
            s.get('nama_anggota', ''),
            float(s.get('total_simpanan') or 0),
        ])

    return wb


@app.route('/laporan')
@login_required
def halaman_laporan():
    ensure_simpanan_schema()
    simpanan = baca_csv(FILE_SIMPANAN)
    pinjaman = baca_csv(FILE_PINJAMAN)
    anggota = baca_csv(FILE_ANGGOTA)
    if not is_current_user_admin():
        id_anggota = get_current_user_id_anggota()
        simpanan = [s for s in simpanan if s.get('id_anggota') == id_anggota]
        pinjaman = [p for p in pinjaman if p.get('id_anggota') == id_anggota]
        anggota = [a for a in anggota if a.get('id_anggota') == id_anggota]
    return render_template('laporan.html', simpanan=simpanan, pinjaman=pinjaman, anggota=anggota)


@app.route('/export/laporan-terpadu')
@admin_required
def export_laporan_terpadu():
    """Upsert anggota dari data pinjaman/simpanan lalu siapkan download Excel 1 file."""
    ensure_simpanan_schema()
    ensure_pinjaman_plafon_schema()

    pinjaman_rows = baca_csv(FILE_PINJAMAN)
    simpanan_rows = baca_csv(FILE_SIMPANAN)
    hasil = upsert_anggota_from_riwayat(pinjaman_rows, simpanan_rows)

    return render_template(
        'export_laporan_terpadu.html',
        added=hasil['added'],
        updated=hasil['updated'],
        download_url=url_for('export_laporan_terpadu_unduh'),
    )


@app.route('/export/laporan-terpadu/unduh')
@admin_required
def export_laporan_terpadu_unduh():
    wb_cls = Workbook
    if wb_cls is None:
        try:
            from openpyxl import Workbook as wb_cls  # type: ignore
        except ModuleNotFoundError:
            return f"Fitur export membutuhkan openpyxl. Jalankan: \"{sys.executable}\" -m pip install openpyxl", 500

    wb = generate_excel_laporan_terpadu()
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        download_name='laporan_terpadu_pinjam_simpan.xlsx',
        as_attachment=True
    )


@app.route('/export/simpanan')
@admin_required
def export_simpanan():
    ensure_simpanan_schema()
    wb_cls = Workbook
    if wb_cls is None:
        try:
            from openpyxl import Workbook as wb_cls  # type: ignore
        except ModuleNotFoundError:
            return f"Fitur export membutuhkan openpyxl. Jalankan: \"{sys.executable}\" -m pip install openpyxl", 500
    wb = wb_cls()
    ws = wb.active
    ws.title = "Data Simpanan"
    ws.append(['No', 'No Anggota', 'Nama', 'Total Simpanan'])
    simpanan = enrich_simpanan_untuk_tampilan(baca_csv(FILE_SIMPANAN), baca_csv(FILE_ANGGOTA))
    for i, s in enumerate(simpanan, 1):
        ws.append([i, s['no_anggota'], s['nama_anggota'], float(s.get('total_simpanan') or 0)])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     download_name='laporan_simpanan.xlsx', as_attachment=True)


@app.route('/export/pinjaman')
@admin_required
def export_pinjaman():
    wb_cls = Workbook
    if wb_cls is None:
        try:
            from openpyxl import Workbook as wb_cls  # type: ignore
        except ModuleNotFoundError:
            return f"Fitur export membutuhkan openpyxl. Jalankan: \"{sys.executable}\" -m pip install openpyxl", 500
    wb = wb_cls()
    ws = wb.active
    ws.title = "Data Pinjaman"
    ws.append(['No', 'Jenis', 'No Anggota', 'Nama', 'Plafon', 'Tenor', 'Cicilan/Bulan', 'Sisa', 'Status'])
    pinjaman = enrich_pinjaman_untuk_tampilan(baca_csv(FILE_PINJAMAN), baca_csv(FILE_ANGGOTA))
    for i, p in enumerate(pinjaman, 1):
        ws.append([
            i,
            p.get('jenis_pinjaman', ''),
            p['no_anggota'],
            p['nama_anggota'],
            float(p.get('plafon') or 0),
            int(float(p.get('tenor_bulan') or 0)),
            float(p.get('cicilan_per_bulan') or 0),
            saldo_pinjaman_aktual(p),
            p.get('status', ''),
        ])

    # Tambahkan SHU Pinjaman
    shu = hitung_shu_pinjaman()
    ws.append([])
    ws.append(['SHU PINJAMAN'])
    ws.append(['SHU Kotor (bunga dibayar)', float(shu['shu_kotor'])])
    ws.append(['SHU Anggota (65%)', float(shu['shu_anggota_65'])])
    ws.append(['SHU Lainnya (35%)', float(shu['shu_lain'])])
    ws.append([])
    ws.append(['Alokasi SHU Anggota 65% (proporsional bunga dibayar)'])
    ws.append(['No', 'No Anggota', 'Nama', 'Bunga dibayar', 'SHU anggota'])
    for i, a in enumerate(shu['alokasi'], 1):
        ws.append([i, a['no_anggota'], a['nama_anggota'], float(a['bunga_dibayar']), float(a['shu_anggota'])])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     download_name='laporan_pinjaman.xlsx', as_attachment=True)


@app.route('/laporan/anggota/<id_anggota>')
@login_required
def laporan_anggota(id_anggota):
    """Detail laporan per anggota."""
    restrict_id_anggota_or_forbid(id_anggota)
    anggota = baca_csv(FILE_ANGGOTA)
    anggota_data = next((a for a in anggota if a['id_anggota'] == id_anggota), None)
    if not anggota_data:
        return "Anggota tidak ditemukan", 404

    ensure_simpanan_schema()
    ensure_simpanan_transaksi_schema()
    anggota_all = baca_csv(FILE_ANGGOTA)
    simpanan_raw = [s for s in baca_csv(FILE_SIMPANAN) if s['id_anggota'] == id_anggota]
    simpanan_transaksi = [s for s in baca_csv(FILE_SIMPANAN_TRANSAKSI) if s.get('id_anggota') == id_anggota]
    for t in simpanan_transaksi:
        ket = (t.get('keterangan') or '').strip()
        if ket:
            t['keterangan'] = ket.replace('Setoran', 'Simpanan').replace('setoran', 'simpanan')
    pinjaman_raw = [p for p in baca_csv(FILE_PINJAMAN) if p['id_anggota'] == id_anggota]
    simpanan = enrich_simpanan_untuk_tampilan(simpanan_raw, anggota_all)
    pinjaman = enrich_pinjaman_untuk_tampilan(pinjaman_raw, anggota_all)

    pinjaman_laporan = []
    total_saldo_diterima = 0.0
    for p in pinjaman:
        row = dict(p)
        plafon = float(row.get('plafon') or 0)
        tenor = int(float(row.get('tenor_bulan') or 0))
        status = (row.get('status') or '').strip()
        provisi = plafon * PROVISI_RATE_LONG_TENOR if tenor > 12 else 0.0
        saldo_diterima = max(plafon - provisi, 0.0)
        if status == 'Disetujui':
            row['saldo_diterima'] = str(round(saldo_diterima, 2))
            total_saldo_diterima += saldo_diterima
        else:
            row['saldo_diterima'] = ''
        pinjaman_laporan.append(row)

    total_simpanan = sum(float(s.get('total_simpanan') or 0) for s in simpanan_raw)
    total_pinjaman = sum(
        saldo_pinjaman_aktual(p)
        for p in pinjaman_raw
        if p.get('status') == 'Disetujui'
    )

    return render_template('cetak_struk.html',
                           anggota=anggota_data,
                           simpanan=simpanan,
                           simpanan_transaksi=sorted(simpanan_transaksi, key=lambda x: x.get('tanggal', ''), reverse=True),
                           pinjaman=pinjaman_laporan,
                           total_simpanan=total_simpanan,
                           total_pinjaman=total_pinjaman,
                           total_saldo_diterima=total_saldo_diterima)


# ══════════════════════════════════════════════
#  JALANKAN APLIKASI
# ══════════════════════════════════════════════
if __name__ == '__main__':
    print("=" * 50)
    print("  APLIKASI KOPERASI BERBASIS WEB")
    print("  Login : http://localhost:5000")
    print("=" * 50)
    app.run(debug=False, host='127.0.0.1', port=5000)