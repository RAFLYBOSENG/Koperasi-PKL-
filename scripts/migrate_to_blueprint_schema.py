#!/usr/bin/env python3
"""
Migration script: Populate roles, permissions, role_permissions, settings dari blueprint.
Ini adalah migration aman yang hanya INSERT dan tidak DROP data existing.

Langkah:
1. Jalankan schema.sql di PostgreSQL dulu.
2. Jalankan script ini: python scripts/migrate_to_blueprint_schema.py
3. Cek hasil di DB: SELECT * FROM roles; SELECT * FROM permissions; SELECT * FROM role_permissions;
"""

import os
import sys
import uuid
from datetime import datetime

# Setup path untuk import koperasi_system
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from koperasi_system.db import db_session
from sqlalchemy import text

# Blueprint role + permission definitions
ROLE_DEFINITIONS = {
    'super_admin': {
        'deskripsi': 'Super Admin - Mengatur seluruh sistem, user, role, konfigurasi, backup, dan audit log.',
        'permissions': {
            'members.view', 'members.manage',
            'savings.deposit.request', 'savings.deposit.input', 'savings.deposit.validate',
            'savings.withdraw.request', 'savings.withdraw.validate',
            'loan.documents.review', 'loan.eligibility.analyze', 'loans.approve',
            'loan.disbursement.input',
            'installments.manage',
            'cash.manage', 'shu.manage', 'shu.validate', 'shu.view',
            'reports.export', 'reports.strategic.view',
            'backup.manage', 'excel.import', 'news.manage', 'users.manage', 'roles.manage',
            'audit.view', 'system.manage',
        }
    },
    'admin_koperasi': {
        'deskripsi': 'Admin Koperasi - Mengelola data anggota, transaksi simpanan, pinjaman, dan proses administrasi.',
        'permissions': {
            'members.view', 'members.manage',
            'savings.deposit.request', 'savings.deposit.input', 'savings.deposit.validate',
            'savings.withdraw.request', 'savings.withdraw.validate',
            'loan.documents.review', 'loan.disbursement.input',
            'installments.manage',
            'reports.export', 'reports.strategic.view',
            'excel.import', 'news.manage',
        }
    },
    'bendahara': {
        'deskripsi': 'Bendahara - Mengelola transaksi keuangan, kas, pembayaran, laporan, dan SHU.',
        'permissions': {
            'members.view',
            'savings.deposit.input', 'savings.deposit.validate',
            'savings.withdraw.validate',
            'loan.eligibility.analyze', 'loan.disbursement.input',
            'installments.manage',
            'cash.manage', 'shu.manage', 'shu.view',
            'reports.export', 'reports.strategic.view',
            'excel.import',
        }
    },
    'ketua_pengurus': {
        'deskripsi': 'Ketua/Pengurus - Menyetujui pinjaman dan melihat laporan strategis.',
        'permissions': {
            'members.view',
            'loans.approve',
            'shu.validate', 'shu.view',
            'reports.export', 'reports.strategic.view',
        }
    },
    'auditor': {
        'deskripsi': 'Auditor - Melihat laporan dan histori transaksi tanpa mengubah data.',
        'permissions': {
            'members.view',
            'shu.view',
            'audit.view',
            'reports.export', 'reports.strategic.view',
        }
    },
    'anggota': {
        'deskripsi': 'Anggota - Melihat data pribadi, simpanan, pinjaman, cicilan, SHU.',
        'permissions': {
            'members.self.view', 'members.self.edit.limited',
            'savings.deposit.request', 'savings.withdraw.request',
            'loan.request',
            'installments.proof.upload',
            'shu.self.view',
            'reports.self.view',
        }
    },
    'admin': {
        'deskripsi': 'Admin (Legacy) - Kompatibilitas dengan sistem lama.',
        'permissions': {
            'members.view', 'members.manage',
            'savings.deposit.request', 'savings.deposit.input', 'savings.deposit.validate',
            'savings.withdraw.request', 'savings.withdraw.validate',
            'loan.documents.review', 'loan.disbursement.input',
            'installments.manage',
            'reports.export', 'reports.strategic.view',
            'shu.view', 'excel.import', 'news.manage',
        }
    },
    'user': {
        'deskripsi': 'User (Legacy) - Kompatibilitas dengan sistem lama.',
        'permissions': {
            'members.self.view', 'members.self.edit.limited',
            'savings.deposit.request', 'savings.withdraw.request',
            'loan.request', 'installments.proof.upload',
            'shu.self.view', 'reports.self.view',
        }
    },
}

# Settings defaults sesuai blueprint
SETTINGS_DEFAULTS = {
    'nominal_simpanan_pokok': ('500000', 'NUMERIC', 'Nominal simpanan pokok wajib (Rp)'),
    'nominal_simpanan_wajib': ('250000', 'NUMERIC', 'Nominal simpanan wajib bulanan (Rp)'),
    'nominal_iuran_sosial': ('10000', 'NUMERIC', 'Nominal iuran sosial bulanan (Rp)'),
    'nominal_hari_koperasi': ('20000', 'NUMERIC', 'Nominal simpanan hari koperasi bulanan (Rp)'),
    'bunga_jangka_pendek': ('1.5', 'NUMERIC', 'Bunga pinjaman jangka pendek (%) per bulan'),
    'bunga_jangka_panjang': ('0.8', 'NUMERIC', 'Bunga pinjaman jangka panjang 13-24 bulan (%) per bulan'),
    'bunga_jangka_panjang_lama': ('0.75', 'NUMERIC', 'Bunga pinjaman jangka panjang 25+ bulan (%) per bulan'),
    'bunga_solusi_cepat': ('2.0', 'NUMERIC', 'Bunga pinjaman solusi cepat (%) per bulan'),
    'bunga_modal_bisnis': ('0.5', 'NUMERIC', 'Bunga pinjaman modal bisnis (%) per bulan'),
    'tenor_min_peminjam': ('2', 'INTEGER', 'Minimal bulan menjadi anggota sebelum boleh pinjam'),
    'tenor_min_topup': ('3', 'INTEGER', 'Minimal bulan sebelum topup pinjaman diizinkan'),
    'shu_cadangan_umum': ('15', 'NUMERIC', 'Persentase SHU untuk cadangan umum'),
    'shu_pasif': ('15', 'NUMERIC', 'Persentase SHU untuk hak pasif semua anggota'),
    'shu_aktif': ('50', 'NUMERIC', 'Persentase SHU untuk hak aktif sesuai jasa anggota'),
    'shu_kesejahteraan_karyawan': ('5', 'NUMERIC', 'Persentase SHU untuk dana kesejahteraan karyawan'),
    'shu_pendidikan': ('0.5', 'NUMERIC', 'Persentase SHU untuk dana pendidikan'),
    'shu_sosial': ('2', 'NUMERIC', 'Persentase SHU untuk dana sosial'),
    'shu_pembangunan': ('0.5', 'NUMERIC', 'Persentase SHU untuk dana pembangunan'),
    'shu_pengurus': ('10', 'NUMERIC', 'Persentase SHU untuk dana pengurus'),
    'shu_risiko': ('2', 'NUMERIC', 'Persentase SHU untuk dana risiko'),
}


def migrate_roles_permissions():
    """Insert roles dan permissions ke database."""
    with db_session() as conn:
        # Baca existing roles
        existing_roles = set()
        try:
            result = conn.execute(text("SELECT role_name FROM roles"))
            existing_roles = {row[0] for row in result}
        except Exception:
            pass  # Tabel mungkin belum ada

        # Insert roles baru
        for role_name, role_data in ROLE_DEFINITIONS.items():
            if role_name in existing_roles:
                print(f"  Role '{role_name}' sudah ada, skip.")
                continue
            role_id = str(uuid.uuid4())
            try:
                conn.execute(
                    text("""
                        INSERT INTO roles (id_role, role_name, deskripsi, created_at, updated_at)
                        VALUES (:id, :name, :desc, NOW(), NOW())
                    """),
                    {'id': role_id, 'name': role_name, 'desc': role_data['deskripsi']}
                )
                print(f"  ✓ Role '{role_name}' ditambahkan.")
            except Exception as e:
                print(f"  ✗ Gagal insert role '{role_name}': {e}")

        # Baca existing permissions
        existing_perms = {}
        try:
            result = conn.execute(text("SELECT id_permission, permission_name FROM permissions"))
            existing_perms = {row[1]: row[0] for row in result}
        except Exception:
            pass

        # Collect semua permissions yang diperlukan
        all_perms = set()
        for role_data in ROLE_DEFINITIONS.values():
            all_perms.update(role_data['permissions'])

        # Insert permissions baru
        for perm_name in sorted(all_perms):
            if perm_name in existing_perms:
                continue
            perm_id = str(uuid.uuid4())
            try:
                conn.execute(
                    text("""
                        INSERT INTO permissions (id_permission, permission_name, created_at)
                        VALUES (:id, :name, NOW())
                    """),
                    {'id': perm_id, 'name': perm_name}
                )
                existing_perms[perm_name] = perm_id
                print(f"  ✓ Permission '{perm_name}' ditambahkan.")
            except Exception as e:
                print(f"  ✗ Gagal insert permission '{perm_name}': {e}")

        # Baca existing role_permissions
        existing_role_perms = set()
        try:
            result = conn.execute(text("""
                SELECT rp.id_role, rp.id_permission 
                FROM role_permissions rp 
                JOIN roles r ON rp.id_role = r.id_role 
                JOIN permissions p ON rp.id_permission = p.id_permission 
                WHERE r.role_name = :role AND p.permission_name = :perm
            """))
            # Ini query lain yang lebih efisien - ambil semua dulu
            result = conn.execute(text("SELECT id_role, id_permission FROM role_permissions"))
            existing_role_perms = {(row[0], row[1]) for row in result}
        except Exception:
            pass

        # Baca roles baru (ID) berdasarkan nama
        role_ids = {}
        try:
            result = conn.execute(text("SELECT id_role, role_name FROM roles"))
            role_ids = {row[1]: row[0] for row in result}
        except Exception as e:
            print(f"  ✗ Gagal baca roles: {e}")
            return

        # Insert role_permissions
        for role_name, role_data in ROLE_DEFINITIONS.items():
            role_id = role_ids.get(role_name)
            if not role_id:
                print(f"  Role ID untuk '{role_name}' tidak ditemukan.")
                continue
            for perm_name in role_data['permissions']:
                perm_id = existing_perms.get(perm_name)
                if not perm_id:
                    print(f"  Permission ID untuk '{perm_name}' tidak ditemukan.")
                    continue
                if (role_id, perm_id) in existing_role_perms:
                    continue
                try:
                    conn.execute(
                        text("""
                            INSERT INTO role_permissions (id_role, id_permission, assigned_at)
                            VALUES (:role_id, :perm_id, NOW())
                        """),
                        {'role_id': role_id, 'perm_id': perm_id}
                    )
                except Exception:
                    pass  # Mungkin sudah ada, tidak masalah

        print(f"  ✓ Role-permission mappings selesai.")


def migrate_settings():
    """Insert settings default ke database."""
    with db_session() as conn:
        # Baca existing settings
        existing_settings = set()
        try:
            result = conn.execute(text("SELECT kunci FROM settings"))
            existing_settings = {row[0] for row in result}
        except Exception:
            pass

        # Insert settings baru
        for kunci, (nilai, tipe, deskripsi) in SETTINGS_DEFAULTS.items():
            if kunci in existing_settings:
                print(f"  Setting '{kunci}' sudah ada, skip.")
                continue
            try:
                conn.execute(
                    text("""
                        INSERT INTO settings (id_setting, kunci, nilai, tipe_data, deskripsi, updated_at)
                        VALUES (:id, :kunci, :nilai, :tipe, :deskripsi, NOW())
                    """),
                    {
                        'id': str(uuid.uuid4()),
                        'kunci': kunci,
                        'nilai': nilai,
                        'tipe': tipe,
                        'deskripsi': deskripsi
                    }
                )
                print(f"  ✓ Setting '{kunci}' = '{nilai}'")
            except Exception as e:
                print(f"  ✗ Gagal insert setting '{kunci}': {e}")


def main():
    print("=" * 60)
    print("Migration: Blueprint Schema - Roles, Permissions, Settings")
    print("=" * 60)

    try:
        print("\n[1/2] Migrating roles and permissions...")
        migrate_roles_permissions()
        print("  ✓ Roles & Permissions migration selesai.")
    except Exception as e:
        print(f"  ✗ Error selama migrasi roles/permissions: {e}")
        sys.exit(1)

    try:
        print("\n[2/2] Migrating settings defaults...")
        migrate_settings()
        print("  ✓ Settings migration selesai.")
    except Exception as e:
        print(f"  ✗ Error selama migrasi settings: {e}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("✓ Semua migrasi blueprint schema berhasil!")
    print("=" * 60)


if __name__ == '__main__':
    main()
