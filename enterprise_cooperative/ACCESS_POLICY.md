# Kebijakan Hak Akses Sistem Koperasi

Sistem mengatur hak akses pengguna ke dalam enam peran utama dengan kewenangan spesifik.

- **Super Admin** memegang kendali tertinggi untuk mengatur seluruh sistem, mengelola pengguna dan role, melakukan backup database, serta memantau konfigurasi dan audit log.
- **Admin Koperasi** bertugas menjalankan operasional administrasi: mengelola data anggota, menginput dan memvalidasi setoran simpanan/angsuran, memeriksa berkas pinjaman, serta melakukan import data dari Excel.
- **Bendahara** berfokus pada keuangan koperasi: mengelola kas, menganalisis kemampuan bayar peminjam, menginput pencairan pinjaman, dan mengelola SHU.
- **Ketua/Pengurus** berperan dalam pengambilan keputusan strategis: menyetujui/menolak pinjaman, melihat serta memvalidasi SHU, dan memantau laporan operasional.
- **Anggota** memiliki akses mandiri untuk melihat data pribadi, memantau simpanan dan hak SHU, mengajukan pinjaman/setoran, serta mengunggah bukti angsuran dengan batas edit pada data sendiri.
- **Auditor** memiliki akses pengawasan baca-saja: melihat data anggota, histori transaksi, laporan, dan data SHU tanpa hak menambah, mengedit, atau mengubah data.

## Implementasi Teknis

- Definisi role: `enterprise_cooperative/core/rbac.py`
- Mapping permission per role: `enterprise_cooperative/core/permissions.py`
- Penegakan akses endpoint: `enterprise_cooperative/core/auth.py` (`require_roles`, `require_permission`)

Dokumen ini menjadi acuan kebijakan akses untuk pengembangan backend, frontend, QA, dan audit internal.
