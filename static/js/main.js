function initializeMainPage() {
    // Toggle sidebar
    const sidebarCollapse = document.getElementById('sidebarCollapse');
    const sidebar = document.getElementById('sidebar');
    const content = document.getElementById('content');

    if (sidebarCollapse) {
        sidebarCollapse.addEventListener('click', function () {
            sidebar.classList.toggle('active');
            content.classList.toggle('active');
        });
    }

    // Auto hide flash messages after 5 seconds
    setTimeout(function () {
        const alerts = document.querySelectorAll('.alert-dismissible');
        alerts.forEach(function (alert) {
            const bsAlert = new bootstrap.Alert(alert);
            bsAlert.close();
        });
    }, 5000);

    // Set current date in navbar
    const nowEl = document.querySelector('.text-muted i.fa-calendar');
    if (nowEl) {
        const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        const today = new Date().toLocaleDateString('id-ID', options);
        const dateTextEl = nowEl.nextElementSibling;
        if (dateTextEl) {
            dateTextEl.textContent = today;
        }
    }

    // ===== FILTER TRANSAKSI LOGIC =====
    const transaksiTable = document.querySelector('table[data-transaksi-table="true"]');
    if (transaksiTable) {
        const filterButtons = document.querySelectorAll('[data-transaksi-filter]');
        let currentFilter = 'all';

        function applyFilter(filterValue) {
            const rows = transaksiTable.querySelectorAll('tbody tr');
            rows.forEach(function (row) {
                const rowType = (row.getAttribute('data-transaksi-type') || '').trim();
                row.style.display = filterValue === 'all' || rowType === filterValue ? '' : 'none';
            });
        }

        function syncButtonStyles(activeFilter) {
            filterButtons.forEach(function (btn) {
                const btnFilter = btn.getAttribute('data-transaksi-filter');
                btn.classList.remove('btn-primary', 'btn-success', 'btn-warning', 'btn-info', 'active');
                btn.classList.remove('btn-outline-primary', 'btn-outline-success', 'btn-outline-warning', 'btn-outline-info');

                if (btnFilter === activeFilter) {
                    btn.classList.add('active');
                    if (btnFilter === 'all') {
                        btn.classList.add('btn-primary');
                    } else if (btnFilter === 'Simpanan') {
                        btn.classList.add('btn-success');
                    } else if (btnFilter === 'Pinjaman') {
                        btn.classList.add('btn-warning');
                    } else if (btnFilter === 'Cicilan') {
                        btn.classList.add('btn-info');
                    }
                } else {
                    if (btnFilter === 'all') {
                        btn.classList.add('btn-outline-primary');
                    } else if (btnFilter === 'Simpanan') {
                        btn.classList.add('btn-outline-success');
                    } else if (btnFilter === 'Pinjaman') {
                        btn.classList.add('btn-outline-warning');
                    } else if (btnFilter === 'Cicilan') {
                        btn.classList.add('btn-outline-info');
                    }
                }
            });
        }

        filterButtons.forEach(function (button) {
            button.addEventListener('click', function (e) {
                e.preventDefault();
                currentFilter = this.getAttribute('data-transaksi-filter') || 'all';
                syncButtonStyles(currentFilter);
                applyFilter(currentFilter);
            });
        });

        applyFilter('all');
        syncButtonStyles('all');
    }

    // ===== STANDARD TABLE SEARCH (for non-transaksi tables) =====
    const pinjamanTable = document.querySelector('table[data-pinjaman-table="true"]');
    if (pinjamanTable) {
        const pinjamanFilterButtons = document.querySelectorAll('[data-pinjaman-filter]');
        const pinjamanJenisButtons = document.querySelectorAll('[data-pinjaman-jenis-filter]');
        const pinjamanSearchInput = document.querySelector('[data-pinjaman-search="true"]');
        let currentPinjamanFilter = 'all';
        let currentPinjamanJenisFilter = 'all';
        let currentPinjamanKeyword = '';

        function applyPinjamanFilter(filterStatus, filterJenis, keyword) {
            const rows = pinjamanTable.querySelectorAll('tbody tr[data-pinjaman-status]');
            let visibleCount = 0;
            rows.forEach(function (row) {
                const rowStatus = (row.getAttribute('data-pinjaman-status') || '').trim();
                const rowJenis = (row.getAttribute('data-pinjaman-jenis') || '').trim();
                const rowText = (row.textContent || '').toLowerCase();
                const statusMatch = filterStatus === 'all' || rowStatus === filterStatus;
                const jenisMatch = filterJenis === 'all' || rowJenis === filterJenis;
                const searchMatch = !keyword || rowText.includes(keyword);
                const isVisible = statusMatch && jenisMatch && searchMatch;
                row.style.display = isVisible ? '' : 'none';
                if (isVisible) {
                    visibleCount += 1;
                }
            });

            const emptyRow = pinjamanTable.querySelector('tr[data-pinjaman-empty="true"]');
            const noResultRow = pinjamanTable.querySelector('tr[data-pinjaman-noresult="true"]');
            if (emptyRow) {
                emptyRow.style.display = '';
            }
            if (noResultRow) {
                noResultRow.style.display = visibleCount === 0 ? '' : 'none';
            }
        }

        function syncPinjamanButtonStyles(activeFilter) {
            pinjamanFilterButtons.forEach(function (btn) {
                const btnFilter = btn.getAttribute('data-pinjaman-filter');
                btn.classList.remove('btn-primary', 'btn-warning', 'btn-success', 'btn-secondary', 'btn-danger', 'active');
                btn.classList.remove('btn-outline-primary', 'btn-outline-warning', 'btn-outline-success', 'btn-outline-secondary', 'btn-outline-danger');

                if (btnFilter === activeFilter) {
                    btn.classList.add('active');
                    if (btnFilter === 'all') {
                        btn.classList.add('btn-primary');
                    } else if (btnFilter === 'Menunggu') {
                        btn.classList.add('btn-warning');
                    } else if (btnFilter === 'Disetujui') {
                        btn.classList.add('btn-success');
                    } else if (btnFilter === 'Lunas') {
                        btn.classList.add('btn-secondary');
                    } else if (btnFilter === 'Ditolak') {
                        btn.classList.add('btn-danger');
                    }
                } else {
                    if (btnFilter === 'all') {
                        btn.classList.add('btn-outline-primary');
                    } else if (btnFilter === 'Menunggu') {
                        btn.classList.add('btn-outline-warning');
                    } else if (btnFilter === 'Disetujui') {
                        btn.classList.add('btn-outline-success');
                    } else if (btnFilter === 'Lunas') {
                        btn.classList.add('btn-outline-secondary');
                    } else if (btnFilter === 'Ditolak') {
                        btn.classList.add('btn-outline-danger');
                    }
                }
            });
        }

        pinjamanFilterButtons.forEach(function (button) {
            button.addEventListener('click', function (e) {
                e.preventDefault();
                currentPinjamanFilter = this.getAttribute('data-pinjaman-filter') || 'all';
                syncPinjamanButtonStyles(currentPinjamanFilter);
                applyPinjamanFilter(currentPinjamanFilter, currentPinjamanJenisFilter, currentPinjamanKeyword);
            });
        });

        function syncPinjamanJenisButtonStyles(activeFilter) {
            pinjamanJenisButtons.forEach(function (btn) {
                const btnFilter = btn.getAttribute('data-pinjaman-jenis-filter');
                btn.classList.remove('btn-primary', 'btn-info', 'btn-warning', 'btn-success', 'btn-dark', 'active');
                btn.classList.remove('btn-outline-primary', 'btn-outline-info', 'btn-outline-warning', 'btn-outline-success', 'btn-outline-dark');

                if (btnFilter === activeFilter) {
                    btn.classList.add('active');
                    if (btnFilter === 'all') {
                        btn.classList.add('btn-primary');
                    } else if (btnFilter === 'Solusi Cepat') {
                        btn.classList.add('btn-info');
                    } else if (btnFilter === 'Jangka Pendek') {
                        btn.classList.add('btn-warning');
                    } else if (btnFilter === 'Jangka Panjang') {
                        btn.classList.add('btn-success');
                    } else if (btnFilter === 'Modal Usaha') {
                        btn.classList.add('btn-dark');
                    }
                } else {
                    if (btnFilter === 'all') {
                        btn.classList.add('btn-outline-primary');
                    } else if (btnFilter === 'Solusi Cepat') {
                        btn.classList.add('btn-outline-info');
                    } else if (btnFilter === 'Jangka Pendek') {
                        btn.classList.add('btn-outline-warning');
                    } else if (btnFilter === 'Jangka Panjang') {
                        btn.classList.add('btn-outline-success');
                    } else if (btnFilter === 'Modal Usaha') {
                        btn.classList.add('btn-outline-dark');
                    }
                }
            });
        }

        pinjamanJenisButtons.forEach(function (button) {
            button.addEventListener('click', function (e) {
                e.preventDefault();
                currentPinjamanJenisFilter = this.getAttribute('data-pinjaman-jenis-filter') || 'all';
                syncPinjamanJenisButtonStyles(currentPinjamanJenisFilter);
                applyPinjamanFilter(currentPinjamanFilter, currentPinjamanJenisFilter, currentPinjamanKeyword);
            });
        });

        if (pinjamanSearchInput) {
            pinjamanSearchInput.addEventListener('input', function () {
                currentPinjamanKeyword = (this.value || '').trim().toLowerCase();
                applyPinjamanFilter(currentPinjamanFilter, currentPinjamanJenisFilter, currentPinjamanKeyword);
            });
        }

        applyPinjamanFilter('all', 'all', '');
        syncPinjamanButtonStyles('all');
        syncPinjamanJenisButtonStyles('all');
    }

    document.querySelectorAll('table').forEach(function (table) {
        const header = table.querySelector('thead');
        if (!header) return;

        // Skip transaksi table - already handled above
        if (table.dataset.transaksiTable === 'true') return;
        if (table.dataset.pinjamanTable === 'true') return;

        // Add search input before table
        const searchInput = document.createElement('input');
        searchInput.type = 'text';
        searchInput.className = 'form-control mb-2';
        searchInput.placeholder = '🔍 Cari di tabel...';
        searchInput.style.maxWidth = '300px';

        const wrapper = table.parentElement;
        wrapper.insertBefore(searchInput, table);

        searchInput.addEventListener('keyup', function () {
            const filter = this.value.toLowerCase();
            const rows = table.querySelectorAll('tbody tr');
            rows.forEach(function (row) {
                const text = row.textContent.toLowerCase();
                row.style.display = text.includes(filter) ? '' : 'none';
            });
        });
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeMainPage);
} else {
    initializeMainPage();
}

window.addEventListener('pageshow', function (event) {
    if (event.persisted && window.IS_AUTHENTICATED) {
        window.location.reload();
    }
});