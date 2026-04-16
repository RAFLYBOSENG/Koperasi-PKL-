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
    document.querySelectorAll('table').forEach(function (table) {
        const header = table.querySelector('thead');
        if (!header) return;

        // Skip transaksi table - already handled above
        if (table.dataset.transaksiTable === 'true') return;

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