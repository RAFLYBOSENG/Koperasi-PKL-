document.addEventListener('DOMContentLoaded', function () {
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
        nowEl.nextElementSibling.textContent = today;
    }

    // DataTable-like search for tables
    document.querySelectorAll('table').forEach(function (table) {
        const header = table.querySelector('thead');
        if (!header) return;

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
});

window.addEventListener('pageshow', function (event) {
    if (event.persisted && window.IS_AUTHENTICATED) {
        window.location.reload();
    }
});