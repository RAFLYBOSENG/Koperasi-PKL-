from flask import Blueprint, current_app, redirect, url_for


def _dispatch(endpoint_name):
    def view(*args, **kwargs):
        return current_app.view_functions[endpoint_name](*args, **kwargs)

    view.__name__ = f'proxy_{endpoint_name}'
    return view


public_portal = Blueprint('public_portal', __name__, url_prefix='/app')
admin_portal = Blueprint('admin_portal', __name__, url_prefix='/admin')


@public_portal.route('/')
def landing():
    return _dispatch('landing')()


@public_portal.route('/login', methods=['GET', 'POST'])
def login():
    return _dispatch('login')()


@public_portal.route('/lupa-password', methods=['GET', 'POST'])
def lupa_password():
    return _dispatch('lupa_password')()


@public_portal.route('/logout')
def logout():
    return _dispatch('logout')()


@public_portal.route('/dashboard')
def dashboard():
    return _dispatch('dashboard')()


@public_portal.route('/simpanan')
def simpanan():
    return _dispatch('halaman_simpanan')()


@public_portal.route('/pinjaman')
def pinjaman():
    return _dispatch('halaman_pinjaman')()


@admin_portal.route('/')
def index():
    return redirect(url_for('admin_portal.dashboard'))


@admin_portal.route('/dashboard')
def dashboard():
    return _dispatch('dashboard')()


@admin_portal.route('/users')
def users():
    return _dispatch('users_index')()


@admin_portal.route('/simpanan')
def simpanan():
    return _dispatch('halaman_simpanan')()


@admin_portal.route('/pinjaman')
def pinjaman():
    return _dispatch('halaman_pinjaman')()


def register_route_aliases(app):
    app.register_blueprint(public_portal)
    app.register_blueprint(admin_portal)