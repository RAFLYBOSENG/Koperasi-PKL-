import os

from koperasi_system.app_system import app


if __name__ == '__main__':
    host = os.getenv('FLASK_HOST', '127.0.0.1')
    port = int(os.getenv('FLASK_PORT', '5000'))
    debug = os.getenv('FLASK_DEBUG', '0') == '1'

    print('=' * 50)
    print('  APLIKASI KOPERASI BERBASIS WEB')
    print('  Server bootstrap: app.py')
    print('  Sistem inti: koperasi_system/app_system.py')
    print(f'  Beranda: http://{host}:{port}/')
    print(f'  Login : http://{host}:{port}/login')
    print('=' * 50)
    app.run(debug=debug, host=host, port=port)
