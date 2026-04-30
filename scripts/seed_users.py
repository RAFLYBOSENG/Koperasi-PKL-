from __future__ import annotations

import argparse

from sqlalchemy import select
from werkzeug.security import generate_password_hash

from enterprise_cooperative.config import get_config_object
from enterprise_cooperative.database import init_database, session_scope
from enterprise_cooperative.models.user import User

DEFAULT_USERS = [
    ('superadmin', 'super_admin'),
    ('adminkoperasi', 'admin_koperasi'),
    ('bendahara', 'bendahara'),
    ('ketuapengurus', 'ketua_pengurus'),
    ('anggota_demo', 'anggota'),
    ('auditor', 'auditor'),
]


def seed_users(default_password: str, overwrite: bool = False) -> None:
    created = 0
    updated = 0
    with session_scope() as db:
        for username, role in DEFAULT_USERS:
            existing = db.scalar(select(User).where(User.username == username))
            if existing is None:
                db.add(
                    User(
                        username=username,
                        password_hash=generate_password_hash(default_password),
                        role=role,
                        is_active=True,
                    )
                )
                created += 1
                continue
            if overwrite:
                existing.password_hash = generate_password_hash(default_password)
                existing.role = role
                existing.is_active = True
                updated += 1
    print(f'Seed selesai. created={created}, updated={updated}, overwrite={overwrite}')


def main() -> None:
    parser = argparse.ArgumentParser(description='Seed default users per role untuk Enterprise Cooperative.')
    parser.add_argument('--password', required=True, help='Password default untuk semua user seed.')
    parser.add_argument('--overwrite', action='store_true', help='Timpa user existing (password/role/is_active).')
    parser.add_argument('--env', default='development', help='Environment config: development/production.')
    args = parser.parse_args()

    config = get_config_object(args.env)
    init_database(config.DATABASE_URI, echo=bool(config.SQLALCHEMY_ECHO))
    seed_users(default_password=args.password, overwrite=args.overwrite)


if __name__ == '__main__':
    main()
