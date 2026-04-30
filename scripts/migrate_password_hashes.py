from __future__ import annotations

import argparse

from sqlalchemy import select
from werkzeug.security import generate_password_hash

from enterprise_cooperative.config import get_config_object
from enterprise_cooperative.database import init_database, session_scope
from enterprise_cooperative.models.user import User


def looks_like_hash(value: str) -> bool:
    return value.startswith('pbkdf2:') or value.startswith('scrypt:')


def migrate_user_passwords(fallback_password: str, overwrite_all: bool = False) -> None:
    migrated = 0
    skipped = 0
    with session_scope() as db:
        users = db.scalars(select(User)).all()
        for user in users:
            current = user.password_hash or ''
            if not overwrite_all and looks_like_hash(current):
                skipped += 1
                continue
            source_password = fallback_password if (overwrite_all or not current.strip()) else current
            user.password_hash = generate_password_hash(source_password)
            migrated += 1
    print(f'Migrasi password selesai. migrated={migrated}, skipped={skipped}, overwrite_all={overwrite_all}')


def main() -> None:
    parser = argparse.ArgumentParser(description='Migrasi password lama/plaintext ke format hash werkzeug.')
    parser.add_argument('--fallback-password', default='ChangeMe123!', help='Password fallback jika field lama kosong.')
    parser.add_argument('--overwrite-all', action='store_true', help='Hash ulang semua akun termasuk yang sudah hash.')
    parser.add_argument('--env', default='development', help='Environment config: development/production.')
    args = parser.parse_args()

    config = get_config_object(args.env)
    init_database(config.DATABASE_URI, echo=bool(config.SQLALCHEMY_ECHO))
    migrate_user_passwords(fallback_password=args.fallback_password, overwrite_all=args.overwrite_all)


if __name__ == '__main__':
    main()
