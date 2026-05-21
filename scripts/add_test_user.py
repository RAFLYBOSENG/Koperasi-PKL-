import uuid
from werkzeug.security import generate_password_hash

id_user = str(uuid.uuid4())
username = 'tester'
password = 'test123'
role = 'user'
id_anggota = '9b6a5ab2-720d-4c05-bb25-0b8629c85dd3'
created_at = '2026-05-21'

pwd_hash = generate_password_hash(password)
line = f"{id_user},{username},{pwd_hash},{role},{id_anggota},{created_at}\n"

with open('data/users.csv', 'a', encoding='utf-8') as fh:
    fh.write(line)

print('Appended user tester with password test123')
