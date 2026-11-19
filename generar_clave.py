"""Genera el hash de una contraseña para pegarlo en secrets.toml.

Uso:  python generar_clave.py
"""
import getpass
import hashlib
import secrets

ITER = 200_000


def hash_clave(clave: str) -> str:
    sal = secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac('sha256', clave.encode(), bytes.fromhex(sal), ITER).hex()
    return f'pbkdf2_sha256${ITER}${sal}${h}'


def verificar(clave: str, guardado: str) -> bool:
    try:
        _, it, sal, h = guardado.split('$')
    except ValueError:
        return False
    calc = hashlib.pbkdf2_hmac('sha256', clave.encode(), bytes.fromhex(sal), int(it)).hex()
    return secrets.compare_digest(calc, h)


if __name__ == '__main__':
    c1 = getpass.getpass('Contraseña nueva: ')
    c2 = getpass.getpass('Repítala: ')
    if c1 != c2 or len(c1) < 8:
        raise SystemExit('Las contraseñas no coinciden o tienen menos de 8 caracteres.')
    print('\nPegue esta línea en secrets.toml como "clave":\n')
    print(hash_clave(c1))
