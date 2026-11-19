"""Contraseñas: una por departamento (en secrets) y una por municipio, calculada.

Las contraseñas municipales no se guardan en ningún lado: se calculan a partir de
una "semilla" secreta y del código del municipio (HMAC-SHA256). Así, 340 claves
distintas salen de un solo valor en los secrets. Para cambiar la clave de un
municipio, se sube su número de versión en [rotacion].
"""
import base64
import hashlib
import hmac
import secrets as _secrets


def clave_municipal(semilla: str, cod_muni, version: int = 1) -> str:
    h = hmac.new(semilla.encode(), f'{int(cod_muni)}:{int(version)}'.encode(), hashlib.sha256).digest()
    s = base64.b32encode(h).decode()[:8]          # letters A-Z and digits 2-7
    return f'{s[:4]}-{s[4:]}'


def _normal(c: str) -> str:
    return ''.join(ch for ch in c.upper() if ch.isalnum())


def coincide(escrita: str, correcta: str) -> bool:
    return _secrets.compare_digest(_normal(escrita).encode(), _normal(correcta).encode())


def igual(escrita: str, correcta: str) -> bool:
    """exact comparison (departmental passwords are case-sensitive)."""
    return _secrets.compare_digest(escrita.encode(), correcta.encode())
