#!/usr/bin/env python3
"""
encrypt_regions.py — Chiffre data_regions.js → data_regions.enc.js (AES-256-GCM)

Usage :
    python3 encrypt_regions.py            # chiffre
    python3 encrypt_regions.py --decrypt  # déchiffre (urgence)
"""

import base64, os, sys, getpass, re
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes


def derive_key(password: str, salt: bytes) -> bytes:
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=200000)
    return kdf.derive(password.encode('utf-8'))


def encrypt():
    try:
        with open('data_regions.js', 'r', encoding='utf-8') as f:
            plaintext = f.read()
    except FileNotFoundError:
        print("✗ data_regions.js introuvable — lancez d'abord : python3 import_regions.py")
        sys.exit(1)

    pwd  = getpass.getpass("Mot de passe     : ")
    pwd2 = getpass.getpass("Confirmer        : ")
    if pwd != pwd2:
        print("✗ Les mots de passe ne correspondent pas")
        sys.exit(1)
    if len(pwd) < 8:
        print("✗ Mot de passe trop court (8 caractères minimum)")
        sys.exit(1)

    salt  = os.urandom(16)
    nonce = os.urandom(12)
    key   = derive_key(pwd, salt)
    ct    = AESGCM(key).encrypt(nonce, plaintext.encode('utf-8'), None)

    payload = base64.b64encode(salt + nonce + ct).decode()
    with open('data_regions.enc.js', 'w', encoding='utf-8') as f:
        f.write('// Données régions chiffrées — ne pas modifier\n')
        f.write(f'const ENCRYPTED_REGIONS = "{payload}";\n')

    print(f"✓ data_regions.enc.js généré ({len(plaintext):,} → {len(payload):,} car. base64)")


def decrypt():
    try:
        with open('data_regions.enc.js', 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        print("✗ data_regions.enc.js introuvable")
        sys.exit(1)

    m = re.search(r'const ENCRYPTED_REGIONS = "(.+?)";', content)
    if not m:
        print("✗ Format data_regions.enc.js invalide")
        sys.exit(1)

    pwd = getpass.getpass("Mot de passe : ")
    raw = base64.b64decode(m.group(1))
    salt, nonce, ct = raw[:16], raw[16:28], raw[28:]

    try:
        key       = derive_key(pwd, salt)
        plaintext = AESGCM(key).decrypt(nonce, ct, None).decode('utf-8')
    except Exception:
        print("✗ Mot de passe incorrect ou fichier corrompu")
        sys.exit(1)

    with open('data_regions.js', 'w', encoding='utf-8') as f:
        f.write(plaintext)
    print("✓ data_regions.js restauré")


if __name__ == '__main__':
    if '--decrypt' in sys.argv:
        decrypt()
    else:
        encrypt()