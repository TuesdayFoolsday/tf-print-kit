"""Sealed box for the courier: X25519 + HKDF-SHA256 + AES-256-GCM.

Format of a sealed file: b"TFC1" | sender ephemeral public key (32) | nonce (12) | ciphertext+tag.
Only the holder of the recipient private key (a GitHub secret) can open it; GCM also proves it wasn't altered.
Keys are raw 32-byte X25519 keys written as base64 (44 characters).
"""
from __future__ import annotations

import base64
import os

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric.x25519 import X25519PrivateKey, X25519PublicKey
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.hkdf import HKDF

MAGIC = b"TFC1"
INFO = b"tf-courier v1"
RAW = dict(encoding=serialization.Encoding.Raw, format=serialization.PublicFormat.Raw)


def _key(shared: bytes, eph_pub: bytes, rcpt_pub: bytes) -> bytes:
    return HKDF(algorithm=hashes.SHA256(), length=32, salt=eph_pub + rcpt_pub, info=INFO).derive(shared)


def new_keypair() -> tuple[str, str]:
    """-> (private_b64, public_b64)"""
    priv = X25519PrivateKey.generate()
    priv_raw = priv.private_bytes(serialization.Encoding.Raw, serialization.PrivateFormat.Raw,
                                  serialization.NoEncryption())
    return base64.b64encode(priv_raw).decode(), base64.b64encode(priv.public_key().public_bytes(**RAW)).decode()


def seal(data: bytes, public_b64: str) -> bytes:
    rcpt_pub = base64.b64decode(public_b64.strip())
    eph = X25519PrivateKey.generate()
    eph_pub = eph.public_key().public_bytes(**RAW)
    key = _key(eph.exchange(X25519PublicKey.from_public_bytes(rcpt_pub)), eph_pub, rcpt_pub)
    nonce = os.urandom(12)
    return MAGIC + eph_pub + nonce + AESGCM(key).encrypt(nonce, data, MAGIC)


def open_box(blob: bytes, private_b64: str) -> bytes:
    if blob[:4] != MAGIC:
        raise ValueError("not a courier file")
    priv = X25519PrivateKey.from_private_bytes(base64.b64decode(private_b64.strip()))
    rcpt_pub = priv.public_key().public_bytes(**RAW)
    eph_pub, nonce, ct = blob[4:36], blob[36:48], blob[48:]
    key = _key(priv.exchange(X25519PublicKey.from_public_bytes(eph_pub)), eph_pub, rcpt_pub)
    return AESGCM(key).decrypt(nonce, ct, MAGIC)  # raises InvalidTag if altered or wrong key
