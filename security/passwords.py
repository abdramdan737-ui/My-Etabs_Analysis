import hashlib, hmac, secrets

_ITERATIONS = 310_000

def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac('sha256', password.encode(), salt, _ITERATIONS)
    return f"pbkdf2_sha256${_ITERATIONS}${salt.hex()}${digest.hex()}"

def verify_password(password: str, encoded: str | None) -> bool:
    if not encoded:
        return False
    try:
        scheme, iterations, salt_hex, digest_hex = encoded.split('$')
        if scheme != 'pbkdf2_sha256': return False
        digest = hashlib.pbkdf2_hmac('sha256', password.encode(), bytes.fromhex(salt_hex), int(iterations))
        return hmac.compare_digest(digest.hex(), digest_hex)
    except (ValueError, TypeError):
        return False


def hash_access_code(code: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac('sha256', code.strip().upper().encode(), salt, 250_000)
    return f"{salt.hex()}${digest.hex()}"

def verify_access_code(code: str, encoded: str) -> bool:
    try:
        salt_hex, digest_hex = encoded.split('$')
        digest = hashlib.pbkdf2_hmac('sha256', code.strip().upper().encode(), bytes.fromhex(salt_hex), 250_000)
        return hmac.compare_digest(digest.hex(), digest_hex)
    except (ValueError, TypeError):
        return False
