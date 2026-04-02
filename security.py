import hashlib, hmac, uuid
from config import SECRET_KEY

def generate_uid():
    return uuid.uuid4().hex

def generate_signature(reference: str, uid: str) -> str:
    msg = f"{reference}|{uid}".encode("utf-8")
    return hmac.new(SECRET_KEY.encode("utf-8"), msg, hashlib.sha256).hexdigest()

def verify_signature(reference: str, uid: str, sig: str) -> bool:
    expected = generate_signature(reference, uid)
    return hmac.compare_digest(expected, sig)
