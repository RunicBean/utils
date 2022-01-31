import hmac
import hashlib
import base64


def decrypt_hmac_sha256(message, secret):
    signature = hmac.new(
        # secret.encode(),
        bytes(secret, 'utf-8'),
        # msg=message.encode(),
        msg=bytes(message, 'utf-8'),
        digestmod=hashlib.sha256
    ).digest()
    base_bytes = base64.b64encode(signature)
    base_str = base_bytes.decode()
    return base_str
