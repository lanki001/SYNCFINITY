# test_syncfinity.py
import pytest
from datetime import datetime, timedelta
from syncfinity_2_0 import LicenseManager, MAX_USERS, CHARSET


# ─── CHARSET TESTS ───────────────────────────────────────────────────────────

def test_charset_has_no_ambiguous_characters():
    """0, 1, O, I are excluded to avoid confusion — verify that"""
    for char in ["0", "1", "O", "I"]:
        assert char not in CHARSET, f"Ambiguous character '{char}' found in CHARSET"

def test_charset_length():
    """Charset must be exactly 32 characters for the encoding math to work"""
    assert len(CHARSET) == 32


# ─── ENCODING / DECODING TESTS ───────────────────────────────────────────────

def test_encode_decode_roundtrip():
    """Encoding then decoding must return the same user_id and timestamp"""
    user_id = 5
    expiry = (datetime.now() + timedelta(days=365)).timestamp()

    data_bytes = LicenseManager.encode_compact_v2(user_id, expiry)
    decoded_user_id, decoded_expiry = LicenseManager.decode_compact_v2(data_bytes)

    assert decoded_user_id == user_id
    # Timestamps are stored in hours, so allow up to 1 hour of drift
    assert abs(decoded_expiry - int(expiry / 3600) * 3600) < 3600

def test_encode_minimum_user_id():
    """User ID of 1 (the minimum) should encode without error"""
    expiry = (datetime.now() + timedelta(days=30)).timestamp()
    data_bytes = LicenseManager.encode_compact_v2(1, expiry)
    assert len(data_bytes) == 4

def test_encode_maximum_user_id():
    """User ID of 100 (the maximum) should encode without error"""
    expiry = (datetime.now() + timedelta(days=30)).timestamp()
    data_bytes = LicenseManager.encode_compact_v2(100, expiry)
    assert len(data_bytes) == 4

def test_encode_invalid_user_id_zero():
    """User ID of 0 is out of range — must raise ValueError"""
    expiry = (datetime.now() + timedelta(days=30)).timestamp()
    with pytest.raises(ValueError):
        LicenseManager.encode_compact_v2(0, expiry)

def test_encode_invalid_user_id_over_max():
    """User ID above MAX_USERS must raise ValueError"""
    expiry = (datetime.now() + timedelta(days=30)).timestamp()
    with pytest.raises(ValueError):
        LicenseManager.encode_compact_v2(MAX_USERS + 1, expiry)


# ─── CHECKSUM TESTS ──────────────────────────────────────────────────────────

def test_checksum_is_3_bytes():
    """Checksum must always return exactly 3 bytes"""
    data = b"\x01\x02\x03\x04"
    checksum = LicenseManager.generate_checksum(data)
    assert len(checksum) == 3

def test_checksum_is_deterministic():
    """Same input must always produce same checksum"""
    data = b"\xAB\xCD\xEF\x01"
    assert LicenseManager.generate_checksum(data) == LicenseManager.generate_checksum(data)

def test_checksum_changes_with_different_data():
    """Different data must produce different checksums"""
    data1 = b"\x01\x00\x00\x00"
    data2 = b"\x02\x00\x00\x00"
    assert LicenseManager.generate_checksum(data1) != LicenseManager.generate_checksum(data2)


# ─── LICENSE VALIDATION TESTS ────────────────────────────────────────────────

def _make_valid_key(user_id=1, days_from_now=365):
    """Helper: generate a real valid license key for testing"""
    expiry = (datetime.now() + timedelta(days=days_from_now)).timestamp()
    data_bytes = LicenseManager.encode_compact_v2(user_id, expiry)
    checksum = LicenseManager.generate_checksum(data_bytes)
    full_bytes = data_bytes + checksum  # 7 bytes total

    # Encode to readable string (same logic as your app's bytes_to_readable)
    result = []
    for byte in full_bytes:
        idx1 = byte % len(CHARSET)
        idx2 = (byte // len(CHARSET)) % len(CHARSET)
        result.append(CHARSET[idx1])
        result.append(CHARSET[idx2])

    return ''.join(result[:16])


def test_valid_license_key_passes():
    """A freshly generated key should pass validation"""
    key = _make_valid_key(user_id=1, days_from_now=365)
    is_valid, user_id, expiry_date, message = LicenseManager.validate_license_key(key)
    assert is_valid is True
    assert user_id == 1

def test_expired_license_key_fails():
    """A key with a past expiry date should fail"""
    key = _make_valid_key(user_id=1, days_from_now=-1)  # expired yesterday
    is_valid, user_id, expiry_date, message = LicenseManager.validate_license_key(key)
    assert is_valid is False
    assert "expired" in message.lower()

def test_wrong_length_key_fails():
    """Keys that are not 16 characters should fail immediately"""
    is_valid, _, _, message = LicenseManager.validate_license_key("TOOSHORT")
    assert is_valid is False
    assert "length" in message.lower()

def test_tampered_key_fails():
    """Changing one character of a valid key should break the checksum"""
    key = _make_valid_key()
    tampered = key[:-1] + ("A" if key[-1] != "A" else "B")
    is_valid, _, _, message = LicenseManager.validate_license_key(tampered)
    assert is_valid is False

def test_key_with_dashes_is_accepted():
    """Keys formatted as XXXX-XXXX-XXXX-XXXX should still validate"""
    key = _make_valid_key()
    formatted = f"{key[0:4]}-{key[4:8]}-{key[8:12]}-{key[12:16]}"
    is_valid, _, _, _ = LicenseManager.validate_license_key(formatted)
    assert is_valid is True

def test_lowercase_key_is_accepted():
    """Keys entered in lowercase should still validate (app uppercases them)"""
    key = _make_valid_key().lower()
    is_valid, _, _, _ = LicenseManager.validate_license_key(key)
    assert is_valid is True