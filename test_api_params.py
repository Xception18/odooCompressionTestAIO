"""Test script untuk cek params API benda_uji/update/"""
import requests
import ssl
import json
from datetime import date

# Disable SSL verify (sama seperti daemon)
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

url = "https://rmc.adhimix.web.id/benda_uji/update/"

# Test data (sample dari struktur database)
test_payload_original = {
    "params": {
        "bjdt_tgl_test": "2026-10-27",
        "bjdt_beban": 150.50,
        "bjdt_berat": 8.25,
        "bjdt_tipe_retak": "A",
        "bjdt_id": 12345,
        "bjdt_beban_mpa": 25.75,
        "bjdt_beban_kg": "15360",
        "bjdt_umur": 28
    }
}

# Test dengan tambahan kodebendauji
test_payload_with_kode = {
    "params": {
        "bjdt_tgl_test": "2026-10-27",
        "bjdt_beban": 150.50,
        "bjdt_berat": 8.25,
        "bjdt_tipe_retak": "A",
        "bjdt_id": 12345,
        "bjdt_beban_mpa": 25.75,
        "bjdt_beban_kg": "15360",
        "bjdt_umur": 28,
        "bjdt_kode": "BU-2026-001"  # Tambahan param kodebendauji
    }
}

print("=" * 70)
print("TEST 1: Original params (8 fields)")
print("=" * 70)
print(f"URL: {url}")
print(f"Payload:\n{json.dumps(test_payload_original, indent=2)}")
print()

try:
    response = requests.post(url, json=test_payload_original, timeout=30, verify=False)
    print(f"Status Code: {response.status_code}")
    print(f"Response Headers: {dict(response.headers)}")
    print(f"Response Body: {response.text[:500]}")
except Exception as e:
    print(f"ERROR: {e}")

print("\n" + "=" * 70)
print("TEST 2: With kodebendauji param (9 fields)")
print("=" * 70)
print(f"Payload:\n{json.dumps(test_payload_with_kode, indent=2)}")
print()

try:
    response = requests.post(url, json=test_payload_with_kode, timeout=30, verify=False)
    print(f"Status Code: {response.status_code}")
    print(f"Response Headers: {dict(response.headers)}")
    print(f"Response Body: {response.text[:500]}")
except Exception as e:
    print(f"ERROR: {e}")

print("\n" + "=" * 70)
print("SELESAI")
print("=" * 70)
