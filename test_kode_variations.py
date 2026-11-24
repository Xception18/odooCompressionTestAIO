"""Test variasi nama param untuk kodebendauji"""
import requests
import json

url = "https://rmc.adhimix.web.id/benda_uji/update/"

# Base params (yang pasti work)
base = {
    "bjdt_tgl_test": "2026-10-27",
    "bjdt_beban": 150.50,
    "bjdt_berat": 8.25,
    "bjdt_tipe_retak": "A",
    "bjdt_id": 510850,
    "bjdt_beban_mpa": 25.75,
    "bjdt_beban_kg": "15360",
    "bjdt_umur": 28
}

# Variasi nama param untuk kodebendauji
variations = [
    ("bjdt_kode", "BU-TEST-001"),
    ("bjdt_kodebendauji", "BU-TEST-002"),
    ("bjdt_kode_benda_uji", "BU-TEST-003"),
    ("kodebendauji", "BU-TEST-004"),
    ("kode_benda_uji", "BU-TEST-005"),
    ("bjdt_code", "BU-TEST-006"),
]

print("=" * 70)
print("TEST VARIASI NAMA PARAM untuk kodebendauji")
print("Ganti bjdt_id dengan ID valid untuk test update real!")
print("=" * 70)

for param_name, value in variations:
    payload = {"params": {**base, param_name: value}}
    
    print(f"\n[{param_name}] = {value}")
    try:
        response = requests.post(url, json=payload, timeout=30, verify=False)
        result = response.json()
        
        # Cek apakah param ada di response
        has_param = param_name in result.get("result", {}).get("data", {})
        echo_value = result.get("result", {}).get("data", {}).get(param_name, "NOT FOUND")
        
        print(f"  Status: {response.status_code}")
        print(f"  Param in response: {has_param}")
        print(f"  Echo value: {echo_value}")
        
    except Exception as e:
        print(f"  ERROR: {e}")

print("\n" + "=" * 70)
print("Cek manual di Odoo untuk field kodebendauji pada bjdt_id tersebut")
print("=" * 70)
