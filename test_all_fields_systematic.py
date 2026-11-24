"""Test systematic semua field DB dengan bjdt_id valid"""
import requests
import json
import sys

url = "https://rmc.adhimix.web.id/benda_uji/update/"

# GANTI DENGAN BJDT_ID VALID DARI ODOO
VALID_BJDT_ID = input("Masukkan bjdt_id valid dari Odoo: ").strip()

if not VALID_BJDT_ID or not VALID_BJDT_ID.isdigit():
    print("Error: bjdt_id harus berupa angka")
    sys.exit(1)

VALID_BJDT_ID = int(VALID_BJDT_ID)

# Base params yang sudah terkonfirmasi work
base_params = {
    "bjdt_id": VALID_BJDT_ID,
    "bjdt_tgl_test": "2026-10-27",
    "bjdt_beban": 150.50,
    "bjdt_berat": 8.25,
    "bjdt_tipe_retak": "A",
    "bjdt_beban_mpa": 25.75,
    "bjdt_beban_kg": "15360",
    "bjdt_umur": 28
}

# Field DB yang belum tested dengan variasi nama param
test_fields = [
    # Format: (field_name, [(param_name, test_value), ...])
    ("idalat", [
        ("bjdt_idalat", "ALT-001"),
        ("bjdt_id_alat", "ALT-001"),
        ("idalat", "ALT-001"),
        ("id_alat", "ALT-001"),
    ]),
    ("tglrencanauji", [
        ("bjdt_tgl_rencana", "2026-10-20"),
        ("bjdt_tgl_rencanauji", "2026-10-20"),
        ("bjdt_tglrencanauji", "2026-10-20"),
        ("tgl_rencana_uji", "2026-10-20"),
        ("bjdt_tgl_rencana_uji", "2026-10-20"),
    ]),
    ("tglbendauji", [
        ("bjdt_tgl_buat", "2026-09-29"),
        ("bjdt_tgl_bendauji", "2026-09-29"),
        ("bjdt_tglbendauji", "2026-09-29"),
        ("tgl_benda_uji", "2026-09-29"),
        ("bjdt_tgl_benda_uji", "2026-09-29"),
    ]),
    ("bujnama", [
        ("bjdt_nama", "SILINDER"),
        ("bjdt_bujnama", "SILINDER"),
        ("bujnama", "SILINDER"),
        ("nama_benda_uji", "SILINDER"),
        ("bjdt_nama_benda_uji", "SILINDER"),
    ]),
    ("kodebendauji", [
        ("bjdt_kode", "BU-2026-999"),
        ("bjdt_kodebendauji", "BU-2026-999"),
        ("kodebendauji", "BU-2026-999"),
        ("kode_benda_uji", "BU-2026-999"),
        ("bjdt_kode_benda_uji", "BU-2026-999"),
    ]),
]

print("=" * 80)
print("SYSTEMATIC TEST - Semua Field Database")
print(f"Target bjdt_id: {VALID_BJDT_ID}")
print("=" * 80)
print("\nPETUNJUK:")
print("1. Script akan test setiap field dengan berbagai variasi nama param")
print("2. Setelah test, CEK MANUAL di Odoo untuk field yang berubah")
print("3. Status 200 TIDAK BERARTI field ter-update, hanya terima request")
print("=" * 80)

input("\nTekan ENTER untuk mulai test...")

results = []

for field_name, param_variations in test_fields:
    print(f"\n{'=' * 80}")
    print(f"TESTING FIELD: {field_name}")
    print(f"{'=' * 80}")
    
    for param_name, test_value in param_variations:
        # Merge base params dengan param yang ditest
        payload = {"params": {**base_params, param_name: test_value}}
        
        try:
            response = requests.post(url, json=payload, timeout=30, verify=False)
            result_data = response.json()
            
            # Cek apakah param ada di response
            has_param = param_name in result_data.get("result", {}).get("data", {})
            echo_value = result_data.get("result", {}).get("data", {}).get(param_name, "NOT IN RESPONSE")
            
            status_icon = "✓" if response.status_code == 200 else "✗"
            echo_icon = "📤" if has_param else "❌"
            
            print(f"{status_icon} [{param_name:30s}] = {test_value:20s} | Echo: {echo_icon} {echo_value}")
            
            results.append({
                "field": field_name,
                "param": param_name,
                "value": test_value,
                "status": response.status_code,
                "echoed": has_param,
                "echo_value": echo_value
            })
            
        except Exception as e:
            print(f"✗ [{param_name:30s}] ERROR: {e}")
            results.append({
                "field": field_name,
                "param": param_name,
                "value": test_value,
                "status": "ERROR",
                "echoed": False,
                "echo_value": str(e)
            })

# Summary
print("\n" + "=" * 80)
print("SUMMARY - Params yang di-echo oleh server (200 OK)")
print("=" * 80)

for field_name, param_variations in test_fields:
    field_results = [r for r in results if r["field"] == field_name and r["echoed"]]
    
    print(f"\n{field_name}:")
    if field_results:
        for r in field_results:
            print(f"  ✓ {r['param']}")
    else:
        print(f"  ❌ TIDAK ADA param yang di-echo")

print("\n" + "=" * 80)
print("LANGKAH SELANJUTNYA:")
print("=" * 80)
print(f"1. Buka Odoo ERP")
print(f"2. Cari record dengan bjdt_id = {VALID_BJDT_ID}")
print(f"3. Cek field-field berikut untuk perubahan:")
print(f"   - idalat")
print(f"   - tglrencanauji")
print(f"   - tglbendauji")
print(f"   - bujnama")
print(f"   - kodebendauji")
print(f"4. Catat param mana yang BENAR-BENAR update field")
print("=" * 80)
