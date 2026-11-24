// Edit Benda Uji untuk Docket Menjadi Belum test


fetch('/web/dataset/call_kw/schedule.truck.mixer.benda.uji/web_save', {
  method: 'POST',
  headers: {'Content-Type': 'application/json'},
  body: JSON.stringify({
    jsonrpc: '2.0',
    method: 'call',
    params: {
      model: 'schedule.truck.mixer.benda.uji',
      method: 'web_save',
      args: [
        [123821],  // Parent ID
        {
          benda_uji_ids: [
            [1, 540747, {
              // bentuk_benda_uji: 1, // ID Bentuk Benda Uji 1 = "Silinder 15 x 30 cm", 12 = "Beam 15 x 15 x 60 cm"
              benda_uji_id: 3,           // ID untuk display_name 3 => "7", 6 => 28
              percentage: 65.0,
              // umur_test_real: 28,
              // kode_benda_uji: "HK - JAKON JO / STA 30+300 - 30+400 / CTB 3% / AD PTB", // KODE Benda Uji
              // hasil_test_berat_kg:{},
              // hasil_test_kg_cm: {},
              // hasil_test_kn: {},
              // hasil_test_mpa: {},
              // jenis_retakan: {},
              // state_line: "rencana",
              // real_test: 0,
              // tanggal_test: "2026-11-28", // RENCANA Tanggal Test
              // realisasi_tanggal_test: {}, // REALISASI Tanggal Test
              // hasil_test: "belum_test",
              // testing_method: "manual",
            }],
            [1, 540748, {
              // bentuk_benda_uji: 1, // ID Bentuk Benda Uji 1 = "Silinder 15 x 30 cm", 12 = "Beam 15 x 15 x 60 cm"
              benda_uji_id: 3,           // ID untuk display_name 3 => "7", 6 => 28
              percentage: 65.0,
              // umur_test_real: 28,
              // kode_benda_uji: "HK - JAKON JO / STA 30+300 - 30+400 / CTB 3% / AD PTB", // KODE Benda Uji
              // hasil_test_berat_kg:{},
              // hasil_test_kg_cm: {},
              // hasil_test_kn: {},
              // hasil_test_mpa: {},
              // jenis_retakan: {},
              // state_line: "rencana",
              // real_test: 0,
              // tanggal_test: "2026-11-28", // RENCANA Tanggal Test
              // realisasi_tanggal_test: {}, // REALISASI Tanggal Test
              // hasil_test: "belum_test",
              // testing_method: "manual",
            }],
          ]
        },
        {}  // specification
      ],
      kwargs: {}
    },
    id: 1
  })
}).then(r => r.json()).then(console.log)