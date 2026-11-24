// Bulk reset Benda Uji → "belum_test" for many dockets
(async function bulkResetDockets() {
  // ===== CONFIG =====
  const DOCKETS = ['DOCKET/PTB/10-2026/11783','DOCKET/PTB/10-2026/11791',];
  const DRY_RUN = false;               // true = only preview, false = actually save
  const DELAY_MS = 300;               // pause between dockets
  const SKIP_ALREADY_RESET = false;   // true = skip lines already "belum_test"
  const OVERRIDE = {benda_uji_id: 3, percentage: 65.0};                // optional, e.g. { benda_uji_id: 3, percentage: 65.0 } forces the same value on all lines
  // ==================

  const PARENT_MODEL = 'schedule.truck.mixer.benda.uji';
  const LINE_MODEL = 'schedule.truck.mixer.benda.uji.line';
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  const rpc = async (model, method, args, kwargs = {}) => {
    const resp = await fetch(`/web/dataset/call_kw/${model}/${method}`, {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        jsonrpc: '2.0',
        method: 'call',
        params: { model, method, args, kwargs },
        id: Math.floor(Math.random() * 100000)
      })
    });
    const data = await resp.json();
    if (data.error) {
      throw new Error(data.error.data?.message || data.error.message);
    }
    return data.result;
  };

  // Build the per-line payload (same fields as your manual script)
  const buildLineCommand = (line) => {
    const bendaUjiId = Array.isArray(line.benda_uji_id) ? line.benda_uji_id[0] : line.benda_uji_id;
    return [1, line.id, {
      benda_uji_id: bendaUjiId,
      percentage: line.percentage,
      hasil_test_berat_kg: {},
      hasil_test_kg_cm: {},
      hasil_test_kn: {},
      hasil_test_mpa: {},
      jenis_retakan: {},
      state_line: 'rencana',
      real_test: 0,
      realisasi_tanggal_test: {},
      hasil_test: 'belum_test',
      testing_method: 'manual',
      umur_test_real: 7,
      ...OVERRIDE
    }];
  };

  const dockets = [...new Set(DOCKETS.map(d => d.trim()).filter(Boolean))];
  const results = [];

  console.log(`🚀 Start: ${dockets.length} docket | DRY_RUN = ${DRY_RUN}`);

  for (let i = 0; i < dockets.length; i++) {
    const docket = dockets[i];
    const row = { docket, parentId: null, lines: 0, status: '', note: '' };
    console.log(`\n[${i + 1}/${dockets.length}] 🔍 ${docket}`);

    try {
      // Step 1: find parent ID
      const found = await rpc(PARENT_MODEL, 'name_search', [docket], { limit: 10 });
      if (!found || found.length === 0) {
        row.status = '❌ NOT FOUND';
        results.push(row);
        continue;
      }

      let match = found.find(f => f[1] === docket);
      if (!match) {
        if (found.length === 1) {
          match = found[0];
        } else {
          row.status = '⚠️ AMBIGUOUS';
          row.note = found.map(f => f[1]).join(' | ');
          results.push(row);
          continue;
        }
      }
      row.parentId = match[0];

      // Step 2: read parent
      const [parent] = await rpc(PARENT_MODEL, 'read', [[row.parentId]], {
        fields: ['id', 'display_name', 'state', 'benda_uji_ids']
      });

      if (!parent.benda_uji_ids.length) {
        row.status = '⚠️ NO LINES';
        results.push(row);
        continue;
      }

      // Step 3: read lines
      let lines = await rpc(LINE_MODEL, 'read', [parent.benda_uji_ids], {
        fields: ['id', 'benda_uji_id', 'percentage', 'hasil_test', 'state_line']
      });

      if (SKIP_ALREADY_RESET) {
        lines = lines.filter(l => l.hasil_test !== 'belum_test');
      }
      if (!lines.length) {
        row.status = '⏭️ SKIPPED';
        row.note = 'all lines already belum_test';
        results.push(row);
        continue;
      }

      const badLines = lines.filter(l => !l.benda_uji_id);
      if (badLines.length && !OVERRIDE.benda_uji_id) {
        row.status = '⚠️ MISSING benda_uji_id';
        row.note = 'line IDs: ' + badLines.map(l => l.id).join(',');
        results.push(row);
        continue;
      }

      row.lines = lines.length;
      const commands = lines.map(buildLineCommand);

      // Step 4: web_save
      if (DRY_RUN) {
        row.status = '🧪 DRY RUN';
        row.note = 'line IDs: ' + lines.map(l => l.id).join(',');
        console.log('   would update:', row.note);
      } else {
        await rpc(PARENT_MODEL, 'web_save', [
          [row.parentId],
          { benda_uji_ids: commands },
          {}
        ]);
        row.status = '✅ UPDATED';
      }
    } catch (err) {
      row.status = '❌ ERROR';
      row.note = err.message;
      console.error('   ', err.message);
    }

    results.push(row);
    if (i < dockets.length - 1) await sleep(DELAY_MS);
  }

  console.log('\n' + '='.repeat(70));
  console.table(results);
  const failed = results.filter(r => r.status.startsWith('❌') || r.status.startsWith('⚠️'));
  console.log(`Done. OK/preview: ${results.length - failed.length} | Need attention: ${failed.length}`);
  if (DRY_RUN) console.log('👉 This was a DRY RUN. Set DRY_RUN = false to actually save.');

  window.lastBulkResult = results;
  return results;
})();