// ✅ WORKING: Script akurat pakai name_search
(async function searchDocket(docketNumber = 'DOCKET/PTB/10-2026/11794') {
  console.log('🔍 Searching:', docketNumber);
  
  try {
    // Step 1: name_search untuk dapat ID
    const searchResp = await fetch('/web/dataset/call_kw/schedule.truck.mixer.benda.uji/name_search', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        jsonrpc: '2.0',
        method: 'call',
        params: {
          model: 'schedule.truck.mixer.benda.uji',
          method: 'name_search',
          args: [docketNumber],
          kwargs: {limit: 10}
        },
        id: Math.floor(Math.random() * 1000)
      })
    });
    
    const searchData = await searchResp.json();
    
    if (searchData.error) {
      console.error('❌ Search error:', searchData.error.message);
      return null;
    }
    
    if (!searchData.result || searchData.result.length === 0) {
      console.error('❌ Docket tidak ditemukan:', docketNumber);
      return null;
    }
    
    const parentId = searchData.result[0][0]; // [id, name] format
    console.log('✅ Parent ID:', parentId);
    
    // Step 2: Read full detail
    const readResp = await fetch('/web/dataset/call_kw/schedule.truck.mixer.benda.uji/read', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        jsonrpc: '2.0',
        method: 'call',
        params: {
          model: 'schedule.truck.mixer.benda.uji',
          method: 'read',
          args: [[parentId]],
          kwargs: {
            fields: [
              'id', 'display_name', 'state', 'rel_docket',
              'project_id', 'kontraktor_id', 'product_id',
              'tanggal_cor', 'benda_uji_ids'
            ]
          }
        },
        id: Math.floor(Math.random() * 1000)
      })
    });
    
    const readData = await readResp.json();
    
    if (readData.error) {
      console.error('❌ Read error:', readData.error.message);
      return null;
    }
    
    const parent = readData.result[0];
    const bendaUjiIds = parent.benda_uji_ids;
    
    // Step 3: Read benda_uji details
    if (bendaUjiIds.length > 0) {
      const detailResp = await fetch('/web/dataset/call_kw/schedule.truck.mixer.benda.uji.line/read', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
          jsonrpc: '2.0',
          method: 'call',
          params: {
            model: 'schedule.truck.mixer.benda.uji.line',
            method: 'read',
            args: [bendaUjiIds],
            kwargs: {
              fields: [
                'id', 'nomor_urut', 'kode_benda_uji',
                'benda_uji_id', 'tanggal_test', 'percentage',
                'bentuk_benda_uji', 'tempat_pengetesan'
              ]
            }
          },
          id: Math.floor(Math.random() * 1000)
        })
      });
      
      const detailData = await detailResp.json();
      
      if (detailData.error) {
        console.error('❌ Detail error:', detailData.error.message);
        return null;
      }
      
      const details = detailData.result;
      
      // Output
      console.log('\n' + '='.repeat(70));
      console.log('📋 HASIL PENCARIAN DOCKET');
      console.log('='.repeat(70));
      console.log('Parent ID       :', parent.id);
      console.log('Docket Number   :', parent.display_name);
      console.log('Status          :', parent.state);
      console.log('Project         :', parent.project_id?.[1] || parent.project_id);
      console.log('Kontraktor      :', parent.kontraktor_id);
      console.log('Product         :', parent.product_id?.[1] || parent.product_id);
      console.log('Tanggal Cor     :', parent.tanggal_cor);
      console.log('Total Benda Uji :', details.length);
      console.log('='.repeat(70));
      
      console.log('\n📦 DETAIL BENDA UJI:');
      console.table(details.map(b => ({
        'ID': b.id,
        'No': b.nomor_urut,
        'Umur': b.benda_uji_id?.[1] || b.benda_uji_id,
        'Tanggal Test': b.tanggal_test,
        '%': b.percentage,
        'Bentuk': b.bentuk_benda_uji?.[1] || b.bentuk_benda_uji
      })));
      
      console.log('\n📝 Benda Uji IDs:', JSON.stringify(details.map(b => b.id)));
      
      // Save
      window.lastDocketData = {
        parentId: parent.id,
        docketNumber: parent.display_name,
        state: parent.state,
        bendaUjiIds: details.map(b => b.id),
        bendaUjiDetails: details,
        fullParent: parent
      };
      
      console.log('\n💾 Data tersimpan di: window.lastDocketData');
      
      return window.lastDocketData;
    }
    
  } catch (err) {
    console.error('❌ Exception:', err);
    return null;
  }
})();