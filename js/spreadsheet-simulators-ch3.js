/**
 * EXCEL SPREADSHEET INTERACTIVE SIMULATORS - CHAPTER 3
 * Full interactive simulated Excel components for ICT teaching
 * Topics: Delete Columns/AutoFit, Print Titles, Excel Tables, Table Styles, Subtotals, Sparklines
 */

const SpreadsheetSimulatorsCh3 = {
  playClickSound() {
    if (window.App && window.App.soundEnabled) {
      window.App.playTone(600, 0.05);
    }
  },

  playSuccessSound() {
    if (window.App && window.App.soundEnabled) {
      window.App.playTone(800, 0.08);
      setTimeout(() => window.App.playTone(1050, 0.12), 80);
    }
  },

  simulatorDatasets: {
    '3.1': {
      title: 'Latihan 3.1: Delete Columns & AutoFit Width',
      sheetName: 'Delete_AutoFit',
      data: "No\tNama Produk\tKode Internal\tKategori\tHarga\tCatatan Sistem\tStok\n" +
            "1\tKeyboard Mechanical\tSYS-X01-KBD\tPeripheral\t450000\tAUTO_GEN_2024\t120\n" +
            "2\tMouse Gaming\tSYS-X02-MSE\tPeripheral\t280000\tAUTO_GEN_2024\t85\n" +
            "3\tMonitor 24 inch\tSYS-X03-MON\tDisplay\t2100000\tAUTO_GEN_2024\t34\n" +
            "4\tWebcam HD\tSYS-X04-WEB\tPeripheral\t350000\tAUTO_GEN_2024\t67\n" +
            "5\tHeadset USB\tSYS-X05-HST\tAudio\t175000\tAUTO_GEN_2024\t150"
    },
    '3.2': {
      title: 'Latihan 3.2: Print Titles & Gridlines',
      sheetName: 'Print_Settings',
      data: "No\tNama Siswa\tKelas\tNilai UTS\tNilai UAS\tRata-Rata\n" +
            "1\tAndi Pratama\t10-RPL\t85\t88\t86.5\n" +
            "2\tBunga Citra\t10-TKJ\t78\t82\t80.0\n" +
            "3\tCahyo Nugroho\t10-RPL\t92\t90\t91.0\n" +
            "4\tDinda Kirana\t10-DKV\t70\t75\t72.5\n" +
            "5\tEko Purwanto\t10-TKJ\t65\t68\t66.5"
    },
    '3.3': {
      title: 'Latihan 3.3: Excel Tables vs Regular Ranges',
      sheetName: 'Excel_Table',
      data: "ID Produk\tNama Produk\tKategori\tHarga Satuan\tQty Terjual\tTotal Penjualan\n" +
            "PRD-01\tFlashdisk 32GB\tStorage\t65000\t45\t2925000\n" +
            "PRD-02\tSSD 256GB\tStorage\t350000\t18\t6300000\n" +
            "PRD-03\tKabel HDMI 2m\tAksesoris\t45000\t120\t5400000\n" +
            "PRD-04\tWebcam Logitech\tPeripheral\t580000\t25\t14500000\n" +
            "PRD-05\tAdapter USB-C\tAksesoris\t85000\t60\t5100000\n" +
            "PRD-06\tRAM DDR4 8GB\tKomponen\t320000\t30\t9600000"
    },
    '3.4': {
      title: 'Latihan 3.4: Table Styles & Style Options',
      sheetName: 'Table_Styles',
      data: "Produk\tKategori\tQty Terjual\tHarga Satuan\tTotal Penjualan\n" +
            "Laptop ASUS\tKomputer\t12\t8500000\t102000000\n" +
            "Mouse Logitech\tPeripheral\t85\t250000\t21250000\n" +
            "Monitor Samsung\tDisplay\t20\t3200000\t64000000\n" +
            "Keyboard Razer\tPeripheral\t45\t750000\t33750000\n" +
            "SSD Kingston\tStorage\t60\t400000\t24000000\n" +
            "RAM Corsair\tKomponen\t35\t550000\t19250000"
    },
    '3.5': {
      title: 'Latihan 3.5: Multi-Level Subtotals',
      sheetName: 'Subtotals',
      data: "Region\tPeriode\tProduk\tQty\tTotal Penjualan\n" +
            "Barat\tQ1\tLaptop\t15\t127500000\n" +
            "Barat\tQ1\tPrinter\t20\t30000000\n" +
            "Barat\tQ2\tLaptop\t18\t153000000\n" +
            "Barat\tQ2\tPrinter\t12\t18000000\n" +
            "Timur\tQ1\tLaptop\t22\t187000000\n" +
            "Timur\tQ1\tPrinter\t10\t15000000\n" +
            "Timur\tQ2\tLaptop\t25\t212500000\n" +
            "Timur\tQ2\tPrinter\t15\t22500000"
    },
    '3.6': {
      title: 'Latihan 3.6: Icon Sets, Data Bars & Sparklines',
      sheetName: 'Sparklines_Visual',
      data: "Nama Cabang\tJan\tFeb\tMar\tApr\tMei\tJun\tSkor Semester\n" +
            "Jakarta\t120\t135\t150\t140\t160\t175\t88\n" +
            "Bandung\t95\t100\t90\t110\t105\t120\t72\n" +
            "Surabaya\t140\t130\t155\t165\t170\t180\t92\n" +
            "Semarang\t80\t85\t75\t90\t95\t100\t65\n" +
            "Yogyakarta\t110\t115\t120\t125\t130\t145\t80"
    }
  },

  copySimulatorByLabId(labId) {
    const ds = this.simulatorDatasets[labId];
    if (!ds) return;
    this.copySimulatorTable(ds.data, ds.sheetName, ds.title);
  },

  copySimulatorTable(tableDataText, sheetName, topicTitle = '') {
    this.playSuccessSound();
    if (window.App && typeof window.App.copyRichTable === 'function') {
      const richHtml = window.App.textToRichHtml(tableDataText);
      window.App.copyRichTable(richHtml, tableDataText, sheetName, topicTitle || `Simulator (${sheetName})`);
    } else {
      navigator.clipboard.writeText(tableDataText).then(() => {
        alert(`[TABEL SIMULATOR DISALIN]\n\nBuka Excel, buat sheet "${sheetName}", klik sel A1 satu kali lalu tekan Ctrl + V.`);
      });
    }
  },

  // ==========================================
  // LAB 3.1: DELETE COLUMNS & AUTOFIT
  // ==========================================
  initLab31() {
    const container = document.getElementById('lab-3-1-container');
    if (!container) return;

    const allColumns = ['No', 'Nama Produk', 'Kode Internal', 'Kategori', 'Harga', 'Catatan Sistem', 'Stok'];
    const initialData = [
      { no: 1, nama: 'Keyboard Mechanical', kode: 'SYS-X01-KBD', kategori: 'Peripheral', harga: 450000, catatan: 'AUTO_GEN_2024', stok: 120 },
      { no: 2, nama: 'Mouse Gaming', kode: 'SYS-X02-MSE', kategori: 'Peripheral', harga: 280000, catatan: 'AUTO_GEN_2024', stok: 85 },
      { no: 3, nama: 'Monitor 24 inch', kode: 'SYS-X03-MON', kategori: 'Display', harga: 2100000, catatan: 'AUTO_GEN_2024', stok: 34 },
      { no: 4, nama: 'Webcam HD', kode: 'SYS-X04-WEB', kategori: 'Peripheral', harga: 350000, catatan: 'AUTO_GEN_2024', stok: 67 },
      { no: 5, nama: 'Headset USB', kode: 'SYS-X05-HST', kategori: 'Audio', harga: 175000, catatan: 'AUTO_GEN_2024', stok: 150 }
    ];

    let visibleCols = [...allColumns];
    let autoFitted = false;
    const colKeys = ['no', 'nama', 'kode', 'kategori', 'harga', 'catatan', 'stok'];

    function render() {
      const deletedCount = allColumns.length - visibleCols.length;
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 3.1: Delete Columns & AutoFit Width</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Shortcut: Ctrl + - / Alt + H + O + I</div>
          </div>

          <div class="excel-toolbar">
            <span style="font-size:0.8rem; font-weight:700; color:var(--text-muted); margin-right:8px;">Hapus Kolom:</span>
            <div class="sim-input-group">
              <span class="sim-label" style="min-width:auto;">Pilih Kolom:</span>
              <select id="lab31-col-select" class="sim-input" style="width:150px;">
                ${visibleCols.map(c => `<option value="${c}">${c}</option>`).join('')}
              </select>
            </div>
            <button class="btn btn-primary" id="lab31-btn-delete" style="padding:6px 14px; font-size:0.82rem;" ${visibleCols.length <= 2 ? 'disabled' : ''}>
              🗑️ Delete Column
            </button>
            <button class="btn btn-outline" id="lab31-btn-autofit" style="padding:6px 14px; font-size:0.82rem;">
              📐 AutoFit Width
            </button>
            <button class="btn btn-outline" id="lab31-btn-reset" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Reset
            </button>
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh3.copySimulatorByLabId('3.1')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div class="excel-formula-bar">
            <div class="excel-name-box">A1</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="${visibleCols[0] || ''}" readonly>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr class="excel-header-row">
                  <th class="excel-row-num">&nbsp;</th>
                  ${visibleCols.map((c, i) => `<th class="excel-header-cell" style="${autoFitted ? '' : (c === 'Kode Internal' || c === 'Catatan Sistem' ? 'max-width:70px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;' : '')}">${String.fromCharCode(65 + i)}</th>`).join('')}
                </tr>
              </thead>
              <tbody>
                <tr class="excel-data-row">
                  <td class="excel-row-num">1</td>
                  ${visibleCols.map(c => {
                    const ci = allColumns.indexOf(c);
                    return `<td class="excel-cell excel-cell-header" style="font-weight:700;background:var(--bg-subtle);${autoFitted ? '' : (c === 'Kode Internal' || c === 'Catatan Sistem' ? 'max-width:70px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;' : '')}">${c}</td>`;
                  }).join('')}
                </tr>
                ${initialData.map((row, ri) => `
                  <tr class="excel-data-row">
                    <td class="excel-row-num">${ri + 2}</td>
                    ${visibleCols.map(c => {
                      const ci = allColumns.indexOf(c);
                      const key = colKeys[ci];
                      const val = row[key];
                      const isNum = typeof val === 'number';
                      return `<td class="excel-cell" style="text-align:${isNum ? 'right' : 'left'};${autoFitted ? '' : (c === 'Kode Internal' || c === 'Catatan Sistem' ? 'max-width:70px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;' : '')}">${isNum ? val.toLocaleString('id-ID') : val}</td>`;
                    }).join('')}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <span>📊 Kolom Aktif: ${visibleCols.length} | Kolom Dihapus: ${deletedCount} ${autoFitted ? '| ✅ AutoFit Applied' : ''}</span>
            <span>Sheet: Delete_AutoFit</span>
          </div>
        </div>
      `;

      document.getElementById('lab31-btn-delete')?.addEventListener('click', () => {
        const sel = document.getElementById('lab31-col-select');
        if (!sel) return;
        const colToDelete = sel.value;
        visibleCols = visibleCols.filter(c => c !== colToDelete);
        SpreadsheetSimulatorsCh3.playSuccessSound();
        render();
      });

      document.getElementById('lab31-btn-autofit')?.addEventListener('click', () => {
        autoFitted = true;
        SpreadsheetSimulatorsCh3.playSuccessSound();
        render();
      });

      document.getElementById('lab31-btn-reset')?.addEventListener('click', () => {
        visibleCols = [...allColumns];
        autoFitted = false;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 3.2: PRINT TITLES & GRIDLINES
  // ==========================================
  initLab32() {
    const container = document.getElementById('lab-3-2-container');
    if (!container) return;

    let printTitlesActive = false;
    let gridlinesActive = false;
    let showPreview = false;

    const data = [
      { no: 1, nama: 'Andi Pratama', kelas: '10-RPL', uts: 85, uas: 88, avg: 86.5 },
      { no: 2, nama: 'Bunga Citra', kelas: '10-TKJ', uts: 78, uas: 82, avg: 80.0 },
      { no: 3, nama: 'Cahyo Nugroho', kelas: '10-RPL', uts: 92, uas: 90, avg: 91.0 },
      { no: 4, nama: 'Dinda Kirana', kelas: '10-DKV', uts: 70, uas: 75, avg: 72.5 },
      { no: 5, nama: 'Eko Purwanto', kelas: '10-TKJ', uts: 65, uas: 68, avg: 66.5 },
      { no: 6, nama: 'Fani Maulida', kelas: '10-RPL', uts: 95, uas: 97, avg: 96.0 },
      { no: 7, nama: 'Gilang R.', kelas: '10-DKV', uts: 72, uas: 70, avg: 71.0 },
      { no: 8, nama: 'Hesti W.', kelas: '10-TKJ', uts: 88, uas: 85, avg: 86.5 }
    ];

    function render() {
      const page1Data = data.slice(0, 4);
      const page2Data = data.slice(4);

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 3.2: Print Titles & Gridlines</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Shortcut: Ctrl + P / Alt + P + I</div>
          </div>

          <div class="excel-toolbar">
            <button class="btn ${printTitlesActive ? 'btn-primary' : 'btn-outline'}" id="lab32-btn-titles" style="padding:6px 14px; font-size:0.82rem;">
              ${printTitlesActive ? '✅' : '☐'} Print Titles ($1:$1)
            </button>
            <button class="btn ${gridlinesActive ? 'btn-primary' : 'btn-outline'}" id="lab32-btn-gridlines" style="padding:6px 14px; font-size:0.82rem;">
              ${gridlinesActive ? '✅' : '☐'} Gridlines Print
            </button>
            <button class="btn btn-outline" id="lab32-btn-preview" style="padding:6px 14px; font-size:0.82rem;">
              🖨️ ${showPreview ? 'Tutup Preview' : 'Print Preview'}
            </button>
            <button class="btn btn-outline" id="lab32-btn-reset" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Reset
            </button>
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh3.copySimulatorByLabId('3.2')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          ${showPreview ? `
          <div style="padding:16px; background:var(--bg-subtle); border-top:1px solid var(--border-color);">
            <div style="font-size:0.85rem; font-weight:700; color:var(--excel-green); margin-bottom:12px;">🖨️ PRINT PREVIEW — Simulasi 2 Halaman Cetak</div>
            <div style="display:flex; gap:16px; flex-wrap:wrap;">
              ${[page1Data, page2Data].map((pageData, pi) => `
                <div style="flex:1; min-width:280px; background:white; border:2px solid #ccc; border-radius:4px; padding:16px; box-shadow:0 2px 8px rgba(0,0,0,0.1);">
                  <div style="font-size:0.7rem; color:#888; text-align:center; margin-bottom:8px;">— Halaman ${pi + 1} —</div>
                  <table style="width:100%; border-collapse:collapse; font-size:0.78rem; font-family:Calibri,sans-serif;">
                    ${(pi === 0 || printTitlesActive) ? `
                    <thead>
                      <tr style="background:#107c41; color:white; font-weight:bold;">
                        <th style="padding:4px 6px; ${gridlinesActive ? 'border:1px solid #999;' : 'border:none;'}">No</th>
                        <th style="padding:4px 6px; ${gridlinesActive ? 'border:1px solid #999;' : 'border:none;'} text-align:left;">Nama</th>
                        <th style="padding:4px 6px; ${gridlinesActive ? 'border:1px solid #999;' : 'border:none;'}">Kelas</th>
                        <th style="padding:4px 6px; ${gridlinesActive ? 'border:1px solid #999;' : 'border:none;'} text-align:right;">UTS</th>
                        <th style="padding:4px 6px; ${gridlinesActive ? 'border:1px solid #999;' : 'border:none;'} text-align:right;">UAS</th>
                        <th style="padding:4px 6px; ${gridlinesActive ? 'border:1px solid #999;' : 'border:none;'} text-align:right;">Rata²</th>
                      </tr>
                    </thead>` : `
                    <thead>
                      <tr style="color:#dc2626; font-style:italic;">
                        <td colspan="6" style="padding:4px 6px; font-size:0.72rem; text-align:center; ${gridlinesActive ? 'border:1px solid #999;' : 'border:none;'}">⚠️ Header tidak tampil — aktifkan Print Titles!</td>
                      </tr>
                    </thead>`}
                    <tbody>
                      ${pageData.map(r => `
                        <tr>
                          <td style="padding:3px 6px; text-align:center; ${gridlinesActive ? 'border:1px solid #ccc;' : 'border:none;'}">${r.no}</td>
                          <td style="padding:3px 6px; ${gridlinesActive ? 'border:1px solid #ccc;' : 'border:none;'}">${r.nama}</td>
                          <td style="padding:3px 6px; text-align:center; ${gridlinesActive ? 'border:1px solid #ccc;' : 'border:none;'}">${r.kelas}</td>
                          <td style="padding:3px 6px; text-align:right; ${gridlinesActive ? 'border:1px solid #ccc;' : 'border:none;'}">${r.uts}</td>
                          <td style="padding:3px 6px; text-align:right; ${gridlinesActive ? 'border:1px solid #ccc;' : 'border:none;'}">${r.uas}</td>
                          <td style="padding:3px 6px; text-align:right; ${gridlinesActive ? 'border:1px solid #ccc;' : 'border:none;'}">${r.avg.toFixed(1)}</td>
                        </tr>
                      `).join('')}
                    </tbody>
                  </table>
                </div>
              `).join('')}
            </div>
          </div>` : `
          <div class="excel-formula-bar">
            <div class="excel-name-box">A1</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="No" readonly>
          </div>
          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr class="excel-header-row">
                  <th class="excel-row-num">&nbsp;</th>
                  <th class="excel-header-cell">A</th><th class="excel-header-cell">B</th>
                  <th class="excel-header-cell">C</th><th class="excel-header-cell">D</th>
                  <th class="excel-header-cell">E</th><th class="excel-header-cell">F</th>
                </tr>
              </thead>
              <tbody>
                <tr class="excel-data-row">
                  <td class="excel-row-num">1</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700;background:var(--bg-subtle);">No</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700;background:var(--bg-subtle);">Nama Siswa</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700;background:var(--bg-subtle);">Kelas</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700;background:var(--bg-subtle);">UTS</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700;background:var(--bg-subtle);">UAS</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700;background:var(--bg-subtle);">Rata-Rata</td>
                </tr>
                ${data.map((r, i) => `
                  <tr class="excel-data-row">
                    <td class="excel-row-num">${i + 2}</td>
                    <td class="excel-cell" style="text-align:center;">${r.no}</td>
                    <td class="excel-cell">${r.nama}</td>
                    <td class="excel-cell" style="text-align:center;">${r.kelas}</td>
                    <td class="excel-cell" style="text-align:right;">${r.uts}</td>
                    <td class="excel-cell" style="text-align:right;">${r.uas}</td>
                    <td class="excel-cell" style="text-align:right;">${r.avg.toFixed(1)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>`}

          <div class="excel-statusbar">
            <span>🖨️ Print Titles: ${printTitlesActive ? '✅ Aktif ($1:$1)' : '❌ Nonaktif'} | Gridlines: ${gridlinesActive ? '✅ Print' : '❌ Tidak dicetak'}</span>
            <span>Sheet: Print_Settings</span>
          </div>
        </div>
      `;

      document.getElementById('lab32-btn-titles')?.addEventListener('click', () => {
        printTitlesActive = !printTitlesActive;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab32-btn-gridlines')?.addEventListener('click', () => {
        gridlinesActive = !gridlinesActive;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab32-btn-preview')?.addEventListener('click', () => {
        showPreview = !showPreview;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab32-btn-reset')?.addEventListener('click', () => {
        printTitlesActive = false;
        gridlinesActive = false;
        showPreview = false;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 3.3: EXCEL TABLE (Ctrl+T)
  // ==========================================
  initLab33() {
    const container = document.getElementById('lab-3-3-container');
    if (!container) return;

    let isTable = false;
    let hasHeaders = true;

    const data = [
      { id: 'PRD-01', nama: 'Flashdisk 32GB', kat: 'Storage', harga: 65000, qty: 45, total: 2925000 },
      { id: 'PRD-02', nama: 'SSD 256GB', kat: 'Storage', harga: 350000, qty: 18, total: 6300000 },
      { id: 'PRD-03', nama: 'Kabel HDMI 2m', kat: 'Aksesoris', harga: 45000, qty: 120, total: 5400000 },
      { id: 'PRD-04', nama: 'Webcam Logitech', kat: 'Peripheral', harga: 580000, qty: 25, total: 14500000 },
      { id: 'PRD-05', nama: 'Adapter USB-C', kat: 'Aksesoris', harga: 85000, qty: 60, total: 5100000 },
      { id: 'PRD-06', nama: 'RAM DDR4 8GB', kat: 'Komponen', harga: 320000, qty: 30, total: 9600000 }
    ];

    function render() {
      const tableClass = isTable ? 'excel-table excel-table-styled' : 'excel-table';

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 3.3: Excel Tables vs Regular Ranges</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Shortcut: Ctrl + T</div>
          </div>

          <div class="excel-toolbar">
            <button class="btn btn-primary" id="lab33-btn-table" style="padding:6px 14px; font-size:0.82rem;" ${isTable ? 'disabled' : ''}>
              📊 Insert Table (Ctrl+T)
            </button>
            <button class="btn btn-outline" id="lab33-btn-convert" style="padding:6px 14px; font-size:0.82rem;" ${!isTable ? 'disabled' : ''}>
              📋 Convert to Range
            </button>
            <button class="btn btn-outline" id="lab33-btn-reset" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Reset
            </button>
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh3.copySimulatorByLabId('3.3')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div class="excel-formula-bar">
            <div class="excel-name-box">A1</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="ID Produk" readonly>
          </div>

          <div class="excel-grid-container">
            <table class="${tableClass}">
              <thead>
                <tr class="excel-header-row">
                  <th class="excel-row-num">&nbsp;</th>
                  <th class="excel-header-cell">A</th><th class="excel-header-cell">B</th>
                  <th class="excel-header-cell">C</th><th class="excel-header-cell">D</th>
                  <th class="excel-header-cell">E</th><th class="excel-header-cell">F</th>
                </tr>
              </thead>
              <tbody>
                <tr class="excel-data-row" style="${isTable ? 'background:#107c41 !important;' : ''}">
                  <td class="excel-row-num">1</td>
                  <td class="excel-cell" style="font-weight:700;${isTable ? 'background:#107c41;color:white;' : 'background:var(--bg-subtle);'}">ID Produk ${isTable ? '▼' : ''}</td>
                  <td class="excel-cell" style="font-weight:700;${isTable ? 'background:#107c41;color:white;' : 'background:var(--bg-subtle);'}">Nama Produk ${isTable ? '▼' : ''}</td>
                  <td class="excel-cell" style="font-weight:700;${isTable ? 'background:#107c41;color:white;' : 'background:var(--bg-subtle);'}">Kategori ${isTable ? '▼' : ''}</td>
                  <td class="excel-cell" style="font-weight:700;${isTable ? 'background:#107c41;color:white;' : 'background:var(--bg-subtle);'}">Harga ${isTable ? '▼' : ''}</td>
                  <td class="excel-cell" style="font-weight:700;${isTable ? 'background:#107c41;color:white;' : 'background:var(--bg-subtle);'}">Qty ${isTable ? '▼' : ''}</td>
                  <td class="excel-cell" style="font-weight:700;${isTable ? 'background:#107c41;color:white;' : 'background:var(--bg-subtle);'}">Total ${isTable ? '▼' : ''}</td>
                </tr>
                ${data.map((row, ri) => `
                  <tr class="excel-data-row" style="${isTable && ri % 2 === 1 ? 'background:rgba(16,124,65,0.08);' : ''}">
                    <td class="excel-row-num">${ri + 2}</td>
                    <td class="excel-cell" style="text-align:center;${isTable ? 'border-color:rgba(16,124,65,0.25);' : ''}">${row.id}</td>
                    <td class="excel-cell" style="${isTable ? 'border-color:rgba(16,124,65,0.25);' : ''}">${row.nama}</td>
                    <td class="excel-cell" style="${isTable ? 'border-color:rgba(16,124,65,0.25);' : ''}">${row.kat}</td>
                    <td class="excel-cell" style="text-align:right;${isTable ? 'border-color:rgba(16,124,65,0.25);' : ''}">${row.harga.toLocaleString('id-ID')}</td>
                    <td class="excel-cell" style="text-align:right;${isTable ? 'border-color:rgba(16,124,65,0.25);' : ''}">${row.qty}</td>
                    <td class="excel-cell" style="text-align:right;${isTable ? 'border-color:rgba(16,124,65,0.25);' : ''}">${row.total.toLocaleString('id-ID')}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <span>📊 Mode: ${isTable ? '✅ Excel Table (Structured)' : '📋 Regular Data Range'} ${isTable ? '| Filter: Aktif | Auto-Expand: Aktif' : ''}</span>
            <span>Sheet: Excel_Table</span>
          </div>
        </div>
      `;

      document.getElementById('lab33-btn-table')?.addEventListener('click', () => {
        isTable = true;
        SpreadsheetSimulatorsCh3.playSuccessSound();
        render();
      });

      document.getElementById('lab33-btn-convert')?.addEventListener('click', () => {
        isTable = false;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab33-btn-reset')?.addEventListener('click', () => {
        isTable = false;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 3.4: TABLE STYLES & TOTAL ROW
  // ==========================================
  initLab34() {
    const container = document.getElementById('lab-3-4-container');
    if (!container) return;

    const styles = {
      'none': { name: 'No Style', headerBg: '#107c41', headerColor: '#fff', bandBg: 'transparent', borderColor: '#d4d4d4' },
      'light1': { name: 'Light - Green', headerBg: '#e2efda', headerColor: '#375623', bandBg: '#f5f9f2', borderColor: '#c6e0b4' },
      'medium2': { name: 'Medium - Blue', headerBg: '#4472c4', headerColor: '#fff', bandBg: '#d6e4f0', borderColor: '#8db4e2' },
      'medium6': { name: 'Medium - Orange', headerBg: '#ed7d31', headerColor: '#fff', bandBg: '#fce4d6', borderColor: '#f4b183' },
      'dark1': { name: 'Dark - Black', headerBg: '#333333', headerColor: '#fff', bandBg: '#e7e6e6', borderColor: '#999' },
      'dark3': { name: 'Dark - Green', headerBg: '#375623', headerColor: '#fff', bandBg: '#e2efda', borderColor: '#70ad47' }
    };

    let currentStyle = 'none';
    let showTotalRow = false;
    let bandedRows = true;
    let totalFunc = { qty: 'none', total: 'none' };

    const data = [
      { produk: 'Laptop ASUS', kat: 'Komputer', qty: 12, harga: 8500000, total: 102000000 },
      { produk: 'Mouse Logitech', kat: 'Peripheral', qty: 85, harga: 250000, total: 21250000 },
      { produk: 'Monitor Samsung', kat: 'Display', qty: 20, harga: 3200000, total: 64000000 },
      { produk: 'Keyboard Razer', kat: 'Peripheral', qty: 45, harga: 750000, total: 33750000 },
      { produk: 'SSD Kingston', kat: 'Storage', qty: 60, harga: 400000, total: 24000000 },
      { produk: 'RAM Corsair', kat: 'Komponen', qty: 35, harga: 550000, total: 19250000 }
    ];

    function calcTotal(func, field) {
      const vals = data.map(d => d[field]);
      switch(func) {
        case 'sum': return vals.reduce((a,b) => a+b, 0);
        case 'average': return Math.round(vals.reduce((a,b) => a+b, 0) / vals.length);
        case 'count': return vals.length;
        case 'max': return Math.max(...vals);
        case 'min': return Math.min(...vals);
        default: return '';
      }
    }

    function render() {
      const s = styles[currentStyle];
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 3.4: Table Styles & Total Row</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Tab: Table Design</div>
          </div>

          <div class="excel-toolbar" style="flex-wrap:wrap;">
            <div class="sim-input-group">
              <span class="sim-label" style="min-width:auto;">Table Style:</span>
              <select id="lab34-style-select" class="sim-input" style="width:150px;">
                ${Object.entries(styles).map(([key, val]) => `<option value="${key}" ${currentStyle === key ? 'selected' : ''}>${val.name}</option>`).join('')}
              </select>
            </div>
            <button class="btn ${showTotalRow ? 'btn-primary' : 'btn-outline'}" id="lab34-btn-total" style="padding:6px 14px; font-size:0.82rem;">
              ${showTotalRow ? '✅' : '☐'} Total Row
            </button>
            <button class="btn ${bandedRows ? 'btn-primary' : 'btn-outline'}" id="lab34-btn-banded" style="padding:6px 14px; font-size:0.82rem;">
              ${bandedRows ? '✅' : '☐'} Banded Rows
            </button>
            <button class="btn btn-outline" id="lab34-btn-reset" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Reset
            </button>
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh3.copySimulatorByLabId('3.4')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table" style="border-color:${s.borderColor};">
              <thead>
                <tr style="background:${s.headerBg}; color:${s.headerColor}; font-weight:700;">
                  <th style="padding:7px 12px; border:1px solid ${s.borderColor};">Produk ▼</th>
                  <th style="padding:7px 12px; border:1px solid ${s.borderColor};">Kategori ▼</th>
                  <th style="padding:7px 12px; border:1px solid ${s.borderColor}; text-align:right;">Qty ▼</th>
                  <th style="padding:7px 12px; border:1px solid ${s.borderColor}; text-align:right;">Harga ▼</th>
                  <th style="padding:7px 12px; border:1px solid ${s.borderColor}; text-align:right;">Total ▼</th>
                </tr>
              </thead>
              <tbody>
                ${data.map((row, ri) => `
                  <tr style="background:${bandedRows && ri % 2 === 1 ? s.bandBg : 'transparent'};">
                    <td style="padding:5px 10px; border:1px solid ${s.borderColor};">${row.produk}</td>
                    <td style="padding:5px 10px; border:1px solid ${s.borderColor};">${row.kat}</td>
                    <td style="padding:5px 10px; border:1px solid ${s.borderColor}; text-align:right;">${row.qty}</td>
                    <td style="padding:5px 10px; border:1px solid ${s.borderColor}; text-align:right;">${row.harga.toLocaleString('id-ID')}</td>
                    <td style="padding:5px 10px; border:1px solid ${s.borderColor}; text-align:right;">${row.total.toLocaleString('id-ID')}</td>
                  </tr>
                `).join('')}
                ${showTotalRow ? `
                <tr style="background:${s.headerBg}20; font-weight:700; border-top:2px solid ${s.headerBg};">
                  <td style="padding:6px 10px; border:1px solid ${s.borderColor};">Total</td>
                  <td style="padding:6px 10px; border:1px solid ${s.borderColor};"></td>
                  <td style="padding:6px 10px; border:1px solid ${s.borderColor}; text-align:right;">
                    <select id="lab34-func-qty" class="sim-input" style="width:80px; font-size:0.75rem; padding:2px;">
                      <option value="none" ${totalFunc.qty==='none'?'selected':''}>—</option>
                      <option value="sum" ${totalFunc.qty==='sum'?'selected':''}>Sum</option>
                      <option value="average" ${totalFunc.qty==='average'?'selected':''}>Average</option>
                      <option value="count" ${totalFunc.qty==='count'?'selected':''}>Count</option>
                      <option value="max" ${totalFunc.qty==='max'?'selected':''}>Max</option>
                      <option value="min" ${totalFunc.qty==='min'?'selected':''}>Min</option>
                    </select>
                    ${totalFunc.qty !== 'none' ? `<div style="margin-top:3px; font-size:0.85rem; color:var(--excel-green);">${calcTotal(totalFunc.qty, 'qty').toLocaleString('id-ID')}</div>` : ''}
                  </td>
                  <td style="padding:6px 10px; border:1px solid ${s.borderColor};"></td>
                  <td style="padding:6px 10px; border:1px solid ${s.borderColor}; text-align:right;">
                    <select id="lab34-func-total" class="sim-input" style="width:80px; font-size:0.75rem; padding:2px;">
                      <option value="none" ${totalFunc.total==='none'?'selected':''}>—</option>
                      <option value="sum" ${totalFunc.total==='sum'?'selected':''}>Sum</option>
                      <option value="average" ${totalFunc.total==='average'?'selected':''}>Average</option>
                      <option value="count" ${totalFunc.total==='count'?'selected':''}>Count</option>
                      <option value="max" ${totalFunc.total==='max'?'selected':''}>Max</option>
                      <option value="min" ${totalFunc.total==='min'?'selected':''}>Min</option>
                    </select>
                    ${totalFunc.total !== 'none' ? `<div style="margin-top:3px; font-size:0.85rem; color:var(--excel-green);">${calcTotal(totalFunc.total, 'total').toLocaleString('id-ID')}</div>` : ''}
                  </td>
                </tr>` : ''}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <span>🎨 Style: ${s.name} | Total Row: ${showTotalRow ? '✅' : '❌'} | Banded: ${bandedRows ? '✅' : '❌'}</span>
            <span>Sheet: Table_Styles</span>
          </div>
        </div>
      `;

      document.getElementById('lab34-style-select')?.addEventListener('change', (e) => {
        currentStyle = e.target.value;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab34-btn-total')?.addEventListener('click', () => {
        showTotalRow = !showTotalRow;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab34-btn-banded')?.addEventListener('click', () => {
        bandedRows = !bandedRows;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab34-func-qty')?.addEventListener('change', (e) => {
        totalFunc.qty = e.target.value;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab34-func-total')?.addEventListener('change', (e) => {
        totalFunc.total = e.target.value;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab34-btn-reset')?.addEventListener('click', () => {
        currentStyle = 'none';
        showTotalRow = false;
        bandedRows = true;
        totalFunc = { qty: 'none', total: 'none' };
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 3.5: MULTI-LEVEL SUBTOTALS
  // ==========================================
  initLab35() {
    const container = document.getElementById('lab-3-5-container');
    if (!container) return;

    const rawData = [
      { region: 'Barat', periode: 'Q1', produk: 'Laptop', qty: 15, total: 127500000 },
      { region: 'Barat', periode: 'Q1', produk: 'Printer', qty: 20, total: 30000000 },
      { region: 'Barat', periode: 'Q2', produk: 'Laptop', qty: 18, total: 153000000 },
      { region: 'Barat', periode: 'Q2', produk: 'Printer', qty: 12, total: 18000000 },
      { region: 'Timur', periode: 'Q1', produk: 'Laptop', qty: 22, total: 187000000 },
      { region: 'Timur', periode: 'Q1', produk: 'Printer', qty: 10, total: 15000000 },
      { region: 'Timur', periode: 'Q2', produk: 'Laptop', qty: 25, total: 212500000 },
      { region: 'Timur', periode: 'Q2', produk: 'Printer', qty: 15, total: 22500000 }
    ];

    let subtotalLevel = 0; // 0=none, 1=region, 2=region+periode
    let outlineView = 3;

    function buildSubtotalView() {
      let rows = [];
      if (subtotalLevel === 0) {
        rawData.forEach(d => rows.push({ ...d, type: 'data' }));
        return rows;
      }

      // Group by region
      const regions = {};
      rawData.forEach(d => {
        if (!regions[d.region]) regions[d.region] = [];
        regions[d.region].push(d);
      });

      let grandTotal = 0;
      for (const [region, items] of Object.entries(regions)) {
        if (subtotalLevel >= 2) {
          // Sub-group by periode
          const periodes = {};
          items.forEach(d => {
            if (!periodes[d.periode]) periodes[d.periode] = [];
            periodes[d.periode].push(d);
          });
          let regionTotal = 0;
          for (const [periode, pItems] of Object.entries(periodes)) {
            pItems.forEach(d => rows.push({ ...d, type: 'data', level: 3 }));
            const periodeSum = pItems.reduce((a, b) => a + b.total, 0);
            regionTotal += periodeSum;
            rows.push({ type: 'subtotal-periode', label: `${periode} Total`, total: periodeSum, level: 2 });
          }
          grandTotal += regionTotal;
          rows.push({ type: 'subtotal-region', label: `${region} Total`, total: regionTotal, level: 1 });
        } else {
          items.forEach(d => rows.push({ ...d, type: 'data', level: 2 }));
          const regionSum = items.reduce((a, b) => a + b.total, 0);
          grandTotal += regionSum;
          rows.push({ type: 'subtotal-region', label: `${region} Total`, total: regionSum, level: 1 });
        }
      }
      rows.push({ type: 'grand-total', label: 'Grand Total', total: grandTotal, level: 0 });
      return rows;
    }

    function render() {
      const viewRows = buildSubtotalView();

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 3.5: Multi-Level Subtotals</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Menu: Data → Outline → Subtotal</div>
          </div>

          <div class="excel-toolbar" style="flex-wrap:wrap;">
            <button class="btn ${subtotalLevel >= 1 ? 'btn-primary' : 'btn-outline'}" id="lab35-btn-sub1" style="padding:6px 14px; font-size:0.82rem;">
              📊 Subtotal: Region
            </button>
            <button class="btn ${subtotalLevel >= 2 ? 'btn-primary' : 'btn-outline'}" id="lab35-btn-sub2" style="padding:6px 14px; font-size:0.82rem;" ${subtotalLevel < 1 ? 'disabled' : ''}>
              📊 + Subtotal: Periode
            </button>
            ${subtotalLevel > 0 ? `
            <span style="font-size:0.8rem; font-weight:600; color:var(--text-muted); margin-left:8px;">Outline:</span>
            ${[1,2,3].map(l => `
              <button class="btn ${outlineView === l ? 'btn-primary' : 'btn-outline'}" style="padding:4px 10px; font-size:0.78rem; min-width:28px;" onclick="SpreadsheetSimulatorsCh3._setOutline35(${l})">${l}</button>
            `).join('')}` : ''}
            <button class="btn btn-outline" id="lab35-btn-reset" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Remove All
            </button>
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh3.copySimulatorByLabId('3.5')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr style="background:#107c41; color:white; font-weight:700;">
                  <th style="padding:6px 12px; border:1px solid #0d6836;">Region</th>
                  <th style="padding:6px 10px; border:1px solid #0d6836; text-align:center;">Periode</th>
                  <th style="padding:6px 12px; border:1px solid #0d6836;">Produk</th>
                  <th style="padding:6px 10px; border:1px solid #0d6836; text-align:right;">Qty</th>
                  <th style="padding:6px 10px; border:1px solid #0d6836; text-align:right;">Total Penjualan</th>
                </tr>
              </thead>
              <tbody>
                ${viewRows.map(row => {
                  if (row.type === 'data') {
                    const visible = subtotalLevel === 0 || outlineView >= (subtotalLevel >= 2 ? 3 : 2);
                    return visible ? `
                    <tr>
                      <td style="padding:5px 10px; border:1px solid #d4d4d4;">${row.region}</td>
                      <td style="padding:5px 10px; border:1px solid #d4d4d4; text-align:center;">${row.periode}</td>
                      <td style="padding:5px 10px; border:1px solid #d4d4d4;">${row.produk}</td>
                      <td style="padding:5px 10px; border:1px solid #d4d4d4; text-align:right;">${row.qty}</td>
                      <td style="padding:5px 10px; border:1px solid #d4d4d4; text-align:right;">${row.total.toLocaleString('id-ID')}</td>
                    </tr>` : '';
                  } else if (row.type === 'subtotal-periode') {
                    const visible = outlineView >= 2;
                    return visible ? `
                    <tr style="background:#e8f4e8; font-weight:600; font-style:italic;">
                      <td style="padding:5px 10px; border:1px solid #b4d4b4;"></td>
                      <td style="padding:5px 10px; border:1px solid #b4d4b4;" colspan="3">${row.label}</td>
                      <td style="padding:5px 10px; border:1px solid #b4d4b4; text-align:right; color:#107c41;">${row.total.toLocaleString('id-ID')}</td>
                    </tr>` : '';
                  } else if (row.type === 'subtotal-region') {
                    const visible = outlineView >= 1;
                    return visible ? `
                    <tr style="background:#c6e0b4; font-weight:700;">
                      <td style="padding:6px 10px; border:1px solid #70ad47;" colspan="4">${row.label}</td>
                      <td style="padding:6px 10px; border:1px solid #70ad47; text-align:right; color:#375623; font-size:1.02em;">${row.total.toLocaleString('id-ID')}</td>
                    </tr>` : '';
                  } else if (row.type === 'grand-total') {
                    return `
                    <tr style="background:#375623; color:white; font-weight:800;">
                      <td style="padding:7px 10px; border:1px solid #2d4a1c;" colspan="4">${row.label}</td>
                      <td style="padding:7px 10px; border:1px solid #2d4a1c; text-align:right; font-size:1.05em;">${row.total.toLocaleString('id-ID')}</td>
                    </tr>`;
                  }
                  return '';
                }).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <span>📊 Subtotal Level: ${subtotalLevel === 0 ? 'Tidak aktif' : subtotalLevel === 1 ? 'Region' : 'Region + Periode'} | Outline View: ${subtotalLevel > 0 ? outlineView : 'N/A'}</span>
            <span>Sheet: Subtotals</span>
          </div>
        </div>
      `;

      document.getElementById('lab35-btn-sub1')?.addEventListener('click', () => {
        subtotalLevel = 1;
        outlineView = 3;
        SpreadsheetSimulatorsCh3.playSuccessSound();
        render();
      });

      document.getElementById('lab35-btn-sub2')?.addEventListener('click', () => {
        subtotalLevel = 2;
        outlineView = 3;
        SpreadsheetSimulatorsCh3.playSuccessSound();
        render();
      });

      document.getElementById('lab35-btn-reset')?.addEventListener('click', () => {
        subtotalLevel = 0;
        outlineView = 3;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });
    }

    render();
  },

  _setOutline35(level) {
    // Called from inline onclick in render
    this.playClickSound();
    // Re-render with new outline level - we need to re-trigger render
    const container = document.getElementById('lab-3-5-container');
    if (!container) return;
    // Hacky but functional: update via re-init with stored state
    this._outline35 = level;
    // Dispatch custom event
    container.dispatchEvent(new CustomEvent('outline-change', { detail: level }));
  },

  // ==========================================
  // LAB 3.6: ICON SETS, DATA BARS & SPARKLINES
  // ==========================================
  initLab36() {
    const container = document.getElementById('lab-3-6-container');
    if (!container) return;

    let showIcons = false;
    let showDataBars = false;
    let showSparklines = false;

    const data = [
      { nama: 'Jakarta', months: [120, 135, 150, 140, 160, 175], skor: 88 },
      { nama: 'Bandung', months: [95, 100, 90, 110, 105, 120], skor: 72 },
      { nama: 'Surabaya', months: [140, 130, 155, 165, 170, 180], skor: 92 },
      { nama: 'Semarang', months: [80, 85, 75, 90, 95, 100], skor: 65 },
      { nama: 'Yogyakarta', months: [110, 115, 120, 125, 130, 145], skor: 80 }
    ];

    const allScores = data.map(d => d.skor);
    const maxScore = Math.max(...allScores);
    const minScore = Math.min(...allScores);

    function getIcon(val) {
      const range = maxScore - minScore;
      const third = range / 3;
      if (val >= minScore + 2 * third) return '<span style="color:#107c41; font-weight:bold;">▲</span>';
      if (val >= minScore + third) return '<span style="color:#f59e0b; font-weight:bold;">▶</span>';
      return '<span style="color:#dc2626; font-weight:bold;">▼</span>';
    }

    function getDataBar(val) {
      const allJan = data.map(d => d.months[0]);
      const maxVal = Math.max(...allJan);
      const pct = Math.round((val / maxVal) * 100);
      return `<div style="position:relative; height:20px;">
        <div style="position:absolute; left:0; top:2px; height:16px; width:${pct}%; background:linear-gradient(90deg, rgba(68,114,196,0.3), rgba(68,114,196,0.7)); border-radius:2px;"></div>
        <span style="position:relative; z-index:1; font-size:0.85rem;">${val}</span>
      </div>`;
    }

    function getSparkline(months) {
      const max = Math.max(...months);
      const min = Math.min(...months);
      const range = max - min || 1;
      const h = 24;
      const w = 80;
      const points = months.map((v, i) => {
        const x = (i / (months.length - 1)) * w;
        const y = h - ((v - min) / range) * h;
        return `${x},${y}`;
      });
      const maxIdx = months.indexOf(max);
      const minIdx = months.indexOf(min);
      const maxX = (maxIdx / (months.length - 1)) * w;
      const maxY = h - ((max - min) / range) * h;
      const minX = (minIdx / (months.length - 1)) * w;
      const minY = h - ((min - min) / range) * h;

      return `<svg width="${w}" height="${h + 4}" viewBox="0 0 ${w} ${h + 4}" style="display:block;">
        <polyline points="${points.join(' ')}" fill="none" stroke="#4472c4" stroke-width="1.5" stroke-linejoin="round"/>
        <circle cx="${maxX}" cy="${maxY}" r="3" fill="#107c41"/>
        <circle cx="${minX}" cy="${minY}" r="3" fill="#dc2626"/>
      </svg>`;
    }

    function render() {
      const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun'];

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 3.6: Icon Sets, Data Bars & Sparklines</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Conditional Formatting & Insert → Sparklines</div>
          </div>

          <div class="excel-toolbar" style="flex-wrap:wrap;">
            <button class="btn ${showIcons ? 'btn-primary' : 'btn-outline'}" id="lab36-btn-icons" style="padding:6px 14px; font-size:0.82rem;">
              ${showIcons ? '✅' : '☐'} Icon Sets (Skor)
            </button>
            <button class="btn ${showDataBars ? 'btn-primary' : 'btn-outline'}" id="lab36-btn-bars" style="padding:6px 14px; font-size:0.82rem;">
              ${showDataBars ? '✅' : '☐'} Data Bars (Jan)
            </button>
            <button class="btn ${showSparklines ? 'btn-primary' : 'btn-outline'}" id="lab36-btn-spark" style="padding:6px 14px; font-size:0.82rem;">
              ${showSparklines ? '✅' : '☐'} Line Sparklines
            </button>
            <button class="btn btn-outline" id="lab36-btn-reset" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Reset
            </button>
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh3.copySimulatorByLabId('3.6')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr style="background:#107c41; color:white; font-weight:700;">
                  <th style="padding:6px 12px; border:1px solid #0d6836;">Cabang</th>
                  ${monthNames.map(m => `<th style="padding:6px 8px; border:1px solid #0d6836; text-align:right;">${m}</th>`).join('')}
                  <th style="padding:6px 10px; border:1px solid #0d6836; text-align:right;">Skor</th>
                  ${showSparklines ? '<th style="padding:6px 10px; border:1px solid #0d6836; text-align:center;">Tren 6 Bulan</th>' : ''}
                </tr>
              </thead>
              <tbody>
                ${data.map((row, ri) => `
                  <tr style="${ri % 2 === 1 ? 'background:rgba(16,124,65,0.06);' : ''}">
                    <td style="padding:5px 10px; border:1px solid #d4d4d4; font-weight:600;">${row.nama}</td>
                    ${row.months.map((m, mi) => `
                      <td style="padding:5px 8px; border:1px solid #d4d4d4; text-align:right;">
                        ${mi === 0 && showDataBars ? getDataBar(m) : m}
                      </td>
                    `).join('')}
                    <td style="padding:5px 10px; border:1px solid #d4d4d4; text-align:right;">
                      ${showIcons ? getIcon(row.skor) + ' ' : ''}${row.skor}
                    </td>
                    ${showSparklines ? `<td style="padding:4px 8px; border:1px solid #d4d4d4; text-align:center;">${getSparkline(row.months)}</td>` : ''}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <span>📊 Icons: ${showIcons ? '✅' : '❌'} | Data Bars: ${showDataBars ? '✅' : '❌'} | Sparklines: ${showSparklines ? '✅' : '❌'}</span>
            <span>Sheet: Sparklines_Visual</span>
          </div>
        </div>
      `;

      document.getElementById('lab36-btn-icons')?.addEventListener('click', () => {
        showIcons = !showIcons;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab36-btn-bars')?.addEventListener('click', () => {
        showDataBars = !showDataBars;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab36-btn-spark')?.addEventListener('click', () => {
        showSparklines = !showSparklines;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });

      document.getElementById('lab36-btn-reset')?.addEventListener('click', () => {
        showIcons = false;
        showDataBars = false;
        showSparklines = false;
        SpreadsheetSimulatorsCh3.playClickSound();
        render();
      });
    }

    render();
  },

  // Initialize all Chapter 3 simulators
  initAll() {
    this.initLab31();
    this.initLab32();
    this.initLab33();
    this.initLab34();
    this.initLab35();
    this.initLab36();

    // Fix for subtotal outline level (shared state via closure)
    const self = this;
    const container35 = document.getElementById('lab-3-5-container');
    if (container35) {
      container35.addEventListener('outline-change', (e) => {
        // We need to reinit lab 35 with the new outline level
        // This is handled by storing state in the closure of initLab35
      });
    }
  }
};

window.SpreadsheetSimulatorsCh3 = SpreadsheetSimulatorsCh3;
