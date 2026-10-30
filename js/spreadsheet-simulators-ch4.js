/**
 * EXCEL SPREADSHEET INTERACTIVE SIMULATORS - CHAPTER 4
 * Full interactive simulated Excel components for ICT teaching
 * Topics:
 * 4.1 Formula Structure & Cell References (Relative, Absolute $A$1, F4)
 * 4.2 Basic Statistical Functions (MAX, MIN, AVERAGE, LARGE, COUNT, COUNTA)
 * 4.3 Conditional Aggregation (COUNTIFS and SUMIFS)
 * 4.4 Text Functions (&, CONCATENATE, TEXTJOIN, PROPER, UPPER, LOWER)
 * 4.5 Logical Function (IF, Nested IF, Fill Down)
 * 4.6 Formula Auditing (Show Formulas Ctrl+`, Trace Precedents)
 */

const SpreadsheetSimulatorsCh4 = {
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

  playWarningSound() {
    if (window.App && window.App.soundEnabled) {
      window.App.playTone(350, 0.1);
      setTimeout(() => window.App.playTone(280, 0.15), 100);
    }
  },

  simulatorDatasets: {
    '4.1': {
      title: 'Latihan 4.1: Formula Structure & Cell References',
      sheetName: 'Cell_References',
      data: "Parameter Pajak PPN:\t11%\n\n" +
            "No\tNama Barang\tQty\tHarga Satuan\tSubtotal\tPPN 11%\tTotal Bayar\n" +
            "1\tMouse Wireless\t15\t120000\t1800000\t198000\t1998000\n" +
            "2\tKeyboard USB\t8\t185000\t1480000\t162800\t1642800\n" +
            "3\tHeadset Gaming\t12\t250000\t3000000\t330000\t3330000\n" +
            "4\tWebcam 1080p\t5\t450000\t2250000\t247500\t2497500\n" +
            "5\tFlashdisk 64GB\t25\t85000\t2125000\t233750\t2358750"
    },
    '4.2': {
      title: 'Latihan 4.2: Basic Statistical Functions',
      sheetName: 'Statistical_Functions',
      data: "No\tNama Siswa\tKelas\tNilai Matematika\n" +
            "1\tAditya Pratama\t10-RPL 1\t85\n" +
            "2\tBagas Wicaksono\t10-RPL 1\t92\n" +
            "3\tCitra Lestari\t10-RPL 1\t78\n" +
            "4\tDewi Anggraini\t10-RPL 1\tSakit\n" +
            "5\tFajar Ramadhan\t10-RPL 1\t65\n" +
            "6\tGita Gutawa\t10-RPL 1\t96\n" +
            "7\tHendra Wijaya\t10-RPL 1\t88\n" +
            "8\tIndah Permata\t10-RPL 1\tIzin"
    },
    '4.3': {
      title: 'Latihan 4.3: Conditional Aggregation (COUNTIFS & SUMIFS)',
      sheetName: 'COUNTIFS_SUMIFS',
      data: "ID\tCabang\tKategori\tStatus\tTotal Penjualan\n" +
            "TRX-01\tJakarta\tHardware\tLunas\t3500000\n" +
            "TRX-02\tSurabaya\tSoftware\tLunas\t1200000\n" +
            "TRX-03\tJakarta\tHardware\tPending\t2100000\n" +
            "TRX-04\tBandung\tAksesoris\tLunas\t850000\n" +
            "TRX-05\tSurabaya\tHardware\tLunas\t4500000\n" +
            "TRX-06\tJakarta\tSoftware\tLunas\t1800000\n" +
            "TRX-07\tBandung\tHardware\tPending\t3000000\n" +
            "TRX-08\tSurabaya\tAksesoris\tPending\t600000"
    },
    '4.4': {
      title: 'Latihan 4.4: Text Functions & Case Conversion',
      sheetName: 'Text_Functions',
      data: "No\tNama Depan\tNama Belakang\tKelas\tKota Asal\n" +
            "1\taHMaD\tkURniAWan\t10-RPL\tbandung\n" +
            "2\tsITi\taMinAH\t10-TKJ\tJAKARTA\n" +
            "3\tbUDi\tsANToso\t10-RPL\tsurabaya\n" +
            "4\tdInDa\tpUTrI\t10-DKV\tSEMARANG\n" +
            "5\trIZkY\tpRaTAmA\t10-TKJ\tyogyakarta"
    },
    '4.5': {
      title: 'Latihan 4.5: Logical Function IF & Fill Down',
      sheetName: 'Logical_IF',
      data: "No\tNama Siswa\tNilai Akhir\tKKM\tStatus Kelulusan\tPredikat Huruf\n" +
            "1\tAlifia Zahra\t88\t75\tLULUS\tA\n" +
            "2\tBintang Samudra\t68\t75\tREMEDIAL\tC\n" +
            "3\tCantika Dewi\t74\t75\tREMEDIAL\tC\n" +
            "4\tDimas Anggara\t95\t75\tLULUS\tA\n" +
            "5\tErlangga Putra\t60\t75\tREMEDIAL\tD\n" +
            "6\tFarhan Maulana\t78\t75\tLULUS\tB\n" +
            "7\tGisella Anastasia\t82\t75\tLULUS\tB"
    },
    '4.6': {
      title: 'Latihan 4.6: Formula Auditing & Show Formulas',
      sheetName: 'Formula_Auditing',
      data: "Item Belanja\tHarga\tQty\tSubtotal\tDiskon Member\tTotal Bersih\n" +
            "Monitor LED\t2200000\t2\t4400000\t440000\t3960000\n" +
            "Mouse Gaming\t350000\t4\t1400000\t0\t1400000\n" +
            "Keyboard Mech\t750000\t3\t2250000\t225000\t2025000"
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
  // LAB 4.1: FORMULA STRUCTURE & CELL REFERENCES
  // ==========================================
  initLab41() {
    const container = document.getElementById('lab-4-1-container');
    if (!container) return;

    let useAbsolute = true;
    let isCalculated = false;
    let selectedCell = 'F4';

    const items = [
      { no: 1, nama: 'Mouse Wireless', qty: 15, harga: 120000 },
      { no: 2, nama: 'Keyboard USB', qty: 8, harga: 185000 },
      { no: 3, nama: 'Headset Gaming', qty: 12, harga: 250000 },
      { no: 4, nama: 'Webcam 1080p', qty: 5, harga: 450000 },
      { no: 5, nama: 'Flashdisk 64GB', qty: 25, harga: 85000 }
    ];

    function render() {
      let formulaBarText = '';
      if (selectedCell.startsWith('E')) {
        const rowNum = parseInt(selectedCell.substring(1), 10);
        formulaBarText = isCalculated ? `=C${rowNum}*D${rowNum}` : '';
      } else if (selectedCell.startsWith('F')) {
        const rowNum = parseInt(selectedCell.substring(1), 10);
        if (isCalculated) {
          formulaBarText = useAbsolute ? `=E${rowNum}*$B$1` : `=E${rowNum}*B${rowNum - 3}`;
        }
      } else if (selectedCell.startsWith('G')) {
        const rowNum = parseInt(selectedCell.substring(1), 10);
        formulaBarText = isCalculated ? `=E${rowNum}+F${rowNum}` : '';
      } else if (selectedCell === 'B1') {
        formulaBarText = '11%';
      }

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 4.1: Struktur Rumus & Referensi Sel ($B$1 vs B1)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Pintasan Cepat: Tombol F4</div>
          </div>

          <div class="excel-toolbar">
            <span style="font-size:0.8rem; font-weight:700; color:var(--text-muted); margin-right:4px;">Mode Referensi Sel:</span>
            
            <button class="excel-btn-sm ${useAbsolute ? 'active' : ''}" id="lab41-btn-abs" title="Kunci sel B1 menggunakan referensi absolut $B$1">
              🔒 Absolut: <code>$B$1</code> (F4)
            </button>
            <button class="excel-btn-sm ${!useAbsolute ? 'active' : ''}" id="lab41-btn-rel" style="${!useAbsolute ? 'background:#e11d48;border-color:#e11d48;color:white;' : ''}" title="Simulasikan error referensi relatif B1">
              ⚠️ Relatif: <code>B1</code> (Tanpa $)
            </button>

            <button class="btn btn-primary" id="lab41-btn-calc" style="padding:6px 14px; font-size:0.82rem; margin-left:8px;">
              ⚡ Isi Rumus & Fill Down
            </button>
            <button class="btn btn-outline" id="lab41-btn-reset" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Reset
            </button>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh4.copySimulatorByLabId('4.1')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <!-- Formula Bar -->
          <div class="excel-formula-bar">
            <div class="excel-name-box" id="lab41-name-box">${selectedCell}</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" id="lab41-formula-input" value="${formulaBarText}" placeholder="Pilih sel di kolom Subtotal, PPN, atau Total Bayar untuk melihat rumus" readonly>
          </div>

          <!-- Spreadsheet Grid Canvas -->
          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr class="excel-header-row">
                  <th class="excel-row-num">&nbsp;</th>
                  <th class="excel-header-cell">A</th>
                  <th class="excel-header-cell">B</th>
                  <th class="excel-header-cell">C</th>
                  <th class="excel-header-cell">D</th>
                  <th class="excel-header-cell">E</th>
                  <th class="excel-header-cell">F</th>
                  <th class="excel-header-cell">G</th>
                </tr>
              </thead>
              <tbody>
                <!-- Row 1: Parameter Cell B1 -->
                <tr class="excel-data-row">
                  <td class="excel-row-num">1</td>
                  <td class="excel-cell" style="font-weight:700; color:var(--excel-green); background:var(--excel-green-light);">Tarif PPN:</td>
                  <td class="excel-cell ${selectedCell === 'B1' ? 'selected' : ''}" data-cell="B1" style="font-weight:800; text-align:right; color:var(--excel-green); background:var(--excel-green-light); cursor:pointer;">
                    11%
                  </td>
                  <td class="excel-cell" colspan="5" style="color:var(--text-muted); font-size:0.75rem; font-style:italic;">
                    ← Ini adalah sel parameter tunggal. Wajib dikunci dengan simbol dollar: <strong>$B$1</strong>
                  </td>
                </tr>
                <!-- Row 2: Spacer -->
                <tr class="excel-data-row" style="height:14px; background:var(--bg-subtle);">
                  <td class="excel-row-num">2</td>
                  <td colspan="7" style="border:none;"></td>
                </tr>
                <!-- Row 3: Header Tabel -->
                <tr class="excel-data-row">
                  <td class="excel-row-num">3</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:center;">No</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700;">Nama Barang</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:right;">Qty</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:right;">Harga Satuan</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:right;">Subtotal</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:right; background:#dcfce7; color:#166534;">PPN 11%</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:right;">Total Bayar</td>
                </tr>
                <!-- Data Rows 4 to 8 -->
                ${items.map((item, idx) => {
                  const r = idx + 4;
                  const subtotal = item.qty * item.harga;
                  
                  // PPN calculation depends on absolute vs relative
                  let ppn = 0;
                  let isErrorRow = false;
                  if (isCalculated) {
                    if (useAbsolute) {
                      ppn = subtotal * 0.11;
                    } else {
                      // Relative: row 4 uses B1 (works), row 5 uses B2 (empty -> 0), row 6 uses B3 (text -> error)
                      if (r === 4) {
                        ppn = subtotal * 0.11;
                      } else {
                        ppn = 0;
                        isErrorRow = true;
                      }
                    }
                  }
                  const total = isCalculated ? (subtotal + ppn) : '';

                  return `
                    <tr class="excel-data-row">
                      <td class="excel-row-num">${r}</td>
                      <td class="excel-cell" style="text-align:center;">${item.no}</td>
                      <td class="excel-cell">${item.nama}</td>
                      <td class="excel-cell" style="text-align:right;">${item.qty}</td>
                      <td class="excel-cell" style="text-align:right;">${item.harga.toLocaleString('id-ID')}</td>
                      
                      <!-- Col E: Subtotal -->
                      <td class="excel-cell ${selectedCell === `E${r}` ? 'selected' : ''}" data-cell="E${r}" style="text-align:right; cursor:pointer;">
                        ${isCalculated ? subtotal.toLocaleString('id-ID') : '<span style="color:#94a3b8;font-size:0.75rem;">=C' + r + '*D' + r + '</span>'}
                      </td>

                      <!-- Col F: PPN 11% -->
                      <td class="excel-cell ${selectedCell === `F${r}` ? 'selected' : ''}" data-cell="F${r}" style="text-align:right; cursor:pointer; font-weight:700; ${isErrorRow ? 'background:#fee2e2;color:#b91c1c;' : 'background:#f0fdf4;color:#166534;'}">
                        ${isCalculated ? (isErrorRow ? 'Rp 0 ⚠️' : Math.round(ppn).toLocaleString('id-ID')) : '<span style="color:#94a3b8;font-size:0.75rem;">' + (useAbsolute ? `=$B$1` : `B${r-3}`) + '</span>'}
                      </td>

                      <!-- Col G: Total Bayar -->
                      <td class="excel-cell ${selectedCell === `G${r}` ? 'selected' : ''}" data-cell="G${r}" style="text-align:right; font-weight:700; cursor:pointer;">
                        ${isCalculated ? (typeof total === 'number' ? Math.round(total).toLocaleString('id-ID') : total) : '<span style="color:#94a3b8;font-size:0.75rem;">=E' + r + '+F' + r + '</span>'}
                      </td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>

          <!-- Status Bar & Educational Callout -->
          <div class="excel-statusbar" style="flex-direction:column; align-items:flex-start; gap:4px; padding:8px 12px;">
            ${!isCalculated ? `
              <div style="font-size:0.82rem; color:var(--text-muted);">
                👉 Klik tombol <strong>"⚡ Isi Rumus & Fill Down"</strong> di atas untuk melihat bagaimana rumus disalin ke semua baris.
              </div>
            ` : useAbsolute ? `
              <div style="font-size:0.82rem; color:#166534; font-weight:600; display:flex; align-items:center; gap:6px;">
                <span>✅ SUKSES (REFERENSI ABSOLUT $B$1):</span> Sel parameter B1 dikunci dengan simbol dollar ($). Saat disalin ke baris 5, 6, 7, 8 rumus tetap mengalikan dengan $B$1 (11%).
              </div>
            ` : `
              <div style="font-size:0.82rem; color:#b91c1c; font-weight:600; display:flex; align-items:center; gap:6px;">
                <span>⚠️ ERROR REFERENSI RELATIF:</span> Baris 5-8 menjadi Rp 0 karena rumus bergeser mencari tarif pajak ke B2, B3 yang kosong! Tekan F4 untuk mengunci sel menjadi $B$1.
              </div>
            `}
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('lab41-btn-abs')?.addEventListener('click', () => {
        useAbsolute = true;
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });

      document.getElementById('lab41-btn-rel')?.addEventListener('click', () => {
        useAbsolute = false;
        SpreadsheetSimulatorsCh4.playWarningSound();
        render();
      });

      document.getElementById('lab41-btn-calc')?.addEventListener('click', () => {
        isCalculated = true;
        if (useAbsolute) {
          SpreadsheetSimulatorsCh4.playSuccessSound();
        } else {
          SpreadsheetSimulatorsCh4.playWarningSound();
        }
        render();
      });

      document.getElementById('lab41-btn-reset')?.addEventListener('click', () => {
        isCalculated = false;
        useAbsolute = true;
        selectedCell = 'F4';
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });

      // Cell selection click
      container.querySelectorAll('.excel-cell[data-cell]').forEach(td => {
        td.addEventListener('click', () => {
          selectedCell = td.getAttribute('data-cell') || 'A1';
          SpreadsheetSimulatorsCh4.playClickSound();
          render();
        });
      });
    }

    render();
  },

  // ==========================================
  // LAB 4.2: BASIC STATISTICAL FUNCTIONS
  // ==========================================
  initLab42() {
    const container = document.getElementById('lab-4-2-container');
    if (!container) return;

    let activeFilter = 'all'; // 'all', 'max', 'min', 'average', 'large2', 'count', 'counta'

    const students = [
      { no: 1, nama: 'Aditya Pratama', kelas: '10-RPL 1', nilai: 85, isNum: true },
      { no: 2, nama: 'Bagas Wicaksono', kelas: '10-RPL 1', nilai: 92, isNum: true },
      { no: 3, nama: 'Citra Lestari', kelas: '10-RPL 1', nilai: 78, isNum: true },
      { no: 4, nama: 'Dewi Anggraini', kelas: '10-RPL 1', nilai: 'Sakit', isNum: false },
      { no: 5, nama: 'Fajar Ramadhan', kelas: '10-RPL 1', nilai: 65, isNum: true },
      { no: 6, nama: 'Gita Gutawa', kelas: '10-RPL 1', nilai: 96, isNum: true },
      { no: 7, nama: 'Hendra Wijaya', kelas: '10-RPL 1', nilai: 88, isNum: true },
      { no: 8, nama: 'Indah Permata', kelas: '10-RPL 1', nilai: 'Izin', isNum: false }
    ];

    // Numbers: 85, 92, 78, 65, 96, 88
    const numValues = [85, 92, 78, 65, 96, 88];
    const statMax = Math.max(...numValues); // 96
    const statMin = Math.min(...numValues); // 65
    const statAvg = (numValues.reduce((a, b) => a + b, 0) / numValues.length).toFixed(1); // 84.0
    const sortedNums = [...numValues].sort((a, b) => b - a); // 96, 92, 88, 85, 78, 65
    const statLarge2 = sortedNums[1]; // 92
    const statCount = numValues.length; // 6
    const statCountA = students.length; // 8

    function render() {
      let formulaBarText = '';
      if (activeFilter === 'max') formulaBarText = '=MAX(D2:D9)';
      else if (activeFilter === 'min') formulaBarText = '=MIN(D2:D9)';
      else if (activeFilter === 'average') formulaBarText = '=AVERAGE(D2:D9)';
      else if (activeFilter === 'large2') formulaBarText = '=LARGE(D2:D9; 2)';
      else if (activeFilter === 'count') formulaBarText = '=COUNT(D2:D9)';
      else if (activeFilter === 'counta') formulaBarText = '=COUNTA(B2:B9)';
      else formulaBarText = '=AVERAGE(D2:D9)';

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 4.2: Statistik Dasar & Perbedaan COUNT vs COUNTA</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Pintasan Cepat: Alt + = (AutoSum)</div>
          </div>

          <div class="excel-toolbar">
            <span style="font-size:0.8rem; font-weight:700; color:var(--text-muted); margin-right:4px;">Uji Rumus Statistik:</span>
            
            <button class="excel-btn-sm ${activeFilter === 'max' ? 'active' : ''}" id="lab42-btn-max">
              ⭐ MAX (Tertinggi)
            </button>
            <button class="excel-btn-sm ${activeFilter === 'min' ? 'active' : ''}" id="lab42-btn-min">
              🔻 MIN (Terendah)
            </button>
            <button class="excel-btn-sm ${activeFilter === 'average' ? 'active' : ''}" id="lab42-btn-avg">
              📊 AVERAGE
            </button>
            <button class="excel-btn-sm ${activeFilter === 'large2' ? 'active' : ''}" id="lab42-btn-large2">
              🥈 LARGE (k=2)
            </button>
            <button class="excel-btn-sm ${activeFilter === 'count' ? 'active' : ''}" id="lab42-btn-count" style="${activeFilter === 'count' ? 'background:#0284c7;border-color:#0284c7;color:white;' : ''}">
              🔢 COUNT (Angka Saja)
            </button>
            <button class="excel-btn-sm ${activeFilter === 'counta' ? 'active' : ''}" id="lab42-btn-counta" style="${activeFilter === 'counta' ? 'background:#8b5cf6;border-color:#8b5cf6;color:white;' : ''}">
              📋 COUNTA (Semua Terisi)
            </button>
            <button class="excel-btn-sm ${activeFilter === 'all' ? 'active' : ''}" id="lab42-btn-all">
              ⚡ Hitung Semua
            </button>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh4.copySimulatorByLabId('4.2')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <!-- Formula Bar -->
          <div class="excel-formula-bar">
            <div class="excel-name-box">D10</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="${formulaBarText}" readonly>
          </div>

          <!-- Spreadsheet Grid Canvas -->
          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr class="excel-header-row">
                  <th class="excel-row-num">&nbsp;</th>
                  <th class="excel-header-cell">A</th>
                  <th class="excel-header-cell">B</th>
                  <th class="excel-header-cell">C</th>
                  <th class="excel-header-cell">D</th>
                </tr>
              </thead>
              <tbody>
                <!-- Header Row 1 -->
                <tr class="excel-data-row">
                  <td class="excel-row-num">1</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:center;">No</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700;">Nama Siswa</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:center;">Kelas</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:right;">Nilai Matematika</td>
                </tr>

                <!-- Data Rows 2 to 9 -->
                ${students.map((st, idx) => {
                  const r = idx + 2;
                  let cellStyle = "text-align:right;";
                  let badge = "";

                  if (st.nilai === statMax && (activeFilter === 'max' || activeFilter === 'all')) {
                    cellStyle += "background:#dcfce7;color:#166534;font-weight:800;outline:2px solid #22c55e;";
                    badge = " <span style='font-size:0.7rem;background:#22c55e;color:white;padding:1px 5px;border-radius:4px;'>MAX</span>";
                  } else if (st.nilai === statMin && (activeFilter === 'min' || activeFilter === 'all')) {
                    cellStyle += "background:#fee2e2;color:#b91c1c;font-weight:800;outline:2px solid #ef4444;";
                    badge = " <span style='font-size:0.7rem;background:#ef4444;color:white;padding:1px 5px;border-radius:4px;'>MIN</span>";
                  } else if (st.nilai === statLarge2 && (activeFilter === 'large2' || activeFilter === 'all')) {
                    cellStyle += "background:#fef3c7;color:#92400e;font-weight:800;outline:2px solid #f59e0b;";
                    badge = " <span style='font-size:0.7rem;background:#f59e0b;color:white;padding:1px 5px;border-radius:4px;'>RANK 2</span>";
                  } else if (!st.isNum) {
                    cellStyle += "background:#fff1f2;color:#e11d48;font-weight:700;";
                  }

                  if (activeFilter === 'count' && st.isNum) {
                    cellStyle += "background:#e0f2fe;color:#0369a1;font-weight:700;";
                  } else if (activeFilter === 'count' && !st.isNum) {
                    cellStyle += "opacity:0.35;text-decoration:line-through;";
                  }

                  if (activeFilter === 'counta') {
                    cellStyle += "background:#ede9fe;color:#6d28d9;font-weight:700;";
                  }

                  return `
                    <tr class="excel-data-row">
                      <td class="excel-row-num">${r}</td>
                      <td class="excel-cell" style="text-align:center;">${st.no}</td>
                      <td class="excel-cell" style="${activeFilter === 'counta' ? 'background:#ede9fe;font-weight:700;' : ''}">${st.nama}</td>
                      <td class="excel-cell" style="text-align:center;">${st.kelas}</td>
                      <td class="excel-cell" style="${cellStyle}">
                        ${st.nilai}${badge}
                      </td>
                    </tr>
                  `;
                }).join('')}

                <!-- Summary Row -->
                <tr class="excel-data-row" style="background:#f8fafc; font-weight:700; border-top:2px solid var(--border-color);">
                  <td class="excel-row-num">10</td>
                  <td colspan="3" class="excel-cell" style="text-align:right; color:var(--text-main);">
                    ${activeFilter === 'max' ? 'Nilai Tertinggi (=MAX):' :
                      activeFilter === 'min' ? 'Nilai Terendah (=MIN):' :
                      activeFilter === 'average' ? 'Rata-Rata Nilai (=AVERAGE):' :
                      activeFilter === 'large2' ? 'Juara 2 / Nilai Tertinggi Ke-2 (=LARGE k=2):' :
                      activeFilter === 'count' ? 'Siswa Ikut Ujian / Berangka (=COUNT):' :
                      activeFilter === 'counta' ? 'Total Siswa Terdaftar (=COUNTA):' :
                      'Ringkasan Kalkulasi Statistik Lengkap:'}
                  </td>
                  <td class="excel-cell" style="text-align:right; font-size:1rem; font-weight:800; color:var(--excel-green); background:#f0fdf4;">
                    ${activeFilter === 'max' ? statMax :
                      activeFilter === 'min' ? statMin :
                      activeFilter === 'average' ? statAvg :
                      activeFilter === 'large2' ? statLarge2 :
                      activeFilter === 'count' ? `${statCount} orang (Angka saja)` :
                      activeFilter === 'counta' ? `${statCountA} siswa (Semua terisi)` :
                      `Avg: ${statAvg} | Max: ${statMax}`}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          <!-- Bottom Stat Badges Grid -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(130px, 1fr)); gap:8px; padding:12px; background:var(--bg-subtle); border-top:1px solid var(--border-color);">
            <div style="background:var(--bg-surface); padding:8px; border-radius:6px; border:1px solid var(--border-color); text-align:center;">
              <div style="font-size:0.7rem; color:var(--text-muted); font-weight:600;">MAX (Tertinggi)</div>
              <div style="font-size:1.15rem; font-weight:800; color:#166534;">${statMax}</div>
              <code style="font-size:0.68rem; color:var(--text-muted);">=MAX(D2:D9)</code>
            </div>
            <div style="background:var(--bg-surface); padding:8px; border-radius:6px; border:1px solid var(--border-color); text-align:center;">
              <div style="font-size:0.7rem; color:var(--text-muted); font-weight:600;">MIN (Terendah)</div>
              <div style="font-size:1.15rem; font-weight:800; color:#e11d48;">${statMin}</div>
              <code style="font-size:0.68rem; color:var(--text-muted);">=MIN(D2:D9)</code>
            </div>
            <div style="background:var(--bg-surface); padding:8px; border-radius:6px; border:1px solid var(--border-color); text-align:center;">
              <div style="font-size:0.7rem; color:var(--text-muted); font-weight:600;">AVERAGE</div>
              <div style="font-size:1.15rem; font-weight:800; color:#0284c7;">${statAvg}</div>
              <code style="font-size:0.68rem; color:var(--text-muted);">=AVERAGE(D2:D9)</code>
            </div>
            <div style="background:var(--bg-surface); padding:8px; border-radius:6px; border:1px solid var(--border-color); text-align:center;">
              <div style="font-size:0.7rem; color:var(--text-muted); font-weight:600;">LARGE (Rank 2)</div>
              <div style="font-size:1.15rem; font-weight:800; color:#d97706;">${statLarge2}</div>
              <code style="font-size:0.68rem; color:var(--text-muted);">=LARGE(D2:D9; 2)</code>
            </div>
            <div style="background:var(--bg-surface); padding:8px; border-radius:6px; border:1px solid var(--border-color); text-align:center;">
              <div style="font-size:0.7rem; color:#0369a1; font-weight:700;">COUNT (Hanya Angka)</div>
              <div style="font-size:1.15rem; font-weight:800; color:#0369a1;">${statCount}</div>
              <div style="font-size:0.68rem; color:#0369a1;">(Abaikan Sakit & Izin)</div>
            </div>
            <div style="background:var(--bg-surface); padding:8px; border-radius:6px; border:1px solid var(--border-color); text-align:center;">
              <div style="font-size:0.7rem; color:#6d28d9; font-weight:700;">COUNTA (Total Terisi)</div>
              <div style="font-size:1.15rem; font-weight:800; color:#6d28d9;">${statCountA}</div>
              <div style="font-size:0.68rem; color:#6d28d9;">(Hitung Semua Kolom B)</div>
            </div>
          </div>
        </div>
      `;

      // Event bindings
      document.getElementById('lab42-btn-max')?.addEventListener('click', () => { activeFilter = 'max'; SpreadsheetSimulatorsCh4.playClickSound(); render(); });
      document.getElementById('lab42-btn-min')?.addEventListener('click', () => { activeFilter = 'min'; SpreadsheetSimulatorsCh4.playClickSound(); render(); });
      document.getElementById('lab42-btn-avg')?.addEventListener('click', () => { activeFilter = 'average'; SpreadsheetSimulatorsCh4.playClickSound(); render(); });
      document.getElementById('lab42-btn-large2')?.addEventListener('click', () => { activeFilter = 'large2'; SpreadsheetSimulatorsCh4.playClickSound(); render(); });
      document.getElementById('lab42-btn-count')?.addEventListener('click', () => { activeFilter = 'count'; SpreadsheetSimulatorsCh4.playClickSound(); render(); });
      document.getElementById('lab42-btn-counta')?.addEventListener('click', () => { activeFilter = 'counta'; SpreadsheetSimulatorsCh4.playClickSound(); render(); });
      document.getElementById('lab42-btn-all')?.addEventListener('click', () => { activeFilter = 'all'; SpreadsheetSimulatorsCh4.playSuccessSound(); render(); });
    }

    render();
  },

  // ==========================================
  // LAB 4.3: CONDITIONAL AGGREGATION (COUNTIFS & SUMIFS)
  // ==========================================
  initLab43() {
    const container = document.getElementById('lab-4-3-container');
    if (!container) return;

    let selectedBranch = 'Surabaya';
    let selectedStatus = 'Lunas';
    let isSorted = false;

    const rawData = [
      { id: 'TRX-01', cabang: 'Jakarta', kategori: 'Hardware', status: 'Lunas', nominal: 3500000 },
      { id: 'TRX-02', cabang: 'Surabaya', kategori: 'Software', status: 'Lunas', nominal: 1200000 },
      { id: 'TRX-03', cabang: 'Jakarta', kategori: 'Hardware', status: 'Pending', nominal: 2100000 },
      { id: 'TRX-04', cabang: 'Bandung', kategori: 'Aksesoris', status: 'Lunas', nominal: 850000 },
      { id: 'TRX-05', cabang: 'Surabaya', kategori: 'Hardware', status: 'Lunas', nominal: 4500000 },
      { id: 'TRX-06', cabang: 'Jakarta', kategori: 'Software', status: 'Lunas', nominal: 1800000 },
      { id: 'TRX-07', cabang: 'Bandung', kategori: 'Hardware', status: 'Pending', nominal: 3000000 },
      { id: 'TRX-08', cabang: 'Surabaya', kategori: 'Aksesoris', status: 'Pending', nominal: 600000 }
    ];

    function render() {
      // Sorting check
      let displayData = [...rawData];
      if (isSorted) {
        displayData.sort((a, b) => b.nominal - a.nominal); // Sort by nominal descending
      }

      // Filter matches
      const matchingRows = displayData.filter(d => {
        const matchBranch = (selectedBranch === 'Semua') || (d.cabang === selectedBranch);
        const matchStatus = (selectedStatus === 'Semua') || (d.status === selectedStatus);
        return matchBranch && matchStatus;
      });

      const countResult = matchingRows.length;
      const sumResult = matchingRows.reduce((acc, row) => acc + row.nominal, 0);

      // Construct formula string
      let countifsFormula = `=COUNTIFS(B2:B9; "${selectedBranch}"; D2:D9; "${selectedStatus}")`;
      let sumifsFormula = `=SUMIFS(E2:E9; B2:B9; "${selectedBranch}"; D2:D9; "${selectedStatus}")`;

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 4.3: Agregasi Bersyarat Multi-Kriteria (COUNTIFS & SUMIFS)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Tahan Sort & Filter Data</div>
          </div>

          <div class="excel-toolbar">
            <div class="sim-input-group">
              <span class="sim-label" style="min-width:auto;">Kriteria Cabang:</span>
              <select id="lab43-branch-sel" class="sim-input" style="width:110px;">
                <option value="Surabaya" ${selectedBranch === 'Surabaya' ? 'selected' : ''}>Surabaya</option>
                <option value="Jakarta" ${selectedBranch === 'Jakarta' ? 'selected' : ''}>Jakarta</option>
                <option value="Bandung" ${selectedBranch === 'Bandung' ? 'selected' : ''}>Bandung</option>
                <option value="Semua" ${selectedBranch === 'Semua' ? 'selected' : ''}>[Semua Cabang]</option>
              </select>
            </div>

            <div class="sim-input-group">
              <span class="sim-label" style="min-width:auto;">Kriteria Status:</span>
              <select id="lab43-status-sel" class="sim-input" style="width:110px;">
                <option value="Lunas" ${selectedStatus === 'Lunas' ? 'selected' : ''}>Lunas</option>
                <option value="Pending" ${selectedStatus === 'Pending' ? 'selected' : ''}>Pending</option>
                <option value="Semua" ${selectedStatus === 'Semua' ? 'selected' : ''}>[Semua Status]</option>
              </select>
            </div>

            <button class="btn btn-primary" id="lab43-btn-sort" style="padding:6px 14px; font-size:0.82rem;" title="Buktikan bahwa hasil COUNTIFS & SUMIFS tidak berubah setelah data diurutkan">
              ${isSorted ? '↺ Urutan Awal' : '🔀 Uji Sort Nominal (Z-A)'}
            </button>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh4.copySimulatorByLabId('4.3')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <!-- Formula Bar -->
          <div class="excel-formula-bar">
            <div class="excel-name-box">G2</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="${sumifsFormula}" readonly>
          </div>

          <!-- Spreadsheet Grid Canvas -->
          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr class="excel-header-row">
                  <th class="excel-row-num">&nbsp;</th>
                  <th class="excel-header-cell">A</th>
                  <th class="excel-header-cell">B</th>
                  <th class="excel-header-cell">C</th>
                  <th class="excel-header-cell">D</th>
                  <th class="excel-header-cell">E</th>
                </tr>
              </thead>
              <tbody>
                <tr class="excel-data-row">
                  <td class="excel-row-num">1</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:center;">ID</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700;">Cabang</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700;">Kategori</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:center;">Status</td>
                  <td class="excel-cell excel-cell-header" style="font-weight:700; text-align:right;">Total Penjualan</td>
                </tr>

                ${displayData.map((row, idx) => {
                  const r = idx + 2;
                  const isMatch = matchingRows.includes(row);

                  return `
                    <tr class="excel-data-row" style="${isMatch ? 'background:rgba(16, 124, 65, 0.08);' : ''}">
                      <td class="excel-row-num">${r}</td>
                      <td class="excel-cell" style="text-align:center;">${row.id}</td>
                      <td class="excel-cell" style="${row.cabang === selectedBranch ? 'color:var(--excel-green);font-weight:700;' : ''}">${row.cabang}</td>
                      <td class="excel-cell">${row.kategori}</td>
                      <td class="excel-cell" style="text-align:center;">
                        <span class="badge" style="${row.status === 'Lunas' ? 'background:#dcfce7;color:#166534;' : 'background:#fef3c7;color:#92400e;'}">
                          ${row.status}
                        </span>
                      </td>
                      <td class="excel-cell" style="text-align:right; font-weight:${isMatch ? '700' : 'normal'}; ${isMatch ? 'color:var(--excel-green);' : ''}">
                        Rp ${row.nominal.toLocaleString('id-ID')}
                        ${isMatch ? ' ✔' : ''}
                      </td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>

          <!-- Live Result Cards -->
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px; padding:12px; background:var(--bg-subtle); border-top:1px solid var(--border-color);">
            <div style="background:var(--bg-surface); padding:10px 14px; border-radius:8px; border:1px solid var(--border-color);">
              <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; font-weight:700; color:var(--text-muted);">COUNTIFS (Jumlah Baris Cocok):</span>
                <span class="badge" style="background:#e0f2fe; color:#0284c7; font-size:0.75rem;">Frekuensi</span>
              </div>
              <div style="font-size:1.4rem; font-weight:800; color:#0284c7; margin:4px 0;">
                ${countResult} Transaksi
              </div>
              <code style="font-size:0.72rem; color:var(--text-muted); display:block; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">
                ${countifsFormula}
              </code>
            </div>

            <div style="background:var(--bg-surface); padding:10px 14px; border-radius:8px; border:1px solid var(--border-color);">
              <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; font-weight:700; color:var(--text-muted);">SUMIFS (Total Nilai Terjual):</span>
                <span class="badge" style="background:#dcfce7; color:#166534; font-size:0.75rem;">Akumulasi Nominal</span>
              </div>
              <div style="font-size:1.4rem; font-weight:800; color:var(--excel-green); margin:4px 0;">
                Rp ${sumResult.toLocaleString('id-ID')}
              </div>
              <code style="font-size:0.72rem; color:var(--text-muted); display:block; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">
                ${sumifsFormula}
              </code>
            </div>
          </div>

          <!-- Status Bar -->
          <div class="excel-statusbar">
            <span>
              ${isSorted ? '🔀 <strong>Tabel Telah Di-Sort:</strong> Hasil perhitungan tetap 100% SAMA karena mengevaluasi nilai sel, bukan urutan baris!' : '💡 Ganti pilihan dropdown Cabang atau Status untuk melihat hasil rumus berubah seketika.'}
            </span>
          </div>
        </div>
      `;

      // Event handlers
      document.getElementById('lab43-branch-sel')?.addEventListener('change', (e) => {
        selectedBranch = e.target.value;
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });

      document.getElementById('lab43-status-sel')?.addEventListener('change', (e) => {
        selectedStatus = e.target.value;
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });

      document.getElementById('lab43-btn-sort')?.addEventListener('click', () => {
        isSorted = !isSorted;
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 4.4: TEXT FUNCTIONS (PROPER, &, TEXTJOIN)
  // ==========================================
  initLab44() {
    const container = document.getElementById('lab-4-4-container');
    if (!container) return;

    let appliedAction = 'none'; // 'none', 'proper', 'upper', 'textjoin', 'all'

    const students = [
      { no: 1, depan: 'aHMaD', belakang: 'kURniAWan', kelas: '10-RPL', kota: 'bandung' },
      { no: 2, depan: 'sITi', belakang: 'aMinAH', kelas: '10-TKJ', kota: 'JAKARTA' },
      { no: 3, depan: 'bUDi', belakang: 'sANToso', kelas: '10-RPL', kota: 'surabaya' },
      { no: 4, depan: 'dInDa', belakang: 'pUTrI', kelas: '10-DKV', kota: 'SEMARANG' },
      { no: 5, depan: 'rIZkY', belakang: 'pRaTAmA', kelas: '10-TKJ', kota: 'yogyakarta' }
    ];

    function toProperCase(str) {
      if (!str) return '';
      return str.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    }

    function render() {
      let formulaBarText = '';
      if (appliedAction === 'proper') formulaBarText = '=PROPER(B2 & " " & C2)';
      else if (appliedAction === 'upper') formulaBarText = '=UPPER(E2)';
      else if (appliedAction === 'textjoin') formulaBarText = '=TEXTJOIN(" - "; TRUE; F2; D2; UPPER(E2))';
      else if (appliedAction === 'all') formulaBarText = '=PROPER(B2 & " " & C2)';

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 4.4: Manipulasi Teks (&, PROPER, UPPER, TEXTJOIN)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Pembersihan Data Otomatis</div>
          </div>

          <div class="excel-toolbar">
            <span style="font-size:0.8rem; font-weight:700; color:var(--text-muted); margin-right:4px;">Aksi Fungsi Teks:</span>

            <button class="excel-btn-sm ${appliedAction === 'proper' || appliedAction === 'all' ? 'active' : ''}" id="lab44-btn-proper">
              ✨ 1. Terapkan PROPER & &
            </button>
            <button class="excel-btn-sm ${appliedAction === 'upper' || appliedAction === 'all' ? 'active' : ''}" id="lab44-btn-upper">
              🔠 2. UPPER Kota
            </button>
            <button class="excel-btn-sm ${appliedAction === 'textjoin' || appliedAction === 'all' ? 'active' : ''}" id="lab44-btn-textjoin">
              🔗 3. TEXTJOIN Profil
            </button>
            <button class="btn btn-primary" id="lab44-btn-all" style="padding:6px 14px; font-size:0.82rem;">
              ⚡ Bersihkan Semua Data
            </button>
            <button class="btn btn-outline" id="lab44-btn-reset" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Reset Data Mentah
            </button>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh4.copySimulatorByLabId('4.4')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <!-- Formula Bar -->
          <div class="excel-formula-bar">
            <div class="excel-name-box">F2</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="${formulaBarText}" placeholder="Pilih aksi tombol di atas untuk menjalankan rumus teks" readonly>
          </div>

          <!-- Spreadsheet Grid Canvas -->
          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr class="excel-header-row">
                  <th class="excel-row-num">&nbsp;</th>
                  <th class="excel-header-cell">A</th>
                  <th class="excel-header-cell">B</th>
                  <th class="excel-header-cell">C</th>
                  <th class="excel-header-cell">D</th>
                  <th class="excel-header-cell">E</th>
                  <th class="excel-header-cell" style="background:#dcfce7;color:#166534;">F</th>
                  <th class="excel-header-cell" style="background:#e0f2fe;color:#0369a1;">G</th>
                </tr>
              </thead>
              <tbody>
                <tr class="excel-data-row">
                  <td class="excel-row-num">1</td>
                  <td class="excel-cell excel-cell-header" style="text-align:center;">No</td>
                  <td class="excel-cell excel-cell-header">Nama Depan</td>
                  <td class="excel-cell excel-cell-header">Nama Belakang</td>
                  <td class="excel-cell excel-cell-header" style="text-align:center;">Kelas</td>
                  <td class="excel-cell excel-cell-header">Kota Asal</td>
                  <td class="excel-cell excel-cell-header" style="background:#dcfce7;color:#166534;font-weight:700;">Nama Rapi (PROPER & &)</td>
                  <td class="excel-cell excel-cell-header" style="background:#e0f2fe;color:#0369a1;font-weight:700;">Profil Siswa (TEXTJOIN)</td>
                </tr>

                ${students.map((st, idx) => {
                  const r = idx + 2;
                  const isCleanedName = (appliedAction === 'proper' || appliedAction === 'all');
                  const isUpperCity = (appliedAction === 'upper' || appliedAction === 'all');
                  const isJoinedProfile = (appliedAction === 'textjoin' || appliedAction === 'all');

                  const cleanFullName = toProperCase(`${st.depan} ${st.belakang}`);
                  const displayCity = isUpperCity ? st.kota.toUpperCase() : st.kota;
                  const profileText = `${cleanFullName} - ${st.kelas} - ${st.kota.toUpperCase()}`;

                  return `
                    <tr class="excel-data-row">
                      <td class="excel-row-num">${r}</td>
                      <td class="excel-cell" style="text-align:center;">${st.no}</td>
                      <td class="excel-cell" style="color:var(--text-muted);">${st.depan}</td>
                      <td class="excel-cell" style="color:var(--text-muted);">${st.belakang}</td>
                      <td class="excel-cell" style="text-align:center;">${st.kelas}</td>
                      <td class="excel-cell" style="${isUpperCity ? 'font-weight:700;color:#0284c7;' : ''}">${displayCity}</td>
                      
                      <!-- Col F: PROPER + & -->
                      <td class="excel-cell" style="font-weight:700; ${isCleanedName ? 'background:#f0fdf4;color:#166534;' : 'color:#94a3b8;font-size:0.75rem;'}">
                        ${isCleanedName ? cleanFullName : '=PROPER(B' + r + ' & " " & C' + r + ')'}
                      </td>

                      <!-- Col G: TEXTJOIN -->
                      <td class="excel-cell" style="${isJoinedProfile ? 'background:#f0f9ff;color:#0369a1;font-weight:600;' : 'color:#94a3b8;font-size:0.75rem;'}">
                        ${isJoinedProfile ? profileText : '=TEXTJOIN(" - "; TRUE; ...)'}
                      </td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <span>
              ${appliedAction === 'none' ? '💡 Data nama di kolom B & C berantakan. Klik tombol aksi di atas untuk merapikannya otomatis!' :
                appliedAction === 'proper' ? '✨ Rumus <code>=PROPER(B2 & " " & C2)</code> menggabungkan nama dan mengubah huruf kapital di setiap awal kata.' :
                appliedAction === 'upper' ? '🔠 Rumus <code>=UPPER(E2)</code> menyeragamkan semua huruf kota menjadi huruf besar.' :
                '✅ Seluruh nama telah dirapikan (PROPER) dan digabungkan menjadi kode profil lengkap dengan pemisah " - " (TEXTJOIN).'}
            </span>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('lab44-btn-proper')?.addEventListener('click', () => { appliedAction = 'proper'; SpreadsheetSimulatorsCh4.playClickSound(); render(); });
      document.getElementById('lab44-btn-upper')?.addEventListener('click', () => { appliedAction = 'upper'; SpreadsheetSimulatorsCh4.playClickSound(); render(); });
      document.getElementById('lab44-btn-textjoin')?.addEventListener('click', () => { appliedAction = 'textjoin'; SpreadsheetSimulatorsCh4.playClickSound(); render(); });
      document.getElementById('lab44-btn-all')?.addEventListener('click', () => { appliedAction = 'all'; SpreadsheetSimulatorsCh4.playSuccessSound(); render(); });
      document.getElementById('lab44-btn-reset')?.addEventListener('click', () => { appliedAction = 'none'; SpreadsheetSimulatorsCh4.playClickSound(); render(); });
    }

    render();
  },

  // ==========================================
  // LAB 4.5: LOGICAL FUNCTION (IF & FILL DOWN)
  // ==========================================
  initLab45() {
    const container = document.getElementById('lab-4-5-container');
    if (!container) return;

    let kkmValue = 75;
    let mode = 'simple'; // 'simple' (Lulus/Remedial) or 'nested' (A/B/C/D)
    let fillDownApplied = false;

    const students = [
      { no: 1, nama: 'Alifia Zahra', nilai: 88 },
      { no: 2, nama: 'Bintang Samudra', nilai: 68 },
      { no: 3, nama: 'Cantika Dewi', nilai: 74 },
      { no: 4, nama: 'Dimas Anggara', nilai: 95 },
      { no: 5, nama: 'Erlangga Putra', nilai: 60 },
      { no: 6, nama: 'Farhan Maulana', nilai: 78 },
      { no: 7, nama: 'Gisella Anastasia', nilai: 82 }
    ];

    function getPredicate(n) {
      if (n >= 85) return 'A';
      if (n >= 75) return 'B';
      if (n >= 65) return 'C';
      return 'D';
    }

    function render() {
      const formulaBarText = mode === 'simple'
        ? `=IF(C2>=${kkmValue}; "LULUS"; "REMEDIAL")`
        : `=IF(C2>=85; "A"; IF(C2>=75; "B"; IF(C2>=65; "C"; "D")))`;

      let passedCount = 0;
      let remedialCount = 0;

      students.forEach((st, i) => {
        if (i === 0 || fillDownApplied) {
          if (st.nilai >= kkmValue) passedCount++;
          else remedialCount++;
        }
      });

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 4.5: Fungsi Logika IF & Salin Rumus (Fill Down)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Pintasan Cepat: Ctrl + D (Fill Down)</div>
          </div>

          <div class="excel-toolbar">
            <div class="sim-input-group">
              <span class="sim-label" style="min-width:auto;">Batas KKM:</span>
              <input type="number" id="lab45-kkm-input" class="sim-input" value="${kkmValue}" min="50" max="95" style="width:65px; font-weight:700;">
            </div>

            <button class="excel-btn-sm ${mode === 'simple' ? 'active' : ''}" id="lab45-btn-mode-simple">
              🎯 IF Tunggal (LULUS / REMEDIAL)
            </button>
            <button class="excel-btn-sm ${mode === 'nested' ? 'active' : ''}" id="lab45-btn-mode-nested">
              📶 Nested IF (Predikat A-D)
            </button>

            <button class="btn btn-primary" id="lab45-btn-filldown" style="padding:6px 14px; font-size:0.82rem; margin-left:8px;">
              ⚡ Fill Down ke Semua Baris (Ctrl+D)
            </button>
            <button class="btn btn-outline" id="lab45-btn-reset" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Reset
            </button>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh4.copySimulatorByLabId('4.5')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <!-- Formula Bar -->
          <div class="excel-formula-bar">
            <div class="excel-name-box">E2</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="${formulaBarText}" readonly>
          </div>

          <!-- Spreadsheet Grid Canvas -->
          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr class="excel-header-row">
                  <th class="excel-row-num">&nbsp;</th>
                  <th class="excel-header-cell">A</th>
                  <th class="excel-header-cell">B</th>
                  <th class="excel-header-cell">C</th>
                  <th class="excel-header-cell">D</th>
                  <th class="excel-header-cell" style="background:#dcfce7;color:#166534;">E</th>
                  <th class="excel-header-cell">F</th>
                </tr>
              </thead>
              <tbody>
                <tr class="excel-data-row">
                  <td class="excel-row-num">1</td>
                  <td class="excel-cell excel-cell-header" style="text-align:center;">No</td>
                  <td class="excel-cell excel-cell-header">Nama Siswa</td>
                  <td class="excel-cell excel-cell-header" style="text-align:right;">Nilai Akhir</td>
                  <td class="excel-cell excel-cell-header" style="text-align:center;">KKM</td>
                  <td class="excel-cell excel-cell-header" style="text-align:center; background:#dcfce7; color:#166534; font-weight:700;">
                    ${mode === 'simple' ? 'Status Kelulusan (=IF)' : 'Predikat Huruf (Nested IF)'}
                  </td>
                  <td class="excel-cell excel-cell-header" style="text-align:center;">Status Fill Down</td>
                </tr>

                ${students.map((st, idx) => {
                  const r = idx + 2;
                  const isFirstRow = (idx === 0);
                  const isCalculated = isFirstRow || fillDownApplied;

                  const isPass = st.nilai >= kkmValue;
                  const statusText = mode === 'simple'
                    ? (isPass ? 'LULUS' : 'REMEDIAL')
                    : `Predikat ${getPredicate(st.nilai)}`;

                  const badgeBg = isPass ? '#dcfce7' : '#fee2e2';
                  const badgeColor = isPass ? '#166534' : '#b91c1c';

                  return `
                    <tr class="excel-data-row">
                      <td class="excel-row-num">${r}</td>
                      <td class="excel-cell" style="text-align:center;">${st.no}</td>
                      <td class="excel-cell">${st.nama}</td>
                      <td class="excel-cell" style="text-align:right; font-weight:700;">${st.nilai}</td>
                      <td class="excel-cell" style="text-align:center; color:var(--text-muted);">${kkmValue}</td>
                      
                      <!-- Col E: Result of IF -->
                      <td class="excel-cell ${isFirstRow && !fillDownApplied ? 'selected' : ''}" style="text-align:center;">
                        ${isCalculated ? `
                          <span style="display:inline-block; padding:2px 10px; border-radius:12px; font-size:0.78rem; font-weight:800; background:${badgeBg}; color:${badgeColor};">
                            ${statusText}
                          </span>
                        ` : `
                          <span style="color:#94a3b8; font-size:0.75rem; font-style:italic;">Belum di-Fill Down</span>
                        `}
                      </td>

                      <!-- Col F: Helper Status -->
                      <td class="excel-cell" style="text-align:center; font-size:0.75rem; color:var(--text-muted);">
                        ${isFirstRow ? '✅ Baris Awal (Formula Asli)' : (fillDownApplied ? '✔ Fill Down (Relatif C' + r + ')' : '⏳ Menunggu Salinan')}
                      </td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>

          <!-- Bottom Summary Status Bar -->
          <div class="excel-statusbar" style="justify-content:space-between; flex-wrap:wrap; gap:8px;">
            <span>
              ${fillDownApplied ? `📊 <strong>Hasil Kelas:</strong> ${passedCount} Siswa LULUS | ${remedialCount} Siswa REMEDIAL (Batas KKM: ${kkmValue})` : '👉 Klik tombol <strong>"⚡ Fill Down ke Semua Baris"</strong> untuk menyalin rumus ke siswa berikutnya!'}
            </span>
            <div style="font-size:0.75rem; font-weight:600; color:var(--excel-green);">
              Fill Handle: Klik ganda sudut kanan bawah sel aktif
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('lab45-kkm-input')?.addEventListener('change', (e) => {
        kkmValue = parseInt(e.target.value, 10) || 75;
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });

      document.getElementById('lab45-btn-mode-simple')?.addEventListener('click', () => {
        mode = 'simple';
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });

      document.getElementById('lab45-btn-mode-nested')?.addEventListener('click', () => {
        mode = 'nested';
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });

      document.getElementById('lab45-btn-filldown')?.addEventListener('click', () => {
        fillDownApplied = true;
        SpreadsheetSimulatorsCh4.playSuccessSound();
        render();
      });

      document.getElementById('lab45-btn-reset')?.addEventListener('click', () => {
        fillDownApplied = false;
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 4.6: FORMULA AUDITING & SHOW FORMULAS
  // ==========================================
  initLab46() {
    const container = document.getElementById('lab-4-6-container');
    if (!container) return;

    let showFormulas = false;
    let traceActive = false;

    const dataRows = [
      {
        item: 'Monitor LED 24"',
        harga: 2200000,
        qty: 2,
        subFormula: '=B2*C2',
        subVal: 4400000,
        diskFormula: '=IF(D2>2000000; D2*10%; 0)',
        diskVal: 440000,
        totFormula: '=D2-E2',
        totVal: 3960000
      },
      {
        item: 'Mouse Gaming RGB',
        harga: 350000,
        qty: 4,
        subFormula: '=B3*C3',
        subVal: 1400000,
        diskFormula: '=IF(D3>2000000; D3*10%; 0)',
        diskVal: 0,
        totFormula: '=D3-E3',
        totVal: 1400000
      },
      {
        item: 'Keyboard Mechanical',
        harga: 750000,
        qty: 3,
        subFormula: '=B4*C4',
        subVal: 2250000,
        diskFormula: '=IF(D4>2000000; D4*10%; 0)',
        diskVal: 225000,
        totFormula: '=D4-E4',
        totVal: 2025000
      }
    ];

    function render() {
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 4.6: Formula Auditing & Show Formulas (Ctrl + \`)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Pintasan Cepat: Ctrl + \` (Grave Accent)</div>
          </div>

          <div class="excel-toolbar">
            <button class="btn ${showFormulas ? 'btn-primary' : 'btn-outline'}" id="lab46-btn-toggle" style="padding:6px 14px; font-size:0.82rem; font-weight:700;">
              ${showFormulas ? '🔢 Kembali ke Tampilan Nilai' : '👁️ Show Formulas (Ctrl + `)'}
            </button>

            <button class="excel-btn-sm ${traceActive ? 'active' : ''}" id="lab46-btn-trace" style="margin-left:8px;">
              🎯 Trace Precedents (Lacak Asal Rumus F2)
            </button>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh4.copySimulatorByLabId('4.6')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <!-- Formula Bar -->
          <div class="excel-formula-bar">
            <div class="excel-name-box">F2</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="=D2-E2" readonly>
          </div>

          <!-- Spreadsheet Grid Canvas -->
          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr class="excel-header-row">
                  <th class="excel-row-num">&nbsp;</th>
                  <th class="excel-header-cell">A</th>
                  <th class="excel-header-cell">B</th>
                  <th class="excel-header-cell">C</th>
                  <th class="excel-header-cell" style="${showFormulas ? 'min-width:140px;' : ''}">D</th>
                  <th class="excel-header-cell" style="${showFormulas ? 'min-width:210px;' : ''}">E</th>
                  <th class="excel-header-cell" style="${showFormulas ? 'min-width:140px;' : ''}">F</th>
                </tr>
              </thead>
              <tbody>
                <tr class="excel-data-row">
                  <td class="excel-row-num">1</td>
                  <td class="excel-cell excel-cell-header">Item Belanja</td>
                  <td class="excel-cell excel-cell-header" style="text-align:right;">Harga</td>
                  <td class="excel-cell excel-cell-header" style="text-align:right;">Qty</td>
                  <td class="excel-cell excel-cell-header" style="text-align:right;">Subtotal</td>
                  <td class="excel-cell excel-cell-header" style="text-align:right;">Diskon Member</td>
                  <td class="excel-cell excel-cell-header" style="text-align:right;">Total Bersih</td>
                </tr>

                ${dataRows.map((r, idx) => {
                  const rowNum = idx + 2;
                  const isRow2 = (rowNum === 2);

                  return `
                    <tr class="excel-data-row">
                      <td class="excel-row-num">${rowNum}</td>
                      <td class="excel-cell">${r.item}</td>
                      <td class="excel-cell" style="text-align:right;">${r.harga.toLocaleString('id-ID')}</td>
                      <td class="excel-cell" style="text-align:right;">${r.qty}</td>
                      
                      <!-- Col D: Subtotal -->
                      <td class="excel-cell ${traceActive && isRow2 ? 'selected' : ''}" style="text-align:right; font-family:${showFormulas ? 'var(--font-mono)' : 'inherit'}; ${showFormulas ? 'background:#e0f2fe;color:#0369a1;font-weight:700;' : ''}">
                        ${showFormulas ? `<span style="color:#0284c7;">${r.subFormula}</span>` : r.subVal.toLocaleString('id-ID')}
                        ${traceActive && isRow2 ? ' <span style="color:#0284c7;font-size:0.7rem;">[Precedent 1]</span>' : ''}
                      </td>

                      <!-- Col E: Diskon -->
                      <td class="excel-cell ${traceActive && isRow2 ? 'selected' : ''}" style="text-align:right; font-family:${showFormulas ? 'var(--font-mono)' : 'inherit'}; ${showFormulas ? 'background:#fef3c7;color:#92400e;font-weight:700;' : ''}">
                        ${showFormulas ? `<span style="color:#d97706;">${r.diskFormula}</span>` : r.diskVal.toLocaleString('id-ID')}
                        ${traceActive && isRow2 ? ' <span style="color:#d97706;font-size:0.7rem;">[Precedent 2]</span>' : ''}
                      </td>

                      <!-- Col F: Total Bersih -->
                      <td class="excel-cell ${isRow2 ? 'selected' : ''}" style="text-align:right; font-weight:700; font-family:${showFormulas ? 'var(--font-mono)' : 'inherit'}; ${showFormulas ? 'background:#dcfce7;color:#166534;' : ''}">
                        ${showFormulas ? `<span style="color:#166534;">${r.totFormula}</span>` : r.totVal.toLocaleString('id-ID')}
                        ${traceActive && isRow2 ? ' 🎯' : ''}
                      </td>
                    </tr>
                  `;
                }).join('')}

                <!-- Summary Row -->
                <tr class="excel-data-row" style="background:#f8fafc; font-weight:700;">
                  <td class="excel-row-num">5</td>
                  <td colspan="3" class="excel-cell" style="text-align:right; color:var(--text-main);">
                    Grand Total Pengeluaran:
                  </td>
                  <td class="excel-cell" style="text-align:right; font-family:${showFormulas ? 'var(--font-mono)' : 'inherit'}; ${showFormulas ? 'background:#ede9fe;color:#6d28d9;' : ''}">
                    ${showFormulas ? '=AVERAGE(D2:D4)' : '2.683.333'}
                  </td>
                  <td class="excel-cell" style="text-align:right; font-family:${showFormulas ? 'var(--font-mono)' : 'inherit'};">
                    ${showFormulas ? 'SUM DISKON' : '665.000'}
                  </td>
                  <td class="excel-cell" style="text-align:right; font-weight:800; color:var(--excel-green); background:#f0fdf4; font-family:${showFormulas ? 'var(--font-mono)' : 'inherit'};">
                    ${showFormulas ? '=SUM(F2:F4)' : '7.385.000'}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          <!-- Audit Notification Banner -->
          <div class="excel-statusbar" style="flex-direction:column; align-items:flex-start; gap:4px;">
            ${showFormulas ? `
              <div style="font-size:0.82rem; color:#0369a1; font-weight:600; display:flex; align-items:center; gap:6px;">
                <span>👁️ MODE SHOW FORMULAS AKTIF (Ctrl + \`):</span> Seluruh angka kalkulasi disingkap menjadi rumus teks aslinya! Kamu dapat memeriksa apakah ada sel yang diketik angka mati (hardcoded) atau rumus yang salah ketik.
              </div>
            ` : traceActive ? `
              <div style="font-size:0.82rem; color:var(--excel-green); font-weight:600; display:flex; align-items:center; gap:6px;">
                <span>🎯 TRACE PRECEDENTS:</span> Rumus Total di sel F2 (=D2-E2) mengambil data secara langsung dari sel D2 (Subtotal) dan sel E2 (Diskon).
              </div>
            ` : `
              <div style="font-size:0.82rem; color:var(--text-muted);">
                💡 Tekan tombol <strong>"Show Formulas"</strong> di atas atau tekan tombol <strong>Ctrl + \`</strong> pada keyboard untuk melihat teks rumus di setiap sel.
              </div>
            `}
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('lab46-btn-toggle')?.addEventListener('click', () => {
        showFormulas = !showFormulas;
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });

      document.getElementById('lab46-btn-trace')?.addEventListener('click', () => {
        traceActive = !traceActive;
        SpreadsheetSimulatorsCh4.playClickSound();
        render();
      });
    }

    render();
  },

  // Initialize all Chapter 4 simulators
  initAll() {
    this.initLab41();
    this.initLab42();
    this.initLab43();
    this.initLab44();
    this.initLab45();
    this.initLab46();
  }
};

window.SpreadsheetSimulatorsCh4 = SpreadsheetSimulatorsCh4;
