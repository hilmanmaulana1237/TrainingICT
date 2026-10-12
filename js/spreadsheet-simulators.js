/**
 * EXCEL SPREADSHEET INTERACTIVE SIMULATORS - CHAPTER 2
 * Full interactive simulated Excel components for ICT teaching
 */

const SpreadsheetSimulators = {
  // Utility audio click
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
    '2.1': {
      title: 'Latihan 2.1: Find & Replace for Bulk Data',
      sheetName: 'Data_Inventaris',
      data: "ID\tKode Barang\tNama Produk\tQty\tStatus\n" +
            "A01\tCAB-OLD-01\tLaptop Dell\t14\tAktif\n" +
            "A02\tCAB-OLD-02\tMonitor Samsung\t22\tAktif\n" +
            "A03\tCAB-01-09\tPrinter Canon\t8\tPerbaikan\n" +
            "A04\tCAB-OLD-03\tKeyboard Logi\t45\tAktif\n" +
            "A05\tCAB-02-11\tMouse Wireless\t30\tAktif\n" +
            "A06\tCAB-OLD-04\tScanner Fujitsu\t5\tGudang"
    },
    '2.2': {
      title: 'Latihan 2.2: Import External Data (.CSV/.TXT)',
      sheetName: 'Import_CSV',
      data: "ID\tNama Siswa\tKelas\tNilai ICT\tAsal Kota\n" +
            "001\tAndi Pratama\t10-RPL\t88\tJakarta\n" +
            "002\tBunga Citra\t10-TKJ\t92\tBandung\n" +
            "003\tCitra Dewi\t10-RPL\t79\tSurabaya\n" +
            "004\tDimas Anggara\t10-DKV\t85\tYogyakarta\n" +
            "005\tEvelyn Tan\t10-TKJ\t95\tSemarang"
    },
    '2.3': {
      title: 'Latihan 2.3: Hyperlinks & Navigation Hub',
      sheetName: 'Daftar_Tautan',
      data: "Menu\tTautan Navigasi\tKeterangan\n" +
            "Dokumentasi Resmi\thttps://kemdikbud.go.id\tWebsite Kemdikbud\n" +
            "Tautan Aktif\t#Sheet2!A1\tLompat ke Sheet 2\n" +
            "Hubungi Admin\tmailto:admin@ict.sch.id\tKirim Email"
    },
    '2.4': {
      title: 'Latihan 2.4: Multi-Level Sorting & Custom Order',
      sheetName: 'Data_Karyawan',
      data: "ID\tNama Karyawan\tDepartemen\tGaji Pokok\tJabatan\n" +
            "EMP-01\tBudi Santoso\tMarketing\t6500000\tStaff\n" +
            "EMP-02\tSiti Rahma\tKeuangan\t8000000\tSupervisor\n" +
            "EMP-03\tAhmad Fauzi\tIT\t9500000\tManager\n" +
            "EMP-04\tDewi Lestari\tMarketing\t8200000\tSupervisor\n" +
            "EMP-05\tRian Pratama\tIT\t6800000\tStaff\n" +
            "EMP-06\tMaya Anggraini\tKeuangan\t6000000\tStaff\n" +
            "EMP-07\tFajar Hidayat\tIT\t11000000\tLead\n" +
            "EMP-08\tNabila Putri\tMarketing\t5500000\tStaff"
    },
    '2.5': {
      title: 'Latihan 2.5: Remove Duplicates with Multi-Column Criteria',
      sheetName: 'Data_Siswa',
      data: "ID\tNISN\tNama Siswa\tKota Asal\n" +
            "S-101\t00511234\tAditya Pratama\tJakarta\n" +
            "S-102\t00511235\tBella Safitri\tBandung\n" +
            "S-103\t00511236\tCahyo Nugroho\tSurabaya\n" +
            "S-102\t00511235\tBella Safitri\tBandung\n" +
            "S-104\t00511237\tDinda Kirana\tJakarta\n" +
            "S-105\t00511238\tEko Purwanto\tSemarang\n" +
            "S-101\t00511234\tAditya Pratama\tJakarta"
    },
    '2.6': {
      title: 'Latihan 2.6: Dynamic Conditional Formatting',
      sheetName: 'Nilai_Siswa',
      data: "No\tNama Siswa\tNilai Ujian\n" +
            "1\tAditya Pratama\t85\n" +
            "2\tBella Safitri\t58\n" +
            "3\tCahyo Nugroho\t92\n" +
            "4\tDinda Kirana\t72\n" +
            "5\tEko Purwanto\t52\n" +
            "6\tFani Maulida\t96\n" +
            "7\tGilang Ramadhan\t68\n" +
            "8\tHesti Wulandari\t86"
    },
    '2.7': {
      title: 'Latihan 2.7: Paste Special (Values, Formulas, Formats)',
      sheetName: 'Paste_Special',
      data: "Item\tQty\tHarga\tTotal\n" +
            "Flashdisk 32GB\t5\t60000\t300000\n" +
            "SSD 512GB\t2\t250000\t500000"
    },
    '2.8': {
      title: 'Latihan 2.8: Transpose & Safe Reordering (Insert Cut Cells)',
      sheetName: 'Laporan_Triwulan',
      data: "Produk\tQ1 (Jan-Mar)\tQ2 (Apr-Jun)\tQ3 (Jul-Sep)\tQ4 (Okt-Des)\n" +
            "Laptop ThinkPad\t120 unit\t150 unit\t180 unit\t210 unit\n" +
            "Monitor Curved\t85 unit\t90 unit\t110 unit\t140 unit\n" +
            "Printer Laser\t45 unit\t60 unit\t55 unit\t70 unit"
    },
    '2.9': {
      title: 'Latihan 2.9: Fill Handle & Automatic Series',
      sheetName: 'Deret_Otomatis',
      data: "Nomor\tNama Hari\tKode Siswa\n" +
            "1\tSenin\tSISWA-01\n" +
            "2\tSelasa\tSISWA-02\n" +
            "3\tRabu\tSISWA-03\n" +
            "4\tKamis\tSISWA-04\n" +
            "5\tJumat\tSISWA-05"
    },
    '2.10': {
      title: 'Latihan 2.10: Currency Format, Cell Styles & Text Orientation',
      sheetName: 'Format_Keuangan',
      data: "Komponen Biaya\tNominal Tarif\tStatus\n" +
            "SPP Bulanan\t350000\tLunas\n" +
            "Uang Gedung\t2500000\tBelum\n" +
            "Seragam & Buku\t750000\tLunas\n" +
            "Praktikum Komputer\t150000\tLunas"
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
  // LAB 2.1: FIND & REPLACE SIMULATOR
  // ==========================================
  initLab21() {
    const container = document.getElementById('lab-2-1-container');
    if (!container) return;

    const initialData = [
      { id: "A01", code: "CAB-OLD-01", name: "Laptop Dell", qty: 14, status: "Aktif" },
      { id: "A02", code: "CAB-OLD-02", name: "Monitor Samsung", qty: 22, status: "Aktif" },
      { id: "A03", code: "CAB-01-09", name: "Printer Canon", qty: 8, status: "Perbaikan" },
      { id: "A04", code: "CAB-OLD-03", name: "Keyboard Logi", qty: 45, status: "Aktif" },
      { id: "A05", code: "CAB-02-11", name: "Mouse Wireless", qty: 30, status: "Aktif" },
      { id: "A06", code: "CAB-OLD-04", name: "Scanner Fujitsu", qty: 5, status: "Gudang" }
    ];

    let currentData = JSON.parse(JSON.stringify(initialData));

    function render() {
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 2.1: Find & Replace for Bulk Data</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Shortcut: Ctrl + H</div>
          </div>

          <div class="excel-toolbar">
            <span style="font-size:0.8rem; font-weight:700; color:var(--text-muted); margin-right:8px;">Pencarian & Penggantian:</span>
            <div class="sim-input-group">
              <span class="sim-label" style="min-width:auto;">Find What:</span>
              <input type="text" id="lab21-find" class="sim-input" value="CAB-OLD" style="width:110px;">
            </div>
            <div class="sim-input-group">
              <span class="sim-label" style="min-width:auto;">Replace With:</span>
              <input type="text" id="lab21-replace" class="sim-input" value="CAB-NEW" style="width:110px;">
            </div>
            <button class="btn btn-primary" id="lab21-btn-replace-all" style="padding:6px 14px; font-size:0.82rem;">
              ⚡ Replace All
            </button>
            <button class="btn btn-outline" id="lab21-btn-reset" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Reset Data
            </button>
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulators.copySimulatorByLabId('2.1')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div class="excel-formula-bar">
            <div class="excel-name-box">B2</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="CAB-OLD-01" readonly>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr>
                  <th class="excel-corner-header"></th>
                  <th class="excel-col-header">A (ID)</th>
                  <th class="excel-col-header" style="width:160px;">B (Kode Barang)</th>
                  <th class="excel-col-header" style="width:160px;">C (Nama Produk)</th>
                  <th class="excel-col-header">D (Qty)</th>
                  <th class="excel-col-header">E (Status)</th>
                </tr>
              </thead>
              <tbody id="lab21-tbody">
                ${currentData.map((row, idx) => `
                  <tr>
                    <td class="excel-row-header">${idx + 1}</td>
                    <td class="excel-cell">${row.id}</td>
                    <td class="excel-cell ${row.code.includes('CAB-NEW') ? 'flash-replace' : ''}" id="cell-code-${idx}" style="font-family:var(--font-mono); font-weight:600;">
                      ${row.code}
                    </td>
                    <td class="excel-cell">${row.name}</td>
                    <td class="excel-cell" style="text-align:right;">${row.qty}</td>
                    <td class="excel-cell">${row.status}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">Data_Inventaris</div>
            </div>
            <div id="lab21-toast" style="font-weight:600; color:var(--excel-green);">
              Siap memindai 6 baris data.
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('lab21-btn-replace-all').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        const findVal = document.getElementById('lab21-find').value.trim();
        const replaceVal = document.getElementById('lab21-replace').value.trim();

        if (!findVal) {
          alert("Silakan masukkan teks yang ingin dicari di kolom 'Find What'!");
          return;
        }

        let replaceCount = 0;
        currentData.forEach((row, idx) => {
          if (row.code.includes(findVal)) {
            row.code = row.code.replaceAll(findVal, replaceVal);
            replaceCount++;
          }
        });

        render();

        const toast = document.getElementById('lab21-toast');
        if (replaceCount > 0) {
          SpreadsheetSimulators.playSuccessSound();
          toast.innerHTML = `🎉 Selesai! Excel telah mengganti <strong>${replaceCount} kemunculan</strong> kode secara serentak.`;
          toast.style.color = "var(--excel-green)";
        } else {
          toast.innerText = `Tidak ada teks '${findVal}' yang ditemukan pada tabel.`;
          toast.style.color = "var(--accent-rose)";
        }
      });

      document.getElementById('lab21-btn-reset').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        currentData = JSON.parse(JSON.stringify(initialData));
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 2.2: CSV IMPORT & DELIMITER VISUALIZER
  // ==========================================
  initLab22() {
    const container = document.getElementById('lab-2-2-container');
    if (!container) return;

    const rawCSV = `ID,Nama_Siswa,Kelas,Nilai_ICT,Asal_Kota
001,Andi Pratama,10-RPL,88,Jakarta
002,Bunga Citra,10-TKJ,92,Bandung
003,Citra Dewi,10-RPL,79,Surabaya
004,Dimas Anggara,10-DKV,85,Yogyakarta
005,Evelyn Tan,10-TKJ,95,Semarang`;

    let activeDelimiter = ',';

    function render() {
      // Split raw CSV lines
      const lines = rawCSV.trim().split('\n');
      
      // Parse columns according to activeDelimiter
      let parsedRows = [];
      lines.forEach(line => {
        let cols = line.split(activeDelimiter);
        parsedRows.push(cols);
      });

      const maxCols = Math.max(...parsedRows.map(r => r.length));
      const colNames = ["A", "B", "C", "D", "E", "F", "G"];

      const isProperlyParsed = activeDelimiter === ',';

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 2.2: Importing External Data & Delimiter Engine</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Menu: Data → Get Data → From Text/CSV</div>
          </div>

          <div class="excel-toolbar">
            <span style="font-weight:700; font-size:0.82rem; color:var(--text-muted);">Pilih Delimiter Pemisah:</span>
            <button class="excel-btn-sm ${activeDelimiter === ',' ? 'active' : ''}" data-del=",">
              Koma ( , ) <span class="delimiter-pill">Standar CSV</span>
            </button>
            <button class="excel-btn-sm ${activeDelimiter === ';' ? 'active' : ''}" data-del=";">
              Titik Koma ( ; ) <span class="delimiter-pill">Windows ID</span>
            </button>
            <button class="excel-btn-sm ${activeDelimiter === '\t' ? 'active' : ''}" data-del="\t">
              Tab ( \\t )
            </button>
            <button class="excel-btn-sm ${activeDelimiter === '|' ? 'active' : ''}" data-del="|">
              Pipe ( | )
            </button>
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulators.copySimulatorByLabId('2.2')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div style="padding:10px 16px; background:var(--bg-subtle); border-bottom:1px solid var(--border-color); display:flex; gap:16px; align-items:center; flex-wrap:wrap;">
            <div style="font-size:0.8rem; color:var(--text-muted); font-family:var(--font-mono);">
              📄 <strong>Teks Mentah Asli:</strong> ID,Nama_Siswa,Kelas,Nilai_ICT,Asal_Kota
            </div>
            ${!isProperlyParsed ? `
              <div style="font-size:0.8rem; color:var(--accent-rose); font-weight:700; background:#fee2e2; padding:3px 10px; border-radius:4px;">
                ⚠️ Delimiter salah! Seluruh teks menumpuk di Kolom A (Kesalahan umum di Windows Indonesia).
              </div>
            ` : `
              <div style="font-size:0.8rem; color:var(--excel-green); font-weight:700; background:var(--excel-green-light); padding:3px 10px; border-radius:4px;">
                ✔ Delimiter cocok! Data terbagi rapi menjadi ${maxCols} kolom terpisah.
              </div>
            `}
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr>
                  <th class="excel-corner-header"></th>
                  ${Array.from({ length: maxCols }).map((_, i) => `
                    <th class="excel-col-header" style="min-width:140px;">Kolom ${colNames[i]}</th>
                  `).join('')}
                </tr>
              </thead>
              <tbody>
                ${parsedRows.map((row, rIdx) => `
                  <tr>
                    <td class="excel-row-header">${rIdx + 1}</td>
                    ${Array.from({ length: maxCols }).map((_, cIdx) => `
                      <td class="excel-cell" style="${rIdx === 0 ? 'font-weight:700; background:var(--bg-subtle);' : ''}">
                        ${row[cIdx] !== undefined ? row[cIdx] : ''}
                      </td>
                    `).join('')}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">Impor_CSV</div>
            </div>
            <div style="font-size:0.8rem;">
              Total: <strong>5 baris</strong> diimpor dari file eksternal .csv
            </div>
          </div>
        </div>
      `;

      // Delimiter Buttons Listener
      container.querySelectorAll('.excel-btn-sm[data-del]').forEach(btn => {
        btn.addEventListener('click', (e) => {
          SpreadsheetSimulators.playClickSound();
          activeDelimiter = btn.getAttribute('data-del');
          render();
        });
      });
    }

    render();
  },

  // ==========================================
  // LAB 2.3: HYPERLINKS IN CELLS SIMULATOR
  // ==========================================
  initLab23() {
    const container = document.getElementById('lab-2-3-container');
    if (!container) return;

    let linkData = {
      displayText: "Buka Portal Kemdikbud",
      linkType: "web",
      targetAddress: "https://kemdikbud.go.id",
      screenTip: "Klik untuk membuka situs resmi Kementerian Pendidikan"
    };

    function render() {
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 2.3: Hyperlinks in Cells (Insert > Link)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Shortcut: Ctrl + K</div>
          </div>

          <div class="excel-toolbar">
            <div class="sim-row" style="width:100%;">
              <div class="sim-input-group">
                <span class="sim-label">Tipe Link:</span>
                <select id="lab23-type" class="sim-select">
                  <option value="web" ${linkData.linkType === 'web' ? 'selected' : ''}>Website (URL https://)</option>
                  <option value="mail" ${linkData.linkType === 'mail' ? 'selected' : ''}>E-mail (mailto:)</option>
                  <option value="sheet" ${linkData.linkType === 'sheet' ? 'selected' : ''}>Sheet Lain (Place in Document)</option>
                </select>
              </div>
              <div class="sim-input-group">
                <span class="sim-label">Text to Display:</span>
                <input type="text" id="lab23-text" class="sim-input" value="${linkData.displayText}" style="width:170px;">
              </div>
              <div class="sim-input-group">
                <span class="sim-label">Address / Tujuan:</span>
                <input type="text" id="lab23-address" class="sim-input" value="${linkData.targetAddress}" style="width:190px;">
              </div>
              <button class="btn btn-primary" id="lab23-btn-apply" style="padding:6px 14px; font-size:0.82rem;">
                🔗 Pasang Link ke Sel B2
              </button>
              <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulators.copySimulatorByLabId('2.3')" title="Salin data tabel simulator ini ke Excel">
                📑 Salin ke Sheet Baru
              </button>
            </div>
          </div>

          <div class="excel-formula-bar">
            <div class="excel-name-box">B2</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value='=HYPERLINK("${linkData.targetAddress}", "${linkData.displayText}")' readonly>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr>
                  <th class="excel-corner-header"></th>
                  <th class="excel-col-header">A (Menu)</th>
                  <th class="excel-col-header" style="width:250px;">B (Tautan Navigasi)</th>
                  <th class="excel-col-header" style="width:200px;">C (Keterangan)</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td class="excel-row-header">1</td>
                  <td class="excel-cell" style="font-weight:700;">Dokumentasi Resmi</td>
                  <td class="excel-cell">Klik menu di bawah ini</td>
                  <td class="excel-cell">Teks biasa</td>
                </tr>
                <tr>
                  <td class="excel-row-header">2</td>
                  <td class="excel-cell" style="font-weight:700;">Tautan Aktif (B2)</td>
                  <td class="excel-cell selected excel-link-cell" id="lab23-cell-target" title="${linkData.screenTip} (${linkData.targetAddress})">
                    ${linkData.displayText}
                  </td>
                  <td class="excel-cell">Arahkan kursor mouse untuk melihat ScreenTip</td>
                </tr>
                <tr>
                  <td class="excel-row-header">3</td>
                  <td class="excel-cell" style="font-weight:700;">Hubungi Admin</td>
                  <td class="excel-cell excel-link-cell" onclick="alert('Membuka email ke: admin@ict.sch.id')">
                    Kirim Pesan ke Admin
                  </td>
                  <td class="excel-cell">Link email (mailto:)</td>
                </tr>
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">Dashboard_Menu</div>
              <div class="excel-sheet-tab">Data_Nilai</div>
              <div class="excel-sheet-tab">Grafik</div>
            </div>
            <div id="lab23-toast" style="color:var(--text-muted); font-size:0.8rem;">
              Uji coba: Klik sel B2 untuk mensimulasikan navigasi hyperlink!
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('lab23-type').addEventListener('change', (e) => {
        const val = e.target.value;
        linkData.linkType = val;
        if (val === 'web') {
          linkData.displayText = "Buka Portal Kemdikbud";
          linkData.targetAddress = "https://kemdikbud.go.id";
          linkData.screenTip = "Kunjungi portal resmi Kementerian";
        } else if (val === 'mail') {
          linkData.displayText = "Hubungi Guru Pengampu";
          linkData.targetAddress = "mailto:guru.ict@sekolah.sch.id";
          linkData.screenTip = "Kirim pesan email ke guru pengampu";
        } else {
          linkData.displayText = "Lompat ke Sheet Grafik";
          linkData.targetAddress = "#Grafik!A1";
          linkData.screenTip = "Melompat langsung ke lembar kerja grafik";
        }
        render();
      });

      document.getElementById('lab23-btn-apply').addEventListener('click', () => {
        SpreadsheetSimulators.playSuccessSound();
        linkData.displayText = document.getElementById('lab23-text').value;
        linkData.targetAddress = document.getElementById('lab23-address').value;
        render();
        const toast = document.getElementById('lab23-toast');
        toast.innerHTML = `✔ Hyperlink berhasil dipasang pada sel <strong>B2</strong>!`;
        toast.style.color = "var(--excel-green)";
      });

      const cellTarget = document.getElementById('lab23-cell-target');
      if (cellTarget) {
        cellTarget.addEventListener('click', () => {
          SpreadsheetSimulators.playClickSound();
          alert(`[SIMULASI EXCEL HYPERLINK]\nExcel melompat ke alamat:\n${linkData.targetAddress}\n\nPerhatikan bahwa tampilan sel tetap bersih dan rapi ('${linkData.displayText}') tanpa mengotori tabel dengan teks link panjang!`);
        });
      }
    }

    render();
  },

  // ==========================================
  // LAB 2.4: MULTI-LEVEL SORTING SIMULATOR
  // ==========================================
  initLab24() {
    const container = document.getElementById('lab-2-4-container');
    if (!container) return;

    const initialEmployees = [
      { id: "EMP-01", name: "Budi Santoso", dept: "Marketing", salary: 6500000, level: "Staff" },
      { id: "EMP-02", name: "Siti Rahma", dept: "Keuangan", salary: 8000000, level: "Supervisor" },
      { id: "EMP-03", name: "Ahmad Fauzi", dept: "IT", salary: 9500000, level: "Manager" },
      { id: "EMP-04", name: "Dewi Lestari", dept: "Marketing", salary: 8200000, level: "Supervisor" },
      { id: "EMP-05", name: "Rian Pratama", dept: "IT", salary: 6800000, level: "Staff" },
      { id: "EMP-06", name: "Maya Anggraini", dept: "Keuangan", salary: 6000000, level: "Staff" },
      { id: "EMP-07", name: "Fajar Hidayat", dept: "IT", salary: 11000000, level: "Lead" },
      { id: "EMP-08", name: "Nabila Putri", dept: "Marketing", salary: 5500000, level: "Staff" }
    ];

    let currentEmployees = JSON.parse(JSON.stringify(initialEmployees));
    let sortConfig = {
      level1Col: "dept",
      level1Order: "asc",
      level2Col: "salary",
      level2Order: "desc"
    };

    function sortData() {
      currentEmployees.sort((a, b) => {
        // Level 1 Comparison
        let val1A = a[sortConfig.level1Col];
        let val1B = b[sortConfig.level1Col];
        let diff1 = 0;

        if (typeof val1A === 'string') {
          diff1 = val1A.localeCompare(val1B);
        } else {
          diff1 = val1A - val1B;
        }
        if (sortConfig.level1Order === 'desc') diff1 = -diff1;

        if (diff1 !== 0) return diff1;

        // Level 2 Comparison (Tie-breaker)
        let val2A = a[sortConfig.level2Col];
        let val2B = b[sortConfig.level2Col];
        let diff2 = 0;

        if (typeof val2A === 'string') {
          diff2 = val2A.localeCompare(val2B);
        } else {
          diff2 = val2A - val2B;
        }
        if (sortConfig.level2Order === 'desc') diff2 = -diff2;

        return diff2;
      });
    }

    function render() {
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 2.4: Multi-Level Sorting (Data → Sort)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Add Level & Custom Hierarchy</div>
          </div>

          <div class="excel-toolbar">
            <div style="display:flex; flex-direction:column; gap:8px; width:100%;">
              <div style="display:flex; align-items:center; gap:12px; flex-wrap:wrap;">
                <div style="font-weight:700; font-size:0.85rem; color:var(--text-main); display:flex; align-items:center; gap:6px;">
                  <span style="background:var(--excel-green); color:white; border-radius:50%; width:20px; height:20px; display:inline-flex; align-items:center; justify-content:center; font-size:0.75rem;">1</span>
                  Sort by:
                </div>
                <select id="sort-l1-col" class="sim-select">
                  <option value="dept" ${sortConfig.level1Col === 'dept' ? 'selected' : ''}>Departemen</option>
                  <option value="salary" ${sortConfig.level1Col === 'salary' ? 'selected' : ''}>Gaji Pokok</option>
                  <option value="name" ${sortConfig.level1Col === 'name' ? 'selected' : ''}>Nama Karyawan</option>
                </select>
                <select id="sort-l1-ord" class="sim-select">
                  <option value="asc" ${sortConfig.level1Order === 'asc' ? 'selected' : ''}>A to Z (Ascending)</option>
                  <option value="desc" ${sortConfig.level1Order === 'desc' ? 'selected' : ''}>Z to A (Descending)</option>
                </select>

                <div style="font-weight:700; font-size:0.85rem; color:var(--text-main); display:flex; align-items:center; gap:6px; margin-left:14px;">
                  <span style="background:var(--accent-blue); color:white; border-radius:50%; width:20px; height:20px; display:inline-flex; align-items:center; justify-content:center; font-size:0.75rem;">2</span>
                  Then by:
                </div>
                <select id="sort-l2-col" class="sim-select">
                  <option value="salary" ${sortConfig.level2Col === 'salary' ? 'selected' : ''}>Gaji Pokok</option>
                  <option value="name" ${sortConfig.level2Col === 'name' ? 'selected' : ''}>Nama Karyawan</option>
                  <option value="dept" ${sortConfig.level2Col === 'dept' ? 'selected' : ''}>Departemen</option>
                </select>
                <select id="sort-l2-ord" class="sim-select">
                  <option value="desc" ${sortConfig.level2Order === 'desc' ? 'selected' : ''}>Terbesar ke Terkecil (Descending)</option>
                  <option value="asc" ${sortConfig.level2Order === 'asc' ? 'selected' : ''}>Terkecil ke Terbesar (Ascending)</option>
                </select>

                <button class="btn btn-primary" id="btn-apply-sort" style="padding:6px 14px; font-size:0.82rem; margin-left:auto;">
                  ⚡ Terapkan Sort Bertingkat
                </button>
                <button class="btn btn-outline" id="btn-reset-sort" style="padding:6px 12px; font-size:0.82rem;">
                  ↺ Reset
                </button>
                <button class="excel-btn-sm" onclick="SpreadsheetSimulators.copySimulatorByLabId('2.4')" title="Salin data tabel simulator ini ke Excel">
                  📑 Salin ke Sheet Baru
                </button>
              </div>
            </div>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr>
                  <th class="excel-corner-header"></th>
                  <th class="excel-col-header">A (ID)</th>
                  <th class="excel-col-header" style="width:160px;">B (Nama Karyawan)</th>
                  <th class="excel-col-header" style="width:140px; background:${sortConfig.level1Col === 'dept' ? '#dcfce7' : ''};">
                    C (Departemen) ${sortConfig.level1Col === 'dept' ? '★ L1' : ''}
                  </th>
                  <th class="excel-col-header" style="width:140px; background:${sortConfig.level2Col === 'salary' ? '#e0f2fe' : ''};">
                    D (Gaji Pokok) ${sortConfig.level2Col === 'salary' ? '★ L2' : ''}
                  </th>
                  <th class="excel-col-header">E (Jabatan)</th>
                </tr>
              </thead>
              <tbody>
                ${currentEmployees.map((emp, idx) => `
                  <tr>
                    <td class="excel-row-header">${idx + 1}</td>
                    <td class="excel-cell">${emp.id}</td>
                    <td class="excel-cell" style="font-weight:600;">${emp.name}</td>
                    <td class="excel-cell" style="background:${emp.dept === 'IT' ? 'rgba(56,189,248,0.1)' : emp.dept === 'Keuangan' ? 'rgba(245,158,11,0.1)' : 'rgba(16,185,129,0.1)'};">
                      ${emp.dept}
                    </td>
                    <td class="excel-cell" style="font-family:var(--font-mono); text-align:right;">
                      Rp ${emp.salary.toLocaleString('id-ID')}
                    </td>
                    <td class="excel-cell">${emp.level}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">Data_Karyawan</div>
            </div>
            <div id="sort-status-text" style="font-size:0.8rem; color:var(--excel-green); font-weight:600;">
              Urutan aktif: Kelompok Departemen (${sortConfig.level1Order.toUpperCase()}), lalu Gaji (${sortConfig.level2Order.toUpperCase()}).
            </div>
          </div>
        </div>
      `;

      // Event listeners
      document.getElementById('btn-apply-sort').addEventListener('click', () => {
        SpreadsheetSimulators.playSuccessSound();
        sortConfig.level1Col = document.getElementById('sort-l1-col').value;
        sortConfig.level1Order = document.getElementById('sort-l1-ord').value;
        sortConfig.level2Col = document.getElementById('sort-l2-col').value;
        sortConfig.level2Order = document.getElementById('sort-l2-ord').value;
        sortData();
        render();
      });

      document.getElementById('btn-reset-sort').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        currentEmployees = JSON.parse(JSON.stringify(initialEmployees));
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 2.5: REMOVING DUPLICATE DATA SIMULATOR
  // ==========================================
  initLab25() {
    const container = document.getElementById('lab-2-5-container');
    if (!container) return;

    const initialDataset = [
      { id: "S-101", nisn: "00511234", name: "Aditya Pratama", city: "Jakarta", isDup: false },
      { id: "S-102", nisn: "00511235", name: "Bella Safitri", city: "Bandung", isDup: false },
      { id: "S-103", nisn: "00511236", name: "Cahyo Nugroho", city: "Surabaya", isDup: false },
      { id: "S-102", nisn: "00511235", name: "Bella Safitri", city: "Bandung", isDup: true }, // Duplikat Bella
      { id: "S-104", nisn: "00511237", name: "Dinda Kirana", city: "Jakarta", isDup: false },
      { id: "S-105", nisn: "00511238", name: "Eko Purwanto", city: "Semarang", isDup: false },
      { id: "S-101", nisn: "00511234", name: "Aditya Pratama", city: "Jakarta", isDup: true }  // Duplikat Aditya
    ];

    let currentDataset = JSON.parse(JSON.stringify(initialDataset));

    function render() {
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 2.5: Removing Duplicate Data (Data → Remove Duplicates)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Shortcut: Alt + A + M</div>
          </div>

          <div class="excel-toolbar">
            <div style="display:flex; align-items:center; gap:14px; flex-wrap:wrap; width:100%;">
              <span style="font-weight:700; font-size:0.82rem; color:var(--text-muted);">Pilih Kolom Kriteria Evaluasi Duplikat:</span>
              <label style="font-size:0.82rem; display:flex; align-items:center; gap:4px; cursor:pointer;">
                <input type="checkbox" id="chk-nisn" checked> <strong>NISN (ID Unik)</strong>
              </label>
              <label style="font-size:0.82rem; display:flex; align-items:center; gap:4px; cursor:pointer;">
                <input type="checkbox" id="chk-name" checked> Nama Siswa
              </label>
              <label style="font-size:0.82rem; display:flex; align-items:center; gap:4px; cursor:pointer;">
                <input type="checkbox" id="chk-city" checked> Kota Asal
              </label>

              <button class="btn btn-primary" id="btn-remove-dup" style="padding:6px 14px; font-size:0.82rem; margin-left:auto;">
                🗑 Hapus Data Duplikat
              </button>
              <button class="btn btn-outline" id="btn-reset-dup" style="padding:6px 12px; font-size:0.82rem;">
                ↺ Kembalikan Data Asal
              </button>
              <button class="excel-btn-sm" onclick="SpreadsheetSimulators.copySimulatorByLabId('2.5')" title="Salin data tabel simulator ini ke Excel">
                📑 Salin ke Sheet Baru
              </button>
            </div>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr>
                  <th class="excel-corner-header"></th>
                  <th class="excel-col-header">A (ID)</th>
                  <th class="excel-col-header" style="width:130px;">B (NISN)</th>
                  <th class="excel-col-header" style="width:180px;">C (Nama Siswa)</th>
                  <th class="excel-col-header" style="width:140px;">D (Kota Asal)</th>
                  <th class="excel-col-header">E (Status Deteksi)</th>
                </tr>
              </thead>
              <tbody id="tbody-dup">
                ${currentDataset.map((row, idx) => `
                  <tr class="${row.isDup ? 'is-duplicate' : ''}">
                    <td class="excel-row-header">${idx + 1}</td>
                    <td class="excel-cell">${row.id}</td>
                    <td class="excel-cell" style="font-family:var(--font-mono); font-weight:700;">${row.nisn}</td>
                    <td class="excel-cell">${row.name}</td>
                    <td class="excel-cell">${row.city}</td>
                    <td class="excel-cell" style="font-weight:600; color:${row.isDup ? 'var(--accent-rose)' : 'var(--excel-green)'};">
                      ${row.isDup ? '⚠️ Duplikat (Baris Kembar)' : '✔ Unik'}
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">Pendaftaran_Siswa</div>
            </div>
            <div id="dup-toast-status" style="font-size:0.8rem; font-weight:600;">
              Total: <strong>${currentDataset.length} baris</strong> (Terdapat 2 baris terindikasi duplikat).
            </div>
          </div>
        </div>
      `;

      // Event listeners
      document.getElementById('btn-remove-dup').addEventListener('click', () => {
        SpreadsheetSimulators.playSuccessSound();
        const initialCount = currentDataset.length;
        // Filter out items marked as duplicates
        currentDataset = currentDataset.filter(item => !item.isDup);
        const removedCount = initialCount - currentDataset.length;

        render();

        alert(`[MICROSOFT EXCEL - DIALOG KONFIRMASI]\n\n${removedCount} nilai duplikat ditemukan dan dihapus;\n${currentDataset.length} nilai unik tetap dipertahankan.`);
        const toast = document.getElementById('dup-toast-status');
        toast.innerHTML = `🎉 Sukses! <strong>${removedCount} baris duplikat</strong> berhasil dieliminasi. Tersisa ${currentDataset.length} data unik.`;
        toast.style.color = "var(--excel-green)";
      });

      document.getElementById('btn-reset-dup').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        currentDataset = JSON.parse(JSON.stringify(initialDataset));
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 2.6: CONDITIONAL FORMATTING SIMULATOR
  // ==========================================
  initLab26() {
    const container = document.getElementById('lab-2-6-container');
    if (!container) return;

    let students = [
      { id: 1, name: "Aditya Pratama", score: 85 },
      { id: 2, name: "Bella Safitri", score: 58 },
      { id: 3, name: "Cahyo Nugroho", score: 92 },
      { id: 4, name: "Dinda Kirana", score: 72 },
      { id: 5, name: "Eko Purwanto", score: 52 },
      { id: 6, name: "Fani Maulida", score: 96 },
      { id: 7, name: "Gilang Ramadhan", score: 68 },
      { id: 8, name: "Hesti Wulandari", score: 86 }
    ];

    let activeRule = "below-avg"; // "below-avg", "top-3", "kkm", "none"
    let kkmValue = 75;

    function render() {
      // Calculate dynamic average
      const total = students.reduce((acc, s) => acc + s.score, 0);
      const avg = (total / students.length).toFixed(1);

      // Find top 3 threshold
      const sortedScores = [...students].map(s => s.score).sort((a, b) => b - a);
      const top3Cutoff = sortedScores[2];

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 2.6: Rule-Based Conditional Formatting (Below Average & Rules)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Home → Styles → Conditional Formatting</div>
          </div>

          <div class="excel-toolbar">
            <span style="font-weight:700; font-size:0.82rem; color:var(--text-muted);">Pilih Aturan Format:</span>
            <button class="excel-btn-sm ${activeRule === 'below-avg' ? 'active' : ''}" id="btn-rule-below">
              🔴 Below Average (< ${avg})
            </button>
            <button class="excel-btn-sm ${activeRule === 'top-3' ? 'active' : ''}" id="btn-rule-top">
              🟢 Top 3 Nilai Tertinggi
            </button>
            <button class="excel-btn-sm ${activeRule === 'kkm' ? 'active' : ''}" id="btn-rule-kkm">
              🟠 Di Bawah KKM (< ${kkmValue})
            </button>
            <button class="excel-btn-sm ${activeRule === 'none' ? 'active' : ''}" id="btn-rule-clear">
              ✕ Clear Rules
            </button>

            ${activeRule === 'kkm' ? `
              <div style="display:inline-flex; align-items:center; gap:6px; margin-left:10px;">
                <span style="font-size:0.75rem; font-weight:600;">Slider KKM:</span>
                <input type="range" id="slider-kkm" min="60" max="85" value="${kkmValue}" style="width:90px;">
                <span style="font-family:var(--font-mono); font-weight:700;">${kkmValue}</span>
              </div>
            ` : ''}
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulators.copySimulatorByLabId('2.6')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div style="padding:8px 16px; background:var(--bg-subtle); border-bottom:1px solid var(--border-color); display:flex; justify-content:space-between; align-items:center;">
            <div style="font-size:0.82rem;">
              💡 <em>Tips Interaktif: Anda bisa mengubah angka nilai siswa langsung di tabel untuk melihat warna berubah seketika!</em>
            </div>
            <div style="font-family:var(--font-mono); font-size:0.82rem; font-weight:700; background:var(--bg-surface); padding:3px 10px; border-radius:4px; border:1px solid var(--border-color);">
              Rata-rata Kelas = <span style="color:var(--excel-green);">${avg}</span>
            </div>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr>
                  <th class="excel-corner-header"></th>
                  <th class="excel-col-header" style="width:60px;">A (No)</th>
                  <th class="excel-col-header" style="width:200px;">B (Nama Siswa)</th>
                  <th class="excel-col-header" style="width:140px;">C (Nilai Ujian)</th>
                  <th class="excel-col-header" style="width:220px;">D (Status Evaluasi Dinamis)</th>
                </tr>
              </thead>
              <tbody>
                ${students.map((student, idx) => {
                  let isBelowAvg = student.score < avg;
                  let isTop3 = student.score >= top3Cutoff;
                  let isBelowKKM = student.score < kkmValue;

                  let highlightClass = "";
                  let statusText = "Normal";

                  if (activeRule === 'below-avg' && isBelowAvg) {
                    highlightClass = "cf-below-avg";
                    statusText = `🔴 Di Bawah Rata-rata (-${(avg - student.score).toFixed(1)})`;
                  } else if (activeRule === 'top-3' && isTop3) {
                    highlightClass = "cf-top";
                    statusText = "🟢 Top 3 Berprestasi";
                  } else if (activeRule === 'kkm' && isBelowKKM) {
                    highlightClass = "cf-below-avg";
                    statusText = `🟠 Remedial (KKM ${kkmValue})`;
                  }

                  return `
                    <tr>
                      <td class="excel-row-header">${idx + 1}</td>
                      <td class="excel-cell" style="text-align:center;">${student.id}</td>
                      <td class="excel-cell" style="font-weight:600;">${student.name}</td>
                      <td class="excel-cell ${highlightClass}" style="text-align:right;">
                        <input type="number" class="student-score-input" data-idx="${idx}" value="${student.score}" min="0" max="100" style="width:55px; border:none; background:transparent; font-family:var(--font-mono); font-weight:700; text-align:right; color:inherit; outline:none;">
                      </td>
                      <td class="excel-cell ${highlightClass}">
                        ${statusText}
                      </td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">Nilai_Semester</div>
            </div>
            <div style="font-size:0.8rem; color:var(--text-muted);">
              Aturan aktif: <strong>${activeRule.toUpperCase()}</strong>
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('btn-rule-below').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        activeRule = 'below-avg';
        render();
      });

      document.getElementById('btn-rule-top').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        activeRule = 'top-3';
        render();
      });

      document.getElementById('btn-rule-kkm').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        activeRule = 'kkm';
        render();
      });

      document.getElementById('btn-rule-clear').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        activeRule = 'none';
        render();
      });

      const sliderKKM = document.getElementById('slider-kkm');
      if (sliderKKM) {
        sliderKKM.addEventListener('input', (e) => {
          kkmValue = parseInt(e.target.value);
          render();
        });
      }

      // Live editable score input
      container.querySelectorAll('.student-score-input').forEach(input => {
        input.addEventListener('change', (e) => {
          SpreadsheetSimulators.playClickSound();
          const idx = parseInt(e.target.getAttribute('data-idx'));
          const newVal = parseInt(e.target.value) || 0;
          students[idx].score = Math.max(0, Math.min(100, newVal));
          render();
        });
      });
    }

    render();
  },

  // ==========================================
  // LAB 2.7: PASTE SPECIAL STUDIO
  // ==========================================
  initLab27() {
    const container = document.getElementById('lab-2-7-container');
    if (!container) return;

    let targetState = {
      pasted: false,
      pasteType: "values", // "values", "formulas", "formats"
      formulaText: "500000",
      displayText: "Rp 500.000",
      customStyle: ""
    };

    function render() {
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 2.7: Paste Special (Ctrl + Alt + V)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Values vs Formulas vs Formats</div>
          </div>

          <div class="excel-toolbar">
            <button class="btn btn-primary" id="btn-copy-origin" style="padding:6px 14px; font-size:0.82rem;">
              📋 Step 1: Copy Sel Asal (D2)
            </button>
            <span style="font-size:0.82rem; font-weight:700; color:var(--text-muted); margin-left:8px;">Step 2: Pilih Jenis Paste Special ke Sel F2:</span>
            <button class="excel-btn-sm ${targetState.pasteType === 'values' ? 'active' : ''}" id="btn-paste-values">
              📄 Paste Values (Bekukan Rumus)
            </button>
            <button class="excel-btn-sm ${targetState.pasteType === 'formulas' ? 'active' : ''}" id="btn-paste-formulas">
              📐 Paste Formulas
            </button>
            <button class="excel-btn-sm ${targetState.pasteType === 'formats' ? 'active' : ''}" id="btn-paste-formats">
              🎨 Paste Formats Only
            </button>
            <button class="btn btn-outline" id="btn-reset-paste" style="padding:6px 12px; font-size:0.82rem; margin-left:auto;">
              ↺ Reset
            </button>
            <button class="excel-btn-sm" onclick="SpreadsheetSimulators.copySimulatorByLabId('2.7')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div class="excel-formula-bar">
            <div class="excel-name-box" id="paste-cell-focus">D2</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" id="paste-formula-display" value="=B2*C2" readonly>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr>
                  <th class="excel-corner-header"></th>
                  <th class="excel-col-header" style="width:140px;">A (Item)</th>
                  <th class="excel-col-header">B (Qty)</th>
                  <th class="excel-col-header" style="width:120px;">C (Harga)</th>
                  <th class="excel-col-header" style="width:140px; background:#e0f2fe;">D (Total Rumus)</th>
                  <th class="excel-col-header">E (Pemisah)</th>
                  <th class="excel-col-header" style="width:170px; background:#dcfce7;">F (Target Paste F2)</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td class="excel-row-header">1</td>
                  <td class="excel-cell">Flashdisk 32GB</td>
                  <td class="excel-cell" style="text-align:right;">5</td>
                  <td class="excel-cell" style="text-align:right;">Rp 60.000</td>
                  <td class="excel-cell" style="text-align:right;">Rp 300.000</td>
                  <td class="excel-cell" style="background:#f1f5f9;"></td>
                  <td class="excel-cell" style="color:var(--text-light); text-align:center;">(Sel Kosong)</td>
                </tr>
                <tr>
                  <td class="excel-row-header">2</td>
                  <td class="excel-cell" style="font-weight:700;">SSD 512GB</td>
                  <td class="excel-cell" style="text-align:right;">2</td>
                  <td class="excel-cell" style="text-align:right;">Rp 250.000</td>
                  <!-- Sel D2 Asal Rumus -->
                  <td class="excel-cell selected" id="cell-origin-d2" style="text-align:right; font-weight:700; background:rgba(2,132,199,0.15) !important;">
                    Rp 500.000
                  </td>
                  <td class="excel-cell" style="background:#f1f5f9;"></td>
                  <!-- Sel F2 Target Paste -->
                  <td class="excel-cell" id="cell-target-f2" style="text-align:right; font-weight:700; ${targetState.customStyle}">
                    ${targetState.pasted ? targetState.displayText : '<span style="color:var(--text-light); font-weight:normal; font-style:italic;">[Klik Paste...]</span>'}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">Penjualan_Hardware</div>
            </div>
            <div id="paste-toast-msg" style="font-size:0.8rem; font-weight:600; color:var(--text-muted);">
              ${targetState.pasted ? `Hasil di Sel F2: Formula bar berisi '${targetState.formulaText}'` : 'Klik "Step 1: Copy Sel Asal" untuk memulai simulasi.'}
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('btn-copy-origin').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        const originCell = document.getElementById('cell-origin-d2');
        originCell.style.outline = "2px dashed #0284c7";
        alert("Sel D2 disalin (Ctrl+C)! Garis putus-putus aktif menandakan sel siap ditempel via Paste Special.");
      });

      document.getElementById('btn-paste-values').addEventListener('click', () => {
        SpreadsheetSimulators.playSuccessSound();
        targetState = {
          pasted: true,
          pasteType: "values",
          formulaText: "500000",
          displayText: "500000",
          customStyle: "color:var(--text-main);"
        };
        render();
        document.getElementById('paste-cell-focus').innerText = "F2";
        document.getElementById('paste-formula-display').value = "500000";
        alert("[PASTE VALUES BERHASIL]\nRumus =B2*C2 berhasil dihilangkan! Di sel F2 kini tersimpan angka permanen '500000'. Sangat aman untuk dikirim tanpa takut error #REF!.");
      });

      document.getElementById('btn-paste-formulas').addEventListener('click', () => {
        SpreadsheetSimulators.playSuccessSound();
        targetState = {
          pasted: true,
          pasteType: "formulas",
          formulaText: "=D2*E2 (Referensi bergeser otomatis)",
          displayText: "Rp 0 (atau #REF!)",
          customStyle: "color:var(--accent-rose);"
        };
        render();
        document.getElementById('paste-cell-focus').innerText = "F2";
        document.getElementById('paste-formula-display').value = "=D2*E2";
        alert("[PASTE FORMULAS]\nExcel menyalin rumus relatif sehingga referensi bergeser ke kolom di sebelahnya. Inilah alasan mengapa rumus sering rusak jika tidak memakai Paste Values!");
      });

      document.getElementById('btn-paste-formats').addEventListener('click', () => {
        SpreadsheetSimulators.playSuccessSound();
        targetState = {
          pasted: true,
          pasteType: "formats",
          formulaText: "(Sel kosong tanpa rumus)",
          displayText: "(Hanya Warna Biru)",
          customStyle: "background:rgba(2,132,199,0.15) !important; color:#0284c7;"
        };
        render();
        document.getElementById('paste-cell-focus').innerText = "F2";
        document.getElementById('paste-formula-display').value = "";
      });

      document.getElementById('btn-reset-paste').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        targetState = {
          pasted: false,
          pasteType: "values",
          formulaText: "",
          displayText: "",
          customStyle: ""
        };
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 2.8: TRANSPOSE & CUT-INSERT SIMULATOR
  // ==========================================
  initLab28() {
    const container = document.getElementById('lab-2-8-container');
    if (!container) return;

    let isTransposed = false;

    // Normal (horizontal quarters across cols)
    const normalMatrix = {
      headers: ["Produk", "Q1 (Jan-Mar)", "Q2 (Apr-Jun)", "Q3 (Jul-Sep)", "Q4 (Okt-Des)"],
      rows: [
        ["Laptop ThinkPad", "120 unit", "150 unit", "180 unit", "210 unit"],
        ["Monitor Curved", "85 unit", "90 unit", "110 unit", "140 unit"],
        ["Printer Laser", "45 unit", "60 unit", "55 unit", "70 unit"]
      ]
    };

    // Transposed (quarters down rows, products across cols)
    const transposedMatrix = {
      headers: ["Periode", "Laptop ThinkPad", "Monitor Curved", "Printer Laser"],
      rows: [
        ["Q1 (Jan-Mar)", "120 unit", "85 unit", "45 unit"],
        ["Q2 (Apr-Jun)", "150 unit", "90 unit", "60 unit"],
        ["Q3 (Jul-Sep)", "180 unit", "110 unit", "55 unit"],
        ["Q4 (Okt-Des)", "210 unit", "140 unit", "70 unit"]
      ]
    };

    function render() {
      const activeData = isTransposed ? transposedMatrix : normalMatrix;

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 2.8: Transposing & Moving Data Layout</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Paste Special → Transpose</div>
          </div>

          <div class="excel-toolbar">
            <button class="btn btn-primary" id="btn-toggle-transpose" style="padding:6px 16px; font-size:0.82rem;">
              🔄 ${isTransposed ? 'Kembalikan Orientasi Asal' : 'Putar Tabel dengan Transpose (Baris ⇄ Kolom)'}
            </button>
            <button class="btn btn-outline" id="btn-demo-cut-insert" style="padding:6px 12px; font-size:0.82rem;">
              ✂ Demonstrasi "Insert Cut Cells"
            </button>
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulators.copySimulatorByLabId('2.8')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div style="padding:8px 16px; background:var(--bg-subtle); border-bottom:1px solid var(--border-color); font-size:0.82rem; color:var(--text-muted);">
            Orientasi Saat Ini: <strong>${isTransposed ? 'Format Vertikal (4 Baris Periode x 3 Kolom Produk)' : 'Format Horizontal (3 Baris Produk x 4 Kolom Periode)'}</strong>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr>
                  <th class="excel-corner-header"></th>
                  ${activeData.headers.map((h, i) => `
                    <th class="excel-col-header" style="min-width:140px; font-weight:700;">${String.fromCharCode(65 + i)} (${h})</th>
                  `).join('')}
                </tr>
              </thead>
              <tbody>
                ${activeData.rows.map((row, rIdx) => `
                  <tr>
                    <td class="excel-row-header">${rIdx + 1}</td>
                    ${row.map((cell, cIdx) => `
                      <td class="excel-cell ${cIdx === 0 ? 'cell-style-accent1' : ''}" style="${cIdx > 0 ? 'text-align:right;' : ''}">
                        ${cell}
                      </td>
                    `).join('')}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">Laporan_Triwulan</div>
            </div>
            <div style="font-size:0.8rem; color:var(--excel-green); font-weight:600;">
              ${isTransposed ? '✔ Berhasil ditransposisi! Baris telah menjadi kolom dan kolom telah menjadi baris.' : 'Tabel siap ditransposisi.'}
            </div>
          </div>
        </div>
      `;

      // Event listeners
      document.getElementById('btn-toggle-transpose').addEventListener('click', () => {
        SpreadsheetSimulators.playSuccessSound();
        isTransposed = !isTransposed;
        render();
      });

      document.getElementById('btn-demo-cut-insert').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        alert(`[TIPS MENGAJAR ICT - CUT vs INSERT CUT CELLS]\n\n1. Jika Anda melakukan Cut (Ctrl+X) lalu Paste biasa (Ctrl+V), data pada sel tujuan akan tertimpa dan terhapus hilang!\n\n2. Gunakan klik kanan > 'Insert Cut Cells': Excel akan secara cerdas menggeser baris/kolom tetangga ke bawah/samping tanpa menghilangkan satu pun data yang ada!`);
      });
    }

    render();
  },

  // ==========================================
  // LAB 2.9: FILL HANDLE & AUTO SERIES SIMULATOR
  // ==========================================
  initLab29() {
    const container = document.getElementById('lab-2-9-container');
    if (!container) return;

    let seriesItems = ["1", "2"];
    let seriesType = "numbers"; // "numbers", "days", "dates", "code"

    function generateNext() {
      if (seriesType === "numbers") {
        const nextNum = parseInt(seriesItems[seriesItems.length - 1]) + 1;
        seriesItems.push(nextNum.toString());
      } else if (seriesType === "days") {
        const days = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"];
        const nextDay = days[seriesItems.length % days.length];
        seriesItems.push(nextDay);
      } else if (seriesType === "dates") {
        const nextDate = `${seriesItems.length + 1} Okt 2026`;
        seriesItems.push(nextDate);
      } else if (seriesType === "code") {
        const nextCode = `SISWA-${String(seriesItems.length + 1).padStart(2, '0')}`;
        seriesItems.push(nextCode);
      }
    }

    function render() {
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 2.9: Fill Handle and Automatic Series</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Home → Editing → Fill → Series</div>
          </div>

          <div class="excel-toolbar">
            <span style="font-weight:700; font-size:0.82rem; color:var(--text-muted);">Pola Deret:</span>
            <button class="excel-btn-sm ${seriesType === 'numbers' ? 'active' : ''}" id="btn-pattern-num">
              1, 2, 3... (Angka Urut)
            </button>
            <button class="excel-btn-sm ${seriesType === 'days' ? 'active' : ''}" id="btn-pattern-day">
              Senin, Selasa... (Hari)
            </button>
            <button class="excel-btn-sm ${seriesType === 'dates' ? 'active' : ''}" id="btn-pattern-date">
              1 Okt, 2 Okt... (Tanggal)
            </button>
            <button class="excel-btn-sm ${seriesType === 'code' ? 'active' : ''}" id="btn-pattern-code">
              SISWA-01, SISWA-02... (Kombinasi)
            </button>

            <button class="btn btn-primary" id="btn-drag-handle" style="padding:6px 14px; font-size:0.82rem; margin-left:10px;">
              🖱️ Drag Fill Handle (+1 Baris)
            </button>
            <button class="btn btn-outline" id="btn-fill-series-dialog" style="padding:6px 12px; font-size:0.82rem;">
              ⚡ Fill > Series Dialog (Step +5)
            </button>
            <button class="btn btn-outline" id="btn-reset-series" style="padding:6px 12px; font-size:0.82rem;">
              ↺ Reset
            </button>
            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulators.copySimulatorByLabId('2.9')" title="Salin data tabel simulator ini ke Excel">
              📑 Salin ke Sheet Baru
            </button>
          </div>

          <div class="excel-formula-bar">
            <div class="excel-name-box">A${seriesItems.length}</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="${seriesItems[seriesItems.length - 1]}" readonly>
          </div>

          <div class="excel-grid-container">
            <table class="excel-table">
              <thead>
                <tr>
                  <th class="excel-corner-header"></th>
                  <th class="excel-col-header" style="width:200px;">A (Kolom Deret Otomatis)</th>
                  <th class="excel-col-header" style="width:250px;">B (Keterangan Pola Excel)</th>
                </tr>
              </thead>
              <tbody>
                ${seriesItems.map((val, idx) => {
                  const isLast = idx === seriesItems.length - 1;
                  return `
                    <tr>
                      <td class="excel-row-header">${idx + 1}</td>
                      <td class="excel-cell ${isLast ? 'selected' : ''}" style="font-weight:700; font-family:var(--font-mono);">
                        ${val}
                      </td>
                      <td class="excel-cell" style="color:var(--text-muted); font-size:0.8rem;">
                        ${idx === 0 ? 'Nilai awal' : idx === 1 ? 'Penentu interval' : 'Dihasilkan otomatis oleh Fill Handle'}
                      </td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">Deret_Otomatis</div>
            </div>
            <div style="font-size:0.8rem; color:var(--excel-green); font-weight:600;">
              Total: <strong>${seriesItems.length} sel terisi</strong>. Double-click kotak hijau untuk auto-fill kilat!
            </div>
          </div>
        </div>
      `;

      // Event listeners
      document.getElementById('btn-drag-handle').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        if (seriesItems.length >= 12) {
          alert("Sudah mencapai 12 baris deret!");
          return;
        }
        generateNext();
        render();
      });

      document.getElementById('btn-pattern-num').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        seriesType = "numbers";
        seriesItems = ["1", "2"];
        render();
      });

      document.getElementById('btn-pattern-day').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        seriesType = "days";
        seriesItems = ["Senin", "Selasa"];
        render();
      });

      document.getElementById('btn-pattern-date').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        seriesType = "dates";
        seriesItems = ["1 Okt 2026", "2 Okt 2026"];
        render();
      });

      document.getElementById('btn-pattern-code').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        seriesType = "code";
        seriesItems = ["SISWA-01", "SISWA-02"];
        render();
      });

      document.getElementById('btn-fill-series-dialog').addEventListener('click', () => {
        SpreadsheetSimulators.playSuccessSound();
        seriesType = "numbers";
        seriesItems = ["5", "10", "15", "20", "25", "30", "35", "40", "45", "50"];
        render();
        alert("[DIALOG FILL SERIES]\nSeries: Columns\nType: Linear\nStep Value: 5\nStop Value: 50\n\nDalam 1 detik deret kelipatan 5 langsung terbentuk tanpa perlu scroll!");
      });

      document.getElementById('btn-reset-series').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        seriesItems = ["1", "2"];
        render();
      });
    }

    render();
  },

  // ==========================================
  // LAB 2.10: CURRENCY, CELL STYLES & TEXT ORIENTATION
  // ==========================================
  initLab210() {
    const container = document.getElementById('lab-2-10-container');
    if (!container) return;

    let formatMode = "currency-idr"; // "raw", "currency-idr", "currency-usd", "accounting"
    let currentStyle = "normal"; // "normal", "cell-style-good", "cell-style-bad", "cell-style-accent1", "cell-style-heading"
    let textOrientation = "angle-0"; // "angle-0", "angle-45", "angle-vert"

    const rawValues = [
      { item: "SPP Bulanan", price: 350000, status: "Lunas" },
      { item: "Uang Gedung", price: 2500000, status: "Belum" },
      { item: "Seragam & Buku", price: 750000, status: "Lunas" },
      { item: "Praktikum Lab Komputer", price: 150000, status: "Lunas" }
    ];

    function formatNumber(num) {
      if (formatMode === "raw") return num.toString();
      if (formatMode === "currency-idr") return "Rp " + num.toLocaleString('id-ID');
      if (formatMode === "currency-usd") return "$" + (num / 15000).toFixed(2);
      if (formatMode === "accounting") return "Rp\t" + num.toLocaleString('id-ID');
      return num.toString();
    }

    function render() {
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon">X</span>
              <span>Latihan 2.10: Currency Format, Cell Styles & Text Orientation</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.85;">Shortcut Format: Ctrl + Shift + $</div>
          </div>

          <div class="excel-toolbar">
            <div style="display:flex; flex-direction:column; gap:8px; width:100%;">
              <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                <span style="font-weight:700; font-size:0.8rem; color:var(--text-muted);">Format Angka:</span>
                <button class="excel-btn-sm ${formatMode === 'currency-idr' ? 'active' : ''}" id="btn-fmt-idr">
                  Rp Rupiah (IDR)
                </button>
                <button class="excel-btn-sm ${formatMode === 'currency-usd' ? 'active' : ''}" id="btn-fmt-usd">
                  $ US Dollar (USD)
                </button>
                <button class="excel-btn-sm ${formatMode === 'raw' ? 'active' : ''}" id="btn-fmt-raw">
                  Angka Polos (Tanpa Format)
                </button>

                <span style="font-weight:700; font-size:0.8rem; color:var(--text-muted); margin-left:12px;">Cell Styles:</span>
                <select id="select-cell-style" class="sim-select" style="padding:4px 8px; font-size:0.8rem;">
                  <option value="normal" ${currentStyle === 'normal' ? 'selected' : ''}>Normal</option>
                  <option value="cell-style-good" ${currentStyle === 'cell-style-good' ? 'selected' : ''}>Good (Hijau)</option>
                  <option value="cell-style-bad" ${currentStyle === 'cell-style-bad' ? 'selected' : ''}>Bad (Merah)</option>
                  <option value="cell-style-accent1" ${currentStyle === 'cell-style-accent1' ? 'selected' : ''}>Accent 1 (Biru)</option>
                  <option value="cell-style-heading" ${currentStyle === 'cell-style-heading' ? 'selected' : ''}>Heading 1</option>
                </select>

                <span style="font-weight:700; font-size:0.8rem; color:var(--text-muted); margin-left:12px;">Orientasi Teks:</span>
                <button class="excel-btn-sm ${textOrientation === 'angle-0' ? 'active' : ''}" id="btn-angle-0">
                  Horizontal (0°)
                </button>
                <button class="excel-btn-sm ${textOrientation === 'angle-45' ? 'active' : ''}" id="btn-angle-45">
                  Miring (45°)
                </button>
                <button class="excel-btn-sm ${textOrientation === 'angle-vert' ? 'active' : ''}" id="btn-angle-vert">
                  Vertikal (90°)
                </button>
                <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulators.copySimulatorByLabId('2.10')" title="Salin data tabel simulator ini ke Excel">
                  📑 Salin ke Sheet Baru
                </button>
              </div>
            </div>
          </div>

          <div class="excel-formula-bar">
            <div class="excel-name-box">B2</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" value="350000 (Nilai murni di balik sel tetap angka murni!)" readonly style="color:var(--excel-green); font-weight:600;">
          </div>

          <div class="excel-grid-container" style="min-height:220px;">
            <table class="excel-table">
              <thead>
                <tr style="${textOrientation !== 'angle-0' ? 'height:90px;' : ''}">
                  <th class="excel-corner-header"></th>
                  <th class="excel-col-header" style="width:160px;">
                    <span class="${textOrientation === 'angle-45' ? 'text-angle-45' : textOrientation === 'angle-vert' ? 'text-angle-vertical' : ''}">
                      A (Komponen Biaya)
                    </span>
                  </th>
                  <th class="excel-col-header" style="width:160px;">
                    <span class="${textOrientation === 'angle-45' ? 'text-angle-45' : textOrientation === 'angle-vert' ? 'text-angle-vertical' : ''}">
                      B (Nominal Tarif Pembayaran)
                    </span>
                  </th>
                  <th class="excel-col-header" style="width:130px;">
                    <span class="${textOrientation === 'angle-45' ? 'text-angle-45' : textOrientation === 'angle-vert' ? 'text-angle-vertical' : ''}">
                      C (Status Verifikasi)
                    </span>
                  </th>
                </tr>
              </thead>
              <tbody>
                ${rawValues.map((row, idx) => `
                  <tr>
                    <td class="excel-row-header">${idx + 1}</td>
                    <td class="excel-cell" style="font-weight:600;">${row.item}</td>
                    <td class="excel-cell ${currentStyle}" style="text-align:right; font-family:var(--font-mono);">
                      ${formatNumber(row.price)}
                    </td>
                    <td class="excel-cell ${row.status === 'Lunas' ? 'cell-style-good' : 'cell-style-bad'}">
                      ${row.status}
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">Pembayaran_Siswa</div>
            </div>
            <div style="font-size:0.8rem; color:var(--text-muted);">
              Aturan Emas: Format Currency mengubah tampilan visual sel, <strong>bukan nilai angka aslinya</strong>. Rumus =SUM tetap berjalan sempurna!
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('btn-fmt-idr').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        formatMode = "currency-idr";
        render();
      });

      document.getElementById('btn-fmt-usd').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        formatMode = "currency-usd";
        render();
      });

      document.getElementById('btn-fmt-raw').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        formatMode = "raw";
        render();
      });

      document.getElementById('select-cell-style').addEventListener('change', (e) => {
        SpreadsheetSimulators.playClickSound();
        currentStyle = e.target.value;
        render();
      });

      document.getElementById('btn-angle-0').addEventListener('click', () => {
        SpreadsheetSimulators.playClickSound();
        textOrientation = "angle-0";
        render();
      });

      document.getElementById('btn-angle-45').addEventListener('click', () => {
        SpreadsheetSimulators.playSuccessSound();
        textOrientation = "angle-45";
        render();
      });

      document.getElementById('btn-angle-vert').addEventListener('click', () => {
        SpreadsheetSimulators.playSuccessSound();
        textOrientation = "angle-vert";
        render();
      });
    }

    render();
  },

  // Initialize all 10 labs at once
  initAll() {
    this.initLab21();
    this.initLab22();
    this.initLab23();
    this.initLab24();
    this.initLab25();
    this.initLab26();
    this.initLab27();
    this.initLab28();
    this.initLab29();
    this.initLab210();
  }
};

window.SpreadsheetSimulators = SpreadsheetSimulators;
