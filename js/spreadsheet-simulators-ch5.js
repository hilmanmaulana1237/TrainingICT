/**
 * EXCEL SPREADSHEET INTERACTIVE SIMULATORS - CHAPTER 5
 * Charts, Sparklines, and Visual Objects
 * 6 Interactive Labs with Live SVG Render, Excel Controls, and Real-Time Data Editing
 */

const SpreadsheetSimulatorsCh5 = {
  playClickSound() {
    if (window.AppCh5 && window.AppCh5.soundEnabled) {
      window.AppCh5.playTone(600, 0.05);
    }
  },

  playSuccessSound() {
    if (window.AppCh5 && window.AppCh5.soundEnabled) {
      window.AppCh5.playTone(800, 0.08);
      setTimeout(() => window.AppCh5.playTone(1050, 0.12), 80);
    }
  },

  playWarningSound() {
    if (window.AppCh5 && window.AppCh5.soundEnabled) {
      window.AppCh5.playTone(350, 0.1);
      setTimeout(() => window.AppCh5.playTone(280, 0.15), 100);
    }
  },

  simulatorDatasets: {
    '5.1': {
      title: 'Latihan 5.1: Basic Chart Creation & Move Chart',
      sheetName: '5.1_GrafikDasar',
      data: "No\tNama Cabang\tKategori Produk\tTotal Omzet (Rp)\n" +
        "1\tJakarta Pusat\tElektronik\t85000000\n" +
        "2\tBandung Kota\tFashion\t42000000\n" +
        "3\tSurabaya Timur\tElektronik\t78000000\n" +
        "4\tYogyakarta\tKuliner\t31000000\n" +
        "5\tMedan Petisah\tElektronik\t64000000\n" +
        "6\tSemarang Barat\tFashion\t39000000"
    },
    '5.2': {
      title: 'Latihan 5.2: Chart Elements & Visual Styles',
      sheetName: '5.2_ElemenGrafik',
      data: "Bulan\tTarget Penjualan (Juta)\tRealisasi Penjualan (Juta)\n" +
        "Januari\t50\t55\n" +
        "Februari\t60\t58\n" +
        "Maret\t65\t72\n" +
        "April\t70\t68\n" +
        "Mei\t75\t82\n" +
        "Juni\t80\t85"
    },
    '5.3': {
      title: 'Latihan 5.3: Switch Row/Column & Dynamic Data',
      sheetName: '5.3_SwitchData',
      data: "Cabang\tKuartal 1\tKuartal 2\tKuartal 3\tKuartal 4\n" +
        "Jakarta\t120\t145\t160\t190\n" +
        "Bandung\t85\t92\t110\t125\n" +
        "Surabaya\t95\t115\t130\t150\n" +
        "Medan\t70\t80\t85\t95"
    },
    '5.4': {
      title: 'Latihan 5.4: Row-Level Sparklines',
      sheetName: '5.4_Sparklines',
      data: "Produk\tJan\tFeb\tMar\tApr\tMei\tJun\n" +
        "Laptop Pro\t15\t18\t22\t19\t25\t30\n" +
        "Monitor 4K\t10\t12\t-4\t8\t14\t11\n" +
        "Keyboard RGB\t25\t20\t18\t15\t12\t10\n" +
        "Mouse Gaming\t30\t-5\t35\t40\t-2\t45\n" +
        "Headset USB\t8\t10\t12\t14\t16\t18"
    },
    '5.5': {
      title: 'Latihan 5.5: Changing Chart Type & Secondary Axis',
      sheetName: '5.5_TipeGrafik',
      data: "Bulan\tUnit Terjual\tTingkat Laba (%)\n" +
        "Januari\t450\t18%\n" +
        "Februari\t520\t22%\n" +
        "Maret\t480\t19%\n" +
        "April\t610\t25%\n" +
        "Mei\t700\t28%\n" +
        "Juni\t750\t30%"
    },
    '5.6': {
      title: 'Latihan 5.6: Pictures & Text Boxes',
      sheetName: '5.6_ObjekVisual',
      data: "Metrik Utama\tNilai Capaian\tStatus Evaluasi\n" +
        "Total Omzet\tRp 520.000.000\tMelampaui Target (+15%)\n" +
        "Jumlah Pelanggan Baru\t1.420 Akun\tSesuai Rencana\n" +
        "Kepuasan Pelanggan (CSAT)\t94.8%\tSangat Baik (Bintang 5)\n" +
        "Tingkat Retensi Member\t88.2%\tPerlu Ditingkatkan di Q3"
    }
  },

  copySimulatorByLabId(labId) {
    const item = this.simulatorDatasets[labId];
    if (!item) return;
    if (window.AppCh5 && window.AppCh5.copyToClipboard) {
      window.AppCh5.copyToClipboard(item.data, item.title, item.sheetName);
    } else {
      navigator.clipboard.writeText(item.data).then(() => {
        alert(`Data tabel ${item.title} berhasil disalin! Buka Excel dan tekan Ctrl+V.`);
      });
    }
  },

  // =========================================================================
  // LAB 5.1: BASIC CHART CREATION & MOVE CHART
  // =========================================================================
  initLab51() {
    const container = document.getElementById('lab-5-1-container');
    if (!container) return;

    let selectionMode = 'ctrl'; // 'all' or 'ctrl' (non-adjacent: Cabang & Omzet)
    let chartCreated = true;
    let chartLocation = 'embedded'; // 'embedded' (on sheet) or 'chart_sheet' (full new sheet)

    const branches = [
      { no: 1, cabang: 'Jakarta Pusat', kategori: 'Elektronik', omzet: 85 },
      { no: 2, cabang: 'Bandung Kota', kategori: 'Fashion', omzet: 42 },
      { no: 3, cabang: 'Surabaya Timur', kategori: 'Elektronik', omzet: 78 },
      { no: 4, cabang: 'Yogyakarta', kategori: 'Kuliner', omzet: 31 },
      { no: 5, cabang: 'Medan Petisah', kategori: 'Elektronik', omzet: 64 },
      { no: 6, cabang: 'Semarang Barat', kategori: 'Fashion', omzet: 39 }
    ];

    function render() {
      const maxVal = 100;

      // Generate SVG Chart
      const svgWidth = 520;
      const svgHeight = 220;
      const chartPadding = { top: 30, right: 20, bottom: 45, left: 50 };
      const plotWidth = svgWidth - chartPadding.left - chartPadding.right;
      const plotHeight = svgHeight - chartPadding.top - chartPadding.bottom;
      const colWidth = (plotWidth / branches.length) * 0.55;
      const colGap = plotWidth / branches.length;

      let barsSvg = '';
      branches.forEach((b, i) => {
        const x = chartPadding.left + i * colGap + (colGap - colWidth) / 2;
        const barH = (b.omzet / maxVal) * plotHeight;
        const y = chartPadding.top + plotHeight - barH;
        barsSvg += `
          <g class="chart-bar-group" style="cursor:pointer;">
            <title>${b.cabang}: ${b.omzet} Juta Rp</title>
            <rect x="${x}" y="${y}" width="${colWidth}" height="${barH}" rx="3" fill="#4f46e5" class="chart-bar">
              <animate attributeName="height" from="0" to="${barH}" dur="0.4s" />
              <animate attributeName="y" from="${chartPadding.top + plotHeight}" to="${y}" dur="0.4s" />
            </rect>
            <text x="${x + colWidth / 2}" y="${y - 6}" font-size="11" font-weight="700" fill="#4f46e5" text-anchor="middle">${b.omzet}M</text>
            <text x="${x + colWidth / 2}" y="${chartPadding.top + plotHeight + 16}" font-size="10" fill="var(--text-muted)" text-anchor="middle">${b.cabang.split(' ')[0]}</text>
          </g>
        `;
      });

      // Gridlines SVG
      let gridSvg = '';
      [0, 25, 50, 75, 100].forEach(val => {
        const yPos = chartPadding.top + plotHeight - (val / maxVal) * plotHeight;
        gridSvg += `
          <line x1="${chartPadding.left}" y1="${yPos}" x2="${svgWidth - chartPadding.right}" y2="${yPos}" stroke="#e2e8f0" stroke-dasharray="3,3" />
          <text x="${chartPadding.left - 8}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${val}M</text>
        `;
      });

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar" style="background:#4f46e5;">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon" style="color:#4f46e5;">X</span>
              <span>Latihan 5.1: Basic Chart Creation & Move Chart (Alt + F1 / F11)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Pintasan Cepat: Alt + F1 (Embedded) | F11 (Chart Sheet)</div>
          </div>

          <div class="excel-toolbar">
            <span class="sim-label">Metode Seleksi:</span>
            <button class="excel-btn-sm ${selectionMode === 'ctrl' ? 'active' : ''}" id="lab51-btn-ctrl" title="Tahan Ctrl: Kolom B dan D terpilih, Kolom C dilewati">
              🎯 Tahan Ctrl (Kolom Terpisah: B & D)
            </button>
            <button class="excel-btn-sm ${selectionMode === 'all' ? 'active' : ''}" id="lab51-btn-all" title="Blok semua kolom B sampai D">
              Semua Kolom (B, C, D)
            </button>

            <span class="sim-divider"></span>

            <button class="btn btn-primary" id="lab51-btn-move" style="padding:6px 14px; font-size:0.82rem; background:#4f46e5; border-color:#4f46e5;">
              ${chartLocation === 'embedded' ? '📑 Move Chart ke Chart Sheet Baru (F11)' : '↩️ Move Chart Kembali ke Sheet1 (Object In)'}
            </button>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh5.copySimulatorByLabId('5.1')">
              📋 Salin ke Excel
            </button>
          </div>

          <!-- Formula Bar / Status -->
          <div class="excel-formula-bar">
            <div class="excel-name-box">B1:D7</div>
            <div class="excel-fx-icon">fx</div>
            <input type="text" class="excel-formula-input" readonly value="${selectionMode === 'ctrl' ? '=SERIES(Sheet1!$D$1; Sheet1!$B$2:$B$7; Sheet1!$D$2:$D$7; 1)  [Kolom C Dilewati via Ctrl]' : '=SERIES(Sheet1!$D$1; Sheet1!$B$2:$B$7; Sheet1!$B$2:$D$7; 1)'}">
          </div>

          <!-- Main Viewport -->
          <div class="excel-grid-container" style="min-height:300px; padding:12px;">
            ${chartLocation === 'chart_sheet' ? `
              <!-- Fullscreen Chart Sheet View -->
              <div style="background:var(--bg-card); border:2px dashed #4f46e5; border-radius:12px; padding:24px; text-align:center;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                  <span style="font-weight:700; color:#4f46e5; font-size:1.05rem;">📊 Lembar Khusus: Chart1 (Chart Sheet Mandiri)</span>
                  <span style="font-size:0.8rem; background:rgba(79, 70, 229, 0.1); color:#4f46e5; padding:4px 10px; border-radius:6px; font-weight:600;">Full Window Display • Bebas Grid Sel</span>
                </div>
                <div style="font-family:var(--font-heading); font-size:1.25rem; font-weight:800; color:var(--text-main); margin-bottom:4px;">
                  Total Omzet Penjualan per Wilayah Cabang
                </div>
                <div style="font-size:0.85rem; color:var(--text-muted); margin-bottom:16px;">Tampilan presentasi layar penuh (hasil pemindahan melalui tombol Move Chart)</div>
                <div style="max-width:700px; margin:0 auto; background:var(--bg-surface); padding:16px; border-radius:10px; border:1px solid var(--border-color);">
                  <svg viewBox="0 0 ${svgWidth} ${svgHeight}" style="width:100%; height:auto;">
                    ${gridSvg}
                    ${barsSvg}
                  </svg>
                </div>
              </div>
            ` : `
              <!-- Embedded Object View (Table + Chart Side by Side) -->
              <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap:18px; align-items:start;">
                <!-- Table with selection highlight -->
                <div style="overflow-x:auto;">
                  <div style="font-size:0.75rem; font-weight:700; color:var(--text-muted); margin-bottom:6px;">
                    DATA SHEET1: ${selectionMode === 'ctrl' ? '✨ Kolom B & D Terpilih (Ctrl Aktif)' : 'Semua Kolom Terpilih'}
                  </div>
                  <table class="excel-table">
                    <thead>
                      <tr>
                        <th class="excel-corner-header"></th>
                        <th class="excel-col-header" style="${selectionMode === 'ctrl' ? 'background:#e0e7ff; color:#3730a3; font-weight:800;' : ''}">B (Cabang)</th>
                        <th class="excel-col-header" style="${selectionMode === 'ctrl' ? 'opacity:0.4;' : ''}">C (Kategori)</th>
                        <th class="excel-col-header" style="${selectionMode === 'ctrl' ? 'background:#e0e7ff; color:#3730a3; font-weight:800;' : ''}">D (Omzet Rp)</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${branches.map((b, i) => `
                        <tr>
                          <td class="excel-row-header">${i + 1}</td>
                          <td class="excel-cell" style="font-weight:600; ${selectionMode === 'ctrl' ? 'background:#f5f3ff;' : ''}">${b.cabang}</td>
                          <td class="excel-cell" style="${selectionMode === 'ctrl' ? 'opacity:0.4; text-decoration:line-through;' : ''}">${b.kategori}</td>
                          <td class="excel-cell" style="text-align:right; font-weight:700; color:#107c41; ${selectionMode === 'ctrl' ? 'background:#f5f3ff;' : ''}">${b.omzet}.000.000</td>
                        </tr>
                      `).join('')}
                    </tbody>
                  </table>
                </div>

                <!-- Embedded Chart Card -->
                <div style="background:var(--bg-surface); border:1px solid var(--border-color); border-radius:10px; padding:14px; box-shadow:0 4px 14px rgba(0,0,0,0.06); position:relative;">
                  <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                    <div style="font-family:var(--font-heading); font-size:0.95rem; font-weight:700; color:var(--text-main);">
                      Total Omzet Penjualan (2-D Clustered Column)
                    </div>
                    <span style="font-size:0.72rem; background:#e0e7ff; color:#4338ca; padding:2px 8px; border-radius:4px; font-weight:700;">
                      Embedded Object
                    </span>
                  </div>
                  <svg viewBox="0 0 ${svgWidth} ${svgHeight}" style="width:100%; height:auto;">
                    ${gridSvg}
                    ${barsSvg}
                  </svg>
                  <div style="margin-top:8px; font-size:0.78rem; color:var(--text-muted); display:flex; justify-content:space-between;">
                    <span>💡 Pintasan: <strong>Alt + F1</strong></span>
                    <span>Pindahkan via: <strong>Chart Design ➔ Move Chart</strong></span>
                  </div>
                </div>
              </div>
            `}
          </div>

          <!-- Statusbar with Sheet Tabs -->
          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab ${chartLocation === 'embedded' ? 'active' : ''}" id="lab51-tab-sheet1">
                5.1_GrafikDasar
              </div>
              ${chartLocation === 'chart_sheet' ? `
                <div class="excel-sheet-tab active" style="background:#4f46e5; color:white; font-weight:700;">
                  📊 Chart1 (Chart Sheet)
                </div>
              ` : ''}
            </div>
            <div style="font-size:0.8rem; color:var(--text-muted);">
              ${chartLocation === 'chart_sheet' ? 'Status: Menampilkan grafik mandiri pada Chart Sheet.' : 'Status: Grafik disematkan (Embedded) di lembar kerja aktif.'}
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('lab51-btn-ctrl').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playClickSound();
        selectionMode = 'ctrl';
        render();
      });

      document.getElementById('lab51-btn-all').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playWarningSound();
        selectionMode = 'all';
        render();
      });

      document.getElementById('lab51-btn-move').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playSuccessSound();
        chartLocation = chartLocation === 'embedded' ? 'chart_sheet' : 'embedded';
        render();
      });

      const tabSheet1 = document.getElementById('lab51-tab-sheet1');
      if (tabSheet1) {
        tabSheet1.addEventListener('click', () => {
          if (chartLocation === 'chart_sheet') {
            chartLocation = 'embedded';
            render();
          }
        });
      }
    }

    render();
  },

  // =========================================================================
  // LAB 5.2: CHART ELEMENTS (+) & VISUAL STYLES
  // =========================================================================
  initLab52() {
    const container = document.getElementById('lab-5-2-container');
    if (!container) return;

    let elements = {
      title: true,
      titleText: 'Evaluasi Target vs Realisasi Semester 1',
      axisTitles: true,
      dataLabels: true,
      gridlines: true,
      legend: true
    };

    let currentStyle = 'modern'; // 'clean', 'modern' (dark), 'gradient'
    let currentPalette = 'colorful'; // 'colorful', 'monochromatic', 'emerald'
    let showElementsMenu = false;

    const dataMonths = [
      { bulan: 'Jan', target: 50, real: 55 },
      { bulan: 'Feb', target: 60, real: 58 },
      { bulan: 'Mar', target: 65, real: 72 },
      { bulan: 'Apr', target: 70, real: 68 },
      { bulan: 'Mei', target: 75, real: 82 },
      { bulan: 'Jun', target: 80, real: 85 }
    ];

    function render() {
      const maxVal = 100;
      const svgWidth = 560;
      const svgHeight = 250;
      const chartPadding = {
        top: elements.title ? 45 : 20,
        right: 25,
        bottom: elements.axisTitles ? 55 : 35,
        left: elements.axisTitles ? 60 : 45
      };

      const plotWidth = svgWidth - chartPadding.left - chartPadding.right;
      const plotHeight = svgHeight - chartPadding.top - chartPadding.bottom;
      const groupWidth = plotWidth / dataMonths.length;
      const barWidth = groupWidth * 0.35;

      // Color selection
      let colorTarget = '#94a3b8';
      let colorReal = '#4f46e5';
      let bgColor = 'var(--bg-surface)';
      let textColor = 'var(--text-main)';

      if (currentStyle === 'modern') {
        bgColor = '#0f172a';
        textColor = '#f8fafc';
      }

      if (currentPalette === 'monochromatic') {
        colorTarget = currentStyle === 'modern' ? '#38bdf8' : '#0284c7';
        colorReal = currentStyle === 'modern' ? '#818cf8' : '#312e81';
      } else if (currentPalette === 'emerald') {
        colorTarget = '#64748b';
        colorReal = '#10b981';
      } else {
        colorTarget = '#f59e0b';
        colorReal = '#4f46e5';
      }

      // Build Bars
      let barsSvg = '';
      dataMonths.forEach((d, i) => {
        const groupX = chartPadding.left + i * groupWidth;
        const xTarget = groupX + (groupWidth - barWidth * 2) / 2;
        const xReal = xTarget + barWidth + 3;

        const hTarget = (d.target / maxVal) * plotHeight;
        const yTarget = chartPadding.top + plotHeight - hTarget;

        const hReal = (d.real / maxVal) * plotHeight;
        const yReal = chartPadding.top + plotHeight - hReal;

        barsSvg += `
          <!-- Target Bar -->
          <rect x="${xTarget}" y="${yTarget}" width="${barWidth}" height="${hTarget}" fill="${colorTarget}" rx="2" opacity="0.85">
            <title>${d.bulan} Target: ${d.target} Juta</title>
          </rect>
          ${elements.dataLabels ? `<text x="${xTarget + barWidth / 2}" y="${yTarget - 4}" font-size="9" font-weight="700" fill="${textColor}" text-anchor="middle">${d.target}</text>` : ''}

          <!-- Realisasi Bar -->
          <rect x="${xReal}" y="${yReal}" width="${barWidth}" height="${hReal}" fill="${colorReal}" rx="2">
            <title>${d.bulan} Realisasi: ${d.real} Juta</title>
          </rect>
          ${elements.dataLabels ? `<text x="${xReal + barWidth / 2}" y="${yReal - 4}" font-size="9" font-weight="700" fill="${colorReal}" text-anchor="middle">${d.real}</text>` : ''}

          <!-- X Label -->
          <text x="${groupX + groupWidth / 2}" y="${chartPadding.top + plotHeight + 16}" font-size="11" font-weight="600" fill="${textColor}" text-anchor="middle">${d.bulan}</text>
        `;
      });

      // Gridlines
      let gridSvg = '';
      if (elements.gridlines) {
        [0, 25, 50, 75, 100].forEach(val => {
          const yPos = chartPadding.top + plotHeight - (val / maxVal) * plotHeight;
          gridSvg += `
            <line x1="${chartPadding.left}" y1="${yPos}" x2="${svgWidth - chartPadding.right}" y2="${yPos}" stroke="${currentStyle === 'modern' ? '#334155' : '#e2e8f0'}" stroke-dasharray="2,2" />
            <text x="${chartPadding.left - 8}" y="${yPos + 4}" font-size="10" fill="${currentStyle === 'modern' ? '#94a3b8' : '#64748b'}" text-anchor="end">${val}</text>
          `;
        });
      }

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar" style="background:#4f46e5;">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon" style="color:#4f46e5;">X</span>
              <span>Latihan 5.2: Chart Elements (+) & Visual Styles</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Gunakan Tombol [+] Hijau untuk Menambah Elemen</div>
          </div>

          <div class="excel-toolbar">
            <span class="sim-label">Chart Styles:</span>
            <button class="excel-btn-sm ${currentStyle === 'clean' ? 'active' : ''}" id="lab52-style-clean">
              ☀️ Style Clean Light
            </button>
            <button class="excel-btn-sm ${currentStyle === 'modern' ? 'active' : ''}" id="lab52-style-modern">
              🌙 Style Modern Dark
            </button>

            <span class="sim-divider"></span>

            <span class="sim-label">Warna:</span>
            <select class="sim-select" id="lab52-color-select" style="font-size:0.8rem; padding:4px 8px;">
              <option value="colorful" ${currentPalette === 'colorful' ? 'selected' : ''}>🎨 Colorful (Orange & Indigo)</option>
              <option value="monochromatic" ${currentPalette === 'monochromatic' ? 'selected' : ''}>🔷 Monochromatic Blue</option>
              <option value="emerald" ${currentPalette === 'emerald' ? 'selected' : ''}>🌿 Emerald Growth</option>
            </select>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh5.copySimulatorByLabId('5.2')">
              📋 Salin ke Excel
            </button>
          </div>

          <!-- Chart Area with Floating [+] Button -->
          <div class="excel-grid-container" style="min-height:320px; padding:20px; background:${bgColor}; transition:all 0.3s ease; position:relative;">
            
            <!-- Excel Floating [+] Chart Elements Button -->
            <div style="position:absolute; top:20px; right:25px; z-index:20;">
              <button id="lab52-btn-plus" style="width:34px; height:34px; border-radius:6px; background:#107c41; color:white; border:none; font-size:1.3rem; font-weight:800; display:flex; align-items:center; justify-content:center; cursor:pointer; box-shadow:0 4px 10px rgba(0,0,0,0.2); transition:transform 0.15s;" title="Klik untuk membuka menu Chart Elements (+)">
                +
              </button>

              <!-- Floating Popout Menu -->
              ${showElementsMenu ? `
                <div style="position:absolute; top:42px; right:0; width:220px; background:var(--bg-surface); border:1px solid var(--border-color); border-radius:8px; padding:12px; box-shadow:0 10px 25px rgba(0,0,0,0.2); font-size:0.82rem; color:var(--text-main); z-index:30;">
                  <div style="font-weight:700; color:#107c41; margin-bottom:8px; display:flex; justify-content:space-between; align-items:center;">
                    <span>Chart Elements</span>
                    <span style="font-size:0.7rem; color:var(--text-muted);">Tab Design</span>
                  </div>
                  <label style="display:flex; align-items:center; gap:8px; margin-bottom:6px; cursor:pointer;">
                    <input type="checkbox" id="lab52-chk-title" ${elements.title ? 'checked' : ''}>
                    <span>Chart Title (Judul)</span>
                  </label>
                  <label style="display:flex; align-items:center; gap:8px; margin-bottom:6px; cursor:pointer;">
                    <input type="checkbox" id="lab52-chk-axis" ${elements.axisTitles ? 'checked' : ''}>
                    <span>Axis Titles (Judul Sumbu)</span>
                  </label>
                  <label style="display:flex; align-items:center; gap:8px; margin-bottom:6px; cursor:pointer;">
                    <input type="checkbox" id="lab52-chk-labels" ${elements.dataLabels ? 'checked' : ''}>
                    <span>Data Labels (Label Angka)</span>
                  </label>
                  <label style="display:flex; align-items:center; gap:8px; margin-bottom:6px; cursor:pointer;">
                    <input type="checkbox" id="lab52-chk-grid" ${elements.gridlines ? 'checked' : ''}>
                    <span>Gridlines (Garis Kisi)</span>
                  </label>
                  <label style="display:flex; align-items:center; gap:8px; cursor:pointer;">
                    <input type="checkbox" id="lab52-chk-legend" ${elements.legend ? 'checked' : ''}>
                    <span>Legend (Keterangan)</span>
                  </label>
                </div>
              ` : ''}
            </div>

            <!-- SVG Chart Render -->
            <div style="max-width:620px; margin:0 auto; position:relative;">
              ${elements.title ? `
                <div style="text-align:center; margin-bottom:10px;">
                  <span style="font-family:var(--font-heading); font-size:1.15rem; font-weight:800; color:${textColor};">
                    ${elements.titleText}
                  </span>
                </div>
              ` : ''}

              <!-- Legend -->
              ${elements.legend ? `
                <div style="display:flex; justify-content:center; gap:20px; font-size:0.8rem; margin-bottom:8px; color:${textColor};">
                  <div style="display:flex; align-items:center; gap:6px;">
                    <span style="display:inline-block; width:12px; height:12px; background:${colorTarget}; border-radius:2px;"></span>
                    <span>Target Penjualan</span>
                  </div>
                  <div style="display:flex; align-items:center; gap:6px;">
                    <span style="display:inline-block; width:12px; height:12px; background:${colorReal}; border-radius:2px;"></span>
                    <span>Realisasi Penjualan</span>
                  </div>
                </div>
              ` : ''}

              <svg viewBox="0 0 ${svgWidth} ${svgHeight}" style="width:100%; height:auto;">
                ${gridSvg}
                ${barsSvg}
                ${elements.axisTitles ? `
                  <!-- Horizontal Axis Title -->
                  <text x="${chartPadding.left + plotWidth / 2}" y="${svgHeight - 8}" font-size="11" font-weight="700" fill="${textColor}" text-anchor="middle">
                    Periode Semester 1 (Bulan)
                  </text>
                  <!-- Vertical Axis Title -->
                  <text x="14" y="${chartPadding.top + plotHeight / 2}" font-size="11" font-weight="700" fill="${textColor}" text-anchor="middle" transform="rotate(-90 14 ${chartPadding.top + plotHeight / 2})">
                    Nominal (Juta Rp)
                  </text>
                ` : ''}
              </svg>
            </div>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">5.2_ElemenGrafik</div>
            </div>
            <div style="font-size:0.8rem; color:var(--text-muted);">
              Klik tombol hijau <strong>[+]</strong> untuk mengaktifkan Data Labels atau mengubah Judul Sumbu.
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      const btnPlus = document.getElementById('lab52-btn-plus');
      if (btnPlus) {
        btnPlus.addEventListener('click', () => {
          SpreadsheetSimulatorsCh5.playClickSound();
          showElementsMenu = !showElementsMenu;
          render();
        });
      }

      const styleClean = document.getElementById('lab52-style-clean');
      if (styleClean) {
        styleClean.addEventListener('click', () => {
          SpreadsheetSimulatorsCh5.playClickSound();
          currentStyle = 'clean';
          render();
        });
      }

      const styleModern = document.getElementById('lab52-style-modern');
      if (styleModern) {
        styleModern.addEventListener('click', () => {
          SpreadsheetSimulatorsCh5.playClickSound();
          currentStyle = 'modern';
          render();
        });
      }

      const colorSelect = document.getElementById('lab52-color-select');
      if (colorSelect) {
        colorSelect.addEventListener('change', (e) => {
          SpreadsheetSimulatorsCh5.playClickSound();
          currentPalette = e.target.value;
          render();
        });
      }

      // Checkbox Listeners
      if (showElementsMenu) {
        document.getElementById('lab52-chk-title').addEventListener('change', (e) => {
          elements.title = e.target.checked;
          render();
        });
        document.getElementById('lab52-chk-axis').addEventListener('change', (e) => {
          elements.axisTitles = e.target.checked;
          render();
        });
        document.getElementById('lab52-chk-labels').addEventListener('change', (e) => {
          elements.dataLabels = e.target.checked;
          render();
        });
        document.getElementById('lab52-chk-grid').addEventListener('change', (e) => {
          elements.gridlines = e.target.checked;
          render();
        });
        document.getElementById('lab52-chk-legend').addEventListener('change', (e) => {
          elements.legend = e.target.checked;
          render();
        });
      }
    }

    render();
  },

  // =========================================================================
  // LAB 5.3: SWITCH ROW/COLUMN & DYNAMIC SOURCE DATA
  // =========================================================================
  initLab53() {
    const container = document.getElementById('lab-5-3-container');
    if (!container) return;

    let orientation = 'by_branch'; // 'by_branch' (X=Cabang, Legend=Quarters) or 'by_quarter' (X=Quarters, Legend=Cabang)

    let salesData = [
      { cabang: 'Jakarta', q1: 120, q2: 145, q3: 160, q4: 190, color: '#4f46e5' },
      { cabang: 'Bandung', q1: 85, q2: 92, q3: 110, q4: 125, color: '#06b6d4' },
      { cabang: 'Surabaya', q1: 95, q2: 115, q3: 130, q4: 150, color: '#10b981' },
      { cabang: 'Medan', q1: 70, q2: 80, q3: 85, q4: 95, color: '#f59e0b' }
    ];

    const quarterColors = ['#4f46e5', '#3b82f6', '#06b6d4', '#10b981'];

    function render() {
      const maxVal = 250;
      const svgWidth = 560;
      const svgHeight = 220;
      const chartPadding = { top: 30, right: 20, bottom: 40, left: 45 };
      const plotWidth = svgWidth - chartPadding.left - chartPadding.right;
      const plotHeight = svgHeight - chartPadding.top - chartPadding.bottom;

      let barsSvg = '';
      let legendHtml = '';

      if (orientation === 'by_branch') {
        // X-Axis = Branches, Legend = Q1, Q2, Q3, Q4
        const groupWidth = plotWidth / salesData.length;
        const barWidth = groupWidth * 0.18;

        legendHtml = `
          <div style="display:flex; justify-content:center; gap:16px; font-size:0.78rem; margin-bottom:8px;">
            ${['Kuartal 1', 'Kuartal 2', 'Kuartal 3', 'Kuartal 4'].map((q, idx) => `
              <div style="display:flex; align-items:center; gap:5px;">
                <span style="display:inline-block; width:10px; height:10px; background:${quarterColors[idx]}; border-radius:2px;"></span>
                <span>${q}</span>
              </div>
            `).join('')}
          </div>
        `;

        salesData.forEach((b, bIdx) => {
          const groupX = chartPadding.left + bIdx * groupWidth;
          const quarters = [b.q1, b.q2, b.q3, b.q4];

          quarters.forEach((qVal, qIdx) => {
            const x = groupX + (groupWidth - barWidth * 4) / 2 + qIdx * (barWidth + 2);
            const h = (qVal / maxVal) * plotHeight;
            const y = chartPadding.top + plotHeight - h;
            barsSvg += `
              <rect x="${x}" y="${y}" width="${barWidth}" height="${h}" fill="${quarterColors[qIdx]}" rx="2">
                <title>${b.cabang} Q${qIdx + 1}: ${qVal}</title>
              </rect>
            `;
          });

          barsSvg += `
            <text x="${groupX + groupWidth / 2}" y="${chartPadding.top + plotHeight + 16}" font-size="11" font-weight="700" fill="var(--text-main)" text-anchor="middle">
              ${b.cabang}
            </text>
          `;
        });
      } else {
        // X-Axis = Quarters (Q1-Q4), Legend = Cabang
        const quarters = ['Kuartal 1', 'Kuartal 2', 'Kuartal 3', 'Kuartal 4'];
        const groupWidth = plotWidth / quarters.length;
        const barWidth = groupWidth * 0.18;

        legendHtml = `
          <div style="display:flex; justify-content:center; gap:16px; font-size:0.78rem; margin-bottom:8px;">
            ${salesData.map(b => `
              <div style="display:flex; align-items:center; gap:5px;">
                <span style="display:inline-block; width:10px; height:10px; background:${b.color}; border-radius:2px;"></span>
                <span>${b.cabang}</span>
              </div>
            `).join('')}
          </div>
        `;

        quarters.forEach((qName, qIdx) => {
          const groupX = chartPadding.left + qIdx * groupWidth;

          salesData.forEach((b, bIdx) => {
            const val = b[`q${qIdx + 1}`];
            const x = groupX + (groupWidth - barWidth * salesData.length) / 2 + bIdx * (barWidth + 2);
            const h = (val / maxVal) * plotHeight;
            const y = chartPadding.top + plotHeight - h;
            barsSvg += `
              <rect x="${x}" y="${y}" width="${barWidth}" height="${h}" fill="${b.color}" rx="2">
                <title>${b.cabang} ${qName}: ${val}</title>
              </rect>
            `;
          });

          barsSvg += `
            <text x="${groupX + groupWidth / 2}" y="${chartPadding.top + plotHeight + 16}" font-size="11" font-weight="700" fill="var(--text-main)" text-anchor="middle">
              ${qName}
            </text>
          `;
        });
      }

      // Gridlines
      let gridSvg = '';
      [0, 50, 100, 150, 200, 250].forEach(val => {
        const yPos = chartPadding.top + plotHeight - (val / maxVal) * plotHeight;
        gridSvg += `
          <line x1="${chartPadding.left}" y1="${yPos}" x2="${svgWidth - chartPadding.right}" y2="${yPos}" stroke="#e2e8f0" stroke-dasharray="2,2" />
          <text x="${chartPadding.left - 6}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${val}</text>
        `;
      });

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar" style="background:#4f46e5;">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon" style="color:#4f46e5;">X</span>
              <span>Latihan 5.3: Switch Row/Column & Live Data Auto-Update</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Edit Angka di Sel Tabel ➔ Grafik Otomatis Berubah</div>
          </div>

          <div class="excel-toolbar">
            <button class="btn btn-primary" id="lab53-btn-switch" style="padding:6px 14px; font-size:0.82rem; background:#4f46e5; border-color:#4f46e5;">
              🔄 Switch Row/Column (${orientation === 'by_branch' ? 'Ubah ke Sumbu Kuartal' : 'Ubah ke Sumbu Cabang'})
            </button>

            <span style="font-size:0.8rem; color:var(--text-muted); margin-left:8px;">
              Sumbu X Saat Ini: <strong>${orientation === 'by_branch' ? 'Wilayah Cabang' : 'Periode Kuartal'}</strong>
            </span>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh5.copySimulatorByLabId('5.3')">
              📋 Salin ke Excel
            </button>
          </div>

          <!-- Split layout: Editable Table on left, Dynamic Chart on right -->
          <div class="excel-grid-container" style="min-height:300px; padding:16px;">
            <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap:20px; align-items:start;">
              
              <!-- Editable Spreadsheet Table -->
              <div>
                <div style="font-size:0.78rem; font-weight:700; color:var(--text-muted); margin-bottom:6px;">
                  ✏️ TABEL SUMBER DINAMIS (Coba ganti angka Jakarta Q1):
                </div>
                <table class="excel-table">
                  <thead>
                    <tr>
                      <th class="excel-corner-header"></th>
                      <th class="excel-col-header">Cabang</th>
                      <th class="excel-col-header">Q1</th>
                      <th class="excel-col-header">Q2</th>
                      <th class="excel-col-header">Q3</th>
                      <th class="excel-col-header">Q4</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${salesData.map((b, bIdx) => `
                      <tr>
                        <td class="excel-row-header">${bIdx + 1}</td>
                        <td class="excel-cell" style="font-weight:700;">${b.cabang}</td>
                        <td class="excel-cell" style="padding:2px 4px;">
                          <input type="number" class="lab53-cell-input" data-branch="${bIdx}" data-q="q1" value="${b.q1}" style="width:55px; border:1px solid #cbd5e1; padding:3px; text-align:right; font-weight:700; border-radius:3px;">
                        </td>
                        <td class="excel-cell" style="padding:2px 4px;">
                          <input type="number" class="lab53-cell-input" data-branch="${bIdx}" data-q="q2" value="${b.q2}" style="width:55px; border:1px solid #cbd5e1; padding:3px; text-align:right; border-radius:3px;">
                        </td>
                        <td class="excel-cell" style="padding:2px 4px;">
                          <input type="number" class="lab53-cell-input" data-branch="${bIdx}" data-q="q3" value="${b.q3}" style="width:55px; border:1px solid #cbd5e1; padding:3px; text-align:right; border-radius:3px;">
                        </td>
                        <td class="excel-cell" style="padding:2px 4px;">
                          <input type="number" class="lab53-cell-input" data-branch="${bIdx}" data-q="q4" value="${b.q4}" style="width:55px; border:1px solid #cbd5e1; padding:3px; text-align:right; border-radius:3px;">
                        </td>
                      </tr>
                    `).join('')}
                  </tbody>
                </table>
                <div style="font-size:0.75rem; color:#107c41; margin-top:8px;">
                  ⚡ <em>Data-linked:</em> Setiap kali nilai diubah, tinggi batang grafik langsung bereaksi secara seketika!
                </div>
              </div>

              <!-- Live Dynamic Chart -->
              <div style="background:var(--bg-surface); border:1px solid var(--border-color); border-radius:10px; padding:14px; box-shadow:0 4px 12px rgba(0,0,0,0.05);">
                <div style="text-align:center; font-family:var(--font-heading); font-size:1rem; font-weight:800; margin-bottom:6px;">
                  ${orientation === 'by_branch' ? 'Penjualan per Kuartal Berdasarkan Cabang' : 'Perbandingan Cabang per Tiap Kuartal'}
                </div>
                ${legendHtml}
                <svg viewBox="0 0 ${svgWidth} ${svgHeight}" style="width:100%; height:auto;">
                  ${gridSvg}
                  ${barsSvg}
                </svg>
              </div>

            </div>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">5.3_SwitchData</div>
            </div>
            <div style="font-size:0.8rem; color:var(--text-muted);">
              Klik <strong>Switch Row/Column</strong> untuk menukar Kategori (Sumbu X) dan Seri Data (Legenda).
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('lab53-btn-switch').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playSuccessSound();
        orientation = orientation === 'by_branch' ? 'by_quarter' : 'by_branch';
        render();
      });

      const inputs = container.querySelectorAll('.lab53-cell-input');
      inputs.forEach(input => {
        input.addEventListener('input', (e) => {
          const bIdx = parseInt(e.target.dataset.branch);
          const qKey = e.target.dataset.q;
          const val = Math.max(0, Math.min(250, parseInt(e.target.value) || 0));
          salesData[bIdx][qKey] = val;
          render();
        });
      });
    }

    render();
  },

  // =========================================================================
  // LAB 5.4: ROW-LEVEL SPARKLINES (LINE, COLUMN, WIN/LOSS)
  // =========================================================================
  initLab54() {
    const container = document.getElementById('lab-5-4-container');
    if (!container) return;

    let sparkType = 'line'; // 'line', 'column', 'winloss'
    let showHighPoint = true;
    let showLowPoint = true;
    let showNegPoints = true;
    let sparklinesVisible = true;

    const sparkProducts = [
      { nama: 'Laptop Pro 14"', vals: [15, 18, 22, 19, 25, 30] },
      { nama: 'Monitor 4K HDR', vals: [10, 12, -4, 8, 14, 11] },
      { nama: 'Keyboard Mech RGB', vals: [25, 20, 18, 15, 12, 10] },
      { nama: 'Mouse Gaming Wireless', vals: [30, -5, 35, 40, -2, 45] },
      { nama: 'Headset Gaming USB', vals: [8, 10, 12, 14, 16, 18] }
    ];

    function createSparklineSvg(values, type) {
      if (!sparklinesVisible) {
        return `<span style="font-size:0.75rem; color:#94a3b8; font-style:italic;">[Kosong - Clear]</span>`;
      }

      const w = 110;
      const h = 26;

      if (type === 'line') {
        const minVal = Math.min(...values);
        const maxVal = Math.max(...values);
        const range = maxVal - minVal || 1;

        const points = values.map((v, i) => {
          const x = 6 + (i / (values.length - 1)) * (w - 12);
          const y = h - 5 - ((v - minVal) / range) * (h - 10);
          return { x, y, v };
        });

        const polylinePoints = points.map(p => `${p.x},${p.y}`).join(' ');

        let markersSvg = '';
        points.forEach((p, idx) => {
          if (showHighPoint && p.v === maxVal) {
            markersSvg += `<circle cx="${p.x}" cy="${p.y}" r="3" fill="#10b981"><title>High: ${p.v}</title></circle>`;
          } else if (showLowPoint && p.v === minVal) {
            markersSvg += `<circle cx="${p.x}" cy="${p.y}" r="3" fill="#ef4444"><title>Low: ${p.v}</title></circle>`;
          }
        });

        return `
          <svg width="${w}" height="${h}" style="vertical-align:middle; overflow:visible;">
            <polyline fill="none" stroke="#4f46e5" stroke-width="2" points="${polylinePoints}" />
            ${markersSvg}
          </svg>
        `;
      } else if (type === 'column') {
        const maxAbs = Math.max(...values.map(v => Math.abs(v))) || 1;
        const colW = (w - 10) / values.length;

        let colsSvg = '';
        values.forEach((v, i) => {
          const x = 5 + i * colW + 1;
          const barH = Math.max(3, (Math.abs(v) / maxAbs) * (h - 6));
          const y = v >= 0 ? h / 2 - barH : h / 2;
          const isNeg = v < 0;
          const colFill = (showNegPoints && isNeg) ? '#ef4444' : '#4f46e5';

          colsSvg += `<rect x="${x}" y="${y}" width="${colW - 2}" height="${barH}" fill="${colFill}" rx="1"><title>${v}</title></rect>`;
        });

        return `
          <svg width="${w}" height="${h}" style="vertical-align:middle;">
            <line x1="2" y1="${h / 2}" x2="${w - 2}" y2="${h / 2}" stroke="#cbd5e1" stroke-width="1" />
            ${colsSvg}
          </svg>
        `;
      } else {
        // Win / Loss Sparkline (Binary + vs -)
        const colW = (w - 10) / values.length;
        let winLossSvg = '';

        values.forEach((v, i) => {
          const x = 5 + i * colW + 1;
          const isPos = v >= 0;
          const barH = 9;
          const y = isPos ? h / 2 - barH - 1 : h / 2 + 1;
          const colFill = isPos ? '#10b981' : '#ef4444';

          winLossSvg += `
            <rect x="${x}" y="${y}" width="${colW - 2}" height="${barH}" fill="${colFill}" rx="1">
              <title>${isPos ? 'Win (+)' : 'Loss (-)'}: ${v}</title>
            </rect>
          `;
        });

        return `
          <svg width="${w}" height="${h}" style="vertical-align:middle;">
            <line x1="2" y1="${h / 2}" x2="${w - 2}" y2="${h / 2}" stroke="#94a3b8" stroke-width="1" stroke-dasharray="2,2" />
            ${winLossSvg}
          </svg>
        `;
      }
    }

    function render() {
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar" style="background:#4f46e5;">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon" style="color:#4f46e5;">X</span>
              <span>Latihan 5.4: Row-Level Sparklines (Grafik Mikro di Dalam Sel)</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Pilihan: Line | Column | Win/Loss</div>
          </div>

          <div class="excel-toolbar">
            <span class="sim-label">Tipe Sparkline:</span>
            <button class="excel-btn-sm ${sparkType === 'line' ? 'active' : ''}" id="lab54-btn-line">
              📈 Line Sparkline
            </button>
            <button class="excel-btn-sm ${sparkType === 'column' ? 'active' : ''}" id="lab54-btn-column">
              📊 Column Sparkline
            </button>
            <button class="excel-btn-sm ${sparkType === 'winloss' ? 'active' : ''}" id="lab54-btn-winloss">
              ⚖️ Win/Loss Sparkline
            </button>

            <span class="sim-divider"></span>

            <button class="btn btn-outline" id="lab54-btn-clear" style="padding:4px 10px; font-size:0.78rem;">
              ${sparklinesVisible ? '🧹 Clear Sparklines (Tab Sparkline)' : '✨ Sisipkan Ulang Sparkline'}
            </button>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh5.copySimulatorByLabId('5.4')">
              📋 Salin ke Excel
            </button>
          </div>

          <!-- Secondary Toolbar: Sparkline Markers -->
          <div style="background:var(--bg-subtle); padding:8px 16px; border-bottom:1px solid var(--border-color); display:flex; gap:16px; align-items:center; font-size:0.8rem;">
            <span style="font-weight:700; color:#4f46e5;">🎯 Penanda Titik (Markers):</span>
            <label style="display:flex; align-items:center; gap:5px; cursor:pointer;">
              <input type="checkbox" id="lab54-chk-high" ${showHighPoint ? 'checked' : ''}>
              <span style="color:#10b981; font-weight:700;">● High Point (Puncak Hijau)</span>
            </label>
            <label style="display:flex; align-items:center; gap:5px; cursor:pointer;">
              <input type="checkbox" id="lab54-chk-low" ${showLowPoint ? 'checked' : ''}>
              <span style="color:#ef4444; font-weight:700;">● Low Point (Terendah)</span>
            </label>
            <label style="display:flex; align-items:center; gap:5px; cursor:pointer;">
              <input type="checkbox" id="lab54-chk-neg" ${showNegPoints ? 'checked' : ''}>
              <span style="color:#ef4444; font-weight:700;">■ Negative Points (Merah Minus)</span>
            </label>
          </div>

          <!-- Table Container -->
          <div class="excel-grid-container" style="min-height:260px; overflow-x:auto;">
            <table class="excel-table">
              <thead>
                <tr>
                  <th class="excel-corner-header"></th>
                  <th class="excel-col-header" style="width:170px;">A (Produk)</th>
                  <th class="excel-col-header">B (Jan)</th>
                  <th class="excel-col-header">C (Feb)</th>
                  <th class="excel-col-header">D (Mar)</th>
                  <th class="excel-col-header">E (Apr)</th>
                  <th class="excel-col-header">F (Mei)</th>
                  <th class="excel-col-header">G (Jun)</th>
                  <th class="excel-col-header" style="background:#e0e7ff; color:#3730a3; width:140px; font-weight:800;">
                    H (Tren Visual Sparkline)
                  </th>
                </tr>
              </thead>
              <tbody>
                ${sparkProducts.map((p, pIdx) => `
                  <tr>
                    <td class="excel-row-header">${pIdx + 1}</td>
                    <td class="excel-cell" style="font-weight:700;">${p.nama}</td>
                    ${p.vals.map(v => `
                      <td class="excel-cell" style="text-align:right; font-family:var(--font-mono); ${v < 0 ? 'color:#ef4444; font-weight:800; background:#fef2f2;' : ''}">
                        ${v}
                      </td>
                    `).join('')}
                    <td class="excel-cell" style="text-align:center; background:#f5f3ff; padding:2px 8px;">
                      ${createSparklineSvg(p.vals, sparkType)}
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">5.4_Sparklines</div>
            </div>
            <div style="font-size:0.8rem; color:var(--text-muted);">
              Tipe Aktif: <strong>${sparkType.toUpperCase()}</strong>. Sparkline hidup di dalam sel dan otomatis beradaptasi saat ukuran kolom diubah.
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('lab54-btn-line').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playClickSound();
        sparkType = 'line';
        sparklinesVisible = true;
        render();
      });

      document.getElementById('lab54-btn-column').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playClickSound();
        sparkType = 'column';
        sparklinesVisible = true;
        render();
      });

      document.getElementById('lab54-btn-winloss').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playClickSound();
        sparkType = 'winloss';
        sparklinesVisible = true;
        render();
      });

      document.getElementById('lab54-btn-clear').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playWarningSound();
        sparklinesVisible = !sparklinesVisible;
        render();
      });

      document.getElementById('lab54-chk-high').addEventListener('change', (e) => {
        showHighPoint = e.target.checked;
        render();
      });
      document.getElementById('lab54-chk-low').addEventListener('change', (e) => {
        showLowPoint = e.target.checked;
        render();
      });
      document.getElementById('lab54-chk-neg').addEventListener('change', (e) => {
        showNegPoints = e.target.checked;
        render();
      });
    }

    render();
  },

  // =========================================================================
  // LAB 5.5: CHANGING CHART TYPE, SIZE, AND COMBO CHART (SECONDARY AXIS)
  // =========================================================================
  initLab55() {
    const container = document.getElementById('lab-5-5-container');
    if (!container) return;

    let chartType = 'combo'; // 'column', 'line', 'bar', 'combo'
    let snapToGrid = false;
    let quickLayout = 1; // 1 (Default), 2 (Labels on top, clean), 5 (With Data Table)

    const comboData = [
      { bulan: 'Jan', unit: 450, margin: 18 },
      { bulan: 'Feb', unit: 520, margin: 22 },
      { bulan: 'Mar', unit: 480, margin: 19 },
      { bulan: 'Apr', unit: 610, margin: 25 },
      { bulan: 'Mei', unit: 700, margin: 28 },
      { bulan: 'Jun', unit: 750, margin: 30 }
    ];

    function render() {
      const maxUnit = 800;
      const maxMargin = 35; // %

      const svgWidth = 560;
      const svgHeight = 240;
      const chartPadding = { top: 35, right: chartType === 'combo' ? 55 : 25, bottom: 40, left: 55 };
      const plotWidth = svgWidth - chartPadding.left - chartPadding.right;
      const plotHeight = svgHeight - chartPadding.top - chartPadding.bottom;

      let chartSvg = '';

      if (chartType === 'column') {
        const groupW = plotWidth / comboData.length;
        const barW = groupW * 0.45;
        comboData.forEach((d, i) => {
          const x = chartPadding.left + i * groupW + (groupW - barW) / 2;
          const h = (d.unit / maxUnit) * plotHeight;
          const y = chartPadding.top + plotHeight - h;
          chartSvg += `
            <rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="#4f46e5" rx="2">
              <title>${d.bulan}: ${d.unit} Unit</title>
            </rect>
            ${quickLayout === 2 ? `<text x="${x + barW / 2}" y="${y - 4}" font-size="10" font-weight="700" fill="#4f46e5" text-anchor="middle">${d.unit}</text>` : ''}
            <text x="${x + barW / 2}" y="${chartPadding.top + plotHeight + 16}" font-size="11" fill="var(--text-muted)" text-anchor="middle">${d.bulan}</text>
          `;
        });
      } else if (chartType === 'line') {
        const step = plotWidth / (comboData.length - 1);
        const points = comboData.map((d, i) => {
          const x = chartPadding.left + i * step;
          const y = chartPadding.top + plotHeight - (d.unit / maxUnit) * plotHeight;
          return { x, y, ...d };
        });

        const lineStr = points.map(p => `${p.x},${p.y}`).join(' ');
        chartSvg += `<polyline fill="none" stroke="#4f46e5" stroke-width="3" points="${lineStr}" />`;

        points.forEach(p => {
          chartSvg += `
            <circle cx="${p.x}" cy="${p.y}" r="4" fill="#ffffff" stroke="#4f46e5" stroke-width="2" />
            <text x="${p.x}" y="${p.y - 8}" font-size="10" font-weight="700" fill="#4f46e5" text-anchor="middle">${p.unit}</text>
            <text x="${p.x}" y="${chartPadding.top + plotHeight + 16}" font-size="11" fill="var(--text-muted)" text-anchor="middle">${p.bulan}</text>
          `;
        });
      } else if (chartType === 'bar') {
        // Horizontal Bar Chart
        const rowH = plotHeight / comboData.length;
        const barH = rowH * 0.55;
        comboData.forEach((d, i) => {
          const y = chartPadding.top + i * rowH + (rowH - barH) / 2;
          const barW = (d.unit / maxUnit) * plotWidth;
          chartSvg += `
            <rect x="${chartPadding.left}" y="${y}" width="${barW}" height="${barH}" fill="#4f46e5" rx="2" />
            <text x="${chartPadding.left + barW + 6}" y="${y + barH / 2 + 4}" font-size="10" font-weight="700" fill="#4f46e5">${d.unit}</text>
            <text x="${chartPadding.left - 6}" y="${y + barH / 2 + 4}" font-size="11" fill="var(--text-muted)" text-anchor="end">${d.bulan}</text>
          `;
        });
      } else {
        // COMBO CHART: Clustered Column (Units on Left Y-Axis) + Line (Margin % on Right Y-Axis)
        const groupW = plotWidth / comboData.length;
        const barW = groupW * 0.45;

        // 1. Column Bars (Units)
        comboData.forEach((d, i) => {
          const x = chartPadding.left + i * groupW + (groupW - barW) / 2;
          const h = (d.unit / maxUnit) * plotHeight;
          const y = chartPadding.top + plotHeight - h;
          chartSvg += `
            <rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="#4f46e5" rx="2" opacity="0.85">
              <title>${d.bulan} Volume: ${d.unit} Unit</title>
            </rect>
            <text x="${x + barW / 2}" y="${chartPadding.top + plotHeight + 16}" font-size="11" fill="var(--text-muted)" text-anchor="middle">${d.bulan}</text>
          `;
        });

        // 2. Line with Markers (Margin %) on Secondary Axis
        const linePoints = comboData.map((d, i) => {
          const x = chartPadding.left + i * groupW + groupW / 2;
          const y = chartPadding.top + plotHeight - (d.margin / maxMargin) * plotHeight;
          return { x, y, ...d };
        });

        const lineStr = linePoints.map(p => `${p.x},${p.y}`).join(' ');
        chartSvg += `<polyline fill="none" stroke="#f59e0b" stroke-width="3" points="${lineStr}" />`;

        linePoints.forEach(p => {
          chartSvg += `
            <circle cx="${p.x}" cy="${p.y}" r="4" fill="#ffffff" stroke="#f59e0b" stroke-width="2.5" />
            <text x="${p.x}" y="${p.y - 8}" font-size="10" font-weight="800" fill="#d97706" text-anchor="middle">${p.margin}%</text>
          `;
        });
      }

      // Left Y-Axis Labels
      let yAxisLeftSvg = '';
      [0, 200, 400, 600, 800].forEach(val => {
        const yPos = chartPadding.top + plotHeight - (val / maxUnit) * plotHeight;
        yAxisLeftSvg += `
          <line x1="${chartPadding.left}" y1="${yPos}" x2="${svgWidth - chartPadding.right}" y2="${yPos}" stroke="#e2e8f0" stroke-dasharray="2,2" />
          <text x="${chartPadding.left - 6}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${val}</text>
        `;
      });

      // Right Y-Axis (Secondary Axis) for Combo Chart
      let yAxisRightSvg = '';
      if (chartType === 'combo') {
        [0, 10, 20, 30, 35].forEach(val => {
          const yPos = chartPadding.top + plotHeight - (val / maxMargin) * plotHeight;
          yAxisRightSvg += `
            <text x="${svgWidth - chartPadding.right + 6}" y="${yPos + 4}" font-size="10" font-weight="700" fill="#d97706" text-anchor="start">${val}%</text>
          `;
        });
      }

      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar" style="background:#4f46e5;">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon" style="color:#4f46e5;">X</span>
              <span>Latihan 5.5: Change Chart Type, Quick Layout, & Sumbu Sekunder</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Tahan Alt Saat Drag untuk Snap to Grid</div>
          </div>

          <div class="excel-toolbar">
            <span class="sim-label">Tipe Grafik:</span>
            <button class="excel-btn-sm ${chartType === 'combo' ? 'active' : ''}" id="lab55-btn-combo" style="font-weight:700;">
              🔀 Combo Chart (Secondary Axis)
            </button>
            <button class="excel-btn-sm ${chartType === 'column' ? 'active' : ''}" id="lab55-btn-col">
              📊 Clustered Column
            </button>
            <button class="excel-btn-sm ${chartType === 'line' ? 'active' : ''}" id="lab55-btn-line">
              📈 Line Chart
            </button>
            <button class="excel-btn-sm ${chartType === 'bar' ? 'active' : ''}" id="lab55-btn-bar">
              ↔️ Bar Chart
            </button>

            <span class="sim-divider"></span>

            <button class="btn btn-outline ${snapToGrid ? 'active' : ''}" id="lab55-btn-snap" style="padding:4px 10px; font-size:0.78rem;">
              ${snapToGrid ? '🧲 Snap to Grid (Alt Aktif)' : '🔓 Snap to Grid (Lepas)'}
            </button>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh5.copySimulatorByLabId('5.5')">
              📋 Salin ke Excel
            </button>
          </div>

          <!-- Chart Area -->
          <div class="excel-grid-container" style="min-height:310px; padding:16px;">
            <div style="max-width:620px; margin:0 auto; background:var(--bg-surface); border:${snapToGrid ? '2px dashed #4f46e5' : '1px solid var(--border-color)'}; border-radius:${snapToGrid ? '0' : '8px'}; padding:14px; position:relative; box-shadow:0 4px 12px rgba(0,0,0,0.05); transition:all 0.2s;">
              
              <!-- Legend & Titles -->
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                <span style="font-family:var(--font-heading); font-size:1rem; font-weight:800; color:var(--text-main);">
                  ${chartType === 'combo' ? 'Kombinasi Volume Penjualan vs Rasio Laba (%)' : 'Performa Volume Penjualan Semester 1'}
                </span>
                
                ${chartType === 'combo' ? `
                  <div style="display:flex; gap:12px; font-size:0.76rem;">
                    <span style="display:flex; align-items:center; gap:4px; color:#4f46e5; font-weight:700;">
                      <span style="width:10px; height:10px; background:#4f46e5; border-radius:2px;"></span> Unit Terjual (Kiri)
                    </span>
                    <span style="display:flex; align-items:center; gap:4px; color:#d97706; font-weight:700;">
                      <span style="width:10px; height:3px; background:#f59e0b;"></span> Rasio Laba % (Kanan)
                    </span>
                  </div>
                ` : ''}
              </div>

              <!-- SVG Render -->
              <svg viewBox="0 0 ${svgWidth} ${svgHeight}" style="width:100%; height:auto;">
                ${yAxisLeftSvg}
                ${yAxisRightSvg}
                ${chartSvg}
              </svg>

              <!-- Quick Layout Footer -->
              <div style="margin-top:10px; padding-top:8px; border-top:1px solid var(--border-color); display:flex; justify-content:space-between; align-items:center; font-size:0.75rem; color:var(--text-muted);">
                <span>Quick Layout: <strong>Layout ${quickLayout}</strong></span>
                <span>💡 Trik Sumbu Sekunder: Memungkinkan dua unit berbeda (Rupiah vs Persen) hidup harmonis dalam satu grafik.</span>
              </div>
            </div>
          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">5.5_TipeGrafik</div>
            </div>
            <div style="font-size:0.8rem; color:var(--text-muted);">
              Sumbu Sekunder (Secondary Axis) menjaga agar persentase tidak menempel di angka 0.
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      document.getElementById('lab55-btn-combo').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playSuccessSound();
        chartType = 'combo';
        render();
      });

      document.getElementById('lab55-btn-col').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playClickSound();
        chartType = 'column';
        render();
      });

      document.getElementById('lab55-btn-line').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playClickSound();
        chartType = 'line';
        render();
      });

      document.getElementById('lab55-btn-bar').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playClickSound();
        chartType = 'bar';
        render();
      });

      document.getElementById('lab55-btn-snap').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playClickSound();
        snapToGrid = !snapToGrid;
        render();
      });
    }

    render();
  },

  // =========================================================================
  // LAB 5.6: PICTURES (ALT TEXT) & FLOATING TEXT BOXES
  // =========================================================================
  initLab56() {
    const container = document.getElementById('lab-5-6-container');
    if (!container) return;

    let showAltTextDrawer = false;
    let altTextValue = 'Logo Resmi SMK Negeri 1 Jakarta - Laporan Kinerja Manajemen 2026';
    let textBoxBg = 'note'; // 'note' (yellow post-it), 'transparent', 'slate'
    let customNoteText = '📌 KESIMPULAN EKSEKUTIF: Target semester 1 tercapai 115%. Prioritas semester 2 adalah ekspansi ke 3 kota baru dan peningkatan retensi member.';

    function render() {
      container.innerHTML = `
        <div class="excel-window">
          <div class="excel-ribbon-bar" style="background:#4f46e5;">
            <div class="excel-ribbon-title">
              <span class="excel-ribbon-icon" style="color:#4f46e5;">X</span>
              <span>Latihan 5.6: Pictures (Alt Text) & Floating Text Box</span>
            </div>
            <div style="font-size:0.75rem; opacity:0.9;">Objek Visual Melayang Bebas di Atas Sel</div>
          </div>

          <div class="excel-toolbar">
            <button class="btn btn-outline" id="lab56-btn-alt" style="padding:5px 12px; font-size:0.8rem;">
              👁️ ${showAltTextDrawer ? 'Tutup Panel Alt Text' : 'Periksa Alt Text Gambar (Aksesibilitas)'}
            </button>

            <span class="sim-divider"></span>

            <span class="sim-label">Format Text Box:</span>
            <button class="excel-btn-sm ${textBoxBg === 'note' ? 'active' : ''}" id="lab56-bg-note">
              📝 Catatan Kuning
            </button>
            <button class="excel-btn-sm ${textBoxBg === 'transparent' ? 'active' : ''}" id="lab56-bg-trans">
              🪟 No Fill (Transparan)
            </button>

            <button class="excel-btn-sm" style="margin-left:auto;" onclick="SpreadsheetSimulatorsCh5.copySimulatorByLabId('5.6')">
              📋 Salin ke Excel
            </button>
          </div>

          <!-- Worksheet Workspace with Floating Objects -->
          <div class="excel-grid-container" style="min-height:360px; padding:20px; position:relative; overflow-x:auto;">
            
            <!-- Grid Background Simulation -->
            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:16px;">
              <div>
                <div style="font-family:var(--font-heading); font-size:1.15rem; font-weight:800; color:var(--text-main);">
                  DASBOR RINGKASAN EKSEKUTIF ICT EXCEL
                </div>
                <div style="font-size:0.8rem; color:var(--text-muted);">Kompilasi Metrik Utama & Dokumen Manajemen</div>
              </div>

              <!-- FLOATING PICTURE OBJECT WITH ALT TEXT BADGE -->
              <div id="lab56-picture-card" style="border:2px solid #4f46e5; border-radius:10px; background:var(--bg-surface); padding:8px 14px; display:flex; align-items:center; gap:10px; box-shadow:0 6px 16px rgba(79, 70, 229, 0.15); cursor:pointer;" title="Klik untuk menguji Alt Text">
                <div style="width:42px; height:42px; background:linear-gradient(135deg, #4f46e5, #06b6d4); border-radius:8px; display:flex; align-items:center; justify-content:center; color:white; font-size:1.4rem;">
                  🏛️
                </div>
                <div>
                  <div style="font-weight:700; font-size:0.85rem; color:#4f46e5;">LOGO INSTITUSI</div>
                  <div style="font-size:0.72rem; color:var(--text-muted);">Picture Format • Alt Text: OK ✅</div>
                </div>
              </div>
            </div>

            <!-- Table Beneath Objects -->
            <div style="max-width:550px; margin-bottom:20px;">
              <table class="excel-table">
                <thead>
                  <tr>
                    <th class="excel-col-header">Metrik Kinerja</th>
                    <th class="excel-col-header" style="text-align:right;">Realisasi</th>
                    <th class="excel-col-header">Status</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td class="excel-cell" style="font-weight:700;">Total Omzet</td>
                    <td class="excel-cell" style="text-align:right; color:#107c41; font-weight:800;">Rp 520.000.000</td>
                    <td class="excel-cell" style="color:#107c41; font-weight:600;">Melampaui Target (+15%)</td>
                  </tr>
                  <tr>
                    <td class="excel-cell" style="font-weight:700;">Pelanggan Baru</td>
                    <td class="excel-cell" style="text-align:right; color:#4f46e5; font-weight:800;">1.420 Akun</td>
                    <td class="excel-cell">Sesuai Rencana</td>
                  </tr>
                  <tr>
                    <td class="excel-cell" style="font-weight:700;">Kepuasan (CSAT)</td>
                    <td class="excel-cell" style="text-align:right; color:#d97706; font-weight:800;">94.8%</td>
                    <td class="excel-cell" style="color:#d97706; font-weight:600;">Sangat Baik (Bintang 5)</td>
                  </tr>
                </tbody>
              </table>
            </div>

            <!-- FLOATING TEXT BOX (Independent of cell grid) -->
            <div style="background:${textBoxBg === 'note' ? '#fef9c3' : textBoxBg === 'transparent' ? 'rgba(255,255,255,0.7)' : '#1e293b'}; color:${textBoxBg === 'slate' ? '#f8fafc' : '#1e293b'}; border:${textBoxBg === 'note' ? '1.5px solid #ca8a04' : textBoxBg === 'transparent' ? '1.5px dashed #4f46e5' : '1.5px solid #334155'}; border-radius:8px; padding:12px 16px; max-width:520px; box-shadow:0 8px 20px rgba(0,0,0,0.08); position:relative;">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; font-size:0.75rem; font-weight:700; color:${textBoxBg === 'note' ? '#854d0e' : '#4f46e5'};">
                <span>🔲 EXCEL TEXT BOX (Insert → Text Box)</span>
                <span style="font-size:0.68rem; opacity:0.8;">Bebas Digeser & Diubah Ukurannya</span>
              </div>
              <textarea id="lab56-textarea" style="width:100%; min-height:60px; background:transparent; border:none; resize:vertical; font-family:var(--font-sans); font-size:0.86rem; line-height:1.5; color:inherit; font-weight:600; outline:none;">${customNoteText}</textarea>
            </div>

            <!-- ALT TEXT DRAWER MODAL -->
            ${showAltTextDrawer ? `
              <div style="position:absolute; top:20px; right:20px; width:280px; background:var(--bg-surface); border:2px solid #4f46e5; border-radius:10px; padding:16px; box-shadow:0 12px 30px rgba(0,0,0,0.25); z-index:40; font-size:0.84rem;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                  <span style="font-weight:800; color:#4f46e5; font-size:0.95rem;">Panel Alt Text</span>
                  <button id="lab56-btn-close-alt" style="background:none; border:none; font-size:1.1rem; cursor:pointer;">✕</button>
                </div>
                <div style="font-size:0.78rem; color:var(--text-muted); margin-bottom:10px;">
                  Bagaimana kamu mendeskripsikan gambar ini untuk orang tunanetra atau gangguan penglihatan?
                </div>
                <textarea id="lab56-alt-input" style="width:100%; height:80px; padding:8px; border:1px solid var(--border-color); border-radius:6px; font-size:0.8rem; margin-bottom:10px; font-family:var(--font-sans);">${altTextValue}</textarea>
                <div style="font-size:0.72rem; color:#107c41; font-weight:700;">
                  ✅ Alt Text tersimpan untuk standardisasi dokumen resmi.
                </div>
              </div>
            ` : ''}

          </div>

          <div class="excel-statusbar">
            <div class="excel-sheet-tabs">
              <div class="excel-sheet-tab active">5.6_ObjekVisual</div>
            </div>
            <div style="font-size:0.8rem; color:var(--text-muted);">
              Text Box dan Gambar melayang di atas sel, menjaga struktur tabel tetap utuh tanpa merusak rumus.
            </div>
          </div>
        </div>
      `;

      // Event Listeners
      const btnAlt = document.getElementById('lab56-btn-alt');
      if (btnAlt) {
        btnAlt.addEventListener('click', () => {
          SpreadsheetSimulatorsCh5.playClickSound();
          showAltTextDrawer = !showAltTextDrawer;
          render();
        });
      }

      const picCard = document.getElementById('lab56-picture-card');
      if (picCard) {
        picCard.addEventListener('click', () => {
          SpreadsheetSimulatorsCh5.playSuccessSound();
          showAltTextDrawer = true;
          render();
        });
      }

      const closeAlt = document.getElementById('lab56-btn-close-alt');
      if (closeAlt) {
        closeAlt.addEventListener('click', () => {
          showAltTextDrawer = false;
          render();
        });
      }

      const altInput = document.getElementById('lab56-alt-input');
      if (altInput) {
        altInput.addEventListener('input', (e) => {
          altTextValue = e.target.value;
        });
      }

      const txtArea = document.getElementById('lab56-textarea');
      if (txtArea) {
        txtArea.addEventListener('input', (e) => {
          customNoteText = e.target.value;
        });
      }

      document.getElementById('lab56-bg-note').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playClickSound();
        textBoxBg = 'note';
        render();
      });

      document.getElementById('lab56-bg-trans').addEventListener('click', () => {
        SpreadsheetSimulatorsCh5.playClickSound();
        textBoxBg = 'transparent';
        render();
      });
    }

    render();
  },

  // Initialize all 6 labs at once
  initAll() {
    this.initLab51();
    this.initLab52();
    this.initLab53();
    this.initLab54();
    this.initLab55();
    this.initLab56();
  }
};

window.SpreadsheetSimulatorsCh5 = SpreadsheetSimulatorsCh5;
