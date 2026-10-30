/**
 * DATA & KURIKULUM PEMBELAJARAN ICT EXCEL - CHAPTER 4
 * Excel Formulas and Functions
 * Versi Khusus Siswa: Panduan Praktik Mandiri & Tutorial Interaktif Langkah demi Langkah di Excel
 */

const CHAPTER_4_DATA = {
  chapterNumber: 4,
  chapterTitle: "Excel Formulas and Functions",
  chapterSubtitle: "Modul Praktik Mandiri & Tutorial Interaktif Microsoft Excel",
  overview: {
    title: "General Theory: Fondasi Kalkulasi, Fungsi Logika, Agregasi, & Audit Rumus Excel",
    description: "Formula dan fungsi adalah inti kekuatan Microsoft Excel dalam otomatisasi perhitungan bisnis dan analisis data. Melalui Chapter 4 ini, siswa akan mempelajari struktur dasar penulisan rumus, perbedaan vital referensi sel relatif ($A$1 vs A1), fungsi statistik utama (MAX, MIN, AVERAGE, LARGE, COUNT, COUNTA), agregasi bersyarat multi-kriteria (COUNTIFS & SUMIFS), manipulasi teks (&, TEXTJOIN, PROPER), pengambilan keputusan otomatis dengan logika IF, serta teknik Formula Auditing menggunakan shortcut Show Formulas (Ctrl+`).",
    coreCompetencies: [
      "Menguasai struktur sintaks rumus Excel dan memahami perbedaan krusial antara Referensi Relatif (A1), Absolut ($A$1), serta Campuran ($A1 / A$1) menggunakan tombol F4",
      "Mampu menggunakan fungsi statistik agregat: MAX, MIN, AVERAGE, LARGE (ranking tertentu), serta membedakan fungsi COUNT (khusus angka) dan COUNTA (seluruh sel terisi)",
      "Menerapkan fungsi kalkulasi bersyarat multi-kriteria COUNTIFS dan SUMIFS yang tahan terhadap pengurutan data (sorting)",
      "Memanipulasi string teks: menggabungkan teks menggunakan operator ampersand (&), CONCATENATE, dan TEXTJOIN, serta merapikan huruf kapital nama dengan PROPER, UPPER, dan LOWER",
      "Mengimplementasikan fungsi logika IF untuk pengambilan keputusan otomatis (Lulus/Remedial) dan menyalin rumus secara efisien menggunakan Fill Down",
      "Melakukan audit dan pelacakan kesalahan rumus lembar kerja menggunakan fitur Show Formulas (Ctrl+`) dan Formula Auditing tools"
    ]
  },
  topics: [
    {
      id: "topic-4-1",
      number: "4.1",
      title: "Formula Structure and Cell References",
      indonesianTitle: "Struktur Rumus & Referensi Sel (Relatif, Absolut, & Campuran)",
      summary: "Every Excel formula starts with an equal sign (=) and can contain operators, cell references, and functions. Cell references can be relative, such as A1, which change when a formula is copied, or absolute, such as $A$1, which always refer to the same cell. Absolute references are useful when a formula must always use the same fixed value or range.",
      shortcut: "F4 (Toggle Absolut / Relatif) / = (Mulai Rumus)",
      ribbonPath: "Formulas → Function Library / Formula Bar",
      theoryExplanation: `Di Microsoft Excel, **Formula (Rumus)** adalah instruksi matematis yang kita berikan kepada program untuk melakukan kalkulasi terhadap nilai-nilai yang ada di dalam lembar kerja.

### 1. Aturan Dasar Struktur Rumus
* **Wajib Dimulai dengan Tanda Sama Dengan (=):** Jika tidak diawali tanda \`=\`, Excel akan menganggap teks tersebut sebagai teks biasa (string) dan tidak melakukan perhitungan apapun.
* **Komponen Rumus:** Rumus dapat memuat operator matematika (\`+\`, \`-\`, \`*\`, \`/\`, \`^\`), referensi sel (misal \`A1\`), nilai konstanta (misal \`100\`), dan nama fungsi (misal \`SUM\`).

### 2. Tiga Jenis Referensi Sel (Cell References)
Pemahaman tentang jenis referensi sel adalah kunci terpenting agar rumus kamu tidak rusak saat disalin (*copy/drag*):

1. **Referensi Relatif (Contoh: A1)**
   * Secara *default*, semua sel di Excel adalah relatif.
   * Ketika rumus disalin ke baris di bawahnya, nomor baris akan otomatis bertambah (\`A1\` menjadi \`A2\`, \`A3\`, dst.).
   * Cocok digunakan untuk menghitung baris per baris, misalnya: \`=C4*D4\` (Jumlah Barang dikali Harga Satuan).

2. **Referensi Absolut (Contoh: $A$1)**
   * Ditandai dengan simbol dollar (\`$\`) di depan huruf kolom dan angka baris.
   * Simbol \`$\` berfungsi mengunci (*lock*) posisi sel sehingga **tidak akan pernah bergeser** ke manapun rumus disalin.
   * Sangat berguna ketika rumus harus merujuk pada sel nilai tetap, seperti: tarif pajak PPN (11%), kurs dollar, atau persentase diskon promo di satu sel tertentu.

3. **Referensi Campuran / Mixed (Contoh: $A1 atau A$1)**
   * \`$A1\`: Kolom A dikunci, tetapi baris dapat bergeser saat disalin ke bawah.
   * \`A$1\`: Baris 1 dikunci, tetapi kolom dapat bergeser saat disalin ke kanan.

### 3. Trik Sakti Tombol F4
Saat sedang mengetik atau mengedit referensi sel di Formula Bar, tekan tombol **F4** di keyboard secara berulang-ulang untuk berpindah mode secara otomatis:
* Tekan F4 ke-1: \`A1\` ➔ \`$A$1\` (Absolut Penuh)
* Tekan F4 ke-2: \`$A$1\` ➔ \`A$1\` (Baris Terkunci)
* Tekan F4 ke-3: \`A$1\` ➔ \`$A1\` (Kolom Terkunci)
* Tekan F4 ke-4: \`$A1\` ➔ \`A1\` (Kembali ke Relatif)`,
      keyPoints: [
        "**Awalan Rumus**: Wajib diawali tanda sama dengan (=). Rumus tanpa tanda = hanya akan dibaca sebagai teks biasa.",
        "**Referensi Relatif (A1)**: Berubah otomatis mengikuti arah salinan rumus (cocok untuk kalkulasi baris per baris).",
        "**Referensi Absolut ($A$1)**: Terkunci permanen pada sel tertentu, tidak bergeser sama sekali saat di-copy/fill down.",
        "**Pintasan F4**: Tombol tercepat untuk mengubah jenis referensi sel tanpa mengetik simbol dollar ($) secara manual."
      ],
      proTips: "Jika laptop kamu memiliki tombol Fn aktif, gunakan kombinasi **Fn + F4** untuk mengunci sel. Selalu gunakan referensi absolut ketika kamu mengambil nilai dari sebuah sel parameter tunggal di luar tabel (seperti sel Tarif Pajak atau Target Penjualan)!",
      commonMistakes: "Lupa mengunci sel nilai tetap (misal lupa membuat $B$1 saat mengalikan pajak). Akibatnya saat rumus disalin ke bawah, referensi bergeser menjadi B2, B3, B4 (sel kosong), sehingga hasil perhitungan pajak pada baris-baris berikutnya menjadi 0 atau memunculkan error #VALUE!.",
      exerciseTask: "Buat tabel perhitungan total belanja dengan mengalikan Qty dan Harga Satuan, lalu hitung Pajak PPN 11% menggunakan referensi absolut ke sel parameter $B$1.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "4.1_CellReferences",
        sampleData: `Parameter Pajak PPN:\t11%\t\t\t\t\t\n\nNo\tNama Barang\tQty\tHarga Satuan\tSubtotal\tPPN 11% ($B$1)\tTotal Bayar\n1\tMouse Wireless\t15\t120000\t\t\t\n2\tKeyboard USB\t8\t185000\t\t\t\n3\tHeadset Gaming\t12\t250000\t\t\t\n4\tWebcam 1080p\t5\t450000\t\t\t\n5\tFlashdisk 64GB\t25\t85000\t\t\t`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <tbody>
    <tr bgcolor="#e8f5e9" style="background-color:#e8f5e9;font-weight:bold;">
      <td colspan="3" style="border:1px solid #0d6836;padding:6px 12px;color:#107c41;">Tarif Pajak PPN (Sel B1):</td>
      <td style="border:1px solid #0d6836;padding:6px 12px;text-align:right;color:#107c41;font-size:12pt;font-weight:bold;">11%</td>
      <td colspan="3" style="border:1px solid #0d6836;padding:6px 12px;color:#64748b;font-size:9pt;font-style:italic;">*Wajib dikunci dengan simbol dollar: $B$1 menggunakan tombol F4</td>
    </tr>
    <tr><td colspan="7" style="border:none;height:12px;"></td></tr>
  </tbody>
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;">No</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;">Nama Barang</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;">Qty</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;">Harga Satuan</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;">Subtotal</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;background-color:#0d6836;">PPN 11% ($B$1)</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;">Total Bayar</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Mouse Wireless</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">15</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">120000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#107c41;font-weight:bold;">=C4*D4</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#d97706;font-weight:bold;">=E4*$B$1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#0284c7;font-weight:bold;">=E4+F4</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Keyboard USB</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">8</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">185000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">3</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Headset Gaming</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">12</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">250000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">4</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Webcam 1080p</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">5</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">450000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">5</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Flashdisk 64GB</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">25</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka aplikasi Microsoft Excel di komputer kamu. Buat **Sheet Baru**, lalu beri nama: `4.1_CellReferences`.",
          "Klik tombol **'📋 Salin Tabel Latihan'** di bawah ini.",
          "Di lembar kerja Excel, klik sel **A1** satu kali saja (jangan double click!), lalu tekan tombol **Ctrl + V** untuk menempelkan tabel.",
          "Pastikan sel **B1** berisi angka `11%` (atau `0.11`) sebagai sel parameter pajak tetap.",
          "**[Hitung Subtotal]** Klik sel **E4** (Subtotal baris Mouse Wireless), ketik rumus: `=C4*D4` lalu tekan **Enter** (Hasil: `1.800.000`). Arahkan kursor mouse ke pojok kanan bawah sel E4 (titik kotak hijau kecil / *Fill Handle*) hingga kursor berubah menjadi tanda tambah hitam `+`, lalu **klik ganda (double-click)** untuk menyalin rumus ke baris 5 sampai 8.",
          "**[Kunci Sel Absolut PPN $B$1]** Klik sel **F4** (kolom PPN), ketik rumus: `=E4*B1`. *PENTING: Jangan tekan Enter dulu!* Langsung tekan tombol **F4** di keyboard kamu (atau **Fn + F4** jika memakai laptop Asus/Lenovo/HP). Perhatikan rumusnya otomatis berubah menjadi `=E4*$B$1`. Sekarang tekan **Enter** (Hasil: `198.000`).",
          "**[Salin PPN ke Bawah]** Arahkan mouse ke pojok kanan bawah sel **F4** hingga kursor menjadi tanda tambah `+`, lalu klik ganda atau tarik ke bawah sampai baris 8. Perhatikan bahwa seluruh baris menghitung pajak secara akurat karena koordinat sel `$B$1` terkunci permanen!",
          "**[Hitung Total Bayar]** Klik sel **G4**, ketik rumus: `=E4+F4` lalu tekan **Enter** (Hasil: `1.998.000`). Klik ganda Fill Handle sel G4 untuk mengisi baris 5 sampai 8.",
          "**[Pemeriksaan Formula]** Tekan pintasan keyboard **Ctrl + `** (tombol backtick di samping angka 1) untuk memeriksa seluruh rumus. Perhatikan sel PPN di baris 8 tetap merujuk pada `=E8*$B$1`. Tekan **Ctrl + `** lagi untuk kembali ke tampilan normal."
        ],
        expectedResult: "Seluruh kolom Subtotal (E4:E8), PPN (F4:F8), dan Total Bayar (G4:G8) terhitung 100% akurat. Sel parameter B1 tidak bergeser saat disalin berkat penguncian absolut $B$1."
      }
    },
    {
      id: "topic-4-2",
      number: "4.2",
      title: "Basic Statistical Functions",
      indonesianTitle: "Fungsi Statistik Dasar (MAX, MIN, AVERAGE, LARGE, COUNT, & COUNTA)",
      summary: "Functions such as MAX, MIN, and AVERAGE are commonly used to summarize numerical data. LARGE finds a value based on its rank, such as the second or third highest value. COUNT counts cells containing numbers, while COUNTA counts cells containing any type of data, including text and numbers.",
      shortcut: "Alt + = (AutoSum) / Shift + F3 (Insert Function)",
      ribbonPath: "Home → Editing → AutoSum (▼) / Formulas → Statistical",
      theoryExplanation: `Dalam analisis data sekolah, toko, maupun kantor, meringkas ratusan baris angka menjadi informasi penting (seperti nilai tertinggi, terendah, dan rata-rata) adalah kebutuhan utama. Excel menyediakan fungsi-fungsi statistik berikut:

### 1. MAX, MIN, dan AVERAGE
* \`=MAX(rentang)\`: Mengambil nilai numerik **tertinggi / maksimum** dari sekumpulan sel. Contoh: mencari skor ujian terbaik siswa.
* \`=MIN(rentang)\`: Mengambil nilai numerik **terendah / minimum** dari sekumpulan sel. Contoh: mencari biaya produksi paling hemat.
* \`=AVERAGE(rentang)\`: Menghitung **rata-rata aritmatika** (penjumlahan seluruh angka dibagi jumlah item numerik).

### 2. Fungsi Ranking: LARGE (dan SMALL)
* \`=LARGE(rentang; k)\`: Mengambil nilai terbesar ke-*k* (ranking tertentu).
  * \`=LARGE(C2:C20; 1)\` menghasilkan nilai terbesar ke-1 (sama dengan MAX).
  * \`=LARGE(C2:C20; 2)\` menghasilkan nilai terbesar ke-2 (Juara 2 / Runner-up).
  * \`=LARGE(C2:C20; 3)\` menghasilkan nilai terbesar ke-3 (Juara 3).
* *Catatan:* Kebalikannya adalah fungsi \`=SMALL(rentang; k)\` untuk mencari nilai terkecil ke-*k*.

### 3. Perbedaan Kritis: COUNT vs COUNTA
Banyak siswa keliru memahami perbedaan kedua fungsi ini:
* **\`=COUNT(rentang)\` (Count Numbers Only):**
  * HANYA menghitung sel yang berisi **data numerik / angka murni**.
  * Mengabaikan sel yang berisi teks (seperti "Sakit", "Izin", "Belum Bayar") serta sel kosong.
  * Sangat cocok untuk menghitung berapa siswa yang sudah memiliki nilai ujian.
* **\`=COUNTA(rentang)\` (Count All / Non-Empty):**
  * Menghitung **semua sel yang TIDAK KOSONG**, baik berisi angka, teks, simbol, atau error.
  * Sangat tepat untuk menghitung total peserta yang terdaftar di daftar hadir.
* **\`=COUNTBLANK(rentang)\`:** Menghitung jumlah sel yang benar-benar kosong.`,
      keyPoints: [
        "**MAX & MIN**: Mengembalikan nilai numerik tertinggi dan terendah pada range.",
        "**AVERAGE**: Menghitung mean aritmatika, otomatis mengabaikan teks dan sel kosong.",
        "**LARGE(array; k)**: Mencari peringkat spesifik (misal k=2 untuk nilai tertinggi ke-2).",
        "**COUNT vs COUNTA**: COUNT hanya menghitung angka, sedangkan COUNTA menghitung semua sel terisi (termasuk teks)."
      ],
      proTips: "Gunakan tombol cepat **Alt + =** di sel tepat di bawah deretan angka untuk menyisipkan rumus \`SUM\` secara instan. Untuk mengganti menjadi AVERAGE atau MAX, klik panah kecil (▼) di samping ikon AutoSum pada tab Home!",
      commonMistakes: "Menggunakan fungsi COUNT pada kolom nama siswa (berisi teks). Rumus =COUNT(B4:B15) akan menghasilkan angka 0 karena COUNT menolak menghitung sel berupa teks! Gunakan =COUNTA(B4:B15) untuk kolom berisi teks.",
      exerciseTask: "Dari data nilai ujian siswa, hitung nilai tertinggi (MAX), terendah (MIN), rata-rata (AVERAGE), nilai juara 2 (LARGE k=2), jumlah yang hadir ujian (COUNT), dan total siswa di kelas (COUNTA).",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "4.2_Statistical",
        sampleData: `No\tNama Siswa\tKelas\tNilai Matematika\n1\tAditya Pratama\t10-RPL 1\t85\n2\tBagas Wicaksono\t10-RPL 1\t92\n3\tCitra Lestari\t10-RPL 1\t78\n4\tDewi Anggraini\t10-RPL 1\tSakit\n5\tFajar Ramadhan\t10-RPL 1\t65\n6\tGita Gutawa\t10-RPL 1\t96\n7\tHendra Wijaya\t10-RPL 1\t88\n8\tIndah Permata\t10-RPL 1\tIzin\n\t\tNilai Tertinggi (MAX):\t\n\t\tNilai Terendah (MIN):\t\n\t\tRata-Rata Ujian (AVERAGE):\t\n\t\tJuara 2 / Ranking 2 (LARGE k=2):\t\n\t\tJumlah Siswa Ikut Ujian (COUNT):\t\n\t\tTotal Siswa Terdaftar (COUNTA):\t`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;">No</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;">Nama Siswa</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;">Kelas</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;">Nilai Matematika</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Aditya Pratama</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL 1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bagas Wicaksono</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL 1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">92</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">3</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Citra Lestari</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL 1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">78</td></tr>
    <tr style="background-color:#fff1f2;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">4</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Dewi Anggraini</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL 1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#e11d48;font-weight:bold;">Sakit</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">5</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Fajar Ramadhan</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL 1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">65</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">6</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Gita Gutawa</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL 1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">96</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">7</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Hendra Wijaya</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL 1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">88</td></tr>
    <tr style="background-color:#fff1f2;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">8</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Indah Permata</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL 1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;color:#e11d48;font-weight:bold;">Izin</td></tr>
    <tr bgcolor="#f8fafc" style="background-color:#f8fafc;font-weight:bold;"><td colspan="3" style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#1e293b;">Nilai Tertinggi (MAX):</td><td style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#107c41;font-weight:bold;">=MAX(D2:D9)</td></tr>
    <tr bgcolor="#f8fafc" style="background-color:#f8fafc;font-weight:bold;"><td colspan="3" style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#1e293b;">Nilai Terendah (MIN):</td><td style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#e11d48;font-weight:bold;">=MIN(D2:D9)</td></tr>
    <tr bgcolor="#f8fafc" style="background-color:#f8fafc;font-weight:bold;"><td colspan="3" style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#1e293b;">Rata-Rata Ujian (AVERAGE):</td><td style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#0284c7;font-weight:bold;">=AVERAGE(D2:D9)</td></tr>
    <tr bgcolor="#f8fafc" style="background-color:#f8fafc;font-weight:bold;"><td colspan="3" style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#1e293b;">Juara 2 / Ranking 2 (LARGE k=2):</td><td style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#d97706;font-weight:bold;">=LARGE(D2:D9; 2)</td></tr>
    <tr bgcolor="#f8fafc" style="background-color:#f8fafc;font-weight:bold;"><td colspan="3" style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#1e293b;">Jumlah Siswa Ikut Ujian (COUNT):</td><td style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#7c3aed;font-weight:bold;">=COUNT(D2:D9)</td></tr>
    <tr bgcolor="#f8fafc" style="background-color:#f8fafc;font-weight:bold;"><td colspan="3" style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#1e293b;">Total Siswa Terdaftar (COUNTA):</td><td style="border:1px solid #cbd5e1;padding:6px 12px;text-align:right;color:#059669;font-weight:bold;">=COUNTA(B2:B9)</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru bernama: `4.2_Statistical`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, buka Excel, klik sel **A1** satu kali, lalu tekan **Ctrl + V**.",
          "**[Nilai Tertinggi - MAX]** Klik sel **D10**, ketik rumus: `=MAX(D2:D9)` lalu tekan **Enter** ➔ Hasil: `96` (skor tertinggi diraih oleh Gita Gutawa).",
          "**[Nilai Terendah - MIN]** Klik sel **D11**, ketik rumus: `=MIN(D2:D9)` lalu tekan **Enter** ➔ Hasil: `65` (skor terendah milik Fajar Ramadhan).",
          "**[Rata-Rata Ujian - AVERAGE]** Klik sel **D12**, ketik rumus: `=AVERAGE(D2:D9)` lalu tekan **Enter** ➔ Hasil: `84`. *Perhatikan: Excel secara otomatis mengabaikan sel teks 'Sakit' dan 'Izin' saat menghitung nilai rata-rata!*",
          "**[Juara 2 / Ranking 2 - LARGE]** Klik sel **D13**, ketik rumus: `=LARGE(D2:D9; 2)` lalu tekan **Enter** ➔ Hasil: `92` (skor Bagas Wicaksono). Angka `2` adalah nilai parameter k untuk ranking terbesar ke-2.",
          "**[Jumlah Siswa Ikut Ujian - COUNT]** Klik sel **D14**, ketik rumus: `=COUNT(D2:D9)` lalu tekan **Enter** ➔ Hasil: `6`. *Ingat: Fungsi COUNT HANYA menghitung sel angka numerik, sehingga 2 siswa yang bertuliskan 'Sakit' dan 'Izin' diabaikan!*",
          "**[Total Siswa Terdaftar - COUNTA]** Klik sel **D15**, ketik rumus: `=COUNTA(B2:B9)` lalu tekan **Enter** ➔ Hasil: `8`. *Fungsi COUNTA menghitung seluruh sel yang terisi (angka maupun nama teks siswa).*"
        ],
        expectedResult: "Seluruh ringkasan statistik berhasil dihitung secara presisi: MAX = 96, MIN = 65, AVERAGE = 84, LARGE(2) = 92, COUNT = 6 orang (angka), dan COUNTA = 8 siswa (termasuk yang sakit/izin)."
      }
    },
    {
      id: "topic-4-3",
      number: "4.3",
      title: "Conditional Aggregation: COUNTIFS and SUMIFS",
      indonesianTitle: "Agregasi Bersyarat Multi-Kriteria: COUNTIFS & SUMIFS",
      summary: "These functions evaluate the actual data values rather than row positions, so the results remain accurate even after the data is sorted. COUNTIFS and SUMIFS calculate values based on one or more conditions. For example, SUMIFS can add values only from rows that meet specific criteria.",
      shortcut: "Shift + F3 (Insert Function) / =SUMIFS(...)",
      ribbonPath: "Formulas → Math & Trig (SUMIFS) / Statistical (COUNTIFS)",
      theoryExplanation: `Ketika data transaksi atau laporan mencapai ratusan baris, kita sering diminta menjawab pertanyaan spesifik, seperti: *"Berapa total penjualan produk Laptop di Cabang Jakarta yang statusnya Lunas?"* atau *"Berapa banyak transaksi di atas 2 juta rupiah?"*.

Fitur **COUNTIFS** dan **SUMIFS** adalah jawaban standar industri untuk kalkulasi bersyarat.

### 1. Mengapa COUNTIFS & SUMIFS Sangat Unggul?
Fungsi ini mengevaluasi **nilai data faktual di dalam sel**, bukan posisi baris. Artinya, ketika tabel data kamu **diurutkan (Sort A-Z), disaring (Filter), atau disisipkan baris baru di tengah-tengah**, hasil kalkulasi COUNTIFS dan SUMIFS **akan selalu 100% akurat dan tidak akan bergeser**.

### 2. Fungsi COUNTIFS (Menghitung Frekuensi Bersyarat)
Menghitung berapa banyak baris data yang memenuhi satu atau beberapa kriteria sekaligus.
* **Sintaks:**
  \`=COUNTIFS(criteria_range1; criteria1; [criteria_range2; criteria2]; ...)\`
* **Contoh Kasus:**
  Menghitung transaksi di cabang "Jakarta" yang berstatus "Lunas":
  \`=COUNTIFS(B2:B20; "Jakarta"; D2:D20; "Lunas")\`

### 3. Fungsi SUMIFS (Menjumlahkan Nilai Bersyarat)
Menjumlahkan angka hanya dari baris-baris yang memenuhi kriteria yang ditentukan.
* **Sintaks:**
  \`=SUMIFS(sum_range; criteria_range1; criteria1; [criteria_range2; criteria2]; ...)\`
* **PENTING - Urutan Parameter SUMIFS:**
  Rentang angka yang akan dijumlahkan (\`sum_range\`) diletakkan di **paling depan**, baru kemudian diikuti pasangan rentang kriteria dan kriterianya yang dipisahkan dengan tanda titik koma (\`;\`).
* **Contoh Kasus:**
  Menjumlahkan omzet penjualan di cabang "Jakarta" dengan status "Lunas":
  \`=SUMIFS(E2:E20; B2:B20; "Jakarta"; D2:D20; "Lunas")\`

### 4. Aturan Penulisan Operator Pembanding
Jika kriteria menggunakan operator perbandingan (\`>\`, \`<\`, \`>=\`, \`<=\`, \`<>\`), operator wajib diapit **tanda petik dua**:
* \`">=1000000"\`: Bernilai satu juta ke atas.
* \`"<>Batal"\`: Semua status yang bukan "Batal".`,
      keyPoints: [
        "**Evaluasi Nilai Aktual**: Kebal terhadap pengurutan data (Sort). Hasil tetap konsisten di mana pun letak barisnya.",
        "**Multi-Kriteria**: Mendukung banyak kondisi logika AND secara simultan dalam satu rumus.",
        "**Struktur SUMIFS**: Parameter \`sum_range\` (rentang angka yang dijumlah) selalu berada di urutan pertama.",
        "**Pemisah Titik Koma (;)**: Di Excel standar Indonesia, pisahkan setiap kriteria dengan titik koma, dan apit kriteria teks dengan tanda kutip dua."
      ],
      proTips: "Alih-alih mengetik nama kriteria secara langsung di dalam rumus (hardcoded seperti `\"Jakarta\"`), kamu bisa mereferensikan ke sel lain (misal `=SUMIFS(E2:E20; B2:B20; G1)`). Dengan cara ini, laporan kamu menjadi dinamis: cukup ganti teks di sel G1, maka totalnya langsung berubah otomatis!",
      commonMistakes: "Tertukar urutan argumen antara SUMIF tunggal dan SUMIFS jamak. Pada SUMIF lama, sum_range ada di belakang. Tetapi pada SUMIFS modern, sum_range WAJIB di depan! Jangan sampai terbalik.",
      exerciseTask: "Dari tabel transaksi penjualan, hitung frekuensi transaksi (COUNTIFS) dan total nominal penjualan (SUMIFS) untuk Cabang Surabaya yang berstatus Lunas menggunakan pemisah titik koma (;).",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "4.3_ConditionalAggregation",
        sampleData: `ID\tCabang\tKategori\tStatus\tTotal Penjualan\t\tAnalisis Bisnis Multi-Kriteria\tHasil Kalkulasi\nTRX-01\tJakarta\tHardware\tLunas\t3500000\t\tJml Transaksi Surabaya Lunas (COUNTIFS):\t\nTRX-02\tSurabaya\tSoftware\tLunas\t1200000\t\tTotal Omzet Surabaya Lunas (SUMIFS):\t\nTRX-03\tJakarta\tHardware\tPending\t2100000\t\tJml Transaksi Penjualan >= 2 Juta (COUNTIFS):\t\nTRX-04\tBandung\tAksesoris\tLunas\t850000\t\tTotal Omzet Jakarta Hardware (SUMIFS):\t\nTRX-05\tSurabaya\tHardware\tLunas\t4500000\t\t\t\nTRX-06\tJakarta\tSoftware\tLunas\t1800000\t\t\t\nTRX-07\tBandung\tHardware\tPending\t3000000\t\t\t\nTRX-08\tSurabaya\tAksesoris\tPending\t600000\t\t\t`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;">ID</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;">Cabang</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;">Kategori</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;">Status</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;">Total Penjualan</th>
      <th style="border:none;background:#ffffff;width:16px;"></th>
      <th style="border:1px solid #0284c7;background:#0284c7;color:#ffffff;padding:6px 14px;text-align:left;">Analisis Bisnis Multi-Kriteria</th>
      <th style="border:1px solid #0284c7;background:#0284c7;color:#ffffff;padding:6px 14px;text-align:right;">Hasil Kalkulasi Formula</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">TRX-01</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Jakarta</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Hardware</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#107c41;font-weight:bold;">Lunas</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">3500000</td><td style="border:none;"></td><td style="border:1px solid #cbd5e1;padding:5px 10px;font-weight:600;">Jml Transaksi Surabaya Lunas:</td><td style="border:1px solid #cbd5e1;padding:5px 10px;text-align:right;color:#0284c7;font-weight:bold;">=COUNTIFS(B2:B9; "Surabaya"; D2:D9; "Lunas")</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">TRX-02</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Surabaya</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Software</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#107c41;font-weight:bold;">Lunas</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">1200000</td><td style="border:none;"></td><td style="border:1px solid #cbd5e1;padding:5px 10px;font-weight:600;">Total Omzet Surabaya Lunas:</td><td style="border:1px solid #cbd5e1;padding:5px 10px;text-align:right;color:#107c41;font-weight:bold;">=SUMIFS(E2:E9; B2:B9; "Surabaya"; D2:D9; "Lunas")</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">TRX-03</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Jakarta</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Hardware</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#f59e0b;font-weight:bold;">Pending</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">2100000</td><td style="border:none;"></td><td style="border:1px solid #cbd5e1;padding:5px 10px;font-weight:600;">Jml Trx Penjualan &gt;= 2 Juta:</td><td style="border:1px solid #cbd5e1;padding:5px 10px;text-align:right;color:#0284c7;font-weight:bold;">=COUNTIFS(E2:E9; "&gt;=2000000")</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">TRX-04</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bandung</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Aksesoris</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#107c41;font-weight:bold;">Lunas</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">850000</td><td style="border:none;"></td><td style="border:1px solid #cbd5e1;padding:5px 10px;font-weight:600;">Total Omzet Jakarta Hardware:</td><td style="border:1px solid #cbd5e1;padding:5px 10px;text-align:right;color:#107c41;font-weight:bold;">=SUMIFS(E2:E9; B2:B9; "Jakarta"; C2:C9; "Hardware")</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">TRX-05</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Surabaya</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Hardware</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#107c41;font-weight:bold;">Lunas</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">4500000</td><td colspan="3" style="border:none;"></td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">TRX-06</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Jakarta</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Software</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#107c41;font-weight:bold;">Lunas</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">1800000</td><td colspan="3" style="border:none;"></td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">TRX-07</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bandung</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Hardware</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#f59e0b;font-weight:bold;">Pending</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">3000000</td><td colspan="3" style="border:none;"></td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">TRX-08</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Surabaya</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Aksesoris</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#f59e0b;font-weight:bold;">Pending</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">600000</td><td colspan="3" style="border:none;"></td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru bernama: `4.3_ConditionalAggregation`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** satu kali di Excel, lalu tekan **Ctrl + V**.",
          "**[Frekuensi Trx Surabaya Lunas - COUNTIFS]** Klik sel **H2**, ketik rumus: `=COUNTIFS(B2:B9; \"Surabaya\"; D2:D9; \"Lunas\")` lalu tekan **Enter** ➔ Hasil: `2`. *(Artinya ada 2 transaksi yang memenuhi kriteria Cabang Surabaya DAN Status Lunas)*.",
          "**[Total Omzet Surabaya Lunas - SUMIFS]** Klik sel **H3**, ketik rumus: `=SUMIFS(E2:E9; B2:B9; \"Surabaya\"; D2:D9; \"Lunas\")` lalu tekan **Enter** ➔ Hasil: `5.700.000` (penjumlahan TRX-02 1.200.000 + TRX-05 4.500.000). *PENTING: Rentang nilai yang dijumlahkan (E2:E9) wajib ditulis di urutan pertama dan gunakan pemisah titik koma (;)!*",
          "**[Transaksi Di Atas 2 Juta]** Klik sel **H4**, ketik rumus: `=COUNTIFS(E2:E9; \">=2000000\")` lalu tekan **Enter** ➔ Hasil: `4`. *Kriteria operator perbandingan wajib diapit tanda kutip dua \">=2000000\"*.",
          "**[Total Omzet Jakarta Hardware]** Klik sel **H5**, ketik rumus: `=SUMIFS(E2:E9; B2:B9; \"Jakarta\"; C2:C9; \"Hardware\")` lalu tekan **Enter** ➔ Hasil: `3.500.000`.",
          "**[Uji Kebal Sort]** Blok tabel transaksi **A1:E9**, buka tab Ribbon **Data → Sort**, lalu urutkan berdasarkan Total Penjualan (*Largest to Smallest*). Perhatikan bahwa nilai kalkulasi di sel H2 sampai H5 tetap 100% tepat dan tidak bergeser, membuktikan keunggulan evaluasi nilai aktual pada COUNTIFS dan SUMIFS!"
        ],
        expectedResult: "Seluruh kalkulasi bersyarat memberikan hasil akurat (2 transaksi Surabaya Lunas senilai Rp 5.700.000), dan nilai rumus tetap konsisten meskipun urutan baris data diacak atau disortir."
      }
    },
    {
      id: "topic-4-4",
      number: "4.4",
      title: "Text Functions: Combining and Changing Text Case",
      indonesianTitle: "Fungsi Teks: Penggabungan (&, CONCATENATE, TEXTJOIN) & Format Huruf (PROPER, UPPER, LOWER)",
      summary: "Text can be combined using the & operator or functions such as CONCATENATE and TEXTJOIN. Separators such as commas and spaces can also be added. The PROPER function changes text so that the first letter of each word is capitalized, which is useful for cleaning up names and titles.",
      shortcut: "& (Ampersand Operator) / =PROPER(...)",
      ribbonPath: "Formulas → Text → CONCAT / TEXTJOIN / PROPER",
      theoryExplanation: `Sering kali data nama siswa, pelanggan, atau alamat yang diimpor dari Google Form atau aplikasi eksternal memiliki format penulisan yang berantakan (misalnya huruf kapital campur aduk: \`aHMaD fAuZI\` atau nama depan dan nama belakang terpisah di kolom berbeda). Excel memiliki fungsi pengolah teks profesional untuk merapikannya.

### 1. Menggabungkan Teks (Combining Text)
Ada 3 cara menggabungkan teks dari beberapa sel:

1. **Operator Ampersand (&):**
   * Cara paling ringkas dan paling sering digunakan.
   * Format: \`=sel1 & "spasi/tanda" & sel2\`
   * Contoh: \`=B2 & " " & C2\` (menggabungkan Nama Depan di B2 dan Nama Belakang di C2 dengan spasi pemisah).

2. **Fungsi CONCATENATE / CONCAT:**
   * Sintaks: \`=CONCATENATE(text1; " "; text2)\` atau \`=CONCAT(B2:C2)\`.

3. **Fungsi Modern: TEXTJOIN (Paling Powerful):**
   * Sintaks: \`=TEXTJOIN(delimiter; ignore_empty; text1; text2; ...)\`
   * \`delimiter\`: Tanda pemisah yang kamu inginkan (misal \`", "\` atau \`" - "\`).
   * \`ignore_empty\`: Masukkan \`TRUE\` agar sel kosong otomatis dilewati tanpa menyisakan pemisah ganda!
   * Contoh: \`=TEXTJOIN(", "; TRUE; B2; C2; D2)\` (menggabungkan Nama, Kelas, dan Kota dengan koma spasi).

### 2. Mengubah Format Huruf (Changing Text Case)
Excel menyediakan 3 fungsi vital untuk menstandardisasi huruf kapital:

1. **PROPER (Title Case / Huruf Kapital di Awal Setiap Kata):**
   * Mengubah huruf pertama setiap kata menjadi **KAPITAL**, dan sisa hurufnya menjadi huruf kecil.
   * Sangat berguna untuk membersihkan nama orang, judul buku, dan nama jalan.
   * Contoh: \`=PROPER("aHMad fAuZi")\` ➔ menghasilkan \`"Ahmad Fauzi"\`.

2. **UPPER (Uppercase / Huruf Kapital Semua):**
   * Mengubah seluruh karakter menjadi huruf besar semua.
   * Contoh: \`=UPPER("jakarta")\` ➔ menghasilkan \`"JAKARTA"\`. Cocok untuk kode plat nomor, kode POS, atau akronim.

3. **LOWER (Lowercase / Huruf Kecil Semua):**
   * Mengubah seluruh teks menjadi huruf kecil semua.
   * Contoh: \`=LOWER("AHMAD@GMAIL.COM")\` ➔ menghasilkan \`"ahmad@gmail.com"\`. Cocok untuk standardisasi alamat email dan username.`,
      keyPoints: [
        "**Operator &**: Metode penggabungan tercepat: \`=B2 & \" \" & C2\`.",
        "**TEXTJOIN**: Fungsi modern cerdas untuk menggabungkan banyak sel dengan satu delimiter dan mengabaikan sel kosong (ignore_empty=TRUE).",
        "**PROPER**: Mengkapitalisasi huruf pertama setiap kata — esensial untuk membersihkan nama orang dan gelar.",
        "**UPPER & LOWER**: Menyeragamkan seluruh teks menjadi kapital penuh atau huruf kecil penuh."
      ],
      proTips: "Kamu bisa menggabungkan fungsi teks di dalam satu rumus sekaligus! Contoh: `=PROPER(B2 & \" \" & C2)` akan menggabungkan nama depan dan nama belakang sekaligus merapikan huruf kapitalnya dalam satu langkah efisien!",
      commonMistakes: "Menggabungkan dua sel tanpa menyisipkan spasi di tengahnya (misal: =B2&C2). Hasilnya nama depan dan belakang akan menempel menjadi 'AhmadFauzi'. Selalu sertakan tanda petik berisi spasi: =B2 & \" \" & C2.",
      exerciseTask: "Gabungkan kolom Nama Depan dan Nama Belakang yang berantakan menjadi Nama Lengkap yang rapi menggunakan formula =PROPER(B2 & \" \" & C2), lalu buat Kode Profil siswa menggunakan TEXTJOIN.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "4.4_TextFunctions",
        sampleData: `No\tNama Depan (Mentah)\tNama Belakang (Mentah)\tKelas\tKota Asal\tNama Rapi (PROPER & &)\tProfil Siswa (TEXTJOIN)\n1\taHMaD\tkURniAWan\t10-RPL\tbandung\t\t\n2\tsITi\taMinAH\t10-TKJ\tJAKARTA\t\t\n3\tbUDi\tsANToso\t10-RPL\tsurabaya\t\t\n4\tdInDa\tpUTrI\t10-DKV\tSEMARANG\t\t\n5\trIZkY\tpRaTAmA\t10-TKJ\tyogyakarta\t\t`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;">No</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;">Nama Depan (Mentah)</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;">Nama Belakang (Mentah)</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;">Kelas</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;">Kota Asal</th>
      <th style="border:1px solid #0d6836;padding:6px 14px;text-align:left;background-color:#0d6836;">Nama Rapi (PROPER & &)</th>
      <th style="border:1px solid #0d6836;padding:6px 14px;text-align:left;background-color:#0d6836;">Profil Siswa (TEXTJOIN)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">aHMaD</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">kURniAWan</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">bandung</td><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#107c41;font-weight:bold;">=PROPER(B2 & " " & C2)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#0284c7;font-weight:bold;">=TEXTJOIN(" - "; TRUE; F2; D2; UPPER(E2))</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">sITi</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">aMinAH</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-TKJ</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">JAKARTA</td><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">3</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">bUDi</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">sANToso</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">surabaya</td><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">4</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">dInDa</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">pUTrI</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-DKV</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">SEMARANG</td><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">5</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">rIZkY</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">pRaTAmA</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-TKJ</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">yogyakarta</td><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#94a3b8;font-style:italic;">(Fill Down)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru bernama: `4.4_TextFunctions`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** satu kali di Excel, lalu tekan **Ctrl + V**.",
          "**[Gabungkan Teks & Rapikan Huruf Kapital]** Klik sel **F2** (kolom Nama Rapi), ketik rumus: `=PROPER(B2 & \" \" & C2)` lalu tekan **Enter** ➔ Hasil: `Ahmad Kurniawan`.",
          "Perhatikan: Teks mentah `aHMaD` dan `kURniAWan` otomatis digabungkan dengan spasi pemisah dan huruf awalnya dibuat kapital rapi oleh fungsi `PROPER`!",
          "Arahkan kursor ke pojok kanan bawah sel F2 hingga menjadi tanda tambah `+`, lalu **klik ganda (double-click)** untuk menyalin rumus ke baris 3 sampai 6.",
          "**[Gabungkan Profil dengan TEXTJOIN]** Klik sel **G2** (kolom Profil Siswa), ketik rumus: `=TEXTJOIN(\" - \"; TRUE; F2; D2; UPPER(E2))` lalu tekan **Enter** ➔ Hasil: `Ahmad Kurniawan - 10-RPL - BANDUNG`.",
          "Perhatikan: Delimiter `\" - \"` memisahkan data dengan rapi, `TRUE` mengabaikan sel kosong, dan fungsi `UPPER(E2)` otomatis membuat nama kota menjadi huruf kapital semua!",
          "Klik ganda Fill Handle sel **G2** ke bawah untuk menyelesaikan seluruh profil siswa."
        ],
        expectedResult: "Seluruh nama mentah yang berantakan berubah rapi dengan format Title Case (PROPER), dan kolom profil siswa tergabung secara terstruktur dan profesional menggunakan fungsi TEXTJOIN."
      }
    },
    {
      id: "topic-4-5",
      number: "4.5",
      title: "Logical Function: IF",
      indonesianTitle: "Fungsi Logika IF: Pengambilan Keputusan Otomatis & Salin Rumus (Fill Down)",
      summary: "The IF function checks a condition and returns one value if the condition is true and another if it is false. Example: =IF(condition; value_if_true; value_if_false). It can be used to create automatic decisions, such as determining whether a student Passes or Fails based on a minimum score. Fill Down can then copy the formula to other rows while automatically adjusting relative references.",
      shortcut: "Ctrl + D (Fill Down) / =IF(...)",
      ribbonPath: "Formulas → Logical → IF",
      theoryExplanation: `Fungsi **IF** adalah salah satu fungsi paling mendasar dan terpenting dalam logika komputasi. Fungsi ini memungkinkan lembar kerja Excel mengambil keputusan otomatis seperti halnya program komputer: *"Jika kondisi X terpenuhi, maka lakukan A; jika tidak terpenuhi, maka lakukan B"*.

### 1. Struktur Sintaks Fungsi IF
Sintaks resmi fungsi IF terdiri dari 3 bagian yang dipisahkan tanda titik koma (\`;\`):
\`=IF(kondisi_logika; nilai_jika_benar; nilai_jika_salah)\`

1. **kondisi_logika (Logical Test):** Pernyataan perbandingan yang menghasilkan TRUE (Benar) atau FALSE (Salah).
   * Contoh: \`C2>=75\` (Apakah nilai siswa di sel C2 lebih besar atau sama dengan KKM 75?).
2. **nilai_jika_benar (Value if True):** Apa yang ditampilkan jika kondisi bernilai TRUE.
   * Jika berupa teks, wajib diapit tanda petik dua, misal: \`"LULUS"\`.
3. **nilai_jika_salah (Value if False):** Apa yang ditampilkan jika kondisi bernilai FALSE.
   * Misal: \`"REMEDIAL"\`.

**Rumus Lengkap Penentuan Kelulusan Siswa (KKM 75):**
\`=IF(C2>=75; "LULUS"; "REMEDIAL")\`

### 2. Menggandakan Rumus dengan Fitur Fill Down
Setelah kamu menulis rumus IF di baris pertama (sel E2), kamu tidak perlu mengetik ulang rumus tersebut untuk baris-baris siswa berikutnya!

Excel menyediakan fitur **Fill Down**:
* **Cara Cepat Mouse:** Klik sel E2, arahkan kursor ke titik kotak hijau kecil di pojok kanan bawah sel (*Fill Handle*) hingga kursor berubah menjadi tanda tambah hitam (+), lalu **klik ganda (double-click)**.
* **Cara Keyboard:** Blok sel E2 sampai E8, lalu tekan **Ctrl + D** (Fill Down).
* **Mekanisme Otomatis:** Karena referensi sel \`C2\` bersifat **relatif**, maka saat di-fill down ke baris 3, rumusnya otomatis menyesuaikan menjadi \`=IF(C3>=75; "LULUS"; "REMEDIAL")\`, baris 4 menjadi \`C4\`, dan seterusnya tanpa kesalahan!

### 3. Pengembangan: IF Bertingkat (Nested IF)
Untuk menentukan predikat nilai (A, B, C, D):
\`=IF(C2>=85; "A"; IF(C2>=75; "B"; IF(C2>=65; "C"; "D")))\``,
      keyPoints: [
        "**3 Parameter IF**: (1) Pengujian logika, (2) Output jika TRUE, (3) Output jika FALSE.",
        "**Pemisah Titik Koma (;)**: Gunakan tanda titik koma (;) untuk memisahkan argumen rumus di Excel standar Indonesia.",
        "**Nilai Teks dalam Rumus**: Teks output seperti \"LULUS\" atau \"REMEDIAL\" wajib diapit tanda kutip dua.",
        "**Fill Down (Ctrl + D)**: Menyalin rumus ke bawah secara instan sambil otomatis menyesuaikan referensi relatif baris per baris.",
        "**Nested IF**: Memasukkan fungsi IF di dalam IF untuk menangani lebih dari 2 kemungkinan kondisi (seperti konversi nilai angka ke huruf A/B/C/D)."
      ],
      proTips: "Di Microsoft Excel komputer Indonesia, pemisah rumus adalah **titik koma (;)**: `=IF(C2>=75; \"LULUS\"; \"REMEDIAL\")`. Jika Excel di komputer kamu menggunakan format English/US, Excel akan meminta tanda koma (,). Cek tooltip bantuan saat mengetik rumus!",
      commonMistakes: "Lupa memberi tanda kutip dua pada teks output, misalnya mengetik =IF(C2>=75; LULUS; REMEDIAL). Excel akan memunculkan error #NAME? karena mengira LULUS dan REMEDIAL adalah nama fungsi yang tidak dikenal. Selalu gunakan tanda kutip: =IF(C2>=75; \"LULUS\"; \"REMEDIAL\").",
      exerciseTask: "Terapkan rumus =IF(C2>=75; \"LULUS\"; \"REMEDIAL\") pada kolom Status Kelulusan, lalu gunakan shortcut Ctrl+D atau klik ganda fill handle untuk mengisi seluruh baris siswa.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "4.5_LogicalIF",
        sampleData: `No\tNama Siswa\tNilai Akhir\tKKM\tStatus Kelulusan (IF)\tPredikat Huruf (Nested IF)\n1\tAlifia Zahra\t88\t75\t\t\n2\tBintang Samudra\t68\t75\t\t\n3\tCantika Dewi\t74\t75\t\t\n4\tDimas Anggara\t95\t75\t\t\n5\tErlangga Putra\t60\t75\t\t\n6\tFarhan Maulana\t78\t75\t\t\n7\tGisella Anastasia\t82\t75\t\t`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;">No</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;">Nama Siswa</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;">Nilai Akhir</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;">KKM</th>
      <th style="border:1px solid #0d6836;padding:6px 14px;text-align:center;background-color:#0d6836;">Status Kelulusan (IF)</th>
      <th style="border:1px solid #0d6836;padding:6px 14px;text-align:center;background-color:#0d6836;">Predikat Huruf (Nested IF)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Alifia Zahra</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">88</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">75</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#107c41;font-weight:bold;">=IF(C2&gt;=75; "LULUS"; "REMEDIAL")</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#0284c7;font-weight:bold;">=IF(C2&gt;=85; "A"; IF(C2&gt;=75; "B"; ...))</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bintang Samudra</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">68</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">75</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down Ctrl+D)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">3</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Cantika Dewi</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">74</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">75</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down Ctrl+D)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">4</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Dimas Anggara</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">95</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">75</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down Ctrl+D)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">5</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Erlangga Putra</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">60</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">75</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down Ctrl+D)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">6</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Farhan Maulana</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">78</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">75</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down Ctrl+D)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">7</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Gisella Anastasia</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">82</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">75</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down Ctrl+D)</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;color:#94a3b8;font-style:italic;">(Fill Down)</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru bernama: `4.5_LogicalIF`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, buka Excel, klik sel **A1** satu kali, lalu tekan **Ctrl + V**.",
          "**[Tentukan Status Kelulusan - Rumus IF]** Klik sel **E2** (kolom Status Kelulusan Alifia Zahra).",
          "Ketik rumus berikut: `=IF(C2>=75; \"LULUS\"; \"REMEDIAL\")` lalu tekan **Enter** ➔ Hasil: `LULUS` (karena nilai Alifia 88 >= KKM 75). *Pemisah resmi yang digunakan adalah titik koma (;)*.",
          "**[Trik Salin Instan Keyboard - Ctrl+D]** Blok rentang sel **E2 sampai E8**, lalu tekan kombinasi keyboard: **Ctrl + D** (Fill Down). Seluruh status kelulusan langsung terisi otomatis!",
          "Perhatikan: Siswa dengan nilai di bawah KKM 75 (Bintang: 68, Cantika: 74, Erlangga: 60) otomatis mendapat status `REMEDIAL`!",
          "**[Konversi Nilai Huruf - Nested IF]** Klik sel **F2** (kolom Predikat Huruf), ketik rumus IF bertingkat: `=IF(C2>=85; \"A\"; IF(C2>=75; \"B\"; IF(C2>=65; \"C\"; \"D\")))` lalu tekan **Enter** ➔ Hasil: `A`.",
          "Klik ganda Fill Handle pada sudut kanan bawah sel **F2** untuk mengisi seluruh predikat siswa ke bawah."
        ],
        expectedResult: "Seluruh status kelulusan siswa dievaluasi secara instan: 4 siswa LULUS dan 3 siswa REMEDIAL, serta predikat huruf A/B/C/D terisi sempurna di seluruh baris berkat fitur Fill Down."
      }
    },
    {
      id: "topic-4-6",
      number: "4.6",
      title: "Formula Auditing: Show Formulas",
      indonesianTitle: "Pemeriksaan Rumus (Formula Auditing): Fitur Show Formulas & Pelacakan Error",
      summary: "Normally, Excel displays the result of a formula rather than the formula itself. Show Formulas (Formulas > Formula Auditing > Show Formulas or Ctrl+`) displays the formulas throughout the worksheet. This feature is useful for checking formulas, finding errors, and understanding how calculations work.",
      shortcut: "Ctrl + ` (Grave Accent / Backtick) / Alt + M + H",
      ribbonPath: "Formulas → Formula Auditing → Show Formulas",
      theoryExplanation: `Secara normal, Microsoft Excel selalu menampilkan **hasil akhir kalkulasi** (seperti angka \`1.500.000\`) di dalam sel, bukan teks rumusnya (\`=C4*D4\`). Hal ini sangat baik untuk membaca laporan, namun menjadi tantangan besar saat kita harus **memeriksa kebenaran rumus, mencari sumber error, atau mengoreksi lembar kerja rekan kerja/siswa**.

### 1. Fitur Show Formulas
Fitur **Show Formulas** mengubah tampilan seluruh worksheet secara instan: alih-alih menampilkan hasil angka, Excel akan **membuka dan menampilkan rumus asli di setiap sel**.
* **Shortcut Super Cepat:** Tekan tombol **Ctrl + \`** (tombol *grave accent* / backtick, terletak tepat di sebelah kiri angka 1 / di bawah tombol Esc).
* **Melalui Ribbon:** Buka tab **Formulas** ➔ di grup *Formula Auditing* ➔ klik **Show Formulas**.
* **Karakteristik Tampilan:** Lebar kolom akan otomatis membesar sedikit agar seluruh teks formula dapat terbaca jelas. Tekan **Ctrl + \`** sekali lagi untuk mengembalikan tampilan ke hasil normal.

### 2. Manfaat Utama Formula Auditing
1. **Mendeteksi Angka Palsu (Hardcoded Data):** Sering kali ada pengguna yang mengetik angka secara manual (bukan menggunakan rumus). Dengan Show Formulas, sel yang bukan rumus akan langsung terlihat mencolok.
2. **Pemeriksaan Konsistensi Rumus:** Memastikan seluruh baris dari atas ke bawah menggunakan rumus yang seragam dan tidak ada yang terlewat.
3. **Mendiagnosis Kode Error Excel yang Sering Muncul:**
   * \`#VALUE!\`: Terjadi ketika tipe data tidak sesuai, misalnya sel angka dikalikan dengan sel berisi teks (\`=100*"Lunas"\`).
   * \`#DIV/0!\`: Terjadi saat suatu angka dibagi dengan angka nol atau sel yang kosong (\`=A1/0\`).
   * \`#REF!\`: Referensi sel sudah terhapus atau dipindahkan secara tidak sengaja.
   * \`#NAME?\`: Salah mengetik nama fungsi (misal mengetik \`=SUMM(...)\` atau lupa tanda kutip pada teks).

### 3. Trace Precedents & Trace Dependents
* **Trace Precedents (Lacak Asal Data):** Menarik panah biru visual yang menunjukkan sel-sel mana saja yang menjadi sumber data untuk rumus sel aktif.
* **Trace Dependents (Lacak Ketergantungan):** Menarik panah yang menunjukkan sel-sel lain mana saja yang bergantung pada hasil sel aktif saat ini.
* **Remove Arrows:** Menghapus seluruh panah pelacak setelah selesai memeriksa.`,
      keyPoints: [
        "**Shortcut Ctrl + `**: Menampilkan semua rumus di lembar kerja secara massal dalam satu detik.",
        "**Pemeriksaan Integritas Data**: Membedakan dengan jelas antara sel yang dihitung dengan rumus dan sel yang diketik manual (hardcoded).",
        "**Analisis Sumber Error**: Membantu membedah akar permasalahan error seperti #VALUE!, #DIV/0!, #REF!, dan #NAME?.",
        "**Trace Precedents / Dependents**: Menampilkan panah biru visual yang menghubungkan alur kalkulasi antar sel."
      ],
      proTips: "Saat kamu mencetak (*print*) tugas Excel siswa untuk penilaian, aktifkan fitur **Show Formulas (Ctrl + `)** terlebih dahulu sebelum mencetak. Dengan begitu, guru atau penguji dapat langsung melihat rumus apa yang kamu gunakan tanpa harus membuka file digitalnya satu per satu!",
      commonMistakes: "Panik ketika seluruh tabel tiba-tiba melebar dan angka-angka berubah menjadi teks rumus panjang. Siswa mengira tabelnya rusak, padahal mereka tidak sengaja menekan kombinasi Ctrl + `. Cukup tekan Ctrl + ` sekali lagi untuk mengembalikannya ke normal.",
      exerciseTask: "Buka lembar kerja latihan yang berisi rumus kalkulasi belanja dan error #DIV/0!, aktifkan Show Formulas dengan shortcut Ctrl+`, periksa asal rumus, lalu lacak ketergantungan sel menggunakan Trace Precedents.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "4.6_ShowFormulas",
        sampleData: `Item Belanja\tHarga\tQty\tSubtotal\tDiskon Member\tTotal Bersih\nMonitor LED\t2200000\t2\t=B2*C2\t=IF(D2>2000000;D2*10%;0)\t=D2-E2\nMouse Gaming\t350000\t4\t=B3*C3\t=IF(D3>2000000;D3*10%;0)\t=D3-E3\nKeyboard Mech\t750000\t3\t=B4*C4\t=IF(D4>2000000;D4*10%;0)\t=D4-E4\nRata-Rata Subtotal:\t\t\t=AVERAGE(D2:D4)\tTotal Pengeluaran:\t=SUM(F2:F4)`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;">Item Belanja</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;">Harga</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;">Qty</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;">Subtotal</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;">Diskon Member</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;">Total Bersih</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Monitor LED</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">2200000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">4400000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">440000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">3960000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">Mouse Gaming</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">350000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">4</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">1400000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">0</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">1400000</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Keyboard Mech</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">750000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">3</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">2250000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">225000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">2025000</td></tr>
    <tr bgcolor="#f8fafc" style="background-color:#f8fafc;font-weight:bold;">
      <td colspan="3" style="border:1px solid #cbd5e1;padding:5px 10px;text-align:right;">Rata-Rata Subtotal:</td>
      <td style="border:1px solid #cbd5e1;padding:5px 10px;text-align:right;color:#0284c7;">2683333</td>
      <td style="border:1px solid #cbd5e1;padding:5px 10px;text-align:right;">Total Pengeluaran:</td>
      <td style="border:1px solid #cbd5e1;padding:5px 10px;text-align:right;color:#107c41;">7385000</td>
    </tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru bernama: `4.6_ShowFormulas`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** satu kali, lalu tekan **Ctrl + V**.",
          "Perhatikan saat ini lembar kerja menampilkan angka hasil perhitungan normal.",
          "Sekarang tekan tombol **Ctrl + \`** pada keyboard kamu (tombol di sebelah kiri angka 1).",
          "Perhatikan keajaibannya! Semua angka di kolom Subtotal, Diskon, dan Total seketika berubah menjadi teks rumus aslinya (misal `=B2*C2`, `=D2-E2`, `=SUM(...)`).",
          "Klik sel **F2** (sel rumus Total Bersih). Buka tab **Formulas** di Ribbon ➔ klik tombol **Trace Precedents**.",
          "Lihat panah biru yang muncul: panah ini menunjuk langsung dari sel D2 dan E2 ke sel F2, membuktikan bahwa F2 dihitung dari kedua sel tersebut!",
          "Klik tombol **Remove Arrows** untuk menghilangkan panah.",
          "Tekan tombol **Ctrl + \`** sekali lagi untuk mengembalikan tampilan ke hasil angka normal."
        ],
        expectedResult: "Tampilan worksheet beralih mulus antara mode hasil angka dan mode audit rumus (Show Formulas), serta panah visual Trace Precedents berhasil memetakan alur asal usul data rumus."
      }
    }
  ],
  quizQuestions: [
    {
      id: 1,
      topicRef: "4.1",
      question: "Seorang siswa membuat rumus perkalian komisi penjualan di sel D2: `=C2*B1`. Sel B1 berisi persentase komisi tetap (5%). Ketika rumus tersebut disalin ke sel D3 menggunakan Fill Down, hasilnya menjadi 0. Mengapa hal ini terjadi dan bagaimana solusinya?",
      options: [
        "Excel mengalami error internal, siswa harus me-restart komputer",
        "Referensi sel B1 bersifat relatif sehingga bergeser menjadi B2 (sel kosong). Solusinya adalah mengunci sel B1 menjadi absolut: `=C2*$B$1` menggunakan tombol F4",
        "Rumus perkalian di Excel tidak boleh menggunakan simbol bintang (*)",
        "Siswa lupa mengubah format sel menjadi mata uang Rupiah"
      ],
      correctIndex: 1,
      explanation: "Referensi relatif (B1) akan otomatis bergeser menjadi B2 saat disalin ke bawah. Karena sel B2 tidak berisi persentase komisi (kosong atau teks), hasil perkalian menjadi 0. Menekan tombol F4 untuk mengunci menjadi $B$1 (referensi absolut) memastikan pengali komisi tetap merujuk pada sel B1 di baris mana pun."
    },
    {
      id: 2,
      topicRef: "4.2",
      question: "Pada daftar hadir ujian 30 siswa, terdapat 25 siswa yang memiliki nilai angka, 3 siswa bertuliskan 'Sakit', dan 2 sel dibiarkan kosong. Jika guru menggunakan rumus `=COUNT(D2:D31)`, berapa hasil yang akan ditampilkan oleh Excel?",
      options: [
        "30 (menghitung seluruh baris tabel)",
        "28 (menghitung seluruh sel yang ada isinya termasuk teks)",
        "25 (karena fungsi COUNT hanya menghitung sel yang berisi data numerik/angka murni)",
        "0 (karena ada teks Sakit yang menyebabkan rumus error)"
      ],
      correctIndex: 2,
      explanation: "Fungsi =COUNT() hanya menghitung sel yang memuat data numerik/angka murni. Teks 'Sakit' dan sel kosong diabaikan, sehingga hasilnya adalah 25. Jika ingin menghitung semua sel terisi termasuk teks 'Sakit', gunakan fungsi =COUNTA() yang akan menghasilkan 28."
    },
    {
      id: 3,
      topicRef: "4.2",
      question: "Dalam perlombaan lari tingkat sekolah, guru ingin menampilkan catatan waktu dari peserta yang meraih juara ke-2 (peringkat kedua tercepat/terkecil). Fungsi Excel apa yang paling tepat digunakan?",
      options: [
        "=MAX(B2:B20)",
        "=LARGE(B2:B20; 2)",
        "=SMALL(B2:B20; 2)",
        "=AVERAGE(B2:B20)"
      ],
      correctIndex: 2,
      explanation: "Dalam perlombaan lari, juara ditentukan oleh waktu paling singkat/terkecil. Fungsi =SMALL(array; 2) mengambil nilai terkecil ke-2 (juara 2). Jika perlombaan berdasarkan skor nilai tertinggi, maka =LARGE(array; 2) yang digunakan."
    },
    {
      id: 4,
      topicRef: "4.3",
      question: "Mengapa fungsi COUNTIFS dan SUMIFS sangat direkomendasikan untuk laporan bisnis yang sering diurutkan (sorting) atau disaring (filtering)?",
      options: [
        "Karena fungsi ini mengunci seluruh baris tabel agar tidak bisa digeser pengguna lain",
        "Karena fungsi ini mengevaluasi nilai data faktual di dalam sel secara logis, bukan bergantung pada nomor urut baris, sehingga hasilnya tetap 100% akurat setelah data diurutkan",
        "Karena fungsi COUNTIFS dan SUMIFS otomatis menghapus baris transaksi yang bernilai minus",
        "Karena fungsi ini hanya bisa bekerja jika file disimpan dalam format CSV"
      ],
      correctIndex: 1,
      explanation: "COUNTIFS dan SUMIFS memeriksa kriteria nilai aktual di seluruh kolom (misal: mencari teks 'Jakarta' dan status 'Lunas'). Di mana pun baris tersebut berpindah setelah di-sort dari A ke Z, kriteria tetap terpenuhi dan kalkulasi menghasilkan angka yang sama persis."
    },
    {
      id: 5,
      topicRef: "4.4",
      question: "Data nama siswa dari sistem tertera sebagai teks campur aduk: 'mUHaMmaD aL-fATiH'. Fungsi apa yang harus digunakan untuk mengubah teks tersebut agar huruf pertama setiap kata menjadi huruf kapital rapi ('Muhammad Al-Fatih')?",
      options: [
        "=UPPER(\"mUHaMmaD aL-fATiH\")",
        "=LOWER(\"mUHaMmaD aL-fATiH\")",
        "=PROPER(\"mUHaMmaD aL-fATiH\")",
        "=TEXTJOIN(\"mUHaMmaD aL-fATiH\")"
      ],
      correctIndex: 2,
      explanation: "Fungsi =PROPER() mengonversi teks ke format Title Case, yaitu membuat huruf pertama di setiap kata menjadi huruf kapital dan seluruh huruf lainnya menjadi huruf kecil. Sangat ideal untuk standardisasi nama orang dan alamat."
    },
    {
      id: 6,
      topicRef: "4.6",
      question: "Saat sedang memeriksa keabsahan perhitungan raport siswa, seorang guru ingin melihat seluruh formula di lembar kerja secara serentak tanpa mengklik sel satu per satu. Shortcut keyboard tercepat apa yang dapat ditekan guru tersebut?",
      options: [
        "Ctrl + Shift + Enter",
        "Ctrl + ` (tombol grave accent / backtick)",
        "Alt + F4",
        "Ctrl + Alt + Delete"
      ],
      correctIndex: 1,
      explanation: "Shortcut Ctrl + ` (tombol backtick di samping angka 1) adalah pintasan resmi Excel untuk mengaktifkan fitur 'Show Formulas'. Seluruh angka hasil seketika berubah menampilkan rumus aslinya, memudahkan pemeriksaan integritas formula dan pencarian sel error."
    }
  ],
  sampleDatasets: [
    {
      id: "dataset-grades-formulas-ch4",
      name: "Data Raport & Rekap Nilai Siswa (IF, Statistik, & Audit)",
      filename: "latihan_raport_siswa_ch4.csv",
      description: "Dataset lengkap nilai ujian siswa untuk latihan fungsi statistik (MAX, MIN, AVERAGE, LARGE, COUNT, COUNTA), logika kelulusan IF, dan Show Formulas.",
      csvContent: `No,Nama Siswa,Kelas,Nilai UTS,Nilai UAS,Rata Rata,Status Kelulusan
1,Aditya Pratama,10-RPL,85,88,86.5,LULUS
2,Bagas Wicaksono,10-RPL,92,94,93.0,LULUS
3,Cantika Dewi,10-TKJ,70,72,71.0,REMEDIAL
4,Dimas Anggara,10-RPL,78,80,79.0,LULUS
5,Erlangga Putra,10-DKV,60,65,62.5,REMEDIAL
6,Fajar Ramadhan,10-TKJ,88,90,89.0,LULUS
7,Gita Gutawa,10-DKV,95,97,96.0,LULUS
8,Hendra Wijaya,10-RPL,72,74,73.0,REMEDIAL
9,Indah Permata,10-TKJ,82,85,83.5,LULUS
10,Jasmine Putri,10-DKV,65,68,66.5,REMEDIAL`
    },
    {
      id: "dataset-sales-commission-ch4",
      name: "Data Transaksi Penjualan & Komisi (Referensi Absolut & SUMIFS)",
      filename: "latihan_penjualan_komisi_ch4.csv",
      description: "Data penjualan multi-cabang untuk latihan penguncian sel komisi/pajak ($B$1) dan kalkulasi multi-kriteria COUNTIFS dan SUMIFS.",
      csvContent: `ID Transaksi,Sales,Cabang,Kategori,Status,Total Penjualan,Komisi 5%
TRX-101,Budi Santoso,Jakarta,Hardware,Lunas,5200000,260000
TRX-102,Dewi Lestari,Surabaya,Software,Lunas,1800000,90000
TRX-103,Ahmad Fauzi,Jakarta,Hardware,Pending,3400000,170000
TRX-104,Siti Aminah,Bandung,Aksesoris,Lunas,950000,47500
TRX-105,Budi Santoso,Surabaya,Hardware,Lunas,4200000,210000
TRX-106,Rudi Hartono,Jakarta,Software,Lunas,2500000,125000
TRX-107,Dewi Lestari,Bandung,Hardware,Pending,3100000,155000
TRX-108,Ahmad Fauzi,Surabaya,Aksesoris,Lunas,1400000,70000`
    },
    {
      id: "dataset-employee-text-audit-ch4",
      name: "Data Karyawan & Pelanggan Kotor (PROPER, TEXTJOIN, & Auditing)",
      filename: "latihan_text_cleaning_ch4.csv",
      description: "Data teks dengan huruf besar-kecil acak untuk latihan penggabungan teks (& / TEXTJOIN), standardisasi huruf (PROPER / UPPER), dan formula auditing.",
      csvContent: `ID,Nama Depan,Nama Belakang,Departemen,Kota Asal,Status Karyawan
KRY-01,aHMaD,kURniAWan,IT Support,bandung,Tetap
KRY-02,sITi,aMinAH,Finance,JAKARTA,Kontrak
KRY-03,bUDi,sANToso,Marketing,surabaya,Tetap
KRY-04,dInDa,pUTrI,Design,SEMARANG,Tetap
KRY-05,rIZkY,pRaTAmA,Logistik,yogyakarta,Kontrak
KRY-06,fITrI,hANdAYAni,HRD,jakarta,Tetap
KRY-07,gUnAWan,wIbOWO,IT Support,surabaya,Kontrak`
    }
  ]
};

window.CHAPTER_4_DATA = CHAPTER_4_DATA;
