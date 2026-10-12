/**
 * DATA & KURIKULUM PEMBELAJARAN ICT EXCEL - CHAPTER 2
 * Data Manipulation, Data Import, Sorting, and Conditional Formatting
 * Versi Khusus Siswa: Panduan Praktik Mandiri & Tutorial Langkah demi Langkah di Excel
 */

const CHAPTER_DATA = {
  chapterNumber: 2,
  chapterTitle: "Data Manipulation, Data Import, Sorting, and Conditional Formatting",
  chapterSubtitle: "Modul Praktik Mandiri & Tutorial Interaktif Microsoft Excel",
  overview: {
    title: "General Theory: Pengantar Manajemen & Manipulasi Data",
    description: "Dalam dunia kerja dan teknologi informasi (ICT), data mentah (raw data) yang baru diambil dari sistem jarang sekali siap langsung diolah. Data sering kali mengandung ketidakkonsistenan, duplikasi baris, format yang salah, atau pemisah kolom yang tidak teratur. Bab ini membekali kamu dengan keahlian penting untuk membersihkan, menstrukturkan, mengimpor, serta memvisualisasikan data secara otomatis menggunakan Microsoft Excel.",
    coreCompetencies: [
      "Mampu melakukan standardisasi dan pembersihan data massal secara cepat dan akurat",
      "Memahami mekanisme impor data eksternal teks (.csv/.txt) dan peran delimiter pemisah kolom",
      "Mengorganisasi data bertingkat (Multi-Level Sorting) dan Custom List secara tepat",
      "Mencegah kesalahan analisis dengan deteksi dan eliminasi baris duplikat",
      "Mengimplementasikan pemformatan bersyarat dinamis (Conditional Formatting) berbasis logika dan statistik",
      "Menguasai teknik pemindahan, transposisi, dan Paste Special untuk efisiensi pengolahan data"
    ]
  },
  topics: [
    {
      id: "topic-2-1",
      number: "2.1",
      title: "Find & Replace for Bulk Data",
      indonesianTitle: "Pencarian & Penggantian Data Massal",
      summary: "Find & Replace can quickly replace multiple occurrences of the same value using Replace All. This is useful for updating codes or IDs consistently across many rows without editing each cell manually.",
      shortcut: "Ctrl + H",
      ribbonPath: "Home → Editing → Find & Select → Replace",
      theoryExplanation: `Fitur **Find & Replace** merupakan instrumen penting dalam standardisasi data. Ketika mengelola ratusan hingga ribuan baris data, mengganti kode lama, prefiks ID produk, nama kota yang salah ketik (typo), atau tanda baca yang tidak konsisten secara manual akan memakan waktu berjam-jam dan sangat rawan kesalahan (*human error*).
      
Dengan perintah **Replace All**, Excel memindai seluruh *worksheet* atau area yang dipilih dalam hitungan milidetik dan mengganti seluruh kemunculan teks target secara serentak.`,
      keyPoints: [
        "**Replace All vs Replace**: Replace All mengganti seluruh kemunculan sekaligus, sedangkan Replace mengevaluasi satu per satu.",
        "**Opsi Match Case**: Membedakan huruf besar dan kecil (misal: 'JKT' tidak akan mengganti 'jkt').",
        "**Match Entire Cell Contents**: Sangat penting agar hanya sel yang bernilai persis sama yang diganti (mencegah teks di tengah kata ikut terganti).",
        "**Look In (Formulas vs Values)**: Menentukan apakah pencarian mencakup rumus di balik sel atau hanya teks yang tampil."
      ],
      proTips: "Selalu seleksi (blok) area kolom data yang ingin kamu ganti terlebih dahulu sebelum menekan **Replace All**. Jika kamu tidak memblok sel tertentu, Excel akan mencari dan mengganti teks yang cocok di seluruh lembar kerja, termasuk teks judul tabel atau rumus!",
      commonMistakes: "Lupa mengaktifkan opsi 'Match entire cell contents' saat ingin mengganti angka '1' menjadi '10', sehingga angka lain seperti '15' atau '21' ikut berubah salah menjadi '105' atau '210'.",
      exerciseTask: "Gantilah semua kode cabang 'CAB-OLD' menjadi 'CAB-NEW' pada tabel data aset menggunakan tombol Replace All.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "2.1_FindReplace",
        sampleData: `Kode Barang\tNama Barang\tLokasi
CAB-OLD-01\tLaptop Dell\tGudang A
CAB-OLD-02\tMonitor Samsung\tGudang B
CAB-01-09\tPrinter Canon\tGudang A
CAB-OLD-03\tKeyboard Logi\tGudang C
CAB-02-11\tMouse Wireless\tGudang B
CAB-OLD-04\tScanner Fujitsu\tGudang A`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:center;background-color:#107c41;color:#ffffff;">Kode Barang</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Nama Barang</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Lokasi</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">CAB-OLD-01</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Laptop Dell</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Gudang A</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">CAB-OLD-02</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Monitor Samsung</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Gudang B</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">CAB-01-09</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Printer Canon</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Gudang A</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">CAB-OLD-03</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Keyboard Logi</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Gudang C</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">CAB-02-11</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Mouse Wireless</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Gudang B</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">CAB-OLD-04</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Scanner Fujitsu</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Gudang A</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka aplikasi Microsoft Excel di komputer kamu. Klik tanda tambah (+) di bagian paling bawah untuk **membuat Sheet Baru**, lalu beri nama: `2.1_FindReplace` (klik kanan tab sheet > Rename).",
          "Klik tombol **'📋 Salin Tabel Latihan'** di bawah ini.",
          "Klik sel **A1** di sheet `2.1_FindReplace` **SATU KALI SAJA** (jangan klik ganda / jangan sampai kursor berkedip di dalam sel!), lalu tekan **Ctrl + V** agar format warna hijau dan border ikut tertempel sempurna.",
          "Sorot/blok kolom **A** (sel A2 sampai A7 yang berisi kode barang).",
          "Tekan tombol pintasan keyboard **Ctrl + H** (atau buka tab menu **Home > Editing > Find & Select > Replace**).",
          "Pada kotak **Find what:** ketikkan `CAB-OLD`.",
          "Pada kotak **Replace with:** ketikkan `CAB-NEW`.",
          "Klik tombol **Replace All**. Excel akan langsung menampilkan notifikasi bahwa semua kode berhasil diperbarui secara serentak!"
        ],
        expectedResult: "Seluruh kode yang berawalan CAB-OLD kini otomatis berubah rapi menjadi CAB-NEW-01, CAB-NEW-02, dst., dengan tampilan tabel bergaris dan berheader hijau profesional."
      }
    },
    {
      id: "topic-2-2",
      number: "2.2",
      title: "Importing External Data",
      indonesianTitle: "Mengimpor Data Eksternal (.txt & .csv)",
      summary: "Excel can import data from external files such as .txt or .csv through Data > Get Data. Excel uses delimiters such as tabs, commas, or semicolons to separate the data into columns.",
      shortcut: "Alt + A + P + T (atau via Ribbon Data)",
      ribbonPath: "Data → Get Data → From File → From Text/CSV",
      theoryExplanation: `Sistem informasi modern (seperti software kasir POS, database perbankan, absensi fingerprint, atau website) biasanya mengekspor laporan dalam format teks terstruktur sederhana seperti **CSV** (*Comma-Separated Values*) atau **TXT** (*Plain Text*).
      
Excel membaca pemisah kolom menggunakan karakter khusus yang disebut **Delimiter** (pemisah). Delimiter standar dunia adalah koma (,), namun di sistem Windows dengan regional setting Indonesia, delimiter default yang dipakai sering kali adalah titik koma (;).`,
      keyPoints: [
        "**Delimiter Populer**: Koma (,), Titik Koma (;), Tab (\\t), dan Garis Tegak (| / Pipe).",
        "**Regional Settings Issue**: Di Indonesia, koma dipakai sebagai penanda desimal (misal: 3,14), sehingga format CSV sering beralih menggunakan titik koma (;) sebagai pemisah kolom.",
        "**Power Query Engine**: Pada Excel modern, menu *From Text/CSV* menggunakan engine Power Query yang memberikan pratinjau langsung sebelum data dimuat (*Load*)."
      ],
      proTips: "Jika kamu membuka file CSV dengan cara klik ganda biasa dan semua teksnya menumpuk di Kolom A, jangan panik! Itu terjadi karena sistem Windows komputermu menggunakan delimiter berbeda. Solusinya, gunakan menu **Data > Get Data > From Text/CSV** agar kamu bisa memilih delimiter yang cocok secara manual.",
      commonMistakes: "Membuka file CSV langsung tanpa memeriksa delimiter pemisah, sehingga seluruh data alamat, nama, dan nilai bercampur jadi satu teks panjang di Kolom A.",
      exerciseTask: "Analisis teks mentah siswa dengan berbagai delimiter dan tentukan delimiter mana yang tepat agar data terbagi menjadi 4 kolom rapi.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "2.2_ImportCSV",
        sampleData: `ID,Nama_Siswa,Kelas,Nilai_ICT,Asal_Kota
001,Andi Pratama,10-RPL,88,Jakarta
002,Bunga Citra,10-TKJ,92,Bandung
003,Citra Dewi,10-RPL,79,Surabaya
004,Dimas Anggara,10-DKV,85,Yogyakarta
005,Evelyn Tan,10-TKJ,95,Semarang`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">ID</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Nama Siswa</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">Kelas</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Nilai ICT</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Asal Kota</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">001</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Andi Pratama</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">88</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Jakarta</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">002</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bunga Citra</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-TKJ</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">92</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bandung</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">003</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Citra Dewi</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">79</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Surabaya</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">004</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Dimas Anggara</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-DKV</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Yogyakarta</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">005</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Evelyn Tan</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-TKJ</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">95</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Semarang</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru dengan mengklik tanda (+) dan beri nama: `2.2_ImportCSV`.",
          "Buka program **Notepad** di komputer kamu. Salin contoh teks CSV di atas, tempelkan ke Notepad, lalu simpan file dengan nama `latihan_data.csv` (pilih *Save as type: All Files*).",
          "Kembali ke Excel di sheet `2.2_ImportCSV`, klik tab menu **Data** di bagian atas.",
          "Pada grup menu sebelah kiri (*Get & Transform Data*), klik ikon **From Text/CSV**.",
          "Pilih file `latihan_data.csv` yang baru saja kamu simpan di Notepad, lalu klik tombol **Import**.",
          "Pada jendela pratinjau Excel yang muncul, periksa bagian **Delimiter**: pastikan terpilih **Comma** (Koma).",
          "Perhatikan bahwa data sudah terbagi rapi menjadi 5 kolom terpisah, lalu klik tombol hijau **Load** di sudut kanan bawah."
        ],
        expectedResult: "Excel akan memuat data ke dalam sheet 2.2_ImportCSV sebagai tabel resmi dengan judul kolom ID, Nama_Siswa, Kelas, Nilai_ICT, dan Asal_Kota yang rapi dan terpisah sempurna."
      }
    },
    {
      id: "topic-2-3",
      number: "2.3",
      title: "Hyperlinks in Cells",
      indonesianTitle: "Tautan & Hyperlink Antar Dokumen",
      summary: "Excel cells can contain hyperlinks to email addresses, websites, files, or locations within a workbook. Use Insert > Link (Ctrl+K). The displayed text can be different from the actual link, making the worksheet cleaner and easier to read.",
      shortcut: "Ctrl + K",
      ribbonPath: "Insert → Link (atau klik kanan sel → Link)",
      theoryExplanation: `Hyperlink di Excel mengubah lembar kerja pasif menjadi **dashboard navigasi interaktif**. Sebuah sel dapat diarahkan ke 4 destinasi utama:
1. **Existing File or Web Page**: Membuka URL website (e.g., https://kemdikbud.go.id) atau dokumen eksternal di komputer.
2. **Place in This Document**: Melompat ke sheet lain atau sel tertentu di dalam workbook yang sama (cocok untuk daftar isi/menu utama).
3. **Create New Document**: Menautkan ke dokumen baru yang otomatis dibuat saat tautan diklik.
4. **E-mail Address**: Membuka aplikasi mail client dengan alamat penerima dan subjek email otomatis (protokol mailto:).

Konsep terpenting adalah memisahkan antara **Text to Display** (teks ramah yang dibaca pengguna) dengan **Address** (alamat URL/tujuan teknis).`,
      keyPoints: [
        "**Text to Display vs Address**: Alamat URL panjang tidak perlu mengotori tampilan; cukup ganti teksnya menjadi 'Klik di Sini' atau 'Lihat Detail'.",
        "**ScreenTip**: Tooltip kecil yang muncul saat kursor mouse diarahkan ke sel sebelum diklik.",
        "**Navigasi Sheet**: Digunakan untuk membuat tombol tombol 'Kembali ke Menu Utama' pada laporan keuangan bersheet banyak."
      ],
      proTips: "Gunakan tipe hyperlink *'Place in This Document'* untuk membuat sheet 'Menu Utama' atau 'Daftar Isi'. Buat kotak sel bertuliskan nama-nama sheet tugas kamu sehingga orang lain bisa berpindah sheet cukup dengan 1 kali klik layaknya aplikasi web!",
      commonMistakes: "Menempelkan alamat web panjang hingga 100 karakter langsung ke sel tanpa menggunakan fitur *Text to Display*, sehingga tabel menjadi sangat lebar dan tidak sedap dipandang.",
      exerciseTask: "Buatlah hyperlink interaktif dengan teks 'Hubungi Guru' yang mengarah ke alamat email dan 'Lihat Nilai' yang melompat ke Sheet 2.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "2.3_Hyperlinks",
        sampleData: `Menu\tTujuan\tKeterangan
Portal Sekolah\tWebsite Resmi\tAkan dibuka di browser
Kirim Email\tAlamat Pengajar\tBuka aplikasi email
Sheet Nilai\tSheet2!A1\tLompat ke Sheet 2`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Menu Navigasi</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Tujuan Tautan</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Keterangan</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#0563c1;text-decoration:underline;">Portal Sekolah</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Website Resmi</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Akan dibuka di browser</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#0563c1;text-decoration:underline;">Kirim Email</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Alamat Pengajar</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Buka aplikasi email</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;color:#0563c1;text-decoration:underline;">Sheet Nilai</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Sheet2!A1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Lompat ke Sheet 2</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru dengan mengklik tanda (+) dan beri nama: `2.3_Hyperlinks`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** di sheet `2.3_Hyperlinks` satu kali, lalu tekan **Ctrl + V**.",
          "Buat satu sheet lagi bernama `Sheet2` (sebagai tujuan lompatan link internal).",
          "Kembali ke sheet `2.3_Hyperlinks`, klik sel **A2** (yang bertuliskan 'Portal Sekolah').",
          "Tekan tombol pintasan **Ctrl + K** di keyboard (atau buka tab **Insert > Link**).",
          "Di sisi kiri dialog, pilih **Existing File or Web Page**. Pada kotak **Address:** ketik `https://kemdikbud.go.id`. Pada kotak **Text to display:** pastikan tertulis `Buka Portal Kemdikbud`, lalu klik **OK**.",
          "Sekarang klik sel **A4** (Sheet Nilai), tekan **Ctrl + K**. Di sisi kiri pilih **Place in This Document**, pilih **Sheet2**, lalu klik **OK**.",
          "Coba klik sel A4 tersebut: Excel akan langsung melompat ke Sheet2 secara instan!"
        ],
        expectedResult: "Teks sel berubah menjadi warna biru bergaris bawah. Mengklik sel A2 membuka website di browser, dan mengklik sel A4 berpindah ke lembar Sheet2."
      }
    },
    {
      id: "topic-2-4",
      number: "2.4",
      title: "Multi-Level Sorting",
      indonesianTitle: "Pengurutan Data Bertingkat (Multi-Level Sort)",
      summary: "The Sort feature can organize data using several columns through Add Level. The first level is the main criterion, followed by additional levels when values are the same. Each level can use ascending or descending order, while Custom List allows a specific order instead of alphabetical sorting.",
      shortcut: "Alt + A + S + S (Data → Sort)",
      ribbonPath: "Data → Sort → Add Level",
      theoryExplanation: `Pengurutan satu kolom (*Single Sort*) sering kali tidak cukup ketika data memiliki banyak entri yang bernilai identik. 
      
Fitur **Multi-Level Sort** memungkinkan kita menentukan prioritas hierarki pengurutan melalui tombol **Add Level**:
* **Level 1 (Sort by)**: Kriteria pengurutan utama.
* **Level 2 (Then by)**: Kriteria pengurutan kedua, yang hanya dieksekusi jika data pada Level 1 bernilai sama.
* **Level 3 (Then by)**: Kriteria pengurutan ketiga untuk memecah data seri di level kedua.

Selain pengurutan standar (A to Z / Z to A / Nilai Terkecil ke Terbesar), Excel mendukung **Custom List** untuk urutan non-abjad seperti jenjang pendidikan (*SD, SMP, SMA*), tingkat urgensi (*Rendah, Sedang, Tinggi*), atau nama hari/bulan.`,
      keyPoints: [
        "**Prioritas Hierarki**: Urutan level sangat menentukan hasil akhir. Mengurutkan 'Departemen lalu Gaji' berbeda dengan 'Gaji lalu Departemen'.",
        "**My data has headers**: Wajib dicentang agar baris judul kolom tidak ikut teracak ke dalam baris data.",
        "**Custom List**: Mengatur urutan logika khusus yang tidak mengikuti urutan alfabetis bahasa.",
        "**Expand the selection**: Peringatan penting Excel saat pengguna hanya menyorot 1 kolom agar kolom di sebelahnya tetap sinkron dan data tidak rusak."
      ],
      proTips: "Jangan pernah memilih opsi *'Continue with current selection'* jika Excel memunculkan dialog 'Sort Warning'! Selalu pilih *'Expand the selection'* agar kolom-kolom lain (seperti nama dan alamat) tetap menyatu dengan gajinya. Jika tidak, data antar-kolom akan tertukar dan rusak fatal!",
      commonMistakes: "Hanya memblok satu kolom saja saat menyortir, sehingga nilai pada kolom tersebut teracak sementara nama orang di kolom sebelahnya tidak ikut berpindah.",
      exerciseTask: "Urutkan tabel pegawai berdasarkan Departemen (A-Z), kemudian berdasarkan Gaji (Terbesar ke Terkecil / Descending).",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "2.4_SortBertingkat",
        sampleData: `ID\tNama Karyawan\tDepartemen\tGaji Pokok\tKota
EMP-01\tBudi Santoso\tMarketing\t6500000\tJakarta
EMP-02\tSiti Rahma\tKeuangan\t8000000\tBandung
EMP-03\tAhmad Fauzi\tIT\t9500000\tSurabaya
EMP-04\tDewi Lestari\tMarketing\t8200000\tJakarta
EMP-05\tRian Pratama\tIT\t6800000\tJakarta
EMP-06\tMaya Anggraini\tKeuangan\t6000000\tSurabaya
EMP-07\tFajar Hidayat\tIT\t11000000\tBandung
EMP-08\tNabila Putri\tMarketing\t5500000\tSurabaya`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">ID</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Nama Karyawan</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Departemen</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;background-color:#107c41;color:#ffffff;">Gaji Pokok</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Kota</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">EMP-01</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Budi Santoso</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Marketing</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">6500000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Jakarta</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">EMP-02</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Siti Rahma</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Keuangan</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">8000000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bandung</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">EMP-03</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Ahmad Fauzi</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">IT</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">9500000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Surabaya</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">EMP-04</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Dewi Lestari</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Marketing</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">8200000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Jakarta</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">EMP-05</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Rian Pratama</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">IT</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">6800000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Jakarta</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">EMP-06</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Maya Anggraini</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Keuangan</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">6000000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Surabaya</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">EMP-07</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Fajar Hidayat</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">IT</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">11000000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bandung</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">EMP-08</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Nabila Putri</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Marketing</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">5500000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Surabaya</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru dengan mengklik tanda (+) dan beri nama: `2.4_SortBertingkat`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** di sheet `2.4_SortBertingkat` satu kali, lalu tekan **Ctrl + V**.",
          "Klik salah satu sel mana saja di dalam tabel tersebut (misal sel C3).",
          "Buka tab menu **Data** di bagian atas, lalu klik tombol besar **Sort** (atau tekan **Alt + A + S + S**).",
          "Pastikan kotak centang **'My data has headers'** di sudut kanan atas dalam posisi **tercentang**.",
          "Pada baris pertama (**Sort by**), pilih kolom **Departemen**, dan atur Order ke **A to Z**.",
          "Klik tombol **Add Level** di bagian atas dialog untuk memunculkan baris kriteria kedua.",
          "Pada baris kedua (**Then by**), pilih kolom **Gaji Pokok**, dan ubah Order ke **Largest to Smallest** (Terbesar ke Terkecil).",
          "Klik tombol **OK**."
        ],
        expectedResult: "Tabel akan dikelompokkan berdasarkan Departemen secara abjad (IT, Keuangan, Marketing), dan di dalam masing-masing departemen gajinya langsung berurutan dari yang tertinggi ke terendah."
      }
    },
    {
      id: "topic-2-5",
      number: "2.5",
      title: "Removing Duplicate Data",
      indonesianTitle: "Menghapus Data Duplikat (Remove Duplicates)",
      summary: "Remove Duplicates removes repeated rows based on selected columns. The columns selected as the criteria determine which rows Excel considers duplicates, so they should be chosen carefully.",
      shortcut: "Alt + A + M",
      ribbonPath: "Data → Data Tools → Remove Duplicates",
      theoryExplanation: `Dalam basis data besar, duplikasi data (*redundancy*) sering terjadi akibat penginputan ganda oleh petugas, penggabungan file dari beberapa cabang, atau formulir pendaftaran daring yang dikirim berulang kali oleh responden yang sama.
      
Fitur **Remove Duplicates** memindai baris-baris data dan secara permanen menghapus baris kedua dan seterusnya yang memiliki nilai identik pada kolom kriteria yang dicentang.`,
      keyPoints: [
        "**Pemilihan Kolom Kriteria**: Jika semua kolom dicentang, baris baru dianggap duplikat jika SEMUA isinya persis sama. Jika hanya kolom 'NIS' atau 'Email' yang dicentang, maka jika ada 2 siswa dengan NIS sama, baris kedua akan langsung dihapus meskipun namanya berbeda.",
        "**Baris yang Dipertahankan**: Excel selalu mempertahankan baris data yang posisinya paling atas (pertama kali ditemukan) dan menghapus baris di bawahnya.",
        "**Pesan Konfirmasi**: Setelah eksekusi, dialog akan menginformasikan jumlah baris duplikat yang terhapus dan sisa baris data unik."
      ],
      proTips: "Selalu buat salinan cadangan lembar kerja (*backup sheet*) sebelum menjalankan **Remove Duplicates**, karena operasi ini menghapus baris data secara fisik dari lembar kerja dan terkadang sulit dibatalkan (Undo) jika file sudah terlanjur disimpan!",
      commonMistakes: "Mencentang kolom kriteria yang salah (misalnya hanya mencentang kolom 'Kota', akibatnya semua siswa yang berasal dari kota yang sama terhapus hingga hanya tersisa 1 orang saja per kota!).",
      exerciseTask: "Gunakan Remove Duplicates pada data pendaftaran siswa dengan kriteria kolom NISN agar setiap siswa hanya terdaftar satu kali.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "2.5_RemoveDuplicates",
        sampleData: `ID\tNISN\tNama Siswa\tKota Asal
S-101\t00511234\tAditya Pratama\tJakarta
S-102\t00511235\tBella Safitri\tBandung
S-103\t00511236\tCahyo Nugroho\tSurabaya
S-102\t00511235\tBella Safitri\tBandung
S-104\t00511237\tDinda Kirana\tJakarta
S-105\t00511238\tEko Purwanto\tSemarang
S-101\t00511234\tAditya Pratama\tJakarta`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">ID</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:center;background-color:#107c41;color:#ffffff;">NISN</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Nama Siswa</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Kota Asal</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">S-101</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">00511234</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Aditya Pratama</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Jakarta</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">S-102</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">00511235</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bella Safitri</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bandung</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">S-103</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">00511236</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Cahyo Nugroho</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Surabaya</td></tr>
    <tr style="background-color:#fee2e2;color:#991b1b;"><td style="border:1px solid #fca5a5;padding:5px 10px;text-align:center;">S-102</td><td style="border:1px solid #fca5a5;padding:5px 10px;text-align:center;font-weight:bold;">00511235</td><td style="border:1px solid #fca5a5;padding:5px 10px;">Bella Safitri (Duplikat)</td><td style="border:1px solid #fca5a5;padding:5px 10px;">Bandung</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">S-104</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">00511237</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Dinda Kirana</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Jakarta</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">S-105</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">00511238</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Eko Purwanto</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Semarang</td></tr>
    <tr style="background-color:#fee2e2;color:#991b1b;"><td style="border:1px solid #fca5a5;padding:5px 10px;text-align:center;">S-101</td><td style="border:1px solid #fca5a5;padding:5px 10px;text-align:center;font-weight:bold;">00511234</td><td style="border:1px solid #fca5a5;padding:5px 10px;">Aditya Pratama (Duplikat)</td><td style="border:1px solid #fca5a5;padding:5px 10px;">Jakarta</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru dengan mengklik tanda (+) dan beri nama: `2.5_RemoveDuplicates`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** di sheet `2.5_RemoveDuplicates` satu kali, lalu tekan **Ctrl + V**.",
          "*(Langkah Pengamanan)* Duplikasi sheet: Klik kanan tab sheet `2.5_RemoveDuplicates` di kiri bawah > pilih **Move or Copy...** > centang kotak **Create a copy** > klik **OK**.",
          "Kembali ke sheet utama `2.5_RemoveDuplicates`, klik salah satu sel di dalam tabel.",
          "Buka tab menu **Data** > klik ikon **Remove Duplicates** (di grup *Data Tools*) atau tekan pintasan **Alt + A + M**.",
          "Pastikan kotak **'My data has headers'** tercentang.",
          "Pilih kolom kriteria penentu: centang kolom **NISN** (karena NISN adalah identitas tunggal siswa).",
          "Klik tombol **OK**."
        ],
        expectedResult: "Excel memunculkan kotak pesan: '2 duplicate values found and removed; 5 unique values remain.' Dua baris kembar otomatis terhapus dan menyisakan data unik."
      }
    },
    {
      id: "topic-2-6",
      number: "2.6",
      title: "Rule-Based Conditional Formatting",
      indonesianTitle: "Format Bersyarat Berbasis Aturan (Conditional Formatting)",
      summary: "Conditional Formatting automatically changes cell formatting based on specific conditions. Unlike manual formatting, it updates when the data changes. For example, the Below Average rule can highlight values below the average.",
      shortcut: "Alt + H + L",
      ribbonPath: "Home → Styles → Conditional Formatting",
      theoryExplanation: `**Conditional Formatting** (Format Bersyarat) adalah fitur visualisasi data dinamis yang menerapkan gaya (warna latar, warna teks, border, ikon) pada sel hanya jika nilai sel tersebut memenuhi kriteria/kondisi tertentu.
      
Keunggulan utama dibandingkan format manual adalah sifatnya yang **Dinamis & Real-Time**: Ketika angka diubah, format warna akan langsung menyesuaikan sendiri seketika tanpa perlu disetel ulang.

Aturan populer di Excel mencakup:
* **Highlight Cells Rules**: Greater Than (>), Less Than (<), Between, Equal To, Text that Contains.
* **Top/Bottom Rules**: Top 10 Items, Top 10%, Above Average, **Below Average** (otomatis menghitung nilai rata-rata kelas lalu mewarnai mereka yang berada di bawahnya).
* **Data Bars & Color Scales**: Gradasi warna panas/dingin seperti heatmap.`,
      keyPoints: [
        "**Sifat Dinamis**: Format bukan properti statis, melainkan kalkulasi logika yang selalu aktif mengevaluasi sel.",
        "**Below Average**: Menghitung rata-rata secara internal di balik layar tanpa perlu membuat sel rumus =AVERAGE() terpisah.",
        "**Clear Rules**: Aturan dapat dihapus dari area terpilih (*Selected Cells*) atau dari seluruh sheet (*Entire Sheet*)."
      ],
      proTips: "Fitur 'Below Average' menghitung nilai rata-rata secara dinamis di balik layar tanpa kamu perlu membuat rumus `=AVERAGE()` manual di lembar kerja. Jika ada siswa yang nilainya diperbaiki setelah remedial, warna merah selnya akan otomatis hilang seketika!",
      commonMistakes: "Menumpuk terlalu banyak aturan format bersyarat tanpa menghapus aturan lama, sehingga warna sel menjadi tumpang tindih dan membingungkan pembaca laporan.",
      exerciseTask: "Terapkan aturan 'Below Average' pada kolom Nilai Ujian agar siswa yang membutuhkan remedial otomatis ditandai dengan warna merah muda.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "2.6_ConditionalFormat",
        sampleData: `No\tNama Siswa\tNilai Ujian
1\tAditya Pratama\t85
2\tBella Safitri\t58
3\tCahyo Nugroho\t92
4\tDinda Kirana\t72
5\tEko Purwanto\t52
6\tFani Maulida\t96
7\tGilang Ramadhan\t68
8\tHesti Wulandari\t86`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">No</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Nama Siswa</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;background-color:#107c41;color:#ffffff;">Nilai Ujian</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Aditya Pratama</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bella Safitri</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;background-color:#ffe4e6;color:#be123c;font-weight:bold;">58</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">3</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Cahyo Nugroho</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">92</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">4</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Dinda Kirana</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">72</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">5</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Eko Purwanto</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;background-color:#ffe4e6;color:#be123c;font-weight:bold;">52</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">6</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Fani Maulida</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">96</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">7</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Gilang Ramadhan</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;background-color:#ffe4e6;color:#be123c;font-weight:bold;">68</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">8</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Hesti Wulandari</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">86</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru dengan mengklik tanda (+) dan beri nama: `2.6_ConditionalFormat`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** di sheet `2.6_ConditionalFormat` satu kali, lalu tekan **Ctrl + V**.",
          "Sorot/blok area angka nilai ujian saja pada sel **C2 sampai C9**.",
          "Buka tab menu **Home**, lalu pada grup *Styles* klik **Conditional Formatting** (atau tekan **Alt + H + L**).",
          "Arahkan kursor mouse ke menu **Top/Bottom Rules**, lalu klik opsi **Below Average**.",
          "Pada kotak dialog yang muncul, biarkan pilihan format default yaitu **'Light Red Fill with Dark Red Text'** (Isian Merah Muda dengan Teks Merah Gelap).",
          "Klik tombol **OK**.",
          "*(Uji Kedinamisan)* Klik sel C3 (nilai Bella yang semula 58 dan berwarna merah). Ketik angka `95` lalu tekan Enter. Perhatikan bagaimana warna merah langsung hilang otomatis!"
        ],
        expectedResult: "Nilai siswa yang berada di bawah nilai rata-rata kelas (Bella: 58, Eko: 52, Gilang: 68) langsung disorot dengan warna merah muda secara otomatis."
      }
    },
    {
      id: "topic-2-7",
      number: "2.7",
      title: "Paste Special",
      indonesianTitle: "Tempel Khusus (Paste Special)",
      summary: "Paste Special (Ctrl+Alt+V) allows users to paste only specific parts of copied data, such as Values, Formulas, or Formats. It is useful for converting formulas into fixed values or applying formatting without changing existing data.",
      shortcut: "Ctrl + Alt + V (atau klik kanan → Paste Special)",
      ribbonPath: "Home → Clipboard → Paste dropdown → Paste Special",
      theoryExplanation: `Saat melakukan Copy (Ctrl+C) dan Paste biasa (Ctrl+V), Excel menyalin segala atribut sel sekaligus: isi, rumus, warna, jenis font, garis tepi, hingga validasi data.
      
**Paste Special** memberikan kendali penuh untuk menempelkan elemen spesifik yang diinginkan saja:
1. **Values (Nilai)**: Hanya menempelkan hasil akhir kalkulasi tanpa membawa rumusnya (mengubah formula dinamis menjadi angka statis permanen).
2. **Formulas (Rumus)**: Hanya menempelkan formula tanpa menyalin format warna atau garis.
3. **Formats (Format)**: Hanya meniru gaya visual (seperti alat Format Painter) tanpa mengubah isi data sel tujuan.
4. **Column Widths**: Menyamakan lebar kolom sel tujuan agar sama persis dengan kolom asal.
5. **Operations (Add, Subtract, Multiply, Divide)**: Melakukan operasi aritmatika massal secara langsung ke sel tujuan.`,
      keyPoints: [
        "**Membekukan Rumus (Freeze Formulas)**: Mencegah error #REF! saat data hasil rumus ingin dipindahkan ke file lain atau dibagikan ke pihak ketiga tanpa mengekspos rumus rahasia.",
        "**Operasi Matematika Instan**: Misalnya menambah bonus 10% ke seluruh sel gaji hanya dengan Paste Special > Multiply.",
        "**Hemat Bandwidth & Performa**: Menghilangkan beban kalkulasi berulang pada dataset puluhan ribu baris."
      ],
      proTips: "Selalu gunakan **Paste Special > Values** sebelum mengirim laporan akhir ke orang lain atau memindahkannya ke sheet baru. Ini akan membekukan rumus menjadi angka permanen dan mencegah munculnya pesan error `#REF!` yang sering terjadi jika sel referensi terhapus atau bergeser.",
      commonMistakes: "Menggunakan Ctrl+V biasa saat menyalin hasil rumus ke kolom lain, sehingga referensi sel bergeser dan hasilnya berubah menjadi 0 atau memunculkan error #VALUE!.",
      exerciseTask: "Salin kolom Total Harga yang berisi rumus =Jumlah*Harga, lalu gunakan Paste Special > Values untuk mengubahnya menjadi angka permanen.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "2.7_PasteSpecial",
        sampleData: `Item\tQty\tHarga\tTotal Rumus
Flashdisk 32GB\t5\t60000\t=B2*C2
SSD 512GB\t2\t250000\t=B3*C3
Mouse Wireless\t10\t85000\t=B4*C4
Keyboard Gaming\t4\t350000\t=B5*C5`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Item</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Qty</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;background-color:#107c41;color:#ffffff;">Harga</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;background-color:#0284c7;color:#ffffff;">Total (Rumus)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Flashdisk 32GB</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">5</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">60000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;background-color:#e0f2fe;font-weight:bold;">=B2*C2</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">SSD 512GB</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">250000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;background-color:#e0f2fe;font-weight:bold;">=B3*C3</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Mouse Wireless</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">10</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;background-color:#e0f2fe;font-weight:bold;">=B4*C4</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">Keyboard Gaming</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">4</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">350000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;background-color:#e0f2fe;font-weight:bold;">=B5*C5</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru dengan mengklik tanda (+) dan beri nama: `2.7_PasteSpecial`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** di sheet `2.7_PasteSpecial` satu kali, lalu tekan **Ctrl + V**. Pastikan di sel D2 tertulis rumus `=B2*C2`, di D3 `=B3*C3`, dst.",
          "Sorot sel **D2 sampai D5**, lalu tekan **Ctrl + C** (garis putus-putus berputar akan muncul mengelilingi sel).",
          "Klik sel target baru di sebelahnya, misalnya klik sel **F2**.",
          "Tekan tombol pintasan **Ctrl + Alt + V** di keyboard (atau klik kanan sel F2 > pilih ikon bertuliskan angka **123 / Values**).",
          "Pada kotak dialog Paste Special, pilih bulatan opsi **Values**, lalu klik **OK**.",
          "Klik salah satu sel di kolom F (misal F2) dan lihat kotak **Formula Bar (fx)** di bagian atas Excel."
        ],
        expectedResult: "Di sel F2 tidak ada lagi rumus =B2*C2, melainkan angka statis murni '300000'. Rumus telah berhasil dibekukan menjadi nilai tetap yang aman dari error #REF!."
      }
    },
    {
      id: "topic-2-8",
      number: "2.8",
      title: "Inserting, Moving, and Transposing Data",
      indonesianTitle: "Menyisipkan, Memindahkan, & Transposisi Data",
      summary: "Rows and columns can be inserted or deleted using Insert or Delete. Data can be moved with Cut and Insert Cut Cells. The Transpose option in Paste Special changes rows into columns or columns into rows.",
      shortcut: "Ctrl + Shift + + (Insert) / Ctrl + - (Delete)",
      ribbonPath: "Home → Cells → Insert / Delete; Paste Special → Transpose",
      theoryExplanation: `Pengorganisasian tata letak (*layout*) tabel merupakan keterampilan dasar penataan lembar kerja:
* **Insert / Delete**: Menambahkan baris kosong di atas baris aktif atau kolom di sebelah kiri kolom aktif.
* **Moving Data (Cut & Insert Cut Cells)**: Memotong data (Ctrl+X) lalu memilih opsi *Insert Cut Cells* memindahkan baris/kolom secara aman tanpa menimpa (*overwrite*) data yang sudah ada sebelumnya.
* **Transposing (Transposisi)**: Memutar orientasi tabel 90 derajat—mengubah baris horizontal menjadi kolom vertikal, atau sebaliknya. Transpose dapat dilakukan melalui opsi *Transpose* pada dialog Paste Special atau fungsi =TRANSPOSE().`,
      keyPoints: [
        "**Insert Cut Cells vs Paste**: Paste biasa setelah Cut akan menimpa sel tujuan, sedangkan Insert Cut Cells menggeser sel tetangga secara aman.",
        "**Transpose**: Sangat berguna ketika format data yang diunduh berbentuk horizontal memanjang ke samping, padahal standar analisis database menghendaki format vertikal ke bawah.",
        "**Struktur Orientasi**: Memudahkan persiapan tabel sebelum dibuatkan Pivot Table atau Grafik."
      ],
      proTips: "Saat ingin memindahkan baris data di antara baris lain, JANGAN gunakan Cut lalu Paste biasa karena data lama di sel tujuan akan langsung terhapus tertimpa! Selalu gunakan klik kanan pada baris tujuan lalu pilih **'Insert Cut Cells'** agar Excel menggeser data yang ada ke bawah secara aman.",
      commonMistakes: "Melakukan Cut lalu Paste langsung di atas sel yang sudah ada isinya, sehingga data yang sudah diketik sebelumnya terhapus permanen.",
      exerciseTask: "Putar tabel data penjualan triwulan dari format baris horizontal menjadi format kolom vertikal menggunakan fitur Transpose.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "2.8_Transpose",
        sampleData: `Produk\tQ1 (Jan-Mar)\tQ2 (Apr-Jun)\tQ3 (Jul-Sep)\tQ4 (Okt-Des)
Laptop ThinkPad\t120 unit\t150 unit\t180 unit\t210 unit
Monitor Curved\t85 unit\t90 unit\t110 unit\t140 unit
Printer Laser\t45 unit\t60 unit\t55 unit\t70 unit`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Produk</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Q1 (Jan-Mar)</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Q2 (Apr-Jun)</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Q3 (Jul-Sep)</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Q4 (Okt-Des)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;font-weight:bold;">Laptop ThinkPad</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">120 unit</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">150 unit</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">180 unit</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">210 unit</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;font-weight:bold;">Monitor Curved</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85 unit</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">90 unit</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">110 unit</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">140 unit</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;font-weight:bold;">Printer Laser</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">45 unit</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">60 unit</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">55 unit</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">70 unit</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru dengan mengklik tanda (+) dan beri nama: `2.8_Transpose`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** di sheet `2.8_Transpose` satu kali, lalu tekan **Ctrl + V**.",
          "Sorot seluruh tabel **A1:E4**, lalu tekan **Ctrl + C** untuk menyalin data.",
          "Klik sel kosong di bawahnya, misalnya sel **A7**.",
          "Tekan tombol pintasan **Ctrl + Alt + V** (atau klik kanan sel A7 > klik tanda panah pada Paste > pilih **Paste Special**).",
          "Di sudut kanan bawah kotak dialog Paste Special, centang kotak kecil **Transpose**.",
          "Klik tombol **OK**.",
          "*(Praktik Insert Cut Cells)*: Sorot baris 3 (Monitor), tekan **Ctrl + X** (Cut). Klik kanan baris 2 (Laptop), lalu klik **Insert Cut Cells**. Baris Monitor berpindah aman ke atas tanpa menimpa baris Laptop!"
        ],
        expectedResult: "Tabel horizontal berputar 90 derajat: Periode Q1-Q4 yang semula membujur di baris atas kini berubah menjadi baris vertikal ke bawah, dan nama produk berubah menjadi judul kolom."
      }
    },
    {
      id: "topic-2-9",
      number: "2.9",
      title: "Fill Handle and Automatic Series",
      indonesianTitle: "Fill Handle & Deret Otomatis (Series)",
      summary: "The Fill Handle is the small square at the bottom right corner of a selected cell. Dragging it allows Excel to continue patterns such as numbers, dates, or text sequences. For more control, use Fill > Series to set the starting value, interval, and ending value.",
      shortcut: "Double Click Fill Handle / Alt + H + F + I + S (Fill Series)",
      ribbonPath: "Home → Editing → Fill → Series",
      theoryExplanation: `**Fill Handle** adalah kotak hijau kecil di sudut kanan bawah sel yang sedang aktif. Fitur ini merupakan salah satu penghemat waktu (*time-saver*) paling ampuh di Microsoft Excel.

Excel memiliki kecerdasan pola (*Pattern Recognition*):
* **Pola Angka**: Jika mengetik 1 lalu 2, drag fill handle akan melanjutkan 3, 4, 5...
* **Pola Penanggalan**: Tanggal bertambah hari, bulan, atau tahun otomatis.
* **Pola Teks Bersiklus**: Nama hari (*Senin, Selasa...*) atau bulan (*Januari, Februari...*).
* **Teks Berkombinasi Angka**: Teks 'Siswa-01' otomatis dilanjutkan menjadi 'Siswa-02', 'Siswa-03'.

Untuk kontrol presisi pembuatan deret ribuan baris tanpa perlu mendrag mouse jauh ke bawah, gunakan dialog **Fill > Series**:
* **Series in**: Rows (ke samping) atau Columns (ke bawah).
* **Type**: Linear (tambah tetap), Growth (kali lipat), Date (tanggal).
* **Step Value**: Nilai kenaikan per langkah (interval).
* **Stop Value**: Batas nilai akhir di mana deret harus berhenti.`,
      keyPoints: [
        "**Trik Double Click**: Mengklik dua kali sudut kanan bawah sel akan otomatis mengisi deret ke bawah sejauh kolom tetangga di sebelah kirinya memiliki data.",
        "**Auto Fill Options Tag**: Ikon kecil yang muncul setelah drag memungkinkan memilih opsi 'Copy Cells', 'Fill Series', 'Fill Formatting Only', atau 'Fill Without Formatting'.",
        "**Fill Series Dialog**: Solusi instan membuat deret angka 1 sampai 10.000 dalam 1 detik tanpa melelahkan tangan."
      ],
      proTips: "Untuk membuat nomor urut 1 sampai 1.000 secara instan tanpa perlu menarik mouse jauh ke bawah, jangan di-drag manual! Cukup ketik angka 1 di sel pertama, buka **Home > Editing > Fill > Series**, pilih *Columns*, isi *Stop value* dengan 1000, lalu tekan Enter. Dalam 1 detik, seribu baris langsung terisi penuh!",
      commonMistakes: "Hanya mengetik angka '1' di satu sel lalu menarik ke bawah tanpa menekan tombol Ctrl atau memilih 'Fill Series', sehingga yang muncul adalah angka '1' semua.",
      exerciseTask: "Buat deret kelipatan 5 dimulai dari angka 5 hingga batas 100 secara otomatis menggunakan menu Fill Series.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "2.9_FillSeries",
        sampleData: `Nomor\tNama Hari\tKode Siswa
1\tSenin\tSISWA-01
2\tSelasa\tSISWA-02`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">Nomor</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Nama Hari</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:center;background-color:#107c41;color:#ffffff;">Kode Siswa</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Senin</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">SISWA-01</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Selasa</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">SISWA-02</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru dengan mengklik tanda (+) dan beri nama: `2.9_FillSeries`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** di sheet `2.9_FillSeries` satu kali, lalu tekan **Ctrl + V**.",
          "Sorot sel angka **A2:A3** sekaligus (keduanya harus diblok agar Excel mengenali interval kenaikannya).",
          "Arahkan kursor mouse tepat ke **sudut kanan bawah sel A3** sampai simbol kursor berubah menjadi tanda tambah hitam pekat (`+`) yang disebut **Fill Handle**.",
          "Klik dan tahan mouse, lalu tarik (drag) ke bawah hingga baris **A10**. Angka 3, 4, 5... hingga 10 akan otomatis tercipta!",
          "*(Mencoba Pola Teks)*: Di sel B2 dan B3 (Senin, Selasa), sorot keduanya lalu tarik fill handle ke bawah: Excel otomatis mengisikan Rabu, Kamis, Jumat, dst.",
          "*(Trik Kilat Fill Series)*: Di sel D1 ketik angka `5`. Buka menu **Home > Editing > Fill > Series**. Pada *Series in* pilih **Columns**, *Step value:* ketik `5`, *Stop value:* ketik `100`, lalu klik **OK**."
        ],
        expectedResult: "Deret kelipatan 5 (5, 10, 15, 20... hingga 100) langsung terbentuk secara instan dan rapi di kolom D."
      }
    },
    {
      id: "topic-2-10",
      number: "2.10",
      title: "Currency Format, Cell Styles, and Text Orientation",
      indonesianTitle: "Format Mata Uang, Cell Styles, & Orientasi Teks",
      summary: "Currency Format displays numbers with a currency symbol and selected decimal places without changing the actual value. Cell Styles provide ready-made combinations of fonts, colors, and borders for consistent formatting. Text Orientation allows text to be rotated, which can save space when column headings are long.",
      shortcut: "Ctrl + Shift + $ (Currency) / Format Cells dialog (Ctrl + 1)",
      ribbonPath: "Home → Number (Currency) / Home → Styles (Cell Styles) / Home → Alignment (Orientation)",
      theoryExplanation: `Aspek estetika, standardisasi, dan keterbacaan (*readability*) adalah pilar profesionalitas lembar kerja:

1. **Currency Format**: 
Menampilkan simbol mata uang (seperti **Rp** atau **$**) serta jumlah desimal **tanpa mengubah nilai numerik asli di balik sel**. Nilai sel tetap berupa angka murni sehingga tetap dapat dihitung oleh rumus matematis (=SUM, =AVERAGE). 

2. **Cell Styles**:
Kumpulan kombinasi siap pakai yang memadukan jenis font, ukuran, warna latar, aksen, dan garis tepi. Membantu menjaga konsistensi gaya visual (seperti gaya Heading, Total, Normal, Good/Bad/Neutral).

3. **Text Orientation (Rotasi Teks)**:
Memutar sudut teks (misal 45 derajat berlawanan arah jarum jam, atau 90 derajat vertikal). Sangat berguna ketika judul kolom memiliki teks panjang sedangkan isi datanya hanya angka kecil atau tanda centang singkat, sehingga lebar kolom dapat dihemat secara drastis tanpa terpotong.`,
      keyPoints: [
        "**Aturan Emas Input Angka**: JANGAN PERNAH mengetik huruf 'Rp' secara manual di keyboard saat menginput harga (misal mengetik 'Rp 50.000') karena Excel akan memperlakukannya sebagai teks biasa dan rumus matematika akan error #VALUE!.",
        "**Perbedaan Currency vs Accounting**: Format Accounting menempatkan simbol mata uang rata kiri rapat dan angka rata kanan, sedangkan Currency menempatkan simbol mata uang menempel tepat di depan angka.",
        "**Text Orientation**: Menghemat ruang horizontal pada tabel kuisioner atau daftar hadir siswa dengan puluhan tanggal."
      ],
      proTips: "ATURAN EMAS: Jangan pernah mengetik huruf 'Rp' atau titik pemisah ribuan secara manual di keyboard saat memasukkan nominal uang! Ketikkan angkanya saja secara murni (misal: 350000), lalu atur tampilannya melalui menu **Home > Number > Currency**. Dengan begitu, rumus penjumlahan `=SUM()` akan tetap berfungsi sempurna!",
      commonMistakes: "Mengetik huruf 'Rp' manual dari keyboard sehingga sel berubah tipe menjadi teks (String), akibatnya ketika dihitung dengan rumus =SUM() hasilnya menjadi 0 atau memunculkan error #VALUE!.",
      exerciseTask: "Formatlah angka menjadi mata uang Rupiah (Rp) tanpa desimal, terapkan Cell Style 'Accent 1', dan putar judul kolom sebesar 45 derajat.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "2.10_CurrencyStyles",
        sampleData: `Komponen Biaya\tNominal Tarif Pembayaran\tStatus Verifikasi
SPP Bulanan\t350000\tLunas
Uang Gedung Sekolah\t2500000\tBelum
Seragam dan Buku Paket\t750000\tLunas
Praktikum Komputer\t150000\tLunas`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Komponen Biaya</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;background-color:#107c41;color:#ffffff;">Nominal Tarif Pembayaran</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:center;background-color:#107c41;color:#ffffff;">Status Verifikasi</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;font-weight:bold;">SPP Bulanan</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">Rp 350.000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;background-color:#dcfce7;color:#166534;font-weight:bold;">Lunas</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;font-weight:bold;">Uang Gedung Sekolah</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">Rp 2.500.000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;background-color:#fee2e2;color:#991b1b;font-weight:bold;">Belum</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;font-weight:bold;">Seragam dan Buku Paket</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">Rp 750.000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;background-color:#dcfce7;color:#166534;font-weight:bold;">Lunas</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;font-weight:bold;">Praktikum Komputer</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;font-weight:bold;">Rp 150.000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;background-color:#dcfce7;color:#166534;font-weight:bold;">Lunas</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru dengan mengklik tanda (+) dan beri nama: `2.10_CurrencyStyles`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** di sheet `2.10_CurrencyStyles` satu kali, lalu tekan **Ctrl + V**.",
          "*(Format Mata Uang)*: Sorot angka nominal di sel **B2 sampai B5**. Buka tab menu **Home**, pada bagian *Number* klik kotak dropdown format (yang bertuliskan General) lalu pilih **Currency** (atau tekan pintasan **Ctrl + Shift + $**). Angka akan berawalan Rp dengan pemisah ribuan yang rapi!",
          "*(Uji Rumus)*: Di sel B6 ketik rumus `=SUM(B2:B5)` lalu tekan Enter. Hasil total langsung terjumlah otomatis karena nominal diformat secara benar sebagai angka murni.",
          "*(Cell Styles)*: Sorot baris judul tabel **A1:C1**. Buka tab **Home > Styles > Cell Styles**, lalu klik gaya **Accent 1** (atau gaya Heading).",
          "*(Text Orientation)*: Tetap blok sel A1:C1, buka tab **Home > Alignment**, klik tombol ikon **Orientation (simbol ab dengan panah miring)**, lalu pilih **Angle Counterclockwise** (Putar 45°)."
        ],
        expectedResult: "Tabel terlihat sangat profesional: judul kolom miring 45 derajat menghemat ruang lembar kerja, dan semua angka berformat Rupiah resmi serta dapat dijumlahkan dengan rumus matematis."
      }
    }
  ],
  quiz: [
    {
      id: 1,
      topicRef: "2.1",
      question: "Apa keuntungan utama menggunakan fitur 'Replace All' pada menu Find & Replace?",
      options: [
        "Menghapus seluruh rumus di dalam lembar kerja sekaligus",
        "Mengganti seluruh kemunculan teks atau kode secara massal dan konsisten dalam hitungan detik",
        "Mengubah jenis font semua sel menjadi Arial",
        "Membuat backup otomatis sebelum data diedit"
      ],
      correctIndex: 1,
      explanation: "Replace All memindai seluruh data dan mengganti semua nilai yang cocok secara serempak, memastikan konsistensi ID atau kode tanpa perlu mengedit ribuan sel satu demi satu."
    },
    {
      id: 2,
      topicRef: "2.2",
      question: "Ketika mengimpor file eksternal .csv atau .txt ke dalam Excel, fungsi dari 'Delimiter' adalah...",
      options: [
        "Menentukan jenis mata uang yang digunakan pada tabel",
        "Sebagai karakter pemisah (seperti koma, tab, atau titik koma) yang membagi data teks menjadi kolom-kolom terpisah",
        "Mengunci file agar tidak bisa diedit oleh orang lain",
        "Menghitung rata-rata nilai secara otomatis saat proses impor berlangsung"
      ],
      correctIndex: 1,
      explanation: "Delimiter adalah tanda baca atau karakter khusus (seperti koma, tab, semicolon) yang digunakan sistem untuk mengenali batas pemisah antar kolom data."
    },
    {
      id: 3,
      topicRef: "2.3",
      question: "Pintasan keyboard (shortcut) standar untuk menyisipkan Hyperlink pada sel Excel adalah...",
      options: [
        "Ctrl + H",
        "Ctrl + L",
        "Ctrl + K",
        "Ctrl + Alt + V"
      ],
      correctIndex: 2,
      explanation: "Ctrl + K adalah shortcut universal di Excel (dan aplikasi Office lainnya) untuk membuka dialog Insert Hyperlink."
    },
    {
      id: 4,
      topicRef: "2.4",
      question: "Pada fitur Multi-Level Sort, kapan kriteria pengurutan pada Level kedua (Then by) akan dieksekusi oleh Excel?",
      options: [
        "Hanya ketika nilai data pada Level pertama bernilai sama (seri/identik)",
        "Selalu dieksekusi terlebih dahulu sebelum Level pertama",
        "Hanya jika kolom Level pertama berisi data tanggal",
        "Kapan saja secara acak"
      ],
      correctIndex: 0,
      explanation: "Level kedua (Then by) berperan sebagai kriteria pemecah seri; ia hanya bekerja pada baris-baris data yang memiliki nilai kembar pada Level pertama."
    },
    {
      id: 5,
      topicRef: "2.5",
      question: "Jika kamu ingin memastikan tidak ada siswa yang terdaftar dua kali berdasarkan nomor induk siswa, kolom kriteria apa yang harus dicentang pada dialog 'Remove Duplicates'?",
      options: [
        "Semua kolom harus dicentang tanpa terkecuali",
        "Hanya kolom Nomor Induk Siswa (NIS/NISN)",
        "Hanya kolom Nama Siswa",
        "Kolom Tanggal Lahir dan Alamat Rumah"
      ],
      correctIndex: 1,
      explanation: "Mencentang kolom NIS/NISN sebagai kriteria utama memastikan bahwa baris dianggap duplikat jika nomor identitasnya sama, meskipun terjadi variasi penulisan pada nama atau kolom lain."
    },
    {
      id: 6,
      topicRef: "2.6",
      question: "Mengapa Conditional Formatting disebut 'dinamis' dibandingkan dengan pemformatan manual biasa?",
      options: [
        "Karena file Excel akan bersuara saat kondisi terpenuhi",
        "Karena warna dan format sel akan otomatis berubah atau hilang seketika saat nilai di dalam sel berubah",
        "Karena sel tersebut tidak bisa diubah lagi setelah diformat",
        "Karena hanya bisa digunakan pada koneksi internet aktif"
      ],
      correctIndex: 1,
      explanation: "Format bersyarat bersifat dinamis karena berbasis evaluasi logika berkelanjutan; jika angka sel diperbaiki melampaui batas kriteria, format warna akan otomatis menyesuaikan tanpa intervensi manual."
    },
    {
      id: 7,
      topicRef: "2.7",
      question: "Kamu ingin menyalin hasil akhir perhitungan rumus ke dokumen lain tanpa membawa rumusnya agar tidak muncul pesan error #REF!. Pilihan Paste Special apa yang tepat?",
      options: [
        "Formulas",
        "Formats",
        "Values",
        "Validation"
      ],
      correctIndex: 2,
      explanation: "Paste Special > Values membekukan hasil rumus menjadi teks atau angka statis permanen tanpa mengikutsertakan formula perhitungannya."
    },
    {
      id: 8,
      topicRef: "2.8",
      question: "Opsi 'Transpose' pada Microsoft Excel berfungsi untuk...",
      options: [
        "Menghapus baris kosong yang ada di antara data",
        "Memutar orientasi data dengan mengubah baris menjadi kolom atau kolom menjadi baris",
        "Menerjemahkan teks bahasa Inggris ke bahasa Indonesia",
        "Mengunci posisi baris judul agar tidak ikut tergulung (scroll)"
      ],
      correctIndex: 1,
      explanation: "Transpose menukar dimensi baris dan kolom: data yang semula membujur horizontal di baris akan diputar menjadi membujur vertikal di kolom, dan sebaliknya."
    },
    {
      id: 9,
      topicRef: "2.9",
      question: "Di mana letak posisi 'Fill Handle' pada sel yang sedang aktif di Excel?",
      options: [
        "Kotak kecil di sudut kiri atas sel",
        "Kotak kecil berwarna hijau di sudut kanan bawah sel",
        "Di tengah-tengah formula bar fx",
        "Di sebelah tab ribbon Insert"
      ],
      correctIndex: 1,
      explanation: "Fill Handle terletak di sudut kanan bawah sel aktif, ditandai dengan titik kotak kecil yang jika kursor diarahkan ke atasnya akan berubah menjadi tanda plus (+) hitam pekat."
    },
    {
      id: 10,
      topicRef: "2.10",
      question: "Mengapa mengetik huruf 'Rp' secara manual menggunakan keyboard pada sel angka harga (misal: 'Rp 50000') adalah sebuah kesalahan fatal dalam Excel?",
      options: [
        "Karena Excel akan menghapus file tersebut secara otomatis",
        "Karena Excel akan menganggap isi sel tersebut sebagai teks biasa, sehingga tidak dapat dihitung oleh rumus matematika seperti =SUM()",
        "Karena font sel akan berubah menjadi miring secara paksa",
        "Karena teks 'Rp' hanya diakui jika menggunakan huruf kecil semua"
      ],
      correctIndex: 1,
      explanation: "Mengetik karakter huruf bersamaan dengan angka membuat Excel mengklasifikasikan sel tersebut sebagai tipe Teks (String), sehingga operasi aritmatika dan fungsi matematika seperti =SUM() akan menghasilkan angka 0 atau error #VALUE!."
    }
  ],
  sampleDatasets: [
    {
      id: "dataset-employee",
      name: "Data Karyawan & Gaji (Multi-Level Sort & Duplikat)",
      filename: "latihan_karyawan_chapter2.csv",
      description: "Data tabel dengan duplikasi nama dan kombinasi departemen untuk latihan Sort bertingkat dan Remove Duplicates.",
      csvContent: `ID,Nama Karyawan,Departemen,Gaji Pokok,Kota
EMP-001,Budi Santoso,Marketing,6500000,Jakarta
EMP-002,Siti Rahma,Keuangan,8000000,Bandung
EMP-003,Ahmad Fauzi,IT,9500000,Surabaya
EMP-002,Siti Rahma,Keuangan,8000000,Bandung
EMP-004,Dewi Lestari,Marketing,5500000,Jakarta
EMP-005,Rian Pratama,IT,7200000,Jakarta
EMP-006,Maya Anggraini,Keuangan,6800000,Surabaya
EMP-004,Dewi Lestari,Marketing,5500000,Jakarta`
    },
    {
      id: "dataset-grades",
      name: "Data Nilai Siswa (Conditional Formatting & Below Average)",
      filename: "latihan_nilai_siswa.csv",
      description: "Data nilai siswa untuk menguji aturan Conditional Formatting (Below Average, Top 3, KKM).",
      csvContent: `No,Nama Siswa,Tugas,UTS,UAS,Rata-Rata
1,Aditya Pratama,85,78,82,81.7
2,Bella Safitri,60,65,58,61.0
3,Cahyo Nugroho,90,92,88,90.0
4,Dinda Kirana,72,70,74,72.0
5,Eko Purwanto,55,50,62,55.7
6,Fani Maulida,95,96,92,94.3
7,Gilang Ramadhan,68,64,70,67.3
8,Hesti Wulandari,88,85,86,86.3`
    },
    {
      id: "dataset-sales",
      name: "Data Penjualan Triwulan (Transpose & Paste Special)",
      filename: "latihan_penjualan_triwulan.csv",
      description: "Data format horizontal yang ideal untuk latihan Transpose ke vertikal dan Paste Values.",
      csvContent: `Produk,Januari,Februari,Maret,Total
Laptop Asus,12,15,18,45
Printer Epson,8,10,7,25
Monitor LG,14,20,16,50
Keyboard Logitech,25,30,28,83`
    }
  ]
};
