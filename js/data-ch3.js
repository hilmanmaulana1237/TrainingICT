/**
 * DATA & KURIKULUM PEMBELAJARAN ICT EXCEL - CHAPTER 3
 * Tables, Table Styles, Subtotals, and Data Formatting
 * Versi Khusus Siswa: Panduan Praktik Mandiri & Tutorial Langkah demi Langkah di Excel
 */

const CHAPTER_3_DATA = {
  chapterNumber: 3,
  chapterTitle: "Tables, Table Styles, Subtotals, and Data Formatting",
  chapterSubtitle: "Modul Praktik Mandiri & Tutorial Interaktif Microsoft Excel",
  overview: {
    title: "General Theory: Pengantar Tabel, Subtotal, dan Pemformatan Data Visual",
    description: "Dalam pengelolaan data profesional, mengubah rentang data biasa menjadi Excel Table memberikan keuntungan luar biasa: filter otomatis, referensi terstruktur, perluasan otomatis, dan gaya visual siap pakai. Chapter ini mengajarkan siswa cara membuat dan mengelola Excel Table, menerapkan Table Styles, menghitung Subtotal bertingkat, serta memvisualisasikan data menggunakan Icon Sets, Data Bars, dan Sparklines.",
    coreCompetencies: [
      "Mampu menghapus kolom yang tidak diperlukan dan mengatur lebar kolom secara optimal menggunakan AutoFit",
      "Memahami pengaturan Print Titles dan Gridlines agar hasil cetak tabel profesional dan mudah dibaca",
      "Menguasai pembuatan Excel Table (Ctrl+T) beserta kelebihan dibanding rentang data biasa",
      "Menerapkan Table Styles dan Table Style Options termasuk Total Row untuk kalkulasi ringkasan otomatis",
      "Menggunakan fitur Subtotal bertingkat (Multi-Level) pada data yang telah diurutkan untuk menghasilkan outline ringkasan",
      "Mengimplementasikan visualisasi data kondisional: Icon Sets, Data Bars, dan Sparklines (Line & Win/Loss)"
    ]
  },
  topics: [
    {
      id: "topic-3-1",
      number: "3.1",
      title: "Deleting Columns & AutoFit Width",
      indonesianTitle: "Menghapus Kolom & Menyesuaikan Lebar Kolom Otomatis",
      summary: "Unneeded columns can be removed using Right-click > Delete. This shifts the columns to the right to fill the empty space. To make the data easier to read, use AutoFit Column Width by double-clicking the column boundary or selecting Home > Format > AutoFit Column Width.",
      shortcut: "Ctrl + - (Delete) / Alt + H + O + I (AutoFit)",
      ribbonPath: "Home → Cells → Format → AutoFit Column Width",
      theoryExplanation: `Ketika bekerja dengan data yang diekspor dari sistem lain atau database, sering kali tabel mengandung **kolom-kolom yang tidak relevan** dengan analisis yang sedang dilakukan. Kolom kosong, kolom kode internal sistem, atau kolom metadata teknis harus dihapus agar tabel menjadi ringkas dan fokus.

**Menghapus Kolom** dilakukan dengan cara:
1. Klik kanan pada header kolom (huruf di atas seperti A, B, C).
2. Pilih **Delete** dari menu konteks.
3. Kolom-kolom di sebelah kanan akan otomatis bergeser ke kiri untuk mengisi ruang kosong.

Setelah kolom tidak relevan dihapus, langkah selanjutnya adalah memastikan data mudah dibaca dengan mengatur lebar kolom. Fitur **AutoFit Column Width** secara otomatis menyesuaikan lebar kolom agar pas dengan konten terpanjang di kolom tersebut.

Ada 2 cara cepat menggunakan AutoFit:
* **Cara Mouse**: Arahkan kursor ke batas kanan header kolom (misal batas antara kolom B dan C) hingga kursor berubah menjadi panah dua arah (↔), lalu **klik ganda** (double-click).
* **Cara Ribbon**: Pilih kolom, buka **Home → Format → AutoFit Column Width**.`,
      keyPoints: [
        "**Delete Column vs Clear Contents**: Delete menghapus seluruh kolom dan menggeser kolom sebelahnya. Clear Contents hanya menghapus isi sel tanpa mengubah struktur kolom.",
        "**AutoFit Column Width**: Menyesuaikan lebar kolom otomatis berdasarkan konten terpanjang. Sangat berguna setelah data baru diimpor.",
        "**Seleksi Multi-Kolom**: Tahan Ctrl sambil mengklik header kolom untuk memilih beberapa kolom sekaligus sebelum menghapus.",
        "**Undo (Ctrl+Z)**: Selalu ingat bahwa penghapusan kolom bisa di-undo segera setelah dilakukan, tetapi tidak setelah file disimpan dan ditutup."
      ],
      proTips: "Untuk meng-AutoFit semua kolom sekaligus, tekan **Ctrl + A** untuk memilih seluruh worksheet, lalu buka **Home → Format → AutoFit Column Width**. Dalam satu klik, semua kolom akan otomatis menyesuaikan lebarnya!",
      commonMistakes: "Menghapus kolom dengan menekan tombol Delete pada keyboard — ini hanya menghapus isi sel (Clear Contents), bukan menghapus kolom itu sendiri. Gunakan klik kanan → Delete atau Ctrl + - untuk menghapus kolom secara struktural.",
      exerciseTask: "Hapus kolom 'Kode Internal' dan 'Catatan Sistem' dari tabel data produk, lalu gunakan AutoFit untuk merapikan semua kolom yang tersisa.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "3.1_DeleteAutoFit",
        sampleData: `No\tNama Produk\tKode Internal\tKategori\tHarga\tCatatan Sistem\tStok
1\tKeyboard Mechanical\tSYS-X01-KBD\tPeripheral\t450000\tAUTO_GEN_2024\t120
2\tMouse Gaming\tSYS-X02-MSE\tPeripheral\t280000\tAUTO_GEN_2024\t85
3\tMonitor 24 inch\tSYS-X03-MON\tDisplay\t2100000\tAUTO_GEN_2024\t34
4\tWebcam HD\tSYS-X04-WEB\tPeripheral\t350000\tAUTO_GEN_2024\t67
5\tHeadset USB\tSYS-X05-HST\tAudio\t175000\tAUTO_GEN_2024\t150`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">No</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Nama Produk</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Kode Internal</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Kategori</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:right;background-color:#107c41;color:#ffffff;">Harga</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Catatan Sistem</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Stok</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Keyboard Mechanical</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">SYS-X01-KBD</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Peripheral</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">450000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">AUTO_GEN_2024</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">120</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Mouse Gaming</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">SYS-X02-MSE</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Peripheral</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">280000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">AUTO_GEN_2024</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">3</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Monitor 24 inch</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">SYS-X03-MON</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Display</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">2100000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">AUTO_GEN_2024</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">34</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">4</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Webcam HD</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">SYS-X04-WEB</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Peripheral</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">350000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">AUTO_GEN_2024</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">67</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">5</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Headset USB</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">SYS-X05-HST</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Audio</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">175000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">AUTO_GEN_2024</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">150</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka aplikasi Microsoft Excel di komputer kamu. Klik tanda tambah (+) di bagian paling bawah untuk **membuat Sheet Baru**, lalu beri nama: `3.1_DeleteAutoFit`.",
          "Klik tombol **'📋 Salin Tabel Latihan'** di bawah ini.",
          "Klik sel **A1** di sheet `3.1_DeleteAutoFit` **SATU KALI SAJA**, lalu tekan **Ctrl + V** agar tabel tertempel sempurna.",
          "Klik header kolom **C** ('Kode Internal') untuk memilih seluruh kolom C.",
          "Klik kanan pada header kolom C, lalu pilih **Delete**. Kolom 'Kode Internal' akan hilang dan kolom di sebelahnya bergeser ke kiri.",
          "Sekarang cari kolom **Catatan Sistem** (yang kini menjadi kolom E setelah penghapusan). Klik header kolom tersebut, klik kanan → **Delete**.",
          "Tekan **Ctrl + A** untuk memilih seluruh data, lalu buka tab **Home → Format → AutoFit Column Width**.",
          "Perhatikan bahwa semua kolom kini sudah rapi dan lebarnya pas dengan konten terpanjang!"
        ],
        expectedResult: "Tabel hanya berisi kolom No, Nama Produk, Kategori, Harga, dan Stok — tanpa kolom yang tidak diperlukan. Semua lebar kolom otomatis menyesuaikan konten."
      }
    },
    {
      id: "topic-3-2",
      number: "3.2",
      title: "Print Titles & Gridlines",
      indonesianTitle: "Pengaturan Cetak: Judul Berulang & Garis Kisi",
      summary: "Print Titles repeats selected rows on every printed page. Gridlines can also be printed through Page Layout > Sheet Options > Gridlines > Print. Use Print Preview to check the page layout before printing.",
      shortcut: "Ctrl + P (Print Preview) / Alt + P + I (Print Titles)",
      ribbonPath: "Page Layout → Print Titles / Sheet Options → Gridlines → Print",
      theoryExplanation: `Saat mencetak tabel data yang sangat panjang (puluhan hingga ratusan baris), halaman kedua dan seterusnya tidak akan memiliki baris judul kolom (*header row*). Hal ini menyulitkan pembaca dokumen cetak karena tidak mengetahui kolom mana yang berisi data apa.

Fitur **Print Titles** menyelesaikan masalah ini dengan mengulangi baris tertentu di bagian atas setiap halaman cetak:
1. Buka tab **Page Layout** → klik **Print Titles** di grup *Page Setup*.
2. Pada kotak **Rows to repeat at top**, klik baris header tabel (biasanya $1:$1).
3. Klik **OK**, maka setiap halaman cetak akan selalu menampilkan baris judul di bagian atas.

**Gridlines (Garis Kisi)** adalah garis bantu tipis yang terlihat di layar Excel tetapi **tidak dicetak secara default**. Untuk mencetak gridlines:
* Buka tab **Page Layout** → di grup *Sheet Options*, centang kotak **Gridlines → Print**.
* Alternatifnya, buka dialog Page Setup → tab Sheet → centang **Gridlines**.

Selalu gunakan **Print Preview** (Ctrl + P) untuk memeriksa tampilan halaman sebelum mencetak, agar memastikan judul berulang dan gridlines sudah tampil dengan benar.`,
      keyPoints: [
        "**Rows to repeat at top**: Biasanya diisi dengan referensi baris header seperti $1:$1 agar judul kolom muncul di setiap halaman cetak.",
        "**Columns to repeat at left**: Opsi tambahan untuk mengulangi kolom tertentu di sisi kiri setiap halaman cetak (berguna untuk tabel horizontal yang sangat lebar).",
        "**Gridlines Print**: Secara default tidak aktif. Harus diaktifkan manual melalui Page Layout → Sheet Options → Gridlines → Print.",
        "**Print Preview (Ctrl+P)**: Wajib diperiksa sebelum mencetak untuk memastikan layout halaman sudah sesuai harapan."
      ],
      proTips: "Jangan lupa bahwa **Print Titles** dan **Gridlines Print** adalah dua pengaturan yang terpisah. Print Titles mengulang baris header, sedangkan Gridlines menampilkan garis bantu antar sel. Keduanya perlu diaktifkan secara terpisah untuk hasil cetak yang sempurna.",
      commonMistakes: "Mencetak tabel panjang tanpa mengaktifkan Print Titles, sehingga halaman kedua dan seterusnya tidak memiliki baris judul dan pembaca bingung membedakan kolom data.",
      exerciseTask: "Atur Print Titles agar Baris 1 terulang di setiap halaman cetak, aktifkan Gridlines Print, lalu periksa hasilnya melalui Print Preview.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "3.2_PrintTitles",
        sampleData: `No\tNama Siswa\tKelas\tNilai UTS\tNilai UAS\tRata-Rata
1\tAndi Pratama\t10-RPL\t85\t88\t86.5
2\tBunga Citra\t10-TKJ\t78\t82\t80.0
3\tCahyo Nugroho\t10-RPL\t92\t90\t91.0
4\tDinda Kirana\t10-DKV\t70\t75\t72.5
5\tEko Purwanto\t10-TKJ\t65\t68\t66.5
6\tFani Maulida\t10-RPL\t95\t97\t96.0
7\tGilang Ramadhan\t10-DKV\t72\t70\t71.0
8\tHesti Wulandari\t10-TKJ\t88\t85\t86.5
9\tIrfan Hakim\t10-RPL\t60\t62\t61.0
10\tJasmine Putri\t10-DKV\t82\t80\t81.0`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">No</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Nama Siswa</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">Kelas</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Nilai UTS</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Nilai UAS</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Rata-Rata</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Andi Pratama</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">88</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">86.5</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bunga Citra</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-TKJ</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">78</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">82</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">80.0</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">3</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Cahyo Nugroho</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-RPL</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">92</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">90</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">91.0</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">4</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Dinda Kirana</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-DKV</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">70</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">75</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">72.5</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">5</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Eko Purwanto</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">10-TKJ</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">65</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">68</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">66.5</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru lalu beri nama: `3.2_PrintTitles`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** satu kali lalu **Ctrl + V**.",
          "Klik tab **Page Layout** di Ribbon atas.",
          "Klik tombol **Print Titles** (di grup *Page Setup*).",
          "Pada kotak **Rows to repeat at top**, klik langsung pada **Baris 1** di worksheet sehingga terisi `$1:$1`.",
          "Klik **OK** untuk menyimpan pengaturan.",
          "Masih di tab **Page Layout**, cari grup **Sheet Options** → centang kotak **Gridlines → Print**.",
          "Tekan **Ctrl + P** untuk membuka Print Preview. Periksa bahwa setiap halaman menampilkan baris judul di atas dan garis kisi terlihat."
        ],
        expectedResult: "Setiap halaman cetak menampilkan baris judul (No, Nama Siswa, Kelas, dll.) di bagian atas, dan garis kisi antar sel tercetak dengan jelas."
      }
    },
    {
      id: "topic-3-3",
      number: "3.3",
      title: "Excel Tables vs Regular Ranges",
      indonesianTitle: "Excel Table vs Rentang Data Biasa",
      summary: "An Excel Table is a structured data range created through Insert > Table (Ctrl+T). It provides automatic filters, structured references, automatic expansion, and built-in styles. If table features are no longer needed, use Table Design > Convert to Range.",
      shortcut: "Ctrl + T (Insert Table)",
      ribbonPath: "Insert → Table (Ctrl+T) / Table Design → Convert to Range",
      theoryExplanation: `**Excel Table** adalah fitur pengelolaan data terstruktur yang sangat powerful. Berbeda dari *data range* biasa yang hanya berupa sekumpulan sel, Table memiliki kemampuan pintar yang otomatis:

**Kelebihan Excel Table vs Data Range Biasa:**
1. **Filter Otomatis**: Setiap kolom header langsung memiliki dropdown filter tanpa perlu mengaktifkannya manual.
2. **Structured References**: Rumus di dalam tabel menggunakan nama kolom (misal: =[@Harga]*[@Qty]) yang lebih mudah dibaca dibanding referensi sel biasa (=C2*D2).
3. **Auto Expansion**: Saat mengetik data baru di baris tepat di bawah tabel, tabel otomatis meluas mencakup baris baru tersebut beserta format dan rumusnya.
4. **Built-in Styles**: Excel menyediakan puluhan desain visual siap pakai untuk tabel.

**Cara Membuat Excel Table:**
1. Pilih (blok) seluruh data termasuk header.
2. Tekan **Ctrl + T** atau buka **Insert → Table**.
3. Excel akan memunculkan dialog: *"Where is the data for your table?"* — pastikan range sudah benar.
4. Centang **"My table has headers"** jika baris pertama berisi judul kolom.

**Convert to Range:**
Jika fitur tabel tidak lagi dibutuhkan, klik di mana saja dalam tabel → buka tab **Table Design** → klik **Convert to Range**. Data dan format visual tetap ada, tetapi fitur otomatis seperti filter dan structured references akan dihilangkan.`,
      keyPoints: [
        "**My table has headers**: Wajib dicentang saat membuat tabel agar baris pertama dikenali sebagai header dan bukan data.",
        "**Banded Rows (Baris Bergaris)**: Pola warna selang-seling otomatis yang memudahkan membaca baris data. Banded rows akan terus berlanjut saat data baru ditambahkan.",
        "**Auto Expansion**: Mengetik data baru di baris tepat di bawah tabel membuat tabel otomatis membesar, termasuk format dan rumus.",
        "**Convert to Range**: Menghapus fitur tabel (filter, structured references) tetapi mempertahankan data dan format visual."
      ],
      proTips: "Setelah mengubah data menjadi Table, perhatikan bahwa tab **Table Design** muncul di Ribbon atas. Tab ini hanya terlihat saat kursor berada di dalam tabel. Dari sini kamu bisa mengganti gaya, menambah Total Row, dan mengelola semua opsi tabel!",
      commonMistakes: "Tidak mencentang 'My table has headers' saat membuat tabel, sehingga Excel menganggap baris pertama sebagai data biasa dan menambahkan header generik seperti Column1, Column2 di atasnya.",
      exerciseTask: "Ubah data penjualan menjadi Excel Table menggunakan Ctrl+T, pastikan headers terdeteksi, lalu coba tambahkan baris data baru di bawah tabel untuk melihat Auto Expansion bekerja.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "3.3_ExcelTable",
        sampleData: `ID Produk\tNama Produk\tKategori\tHarga Satuan\tQty Terjual\tTotal Penjualan
PRD-01\tFlashdisk 32GB\tStorage\t65000\t45\t2925000
PRD-02\tSSD 256GB\tStorage\t350000\t18\t6300000
PRD-03\tKabel HDMI 2m\tAksesoris\t45000\t120\t5400000
PRD-04\tWebcam Logitech\tPeripheral\t580000\t25\t14500000
PRD-05\tAdapter USB-C\tAksesoris\t85000\t60\t5100000
PRD-06\tRAM DDR4 8GB\tKomponen\t320000\t30\t9600000`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">ID Produk</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Nama Produk</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Kategori</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Harga Satuan</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Qty Terjual</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Total Penjualan</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">PRD-01</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Flashdisk 32GB</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Storage</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">65000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">45</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">2925000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">PRD-02</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">SSD 256GB</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Storage</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">350000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">18</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">6300000</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">PRD-03</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Kabel HDMI 2m</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Aksesoris</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">45000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">120</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">5400000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">PRD-04</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Webcam Logitech</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Peripheral</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">580000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">25</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">14500000</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">PRD-05</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Adapter USB-C</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Aksesoris</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">60</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">5100000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">PRD-06</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">RAM DDR4 8GB</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Komponen</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">320000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">30</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">9600000</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru bernama: `3.3_ExcelTable`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** satu kali, lalu **Ctrl + V**.",
          "Pilih (blok) seluruh data termasuk header: dari sel **A1** hingga sel **F7**.",
          "Tekan **Ctrl + T** pada keyboard (atau buka menu **Insert → Table**).",
          "Pada dialog yang muncul, pastikan **'My table has headers'** tercentang, lalu klik **OK**.",
          "Perhatikan: filter dropdown otomatis muncul di setiap header kolom, dan tabel mendapatkan gaya visual banded rows!",
          "Coba ketik data baru di baris tepat di bawah tabel (sel A8). Perhatikan tabel otomatis meluas (*Auto Expansion*).",
          "Untuk mengubah kembali ke rentang biasa: klik dalam tabel → tab **Table Design** → klik **Convert to Range** → klik **Yes**."
        ],
        expectedResult: "Data berubah menjadi Excel Table resmi dengan filter dropdown di header, banded rows otomatis, dan tabel secara otomatis meluas saat data baru ditambahkan di bawahnya."
      }
    },
    {
      id: "topic-3-4",
      number: "3.4",
      title: "Table Styles & Style Options",
      indonesianTitle: "Gaya Tabel & Opsi Tampilan Tabel",
      summary: "Table Styles provide ready-made formatting for headers, borders, and banded rows. Table Style Options allow users to control elements such as Header Row, Banded Rows, First Column, Last Column, and Total Row with summary calculations.",
      shortcut: "Alt + J + T (Table Design Tab)",
      ribbonPath: "Table Design → Table Styles / Table Style Options → Total Row",
      theoryExplanation: `Setelah data diubah menjadi Excel Table, tab **Table Design** menyediakan dua fitur utama untuk mengkustomisasi tampilan dan fungsi tabel:

**1. Table Styles (Galeri Gaya Tabel)**
Table Styles adalah koleksi desain visual siap pakai yang mengatur warna header, border, dan pola banded rows secara konsisten. Excel menyediakan 3 kategori gaya:
* **Light**: Gaya minimalis dengan warna lembut, cocok untuk dokumen formal.
* **Medium**: Gaya menengah dengan kontras warna yang seimbang.
* **Dark**: Gaya dengan header gelap dan baris berwarna kuat, cocok untuk presentasi.

**2. Table Style Options (Opsi Gaya Tabel)**
Di dalam tab Table Design, grup **Table Style Options** memungkinkan pengaturan elemen individual:
* **Header Row**: Menampilkan/menyembunyikan baris header dengan filter.
* **Total Row**: Menambahkan baris ringkasan di bagian bawah tabel untuk kalkulasi otomatis.
* **Banded Rows**: Mengaktifkan/menonaktifkan pola warna selang-seling pada baris.
* **First Column / Last Column**: Memberikan format tebal pada kolom pertama atau terakhir.
* **Filter Button**: Menampilkan/menyembunyikan tombol dropdown filter di header.

**Total Row** adalah fitur yang sangat berguna — saat diaktifkan, baris terakhir tabel berubah menjadi baris kalkulasi otomatis. Klik dropdown di setiap sel Total Row untuk memilih fungsi ringkasan:
* **Sum** (Jumlah), **Average** (Rata-rata), **Count** (Hitung), **Max** (Nilai tertinggi), **Min** (Nilai terendah), dan lainnya.
* Label teks di kolom pertama Total Row juga bisa dikustomisasi.`,
      keyPoints: [
        "**Table Styles Gallery**: Tiga kategori desain (Light, Medium, Dark) yang bisa diubah kapan saja tanpa mengubah data.",
        "**Total Row**: Baris ringkasan otomatis di bawah tabel dengan pilihan fungsi: Sum, Average, Count, Max, Min, dan lainnya.",
        "**Banded Rows / Columns**: Pola warna selang-seling untuk memudahkan pembacaan baris. Bisa diaktifkan atau dinonaktifkan melalui checkbox.",
        "**First / Last Column**: Opsi untuk memberikan format visual tebal (bold) pada kolom pertama atau kolom terakhir tabel."
      ],
      proTips: "Gunakan **Total Row** untuk menampilkan ringkasan cepat tanpa perlu menulis rumus manual. Klik sel di baris Total, lalu pilih fungsi dari dropdown. Misalnya, pilih **Sum** di kolom Total Penjualan untuk langsung melihat total keseluruhan!",
      commonMistakes: "Mengubah Table Style saat berpikir bahwa data juga akan berubah. Perlu diingat: Table Styles hanya mengubah tampilan visual (warna, border, font), data di dalam sel tetap sama persis.",
      exerciseTask: "Terapkan Table Style 'Medium 2' pada tabel data, aktifkan Total Row, lalu pilih fungsi Sum untuk kolom Total Penjualan dan Average untuk kolom Qty.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "3.4_TableStyles",
        sampleData: `Produk\tKategori\tQty Terjual\tHarga Satuan\tTotal Penjualan
Laptop ASUS\tKomputer\t12\t8500000\t102000000
Mouse Logitech\tPeripheral\t85\t250000\t21250000
Monitor Samsung\tDisplay\t20\t3200000\t64000000
Keyboard Razer\tPeripheral\t45\t750000\t33750000
SSD Kingston\tStorage\t60\t400000\t24000000
RAM Corsair\tKomponen\t35\t550000\t19250000`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Produk</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Kategori</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Qty Terjual</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Harga Satuan</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Total Penjualan</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Laptop ASUS</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Komputer</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">12</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">8500000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">102000000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">Mouse Logitech</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Peripheral</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">250000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">21250000</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Monitor Samsung</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Display</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">20</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">3200000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">64000000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">Keyboard Razer</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Peripheral</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">45</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">750000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">33750000</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">SSD Kingston</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Storage</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">60</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">400000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">24000000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">RAM Corsair</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Komponen</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">35</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">550000</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">19250000</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru bernama: `3.4_TableStyles`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** satu kali, lalu **Ctrl + V**.",
          "Blok seluruh data (A1:E7), tekan **Ctrl + T** untuk membuat Excel Table, pastikan centang **'My table has headers'**.",
          "Klik di mana saja dalam tabel. Perhatikan tab **Table Design** muncul di Ribbon atas.",
          "Di grup **Table Styles**, klik galeri panah bawah (▼) untuk membuka seluruh koleksi gaya. Pilih gaya **Medium 2** (biru muda).",
          "Di grup **Table Style Options**, centang **Total Row**. Perhatikan baris ringkasan muncul di bagian bawah tabel!",
          "Klik sel di kolom **Total Penjualan** pada baris Total. Klik dropdown kecil yang muncul, pilih **Sum**.",
          "Klik sel di kolom **Qty Terjual** pada baris Total. Pilih **Average** dari dropdown.",
          "Coba juga centang/hapus centang **Banded Rows** dan **First Column** untuk melihat efeknya!"
        ],
        expectedResult: "Tabel menampilkan gaya Medium 2 dengan baris Total di bawah yang menunjukkan Sum total penjualan dan Average qty terjual secara otomatis."
      }
    },
    {
      id: "topic-3-5",
      number: "3.5",
      title: "Multi-Level Subtotals",
      indonesianTitle: "Subtotal Bertingkat pada Data Terurut",
      summary: "The Subtotal feature (Data > Outline > Subtotal) automatically calculates summaries such as Sum or Average for groups of data. For multiple levels, the data must be sorted by the grouping column first. Excel creates an outline structure that allows users to show or hide detailed data.",
      shortcut: "Alt + A + B (Data → Subtotal)",
      ribbonPath: "Data → Outline → Subtotal",
      theoryExplanation: `Fitur **Subtotal** adalah alat ringkasan otomatis yang menghitung nilai agregat (Sum, Average, Count, Max, Min) untuk setiap grup data berdasarkan kolom pengelompokan tertentu.

**Prasyarat Penting:** Data **harus diurutkan (sort)** berdasarkan kolom pengelompokan terlebih dahulu sebelum menerapkan Subtotal. Jika data belum terurut, Subtotal akan menghasilkan ringkasan yang salah karena grup data terpecah.

**Langkah Menerapkan Subtotal:**
1. **Urutkan data** berdasarkan kolom pengelompokan (misal: Sort by Region A→Z).
2. Buka **Data → Outline → Subtotal**.
3. Pada dialog Subtotal:
   - **At each change in**: Pilih kolom pengelompokan (misal: Region).
   - **Use function**: Pilih fungsi ringkasan (Sum, Average, Count, dsb).
   - **Add subtotal to**: Centang kolom-kolom yang akan dihitung ringkasannya.

**Multi-Level Subtotals:**
Untuk subtotal bertingkat (misal: pertama per Region, lalu per Periode di dalam setiap Region):
1. **Urutkan data** berdasarkan dua kolom: Region terlebih dahulu, lalu Periode.
2. Terapkan Subtotal pertama berdasarkan **Region**.
3. Terapkan Subtotal kedua berdasarkan **Periode** — **PENTING: hilangkan centang "Replace current subtotals"** agar subtotal sebelumnya tidak terhapus.

Setelah Subtotal diterapkan, Excel membuat **Outline (tingkat kerangka)** di sisi kiri worksheet:
* **Level 1**: Hanya menampilkan Grand Total keseluruhan.
* **Level 2**: Menampilkan subtotal per grup.
* **Level 3**: Menampilkan semua data detail beserta subtotal.`,
      keyPoints: [
        "**Data Harus Terurut**: Subtotal hanya bekerja benar jika data sudah di-sort berdasarkan kolom pengelompokan. Jika tidak, grup data terpecah dan ringkasan salah.",
        "**Replace current subtotals**: Hilangkan centang opsi ini saat menambahkan level subtotal kedua agar level sebelumnya tidak terhapus.",
        "**Outline Levels (1, 2, 3)**: Tombol angka di pojok kiri atas worksheet yang memungkinkan pengguna menyembunyikan/menampilkan detail data.",
        "**Remove All**: Tombol untuk menghapus seluruh subtotal dan mengembalikan data ke bentuk semula."
      ],
      proTips: "Saat menerapkan Multi-Level Subtotal, selalu terapkan level yang **lebih besar** (misal: Region) terlebih dahulu, baru kemudian level yang **lebih kecil** (misal: Periode). Dan jangan lupa untuk **menghilangkan centang 'Replace current subtotals'** pada level kedua!",
      commonMistakes: "Menerapkan Subtotal pada data yang belum diurutkan, sehingga data dengan nilai grup yang sama tersebar di berbagai tempat dan Subtotal menganggap setiap kemunculan terpisah sebagai grup berbeda.",
      exerciseTask: "Urutkan data berdasarkan Region lalu Periode, terapkan Subtotal Sum per Region, lalu tambahkan Subtotal Average per Periode tanpa menghapus subtotal Region.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "3.5_Subtotals",
        sampleData: `Region\tPeriode\tProduk\tQty\tTotal Penjualan
Barat\tQ1\tLaptop\t15\t127500000
Barat\tQ1\tPrinter\t20\t30000000
Barat\tQ2\tLaptop\t18\t153000000
Barat\tQ2\tPrinter\t12\t18000000
Timur\tQ1\tLaptop\t22\t187000000
Timur\tQ1\tPrinter\t10\t15000000
Timur\tQ2\tLaptop\t25\t212500000
Timur\tQ2\tPrinter\t15\t22500000`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Region</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:center;background-color:#107c41;color:#ffffff;">Periode</th>
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Produk</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Qty</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Total Penjualan</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Barat</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">Q1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Laptop</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">15</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">127500000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">Barat</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">Q1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Printer</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">20</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">30000000</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Barat</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">Q2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Laptop</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">18</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">153000000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">Barat</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">Q2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Printer</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">12</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">18000000</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Timur</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">Q1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Laptop</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">22</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">187000000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">Timur</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">Q1</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Printer</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">10</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">15000000</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Timur</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">Q2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Laptop</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">25</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">212500000</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">Timur</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">Q2</td><td style="border:1px solid #d4d4d4;padding:5px 10px;">Printer</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">15</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">22500000</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru bernama: `3.5_Subtotals`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** satu kali, lalu **Ctrl + V**.",
          "**PENTING: Urutkan data terlebih dahulu!** Blok semua data, buka **Data → Sort**. Sort by **Region** (A to Z), kemudian **Add Level** → Then by **Periode** (A to Z). Klik **OK**.",
          "Buka **Data → Outline → Subtotal** (atau tekan Alt + A + B).",
          "Pada dialog Subtotal: pilih **At each change in: Region**, **Use function: Sum**, centang kolom **Total Penjualan**. Klik **OK**.",
          "Perhatikan baris Subtotal muncul di bawah setiap grup Region (Barat Sum dan Timur Sum)!",
          "Sekarang buka kembali **Data → Subtotal**. Pilih **At each change in: Periode**, **Use function: Sum**, centang **Total Penjualan**. **HILANGKAN centang 'Replace current subtotals'**. Klik **OK**.",
          "Gunakan tombol angka **1, 2, 3, 4** di pojok kiri atas untuk melihat berbagai tingkat detail outline!"
        ],
        expectedResult: "Data menampilkan Subtotal bertingkat: Sum per Periode di dalam setiap Region, Sum per Region, dan Grand Total keseluruhan. Outline level 1-4 berfungsi untuk show/hide detail."
      }
    },
    {
      id: "topic-3-6",
      number: "3.6",
      title: "Icon Sets, Data Bars & Sparklines",
      indonesianTitle: "Visualisasi Kondisional: Icon Sets, Data Bars, dan Sparklines",
      summary: "Icon Sets use symbols such as arrows to show whether values are relatively high, medium, or low. Data Bars display colored bars inside cells. Sparklines are small charts placed inside a single cell — Line Sparklines for trends, Win/Loss Sparklines for binary results.",
      shortcut: "Alt + H + L (Conditional Formatting) / Alt + N + SL (Sparklines)",
      ribbonPath: "Home → Conditional Formatting → Icon Sets / Data Bars | Insert → Sparklines",
      theoryExplanation: `Selain Conditional Formatting berbasis warna sel dan font yang sudah dipelajari di Chapter 2, Excel menyediakan **tiga teknik visualisasi data dalam sel** yang sangat powerful:

**1. Icon Sets (Ikon Indikator)**
Icon Sets menampilkan simbol visual kecil di dalam sel berdasarkan posisi relatif nilai tersebut terhadap nilai lain di rentang yang sama. Contoh ikon populer:
* **Panah Tiga Arah**: ↑ (tinggi), → (sedang), ↓ (rendah)
* **Traffic Lights**: 🟢 (baik), 🟡 (sedang), 🔴 (buruk)
* **Bintang Rating**: ★★★ (3 bintang), ★★ (2 bintang), ★ (1 bintang)

Cara menerapkan: Pilih rentang data → **Home → Conditional Formatting → Icon Sets** → pilih set ikon yang diinginkan.

**2. Data Bars (Batang Data Dalam Sel)**
Data Bars menampilkan batang horizontal berwarna di dalam sel. Panjang batang proporsional dengan nilai sel — semakin besar nilai, semakin panjang batangnya. Tersedia dalam dua varian:
* **Gradient Fill**: Batang dengan gradasi warna dari kiri ke kanan.
* **Solid Fill**: Batang dengan warna solid rata.

Cara menerapkan: Pilih rentang data → **Home → Conditional Formatting → Data Bars** → pilih gaya yang diinginkan.

**3. Sparklines (Grafik Mini Dalam Sel)**
Sparklines adalah grafik miniatur yang ditempatkan di dalam satu sel untuk menampilkan tren atau pola data. Ada tiga jenis:
* **Line Sparkline**: Garis yang menunjukkan tren naik/turun seiring waktu. Cocok untuk data penjualan bulanan.
* **Column Sparkline**: Batang mini untuk membandingkan nilai dalam satu seri.
* **Win/Loss Sparkline**: Menampilkan hanya dua kondisi (menang/kalah, lulus/gagal, positif/negatif) dalam bentuk batang ke atas dan ke bawah.

Cara membuat: Pilih sel tujuan → **Insert → Sparklines** → pilih jenis → tentukan **Data Range** (rentang sumber data) dan **Location Range** (sel tempat sparkline ditampilkan).

Warna dan titik sorotan (High Point, Low Point, First Point, Last Point) dapat diatur melalui **Sparkline Tools → Design**.`,
      keyPoints: [
        "**Icon Sets**: Menampilkan simbol visual (panah, lampu lalu lintas, bintang) berdasarkan posisi relatif nilai dalam rentang data.",
        "**Data Bars**: Batang horizontal dalam sel yang proporsional terhadap nilai — tersedia dalam Gradient Fill dan Solid Fill.",
        "**Line Sparklines**: Grafik garis mini dalam satu sel untuk menunjukkan tren data dari waktu ke waktu. Sangat cocok untuk data time-series.",
        "**Win/Loss Sparklines**: Grafik mini yang hanya menunjukkan dua kondisi (positif/negatif), cocok untuk hasil lulus/gagal atau menang/kalah.",
        "**Sparkline Design**: Warna garis, titik sorotan (High, Low, First, Last Point), dan tipe marker dapat dikustomisasi melalui Sparkline Tools."
      ],
      proTips: "Sparklines sangat berguna untuk dashboard ringkas. Tempatkan data bulanan di kolom B-M (Januari-Desember), lalu buat Line Sparkline di kolom N yang merangkum tren setahun penuh dalam satu sel! Aktifkan **High Point** dan **Low Point** di Sparkline Design untuk menyorot bulan terbaik dan terburuk.",
      commonMistakes: "Mencampur Sparklines dengan Conditional Formatting di sel yang sama. Sparklines adalah objek grafik terpisah yang ditimpa di atas sel, sedangkan Data Bars dan Icon Sets adalah bagian dari Conditional Formatting. Keduanya bisa bentrok secara visual jika diterapkan bersamaan.",
      exerciseTask: "Terapkan Icon Sets (3 Arrows) pada kolom Skor, Data Bars (Gradient Fill Biru) pada kolom Penjualan, dan buat Line Sparkline dari data bulanan di kolom terpisah.",
      desktopTutorial: {
        title: "Tutorial Mencoba Langsung di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "3.6_Sparklines",
        sampleData: `Nama Cabang\tJan\tFeb\tMar\tApr\tMei\tJun\tSkor Semester
Jakarta\t120\t135\t150\t140\t160\t175\t88
Bandung\t95\t100\t90\t110\t105\t120\t72
Surabaya\t140\t130\t155\t165\t170\t180\t92
Semarang\t80\t85\t75\t90\t95\t100\t65
Yogyakarta\t110\t115\t120\t125\t130\t145\t80`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">
  <thead>
    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">Nama Cabang</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Jan</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Feb</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Mar</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Apr</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Mei</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Jun</th>
      <th style="border:1px solid #0d6836;padding:6px 10px;text-align:right;background-color:#107c41;color:#ffffff;">Skor Semester</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Jakarta</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">120</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">135</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">150</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">140</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">160</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">175</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">88</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">Bandung</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">95</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">100</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">90</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">110</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">105</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">120</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">72</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Surabaya</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">140</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">130</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">155</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">165</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">170</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">180</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">92</td></tr>
    <tr style="background-color:#f9fbf9;"><td style="border:1px solid #d4d4d4;padding:5px 10px;">Semarang</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">80</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">85</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">75</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">90</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">95</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">100</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">65</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;">Yogyakarta</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">110</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">115</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">120</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">125</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">130</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">145</td><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:right;">80</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka file Excel kamu, buat Sheet Baru bernama: `3.6_Sparklines`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, klik sel **A1** satu kali, lalu **Ctrl + V**.",
          "**Icon Sets**: Blok kolom **Skor Semester** (H2:H6). Buka **Home → Conditional Formatting → Icon Sets** → pilih **3 Arrows (Colored)**.",
          "**Data Bars**: Blok kolom **Jan** (B2:B6). Buka **Home → Conditional Formatting → Data Bars** → pilih **Gradient Fill → Blue**.",
          "**Sparklines**: Klik sel **I2** (kolom kosong di sebelah kanan). Buka tab **Insert → Line Sparkline**.",
          "Pada dialog Sparkline, isi **Data Range** dengan `B2:G2` dan **Location Range** dengan `I2`. Klik **OK**.",
          "Kursor masih di I2. Di tab **Sparkline → Design**, centang **High Point** dan **Low Point** untuk menyoroti titik tertinggi dan terendah.",
          "Salin sel I2 ke bawah (I3:I6) menggunakan Fill Handle untuk membuat Sparkline semua cabang.",
          "Di sel I1, ketik judul kolom: `Tren 6 Bulan`."
        ],
        expectedResult: "Kolom Skor menampilkan ikon panah (hijau ↑ untuk tinggi, kuning → sedang, merah ↓ rendah). Kolom Jan menampilkan Data Bars biru. Kolom Tren 6 Bulan menampilkan Line Sparkline dengan titik tertinggi dan terendah yang disorot."
      }
    }
  ],
  quiz: [
    {
      id: 1,
      topicRef: "3.1",
      question: "Apa perbedaan utama antara menekan tombol Delete di keyboard dan menggunakan klik kanan > Delete pada sebuah kolom di Excel?",
      options: [
        "Keduanya melakukan hal yang sama persis",
        "Tombol Delete menghapus isi sel saja (Clear Contents), sedangkan klik kanan > Delete menghapus seluruh struktur kolom dan menggeser kolom sebelahnya",
        "Klik kanan > Delete hanya menghapus satu sel, bukan seluruh kolom",
        "Tombol Delete menghapus kolom beserta format dan isinya secara permanen tanpa bisa di-undo"
      ],
      correctIndex: 1,
      explanation: "Menekan tombol Delete pada keyboard hanya melakukan Clear Contents (menghapus isi/teks dalam sel), sedangkan klik kanan > Delete pada header kolom akan menghapus seluruh kolom secara struktural dan menggeser kolom di sebelah kanannya ke kiri."
    },
    {
      id: 2,
      topicRef: "3.2",
      question: "Seorang guru mencetak daftar nilai yang berisi 50 siswa (3 halaman cetak). Halaman kedua dan ketiga tidak memiliki baris judul kolom. Fitur apa yang harus digunakan untuk mengatasi masalah ini?",
      options: [
        "Conditional Formatting → Highlight Cell Rules",
        "Page Layout → Print Titles → Rows to repeat at top",
        "View → Freeze Panes → Freeze Top Row",
        "Home → Format → Row Height"
      ],
      correctIndex: 1,
      explanation: "Print Titles > Rows to repeat at top ($1:$1) akan mengulangi baris header di bagian atas setiap halaman cetak, sehingga pembaca selalu mengetahui arti setiap kolom. Freeze Panes hanya berfungsi di layar, bukan saat mencetak."
    },
    {
      id: 3,
      topicRef: "3.3",
      question: "Apa keuntungan utama menggunakan Excel Table (Ctrl+T) dibandingkan data range biasa?",
      options: [
        "Excel Table membuat data tidak bisa diedit oleh pengguna lain",
        "Excel Table menyediakan filter otomatis, structured references, auto expansion, dan built-in styles secara bersamaan",
        "Excel Table hanya bisa berisi data angka, bukan teks",
        "Excel Table mengunci seluruh worksheet agar tidak bisa dimodifikasi"
      ],
      correctIndex: 1,
      explanation: "Excel Table memberikan empat keuntungan utama sekaligus: filter dropdown otomatis di setiap header, structured references untuk rumus yang lebih mudah dibaca, auto expansion saat data baru ditambahkan, dan akses ke galeri Table Styles."
    },
    {
      id: 4,
      topicRef: "3.4",
      question: "Fitur Total Row pada Excel Table memungkinkan pengguna untuk...",
      options: [
        "Menambahkan baris kosong di akhir tabel untuk catatan kaki",
        "Menampilkan baris ringkasan dengan fungsi otomatis seperti Sum, Average, Count, atau Max di bagian bawah tabel",
        "Menghitung jumlah karakter di setiap sel tabel",
        "Membekukan baris terakhir tabel agar tidak tergulung saat scroll"
      ],
      correctIndex: 1,
      explanation: "Total Row adalah baris ringkasan otomatis di bagian bawah tabel. Setiap sel di baris Total memiliki dropdown untuk memilih fungsi ringkasan: Sum, Average, Count, Max, Min, dan lainnya — tanpa perlu menulis rumus manual."
    },
    {
      id: 5,
      topicRef: "3.5",
      question: "Mengapa data HARUS diurutkan (sort) terlebih dahulu sebelum menerapkan fitur Subtotal?",
      options: [
        "Karena Subtotal tidak bisa bekerja pada data yang memiliki lebih dari 100 baris",
        "Karena jika data belum terurut, grup data yang sama akan tersebar di berbagai tempat dan Subtotal menganggap setiap kemunculan sebagai grup terpisah",
        "Karena Subtotal hanya bisa diterapkan pada kolom pertama tabel saja",
        "Karena Excel akan menghapus baris duplikat secara otomatis saat Subtotal diterapkan"
      ],
      correctIndex: 1,
      explanation: "Subtotal bekerja berdasarkan perubahan nilai pada kolom pengelompokan secara berurutan. Jika data belum terurut, nilai grup yang sama muncul di lokasi berbeda dan Subtotal akan membuat ringkasan terpisah untuk setiap kemunculan, menghasilkan output yang salah."
    },
    {
      id: 6,
      topicRef: "3.6",
      question: "Apa perbedaan utama antara Line Sparkline dan Win/Loss Sparkline di Excel?",
      options: [
        "Line Sparkline hanya bisa menampilkan 10 titik data, sedangkan Win/Loss tidak dibatasi",
        "Line Sparkline menunjukkan tren data dalam bentuk garis kontinu, sedangkan Win/Loss Sparkline hanya menampilkan dua kondisi (positif/negatif) dalam bentuk batang ke atas dan ke bawah",
        "Win/Loss Sparkline hanya bisa digunakan pada data teks, bukan angka",
        "Keduanya menghasilkan tampilan yang persis sama"
      ],
      correctIndex: 1,
      explanation: "Line Sparkline menampilkan tren data sebagai garis yang naik dan turun sesuai nilai aktual, cocok untuk melihat perubahan dari waktu ke waktu. Win/Loss Sparkline menyederhanakan data menjadi hanya dua kondisi — positif (batang ke atas) atau negatif (batang ke bawah) — cocok untuk hasil lulus/gagal atau menang/kalah."
    }
  ],
  sampleDatasets: [
    {
      id: "dataset-products-ch3",
      name: "Data Produk & Penjualan (Table Styles & Total Row)",
      filename: "latihan_produk_chapter3.csv",
      description: "Data tabel produk dengan kategori untuk latihan membuat Excel Table, menerapkan Table Styles, dan menggunakan Total Row.",
      csvContent: `ID Produk,Nama Produk,Kategori,Harga Satuan,Qty Terjual,Total Penjualan
PRD-01,Flashdisk 32GB,Storage,65000,45,2925000
PRD-02,SSD 256GB,Storage,350000,18,6300000
PRD-03,Kabel HDMI 2m,Aksesoris,45000,120,5400000
PRD-04,Webcam Logitech,Peripheral,580000,25,14500000
PRD-05,Adapter USB-C,Aksesoris,85000,60,5100000
PRD-06,RAM DDR4 8GB,Komponen,320000,30,9600000
PRD-07,Mouse Gaming,Peripheral,250000,85,21250000
PRD-08,Monitor 24 inch,Display,3200000,20,64000000`
    },
    {
      id: "dataset-sales-region-ch3",
      name: "Data Penjualan per Region (Multi-Level Subtotals)",
      filename: "latihan_subtotal_region.csv",
      description: "Data penjualan dengan kolom Region dan Periode untuk latihan Subtotal bertingkat dan Outline Levels.",
      csvContent: `Region,Periode,Produk,Qty,Total Penjualan
Barat,Q1,Laptop,15,127500000
Barat,Q1,Printer,20,30000000
Barat,Q2,Laptop,18,153000000
Barat,Q2,Printer,12,18000000
Timur,Q1,Laptop,22,187000000
Timur,Q1,Printer,10,15000000
Timur,Q2,Laptop,25,212500000
Timur,Q2,Printer,15,22500000`
    },
    {
      id: "dataset-branch-sparkline",
      name: "Data Cabang Bulanan (Icon Sets, Data Bars & Sparklines)",
      filename: "latihan_sparkline_cabang.csv",
      description: "Data performa cabang per bulan untuk latihan Icon Sets, Data Bars, dan pembuatan Sparklines.",
      csvContent: `Nama Cabang,Jan,Feb,Mar,Apr,Mei,Jun,Skor Semester
Jakarta,120,135,150,140,160,175,88
Bandung,95,100,90,110,105,120,72
Surabaya,140,130,155,165,170,180,92
Semarang,80,85,75,90,95,100,65
Yogyakarta,110,115,120,125,130,145,80`
    }
  ]
};
