/**
 * DATA & KURIKULUM PEMBELAJARAN ICT EXCEL - CHAPTER 5
 * Charts, Sparklines, and Visual Objects
 * Versi Siswa & Guru: Modul Praktik Mandiri, Visualisasi Data, dan Tutorial Interaktif Langkah demi Langkah
 */

const CHAPTER_5_DATA = {
  chapterNumber: 5,
  chapterTitle: "Charts, Sparklines, and Visual Objects",
  chapterSubtitle: "Modul Praktik Mandiri & Tutorial Interaktif Visualisasi Data Microsoft Excel",
  overview: {
    title: "General Theory: Transformasi Data Numerik Menjadi Informasi Visual & Grafis Profesional",
    description: "Visualisasi data adalah jembatan vital yang mengubah deretan angka mentah pada lembar kerja menjadi cerita visual yang mudah dipahami oleh pengambil keputusan. Melalui Chapter 5, siswa akan menguasai teknik pembuatan grafik standar (Column, Bar, Line, Pie), pemilihan kolom terpisah (non-adjacent) dengan tombol Ctrl, penempatan grafik pada Chart Sheet khusus, pengaturan elemen grafik melalui tombol Chart Elements (+), pengubahan skema warna dan gaya visual, teknik manipulasi sumbu data (Switch Row/Column), pembuatan tren mini di dalam sel menggunakan Sparklines (Line, Column, Win/Loss), pengubahan tipe grafik dan Quick Layout, serta penambahan objek visual bebas seperti gambar ber-Alt Text dan Text Box untuk menyusun laporan dasbor eksekutif yang informatif.",
    coreCompetencies: [
      "Mampu membuat grafik dasar dari data terpilih melalui tab Insert, memilih kolom terpisah menggunakan tombol Ctrl, dan memindahkan grafik ke lembar khusus (Chart Sheet) menggunakan Move Chart",
      "Menguasai pengelolaan elemen grafik (Chart Title, Axis Titles, Data Labels, Gridlines, Legend) melalui tombol Chart Elements (+) serta menerapkan Chart Styles dan Chart Colors",
      "Memahami mekanisme Switch Row/Column untuk menukar peran Baris dan Kolom antara Kategori (Sumbu X) dan Seri Data (Legenda), serta memahami pembaharuan grafik otomatis saat data sumber berubah",
      "Mampu menyisipkan dan memformat Row-Level Sparklines (Line, Column, Win/Loss) di dalam sel untuk menampilkan tren visual tanpa memakan ruang lembar kerja",
      "Dapat mengubah tipe grafik (Change Chart Type), mengatur ukuran grafik secara presisi, menerapkan Quick Layout, dan memodifikasi Shape Styles pada area grafik",
      "Mampu menyisipkan gambar (Pictures), memberikan teks alternatif (Alt Text) untuk aksesibilitas, serta menempatkan Text Box melayang yang dapat dikustomisasi secara bebas"
    ]
  },
  topics: [
    {
      id: "topic-5-1",
      number: "5.1",
      title: "Basic Chart Creation",
      indonesianTitle: "Pembuatan Grafik Dasar & Pemilihan Kolom Terpisah (Ctrl)",
      summary: "Charts are created from selected worksheet data through Insert > Charts. Users can select non-adjacent columns by holding Ctrl when choosing the data. A chart can be moved within the worksheet or placed on a separate Chart Sheet using Chart Design > Move Chart.",
      shortcut: "Alt + F1 (Grafik Cepat di Sheet) / F11 (Grafik di Sheet Baru)",
      ribbonPath: "Insert → Charts → 2-D Column / Chart Design → Move Chart",
      theoryExplanation: `Grafik (**Chart**) adalah representasi visual dari data angka di lembar kerja Excel. Dengan grafik, perbandingan nilai, tren kenaikan/penurunan, dan proporsi data dapat dipahami hanya dalam hitungan detik.

### 1. Prinsip Pemilihan Data Sebelum Membuat Grafik
Sebelum mengklik ikon grafik pada tab **Insert**, kamu wajib memblok data terlebih dahulu:
* **Blok Data Standar (Sel Bersebelahan):** Blok header kolom dan baris data yang ingin ditampilkan (misal kolom Nama Cabang dan Omzet Penjualan).
* **Blok Kolom Terpisah (Non-Adjacent Selection):** Seringkali kita hanya ingin membandingkan Kolom A (Nama Produk) dan Kolom D (Total Laba), sementara Kolom B dan C (Harga Pokok & Qty) ingin dilewati.
  * **Caranya:** Blok Kolom A1:A8 terlebih dahulu ➔ Tekan dan tahan tombol **Ctrl** di keyboard ➔ Blok Kolom D1:D8 ➔ Lepas tombol **Ctrl**.
  * *Perhatian:* Jumlah baris yang diblok pada kedua kolom **wajib sama persis** agar grafik tidak keliru membaca kategori!

### 2. Pintasan Keyboard Tercepat Membuat Grafik
Excel menyediakan 2 tombol sakti untuk menyisipkan grafik seketika setelah data diblok:
* **Alt + F1:** Menghasilkan grafik batang (*Embedded Clustered Column*) langsung di samping tabel pada sheet yang sama.
* **F11:** Memindahkan grafik ke **lembar kerja khusus (Chart Sheet)** baru secara otomatis bernama *Chart1*.

### 3. Memindahkan Grafik (Move Chart)
Secara default, grafik akan diletakkan mengambang (*floating object*) di lembar kerja yang sama. Kamu dapat memindahkannya melalui tab **Chart Design → Move Chart**:
1. **New Sheet:** Menempatkan grafik di lembar kerja khusus berskala penuh (*Chart Sheet*). Sangat ideal untuk materi presentasi atau cetakan mandiri tanpa terganggu baris grid sel.
2. **Object In:** Menempatkan grafik sebagai objek pendamping tabel di sheet tertentu (misal: *Sheet1* atau sheet dasbor).`,
      keyPoints: [
        "**Blok Data Dahulu**: Excel membutuhkan referensi data dan label judul sebelum jenis grafik dapat dipilih.",
        "**Pintasan Ctrl (Kolom Terpisah)**: Tahan tombol Ctrl untuk memilih rentang kolom yang tidak bersebelahan tanpa mengikutsertakan kolom di tengahnya.",
        "**Pintasan Alt + F1**: Membuat grafik 2-D Column instan di lembar kerja aktif.",
        "**Pintasan F11**: Membuat grafik otomatis pada lembar grafik mandiri (*Chart Sheet*).",
        "**Fitur Move Chart**: Memindahkan posisi grafik antara objek sel (*Embedded*) atau lembar khusus (*Chart Sheet*)."
      ],
      proTips: "Saat memblok tabel untuk membuat grafik, selalu sertakan baris Judul Kolom (Header) dan kolom Label Kategori (seperti Nama Cabang/Bulan). Excel akan otomatis menggunakan header tersebut sebagai nama Legenda dan teks label sumbu horizontal tanpa perlu diketik manual!",
      commonMistakes: "Memblok baris Total/Grand Total bersamaan dengan data rincian. Jika baris Total ikut diblok, batang grafik Total akan menjulang sangat tinggi dan menenggelamkan batang data rincian lainnya sehingga skala grafik menjadi rusak.",
      exerciseTask: "Dari tabel penjualan cabang, gunakan tombol Ctrl untuk memilih kolom Nama Cabang (A3:A8) dan Total Omzet (D3:D8) dengan melewati kolom Qty dan Diskon, lalu buat grafik 2-D Clustered Column dan pindahkan ke Chart Sheet menggunakan Move Chart.",
      desktopTutorial: {
        title: "Tutorial Praktik Mandiri di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "5.1_GrafikDasar",
        sampleData: `No\tNama Cabang\tKategori Produk\tTotal Omzet (Rp)\n1\tJakarta Pusat\tElektronik\t85000000\n2\tBandung Kota\tFashion\t42000000\n3\tSurabaya Timur\tElektronik\t78000000\n4\tYogyakarta\tKuliner\t31000000\n5\tMedan Petisah\tElektronik\t64000000\n6\tSemarang Barat\tFashion\t39000000`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #4f46e5;width:100%;">
  <thead>
    <tr bgcolor="#4f46e5" style="background-color:#4f46e5;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #4338ca;padding:6px 10px;text-align:center;">No</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:left;">Nama Cabang (Kolom B)</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:left;">Kategori Produk (Kolom C)</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:right;">Total Omzet (Rp) (Kolom D)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">1</td><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;color:#4f46e5;">Jakarta Pusat</td><td style="border:1px solid #d4d4d4;padding:5px 14px;color:#64748b;">Elektronik</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#107c41;">85.000.000</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">2</td><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;color:#4f46e5;">Bandung Kota</td><td style="border:1px solid #d4d4d4;padding:5px 14px;color:#64748b;">Fashion</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#107c41;">42.000.000</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">3</td><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;color:#4f46e5;">Surabaya Timur</td><td style="border:1px solid #d4d4d4;padding:5px 14px;color:#64748b;">Elektronik</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#107c41;">78.000.000</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">4</td><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;color:#4f46e5;">Yogyakarta</td><td style="border:1px solid #d4d4d4;padding:5px 14px;color:#64748b;">Kuliner</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#107c41;">31.000.000</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">5</td><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;color:#4f46e5;">Medan Petisah</td><td style="border:1px solid #d4d4d4;padding:5px 14px;color:#64748b;">Elektronik</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#107c41;">64.000.000</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:center;">6</td><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;color:#4f46e5;">Semarang Barat</td><td style="border:1px solid #d4d4d4;padding:5px 14px;color:#64748b;">Fashion</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#107c41;">39.000.000</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka Microsoft Excel di komputer kamu, buat Sheet Baru bernama: `5.1_GrafikDasar`.",
          "Klik tombol **'📋 Salin Tabel Latihan'**, buka Excel, klik sel **A1** satu kali, lalu tekan **Ctrl + V**.",
          "**[Seleksi Kolom Terpisah]** Blok sel **B1:B7** (Nama Cabang) terlebih dahulu.",
          "Tekan dan **tahan tombol Ctrl** di keyboard, lalu blok rentang sel **D1:D7** (Total Omzet), kemudian lepaskan tombol Ctrl. *(Kolom C Kategori dilewati)*.",
          "**[Buat Grafik Cepat]** Tekan shortcut keyboard: **Alt + F1**. Grafik batang 2-D Column langsung muncul di samping tabelmu!",
          "*(Alternatif Tab Ribbon)*: Klik tab **Insert** ➔ grup **Charts** ➔ klik ikon **Insert Column or Bar Chart** ➔ pilih **2-D Clustered Column**.",
          "**[Pindahkan ke Chart Sheet]** Klik grafik tersebut ➔ buka tab **Chart Design** di paling kanan Ribbon ➔ klik tombol **Move Chart** (di pojok kanan atas).",
          "Pilih opsi radio button **'New sheet:'** dan beri nama `Grafik_Omzet_Cabang` lalu klik **OK**.",
          "Perhatikan: Grafik kini tampil dalam satu halaman lembar kerja khusus yang elegan dan bersih tanpa garis-garis grid sel!"
        ],
        expectedResult: "Grafik 2-D Clustered Column berhasil dibuat dari kolom non-adjacent (Nama Cabang & Omzet) dan dipindahkan dengan rapi ke lembar Chart Sheet tersendiri."
      }
    },
    {
      id: "topic-5-2",
      number: "5.2",
      title: "Chart Elements and Visual Styles",
      indonesianTitle: "Elemen Grafik (+) & Gaya Visual (Styles & Colors)",
      summary: "Charts contain several elements, including Chart Title, Axis Titles, Data Labels, and Legend. These elements can be added or edited using Chart Elements (+) or the Chart Design tab. Chart Styles provide ready-made visual designs, while Chart Colors allows users to change the color scheme of the data series.",
      shortcut: "Tombol [+] di Pojok Kanan Atas Grafik",
      ribbonPath: "Chart Design → Add Chart Element / Chart Styles / Change Colors",
      theoryExplanation: `Sebuah grafik profesional tidak hanya menampilkan batang atau garis, melainkan harus memiliki informasi pendukung yang jelas agar pembaca langsung memahami konteks data tanpa bertanya-tanya.

### 1. Komponen Utama Elemen Grafik (Chart Elements)
Ketika kamu mengklik grafik di Excel, akan muncul tombol mengambang berlambang tambah hijau (**+**) di sudut kanan atas grafik. Inilah tombol **Chart Elements**:
* **Chart Title (Judul Grafik):** Menjelaskan inti dari grafik (misal: *"Laporan Performa Penjualan Semester 1 2026"*).
* **Axis Titles (Judul Sumbu):**
  * *Primary Horizontal:* Memberi keterangan sumbu X (misal: *"Wilayah Cabang"*).
  * *Primary Vertical:* Memberi keterangan sumbu Y beserta satuan ukurnya (misal: *"Total Omzet (Juta Rupiah)"*).
* **Data Labels (Label Angka):** Menampilkan nilai angka pasti tepat di atas atau di dalam batang grafik, sehingga audiens tidak perlu menarik garis pandang ke sumbu Y.
* **Data Table (Tabel Data di Bawah Grafik):** Menyertakan rangkuman tabel mini di bawah area grafik.
* **Gridlines (Garis Kisi):** Garis bantu horizontal/vertikal untuk mempermudah perbandingan ketinggian.
* **Legend (Legenda/Keterangan Warna):** Menjelaskan representasi warna data (misal: Biru = Target, Oranye = Realisasi). Dapat diatur posisinya di Top, Bottom, Left, atau Right.

### 2. Mempercantik Tampilan: Chart Styles & Chart Colors
Excel menyediakan koleksi tema visual siap pakai tanpa perlu mendesain dari nol:
* **Chart Styles:** Terletak pada tab **Chart Design → Chart Styles**. Tersedia gaya bernuansa *Modern Dark/Slate*, gaya berbayang 3D lembut, garis bertepi halus, hingga latar belakang transparan.
* **Change Colors (Skema Palet Warna):**
  * *Colorful:* Membedakan setiap seri data dengan warna kontras (Biru, Oranye, Abu-abu, Kuning).
  * *Monochromatic:* Menggunakan gradasi satu warna utama (misal: Navy tua, Biru medium, Biru muda). Sangat elegan untuk laporan resmi instansi dan korporasi.`,
      keyPoints: [
        "**Tombol [+] Chart Elements**: Akses instan tercepat untuk mencentang atau menghilangkan elemen grafik.",
        "**Axis Titles**: Wajib dicantumkan untuk memperjelas unit pengukuran (apakah Rupiah, Persen, atau Lusin).",
        "**Data Labels**: Menampilkan nilai konkret langsung pada batang/titik grafik.",
        "**Chart Styles**: Koleksi preset visual profesional siap pakai dengan pencahayaan dan kontras seimbang.",
        "**Change Colors**: Memilih palet warna yang selaras dengan identitas brand atau laporan (Colorful vs Monochromatic)."
      ],
      proTips: "Jika kamu sudah mengaktifkan Data Labels (angka di ujung batang), pertimbangkan untuk menyembunyikan sumbu angka vertikal (Primary Vertical Axis) dan garis kisi (Gridlines) agar tampilan dasbor grafik terlihat bersih, modern, dan minimalis!",
      commonMistakes: "Memberi judul grafik yang terlalu umum seperti 'Grafik 1' atau membiarkan judul bertuliskan 'Chart Title'. Selalu ganti judul grafik dengan pesan kesimpulan bisnis yang informatif.",
      exerciseTask: "Buka grafik penjualan, gunakan tombol Chart Elements (+) untuk menambahkan Judul Grafik, Judul Sumbu Horizontal dan Vertikal, aktifkan Data Labels di bagian Outside End, serta pilih Chart Style bernuansa Dark/Modern dengan palet warna Monochromatic.",
      desktopTutorial: {
        title: "Tutorial Praktik Mandiri di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "5.2_ElemenGrafik",
        sampleData: `Bulan\tTarget Penjualan (Juta)\tRealisasi Penjualan (Juta)\nJanuari\t50\t55\nFebruari\t60\t58\nMaret\t65\t72\nApril\t70\t68\nMei\t75\t82\nJuni\t80\t85`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #4f46e5;width:100%;">
  <thead>
    <tr bgcolor="#4f46e5" style="background-color:#4f46e5;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #4338ca;padding:6px 12px;text-align:left;">Bulan</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:right;">Target Penjualan (Juta Rp)</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:right;">Realisasi Penjualan (Juta Rp)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">Januari</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">50</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#107c41;">55</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">Februari</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">60</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#e11d48;">58</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">Maret</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">65</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#107c41;">72</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">April</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">70</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#e11d48;">68</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">Mei</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">75</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#107c41;">82</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">Juni</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">80</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;font-weight:bold;color:#107c41;">85</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka Microsoft Excel, buat Sheet Baru bernama: `5.2_ElemenGrafik`.",
          "Salin data latihan di atas (klik tombol **'📋 Salin Tabel Latihan'**), lalu paste di sel **A1**.",
          "Blok seluruh tabel dari sel **A1 sampai C7**.",
          "Buka tab **Insert** ➔ grup **Charts** ➔ pilih **2-D Clustered Column**.",
          "**[Gunakan Tombol Elemen +]** Klik grafik, lalu klik ikon tanda tambah hijau (**+**) di sudut kanan atas grafik.",
          "Centang kotak **Axis Titles** ➔ ketik `Periode Bulan` pada judul sumbu bawah, dan ketik `Nominal (Juta Rp)` pada judul sumbu kiri.",
          "Klik panah kecil (▶) di samping **Data Labels** ➔ pilih opsi **Outside End** agar angka penjualan tampil tepat di atas puncak batang.",
          "Klik dua kali pada teks **Chart Title**, ganti judulnya menjadi: `Evaluasi Target vs Realisasi Semester 1`.",
          "**[Terapkan Gaya & Warna]** Buka tab **Chart Design** ➔ pilih salah satu **Chart Style** yang memiliki latar gelap/elegan (misal Style 9 atau 8).",
          "Klik menu **Change Colors** di tab Chart Design ➔ pilih palet **Monochromatic Palette 1** untuk nuansa biru profesional."
        ],
        expectedResult: "Grafik memiliki judul deskriptif, judul sumbu yang informatif, label angka di setiap batang, serta gaya visual modern yang siap dipresentasikan."
      }
    },
    {
      id: "topic-5-3",
      number: "5.3",
      title: "Switch Row/Column and Edit Source Data",
      indonesianTitle: "Tukar Sumbu Data (Switch Row/Column) & Pembaruan Otomatis",
      summary: "Switch Row/Column changes how Excel organizes the selected data in a chart. It can switch rows and columns between categories and data series without changing the original data. When source data changes, the chart updates automatically. The chart's data range can also be expanded through Chart Design > Select Data.",
      shortcut: "Chart Design → Switch Row/Column",
      ribbonPath: "Chart Design → Data → Switch Row/Column / Select Data",
      theoryExplanation: `Seringkali setelah grafik terbuat, sudut pandang analisis yang kita inginkan belum sesuai: Apakah kita ingin membandingkan perkembangan antar kuartal dari setiap cabang, ATAU membandingkan performa seluruh cabang di setiap kuartal? Fitur **Switch Row/Column** menjawab kebutuhan ini tanpa mengubah susunan tabel asli!

### 1. Cara Kerja Fitur Switch Row/Column
Tombol **Switch Row/Column** berada di tab **Chart Design → grup Data**:
* **Kondisi Awal (Berdasarkan Kolom):** Sumbu Horizontal (X) menampilkan nama Cabang, sedangkan Legenda warna mewakili Kuartal (Q1, Q2, Q3, Q4). Kita melihat perbandingan kuartal per masing-masing cabang.
* **Setelah Diklik (Berdasarkan Baris):** Sumbu Horizontal (X) seketika berganti menjadi Kuartal (Q1, Q2, Q3, Q4), sedangkan Legenda warna mewakili nama Cabang. Sudut pandang analisis kini berubah menjadi kompetisi antar cabang dari waktu ke waktu!
* **Keunggulan:** Tabel angka di lembar kerjamu sama sekali tidak berubah posisinya. Hanya interpretasi grafik yang ditukar.

### 2. Sifat Dinamis & Real-Time Grafik Excel
Grafik di Excel terhubung secara langsung (*live-linked*) dengan sel sumber datanya:
* Jika kamu mengubah angka penjualan di tabel (misal sel C3 dari 45 juta menjadi 90 juta), batang grafik akan **langsung naik secara real-time** seketika kamu menekan tombol Enter!
* Tidak perlu membuat ulang grafik dari awal hanya karena ada revisi data.

### 3. Memperluas Rentang Data (Select Data Source)
Jika bisnismu menambah kuartal baru (Q4) atau cabang baru di luar area yang diblok sebelumnya:
1. Klik grafik ➔ buka tab **Chart Design** ➔ klik tombol **Select Data**.
2. Pada kotak dialog *Select Data Source*, kamu bisa memperbarui kotak **Chart data range** dengan memblok ulang tabel baru, atau menambah seri data secara spesifik.
3. Garis batas berwarna (*colored bounding box*) di lembar kerja juga bisa ditarik sudutnya menggunakan mouse untuk memperluas rentang secara visual!`,
      keyPoints: [
        "**Switch Row/Column**: Menukar posisi Kategori (Sumbu X) dan Seri Data (Legenda) secara instan tanpa mengacak tabel asli.",
        "**Koneksi Real-Time**: Perubahan nilai pada sel tabel sumber seketika memperbarui tinggi/posisi elemen grafik secara otomatis.",
        "**Dialog Select Data**: Mengelola penambahan seri data baru, pengeditan label sumbu, atau perluasan rentang sel sumber.",
        "**Bounding Box Interaktif**: Rentang data pada tabel ditandai kotak warna yang dapat ditarik untuk menambah baris/kolom baru ke grafik."
      ],
      proTips: "Jika tombol Switch Row/Column terlihat abu-abu (dinonaktifkan), pastikan grafik yang kamu pilih sedang dalam keadaan aktif dan bukan tipe grafik tertentu yang hanya memiliki satu seri data (seperti grafik Pie standar).",
      commonMistakes: "Mengetik ulang seluruh data tabel secara manual dalam format terbalik (transpose) hanya untuk mengubah tampilan grafik, padahal cukup satu kali klik tombol Switch Row/Column!",
      exerciseTask: "Buat grafik kolom dari data penjualan multi-kuartal, gunakan tombol Switch Row/Column untuk mengamati perbedaan perspektif analisis, lalu ubah salah satu nilai sel untuk membuktikan keandalan auto-update grafik Excel.",
      desktopTutorial: {
        title: "Tutorial Praktik Mandiri di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "5.3_SwitchData",
        sampleData: `Cabang\tKuartal 1\tKuartal 2\tKuartal 3\tKuartal 4\nJakarta\t120\t145\t160\t190\nBandung\t85\t92\t110\t125\nSurabaya\t95\t115\t130\t150\nMedan\t70\t80\t85\t95`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #4f46e5;width:100%;">
  <thead>
    <tr bgcolor="#4f46e5" style="background-color:#4f46e5;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:left;">Cabang</th>
      <th style="border:1px solid #4338ca;padding:6px 12px;text-align:right;">Kuartal 1 (Q1)</th>
      <th style="border:1px solid #4338ca;padding:6px 12px;text-align:right;">Kuartal 2 (Q2)</th>
      <th style="border:1px solid #4338ca;padding:6px 12px;text-align:right;">Kuartal 3 (Q3)</th>
      <th style="border:1px solid #4338ca;padding:6px 12px;text-align:right;">Kuartal 4 (Q4)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Jakarta</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">120</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">145</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">160</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;font-weight:bold;color:#4f46e5;">190</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Bandung</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">85</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">92</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">110</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;font-weight:bold;color:#4f46e5;">125</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Surabaya</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">95</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">115</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">130</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;font-weight:bold;color:#4f46e5;">150</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Medan</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">70</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">80</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;">85</td><td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:right;font-weight:bold;color:#4f46e5;">95</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka Microsoft Excel, buat Sheet Baru bernama: `5.3_SwitchData`.",
          "Salin data tabel di atas (klik tombol **'📋 Salin Tabel Latihan'**), lalu paste di sel **A1**.",
          "Blok seluruh data tabel dari sel **A1 sampai E5**.",
          "Buka tab **Insert** ➔ grup **Charts** ➔ pilih **2-D Clustered Column**.",
          "Perhatikan tampilan awal: Sumbu X menampilkan nama-nama Cabang (Jakarta, Bandung, Surabaya, Medan), dan Legenda warna mewakili Kuartal.",
          "**[Gunakan Switch Row/Column]** Klik grafik ➔ buka tab **Chart Design** ➔ di grup Data, klik tombol **Switch Row/Column**.",
          "Amati transformasinya: Sumbu X kini berganti menampilkan urutan Kuartal (Q1, Q2, Q3, Q4), sedangkan Legenda warna mewakili nama-nama Cabang!",
          "**[Uji Live Auto-Update]** Klik sel **B2** (penjualan Jakarta Q1 yang awalnya 120), ganti nilainya menjadi `250`, lalu tekan **Enter**.",
          "Perhatikan batang grafik Jakarta pada Kuartal 1 langsung melonjak naik secara seketika mengikuti angka baru tanpa perlu refresh!"
        ],
        expectedResult: "Siswa memahami bagaimana sudut pandang analisis dapat ditukar seketika dengan Switch Row/Column dan membuktikan grafik terhubung dinamis dengan data sumber."
      }
    },
    {
      id: "topic-5-4",
      number: "5.4",
      title: "Row-Level Sparklines",
      indonesianTitle: "Sparklines Tingkat Baris (Line, Column, & Win/Loss)",
      summary: "Sparklines are small charts placed directly inside cells to show data trends. Column Sparklines are useful for comparing values across periods, while Win/Loss Sparklines show positive or negative results without focusing on the actual value size.",
      shortcut: "Insert → Sparklines (Line / Column / Win/Loss)",
      ribbonPath: "Insert → Sparklines → Line / Column / Win/Loss",
      theoryExplanation: `Terkadang kita memiliki tabel berisi puluhan atau ratusan baris data (misalnya daftar 100 produk atau 50 siswa). Jika kita membuat grafik batang standar untuk setiap baris, lembar kerja akan menjadi sangat penuh dan kacau. Solusi cerdas Microsoft Excel untuk kasus ini adalah **Sparklines**!

### 1. Apa Itu Sparklines?
**Sparklines** adalah grafik mini berukuran sebesar sel tunggal yang diletakkan tepat di baris data yang bersangkutan. Sparklines berfungsi menampilkan pola tren (*trendline*), fluktuasi, atau perbandingan nilai dalam sekejap tanpa memerlukan sumbu, judul, maupun legenda.

### 2. Tiga Jenis Sparklines di Excel
Terdapat pada tab **Insert → grup Sparklines**:
1. **Line Sparkline (Garis):** Sangat efektif untuk melihat tren kontinyu dari waktu ke waktu (misal: pergerakan harga saham harian atau tren nilai bulanan siswa).
2. **Column Sparkline (Kolom/Batang Mini):** Sangat ideal untuk membandingkan nominal antar periode diskrit (misal: omzet penjualan kuartal 1 s.d. kuartal 4).
3. **Win/Loss Sparkline (Menang/Kalah):**
   * Jenis sparkline khusus yang menampilkan hasil **Positif vs Negatif** secara biner.
   * Batang yang mengarah ke atas (biasanya biru/hijau) menunjukkan nilai positif (untung/laba).
   * Batang yang mengarah ke bawah (biasanya merah) menunjukkan nilai negatif (rugi/penurunan).
   * *Keunikan:* Win/Loss mengabaikan besaran nominal angka dan murni berfokus pada status untung atau rugi.

### 3. Menandai Titik Kritis (Marker Points)
Setelah Sparklines dibuat, tab khusus **Sparkline** akan muncul di Ribbon. Kamu dapat mencentang penanda penting:
* **High Point:** Memberi warna khusus pada nilai puncak tertinggi.
* **Low Point:** Memberi warna khusus pada titik terendah.
* **Negative Points:** Memberi warna merah mencolok pada nilai yang bernilai minus.
* **First / Last Point:** Menandai titik awal dan titik akhir periode.`,
      keyPoints: [
        "**Grafik Mikro di Dalam Sel**: Sparkline hidup di dalam satu sel dan ukurannya otomatis mengikuti lebar kolom dan tinggi baris.",
        "**Line Sparkline**: Memvisualisasikan fluktuasi tren berkelanjutan.",
        "**Column Sparkline**: Menampilkan perbandingan volume antar periode waktu.",
        "**Win/Loss Sparkline**: Menyoroti status biner positif (laba) vs negatif (rugi) tanpa terpengaruh perbedaan skala angka.",
        "**Penanda Khusus (High/Low Points)**: Memudahkan mata mendeteksi rekor tertinggi dan titik terlemah secara instan."
      ],
      proTips: "Untuk menghapus Sparklines, menekan tombol Delete di keyboard TIDAK AKAN berhasil karena Sparkline adalah objek visual sel! Caranya: klik sel yang memuat Sparkline ➔ buka tab **Sparkline** di Ribbon ➔ klik tombol **Clear** (ikon penghapus).",
      commonMistakes: "Memasukkan sel Total ke dalam Data Range Sparkline. Sama seperti grafik biasa, jika angka total ikut dimasukkan ke dalam deret sparkline, titik terakhir akan melonjak drastis dan merusak representasi tren periode sebenarnya.",
      exerciseTask: "Buat kolom Sparkline Tren (Line) dan Sparkline Status Laba/Rugi (Win/Loss) untuk menganalisis performa 5 produk selama 6 bulan terakhir, serta tandai High Point dan Negative Points dengan warna kontras.",
      desktopTutorial: {
        title: "Tutorial Praktik Mandiri di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "5.4_Sparklines",
        sampleData: `Produk\tJan\tFeb\tMar\tApr\tMei\tJun\tTren Nilai (Line)\tStatus Laba (Win/Loss)\nLaptop Pro\t15\t18\t22\t19\t25\t30\t\t\nMonitor 4K\t10\t12\t-4\t8\t14\t11\t\t\nKeyboard RGB\t25\t20\t18\t15\t12\t10\t\t\nMouse Gaming\t30\t-5\t35\t40\t-2\t45\t\t\nHeadset USB\t8\t10\t12\t14\t16\t18\t\t`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #4f46e5;width:100%;">
  <thead>
    <tr bgcolor="#4f46e5" style="background-color:#4f46e5;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #4338ca;padding:6px 12px;text-align:left;">Produk</th>
      <th style="border:1px solid #4338ca;padding:6px 8px;text-align:right;">Jan</th>
      <th style="border:1px solid #4338ca;padding:6px 8px;text-align:right;">Feb</th>
      <th style="border:1px solid #4338ca;padding:6px 8px;text-align:right;">Mar</th>
      <th style="border:1px solid #4338ca;padding:6px 8px;text-align:right;">Apr</th>
      <th style="border:1px solid #4338ca;padding:6px 8px;text-align:right;">Mei</th>
      <th style="border:1px solid #4338ca;padding:6px 8px;text-align:right;">Jun</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:center;background-color:#3730a3;">Tren Nilai (Line)</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:center;background-color:#3730a3;">Status Laba (Win/Loss)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">Laptop Pro</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">15</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">18</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">22</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">19</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">25</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">30</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:center;color:#4f46e5;font-weight:bold;">📈 [Sparkline]</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:center;color:#107c41;font-weight:bold;">⚖️ [Win/Loss]</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">Monitor 4K</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">10</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">12</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;color:#e11d48;font-weight:bold;">-4</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">8</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">14</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">11</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:center;color:#4f46e5;font-weight:bold;">📈 [Sparkline]</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:center;color:#e11d48;font-weight:bold;">⚖️ [Win/Loss]</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">Keyboard RGB</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">25</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">20</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">18</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">15</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">12</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">10</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:center;color:#4f46e5;font-weight:bold;">📈 [Sparkline]</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:center;color:#107c41;font-weight:bold;">⚖️ [Win/Loss]</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">Mouse Gaming</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">30</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;color:#e11d48;font-weight:bold;">-5</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">35</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">40</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;color:#e11d48;font-weight:bold;">-2</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">45</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:center;color:#4f46e5;font-weight:bold;">📈 [Sparkline]</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:center;color:#e11d48;font-weight:bold;">⚖️ [Win/Loss]</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 12px;font-weight:bold;">Headset USB</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">8</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">10</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">12</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">14</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">16</td><td style="border:1px solid #d4d4d4;padding:5px 8px;text-align:right;">18</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:center;color:#4f46e5;font-weight:bold;">📈 [Sparkline]</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:center;color:#107c41;font-weight:bold;">⚖️ [Win/Loss]</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka Microsoft Excel, buat Sheet Baru bernama: `5.4_Sparklines`.",
          "Salin data tabel di atas (klik tombol **'📋 Salin Tabel Latihan'**), lalu paste di sel **A1**.",
          "**[Buat Line Sparkline]** Klik sel **H2** (kolom Tren Nilai untuk Laptop Pro).",
          "Buka tab **Insert** ➔ di grup Sparklines, klik **Line**.",
          "Pada kotak dialog Create Sparklines: isi **Data Range** dengan `B2:G2`, dan **Location Range** otomatis terisi `$H$2`. Klik **OK**.",
          "Gunakan fitur Fill Down (tarik kotak kecil hijau di sudut sel H2) ke bawah sampai sel **H6**.",
          "**[Kustomisasi Titik Kritis]** Di tab **Sparkline** yang muncul di Ribbon, centang opsi **High Point** dan **Low Point**.",
          "**[Buat Win/Loss Sparkline]** Klik sel **I2** (kolom Status Laba).",
          "Buka tab **Insert** ➔ grup Sparklines ➔ klik **Win/Loss**.",
          "Isi Data Range dengan `B2:G2` dan klik **OK**, lalu tarik ke bawah sampai sel **I6**.",
          "Di tab Sparkline, centang **Negative Points** ➔ klik **Marker Color** ➔ **Negative Points** ➔ pilih warna Merah Terang.",
          "Perhatikan: Pada Monitor 4K dan Mouse Gaming yang memiliki angka minus (-4, -5, -2), batang Win/Loss seketika mengarah ke bawah dengan warna merah!"
        ],
        expectedResult: "Tabel dilengkapi grafik mikro Line dengan titik rekor tertinggi/terendah dan grafik biner Win/Loss yang menyoroti angka defisit secara kontras."
      }
    },
    {
      id: "topic-5-5",
      number: "5.5",
      title: "Changing Chart Type, Size, and Layout",
      indonesianTitle: "Mengubah Tipe Grafik, Ukuran, & Quick Layout",
      summary: "A chart type can be changed through Chart Design > Change Chart Type without creating a new chart. The chart size can be adjusted by dragging its corners. Quick Layout changes the arrangement of chart elements, while Shape Style controls the appearance of the chart area, such as borders, background colors, and effects.",
      shortcut: "Chart Design → Change Chart Type / Alt + Drag (Snap to Grid)",
      ribbonPath: "Chart Design → Change Chart Type / Quick Layout / Format → Size",
      theoryExplanation: `Seringkali setelah sebuah grafik selesai dibuat, format awal (misalnya Clustered Column) kurang efektif dalam menyampaikan pesan yang diinginkan. Daripada menghapus dan membuat ulang grafik, Excel memungkinkan kita mengubah tipe grafik secara instan serta menata tata letak elemennya dalam hitungan detik.

### 1. Mengubah Tipe Grafik (Change Chart Type)
Terdapat pada tab **Chart Design → Change Chart Type** (atau klik kanan pada grafik ➔ pilih *Change Chart Type...*):
* **Column ➔ Line:** Mengubah diagram batang menjadi grafik garis. Sangat ideal jika data memiliki rentang waktu yang rapat dan ingin menonjolkan laju pertumbuhan.
* **Bar Chart (Batang Horizontal):** Pilihan tepat jika label nama kategori sangat panjang (seperti nama jabatan atau nama produk yang kompleks) agar tidak bertabrakan secara vertikal.
* **Combo Chart (Grafik Kombinasi):**
  * Menggabungkan 2 tipe grafik sekaligus (misalnya batang untuk *Total Penjualan* dan garis untuk *Persentase Pertumbuhan*).
  * **Secondary Axis (Sumbu Sekunder):** Mengaktifkan sumbu Y kedua di sebelah kanan grafik agar data bernilai puluhan juta rupiah dan data bernilai persen (0% s.d. 100%) dapat ditampilkan bersamaan secara proporsional tanpa saling menenggelamkan!

### 2. Mengatur Ukuran Grafik & Trik Snap to Grid (Alt + Drag)
* **Mengubah Ukuran Manual:** Arahkan kursor mouse ke lingkaran kecil di sudut atau tepi bingkai grafik (*sizing handles*), lalu tarik sesuai keinginan.
* **Ukuran Presisi:** Buka tab **Format ➔ grup Size** di pojok kanan Ribbon untuk mengetikkan tinggi (*Height*) dan lebar (*Width*) pasti dalam satuan sentimeter.
* **Trik Sakti Snap to Grid (Tekan Alt):** Tahan tombol **Alt** saat menggeser atau mengubah ukuran grafik dengan mouse. Bingkai grafik akan otomatis **menempel rapi (*snap*) persis pada garis kisi sel Excel**, menghasilkan dasbor yang sejajar dan simetris secara sempurna!

### 3. Quick Layout & Shape Styles
* **Quick Layout:** Menu siap pakai di tab **Chart Design** yang mengubah susunan posisi elemen grafik (letak legenda, tabel data terpasang di bawah, label di tengah batang) hanya dalam satu klik.
* **Shape Styles (Format Area Grafik):** Tab **Format** memungkinkan kamu memberi warna latar belakang khusus (*Shape Fill*), bingkai tepi melengkung (*Shape Outline*), dan efek bayangan modern (*Shape Effects / Soft Shadow*).`,
      keyPoints: [
        "**Change Chart Type**: Mengonversi tipe grafik seketika tanpa kehilangan pengaturan data sumber.",
        "**Combo Chart & Secondary Axis**: Menggabungkan batang dan garis untuk membandingkan dua metrik dengan satuan berbeda (misal: Nominal Rupiah vs Persen).",
        "**Pintasan Alt + Drag (Snap to Grid)**: Menyelaraskan batas tepi grafik tepat di garis kisi sel agar dasbor terlihat rapi presisi.",
        "**Quick Layout**: Mengatur komposisi letak judul, legenda, dan label dalam berbagai preset standar.",
        "**Shape Formatting**: Memberikan sentuhan estetika pada bingkai grafik dengan efek fill dan border elegan."
      ],
      proTips: "Saat membuat Combo Chart antara Penjualan (Rupiah) dan Rasio Profit (%), selalu centang kotak 'Secondary Axis' pada seri Rasio Profit. Tanpa sumbu sekunder, garis persentase (misal 15% atau 0.15) akan tampak menempel datar di garis angka nol karena skala sumbu primer bernilai jutaan!",
      commonMistakes: "Menggunakan grafik 3D yang berlebihan. Grafik 3D seperti 3D Pie atau 3D Column seringkali mendistorsi sudut pandang dan membuat perbandingan nilai sulit diukur secara akurat oleh mata manusia.",
      exerciseTask: "Ubah grafik batang penjualan menjadi Combo Chart (Volume Penjualan berbentuk Clustered Column dan Tingkat Keuntungan berbentuk Line pada Secondary Axis), rapikan posisi grafik menggunakan trik Alt + Drag agar sejajar dengan sel, dan terapkan Quick Layout.",
      desktopTutorial: {
        title: "Tutorial Praktik Mandiri di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "5.5_TipeGrafik",
        sampleData: `Bulan\tUnit Terjual\tTingkat Laba (%)\nJanuari\t450\t18%\nFebruari\t520\t22%\nMaret\t480\t19%\nApril\t610\t25%\nMei\t700\t28%\nJuni\t750\t30%`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #4f46e5;width:100%;">
  <thead>
    <tr bgcolor="#4f46e5" style="background-color:#4f46e5;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:left;">Bulan</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:right;">Unit Terjual (Volume)</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:right;">Tingkat Laba (%)</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Januari</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">450</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;color:#107c41;font-weight:bold;">18%</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Februari</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">520</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;color:#107c41;font-weight:bold;">22%</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Maret</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">480</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;color:#107c41;font-weight:bold;">19%</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">April</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">610</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;color:#107c41;font-weight:bold;">25%</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Mei</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">700</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;color:#107c41;font-weight:bold;">28%</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Juni</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;">750</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;color:#107c41;font-weight:bold;">30%</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka Microsoft Excel, buat Sheet Baru bernama: `5.5_TipeGrafik`.",
          "Salin data tabel di atas (klik tombol **'📋 Salin Tabel Latihan'**), lalu paste di sel **A1**.",
          "Blok seluruh tabel dari sel **A1 sampai C7**.",
          "Buat grafik standar: tab **Insert** ➔ grup Charts ➔ **2-D Clustered Column**.",
          "Perhatikan: Batang 'Tingkat Laba (%)' tampak sangat pipih di dasar grafik karena angkanya berupa desimal (0.18 - 0.30) sedangkan unit terjual bernilai ratusan.",
          "**[Ubah Menjadi Combo Chart]** Klik grafik ➔ buka tab **Chart Design** ➔ klik **Change Chart Type**.",
          "Di panel sebelah kiri, pilih kategori paling bawah: **Combo**.",
          "Atur seri data: Seri 'Unit Terjual' tetap pilih **Clustered Column**, sedangkan seri 'Tingkat Laba (%)' ubah menjadi **Line** dan **CENTANG kotak 'Secondary Axis'** di sampingnya!",
          "Klik **OK**. Perhatikan: Sumbu kanan kini muncul khusus untuk persentase (0% s.d. 35%), dan garis laba tampil melayang secara proporsional di atas batang penjualan!",
          "**[Trik Alt + Drag (Snap to Grid)]** Tahan tombol **Alt** di keyboard, lalu tarik sudut kanan bawah grafik agar menempel persis di sudut sel **I16**.",
          "Buka menu **Quick Layout** pada tab Chart Design ➔ pilih **Layout 2** atau **Layout 5** untuk melihat tata letak alternatif."
        ],
        expectedResult: "Tercipta grafik kombinasi (Combo Chart) profesional dengan sumbu sekunder yang seimbang, berukuran presisi menempel garis kisi sel (Snap to Grid)."
      }
    },
    {
      id: "topic-5-6",
      number: "5.6",
      title: "Pictures and Text Boxes",
      indonesianTitle: "Penyisipan Gambar (Alt Text) & Kotak Teks Melayang (Text Box)",
      summary: "Excel worksheets can also contain pictures and text boxes through the Insert tab. Pictures can be positioned freely and provided with Alt Text for accessibility. Text Boxes allow text to be placed anywhere on the worksheet independently of cells. Their size, font, color, and alignment can be customized as needed.",
      shortcut: "Insert → Illustrations → Pictures / Insert → Text → Text Box",
      ribbonPath: "Insert → Illustrations → Pictures / Insert → Text → Text Box",
      theoryExplanation: `Sebuah lembar kerja Excel modern yang difungsikan sebagai Laporan Dasbor Manajemen (*Executive Dashboard*) tidak hanya memuat kisi tabel dan grafik, melainkan membutuhkan elemen ilustrasi pendukung seperti logo institusi, ikon penjelas, kartu ringkasan KPI (*callout card*), dan kotak catatan eksekutif (*Text Box*).

### 1. Menyisipkan & Memformat Gambar (Pictures)
Gambar dapat disisipkan melalui tab **Insert → Illustrations → Pictures**:
* **Sumber Gambar:** *This Device* (file di komputer kamu), *Stock Images*, atau *Online Pictures*.
* **Penempatan Bebas:** Gambar adalah objek mengambang (*floating object*) yang dapat digeser, diputar (*rotate*), diubah ukuran skalanya secara bebas, dan ditumpuk (*layering*) di atas atau di samping grafik.
* **Format Picture:** Tab **Picture Format** menyediakan fitur pemotongan (*Crop*), pembingkaian gaya foto (*Picture Styles*), koreksi kecerahan (*Corrections*), dan transparansi.
* **Alt Text (Teks Alternatif untuk Aksesibilitas):**
  * Fitur penting berstandar internasional: klik kanan pada gambar ➔ pilih **View Alt Text...**.
  * Tulis deskripsi singkat mengenai isi gambar (misal: *"Logo Resmi SMK Negeri 1 Jakarta"*).
  * Alt Text memungkinkan pembaca tunanetra yang menggunakan pembaca layar (*screen reader*) memahami konteks gambar, serta menjaga kepatuhan aksesibilitas dokumen resmi.

### 2. Kotak Teks Melayang (Text Box)
Seringkali kita ingin meletakkan catatan penting, kesimpulan rapat, atau label judul besar di posisi tertentu tanpa terikat oleh ukuran lebar kolom atau tinggi baris sel Excel. Solusinya adalah **Text Box**:
* **Cara Menyisipkan:** Tab **Insert ➔ grup Text ➔ Text Box** ➔ klik dan tarik (*drag*) kursor mouse di lembar kerja untuk menggambar kotak.
* **Keleluasaan Penuh:**
  * Ukuran teks, jenis font, perataan (*alignment*), dan warna teks dapat diatur persis seperti di Microsoft Word.
  * **Transparansi / No Fill:** Melalui tab **Shape Format → Shape Fill → No Fill** dan **Shape Outline → No Outline**, kotak teks akan menjadi transparan sehingga tulisan melayang rapi di atas grafik atau sel tanpa menutupi latar belakang.
  * **Koneksi Dinamis ke Nilai Sel:** Kamu bahkan bisa membuat Text Box dinamis: klik garis batas Text Box ➔ ketik di Formula Bar: \`=C2\` ➔ tekan Enter. Teks di dalam kotak akan otomatis menampilkan isi dari sel C2!`,
      keyPoints: [
        "**Objek Melayang Bebas**: Gambar dan Text Box berada di lapisan atas lembar kerja tanpa terhalang batas garis sel.",
        "**Pentingnya Alt Text**: Memberikan deskripsi teks pada gambar untuk standar aksesibilitas pembaca layar (*screen readers*).",
        "**Independensi Text Box**: Memungkinkan penulisan kesimpulan, instruksi tugas, atau callout KPI di posisi manapun secara presisi.",
        "**Shape Fill & Outline**: Text Box dapat diatur tanpa latar (*No Fill*) dan tanpa garis (*No Outline*) untuk menghasilkan tipografi melayang modern.",
        "**Formula Linking**: Text Box dapat disambungkan secara dinamis ke rumus sel Excel melalui Formula Bar."
      ],
      proTips: "Untuk mengunci posisi gambar atau Text Box agar tidak berantakan saat ada kolom yang disembunyikan atau diubah lebarnya: klik kanan objek ➔ Format Shape/Picture ➔ buka tab Size & Properties ➔ pilih opsi 'Don't move or size with cells'!",
      commonMistakes: "Menggabungkan (*merge*) puluhan sel tabel hanya untuk mengetikkan catatan panjang. Menggunakan Merge Cells berisiko merusak rumus dan operasi sorting data. Gunakan Text Box melayang yang jauh lebih aman dan fleksibel!",
      exerciseTask: "Sisipkan logo atau ilustrasi ke dalam lembar dasbor, lengkapi dengan Alt Text deskriptif, lalu tambahkan Text Box transparan berisi 'Ringkasan Eksekutif & Kesimpulan Performa Penjualan Semester 1' yang diposisikan rapi di samping grafik.",
      desktopTutorial: {
        title: "Tutorial Praktik Mandiri di Microsoft Excel Laptop/PC:",
        recommendedSheetName: "5.6_ObjekVisual",
        sampleData: `Metrik Utama\tNilai Capaian\tStatus Evaluasi\nTotal Omzet\tRp 520.000.000\tMelampaui Target (+15%)\nJumlah Pelanggan Baru\t1.420 Akun\tSesuai Rencana\nKepuasan Pelanggan (CSAT)\t94.8%\tSangat Baik (Bintang 5)\nTingkat Retensi Member\t88.2%\tPerlu Ditingkatkan di Q3`,
        sampleHtml: `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #4f46e5;width:100%;">
  <thead>
    <tr bgcolor="#4f46e5" style="background-color:#4f46e5;color:#ffffff;font-weight:bold;">
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:left;">Metrik Utama</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:right;">Nilai Capaian</th>
      <th style="border:1px solid #4338ca;padding:6px 14px;text-align:left;">Status Evaluasi</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Total Omzet</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;color:#107c41;font-weight:bold;">Rp 520.000.000</td><td style="border:1px solid #d4d4d4;padding:5px 14px;color:#107c41;font-weight:bold;">Melampaui Target (+15%)</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Jumlah Pelanggan Baru</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;color:#4f46e5;font-weight:bold;">1.420 Akun</td><td style="border:1px solid #d4d4d4;padding:5px 14px;">Sesuai Rencana</td></tr>
    <tr><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Kepuasan Pelanggan (CSAT)</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;color:#d97706;font-weight:bold;">94.8%</td><td style="border:1px solid #d4d4d4;padding:5px 14px;color:#d97706;font-weight:bold;">Sangat Baik (Bintang 5)</td></tr>
    <tr style="background-color:#f8fafc;"><td style="border:1px solid #d4d4d4;padding:5px 14px;font-weight:bold;">Tingkat Retensi Member</td><td style="border:1px solid #d4d4d4;padding:5px 14px;text-align:right;color:#e11d48;font-weight:bold;">88.2%</td><td style="border:1px solid #d4d4d4;padding:5px 14px;color:#e11d48;font-weight:bold;">Perlu Ditingkatkan di Q3</td></tr>
  </tbody>
</table>`,
        steps: [
          "Buka Microsoft Excel, buat Sheet Baru bernama: `5.6_ObjekVisual`.",
          "Salin data tabel di atas (klik tombol **'📋 Salin Tabel Latihan'**), lalu paste di sel **A1**.",
          "**[Sisipkan Gambar / Logo]** Buka tab **Insert** ➔ di grup Illustrations, klik **Pictures** ➔ pilih gambar apa saja di komputer kamu (atau pilih *Stock Images / Online Pictures*).",
          "Atur ukuran gambar dengan menarik sudut lingkarannya agar pas diletakkan di atas judul tabel (misal sel **E1**).",
          "**[Beri Alt Text]** Klik kanan pada gambar ➔ pilih menu **View Alt Text...** ➔ ketik deskripsi: `Logo Perusahaan Resmi Laporan Kinerja 2026` ➔ tutup panel Alt Text.",
          "**[Sisipkan Text Box Catatan Eksekutif]** Buka tab **Insert** ➔ di grup Text (sebelah kanan), klik **Text Box**.",
          "Arahkan kursor ke sel **E3**, klik dan tahan (*drag*) mouse hingga membentuk kotak persegi panjang dengan ukuran sekitar lebar 6 kolom dan tinggi 5 baris.",
          "Ketik catatan berikut di dalam Text Box: `📌 CATATAN MANAJEMEN: Omzet semester 1 sukses melampaui target sebesar 15%. Fokus strategi kuartal 3 adalah meningkatkan kembali retensi member dengan program loyalty khusus.`",
          "**[Format Tampilan Text Box]** Blok tulisan di dalam Text Box ➔ beri warna biru tua dan tebalkan judulnya.",
          "Klik garis tepi Text Box ➔ tab **Shape Format** ➔ pada **Shape Fill** pilih warna latar kuning lembut (atau *No Fill*) dan beri garis tepi tipis (*Shape Outline*)."
        ],
        expectedResult: "Tabel laporan bertransformasi menjadi dasbor eksekutif yang kaya visual dengan logo ber-Alt Text dan kartu catatan manajemen (Text Box) yang informatif."
      }
    }
  ],
  quizQuestions: [
    {
      id: 1,
      topicRef: "5.1",
      question: "Seorang siswa ingin membuat grafik perbandingan antara Kolom A (Nama Barang) dan Kolom D (Total Harga), namun ingin melewati Kolom B dan C. Tombol keyboard apa yang wajib ditekan dan ditahan saat memblok Kolom D?",
      options: [
        "Tombol Shift",
        "Tombol Alt",
        "Tombol Ctrl (Control)",
        "Tombol Tab"
      ],
      correctIndex: 2,
      explanation: "Tombol Ctrl (Control) digunakan untuk melakukan non-adjacent selection (memilih sel/kolom yang posisinya terpisah). Dengan menahan Ctrl, kita dapat memilih rentang A1:A8 lalu D1:D8 tanpa mengikutsertakan kolom B dan C."
    },
    {
      id: 2,
      topicRef: "5.1",
      question: "Apa fungsi dari tombol pintasan F11 setelah seorang pengguna memblok tabel data di Microsoft Excel?",
      options: [
        "Menyimpan file lembar kerja dalam format PDF secara otomatis",
        "Membuat grafik Clustered Column secara instan pada lembar kerja khusus baru (Chart Sheet)",
        "Menghapus seluruh angka nol di dalam tabel terpilih",
        "Membuka dialog pencarian fungsi matematika"
      ],
      correctIndex: 1,
      explanation: "Tombol F11 adalah shortcut bawaan Excel untuk membuat grafik kolom secara otomatis pada lembar grafik mandiri tersendiri yang disebut Chart Sheet (misal: Chart1). Untuk membuat grafik instan di lembar kerja yang sama, pintasannya adalah Alt + F1."
    },
    {
      id: 3,
      topicRef: "5.2",
      question: "Elemen grafik manakah yang paling tepat diaktifkan jika presenter ingin agar audiens langsung mengetahui angka nominal konkret di ujung setiap batang tanpa harus menarik garis mata ke sumbu angka vertikal?",
      options: [
        "Legend (Legenda)",
        "Gridlines (Garis Kisi)",
        "Data Labels (Label Data)",
        "Error Bars"
      ],
      correctIndex: 2,
      explanation: "Data Labels menampilkan nilai angka eksak dari data sumber secara langsung di samping, di dalam, atau di atas batang/titik data grafik, sehingga pembaca tidak perlu memperkirakan angka dari sumbu vertikal."
    },
    {
      id: 4,
      topicRef: "5.3",
      question: "Apa yang terjadi pada tampilan grafik saat tombol 'Switch Row/Column' pada tab Chart Design diklik?",
      options: [
        "Angka-angka di tabel lembar kerja asli ditukar posisinya secara permanen",
        "Sumbu Kategori (Sumbu X) dan Seri Data (Legenda) bertukar posisi dalam pengorganisasian grafik tanpa mengubah tabel aslinya",
        "Grafik batang secara otomatis terhapus dan digantikan grafik pie",
        "Semua nilai minus diubah menjadi nilai positif"
      ],
      correctIndex: 1,
      explanation: "Fitur Switch Row/Column menukar orientasi pengorganisasian data antara Baris dan Kolom pada grafik. Apa yang tadinya menjadi label sumbu horizontal kini menjadi legenda, dan sebaliknya, sementara data asli di worksheet tetap aman tanpa perubahan."
    },
    {
      id: 5,
      topicRef: "5.4",
      question: "Perusahaan retail memiliki data riwayat untung/rugi bulanan (+ dan -). Manajemen ingin melihat ringkasan visual mini di dalam sel yang secara khusus menonjolkan kapan terjadi laba dan kapan terjadi rugi tanpa terganggu perbedaan besar kecilnya nominal. Tipe Sparkline apa yang paling ideal?",
      options: [
        "Line Sparkline",
        "Column Sparkline",
        "Win/Loss Sparkline",
        "Pie Sparkline"
      ],
      correctIndex: 2,
      explanation: "Win/Loss Sparkline adalah grafik mini khusus yang hanya menampilkan hasil biner: batang ke atas untuk nilai positif (laba/untung) dan batang ke bawah untuk nilai negatif (rugi), dengan ukuran ketinggian batang yang sama rata tanpa memperdulikan perbedaan skala nominal."
    },
    {
      id: 6,
      topicRef: "5.6",
      question: "Mengapa memberikan teks alternatif (Alt Text) pada gambar atau objek visual di Microsoft Excel sangat penting menurut standar dokumen resmi?",
      options: [
        "Agar ukuran file Excel menjadi lebih kecil saat dikirim via email",
        "Membuat gambar otomatis bergerak saat presentasi dibuka",
        "Mendukung aksesibilitas agar orang dengan gangguan penglihatan yang menggunakan screen reader dapat memahami informasi di balik gambar tersebut",
        "Mencegah gambar tersebut disalin atau diedit oleh pengguna lain"
      ],
      correctIndex: 2,
      explanation: "Alt Text (Alternative Text) adalah standar aksesibilitas digital internasional. Teks deskriptif ini dibacakan oleh perangkat lunak screen reader kepada pengguna tunanetra, memastikan semua orang mendapatkan informasi yang utuh dari visual yang disematkan."
    }
  ],
  sampleDatasets: [
    {
      id: "dataset-branch-sales-ch5",
      name: "Data Penjualan Multi-Cabang & Multi-Kuartal (Grafik & Switch Row/Col)",
      filename: "latihan_grafik_penjualan_ch5.csv",
      description: "Data penjualan multi-wilayah dan multi-kuartal untuk latihan membuat grafik kolom 2D, non-adjacent selection, switch row/column, dan penyesuaian elemen grafik.",
      csvContent: `Wilayah Cabang,Kuartal 1,Kuartal 2,Kuartal 3,Kuartal 4,Total Tahunan
DKI Jakarta,120000000,145000000,160000000,190000000,615000000
Jawa Barat (Bandung),85000000,92000000,110000000,125000000,412000000
Jawa Timur (Surabaya),95000000,115000000,130000000,150000000,490000000
DI Yogyakarta,45000000,52000000,60000000,68000000,225000000
Sumatera Utara (Medan),70000000,80000000,85000000,95000000,330000000
Sulawesi Selatan (Makassar),55000000,62000000,74000000,82000000,273000000`
    },
    {
      id: "dataset-monthly-sparklines-ch5",
      name: "Data Tren Kinerja Bulanan & Untung-Rugi (Line & Win/Loss Sparklines)",
      filename: "latihan_sparklines_kinerja_ch5.csv",
      description: "Data fluktuasi kinerja operasional 6 bulan dengan nilai positif dan negatif untuk latihan membuat Line Sparklines dan Win/Loss Sparklines di dalam sel.",
      csvContent: `Kode Unit,Divisi Operasional,Bulan 1,Bulan 2,Bulan 3,Bulan 4,Bulan 5,Bulan 6
DIV-01,Divisi Hardware & Server,25,32,28,35,42,48
DIV-02,Divisi Cloud & Hosting,15,18,-5,12,20,16
DIV-03,Divisi Konsultasi IT,40,35,30,25,20,15
DIV-04,Divisi Cybersecurity,10,-8,15,22,-3,28
DIV-05,Divisi Software App,50,55,62,58,65,70
DIV-06,Divisi Helpdesk Support,12,14,16,18,20,22`
    },
    {
      id: "dataset-combo-dashboard-ch5",
      name: "Data Metrik Dasbor Target vs Realisasi (Combo Chart & Objek Visual)",
      filename: "latihan_combo_dashboard_ch5.csv",
      description: "Dataset komparasi volume penjualan dan rasio margin laba untuk latihan Combo Chart dengan Secondary Axis, Text Box KPI, dan gambar ber-Alt Text.",
      csvContent: `Periode Bulan,Volume Penjualan (Unit),Target Minimum,Rasio Margin Laba (%)
Januari 2026,450,400,18.5%
Februari 2026,520,450,22.0%
Maret 2026,480,450,19.2%
April 2026,610,500,25.4%
Mei 2026,700,550,28.1%
Juni 2026,750,600,30.5%`
    }
  ]
};

window.CHAPTER_5_DATA = CHAPTER_5_DATA;

