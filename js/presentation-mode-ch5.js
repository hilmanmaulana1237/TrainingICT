/**
 * TEACHER PRESENTATION MODE (SLIDESHOW & PROJECTOR SUITE) - CHAPTER 5
 * Optimized for classroom projectors, interactive smartboards, and ICT instruction
 * Built for Chapter 5: Charts, Sparklines, and Visual Objects
 */

const PresentationModeCh5 = {
  currentSlideIndex: 0,
  slides: [],
  laserActive: false,
  overlayEl: null,
  laserDotEl: null,

  init() {
    this.overlayEl = document.getElementById('presentation-overlay');
    this.laserDotEl = document.getElementById('laser-pointer-dot');
    this.buildSlideDeck();
    this.bindEvents();
  },

  buildSlideDeck() {
    this.slides = [];

    // Slide 0: Chapter Cover Slide
    this.slides.push({
      number: "COVER",
      badge: "CHAPTER 5 ICT",
      title: CHAPTER_5_DATA.chapterTitle,
      subtitle: CHAPTER_5_DATA.chapterSubtitle,
      bullets: [
        "5.1 Pembuatan Grafik Dasar (Insert Charts, Seleksi Terpisah Ctrl, Move Chart)",
        "5.2 Elemen Grafik & Gaya Visual (Tombol [+], Axis Titles, Data Labels, Styles & Colors)",
        "5.3 Switch Row/Column & Pembaruan Otomatis (Tukar Kategori & Legenda, Live Update)",
        "5.4 Row-Level Sparklines (Tren Mikro Sel: Line, Column, Win/Loss, High/Low Points)",
        "5.5 Mengubah Tipe Grafik & Ukuran (Change Chart Type, Combo Chart & Secondary Axis, Quick Layout)",
        "5.6 Objek Visual Bebas (Penyisipan Gambar ber-Alt Text & Text Box Melayang)"
      ],
      notes: "Bab 5 berfokus pada data storytelling: bagaimana menerjemahkan tabel data numerik mentah menjadi grafik yang jelas, estetik, dan mudah dipahami dalam pengambilan keputusan bisnis."
    });

    // Slide 1: General Theory
    this.slides.push({
      number: "THEORY",
      badge: "Pengantar Umum",
      title: CHAPTER_5_DATA.overview.title,
      subtitle: "Prinsip Visualisasi Data & Desain Dasbor Microsoft Excel",
      bullets: [
        "Grafik mengubah deretan angka membosankan menjadi pola visual yang dapat dianalisis dalam hitungan detik.",
        "Pemilihan data terpisah (non-adjacent) dilakukan dengan menahan tombol Ctrl tanpa mengubah posisi tabel.",
        "Fitur Switch Row/Column memungkinkan pergantian sudut pandang analisis tanpa merevisi data sumber.",
        "Sparklines menghadirkan grafik mikro di dalam sel untuk memantau ratusan tren baris data secara efisien.",
        "Sumbu Sekunder (Secondary Axis) pada Combo Chart menyelesaikan masalah perbandingan dua metrik berukuran skala timpang (Rupiah vs Persen)."
      ],
      notes: "Pesan Guru: Ingatkan siswa bahwa grafik yang baik bukan yang paling ramai warnanya, melainkan yang paling cepat menyampaikan kesimpulan data dengan jujur dan jelas."
    });

    // Slides 2 to 7: All 6 topics
    CHAPTER_5_DATA.topics.forEach(topic => {
      this.slides.push({
        number: topic.number,
        badge: `Shortcut: ${topic.shortcut}`,
        title: `${topic.number} ${topic.title}`,
        subtitle: topic.indonesianTitle,
        bullets: topic.keyPoints.map(p => p.replace(/\*\*/g, '').replace(/\*/g, '')),
        ribbon: topic.ribbonPath,
        notes: `💡 TIPS PRAKTIK: ${(topic.proTips || topic.teacherTips || '').replace(/\*\*/g, '').replace(/\*/g, '')}\n\n⚠️ PERHATIAN: ${(topic.commonMistakes || '').replace(/\*\*/g, '').replace(/\*/g, '')}`
      });
    });

    // Final Slide: Summary & Quiz
    this.slides.push({
      number: "EVALUASI",
      badge: "Kuis Praktik",
      title: "Evaluasi & Kuis Pemahaman Siswa",
      subtitle: "Uji Pemahaman Visualisasi Data & Objek Visual Chapter 5",
      bullets: [
        "Kerjakan 6 pertanyaan studi kasus interaktif untuk mengukur pemahaman konsep.",
        "Pahami pembahasan setiap nomor untuk memantapkan kompetensi visualisasi spreadsheet.",
        "Unduh 3 paket dataset latihan CSV untuk bereksperimen langsung di Microsoft Excel desktop."
      ],
      notes: "Selesaikan kuis evaluasi formatif Chapter 5 (KKM 75) sebagai syarat kelulusan unit visualisasi data."
    });
  },

  open(slideIndex = 0) {
    this.currentSlideIndex = slideIndex;
    if (this.overlayEl) {
      this.overlayEl.classList.add('active');
      document.body.style.overflow = 'hidden';
      this.renderSlide();
    }
  },

  close() {
    if (this.overlayEl) {
      this.overlayEl.classList.remove('active');
      document.body.style.overflow = '';
      this.disableLaser();
    }
  },

  nextSlide() {
    if (this.currentSlideIndex < this.slides.length - 1) {
      this.currentSlideIndex++;
      this.renderSlide();
      if (window.AppCh5) window.AppCh5.playTone(750, 0.05);
    }
  },

  prevSlide() {
    if (this.currentSlideIndex > 0) {
      this.currentSlideIndex--;
      this.renderSlide();
      if (window.AppCh5) window.AppCh5.playTone(600, 0.05);
    }
  },

  toggleLaser() {
    this.laserActive = !this.laserActive;
    const btn = document.getElementById('btn-pres-laser');
    if (this.laserActive) {
      if (btn) btn.style.background = '#e11d48';
      if (this.overlayEl) this.overlayEl.style.cursor = 'none';
      if (this.laserDotEl) this.laserDotEl.style.display = 'block';
    } else {
      this.disableLaser();
    }
  },

  disableLaser() {
    this.laserActive = false;
    const btn = document.getElementById('btn-pres-laser');
    if (btn) btn.style.background = '';
    if (this.overlayEl) this.overlayEl.style.cursor = 'default';
    if (this.laserDotEl) this.laserDotEl.style.display = 'none';
  },

  renderSlide() {
    const slide = this.slides[this.currentSlideIndex];
    if (!slide) return;

    const renderEl = document.getElementById('pres-slide-render');
    const counterEl = document.getElementById('pres-slide-counter');

    if (counterEl) {
      counterEl.textContent = `${this.currentSlideIndex + 1} / ${this.slides.length}`;
    }

    if (renderEl) {
      renderEl.innerHTML = `
        <div class="pres-slide-card">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px;">
            <span class="pres-badge">${slide.badge}</span>
            ${slide.ribbon ? `<span style="font-size:0.85rem; background:#1e293b; color:#94a3b8; padding:4px 12px; border-radius:6px; font-family:var(--font-mono);">📍 ${slide.ribbon}</span>` : ''}
          </div>

          <h2 class="pres-title">${slide.title}</h2>
          <div class="pres-subtitle">${slide.subtitle}</div>

          <div class="pres-bullets-box">
            ${slide.bullets.map(b => `
              <div class="pres-bullet-item">
                <span class="pres-bullet-dot"></span>
                <span>${b}</span>
              </div>
            `).join('')}
          </div>

          ${slide.notes ? `
            <div class="pres-notes-box">
              ${slide.notes.replace(/\n/g, '<br>')}
            </div>
          ` : ''}
        </div>
      `;
    }
  },

  bindEvents() {
    const btnOpen = document.getElementById('btn-open-presentation');
    if (btnOpen) {
      btnOpen.addEventListener('click', () => this.open(0));
    }

    const btnClose = document.getElementById('btn-pres-close');
    if (btnClose) {
      btnClose.addEventListener('click', () => this.close());
    }

    const btnNext = document.getElementById('btn-pres-next');
    if (btnNext) {
      btnNext.addEventListener('click', () => this.nextSlide());
    }

    const btnPrev = document.getElementById('btn-pres-prev');
    if (btnPrev) {
      btnPrev.addEventListener('click', () => this.prevSlide());
    }

    const btnLaser = document.getElementById('btn-pres-laser');
    if (btnLaser) {
      btnLaser.addEventListener('click', () => this.toggleLaser());
    }

    // Keyboard Navigation
    document.addEventListener('keydown', (e) => {
      if (!this.overlayEl || !this.overlayEl.classList.contains('active')) return;

      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
        e.preventDefault();
        this.nextSlide();
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        e.preventDefault();
        this.prevSlide();
      } else if (e.key === 'Escape') {
        this.close();
      } else if (e.key === 'l' || e.key === 'L') {
        this.toggleLaser();
      }
    });

    // Virtual Laser Pointer Mouse Tracking
    document.addEventListener('mousemove', (e) => {
      if (!this.laserActive || !this.laserDotEl) return;
      this.laserDotEl.style.left = `${e.clientX}px`;
      this.laserDotEl.style.top = `${e.clientY}px`;
    });
  }
};

window.PresentationModeCh5 = PresentationModeCh5;

