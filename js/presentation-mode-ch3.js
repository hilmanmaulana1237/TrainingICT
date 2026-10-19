/**
 * TEACHER PRESENTATION MODE (SLIDESHOW & PROJECTOR SUITE) - CHAPTER 3
 * Specifically optimized for classroom projectors, interactive smartboards, and ICT instruction
 * Built for Chapter 3: Tables, Table Styles, Subtotals, and Data Formatting
 */

const PresentationModeCh3 = {
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
      badge: "CHAPTER 3 ICT",
      title: CHAPTER_3_DATA.chapterTitle,
      subtitle: CHAPTER_3_DATA.chapterSubtitle,
      bullets: [
        "3.1 Menghapus kolom tidak relevan & AutoFit Column Width",
        "3.2 Pengaturan cetak: Print Titles ($1:$1) & Gridlines cetak",
        "3.3 Excel Tables (Ctrl+T) vs Rentang data biasa",
        "3.4 Table Styles, Banded Rows, & Total Row kalkulasi otomatis",
        "3.5 Multi-Level Subtotals bertingkat & Outline Levels",
        "3.6 Visualisasi data langsung di sel: Icon Sets, Data Bars, & Sparklines"
      ],
      notes: "Bab 3 memandu siswa menguasai pengorganisasian data tingkat mahir: Excel Table, perhitungan ringkasan terstruktur dengan Subtotal, dan visualisasi data langsung di dalam sel."
    });

    // Slide 1: General Theory
    this.slides.push({
      number: "THEORY",
      badge: "Pengantar Umum",
      title: CHAPTER_3_DATA.overview.title,
      subtitle: "Keunggulan Excel Table & Pemformatan Data Visual",
      bullets: [
        "Data range biasa rentan salah format dan tidak otomatis memperluas rumus saat baris baru bertambah.",
        "Excel Table menyediakan fitur Filter otomatis, Structured References, Banded Rows, dan Auto-Expansion.",
        "Subtotal bertingkat dan Sparklines memudahkan visualisasi tren data per grup tanpa membuat chart besar."
      ],
      notes: "Konsep Inti: Mengubah data biasa menjadi Excel Table adalah standar profesional untuk efisiensi dan kerapian laporan kerja."
    });

    // Slides 2 to 7: All 6 topics
    CHAPTER_3_DATA.topics.forEach(topic => {
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
      subtitle: "Uji Pemahaman Melalui 6 Studi Kasus Nyata Chapter 3",
      bullets: [
        "Kerjakan 6 pertanyaan studi kasus di mode kuis interaktif.",
        "Dapatkan evaluasi instan dengan pembahasan komprehensif di setiap nomor.",
        "Unduh file CSV latihan untuk praktik langsung di software Microsoft Excel laptop kamu."
      ],
      notes: "Selesaikan kuis evaluasi formatif untuk mengukur penguasaan standar kompetensi Chapter 3 (KKM 75)."
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
    }
  },

  prevSlide() {
    if (this.currentSlideIndex > 0) {
      this.currentSlideIndex--;
      this.renderSlide();
    }
  },

  renderSlide() {
    const slide = this.slides[this.currentSlideIndex];
    const container = document.getElementById('pres-slide-render');
    if (!container) return;

    container.innerHTML = `
      <div class="pres-slide-card">
        <div class="pres-slide-header">
          <span class="pres-slide-number">${slide.number}</span>
          <div>
            <div style="font-size:0.85rem; font-weight:700; color:var(--accent-mint); text-transform:uppercase; letter-spacing:0.05em;">
              ${slide.badge || ''}
            </div>
            <h2 class="pres-slide-title">${slide.title}</h2>
            ${slide.subtitle ? `<div style="font-size:1.1rem; color:#94a3b8; margin-top:2px;">${slide.subtitle}</div>` : ''}
          </div>
        </div>

        ${slide.ribbon ? `
          <div style="display:inline-flex; align-items:center; gap:8px; background:rgba(16,124,65,0.2); border:1px solid rgba(16,185,129,0.3); padding:6px 14px; border-radius:8px; font-family:var(--font-mono); font-size:0.9rem; color:#86efac; margin-bottom:24px;">
            📌 Jalur Menu: <strong>${slide.ribbon}</strong>
          </div>
        ` : ''}

        <ul class="pres-slide-bullets">
          ${slide.bullets.map(b => `<li>${b}</li>`).join('')}
        </ul>

        ${slide.notes ? `
          <div class="pres-teacher-notes">
            <div style="font-weight:700; margin-bottom:4px; display:flex; align-items:center; gap:6px;">
              💡 Catatan Penting & Tips Praktik:
            </div>
            <div style="white-space:pre-line; line-height:1.5;">${slide.notes}</div>
          </div>
        ` : ''}
      </div>
    `;

    // Update Counter
    const counterEl = document.getElementById('pres-slide-counter');
    if (counterEl) {
      counterEl.innerText = `${this.currentSlideIndex + 1} / ${this.slides.length}`;
    }
  },

  toggleLaser() {
    this.laserActive = !this.laserActive;
    const btn = document.getElementById('btn-pres-laser');
    if (btn) {
      btn.style.background = this.laserActive ? '#ef4444' : '';
      btn.style.color = this.laserActive ? '#ffffff' : '';
    }

    if (this.laserActive) {
      if (this.laserDotEl) this.laserDotEl.style.display = 'block';
      document.body.style.cursor = 'none';
    } else {
      this.disableLaser();
    }
  },

  disableLaser() {
    this.laserActive = false;
    if (this.laserDotEl) this.laserDotEl.style.display = 'none';
    document.body.style.cursor = '';
    const btn = document.getElementById('btn-pres-laser');
    if (btn) {
      btn.style.background = '';
      btn.style.color = '';
    }
  },

  bindEvents() {
    // Keyboard controls
    document.addEventListener('keydown', (e) => {
      if (!this.overlayEl || !this.overlayEl.classList.contains('active')) return;

      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
        e.preventDefault();
        this.nextSlide();
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        e.preventDefault();
        this.prevSlide();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.close();
      } else if (e.key === 'l' || e.key === 'L') {
        this.toggleLaser();
      }
    });

    // Mouse move for laser pointer
    document.addEventListener('mousemove', (e) => {
      if (this.laserActive && this.laserDotEl) {
        this.laserDotEl.style.left = `${e.clientX}px`;
        this.laserDotEl.style.top = `${e.clientY}px`;
      }
    });

    // Next / Prev buttons
    const btnNext = document.getElementById('btn-pres-next');
    const btnPrev = document.getElementById('btn-pres-prev');
    const btnClose = document.getElementById('btn-pres-close');
    const btnLaser = document.getElementById('btn-pres-laser');

    if (btnNext) btnNext.addEventListener('click', () => this.nextSlide());
    if (btnPrev) btnPrev.addEventListener('click', () => this.prevSlide());
    if (btnClose) btnClose.addEventListener('click', () => this.close());
    if (btnLaser) btnLaser.addEventListener('click', () => this.toggleLaser());
  }
};

window.PresentationMode = PresentationModeCh3;
window.PresentationModeCh3 = PresentationModeCh3;
