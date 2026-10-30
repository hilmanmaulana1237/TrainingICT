/**
 * TEACHER PRESENTATION MODE (SLIDESHOW & PROJECTOR SUITE) - CHAPTER 4
 * Specifically optimized for classroom projectors, interactive smartboards, and ICT instruction
 * Built for Chapter 4: Excel Formulas and Functions
 */

const PresentationModeCh4 = {
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
      badge: "CHAPTER 4 ICT",
      title: CHAPTER_4_DATA.chapterTitle,
      subtitle: CHAPTER_4_DATA.chapterSubtitle,
      bullets: [
        "4.1 Struktur Rumus & Referensi Sel (Relatif A1, Absolut $A$1, tombol F4)",
        "4.2 Fungsi Statistik Dasar (MAX, MIN, AVERAGE, LARGE, COUNT vs COUNTA)",
        "4.3 Agregasi Bersyarat Multi-Kriteria (COUNTIFS & SUMIFS stabil terhadap Sort)",
        "4.4 Fungsi Teks: Penggabungan (&, CONCATENATE, TEXTJOIN) & Format Huruf (PROPER, UPPER, LOWER)",
        "4.5 Fungsi Logika IF: Keputusan Otomatis & Salin Rumus Fill Down (Ctrl+D)",
        "4.6 Formula Auditing: Show Formulas (Ctrl+`) & Pelacakan Alur Sel (Trace Precedents)"
      ],
      notes: "Bab 4 adalah inti kecakapan spreadsheet: menguasai kalkulasi otomatis, fungsi statistik, logika IF, agregasi multi-kriteria, dan audit rumus secara profesional."
    });

    // Slide 1: General Theory
    this.slides.push({
      number: "THEORY",
      badge: "Pengantar Umum",
      title: CHAPTER_4_DATA.overview.title,
      subtitle: "Fondasi Formula, Referensi Sel, & Fungsi Logika Excel",
      bullets: [
        "Semua rumus Excel wajib dimulai dengan tanda sama dengan (=).",
        "Referensi relatif (A1) bergeser saat disalin; referensi absolut ($A$1) terkunci permanen pada sel tertentu.",
        "COUNTIFS dan SUMIFS mengevaluasi nilai sel aktual, sehingga kalkulasi tetap presisi meski data di-sort.",
        "Show Formulas (Ctrl+`) membuka seluruh teks rumus lembar kerja untuk audit dan pemeriksaan integritas data."
      ],
      notes: "Pesan Guru: Ajarkan siswa pentingnya tombol F4 untuk mengunci sel sebelum mereka menyalin rumus ke ratusan baris data."
    });

    // Slides 2 to 7: All 6 topics
    CHAPTER_4_DATA.topics.forEach(topic => {
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
      subtitle: "Uji Pemahaman Melalui 6 Kasus Nyata Rumus & Fungsi Chapter 4",
      bullets: [
        "Kerjakan 6 pertanyaan studi kasus di mode kuis interaktif.",
        "Dapatkan evaluasi instan dengan pembahasan komprehensif di setiap nomor.",
        "Unduh file CSV latihan untuk praktik langsung di software Microsoft Excel komputer kamu."
      ],
      notes: "Selesaikan kuis evaluasi formatif untuk mengukur penguasaan standar kompetensi Chapter 4 (KKM 75)."
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
      if (window.App) window.App.playTone(750, 0.05);
    }
  },

  prevSlide() {
    if (this.currentSlideIndex > 0) {
      this.currentSlideIndex--;
      this.renderSlide();
      if (window.App) window.App.playTone(600, 0.05);
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
    const renderEl = document.getElementById('pres-slide-render');
    const counterEl = document.getElementById('pres-slide-counter');

    if (counterEl) {
      counterEl.innerText = `${this.currentSlideIndex + 1} / ${this.slides.length}`;
    }

    if (!renderEl || !slide) return;

    renderEl.innerHTML = `
      <div class="pres-slide-card">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
          <span class="badge" style="background:#0284c7; color:white; font-size:0.85rem; padding:4px 12px; font-weight:700;">
            ${slide.badge}
          </span>
          <span style="font-family:var(--font-mono); font-size:0.9rem; color:#94a3b8;">
            SLIDE ${slide.number}
          </span>
        </div>

        <h2 style="font-size:2.2rem; font-family:var(--font-heading); color:#f8fafc; margin-bottom:6px; line-height:1.2;">
          ${slide.title}
        </h2>
        <div style="font-size:1.15rem; color:#38bdf8; font-weight:600; margin-bottom:24px;">
          ${slide.subtitle}
        </div>

        <div style="background:#1e293b; border-radius:12px; padding:20px 24px; border:1px solid #334155; margin-bottom:20px;">
          <h4 style="color:#e2e8f0; font-size:1rem; font-weight:700; margin-bottom:12px; display:flex; align-items:center; gap:8px;">
            <span>📌</span> Konsep & Pembahasan Utama:
          </h4>
          <ul style="margin:0; padding-left:22px; color:#cbd5e1; font-size:1.05rem; line-height:1.7;">
            ${slide.bullets.map(b => `<li style="margin-bottom:8px;">${b}</li>`).join('')}
          </ul>
        </div>

        ${slide.ribbon ? `
          <div style="display:flex; align-items:center; gap:8px; font-size:0.85rem; color:#94a3b8; background:#0f172a; padding:8px 14px; border-radius:6px; border:1px solid #334155;">
            <span>🧭 <strong>Ribbon Menu:</strong></span>
            <code>${slide.ribbon}</code>
          </div>
        ` : ''}

        ${slide.notes ? `
          <div style="margin-top:16px; padding:12px 16px; background:rgba(16, 185, 129, 0.1); border-left:4px solid #10b981; border-radius:4px; font-size:0.9rem; color:#a7f3d0; white-space:pre-line;">
            ${slide.notes}
          </div>
        ` : ''}
      </div>
    `;
  },

  bindEvents() {
    // Open & Close
    document.getElementById('btn-open-presentation')?.addEventListener('click', () => {
      this.open(0);
    });

    document.getElementById('btn-pres-close')?.addEventListener('click', () => {
      this.close();
    });

    // Navigation buttons
    document.getElementById('btn-pres-next')?.addEventListener('click', () => {
      this.nextSlide();
    });

    document.getElementById('btn-pres-prev')?.addEventListener('click', () => {
      this.prevSlide();
    });

    // Laser pointer toggle
    document.getElementById('btn-pres-laser')?.addEventListener('click', () => {
      this.toggleLaser();
    });

    // Mouse movement for virtual laser
    window.addEventListener('mousemove', (e) => {
      if (this.laserActive && this.laserDotEl) {
        this.laserDotEl.style.left = `${e.clientX}px`;
        this.laserDotEl.style.top = `${e.clientY}px`;
      }
    });

    // Keyboard navigation
    window.addEventListener('keydown', (e) => {
      if (!this.overlayEl || !this.overlayEl.classList.contains('active')) return;

      if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {
        e.preventDefault();
        this.nextSlide();
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        e.preventDefault();
        this.prevSlide();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.close();
      } else if (e.key.toLowerCase() === 'l') {
        e.preventDefault();
        this.toggleLaser();
      }
    });
  }
};

window.PresentationMode = PresentationModeCh4;
window.PresentationModeCh4 = PresentationModeCh4;
