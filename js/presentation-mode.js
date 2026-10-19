/**
 * TEACHER PRESENTATION MODE (SLIDESHOW & PROJECTOR SUITE)
 * Specifically optimized for classroom projectors, interactive smartboards, and ICT instruction
 */

const PresentationMode = {
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
      badge: "CHAPTER 2 ICT",
      title: CHAPTER_DATA.chapterTitle,
      subtitle: CHAPTER_DATA.chapterSubtitle,
      bullets: [
        "Standardisasi dan pembersihan data massal (Data Cleansing)",
        "Integrasi file eksternal .CSV/.TXT & peran Delimiter",
        "Pengorganisasian data bertingkat (Multi-Level Sort) & Custom List",
        "Format bersyarat dinamis (Conditional Formatting & Below Average)",
        "Optimalisasi kerja: Transpose, Paste Special, dan Fill Handle Series"
      ],
      notes: "Bab ini membekali kamu dengan keahlian penting untuk membersihkan, menstrukturkan, mengimpor, dan mengolah data secara profesional di Microsoft Excel."
    });

    // Slide 1: General Theory
    this.slides.push({
      number: "THEORY",
      badge: "Pengantar Umum",
      title: CHAPTER_DATA.overview.title,
      subtitle: "Mengapa Data Perlu Dimanipulasi & Distandarisasi?",
      bullets: [
        "Data mentah (Raw Data) dari lapangan sering kali acak, kotor, dan terduplikasi.",
        "Manipulasi data di Excel bukan berarti memalsukan, melainkan menstrukturkan data agar siap dibaca dan diolah rumus.",
        "Data yang rapi mempercepat pembuatan laporan dan mencegah kesalahan kalkulasi fatal."
      ],
      notes: "Konsep Inti: Data mentah dari sistem kasir atau form pendaftaran sering kali tidak rapi. Manipulasi data bertujuan menstrukturkan tabel agar siap dianalisis dan diolah rumus secara aman."
    });

    // Slides 2 to 11: All 10 topics
    CHAPTER_DATA.topics.forEach(topic => {
      this.slides.push({
        number: topic.number,
        badge: `Shortcut: ${topic.shortcut}`,
        title: `${topic.number} ${topic.title}`,
        subtitle: topic.indonesianTitle,
        bullets: topic.keyPoints.map(p => p.replace(/\*\*/g, '').replace(/\*/g, '')),
        ribbon: topic.ribbonPath,
        notes: `💡 TIPS PRAKTIK: ${(topic.proTips || topic.teacherTips).replace(/\*\*/g, '').replace(/\*/g, '')}\n\n⚠️ PERHATIAN: ${topic.commonMistakes.replace(/\*\*/g, '').replace(/\*/g, '')}`
      });
    });

    // Final Slide: Summary & Quiz
    this.slides.push({
      number: "EVALUASI",
      badge: "Kuis Praktik",
      title: "Evaluasi & Kuis Pemahaman Siswa",
      subtitle: "Uji Pemahaman Melalui 10 Studi Kasus Nyata",
      bullets: [
        "Kerjakan 10 pertanyaan studi kasus di mode kuis interaktif.",
        "Dapatkan evaluasi instan dengan pembahasan komprehensif di setiap nomor.",
        "Unduh file CSV latihan untuk praktik langsung di software Microsoft Excel laptop kamu."
      ],
      notes: "Uji pemahamanmu setelah mempelajari materi dengan menyelesaikan 10 soal kuis studi kasus dan praktikkan di Microsoft Excel desktop."
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

window.PresentationMode = PresentationMode;
