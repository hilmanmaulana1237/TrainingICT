/**
 * MAIN APP CONTROLLER - ICT EXCEL CHAPTER 5
 * Charts, Sparklines, and Visual Objects
 * Orchestrates navigation, curriculum rendering, quizzes, downloads, audio, and teacher tools
 */

const AppCh5 = {
  soundEnabled: true,
  audioCtx: null,
  completedTopics: new Set(),
  timerInterval: null,
  timerSecondsLeft: 90 * 60, // Default 90 minutes ICT class
  timerRunning: false,

  init() {
    this.loadState();
    this.renderSidebarNav();
    this.renderTopicsContent();
    this.initAudioContext();
    this.bindGlobalEvents();
    this.initQuiz();
    this.updateProgressDisplay();
    this.initTimer();

    // Initialize simulators after DOM is ready
    if (window.SpreadsheetSimulatorsCh5) {
      window.SpreadsheetSimulatorsCh5.initAll();
    }

    // Initialize presentation mode
    if (window.PresentationModeCh5) {
      window.PresentationModeCh5.init();
    }
  },

  // Audio synthesizer via Web Audio API (zero external assets required)
  initAudioContext() {
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (AudioCtx) {
        this.audioCtx = new AudioCtx();
      }
    } catch (e) {
      console.warn("Web Audio not supported", e);
    }
  },

  playTone(freq = 600, duration = 0.08) {
    if (!this.soundEnabled || !this.audioCtx) return;
    try {
      if (this.audioCtx.state === 'suspended') {
        this.audioCtx.resume();
      }
      const osc = this.audioCtx.createOscillator();
      const gain = this.audioCtx.createGain();
      osc.type = 'sine';
      osc.frequency.setValueAtTime(freq, this.audioCtx.currentTime);
      gain.gain.setValueAtTime(0.08, this.audioCtx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, this.audioCtx.currentTime + duration);
      osc.connect(gain);
      gain.connect(this.audioCtx.destination);
      osc.start();
      osc.stop(this.audioCtx.currentTime + duration);
    } catch (e) {}
  },

  loadState() {
    // Theme preference
    const savedTheme = localStorage.getItem('ict_excel_theme');
    if (savedTheme === 'dark') {
      document.body.classList.add('dark-mode');
    }

    // Sound preference
    const savedSound = localStorage.getItem('ict_excel_sound');
    if (savedSound !== null) {
      this.soundEnabled = savedSound === 'true';
    }

    // Completed topics for Chapter 5
    const savedCompleted = localStorage.getItem('ict_excel_ch5_completed');
    if (savedCompleted) {
      try {
        const arr = JSON.parse(savedCompleted);
        this.completedTopics = new Set(arr);
      } catch (e) {}
    }
  },

  saveState() {
    localStorage.setItem('ict_excel_theme', document.body.classList.contains('dark-mode') ? 'dark' : 'light');
    localStorage.setItem('ict_excel_sound', this.soundEnabled.toString());
    localStorage.setItem('ict_excel_ch5_completed', JSON.stringify([...this.completedTopics]));
  },

  toggleTheme() {
    document.body.classList.toggle('dark-mode');
    this.playTone(500, 0.05);
    this.saveState();
  },

  toggleSound() {
    this.soundEnabled = !this.soundEnabled;
    const btn = document.getElementById('btn-toggle-sound');
    if (btn) {
      btn.innerHTML = this.soundEnabled ? '🔊' : '🔇';
      btn.title = this.soundEnabled ? 'Suara Aktif' : 'Suara Dimatikan';
    }
    if (this.soundEnabled) this.playTone(880, 0.1);
    this.saveState();
  },

  // Render Sidebar
  renderSidebarNav() {
    const navList = document.getElementById('sidebar-nav-list');
    if (!navList) return;

    let html = `
      <li>
        <a href="#overview" class="nav-link active" data-target="overview">
          <span class="nav-icon">📖</span>
          <span class="nav-label">General Theory (Visualisasi Data)</span>
        </a>
      </li>
      <li>
        <a href="#chart-quickstart" class="nav-link" data-target="chart-quickstart" style="color:#4f46e5; font-weight:700;">
          <span class="nav-icon">🔰</span>
          <span class="nav-label">Panduan Visual & Jenis Grafik</span>
        </a>
      </li>
    `;

    CHAPTER_5_DATA.topics.forEach((topic) => {
      const isDone = this.completedTopics.has(topic.id);
      html += `
        <li>
          <a href="#${topic.id}" class="nav-link" data-target="${topic.id}">
            <span class="nav-icon">${isDone ? '✅' : '📊'}</span>
            <span class="nav-label">${topic.number} ${topic.title}</span>
          </a>
        </li>
      `;
    });

    html += `
      <li style="margin-top:14px; border-top:1px solid var(--border-color); padding-top:10px;">
        <a href="#quiz-section" class="nav-link" data-target="quiz-section">
          <span class="nav-icon">📝</span>
          <span class="nav-label">Kuis Praktik Evaluasi Siswa</span>
        </a>
      </li>
    `;

    navList.innerHTML = html;

    // Attach click events for active state & smooth scrolling
    const links = navList.querySelectorAll('.nav-link');
    links.forEach(link => {
      link.addEventListener('click', (e) => {
        links.forEach(l => l.classList.remove('active'));
        link.classList.add('active');
        this.playTone(700, 0.04);

        // Auto close mobile drawer if open
        const sidebar = document.getElementById('app-sidebar');
        if (sidebar && sidebar.classList.contains('mobile-open')) {
          sidebar.classList.remove('mobile-open');
        }
      });
    });
  },

  // Helper for Markdown formatting
  formatInlineMarkdown(text) {
    if (!text) return '';
    return text
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      .replace(/\*(.*?)\*/g, '<em>$1</em>')
      .replace(/`([^`]+)`/g, '<code>$1</code>');
  },

  formatMarkdown(text) {
    if (!text) return '';
    const paragraphs = text.split('\n\n');

    return paragraphs.map(p => {
      p = p.trim();
      if (!p) return '';

      if (p.startsWith('### ')) {
        return `<h3 style="font-family:var(--font-heading); font-size:1.1rem; font-weight:700; color:var(--text-main); margin-top:18px; margin-bottom:10px;">${this.formatInlineMarkdown(p.replace('### ', ''))}</h3>`;
      }
      if (p.startsWith('## ')) {
        return `<h2 style="font-family:var(--font-heading); font-size:1.25rem; font-weight:800; color:var(--text-main); margin-top:22px; margin-bottom:12px;">${this.formatInlineMarkdown(p.replace('## ', ''))}</h2>`;
      }

      const lines = p.split('\n');
      let html = '';
      let bulletItems = [];
      let orderedItems = [];

      const flushLists = () => {
        if (bulletItems.length > 0) {
          html += `<ul style="margin-left:20px; margin-bottom:12px; line-height:1.65;">${bulletItems.map(item => `<li>${item}</li>`).join('')}</ul>`;
          bulletItems = [];
        }
        if (orderedItems.length > 0) {
          html += `<ol style="margin-left:20px; margin-bottom:12px; line-height:1.65;">${orderedItems.map(item => `<li>${item}</li>`).join('')}</ol>`;
          orderedItems = [];
        }
      };

      lines.forEach(line => {
        line = line.trim();
        if (line.startsWith('* ') || line.startsWith('- ')) {
          if (orderedItems.length > 0) flushLists();
          const clean = line.replace(/^[\*\-]\s+/, '');
          bulletItems.push(this.formatInlineMarkdown(clean));
        } else if (/^\d+\.\s+/.test(line)) {
          if (bulletItems.length > 0) flushLists();
          const clean = line.replace(/^\d+\.\s+/, '');
          orderedItems.push(this.formatInlineMarkdown(clean));
        } else {
          flushLists();
          html += `<p style="margin-bottom:10px; line-height:1.65;">${this.formatInlineMarkdown(line)}</p>`;
        }
      });

      flushLists();
      return html;
    }).join('');
  },

  // Render Topics
  renderTopicsContent() {
    const container = document.getElementById('topics-render-container');
    if (!container) return;

    let html = "";

    CHAPTER_5_DATA.topics.forEach((topic) => {
      const isDone = this.completedTopics.has(topic.id);
      const dt = topic.desktopTutorial;

      html += `
        <article class="topic-card" id="${topic.id}">
          <header class="topic-header-bar" style="border-left:4px solid #4f46e5;">
            <div class="topic-title-group">
              <div class="topic-num-badge" style="background:#4f46e5;">${topic.number}</div>
              <div class="topic-headings">
                <h2>${topic.title}</h2>
                <div class="topic-subheading">${topic.indonesianTitle}</div>
              </div>
            </div>

            <div class="topic-header-badges">
              <span class="shortcut-badge" title="Pintasan Keyboard">
                ⌨️ <kbd>${topic.shortcut}</kbd>
              </span>
              <span class="ribbon-badge" title="Lokasi di Menu Excel">
                🧭 ${topic.ribbonPath}
              </span>
            </div>
          </header>

          <div class="topic-body">
            <!-- Theory & Core Concept -->
            <div class="topic-summary-box" style="border-left-color:#4f46e5;">
              "${topic.summary}"
            </div>

            <div class="theory-text">
              ${this.formatMarkdown(topic.theoryExplanation)}
            </div>

            <!-- Key Points -->
            <div>
              <h3 style="font-size:1rem; font-weight:700; margin-bottom:8px; color:var(--text-main);">
                📌 Poin Kunci & Karakteristik Fitur:
              </h3>
              <ul class="key-points-list">
                ${topic.keyPoints.map(kp => `<li>${this.formatInlineMarkdown(kp)}</li>`).join('')}
              </ul>
            </div>

            <!-- Interactive Spreadsheet Simulator Lab -->
            <div style="margin-top:14px;">
              <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
                <h4 style="font-family:var(--font-heading); font-size:1.05rem; font-weight:700; color:#4f46e5; display:flex; align-items:center; gap:8px;">
                  💻 Lab Interaktif & Simulator Langsung
                </h4>
                <span style="font-size:0.75rem; color:var(--text-muted);">Simulasi Real-time Excel Browser</span>
              </div>

              <!-- Lab Container placeholder -->
              <div id="lab-${topic.number.replace('.', '-')}-container"></div>
            </div>

            <!-- Guided Desktop Tutorial for Real Excel -->
            ${dt ? `
            <div class="excel-desktop-guide-box">
              <div class="guide-box-header">
                <div>
                  <h4 class="guide-box-title">
                    <span>🖥️</span> ${dt.title}
                  </h4>
                  <div class="sheet-target-pill" style="margin-top:6px; margin-bottom:0;">
                    <span>Target Sheet:</span>
                    <code>${dt.recommendedSheetName}</code>
                    <button class="btn btn-outline" style="padding:2px 8px; font-size:0.75rem; margin-left:6px;" onclick="AppCh5.copySheetName('${dt.recommendedSheetName}', this)">📋 Salin Nama Sheet</button>
                  </div>
                </div>

                <div style="display:flex; gap:8px; flex-wrap:wrap;">
                  <button class="copy-practice-btn" onclick="AppCh5.copyTopicTable('${topic.id}', '${dt.recommendedSheetName}', '${topic.title}', this)">
                    📋 Salin Tabel Latihan ke Excel (Ctrl+V)
                  </button>
                </div>
              </div>

              <!-- Excel Paste Pro-Tip Box -->
              <div class="excel-paste-pro-tip">
                <div class="paste-tip-badge" style="background:#4f46e5;">💡 PANDUAN PRAKTIK DI MICROSOFT EXCEL</div>
                <div class="paste-tip-content">
                  <p>Cara menempelkan data latihan dan membuat visualisasi agar langsung presisi:</p>
                  <ol class="paste-tip-steps">
                    <li>Buka aplikasi Microsoft Excel, buat sheet baru bernama: <code>${dt.recommendedSheetName}</code>.</li>
                    <li>Klik sel <strong>A1 SATU KALI SAJA</strong> (pastikan tidak dobel klik agar kursor tidak berkedip di dalam sel).</li>
                    <li>Tekan <strong>Ctrl + V</strong> di keyboard kamu. Tabel latihan langsung tertempel sempurna!</li>
                  </ol>
                  <div class="paste-warning-note">
                    ⚠️ <strong>Tips Pintasan Cepat:</strong> Tekan <strong>Alt + F1</strong> untuk membuat grafik di lembar aktif atau <strong>F11</strong> untuk membuat grafik di lembar khusus (*Chart Sheet*).
                  </div>
                </div>
              </div>

              <!-- Visual Sample Table Preview -->
              ${dt.sampleHtml ? `
              <div class="desktop-sample-table-container">
                <div class="sample-table-header-info">
                  <span class="sample-table-badge" style="color:#4f46e5;">
                    <span>👁️</span> Pratinjau Desain Tabel Excel
                  </span>
                  <span class="sample-table-hint">
                    Tampilan tabel setelah ditempelkan ke lembar kerja Excel kamu:
                  </span>
                </div>
                <div class="sample-table-wrapper">
                  ${dt.sampleHtml}
                </div>
              </div>
              ` : ''}

              <!-- Step-by-Step Instructions -->
              <div style="margin-bottom:18px;">
                <div style="font-size:0.95rem; font-weight:700; color:var(--text-main); margin-bottom:12px; display:flex; align-items:center; gap:8px;">
                  <span>📋 Langkah-Langkah Praktik Terpandu di Excel Komputer:</span>
                </div>
                <ul class="desktop-steps-list">
                  ${dt.steps.map((step, idx) => {
                    const formulaMatch = step.match(/`(=[^`]+)`/);
                    let formulaButton = '';
                    if (formulaMatch && formulaMatch[1]) {
                      const cleanFormula = formulaMatch[1];
                      formulaButton = `<button class="copy-formula-btn" onclick="AppCh5.copyFormulaText('${cleanFormula.replace(/'/g, "\\'")}', this)" title="Salin rumus ${cleanFormula}">📋 Salin Rumus</button>`;
                    }
                    return `
                      <li class="desktop-step-item">
                        <span class="step-num-badge" style="background:#4f46e5;">Langkah ${idx + 1}</span>
                        <div class="step-text-content">
                          ${this.formatInlineMarkdown(step)}
                          ${formulaButton}
                        </div>
                      </li>
                    `;
                  }).join('')}
                </ul>
              </div>

              <!-- Expected Result Box -->
              <div class="expected-result-card" style="border-left-color:#4f46e5;">
                <span class="expected-badge" style="color:#4f46e5;">🎯 Target Hasil Akhir:</span>
                <p class="expected-text">
                  ${dt.expectedResult}
                </p>
              </div>
            </div>
            ` : ''}

            <!-- Topic Footer & Completion Button -->
            <footer class="topic-footer-bar">
              <button class="topic-check-btn ${isDone ? 'completed' : ''}" onclick="AppCh5.toggleTopicCompletion('${topic.id}', this)">
                <span class="check-icon">${isDone ? '✓' : '○'}</span>
                <span>${isDone ? 'Telah Selesai Dipraktikkan' : 'Tandai Selesai Topik Ini'}</span>
              </button>

              <span style="font-size:0.8rem; color:var(--text-muted);">
                ${isDone ? '🎉 Kompetensi topik ini telah tercapai.' : 'Klik jika kamu telah memahami materi & mencoba simulasinya.'}
              </span>
            </footer>
          </div>
        </article>
      `;
    });

    container.innerHTML = html;
  },

  // Toggle Topic Completion
  toggleTopicCompletion(topicId, btnEl) {
    if (this.completedTopics.has(topicId)) {
      this.completedTopics.delete(topicId);
      this.playTone(400, 0.06);
    } else {
      this.completedTopics.add(topicId);
      this.playTone(850, 0.08);
      setTimeout(() => this.playTone(1100, 0.12), 90);
    }

    this.saveState();
    this.updateProgressDisplay();
    this.renderSidebarNav();

    if (btnEl) {
      const isDone = this.completedTopics.has(topicId);
      btnEl.classList.toggle('completed', isDone);
      btnEl.querySelector('.check-icon').textContent = isDone ? '✓' : '○';
      btnEl.querySelector('span:last-child').textContent = isDone ? 'Telah Selesai Dipraktikkan' : 'Tandai Selesai Topik Ini';
    }
  },

  // Update Progress Display
  updateProgressDisplay() {
    const total = CHAPTER_5_DATA.topics.length;
    const completed = this.completedTopics.size;
    const percent = Math.round((completed / total) * 100);

    const barFill = document.getElementById('progress-bar-fill');
    const textLabel = document.getElementById('progress-text');
    const heroStat = document.getElementById('hero-stat-progress');

    if (barFill) barFill.style.width = `${percent}%`;
    if (textLabel) textLabel.textContent = `${completed} / ${total} Topik (${percent}%)`;
    if (heroStat) heroStat.textContent = `${percent}%`;
  },

  // Copy Helpers
  copyToClipboard(text, title, sheetName) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => {
        this.playTone(900, 0.08);
        this.showCopySuccessModal(title, sheetName);
      }).catch(() => {
        this.fallbackCopy(text, title, sheetName);
      });
    } else {
      this.fallbackCopy(text, title, sheetName);
    }
  },

  fallbackCopy(text, title, sheetName) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand('copy');
      this.playTone(900, 0.08);
      this.showCopySuccessModal(title, sheetName);
    } catch (err) {
      alert("Gagal menyalin otomatis. Silakan blok tabel secara manual.");
    }
    document.body.removeChild(ta);
  },

  copyTopicTable(topicId, sheetName, topicTitle, btnEl) {
    const topic = CHAPTER_5_DATA.topics.find(t => t.id === topicId);
    if (!topic || !topic.desktopTutorial) return;

    const data = topic.desktopTutorial.sampleData;
    this.copyToClipboard(data, `Latihan ${topic.number}: ${topicTitle}`, sheetName);

    if (btnEl) {
      const orig = btnEl.innerHTML;
      btnEl.innerHTML = `✓ Berhasil Disalin! Buka Excel & Ctrl+V`;
      btnEl.style.background = "#107c41";
      btnEl.style.color = "#ffffff";
      setTimeout(() => {
        btnEl.innerHTML = orig;
        btnEl.style.background = "";
        btnEl.style.color = "";
      }, 3000);
    }
  },

  copyFormulaText(formulaText, btnEl) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(formulaText).then(() => {
        this.playTone(950, 0.06);
        if (btnEl) {
          btnEl.classList.add('copied');
          btnEl.innerHTML = '✓ Tersalin!';
          setTimeout(() => {
            btnEl.classList.remove('copied');
            btnEl.innerHTML = '📋 Salin Rumus';
          }, 2000);
        }
      });
    }
  },

  copySheetName(sheetName, btnEl) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(sheetName).then(() => {
        this.playTone(900, 0.06);
        if (btnEl) {
          btnEl.textContent = '✓ Disalin!';
          setTimeout(() => {
            btnEl.textContent = '📋 Salin Nama Sheet';
          }, 2000);
        }
      });
    }
  },

  showCopySuccessModal(title, sheetName) {
    const modalBody = document.getElementById('copy-success-modal-body');
    if (!modalBody) return;

    modalBody.innerHTML = `
      <div style="text-align:center; padding:10px 0;">
        <div style="font-size:3rem; margin-bottom:8px;">📋✨</div>
        <h4 style="font-family:var(--font-heading); font-size:1.15rem; font-weight:800; color:var(--text-main); margin-bottom:6px;">
          Tabel Data Siap Ditempel ke Excel!
        </h4>
        <p style="font-size:0.88rem; color:var(--text-muted); margin-bottom:16px;">
          Data untuk <strong>${title}</strong> telah tersimpan di clipboard komputer kamu.
        </p>

        <div style="background:var(--bg-subtle); border:1px solid var(--border-color); border-radius:8px; padding:14px; text-align:left; font-size:0.85rem; line-height:1.6; margin-bottom:18px;">
          <strong style="color:#4f46e5;">3 Langkah Mudah di Aplikasi Excel:</strong>
          <ol style="margin-left:18px; margin-top:6px;">
            <li>Buka Microsoft Excel, buat sheet baru bernama: <code style="background:var(--bg-surface); padding:2px 6px; border-radius:4px; font-weight:700;">${sheetName}</code></li>
            <li>Klik sel <strong>A1 SATU KALI</strong> (jangan double click).</li>
            <li>Tekan shortcut keyboard: <kbd style="background:var(--bg-surface); padding:2px 6px; border-radius:4px; border:1px solid var(--border-color); font-weight:700;">Ctrl</kbd> + <kbd style="background:var(--bg-surface); padding:2px 6px; border-radius:4px; border:1px solid var(--border-color); font-weight:700;">V</kbd></li>
          </ol>
        </div>

        <button class="btn btn-primary" style="padding:8px 24px; font-weight:700; background:#4f46e5; border-color:#4f46e5;" onclick="AppCh5.closeModal('copy-success-modal')">
          Siap, Saya Mengerti! 👍
        </button>
      </div>
    `;

    this.openModal('copy-success-modal');
  },

  // Modal Control
  openModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
      modal.classList.add('active');
      document.body.style.overflow = 'hidden';
      this.playTone(700, 0.05);
    }
  },

  closeModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
      modal.classList.remove('active');
      document.body.style.overflow = '';
      this.playTone(500, 0.04);
    }
  },

  // Interactive Quiz
  quizState: {
    userAnswers: {},
    submitted: false
  },

  initQuiz() {
    const container = document.getElementById('quiz-questions-container');
    if (!container) return;

    let html = "";
    CHAPTER_5_DATA.quizQuestions.forEach((q, qIdx) => {
      html += `
        <div class="quiz-question-card" id="quiz-card-${q.id}">
          <div class="quiz-question-header">
            <span class="quiz-q-num">Pertanyaan ${qIdx + 1} / 6</span>
            <span class="quiz-topic-ref">Topik ${q.topicRef}</span>
          </div>
          <div class="quiz-q-text">${q.question}</div>
          <div class="quiz-options-list">
            ${q.options.map((opt, optIdx) => `
              <label class="quiz-option-label" id="opt-label-${q.id}-${optIdx}">
                <input type="radio" name="quiz-q-${q.id}" value="${optIdx}" onchange="AppCh5.selectQuizOption(${q.id}, ${optIdx})">
                <span class="opt-letter">${String.fromCharCode(65 + optIdx)}</span>
                <span class="opt-text">${opt}</span>
              </label>
            `).join('')}
          </div>
          <div class="quiz-explanation-box" id="explanation-${q.id}">
            💡 <strong>Pembahasan:</strong> ${q.explanation}
          </div>
        </div>
      `;
    });

    container.innerHTML = html;
  },

  selectQuizOption(qId, optIdx) {
    this.quizState.userAnswers[qId] = optIdx;
    this.playTone(650, 0.04);

    const card = document.getElementById(`quiz-card-${qId}`);
    if (card) {
      const labels = card.querySelectorAll('.quiz-option-label');
      labels.forEach((l, idx) => {
        l.classList.toggle('selected', idx === optIdx);
      });
    }
  },

  submitQuiz() {
    const questions = CHAPTER_5_DATA.quizQuestions;
    const total = questions.length;
    let score = 0;

    questions.forEach(q => {
      const selected = this.quizState.userAnswers[q.id];
      const card = document.getElementById(`quiz-card-${q.id}`);
      const explanation = document.getElementById(`explanation-${q.id}`);

      if (explanation) explanation.classList.add('visible');

      if (selected !== undefined) {
        const isCorrect = selected === q.correctIndex;
        if (isCorrect) score++;

        const selectedLabel = document.getElementById(`opt-label-${q.id}-${selected}`);
        const correctLabel = document.getElementById(`opt-label-${q.id}-${q.correctIndex}`);

        if (selectedLabel) {
          selectedLabel.classList.add(isCorrect ? 'correct' : 'incorrect');
        }
        if (correctLabel && !isCorrect) {
          correctLabel.classList.add('correct');
        }
      }
    });

    const percent = Math.round((score / total) * 100);
    this.showQuizResultModal(score, total, percent);
    this.quizState.submitted = true;

    if (score >= 4) {
      this.playTone(800, 0.1);
      setTimeout(() => this.playTone(1050, 0.15), 100);
    } else {
      this.playTone(400, 0.15);
    }
  },

  resetQuiz() {
    this.quizState.userAnswers = {};
    this.quizState.submitted = false;
    this.initQuiz();
    const btnSubmit = document.getElementById('btn-submit-quiz');
    if (btnSubmit) btnSubmit.style.display = 'block';
  },

  showQuizResultModal(score, total, percent) {
    const modalBox = document.getElementById('quiz-result-modal-box');
    if (!modalBox) return;

    const isPassed = percent >= 75;

    modalBox.innerHTML = `
      <div class="modal-header">
        <h3 class="modal-title">📊 Hasil Evaluasi Mandiri Chapter 5</h3>
        <button class="btn-icon modal-close-btn" onclick="AppCh5.closeModal('quiz-result-modal')">&times;</button>
      </div>
      <div class="modal-body" style="text-align:center; padding:20px 0;">
        <div style="font-size:3.5rem; margin-bottom:8px;">${isPassed ? '🏆' : '📚'}</div>
        <h3 style="font-size:1.6rem; font-weight:800; color:${isPassed ? '#107c41' : '#d97706'}; margin-bottom:6px;">
          ${isPassed ? 'LULUS KOMPETENSI!' : 'PERLU REVIEW MATERI'}
        </h3>
        <p style="color:var(--text-muted); margin-bottom:16px;">
          ${isPassed ? 'Selamat! Kamu telah menguasai visualisasi data, grafik, sparklines, dan objek visual Excel.' : 'Kamu hampir berhasil! Pelajari kembali pembahasan di bawah kuis dan coba lagi.'}
        </p>

        <div style="display:flex; justify-content:center; gap:16px; margin-bottom:20px;">
          <div style="background:var(--bg-subtle); padding:12px 20px; border-radius:10px; border:1px solid var(--border-color);">
            <div style="font-size:0.8rem; color:var(--text-muted);">Skor Akhir</div>
            <div style="font-size:1.8rem; font-weight:800; color:#4f46e5;">${percent}%</div>
          </div>
          <div style="background:var(--bg-subtle); padding:12px 20px; border-radius:10px; border:1px solid var(--border-color);">
            <div style="font-size:0.8rem; color:var(--text-muted);">Jawaban Benar</div>
            <div style="font-size:1.8rem; font-weight:800; color:var(--text-main);">${score} / ${total}</div>
          </div>
        </div>

        <div style="display:flex; justify-content:center; gap:12px;">
          <button class="btn btn-outline" onclick="AppCh5.closeModal('quiz-result-modal'); document.getElementById('quiz-section').scrollIntoView({behavior:'smooth'});">
            🔍 Lihat Pembahasan
          </button>
          <button class="btn btn-primary" style="background:#4f46e5; border-color:#4f46e5;" onclick="AppCh5.resetQuiz(); AppCh5.closeModal('quiz-result-modal');">
            ↺ Coba Kuis Lagi
          </button>
        </div>
      </div>
    `;

    this.openModal('quiz-result-modal');
  },

  // Classroom Timer
  initTimer() {
    const btnToggle = document.getElementById('btn-timer-toggle');
    const btnReset = document.getElementById('btn-timer-reset');

    if (btnToggle) {
      btnToggle.addEventListener('click', () => {
        if (this.timerRunning) {
          this.pauseTimer();
        } else {
          this.startTimer();
        }
      });
    }

    if (btnReset) {
      btnReset.addEventListener('click', () => {
        this.resetTimer();
      });
    }

    this.updateTimerDisplay();
  },

  startTimer() {
    this.timerRunning = true;
    const btn = document.getElementById('btn-timer-toggle');
    if (btn) btn.textContent = '⏸';
    this.playTone(800, 0.05);

    this.timerInterval = setInterval(() => {
      if (this.timerSecondsLeft > 0) {
        this.timerSecondsLeft--;
        this.updateTimerDisplay();
      } else {
        this.pauseTimer();
        this.playTone(900, 0.3);
        alert("Waktu praktik kelas 90 menit telah selesai! Evaluasi hasil visualisasi grafik kamu.");
      }
    }, 1000);
  },

  pauseTimer() {
    this.timerRunning = false;
    clearInterval(this.timerInterval);
    const btn = document.getElementById('btn-timer-toggle');
    if (btn) btn.textContent = '▶';
    this.playTone(500, 0.05);
  },

  resetTimer() {
    this.pauseTimer();
    this.timerSecondsLeft = 90 * 60;
    this.updateTimerDisplay();
    this.playTone(600, 0.05);
  },

  updateTimerDisplay() {
    const mins = Math.floor(this.timerSecondsLeft / 60);
    const secs = this.timerSecondsLeft % 60;
    const el = document.getElementById('timer-countdown');
    if (el) {
      el.textContent = `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
    }
  },

  // Practice Datasets Modal
  renderDatasetsList() {
    const listEl = document.getElementById('datasets-list-container');
    if (!listEl) return;

    let html = '';
    CHAPTER_5_DATA.sampleDatasets.forEach((ds) => {
      html += `
        <div class="dataset-card" style="border:1px solid var(--border-color); border-radius:8px; padding:14px; margin-bottom:12px; background:var(--bg-surface);">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:6px;">
            <div>
              <h4 style="font-size:0.95rem; font-weight:700; color:var(--text-main); margin-bottom:2px;">${ds.name}</h4>
              <span style="font-size:0.75rem; color:#4f46e5; font-family:var(--font-mono);">${ds.filename}</span>
            </div>
            <button class="btn btn-outline" style="padding:4px 12px; font-size:0.8rem;" onclick="AppCh5.downloadCsv('${ds.id}')">
              ⬇️ Unduh CSV
            </button>
          </div>
          <p style="font-size:0.82rem; color:var(--text-muted); line-height:1.4;">${ds.description}</p>
        </div>
      `;
    });

    listEl.innerHTML = html;
  },

  downloadCsv(datasetId) {
    const ds = CHAPTER_5_DATA.sampleDatasets.find(d => d.id === datasetId);
    if (!ds) return;

    const blob = new Blob([ds.csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", ds.filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    this.playTone(850, 0.08);
  },

  bindGlobalEvents() {
    // Theme toggle
    const btnTheme = document.getElementById('btn-toggle-theme');
    if (btnTheme) btnTheme.addEventListener('click', () => this.toggleTheme());

    // Sound toggle
    const btnSound = document.getElementById('btn-toggle-sound');
    if (btnSound) btnSound.addEventListener('click', () => this.toggleSound());

    // Mobile menu toggle
    const btnMobile = document.getElementById('btn-mobile-menu');
    const sidebar = document.getElementById('app-sidebar');
    if (btnMobile && sidebar) {
      btnMobile.addEventListener('click', () => {
        sidebar.classList.toggle('mobile-open');
        this.playTone(600, 0.04);
      });
    }

    // Modal close buttons
    const closeBtns = document.querySelectorAll('.modal-close-btn');
    closeBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        const modal = btn.closest('.modal-overlay');
        if (modal) this.closeModal(modal.id);
      });
    });

    // Close modal on click backdrop
    const modals = document.querySelectorAll('.modal-overlay');
    modals.forEach(m => {
      m.addEventListener('click', (e) => {
        if (e.target === m) this.closeModal(m.id);
      });
    });

    // Shortcuts modal button
    const btnShortcuts = document.getElementById('btn-open-shortcuts');
    if (btnShortcuts) {
      btnShortcuts.addEventListener('click', () => this.openModal('shortcuts-modal'));
    }

    // Datasets modal button
    const btnDatasets = document.getElementById('btn-open-datasets');
    if (btnDatasets) {
      btnDatasets.addEventListener('click', () => {
        this.renderDatasetsList();
        this.openModal('datasets-modal');
      });
    }

    // Submit Quiz button
    const btnSubmitQuiz = document.getElementById('btn-submit-quiz');
    if (btnSubmitQuiz) {
      btnSubmitQuiz.addEventListener('click', () => this.submitQuiz());
    }
  }
};

window.AppCh5 = AppCh5;

// Launch on DOM ready
document.addEventListener('DOMContentLoaded', () => {
  AppCh5.init();
});

