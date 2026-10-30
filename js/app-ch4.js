/**
 * MAIN APP CONTROLLER - ICT EXCEL CHAPTER 4
 * Excel Formulas and Functions
 * Orchestrates navigation, curriculum rendering, quizzes, downloads, audio, and teacher tools
 */

const AppCh4 = {
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
    if (window.SpreadsheetSimulatorsCh4) {
      window.SpreadsheetSimulatorsCh4.initAll();
    }

    // Initialize presentation mode
    if (window.PresentationMode) {
      window.PresentationMode.init();
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

    // Completed topics for Chapter 4
    const savedCompleted = localStorage.getItem('ict_excel_ch4_completed');
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
    localStorage.setItem('ict_excel_ch4_completed', JSON.stringify([...this.completedTopics]));
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
          <div class="nav-text-group">
            <span class="nav-title">General Theory</span>
            <span class="nav-sub">Fondasi Rumus & Fungsi</span>
          </div>
        </a>
      </li>
      <li>
        <a href="#beginner-quickstart" class="nav-link" data-target="beginner-quickstart">
          <span class="nav-icon">🔰</span>
          <div class="nav-text-group">
            <span class="nav-title">Panduan Kilat Pemula</span>
            <span class="nav-sub">Aturan Emas, F4, & Error</span>
          </div>
        </a>
      </li>
    `;

    CHAPTER_4_DATA.topics.forEach((topic) => {
      const isDone = this.completedTopics.has(topic.id);
      html += `
        <li>
          <a href="#${topic.id}" class="nav-link" data-target="${topic.id}">
            <span class="nav-badge ${isDone ? 'done' : ''}">${topic.number}</span>
            <div class="nav-text-group">
              <span class="nav-title">${topic.title}</span>
              <span class="nav-sub">${topic.indonesianTitle}</span>
            </div>
            ${isDone ? '<span class="nav-check">✔</span>' : ''}
          </a>
        </li>
      `;
    });

    html += `
      <li style="margin-top:10px;">
        <a href="#quiz-section" class="nav-link" data-target="quiz-section">
          <span class="nav-icon">📝</span>
          <div class="nav-text-group">
            <span class="nav-title">Evaluasi & Kuis Formatif</span>
            <span class="nav-sub">6 Studi Kasus Nyata</span>
          </div>
        </a>
      </li>
    `;

    navList.innerHTML = html;

    // Click handler for links
    navList.querySelectorAll('.nav-link').forEach(link => {
      link.addEventListener('click', (e) => {
        navList.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
        link.classList.add('active');
        this.playTone(700, 0.04);
        
        // Auto close mobile drawer if open
        const sidebar = document.getElementById('app-sidebar');
        if (sidebar && window.innerWidth <= 992) {
          sidebar.classList.remove('open');
        }
      });
    });
  },

  // Markdown Formatter Helpers
  formatInlineMarkdown(text) {
    if (!text) return '';
    return text
      // Bold: **text**
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      // Italic: *text*
      .replace(/\*([^*\n]+)\*/g, '<em>$1</em>')
      // Inline code: `code`
      .replace(/`([^`\n]+)`/g, '<code class="inline-code">$1</code>');
  },

  formatMarkdown(text) {
    if (!text) return '';

    // Split text into paragraphs/sections by double newline
    const paragraphs = text.trim().split(/\n\s*\n/);

    return paragraphs.map(para => {
      const lines = para.split('\n').map(l => l.trim()).filter(Boolean);
      if (lines.length === 0) return '';

      let html = '';
      let bulletItems = [];
      let orderedItems = [];

      const flushLists = () => {
        if (bulletItems.length > 0) {
          html += `<ul class="theory-bullet-list">${bulletItems.map(item => `<li>${item}</li>`).join('')}</ul>`;
          bulletItems = [];
        }
        if (orderedItems.length > 0) {
          html += `<ol class="theory-numbered-list">${orderedItems.map(item => `<li>${item}</li>`).join('')}</ol>`;
          orderedItems = [];
        }
      };

      lines.forEach(line => {
        if (/^###\s+/.test(line)) {
          flushLists();
          html += `<h3 style="font-size:1.1rem; font-weight:700; color:var(--excel-green); margin:16px 0 8px;">${this.formatInlineMarkdown(line.replace(/^###\s+/, ''))}</h3>`;
        } else if (/^(\*|-)\s+/.test(line)) {
          if (orderedItems.length > 0) flushLists();
          const clean = line.replace(/^(\*|-)\s+/, '');
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

    CHAPTER_4_DATA.topics.forEach((topic) => {
      const isDone = this.completedTopics.has(topic.id);
      const dt = topic.desktopTutorial;

      html += `
        <article class="topic-card" id="${topic.id}">
          <header class="topic-header-bar">
            <div class="topic-title-group">
              <div class="topic-num-badge">${topic.number}</div>
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
            <div class="topic-summary-box">
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
                <h4 style="font-family:var(--font-heading); font-size:1.05rem; font-weight:700; color:var(--excel-green); display:flex; align-items:center; gap:8px;">
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
                    <button class="btn btn-outline" style="padding:2px 8px; font-size:0.75rem; margin-left:6px;" onclick="App.copySheetName('${dt.recommendedSheetName}', this)">📋 Salin Nama Sheet</button>
                  </div>
                </div>

                <div style="display:flex; gap:8px; flex-wrap:wrap;">
                  <button class="copy-practice-btn" onclick="App.copyTopicTable('${topic.id}', '${dt.recommendedSheetName}', '${topic.title}', this)">
                    📋 Salin Tabel Latihan ke Excel (Ctrl+V)
                  </button>
                </div>
              </div>

              <!-- Excel Paste Pro-Tip Box -->
              <div class="excel-paste-pro-tip">
                <div class="paste-tip-badge">💡 PANDUAN MENEMPEL TABEL DI EXCEL</div>
                <div class="paste-tip-content">
                  <p>Cara menempelkan data latihan agar langsung rapi dengan garis hijau &amp; format sel siap pakai:</p>
                  <ol class="paste-tip-steps">
                    <li>Buka aplikasi Microsoft Excel, buat sheet baru bernama: <code>${dt.recommendedSheetName}</code>.</li>
                    <li>Klik sel <strong>A1 SATU KALI SAJA</strong> (pastikan tidak dobel klik agar kursor tidak berkedip di dalam sel).</li>
                    <li>Tekan <strong>Ctrl + V</strong> di keyboard kamu. Tabel latihan langsung tertempel sempurna!</li>
                  </ol>
                  <div class="paste-warning-note">
                    ⚠️ <strong>Tips Regional Setting:</strong> Jika Excel memunculkan pesan error saat menekan Enter, ganti tanda koma (<code>,</code>) dengan titik koma (<code>;</code>) sesuai pengaturan Windows kamu!
                  </div>
                </div>
              </div>

              <!-- Visual Sample Table Preview -->
              ${dt.sampleHtml ? `
              <div class="desktop-sample-table-container">
                <div class="sample-table-header-info">
                  <span class="sample-table-badge">
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
                      formulaButton = `<button class="copy-formula-btn" onclick="App.copyFormulaText('${cleanFormula.replace(/'/g, "\\'")}', this)" title="Salin rumus ${cleanFormula}">📋 Salin Rumus</button>`;
                    }
                    return `
                      <li class="desktop-step-item">
                        <span class="step-num-badge">Langkah ${idx + 1}</span>
                        <div style="flex-grow:1; line-height:1.65;">
                          ${this.formatInlineMarkdown(step)}
                          ${formulaButton}
                        </div>
                      </li>
                    `;
                  }).join('')}
                </ul>
              </div>

              <!-- Expected Result Box -->
              <div class="expected-result-box">
                <strong>🎯 Verifikasi Hasil Akhir:</strong>
                <div style="margin-top:4px;">${this.formatInlineMarkdown(dt.expectedResult)}</div>
              </div>
            </div>
            ` : ''}

            <!-- Student Pro-Tips & Attention Callout -->
            <div class="student-tips-box">
              <div class="student-tips-title">
                <span>💡 Tips Penting & Trik Praktis:</span>
              </div>
              <div class="student-tips-content">
                ${this.formatInlineMarkdown(topic.proTips || topic.teacherTips)}
              </div>
              <div class="common-pitfall-box">
                <strong>⚠️ Perhatian & Kesalahan yang Sering Terjadi:</strong> ${this.formatInlineMarkdown(topic.commonMistakes)}
              </div>
            </div>

            <!-- Exercise Completion Bar -->
            <div class="exercise-action-bar">
              <div class="exercise-label">
                <span>🎯 <strong>Tantangan Mandiri:</strong> ${this.formatInlineMarkdown(topic.exerciseTask)}</span>
              </div>
              <button class="complete-topic-btn ${isDone ? 'is-completed' : ''}" data-topic-id="${topic.id}">
                ${isDone ? '✔ Topik Selesai Dipelajari' : 'Tandai Selesai'}
              </button>
            </div>
          </div>
        </article>
      `;
    });

    container.innerHTML = html;

    // Bind topic complete buttons
    container.querySelectorAll('.complete-topic-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const topicId = btn.getAttribute('data-topic-id');
        this.toggleTopicComplete(topicId, btn);
      });
    });
  },

  getTopicSampleData(topicId) {
    const topic = CHAPTER_4_DATA.topics.find(t => t.id === topicId);
    return topic && topic.desktopTutorial ? topic.desktopTutorial.sampleData : '';
  },

  copyTopicTable(topicId, sheetName, topicTitle, btnEl) {
    const topic = CHAPTER_4_DATA.topics.find(t => t.id === topicId);
    if (!topic || !topic.desktopTutorial) return;
    const dt = topic.desktopTutorial;
    this.copyRichTable(dt.sampleHtml, dt.sampleData, sheetName, topicTitle);
  },

  copyFormulaText(formulaText, btnEl) {
    if (!formulaText) return;
    navigator.clipboard.writeText(formulaText).then(() => {
      this.playTone(850, 0.08);
      if (btnEl) {
        const oldHtml = btnEl.innerHTML;
        btnEl.innerHTML = "✔ Tersalin!";
        btnEl.classList.add("copied");
        setTimeout(() => {
          btnEl.innerHTML = oldHtml;
          btnEl.classList.remove("copied");
        }, 1500);
      }
    }).catch(err => {
      console.warn("Clipboard copy failed", err);
    });
  },

  copySheetName(sheetName, btnEl) {
    if (!sheetName) return;
    navigator.clipboard.writeText(sheetName).then(() => {
      this.playTone(800, 0.08);
      if (btnEl) {
        const oldHtml = btnEl.innerHTML;
        btnEl.innerHTML = "✔ Sheet Tersalin!";
        setTimeout(() => {
          btnEl.innerHTML = oldHtml;
        }, 1500);
      }
    });
  },

  toggleTopicComplete(topicId, btnEl) {
    if (this.completedTopics.has(topicId)) {
      this.completedTopics.delete(topicId);
      btnEl.classList.remove('is-completed');
      btnEl.innerText = "Tandai Selesai";
      this.playTone(450, 0.05);
    } else {
      this.completedTopics.add(topicId);
      btnEl.classList.add('is-completed');
      btnEl.innerText = "✔ Topik Selesai Dipelajari";
      this.playTone(950, 0.1);
    }

    this.saveState();
    this.updateProgressDisplay();
    this.renderSidebarNav();
  },

  updateProgressDisplay() {
    const total = CHAPTER_4_DATA.topics.length;
    const completed = this.completedTopics.size;
    const percentage = Math.round((completed / total) * 100);

    const progressFill = document.getElementById('progress-bar-fill');
    const progressText = document.getElementById('progress-text');
    const heroProgressStat = document.getElementById('hero-stat-progress');

    if (progressFill) progressFill.style.width = `${percentage}%`;
    if (progressText) progressText.innerText = `${completed} / ${total} Topik (${percentage}%)`;
    if (heroProgressStat) heroProgressStat.innerText = `${percentage}%`;
  },

  // Classroom Pacing Timer
  initTimer() {
    const timerDisplay = document.getElementById('timer-countdown');
    const btnToggleTimer = document.getElementById('btn-timer-toggle');
    const btnResetTimer = document.getElementById('btn-timer-reset');

    const updateDisplay = () => {
      if (!timerDisplay) return;
      const mins = Math.floor(this.timerSecondsLeft / 60);
      const secs = this.timerSecondsLeft % 60;
      timerDisplay.innerText = `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
    };

    updateDisplay();

    if (btnToggleTimer) {
      btnToggleTimer.addEventListener('click', () => {
        this.playTone(600, 0.05);
        if (this.timerRunning) {
          clearInterval(this.timerInterval);
          this.timerRunning = false;
          btnToggleTimer.innerText = "▶";
          btnToggleTimer.title = "Mulai Timer";
        } else {
          this.timerRunning = true;
          btnToggleTimer.innerText = "⏸";
          btnToggleTimer.title = "Jeda Timer";
          this.timerInterval = setInterval(() => {
            if (this.timerSecondsLeft > 0) {
              this.timerSecondsLeft--;
              updateDisplay();
            } else {
              clearInterval(this.timerInterval);
              this.timerRunning = false;
              btnToggleTimer.innerText = "▶";
              alert("⏰ Waktu sesi pembelajaran ICT Chapter 4 telah selesai!");
            }
          }, 1000);
        }
      });
    }

    if (btnResetTimer) {
      btnResetTimer.addEventListener('click', () => {
        this.playTone(450, 0.08);
        clearInterval(this.timerInterval);
        this.timerRunning = false;
        this.timerSecondsLeft = 90 * 60;
        if (btnToggleTimer) {
          btnToggleTimer.innerText = "▶";
          btnToggleTimer.title = "Mulai Timer";
        }
        updateDisplay();
      });
    }
  },

  // Quiz Engine
  selectedAnswers: {},
  quizSubmitted: false,

  initQuiz() {
    const container = document.getElementById('quiz-questions-list');
    if (!container) return;

    let html = "";
    CHAPTER_4_DATA.quizQuestions.forEach((q, idx) => {
      html += `
        <div class="quiz-question-card" id="quiz-card-${q.id}">
          <div class="quiz-question-header">
            <span class="quiz-q-num">Soal ${idx + 1}</span>
            <span class="quiz-q-topic">Ref. Topik ${q.topicRef}</span>
          </div>

          <div class="quiz-question-text">
            ${q.question}
          </div>

          <div class="quiz-options-list">
            ${q.options.map((opt, optIdx) => `
              <div class="quiz-option" data-qid="${q.id}" data-opt-idx="${optIdx}">
                <div class="quiz-opt-indicator">${String.fromCharCode(65 + optIdx)}</div>
                <div class="quiz-opt-text">${opt}</div>
              </div>
            `).join('')}
          </div>

          <div class="quiz-feedback-box" id="quiz-feedback-${q.id}" style="display:none;"></div>
        </div>
      `;
    });

    container.innerHTML = html;

    // Bind option click
    container.querySelectorAll('.quiz-option').forEach(el => {
      el.addEventListener('click', () => {
        if (this.quizSubmitted) return;
        const qid = parseInt(el.getAttribute('data-qid'), 10);
        const optIdx = parseInt(el.getAttribute('data-opt-idx'), 10);

        this.selectedAnswers[qid] = optIdx;
        this.playTone(680, 0.04);

        const parentCard = document.getElementById(`quiz-card-${qid}`);
        parentCard.querySelectorAll('.quiz-option').forEach(o => o.classList.remove('selected'));
        el.classList.add('selected');

        this.updateQuizProgressBar();
      });
    });

    // Submit button
    document.getElementById('btn-submit-quiz')?.addEventListener('click', () => {
      this.submitQuiz();
    });

    // Reset button
    document.getElementById('btn-reset-quiz')?.addEventListener('click', () => {
      this.resetQuiz();
    });
  },

  updateQuizProgressBar() {
    const total = CHAPTER_4_DATA.quizQuestions.length;
    const answered = Object.keys(this.selectedAnswers).length;
    const percentage = Math.round((answered / total) * 100);

    const progressFill = document.getElementById('quiz-progress-fill');
    const progressText = document.getElementById('quiz-progress-text');
    const badgeStatus = document.getElementById('quiz-badge-status');

    if (progressFill) progressFill.style.width = `${percentage}%`;
    if (progressText) progressText.innerText = `Progres Pengerjaan: ${answered} / ${total} Soal Terjawab (${percentage}%)`;
    if (badgeStatus) {
      if (answered === total) {
        badgeStatus.innerText = "Status: Siap Dikumpulkan!";
        badgeStatus.style.color = "var(--excel-green)";
        badgeStatus.style.borderColor = "var(--excel-green)";
      } else {
        badgeStatus.innerText = "Status: Belum Dikumpulkan";
        badgeStatus.style.color = "var(--text-muted)";
      }
    }
  },

  submitQuiz() {
    const total = CHAPTER_4_DATA.quizQuestions.length;
    const answered = Object.keys(this.selectedAnswers).length;

    if (answered < total) {
      this.playTone(350, 0.15);
      const confirmSubmit = confirm(`Kamu baru menjawab ${answered} dari ${total} soal. Apakah kamu yakin ingin mengumpulkan sekarang?`);
      if (!confirmSubmit) return;
    }

    this.quizSubmitted = true;
    let correctCount = 0;

    CHAPTER_4_DATA.quizQuestions.forEach(q => {
      const userAns = this.selectedAnswers[q.id];
      const isCorrect = (userAns === q.correctIndex);
      if (isCorrect) correctCount++;

      const card = document.getElementById(`quiz-card-${q.id}`);
      const feedback = document.getElementById(`quiz-feedback-${q.id}`);

      // Highlight options
      card.querySelectorAll('.quiz-option').forEach((optEl, idx) => {
        optEl.classList.remove('correct', 'incorrect');
        if (idx === q.correctIndex) {
          optEl.classList.add('correct');
        } else if (idx === userAns && !isCorrect) {
          optEl.classList.add('incorrect');
        }
      });

      // Show feedback
      if (feedback) {
        feedback.style.display = 'block';
        feedback.className = `quiz-feedback-box ${isCorrect ? 'correct' : 'incorrect'}`;
        feedback.innerHTML = `
          <strong>${isCorrect ? '✔ Jawaban Kamu Benar!' : '✖ Kurang Tepat!'}</strong>
          <p style="margin-top:6px; font-size:0.88rem; line-height:1.5;">${q.explanation}</p>
        `;
      }
    });

    const finalScore = Math.round((correctCount / total) * 100);
    const passed = finalScore >= 75;

    if (passed) {
      this.playTone(800, 0.08);
      setTimeout(() => this.playTone(1100, 0.15), 100);
    } else {
      this.playTone(400, 0.12);
    }

    // Show result modal
    this.showQuizResultModal(finalScore, correctCount, total, passed);
  },

  showQuizResultModal(score, correct, total, passed) {
    const modalBox = document.getElementById('quiz-result-modal-box');
    if (!modalBox) return;

    modalBox.innerHTML = `
      <div class="modal-header">
        <h3 class="modal-title">${passed ? '🎉 Selamat! Hasil Kuis Memuaskan' : '📚 Terus Semangat Belajar!'}</h3>
        <button class="btn-icon modal-close-btn">&times;</button>
      </div>
      <div class="modal-body" style="text-align:center; padding:24px 16px;">
        <div style="font-size:3.5rem; font-weight:800; font-family:var(--font-heading); color:${passed ? 'var(--excel-green)' : '#e11d48'}; line-height:1;">
          ${score}
        </div>
        <div style="font-size:1.05rem; font-weight:600; color:var(--text-main); margin-top:8px;">
          Skor Akhir Evaluasi Mandiri Chapter 4
        </div>
        <p style="color:var(--text-muted); font-size:0.9rem; margin-top:6px;">
          Kamu berhasil menjawab dengan benar <strong>${correct}</strong> dari <strong>${total}</strong> soal. (Standar KKM: 75)
        </p>

        <div style="margin:20px 0; padding:12px; background:var(--bg-subtle); border-radius:8px; font-size:0.85rem; text-align:left;">
          ${passed
            ? '🏆 <strong>Luar Biasa!</strong> Kamu telah menguasai konsep dasar rumus, penguncian absolut $A$1, fungsi statistik dasar, agregasi COUNTIFS/SUMIFS, manipulasi teks, logika IF, dan audit formula.'
            : '💡 <strong>Tips Perbaikan:</strong> Pelajari kembali materi penguncian sel absolut ($A$1) dan urutan parameter fungsi SUMIFS di atas sebelum mencoba kuis kembali.'
          }
        </div>

        <button class="btn btn-primary" onclick="App.closeAllModals()" style="padding:10px 24px; font-size:0.9rem;">
          Lihat Pembahasan Lengkap di Lembar Soal
        </button>
      </div>
    `;

    this.openModal('quiz-result-modal');
  },

  resetQuiz() {
    this.selectedAnswers = {};
    this.quizSubmitted = false;
    this.playTone(550, 0.05);

    CHAPTER_4_DATA.quizQuestions.forEach(q => {
      const card = document.getElementById(`quiz-card-${q.id}`);
      const feedback = document.getElementById(`quiz-feedback-${q.id}`);

      if (card) {
        card.querySelectorAll('.quiz-option').forEach(o => {
          o.classList.remove('selected', 'correct', 'incorrect');
        });
      }
      if (feedback) {
        feedback.style.display = 'none';
        feedback.innerHTML = '';
      }
    });

    this.updateQuizProgressBar();
  },

  // Modal System
  openModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
      modal.classList.add('active');
      document.body.style.overflow = 'hidden';
      this.playTone(700, 0.04);
    }
  },

  closeAllModals() {
    document.querySelectorAll('.modal-overlay').forEach(m => m.classList.remove('active'));
    document.body.style.overflow = '';
  },

  // Rich HTML Table Copy for Native Excel Paste
  textToRichHtml(text) {
    const lines = text.trim().split('\n');
    let html = '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;width:100%;">';
    let isHeader = true;

    lines.forEach(line => {
      if (!line.trim()) return;
      const delimiter = line.includes('\t') ? '\t' : ',';
      const cells = line.split(delimiter).map(c => c.trim().replace(/^"|"$/g, ''));

      if (isHeader) {
        html += '<thead bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;"><tr>';
        cells.forEach(c => {
          html += `<th style="border:1px solid #0d6836;padding:6px 12px;text-align:left;background-color:#107c41;color:#ffffff;">${c}</th>`;
        });
        html += '</tr></thead><tbody>';
        isHeader = false;
      } else {
        html += '<tr>';
        cells.forEach(c => {
          const isNum = !isNaN(Number(c.replace(/\./g, '').replace(/,/g, ''))) && c !== '';
          const align = isNum ? 'right' : 'left';
          html += `<td style="border:1px solid #d4d4d4;padding:5px 12px;text-align:${align};">${c}</td>`;
        });
        html += '</tr>';
      }
    });

    html += '</tbody></table>';
    return html;
  },

  async copyRichTable(htmlContent, plainContent, sheetName = 'Sheet Baru', topicTitle = '') {
    if (!htmlContent && plainContent) {
      htmlContent = this.textToRichHtml(plainContent);
    }

    const fullHtml = `<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style>
  table { border-collapse: collapse; font-family: Calibri, 'Segoe UI', Arial, sans-serif; font-size: 11pt; }
  th { background-color: #107c41; color: #ffffff; font-weight: bold; border: 1px solid #0d6836; padding: 6px 14px; text-align: left; }
  td { border: 1px solid #d4d4d4; padding: 5px 12px; }
</style>
</head>
<body>
${htmlContent}
</body>
</html>`;

    let copied = false;

    // 1. Try modern Clipboard API
    if (navigator.clipboard && window.ClipboardItem) {
      try {
        const item = new ClipboardItem({
          'text/html': new Blob([fullHtml], { type: 'text/html' }),
          'text/plain': new Blob([plainContent || ''], { type: 'text/plain' })
        });
        await navigator.clipboard.write([item]);
        copied = true;
      } catch (err) {
        console.warn('ClipboardItem write failed, trying fallback...', err);
      }
    }

    // 2. DOM selection fallback
    if (!copied) {
      try {
        const container = document.createElement('div');
        container.innerHTML = fullHtml;
        container.style.position = 'fixed';
        container.style.left = '-9999px';
        container.style.top = '-9999px';
        container.style.opacity = '0';
        container.style.pointerEvents = 'none';
        document.body.appendChild(container);

        const range = document.createRange();
        range.selectNode(container);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);

        copied = document.execCommand('copy');
        sel.removeAllRanges();
        document.body.removeChild(container);
      } catch (err2) {
        console.warn('execCommand copy fallback failed:', err2);
      }
    }

    if (copied) {
      this.playSuccessSound();
      this.showCopySuccessModal(sheetName, topicTitle);
    } else {
      alert("Tabel berhasil disalin. Buka Excel lalu tekan Ctrl + V.");
    }
  },

  showCopySuccessModal(sheetName, topicTitle) {
    const modalBody = document.getElementById('copy-success-modal-body');
    if (!modalBody) return;

    modalBody.innerHTML = `
      <div style="background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px; padding:12px 16px; margin-bottom:16px;">
        <div style="font-weight:700; color:#166534; font-size:0.95rem; margin-bottom:4px;">
          ✔ Data Tabel Berhasil Disalin ke Clipboard Komputer!
        </div>
        <div style="font-size:0.85rem; color:#15803d;">
          Tabel telah diformat dalam gaya Microsoft Excel resmi (lengkap dengan warna header hijau dan garis tabel rapi).
        </div>
      </div>

      <div style="font-size:0.88rem; font-weight:700; color:var(--text-main); margin-bottom:8px;">
        Cara Menempelkan (*Paste*) ke Microsoft Excel Laptop/PC:
      </div>

      <ol style="margin:0; padding-left:20px; font-size:0.86rem; line-height:1.7; color:var(--text-main);">
        <li>Buka aplikasi <strong>Microsoft Excel</strong> di laptop atau komputer kamu.</li>
        <li>Buat Sheet baru lalu beri nama: <code style="background:var(--bg-subtle); padding:2px 6px; border-radius:4px; font-weight:700; color:var(--excel-green);">${sheetName}</code>.</li>
        <li>Klik sel <strong>A1</strong> <strong>SATU KALI SAJA</strong> (jangan double click agar kursor tidak masuk ke dalam sel).</li>
        <li>Tekan tombol keyboard: <kbd style="background:var(--bg-subtle); border:1px solid var(--border-color); padding:2px 8px; border-radius:4px; font-weight:700;">Ctrl</kbd> + <kbd style="background:var(--bg-subtle); border:1px solid var(--border-color); padding:2px 8px; border-radius:4px; font-weight:700;">V</kbd>.</li>
        <li>Tabel akan langsung tertempel sempurna dan siap digunakan untuk praktik!</li>
      </ol>

      <div style="display:flex; justify-content:flex-end; margin-top:20px;">
        <button class="btn btn-primary" onclick="App.closeAllModals()" style="padding:8px 20px; font-size:0.85rem;">
          Saya Mengerti, Kembali ke Modul
        </button>
      </div>
    `;

    this.openModal('copy-success-modal');
  },

  // CSV Downloader & Dataset Handling
  downloadCSV(datasetId) {
    const dataset = CHAPTER_4_DATA.sampleDatasets.find(d => d.id === datasetId);
    if (!dataset) return;

    this.playTone(850, 0.08);
    const blob = new Blob([dataset.csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = dataset.filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  },

  copyDataset(datasetId) {
    const dataset = CHAPTER_4_DATA.sampleDatasets.find(d => d.id === datasetId);
    if (!dataset) return;

    const richHtml = this.textToRichHtml(dataset.csvContent);
    this.copyRichTable(richHtml, dataset.csvContent, dataset.id, dataset.name);
  },

  // Global Event Listeners
  bindGlobalEvents() {
    // Theme toggle
    document.getElementById('btn-toggle-theme')?.addEventListener('click', () => {
      this.toggleTheme();
    });

    // Sound toggle
    document.getElementById('btn-toggle-sound')?.addEventListener('click', () => {
      this.toggleSound();
    });

    // Mobile menu toggle
    document.getElementById('btn-mobile-menu')?.addEventListener('click', () => {
      const sidebar = document.getElementById('app-sidebar');
      if (sidebar) {
        sidebar.classList.toggle('open');
        this.playTone(600, 0.04);
      }
    });

    // Datasets modal open
    document.getElementById('btn-open-datasets')?.addEventListener('click', () => {
      this.renderDatasetsModal();
      this.openModal('datasets-modal');
    });

    // Shortcuts modal open
    document.getElementById('btn-open-shortcuts')?.addEventListener('click', () => {
      this.openModal('shortcuts-modal');
    });

    // Close buttons on all modals
    document.querySelectorAll('.modal-close-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        this.closeAllModals();
        this.playTone(450, 0.04);
      });
    });

    // Close modal on click outside box
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
      overlay.addEventListener('click', (e) => {
        if (e.target === overlay) {
          this.closeAllModals();
        }
      });
    });

    // Presentation mode trigger
    document.getElementById('btn-open-presentation')?.addEventListener('click', () => {
      if (window.PresentationMode) {
        window.PresentationMode.open(0);
      }
    });
  },

  renderDatasetsModal() {
    const container = document.getElementById('datasets-list-container');
    if (!container) return;

    container.innerHTML = CHAPTER_4_DATA.sampleDatasets.map(ds => `
      <div style="background:var(--bg-subtle); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:14px 16px; margin-bottom:12px;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; flex-wrap:wrap; gap:8px;">
          <h4 style="font-size:0.95rem; font-weight:700; color:var(--text-main); margin:0;">
            ${ds.name}
          </h4>
          <span style="font-size:0.75rem; font-family:var(--font-mono); background:var(--bg-surface); padding:2px 8px; border-radius:4px; border:1px solid var(--border-color); color:var(--excel-green); font-weight:700;">
            ${ds.filename}
          </span>
        </div>
        <p style="font-size:0.85rem; color:var(--text-muted); margin-bottom:12px;">${ds.description}</p>
        <div style="display:flex; gap:10px; flex-wrap:wrap;">
          <button class="btn btn-primary" style="padding:6px 14px; font-size:0.82rem;" onclick="App.downloadCSV('${ds.id}')">
            ⬇ Unduh File .CSV
          </button>
          <button class="btn btn-outline" style="padding:6px 14px; font-size:0.82rem;" onclick="App.copyDataset('${ds.id}')">
            📋 Salin ke Clipboard (Siap Paste ke Excel)
          </button>
        </div>
      </div>
    `).join('');
  }
};

window.App = AppCh4;
window.AppCh4 = AppCh4;

// Bootstrap on DOMContentLoaded
document.addEventListener('DOMContentLoaded', () => {
  AppCh4.init();
});
