/**
 * MAIN APP CONTROLLER - ICT EXCEL CHAPTER 3
 * Tables, Table Styles, Subtotals, and Data Formatting
 * Orchestrates navigation, curriculum rendering, quizzes, downloads, audio, and teacher tools
 */

const AppCh3 = {
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
    if (window.SpreadsheetSimulatorsCh3) {
      window.SpreadsheetSimulatorsCh3.initAll();
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

    // Completed topics for Chapter 3
    const savedCompleted = localStorage.getItem('ict_excel_ch3_completed');
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
    localStorage.setItem('ict_excel_ch3_completed', JSON.stringify([...this.completedTopics]));
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
          <div class="nav-link-left">
            <span class="nav-badge-num">GEN</span>
            <span>General Theory</span>
          </div>
        </a>
      </li>
    `;

    CHAPTER_3_DATA.topics.forEach(t => {
      const isDone = this.completedTopics.has(t.id);
      html += `
        <li>
          <a href="#${t.id}" class="nav-link ${isDone ? 'completed' : ''}" data-target="${t.id}">
            <div class="nav-link-left">
              <span class="nav-badge-num">${t.number}</span>
              <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:180px;">${t.title}</span>
            </div>
            <span class="nav-check-icon">✓</span>
          </a>
        </li>
      `;
    });

    html += `
      <li style="margin-top:10px; padding-top:10px; border-top:1px solid var(--border-color);">
        <a href="#quiz-section" class="nav-link" data-target="quiz-section">
          <div class="nav-link-left">
            <span class="nav-badge-num" style="background:var(--accent-amber); color:white;">Q</span>
            <span style="font-weight:700;">Kuis & Evaluasi</span>
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
        if (/^(\*|-)\s+/.test(line)) {
          if (orderedItems.length > 0) flushLists();
          const clean = line.replace(/^(\*|-)\s+/, '');
          bulletItems.push(this.formatInlineMarkdown(clean));
        } else if (/^\d+\.\s+/.test(line)) {
          if (bulletItems.length > 0) flushLists();
          const clean = line.replace(/^\d+\.\s+/, '');
          orderedItems.push(this.formatInlineMarkdown(clean));
        } else {
          flushLists();
          html += `<p>${this.formatInlineMarkdown(line)}</p>`;
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

    CHAPTER_3_DATA.topics.forEach((topic) => {
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
            <div style="background:var(--bg-subtle); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:16px; margin-top:16px;">
              <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:12px;">
                <div>
                  <h4 style="font-family:var(--font-heading); font-size:1.05rem; font-weight:700; color:var(--excel-green); display:flex; align-items:center; gap:6px;">
                    <span>🖥️</span> ${dt.title}
                  </h4>
                  <div style="font-size:0.8rem; color:var(--text-muted); margin-top:2px;">
                    Rekomendasi Nama Sheet: <code style="background:var(--bg-surface); padding:2px 6px; border-radius:4px; font-weight:700; color:var(--excel-green); border:1px solid var(--border-color);">${dt.recommendedSheetName}</code>
                  </div>
                </div>
                <button class="btn btn-outline" style="padding:6px 14px; font-size:0.82rem; font-weight:700;" onclick="App.copyRichTable(null, App.getTopicSampleData('${topic.id}'), '${dt.recommendedSheetName}', '${topic.title}')">
                  📋 Salin Tabel Latihan ke Excel (Ctrl+V)
                </button>
              </div>

              <div style="margin-bottom:14px;">
                <div style="font-size:0.82rem; font-weight:700; color:var(--text-main); margin-bottom:6px;">Langkah-Langkah Praktik di Excel PC/Laptop:</div>
                <ol style="margin:0; padding-left:20px; font-size:0.88rem; line-height:1.6; color:var(--text-main);">
                  ${dt.steps.map(step => `<li style="margin-bottom:6px;">${this.formatInlineMarkdown(step)}</li>`).join('')}
                </ol>
              </div>

              ${dt.sampleHtml ? `
              <details style="margin-top:10px; font-size:0.84rem;">
                <summary style="cursor:pointer; color:var(--excel-green); font-weight:600; padding:4px 0;">
                  👁️ Lihat Pratinjau Tabel Latihan
                </summary>
                <div style="overflow-x:auto; margin-top:10px; border-radius:6px; border:1px solid var(--border-color); background:var(--bg-surface); padding:12px;">
                  ${dt.sampleHtml}
                </div>
              </details>
              ` : ''}
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
    const topic = CHAPTER_3_DATA.topics.find(t => t.id === topicId);
    return topic && topic.desktopTutorial ? topic.desktopTutorial.sampleData : '';
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
    const total = CHAPTER_3_DATA.topics.length;
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
              alert("⏰ Waktu sesi pembelajaran ICT Chapter 3 telah selesai!");
            }
          }, 1000);
        }
      });
    }

    if (btnResetTimer) {
      btnResetTimer.addEventListener('click', () => {
        this.playTone(500, 0.05);
        clearInterval(this.timerInterval);
        this.timerRunning = false;
        this.timerSecondsLeft = 90 * 60;
        if (btnToggleTimer) btnToggleTimer.innerText = "▶";
        updateDisplay();
      });
    }
  },

  // Quiz Engine
  initQuiz() {
    const quizContainer = document.getElementById('quiz-questions-list');
    if (!quizContainer) return;

    let userAnswers = {}; // questionId -> selectedIndex
    const totalQuestions = CHAPTER_3_DATA.quiz.length;

    function updateQuizProgress() {
      const answeredCount = Object.keys(userAnswers).length;
      const percent = Math.round((answeredCount / totalQuestions) * 100);
      
      const progText = document.getElementById('quiz-progress-text');
      if (progText) {
        progText.innerText = `Progres Pengerjaan: ${answeredCount} / ${totalQuestions} Soal Terjawab (${percent}%)`;
      }
      
      const progFill = document.getElementById('quiz-progress-fill');
      if (progFill) {
        progFill.style.width = `${percent}%`;
      }

      const badgeStatus = document.getElementById('quiz-badge-status');
      if (badgeStatus) {
        if (answeredCount === totalQuestions) {
          badgeStatus.style.color = 'var(--excel-green)';
          badgeStatus.style.borderColor = 'rgba(16,124,65,0.4)';
          badgeStatus.innerText = 'Status: Siap Dikumpulkan (100%)';
        } else if (answeredCount > 0) {
          badgeStatus.style.color = 'var(--accent-amber)';
          badgeStatus.innerText = `Status: Sedang Mengerjakan (${answeredCount}/${totalQuestions})`;
        } else {
          badgeStatus.style.color = 'var(--text-muted)';
          badgeStatus.innerText = 'Status: Belum Dikumpulkan';
        }
      }
    }

    function renderQuestions() {
      quizContainer.innerHTML = CHAPTER_3_DATA.quiz.map((q, qIdx) => `
        <div class="quiz-card" id="quiz-card-${q.id}">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
            <span style="font-family:var(--font-mono); font-size:0.8rem; font-weight:700; color:var(--excel-green); background:var(--excel-green-light); padding:3px 10px; border-radius:4px; border:1px solid rgba(16,124,65,0.2);">
              Soal ${qIdx + 1} dari ${totalQuestions} • Topik ${q.topicRef}
            </span>
            <span id="quiz-status-${q.id}" style="font-size:0.82rem; font-weight:700;">
              ${userAnswers[q.id] !== undefined ? `<span style="color:var(--excel-green);">✓ Pilihan Tersimpan (${String.fromCharCode(65 + userAnswers[q.id])})</span>` : `<span style="color:var(--text-muted); font-size:0.75rem;">Belum dijawab</span>`}
            </span>
          </div>

          <div class="quiz-question-text">
            ${q.question}
          </div>

          <div class="quiz-options">
            ${q.options.map((opt, optIdx) => `
              <button type="button" class="quiz-option-btn ${userAnswers[q.id] === optIdx ? 'selected' : ''}" data-qid="${q.id}" data-optid="${optIdx}">
                <span class="quiz-opt-letter">${String.fromCharCode(65 + optIdx)}.</span>
                <span class="quiz-opt-text">${opt}</span>
                <span class="quiz-opt-radio"></span>
              </button>
            `).join('')}
          </div>

          <div class="quiz-explanation-box" id="quiz-explain-${q.id}">
            <div style="font-weight:700; color:var(--excel-green); margin-bottom:4px; display:flex; align-items:center; gap:6px;">
              <span>💡 Pembahasan Kunci Jawaban:</span>
            </div>
            <div>${q.explanation}</div>
          </div>
        </div>
      `).join('');

      // Add click listeners to options
      quizContainer.querySelectorAll('.quiz-option-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          const qId = parseInt(btn.getAttribute('data-qid'));
          const optId = parseInt(btn.getAttribute('data-optid'));
          userAnswers[qId] = optId;

          // Update active style in question card
          const card = document.getElementById(`quiz-card-${qId}`);
          if (card) {
            card.querySelectorAll('.quiz-option-btn').forEach(b => b.classList.remove('selected'));
            btn.classList.add('selected');
            
            const statusText = document.getElementById(`quiz-status-${qId}`);
            if (statusText) {
              statusText.innerHTML = `<span style="color:var(--excel-green);">✓ Pilihan Tersimpan (${String.fromCharCode(65 + optId)})</span>`;
            }
          }

          AppCh3.playTone(750, 0.05);
          updateQuizProgress();
        });
      });

      updateQuizProgress();
    }

    renderQuestions();

    // Submit Quiz button
    const btnSubmit = document.getElementById('btn-submit-quiz');
    if (btnSubmit) {
      btnSubmit.onclick = () => {
        const answeredCount = Object.keys(userAnswers).length;

        if (answeredCount === 0) {
          alert("Silakan pilih minimal 1 jawaban terlebih dahulu sebelum mengumpulkan evaluasi kuis!");
          return;
        }

        if (answeredCount < totalQuestions) {
          const confirmed = confirm(
            `Anda baru menjawab ${answeredCount} dari ${totalQuestions} soal.\n\n` +
            `Soal yang belum dijawab akan dihitung salah (skor 0).\n` +
            `Apakah Anda yakin ingin menyelesaikan dan menghitung skor sekarang?`
          );
          if (!confirmed) return;
        }

        // Calculate score
        let correctCount = 0;
        CHAPTER_3_DATA.quiz.forEach(q => {
          const selected = userAnswers[q.id];
          const isCorrect = selected === q.correctIndex;
          if (isCorrect) correctCount++;

          const card = document.getElementById(`quiz-card-${q.id}`);
          const explainBox = document.getElementById(`quiz-explain-${q.id}`);
          const statusText = document.getElementById(`quiz-status-${q.id}`);

          if (explainBox) explainBox.style.display = 'block';

          if (card) {
            const buttons = card.querySelectorAll('.quiz-option-btn');
            buttons.forEach(b => {
              b.disabled = true;
              b.style.cursor = 'default';
              const bOptId = parseInt(b.getAttribute('data-optid'));
              if (bOptId === q.correctIndex) {
                b.classList.add('selected-correct');
              } else if (bOptId === selected && !isCorrect) {
                b.classList.add('selected-wrong');
              }
            });
          }

          if (statusText) {
            if (selected === undefined) {
              statusText.innerHTML = `<span style="color:#f59e0b; font-weight:700;">⚠️ Tidak Dijawab (0)</span>`;
            } else if (isCorrect) {
              statusText.innerHTML = `<span style="color:var(--excel-green); font-weight:700;">✔ Benar (+${Math.round(100/totalQuestions)})</span>`;
            } else {
              statusText.innerHTML = `<span style="color:var(--accent-rose); font-weight:700;">✕ Salah (0)</span>`;
            }
          }
        });

        const finalScore = Math.round((correctCount / totalQuestions) * 100);

        if (finalScore >= 75) {
          AppCh3.playTone(850, 0.1);
          setTimeout(() => AppCh3.playTone(1100, 0.15), 100);
        } else {
          AppCh3.playTone(400, 0.15);
        }

        // Show Score Modal
        const modalBox = document.getElementById('quiz-result-modal-box');
        if (modalBox) {
          modalBox.innerHTML = `
            <div class="modal-body" style="padding:28px 24px; text-align:center;">
              <div style="font-size:3.5rem; margin-bottom:8px;">${finalScore >= 75 ? '🏆' : '📚'}</div>
              <div class="quiz-score-badge">${finalScore}</div>
              <div style="display:inline-block; font-size:0.85rem; font-weight:700; padding:4px 14px; border-radius:999px; margin-bottom:14px; background:${finalScore >= 75 ? '#dcfce7' : '#fee2e2'}; color:${finalScore >= 75 ? '#166534' : '#991b1b'};">
                ${finalScore >= 75 ? 'STATUS: LULUS / TUNTAS (KKM 75)' : 'STATUS: PERLU PENDALAMAN (REMEDIAL)'}
              </div>
              <h3 style="font-family:var(--font-heading); font-size:1.35rem; font-weight:800; margin-bottom:8px; color:var(--text-main);">
                ${finalScore >= 80 ? 'Luar Biasa! Kompetensi Sangat Memuaskan' : finalScore >= 75 ? 'Bagus! Kompetensi Tuntas KKM' : 'Tetap Semangat, Pelajari Lagi Materi!'}
              </h3>
              <p style="color:var(--text-muted); font-size:0.92rem; margin-bottom:24px; line-height:1.5;">
                Anda menjawab benar <strong>${correctCount} dari ${totalQuestions} soal</strong> evaluasi Microsoft Excel Chapter 3.
              </p>
              <div style="display:flex; justify-content:center; gap:12px; flex-wrap:wrap;">
                <button class="btn btn-primary" onclick="App.closeModal('quiz-result-modal'); document.getElementById('quiz-section').scrollIntoView({behavior:'smooth'});" style="padding:10px 22px;">
                  📖 Lihat Pembahasan Kunci Jawaban
                </button>
                <button class="btn btn-outline" onclick="App.copyQuizReport(${finalScore}, ${correctCount}, ${totalQuestions})" style="padding:10px 20px;">
                  📋 Salin Nilai Laporan
                </button>
              </div>
            </div>
          `;
          AppCh3.openModal('quiz-result-modal');
        }
      };
    }

    // Reset Quiz
    const btnResetQuiz = document.getElementById('btn-reset-quiz');
    if (btnResetQuiz) {
      btnResetQuiz.onclick = () => {
        userAnswers = {};
        renderQuestions();
        AppCh3.playTone(600, 0.05);
        const quizSec = document.getElementById('quiz-section');
        if (quizSec) quizSec.scrollIntoView({ behavior: 'smooth' });
      };
    }
  },

  copyQuizReport(score, correct, total) {
    const text = `LAPORAN EVALUASI SISWA - ICT EXCEL CHAPTER 3\n` +
      `Materi: Tables, Table Styles, Subtotals, and Data Formatting\n` +
      `Skor Akhir: ${score} / 100\n` +
      `Jawaban Benar: ${correct} dari ${total} Soal\n` +
      `Status KKM (75): ${score >= 75 ? 'LULUS / TUNTAS' : 'REMEDIAL'}\n` +
      `Tanggal: ${new Date().toLocaleDateString('id-ID')}`;

    navigator.clipboard.writeText(text).then(() => {
      alert("Laporan nilai berhasil disalin ke clipboard! Siap ditempel ke catatan tugas atau buku catatan belajarmu.");
    });
  },

  // Modal Controllers
  openModal(modalId) {
    const el = document.getElementById(modalId);
    if (el) {
      el.classList.add('active');
      document.body.style.overflow = 'hidden';
      this.playTone(700, 0.05);
    }
  },

  closeModal(modalId) {
    const el = document.getElementById(modalId);
    if (el) {
      el.classList.remove('active');
      document.body.style.overflow = '';
      this.playTone(500, 0.04);
    }
  },

  // Practice Datasets Download / Copy
  downloadCSV(datasetId) {
    const dataset = CHAPTER_3_DATA.sampleDatasets.find(d => d.id === datasetId);
    if (!dataset) return;

    const blob = new Blob([dataset.csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", dataset.filename);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    this.playTone(900, 0.08);
  },

  // Helper to convert plain TSV / CSV text into styled HTML table with green headers and gridlines
  textToRichHtml(textData) {
    if (!textData) return '';
    const lines = textData.trim().split('\n');
    if (lines.length === 0) return '';

    const isCsv = lines[0].includes(',') && !lines[0].includes('\t');
    const separator = isCsv ? ',' : '\t';

    let html = `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;border:1px solid #107c41;">\n`;
    
    // Header Row
    const headers = lines[0].split(separator);
    html += `  <thead>\n    <tr bgcolor="#107c41" style="background-color:#107c41;color:#ffffff;font-weight:bold;">\n`;
    headers.forEach(h => {
      html += `      <th bgcolor="#107c41" style="border:1px solid #0d6836;padding:6px 12px;background-color:#107c41;color:#ffffff;font-weight:bold;text-align:left;">${h.trim()}</th>\n`;
    });
    html += `    </tr>\n  </thead>\n  <tbody>\n`;

    // Data Rows
    for (let i = 1; i < lines.length; i++) {
      const row = lines[i].split(separator);
      const isZebra = i % 2 === 0;
      const rowBg = isZebra ? '#f9fbf9' : '#ffffff';
      html += `    <tr bgcolor="${rowBg}" style="background-color:${rowBg};">\n`;
      row.forEach(cell => {
        const val = cell.trim();
        const isNum = !isNaN(Number(val.replace(/[^0-9.-]+/g, ''))) && val.length > 0 && !val.startsWith('00') && !val.startsWith('PRD-') && !val.startsWith('SYS-') && !val.startsWith('Q');
        const align = isNum ? 'right' : 'left';
        html += `      <td style="border:1px solid #d4d4d4;padding:5px 10px;text-align:${align};">${val}</td>\n`;
      });
      html += `    </tr>\n`;
    }
    html += `  </tbody>\n</table>`;
    return html;
  },

  // Rich Clipboard Copy Engine
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

    // 3. Fallback to writeText
    if (!copied && navigator.clipboard && navigator.clipboard.writeText) {
      try {
        await navigator.clipboard.writeText(plainContent || '');
        copied = true;
      } catch (err3) {
        console.error('All clipboard methods failed:', err3);
      }
    }

    if (copied) {
      this.playTone(850, 0.08);
      setTimeout(() => this.playTone(1100, 0.12), 80);
      this.showCopySuccessModal(topicTitle, sheetName);
    } else {
      alert("Gagal menyalin otomatis. Silakan blok tabel pratinjau secara manual dan tekan Ctrl + C.");
    }

    return copied;
  },

  showCopySuccessModal(topicTitle, sheetName) {
    const modal = document.getElementById('copy-success-modal');
    const modalBody = document.getElementById('copy-success-modal-body');
    if (!modal || !modalBody) return;

    modalBody.innerHTML = `
      <div style="text-align:center; margin-bottom:16px;">
        <div style="width:52px; height:52px; border-radius:50%; background:#dcfce7; color:#166534; display:flex; align-items:center; justify-content:center; font-size:1.6rem; font-weight:bold; margin:0 auto 10px; border:2px solid #86efac;">
          ✓
        </div>
        <h4 style="font-family:var(--font-heading); font-size:1.2rem; font-weight:800; color:var(--text-main); margin-bottom:4px;">
          Tabel & Format Desain Lengkap Berhasil Disalin!
        </h4>
        <p style="font-size:0.86rem; color:var(--text-muted);">
          Data tabel ${topicTitle ? `untuk <strong>${topicTitle}</strong>` : ''} sudah tersalin ke clipboard lengkap dengan header hijau, border garis, dan warna sel.
        </p>
      </div>

      <div style="background:var(--excel-green-light); border:1.5px solid rgba(16,124,65,0.3); border-radius:var(--radius-md); padding:16px; margin-bottom:16px;">
        <div style="font-size:0.8rem; font-weight:800; color:var(--excel-green); letter-spacing:0.05em; text-transform:uppercase; margin-bottom:10px; display:flex; align-items:center; gap:6px;">
          <span>🎯</span> CARA MENEMPELKAN DI EXCEL AGAR DESAIN PERSIS SAMA:
        </div>
        <ol style="margin:0; padding-left:18px; font-size:0.88rem; line-height:1.6; color:var(--text-main);">
          <li style="margin-bottom:6px;">
            Buka aplikasi Microsoft Excel, lalu buat Sheet Baru bernama: <br>
            <code style="background:var(--bg-surface); padding:2px 8px; border-radius:4px; font-weight:700; color:var(--excel-green); border:1px solid rgba(16,124,65,0.35);">${sheetName}</code>
          </li>
          <li style="margin-bottom:6px;">
            Klik sel <strong>A1 SATU KALI SAJA</strong> di sheet tersebut.<br>
            <span style="font-size:0.8rem; color:#dc2626; font-weight:700;">
              ⚠️ PENTING: Jangan klik ganda / jangan sampai kursor berkedip di dalam sel!
            </span>
          </li>
          <li>
            Tekan <strong>Ctrl + V</strong> di keyboard.
          </li>
        </ol>
      </div>

      <div style="background:var(--bg-subtle); border-radius:var(--radius-sm); padding:10px 14px; font-size:0.8rem; color:var(--text-muted); line-height:1.5; margin-bottom:18px;">
        💡 <strong>Mengapa sebelumnya desain hilang?</strong> Ketika kursor teks berkedip di dalam sel (klik ganda), Excel menganggap kamu sedang mengetik teks polos biasa sehingga menolak format warna tabel. Dengan mengklik sel A1 <strong>satu kali saja</strong>, Excel akan menempelkan seluruh desain visual tabel secara utuh!
      </div>

      <div style="display:flex; justify-content:flex-end; gap:10px;">
        <button class="btn btn-primary" onclick="App.closeModal('copy-success-modal')" style="padding:8px 22px; font-size:0.9rem;">
          Siap, Buka Excel & Tempelkan (Ctrl + V)
        </button>
      </div>
    `;

    this.openModal('copy-success-modal');
  },

  copyDataset(datasetId) {
    const dataset = CHAPTER_3_DATA.sampleDatasets.find(d => d.id === datasetId);
    if (!dataset) return;

    const richHtml = this.textToRichHtml(dataset.csvContent);
    this.copyRichTable(richHtml, dataset.csvContent, dataset.id, dataset.name);
  },

  // Global UI Events
  bindGlobalEvents() {
    // Theme toggle button
    const btnTheme = document.getElementById('btn-toggle-theme');
    if (btnTheme) btnTheme.addEventListener('click', () => this.toggleTheme());

    // Sound toggle button
    const btnSound = document.getElementById('btn-toggle-sound');
    if (btnSound) btnSound.addEventListener('click', () => this.toggleSound());

    // Presenter mode button in topbar
    const btnPresent = document.getElementById('btn-open-presentation');
    if (btnPresent) {
      btnPresent.addEventListener('click', () => {
        this.playTone(800, 0.08);
        if (window.PresentationMode) {
          window.PresentationMode.open(0);
        }
      });
    }

    // Mobile Hamburger
    const btnMenu = document.getElementById('btn-mobile-menu');
    const sidebar = document.getElementById('app-sidebar');
    if (btnMenu && sidebar) {
      btnMenu.addEventListener('click', () => {
        sidebar.classList.toggle('open');
      });
    }

    // Modal Triggers
    const btnShortcuts = document.getElementById('btn-open-shortcuts');
    if (btnShortcuts) {
      btnShortcuts.addEventListener('click', () => this.openModal('shortcuts-modal'));
    }

    const btnDatasets = document.getElementById('btn-open-datasets');
    if (btnDatasets) {
      btnDatasets.addEventListener('click', () => this.openModal('datasets-modal'));
    }

    // Modal Close buttons
    document.querySelectorAll('.modal-close-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const modal = btn.closest('.modal-overlay');
        if (modal) this.closeModal(modal.id);
      });
    });

    // Close modal when clicking outside box
    document.querySelectorAll('.modal-overlay').forEach(modal => {
      modal.addEventListener('click', (e) => {
        if (e.target === modal) {
          this.closeModal(modal.id);
        }
      });
    });

    // Render Datasets in Modal
    const datasetModalBody = document.getElementById('datasets-list-container');
    if (datasetModalBody) {
      datasetModalBody.innerHTML = CHAPTER_3_DATA.sampleDatasets.map(ds => `
        <div style="background:var(--bg-subtle); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:16px; margin-bottom:14px;">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
            <h4 style="font-family:var(--font-heading); font-size:1.05rem; font-weight:700;">${ds.name}</h4>
            <span style="font-family:var(--font-mono); font-size:0.75rem; background:var(--bg-surface); padding:2px 8px; border-radius:4px; border:1px solid var(--border-color);">
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
  }
};

window.App = AppCh3;
window.AppCh3 = AppCh3;

// Bootstrap on DOMContentLoaded
document.addEventListener('DOMContentLoaded', () => {
  AppCh3.init();
});
