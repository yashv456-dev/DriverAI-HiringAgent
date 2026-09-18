/**
 * DriverAI Hiring Agent — Frontend Controller & Interactivity
 * Apple VisionOS / Liquid Frosted Glassmorphism Edition
 */

document.addEventListener('DOMContentLoaded', () => {
  // ADDED 2026-09-18: declared up here, not beside applyRoleCount, because
  // VIEW_META below interpolates it at definition time -- a `const` further
  // down put it in the temporal dead zone and threw before any view rendered.
  const JD_COUNT = { n: 0 };
  // ── State Management ─────────────────────────────────────────────────────────
  const state = {
    activeView: 'dashboard',
    selectedFile: null,
    candidates: [],
    roles: [],
    categories: [],
    stats: {},
    currentScoredCandidate: null,
    chatHistory: [],
    compareState: {
      selectedAppIds: [],
      activeTargetRole: '',
      cachedComparison: null,
    },
  };

  // ── Sample Resumes for 1-Click Instant Testing ────────────────────────────────
  const SAMPLE_RESUMES = {
    alex: {
      name: 'Alex Rivera',
      text: `ALEX RIVERA
Email: alex.rivera@example.com | Phone: (415) 555-0143
Location: San Francisco, CA, United States | LinkedIn: linkedin.com/in/alexrivera-ai

SUMMARY
Senior Machine Learning Engineer with 6+ years of experience designing, training, and deploying large-scale deep learning models, LLM fine-tuning pipelines, and RAG architectures in production environments.

EXPERIENCE
Senior ML Engineer — NeuralScale AI, San Francisco, CA (2022 – Present)
- Architected and optimized real-time LLM inference pipelines with TensorRT-LLM and vLLM, reducing latency by 45%.
- Fine-tuned transformer models using LoRA and QLoRA for custom domain retrieval tasks.
- Scaled vector search infrastructure using Milvus and Pinecone across 50M+ document embeddings.

Machine Learning Engineer — Apex Data Systems, Mountain View, CA (2019 – 2022)
- Implemented PyTorch deep learning vision and NLP models for automated document intelligence.
- Built distributed training workflows on AWS EC2 GPU clusters with PyTorch FSDP and Ray.

EDUCATION
Master of Science in Computer Science — Stanford University (2017 – 2019)
Bachelor of Science in Electrical Engineering & CS — UC Berkeley (2013 – 2017)

SKILLS
Machine Learning: PyTorch, TensorFlow, Transformers, HuggingFace, TensorRT, vLLM, LangChain, LlamaIndex, RAG
Languages: Python, C++, CUDA, SQL
Cloud & Ops: AWS, GCP, Docker, Kubernetes, Ray, Triton Inference Server`
    },
    sarah: {
      name: 'Sarah Chen',
      text: `SARAH CHEN
Email: sarah.chen@techmail.io | Phone: (206) 555-0892
Location: Seattle, WA, USA | GitHub: github.com/schen-distributed

SUMMARY
Staff Distributed Systems & Infrastructure Engineer with 8+ years experience architecting high-throughput, low-latency microservices, consensus protocols (Raft), and distributed storage systems in Rust and Go.

EXPERIENCE
Staff Infrastructure Engineer — CloudMesh Systems, Seattle, WA (2021 – Present)
- Designed low-latency distributed event streaming engine in Rust processing 2M+ events/sec with sub-millisecond p99 latency.
- Led migration of 140+ microservices to Kubernetes service mesh with Envoy and Istio.
- Implemented distributed consensus and raft-based state replication for multi-region active-active clusters.

Senior Backend Engineer — DataStream Corp, Austin, TX (2017 – 2021)
- Built high-throughput telemetry ingestion pipeline in Go and Apache Kafka handling 10TB/day.
- Reduced cloud infrastructure footprint by 35% through custom gRPC memory pooling and zero-copy deserialization.

EDUCATION
Bachelor of Science in Computer Science — University of Washington (2013 – 2017)

SKILLS
Languages: Rust, Go (Golang), C++, Python, SQL
Distributed Systems: Raft, Consensus, Kafka, gRPC, Protocol Buffers, Redis, Cassandra
Infrastructure: Kubernetes, Docker, Terraform, AWS, Linux Kernel eBPF, Prometheus, Grafana`
    },
    marco: {
      name: 'Marco Rossi',
      text: `MARCO ROSSI
Email: marco.rossi@milan-tech.it | Phone: +39 02 555 1234
Location: Milan, Lombardy, Italy

SUMMARY
Senior Full Stack Engineer with 5+ years building modern web applications with React, Next.js, and TypeScript.

EXPERIENCE
Senior Software Engineer — Milano Digital Labs, Milan, Italy (2020 – Present)
- Developed responsive web portals using Next.js and Tailwind CSS.
- Maintained REST APIs in Node.js and PostgreSQL.

EDUCATION
Bachelor of Science in Computer Engineering — Politecnico di Milano (2015 – 2019)

SKILLS
React, TypeScript, Node.js, Next.js, CSS, PostgreSQL, Git`
    }
  };

  // ── DOM Elements ─────────────────────────────────────────────────────────────
  const navItems = document.querySelectorAll('.nav-item');
  const viewPanels = document.querySelectorAll('.view-panel');
  const pageTitle = document.getElementById('page-title');
  const pageSubtitle = document.getElementById('page-subtitle');
  const globalSearch = document.getElementById('global-search');
  const btnQuickUpload = document.getElementById('btn-quick-upload');
  const spotlight = document.getElementById('cursor-spotlight');

  // Analyzer Elements
  const dropzoneMain = document.getElementById('dropzone-main');
  const fileInputMain = document.getElementById('file-input-main');
  const dropzoneMini = document.getElementById('dropzone-mini');
  const fileInputMini = document.getElementById('file-input-mini');
  const btnBrowseFile = document.getElementById('btn-browse-file');
  const selectedFileBanner = document.getElementById('selected-file-banner');
  const fileBannerName = document.getElementById('file-banner-name');
  const fileBannerSize = document.getElementById('file-banner-size');
  const btnRemoveFile = document.getElementById('btn-remove-file');
  const btnRunAnalysis = document.getElementById('btn-run-analysis');
  const pasteResumeTextarea = document.getElementById('paste-resume-textarea');
  const chkSaveCandidate = document.getElementById('chk-save-candidate');
  const analysisResults = document.getElementById('analysis-results');
  const analysisLoader = document.getElementById('analysis-loader');
  const analysisContent = document.getElementById('analysis-content');
  const analysisError = document.getElementById('analysis-error');
  const tabUploadFile = document.getElementById('tab-upload-file');
  const tabPasteText = document.getElementById('tab-paste-text');
  const contentUploadFile = document.getElementById('content-upload-file');
  const contentPasteText = document.getElementById('content-paste-text');

  // Candidate Pipeline Elements
  const filterCandSearch = document.getElementById('filter-candidate-search');
  const filterCandStatus = document.getElementById('filter-candidate-status');
  const filterMinScore = document.getElementById('filter-min-score');
  const scoreSliderVal = document.getElementById('score-slider-val');
  const candidatesTableBody = document.getElementById('candidates-table-tbody');
  const btnRefreshCandidates = document.getElementById('btn-refresh-candidates');

  // Roles Elements
  const rolesCardsGrid = document.getElementById('roles-cards-grid');
  const rolesCategoryPills = document.getElementById('roles-category-pills');
  const filterRolesSearch = document.getElementById('filter-roles-search');

  // Copilot Elements
  const copilotMessages = document.getElementById('copilot-messages');
  const copilotInput = document.getElementById('copilot-input');
  const btnSendChat = document.getElementById('btn-send-chat');
  const btnClearChat = document.getElementById('btn-clear-chat');
  const promptChips = document.querySelectorAll('.prompt-chip');

  // Drawer Elements
  const drawerOverlay = document.getElementById('candidate-drawer-overlay');
  const btnCloseDrawer = document.getElementById('btn-close-drawer');
  const drawerName = document.getElementById('drawer-name');
  const drawerStatusBadge = document.getElementById('drawer-status-badge');
  const drawerBody = document.getElementById('drawer-body');

  // ── Dynamic Interactive Cursor Light & Specular Sheen ───────────────────────
  window.addEventListener('mousemove', (e) => {
    if (spotlight) {
      spotlight.style.left = `${e.clientX}px`;
      spotlight.style.top = `${e.clientY}px`;
    }

    const cards = document.querySelectorAll('.glass-card');
    cards.forEach((card) => {
      const rect = card.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      card.style.setProperty('--mouse-card-x', `${x}px`);
      card.style.setProperty('--mouse-card-y', `${y}px`);
    });
  });

  // ── Navigation Router ────────────────────────────────────────────────────────

  const viewMetadata = {
    dashboard: {
      title: 'Executive Dashboard',
      subtitle: `Real-time candidate intelligence and ${JD_COUNT.n}-role matching`,
    },
    analyzer: {
      title: 'AI Resume Analyzer',
      subtitle: 'Upload any candidate resume for instant Gemini extraction & scoring',
    },
    candidates: {
      title: 'Candidate Pipeline',
      subtitle: 'Search, filter, and review stored candidate scorecards',
    },
    compare: {
      title: 'Candidate Comparison Matrix & Skill Radar',
      subtitle: 'Side-by-side 6-axis technical spider radar and Gemini head-to-head executive verdict',
    },
    roles: {
      title: `${JD_COUNT.n} Open Job Roles`,
      subtitle: 'Browse company job descriptions and required skills',
    },
    copilot: {
      title: 'Recruiter Copilot',
      subtitle: 'AI hiring assistant powered by Gemini 3.1 Flash-Lite',
    },
    settings: {
      title: 'System Diagnostics & Settings',
      subtitle: 'Inspect engine health, active models, and configurations',
    },
  };

  function switchView(viewName) {
    state.activeView = viewName;

    // Update Nav
    navItems.forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.view === viewName);
    });

    // Update View Panels
    viewPanels.forEach((panel) => {
      panel.classList.toggle('active', panel.id === `view-${viewName}`);
    });

    // Update Header
    const meta = viewMetadata[viewName] || { title: 'DriverAI Hiring', subtitle: '' };
    pageTitle.textContent = meta.title;
    pageSubtitle.textContent = meta.subtitle;

    // Trigger View Refresh
    if (viewName === 'dashboard') loadStats();
    if (viewName === 'candidates') loadCandidates();
    if (viewName === 'roles') loadRoles();
    if (viewName === 'compare') initCompareView();
  }
  window.switchView = switchView;

  navItems.forEach((btn) => {
    btn.addEventListener('click', () => switchView(btn.dataset.view));
  });

  if (btnQuickUpload) {
    btnQuickUpload.addEventListener('click', () => {
      switchView('analyzer');
      fileInputMain.click();
    });
  }

  // ── Data-source mode switch (added 2026-09-18) ─────────────────────────────
  // Testing   = local SQLite store.
  // Developer = live SharePoint master over Microsoft Graph, read-only.
  // The MS badge is painted from a real Graph probe returned by the server, so it
  // cannot claim "connected" the way the hardcoded diagnostics panel does.
  const modeSelect = document.getElementById('mode-select');
  const msStatusEl = document.getElementById('ms-status');

  function paintMsStatus(ms) {
    if (!msStatusEl || !ms) return;
    let dot = 'text-dim', label = 'MS: not configured';
    if (ms.reachable) { dot = 'text-emerald'; label = 'MS: connected'; }
    else if (ms.configured) { dot = 'text-amber'; label = 'MS: unreachable'; }
    msStatusEl.innerHTML =
      `<i class="fa-solid fa-circle ${dot}" style="font-size: 7px;"></i> ${escapeHtml(label)}`;
    msStatusEl.title = ms.detail || '';
  }


  // ADDED 2026-09-18: the role count was hardcoded as 105 in a dozen places while
  // the cache actually holds 106, and it changes whenever a JD is added to the
  // SharePoint folder. Elements carry their own template in data-jd-count ("{n}
  // Open Roles") and this fills it once /api/health reports the real number.

  function applyRoleCount(n) {
    if (!n) return;
    JD_COUNT.n = n;
    document.querySelectorAll('[data-jd-count]').forEach((el) => {
      el.textContent = el.getAttribute('data-jd-count').replace('{n}', n);
    });
    const ph = document.getElementById('filter-roles-search-ph');
    if (ph) ph.placeholder = `Search ${n} roles by title or required skill (e.g. PyTorch, Rust, AWS)...`;
    const kpi = document.getElementById('kpi-roles-count');
    if (kpi) kpi.textContent = n;
  }

  // ADDED 2026-09-18: paint the AI badge and JD count from the server instead of
  // the hardcoded "Gemini 3.1 Active" / "105 JDs" that used to sit in the markup.
  // A quota-exhausted brain now reads "AI: unavailable", not "Active".
  async function refreshHeaderStatus() {
    const aiEl = document.getElementById('ai-status');
    const jdEl = document.getElementById('jd-count');
    try {
      const d = await (await fetch('/api/health')).json();
      renderDiagnostics(d);
      applyRoleCount(d.roles_loaded);
      if (jdEl) {
        jdEl.innerHTML =
          `<i class="fa-solid fa-briefcase"></i> ${d.roles_loaded} JDs`;
      }
      if (aiEl) {
        const ai = d.ai_engine || {};
        const ok = !!ai.connected;
        aiEl.innerHTML =
          `<i class="fa-solid fa-circle ${ok ? 'text-emerald' : 'text-amber'}" style="font-size: 7px;"></i> ` +
          `${escapeHtml(ai.model || 'AI')}${ok ? '' : ' unavailable'}`;
        aiEl.title = ai.status || '';
      }
    } catch (e) {
      if (aiEl) aiEl.textContent = 'AI: unknown';
    }
  }

  // ADDED 2026-09-18: render System Diagnostics from /api/health.
  // Every row here reports an observed value. Where the truth is bad news -- mail
  // not suppressed, AI unreachable, Microsoft down -- the row says so in amber or
  // red rather than showing a reassuring default.
  function renderDiagnostics(d) {
    const host = document.getElementById('diagnostics-list');
    if (!host || !d) return;
    const ai = d.ai_engine || {}, db = d.database || {}, mail = d.mail || {}, ms = d.microsoft || {};

    const row = (label, ok, text, icon) => `
      <div class="diag-item">
        <span class="diag-label">${escapeHtml(label)}</span>
        <span class="diag-status ${ok ? 'status-ok' : 'text-amber'}">
          <i class="fa-solid ${icon || (ok ? 'fa-circle-check' : 'fa-triangle-exclamation')}"></i>
          ${escapeHtml(text)}
        </span>
      </div>`;

    // Suppressed mail is the SAFE state, so ok=true there. Sending is the state
    // that has to shout, because it is the one that reaches real applicants.
    const mailSending = !mail.suppressed;

    host.innerHTML = [
      row('AI Engine', !!ai.connected,
          ai.connected ? `${ai.model} (active)` : `${ai.model || 'AI'} unreachable — ${ai.status || 'no response'}`),
      row('Job Descriptions Cache', (d.roles_loaded || 0) > 0, `${d.roles_loaded} roles loaded`),
      row('Microsoft / SharePoint', !!ms.reachable,
          ms.reachable ? 'connected' : (ms.configured ? 'configured but unreachable' : 'not configured (local only)')),
      row('Data Source', true,
          d.mode === 'developer' ? 'Developer — live SharePoint' : 'Testing — local store'),
      row('Local Database', !!db.exists,
          db.exists ? `SQLite (${db.name})` : `missing (${db.name})`),
      row('Admin Email', !!mail.admin_email, mail.admin_email || 'not set'),
      row('Outbound Applicant Mail', !mailSending,
          mailSending ? 'SENDING — real applicants will be emailed' : 'suppressed (safe)',
          mailSending ? 'fa-paper-plane' : 'fa-shield'),
    ].join('');
  }

  async function refreshMode() {
    try {
      const r = await fetch('/api/mode');
      const d = await r.json();
      if (modeSelect) modeSelect.value = d.mode;
      paintMsStatus(d.microsoft);
    } catch (e) {
      if (msStatusEl) msStatusEl.textContent = 'MS: unknown';
    }
    refreshHeaderStatus();
  }

  if (modeSelect) {
    modeSelect.addEventListener('change', async () => {
      const wanted = modeSelect.value;
      try {
        const r = await fetch('/api/mode', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mode: wanted }),
        });
        const d = await r.json();
        if (!r.ok) {
          // Server refused (no reachable tenant). Snap back rather than showing
          // an empty table that reads as "no candidates".
          await refreshMode();
          showToast(d.detail || 'Could not switch mode', 'error');
          return;
        }
        paintMsStatus(d.microsoft);
        showToast(`Source: ${wanted === 'developer' ? 'Microsoft SharePoint (live)' : 'local test store'}`);
        loadCandidates();
        loadStats();
      } catch (e) {
        await refreshMode();
        showToast('Could not switch mode', 'error');
      }
    });
  }
  refreshMode();

  // ADDED 2026-09-18: #btn-refresh-health existed in the markup with no listener
  // anywhere, so "Check Health" did nothing at all when clicked.
  const btnRefreshHealth = document.getElementById('btn-refresh-health');
  if (btnRefreshHealth) {
    btnRefreshHealth.addEventListener('click', async () => {
      btnRefreshHealth.disabled = true;
      await refreshMode();          // repaints badges AND diagnostics
      btnRefreshHealth.disabled = false;
      showToast('Health checked');
    });
  }

  const btnViewAllRoles = document.getElementById('btn-view-all-roles');
  if (btnViewAllRoles) {
    btnViewAllRoles.addEventListener('click', () => switchView('roles'));
  }

  const btnViewAllCandidates = document.getElementById('btn-view-all-candidates');
  if (btnViewAllCandidates) {
    btnViewAllCandidates.addEventListener('click', () => switchView('candidates'));
  }

  // ── Global Search ────────────────────────────────────────────────────────────
  if (globalSearch) {
    globalSearch.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        const query = globalSearch.value.trim();
        if (query) {
          switchView('candidates');
          filterCandSearch.value = query;
          loadCandidates();
        }
      }
    });
  }

  // ── 1-Click Sample Resumes Handlers ──────────────────────────────────────────
  document.querySelectorAll('.sample-chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      const sampleKey = chip.dataset.sample;
      const sample = SAMPLE_RESUMES[sampleKey];
      if (!sample) return;

      switchView('analyzer');
      tabPasteText.click();
      pasteResumeTextarea.value = sample.text;
      updateAnalyzerButtonState();

      // Trigger instant analysis for seamless demo
      btnRunAnalysis.click();
    });
  });

  // ── Dashboard Data Loader ───────────────────────────────────────────────────

  async function loadStats() {
    try {
      const res = await fetch('/api/stats');
      if (!res.ok) return;
      const data = await res.json();
      state.stats = data;

      document.getElementById('kpi-total-candidates').textContent = data.total_candidates;
      document.getElementById('kpi-scored-candidates').textContent = data.scored_candidates;
      document.getElementById('kpi-avg-score').textContent = `${data.average_score}%`;
      document.getElementById('kpi-roles-count').textContent = data.open_roles_count;
      document.getElementById('nav-candidate-count').textContent = data.total_candidates;

      // Top roles widget
      const topRolesList = document.getElementById('dashboard-top-roles');
      if (data.top_roles && data.top_roles.length > 0) {
        topRolesList.innerHTML = data.top_roles.map((r) => `
          <div class="top-role-item">
            <span class="top-role-name">${r.role}</span>
            <span class="top-role-count">${r.count} matched</span>
          </div>
        `).join('');
      } else {
        topRolesList.innerHTML = `<div class="empty-state-sm">No candidate matches logged yet. Use the analyzer to score resumes!</div>`;
      }

      // Recent candidates table
      const tbody = document.getElementById('dashboard-recent-tbody');
      if (data.recent_candidates && data.recent_candidates.length > 0) {
        tbody.innerHTML = data.recent_candidates.map((c) => `
          <tr>
            <td><strong>${escapeHtml(c.full_name)}</strong><br><span class="app-id-pill">${c.app_id}</span></td>
            <td>${escapeHtml(c.role_1 || 'General Technical')}</td>
            <td><span class="score-badge ${getScoreBadgeClass(c.score)}">${c.score}%</span></td>
            <td>${escapeHtml(c.location || 'USA')}</td>
            <td><span class="badge ${getStatusBadgeClass(c.status)}">${c.status}</span></td>
            <td>${c.created_at ? new Date(c.created_at).toLocaleDateString() : 'Today'}</td>
            <td><button class="btn btn-secondary btn-sm" onclick="window.viewCandidate('${c.app_id}')">View</button></td>
          </tr>
        `).join('');
      }
    } catch (err) {
      console.error('Failed to load stats:', err);
    }
  }

  // ── Resume Analyzer Flow ────────────────────────────────────────────────────

  // Tabs
  tabUploadFile.addEventListener('click', () => {
    tabUploadFile.classList.add('active');
    tabPasteText.classList.remove('active');
    contentUploadFile.classList.add('active');
    contentPasteText.classList.remove('active');
    updateAnalyzerButtonState();
  });

  tabPasteText.addEventListener('click', () => {
    tabPasteText.classList.add('active');
    tabUploadFile.classList.remove('active');
    contentPasteText.classList.add('active');
    contentUploadFile.classList.remove('active');
    updateAnalyzerButtonState();
  });

  pasteResumeTextarea.addEventListener('input', updateAnalyzerButtonState);

  function updateAnalyzerButtonState() {
    const isFileTab = tabUploadFile.classList.contains('active');
    if (isFileTab) {
      btnRunAnalysis.disabled = !state.selectedFile;
    } else {
      btnRunAnalysis.disabled = pasteResumeTextarea.value.trim().length < 20;
    }
  }

  // Dropzone Setup
  function setupDropzone(dropzoneEl, fileInputEl) {
    if (!dropzoneEl || !fileInputEl) return;

    ['dragenter', 'dragover'].forEach((eventName) => {
      dropzoneEl.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropzoneEl.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach((eventName) => {
      dropzoneEl.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropzoneEl.classList.remove('dragover');
      });
    });

    dropzoneEl.addEventListener('drop', (e) => {
      const files = e.dataTransfer.files;
      if (files.length > 0) {
        handleSelectedFile(files[0]);
      }
    });

    dropzoneEl.addEventListener('click', (e) => {
      if (e.target.closest('button') || e.target.tagName === 'INPUT') return;
      fileInputEl.click();
    });

    fileInputEl.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handleSelectedFile(e.target.files[0]);
      }
    });
  }

  setupDropzone(dropzoneMain, fileInputMain);
  setupDropzone(dropzoneMini, fileInputMini);

  if (btnBrowseFile) {
    btnBrowseFile.addEventListener('click', () => fileInputMain.click());
  }

  function handleSelectedFile(file) {
    state.selectedFile = file;
    fileBannerName.textContent = file.name;
    fileBannerSize.textContent = `${(file.size / (1024 * 1024)).toFixed(2)} MB`;
    selectedFileBanner.classList.remove('hidden');
    dropzoneMain.classList.add('hidden');
    updateAnalyzerButtonState();
  }

  if (btnRemoveFile) {
    btnRemoveFile.addEventListener('click', () => {
      state.selectedFile = null;
      fileInputMain.value = '';
      fileInputMini.value = '';
      selectedFileBanner.classList.add('hidden');
      dropzoneMain.classList.remove('hidden');
      updateAnalyzerButtonState();
    });
  }

  // Run Analysis Execution
  if (btnRunAnalysis) {
    btnRunAnalysis.addEventListener('click', async () => {
      const isFileTab = tabUploadFile.classList.contains('active');
      const formData = new FormData();
      formData.append('save_candidate', chkSaveCandidate.checked ? 'true' : 'false');

      if (isFileTab) {
        if (!state.selectedFile) return;
        formData.append('file', state.selectedFile);
      } else {
        const text = pasteResumeTextarea.value.trim();
        if (text.length < 20) return;
        formData.append('resume_text', text);
      }

      // Show Loading Skeleton
      analysisResults.classList.remove('hidden');
      analysisLoader.classList.remove('hidden');
      analysisContent.classList.add('hidden');
      analysisError.classList.add('hidden');
      btnRunAnalysis.disabled = true;

      try {
        const res = await fetch('/api/score-resume', {
          method: 'POST',
          body: formData,
        });

        if (!res.ok) {
          const errData = await res.json();
          throw new Error(errData.detail || 'Analysis failed');
        }

        const data = await res.json();
        state.currentScoredCandidate = data;
        renderAnalysisResults(data);
      } catch (err) {
        // No alert(): a modal dialog freezes the tab for anything driving the page.
        analysisError.innerHTML = `<span class="badge badge-danger"><i class="fa-solid fa-circle-xmark"></i> Analysis Error</span> ${escapeHtml(err.message)}`;
        analysisError.classList.remove('hidden');
      } finally {
        analysisLoader.classList.add('hidden');
        btnRunAnalysis.disabled = false;
      }
    });
  }

  function renderAnalysisResults(data) {
    analysisContent.classList.remove('hidden');

    const candidate = data.candidate || {};
    const scores = data.scores || [];
    const topRole = scores[0] || {};

    // Avatar Initials
    const name = candidate.full_name || 'Candidate';
    document.getElementById('res-name').textContent = name;
    document.getElementById('res-avatar').textContent = name.split(' ').map((p) => p[0]).join('').substring(0, 2).toUpperCase();
    document.getElementById('res-target-role').textContent = topRole.role || candidate.looking_for_role || 'General Engineer';

    // Location & Geo Badge
    document.getElementById('res-location').textContent = candidate.location || 'USA';
    document.getElementById('res-phone').textContent = candidate.phone || 'Not Provided';
    document.getElementById('res-education').textContent = candidate.education || 'B.S. / Equivalent';
    document.getElementById('res-experience').textContent = candidate.experience || 'Not extracted';

    // Badges
    const badgesEl = document.getElementById('res-badges');
    // Backend already applied the geo rule (handles 'United States', 'US', ...).
    const isUSA = data.status ? !/non-usa|ineligible/i.test(data.status) : (candidate.country || 'USA').toUpperCase() === 'USA';
    badgesEl.innerHTML = `
      <span class="badge ${isUSA ? 'badge-success' : 'badge-danger'}">
        <i class="fa-solid ${isUSA ? 'fa-circle-check' : 'fa-circle-xmark'}"></i>
        ${isUSA ? 'USA Location Verified' : 'Non-USA Location'}
      </span>
      <span class="badge badge-source">Gemini 3.1 Flash-Lite</span>
    `;

    // Skills Cloud
    const skillsCloud = document.getElementById('res-skills-cloud');
    const skillsCount = document.getElementById('res-skills-count');
    const skills = candidate.skills || [];
    skillsCount.textContent = skills.length;
    skillsCloud.innerHTML = skills.map((s) => `<span class="skill-chip">${escapeHtml(s)}</span>`).join('');

    // Top Role Matches List
    const topScoreBadge = document.getElementById('res-top-score-badge');
    topScoreBadge.textContent = `Top Match: ${topRole.score || 0}%`;

    const matchesList = document.getElementById('res-matches-list');
    matchesList.innerHTML = scores.slice(0, 4).map((m, idx) => `
      <div class="match-item">
        <div class="match-item-top">
          <span class="match-role-title">#${idx + 1} ${escapeHtml(m.role)}</span>
          <span class="match-score-pill ${getScoreBadgeClass(m.score)}">${m.score}% Match</span>
        </div>
        <div class="match-progress-bar">
          <div class="match-progress-fill" style="width: ${m.score}%;"></div>
        </div>
        <p class="match-reasoning">${escapeHtml(m.reasoning || 'Strong contextual fit based on semantic skills match.')}</p>
      </div>
    `).join('');

    // Ask Copilot Button Handler
    const btnCopilotAsk = document.getElementById('btn-copilot-ask-cand');
    if (btnCopilotAsk) {
      btnCopilotAsk.onclick = () => {
        switchView('copilot');
        copilotInput.value = `Evaluate ${candidate.full_name} for the ${topRole.role} position. Highlight their top 3 strengths and recommend next interview questions.`;
        btnSendChat.click();
      };
    }

    // Draft Invite Button Handler
    const btnDraftInvite = document.getElementById('btn-draft-invite');
    if (btnDraftInvite) {
      btnDraftInvite.onclick = () => {
        switchView('copilot');
        copilotInput.value = `Draft an enthusiastic and professional interview invitation email for ${candidate.full_name} for the ${topRole.role} role at DriverAI.`;
        btnSendChat.click();
      };
    }
  }

  // ── Candidate Pipeline Loader & Filter ──────────────────────────────────────

  async function loadCandidates() {
    try {
      const q = filterCandSearch ? filterCandSearch.value.trim() : '';
      const status = filterCandStatus ? filterCandStatus.value : 'all';
      const minScore = filterMinScore ? filterMinScore.value : 0;

      let url = `/api/candidates?min_score=${minScore}`;
      if (q) url += `&query=${encodeURIComponent(q)}`;
      if (status !== 'all') url += `&status=${encodeURIComponent(status)}`;

      const res = await fetch(url);
      if (!res.ok) return;
      const data = await res.json();
      state.candidates = data.candidates || [];

      renderCandidatesTable(state.candidates);
    } catch (err) {
      console.error('Failed to load candidates:', err);
    }
  }

  function renderCandidatesTable(candidates) {
    if (!candidatesTableBody) return;

    if (candidates.length === 0) {
      candidatesTableBody.innerHTML = `
        <tr>
          <td colspan="8" class="text-center py-4 text-muted">
            <i class="fa-solid fa-folder-open me-2"></i> No candidates match the current filters.
          </td>
        </tr>
      `;
      return;
    }

    candidatesTableBody.innerHTML = candidates.map((c) => {
      const scoreVal = Number(c.match_score !== undefined ? c.match_score : (c.score !== undefined ? c.score : 75));
      const roleVal = c.suggested_role_1 || c.role_1 || 'General Engineer';
      return `
      <tr>
        <td><span class="app-id-pill">${c.app_id}</span></td>
        <td><strong>${escapeHtml(c.full_name)}</strong><br><small class="text-muted">${escapeHtml(c.email || '')}</small></td>
        <td>${escapeHtml(roleVal)}</td>
        <td><span class="score-badge ${getScoreBadgeClass(scoreVal)}">${scoreVal}%</span></td>
        <td>${escapeHtml(c.location || 'USA')}</td>
        <td><span class="badge ${getStatusBadgeClass(c.status)}">${c.status}</span></td>
        <td>${escapeHtml(c.education || 'Not extracted')}</td>
        <td><button class="btn btn-secondary btn-sm" onclick="window.viewCandidate('${c.app_id}')">Scorecard</button></td>
      </tr>
    `;
    }).join('');
  }

  if (filterCandSearch) filterCandSearch.addEventListener('input', debounce(loadCandidates, 300));
  if (filterCandStatus) filterCandStatus.addEventListener('change', loadCandidates);
  if (filterMinScore) {
    filterMinScore.addEventListener('input', () => {
      scoreSliderVal.textContent = `${filterMinScore.value}%`;
      loadCandidates();
    });
  }
  if (btnRefreshCandidates) btnRefreshCandidates.addEventListener('click', loadCandidates);

  // ── View Candidate Drawer Modal ─────────────────────────────────────────────

  // ── View Candidate Drawer Modal (Tabbed with 1-Click Outreach) ───────────────

  let activeCandidateData = null;

  window.viewCandidate = async function(appId) {
    try {
      const res = await fetch(`/api/candidates/${appId}`);
      if (!res.ok) throw new Error('Candidate not found');
      const cand = await res.json();
      activeCandidateData = cand;

      drawerName.textContent = cand.full_name || 'Candidate Profile';
      drawerStatusBadge.className = `badge ${getStatusBadgeClass(cand.status)}`;
      drawerStatusBadge.textContent = cand.status;

      renderDrawerTab('overview');
      drawerOverlay.classList.remove('hidden');
    } catch (err) {
      alert(`Error loading candidate: ${err.message}`);
    }
  };

  window.viewCandidateDetails = window.viewCandidate;

  function renderDrawerTab(tabName) {
    if (!activeCandidateData) return;
    const cand = activeCandidateData;

    // Update active tab button
    document.querySelectorAll('.drawer-nav-tab').forEach(b => {
      b.classList.toggle('active', b.getAttribute('data-dtab') === tabName);
    });

    if (tabName === 'overview') {
      const skills = Array.isArray(cand.skills) ? cand.skills : (cand.skills ? String(cand.skills).split(',').map(s => s.trim()) : []);
      drawerBody.innerHTML = `
        <div class="profile-details-grid">
          <div class="detail-item">
            <span class="detail-label"><i class="fa-solid fa-id-badge"></i> Application ID</span>
            <span class="detail-value font-mono">${cand.app_id}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label"><i class="fa-solid fa-envelope"></i> Email</span>
            <span class="detail-value">${escapeHtml(cand.email || 'Not extracted')}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label"><i class="fa-solid fa-location-dot"></i> Location</span>
            <span class="detail-value">${escapeHtml(cand.location || 'USA')}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label"><i class="fa-solid fa-phone"></i> Phone</span>
            <span class="detail-value">${escapeHtml(cand.phone || 'Not extracted')}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label"><i class="fa-solid fa-briefcase"></i> Top Role #1</span>
            <span class="detail-value text-cyan">${escapeHtml(cand.role_1 || 'General Match')}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label"><i class="fa-solid fa-award"></i> Match Score</span>
            <span class="detail-value"><span class="score-badge ${getScoreBadgeClass(cand.score)}">${cand.score}%</span></span>
          </div>
        </div>

        <div class="skills-section mt-3">
          <div class="section-heading"><i class="fa-solid fa-microchip text-cyan"></i> Extracted Technical Stack</div>
          <div class="skills-chip-cloud">
            ${skills.map(s => `<span class="skill-chip">${escapeHtml(s)}</span>`).join('')}
          </div>
        </div>

        <div class="matches-list mt-3">
          <div class="section-heading"><i class="fa-solid fa-ranking-star text-amber"></i> Matched DriverAI Roles</div>
          <div class="match-item">
            <div class="match-item-top">
              <strong>#1 ${escapeHtml(cand.role_1 || 'General')}</strong>
              <span class="score-badge ${getScoreBadgeClass(cand.score)}">${cand.score}%</span>
            </div>
            ${cand.reasoning_1 ? `<p class="match-reasoning mt-1">${escapeHtml(cand.reasoning_1)}</p>` : `<p class="match-reasoning mt-1 text-dim">High semantic overlap against core technical responsibilities and distributed systems requirements.</p>`}
          </div>
          ${cand.role_2 ? `
            <div class="match-item mt-2">
              <div class="match-item-top">
                <strong>#2 ${escapeHtml(cand.role_2)}</strong>
                <span class="score-badge ${getScoreBadgeClass(cand.score_2 || 0)}">${cand.score_2 || 0}%</span>
              </div>
            </div>
          ` : ''}
        </div>

        <div class="results-actions mt-3" style="display: flex; gap: 8px;">
          <button class="btn btn-secondary" onclick="window.askCopilotAbout('${escapeHtml(cand.full_name)}', '${escapeHtml(cand.role_1 || '')}')">
            <i class="fa-solid fa-comments text-cyan"></i> Ask Copilot
          </button>
          <button class="btn btn-primary" onclick="renderDrawerTab('outreach')">
            <i class="fa-solid fa-paper-plane"></i> 1-Click Outreach
          </button>
          <!-- ADDED 2026-09-18: open the actual CV. Everything else in this drawer
               is data ABOUT the document; this is the document. -->
          <button class="btn btn-secondary" onclick="window.openResume('${escapeHtml(cand.app_id)}')">
            <i class="fa-solid fa-file-pdf text-amber"></i> Open Resume
          </button>
        </div>
      `;
    } else if (tabName === 'resume') {
      drawerBody.innerHTML = `
        <div class="outreach-card">
          <div class="outreach-header">
            <div class="outreach-title"><i class="fa-solid fa-file-lines text-cyan"></i> Parsed Resume Summary</div>
          </div>
          <div style="background-color: var(--color-canvas-subtle); padding: 12px; border-radius: 6px; font-family: var(--font-mono); font-size: 12px; line-height: 1.6; white-space: pre-wrap; max-height: 420px; overflow-y: auto;">${escapeHtml(cand.notes || cand.summary || 'Full resume text extracted during ingestion.')}</div>
        </div>
      `;
    } else if (tabName === 'outreach') {
      const emailName = cand.full_name || 'Candidate';
      const emailRole = cand.suggested_role_1 || cand.role_1 || 'Software Engineer';
      const emailScore = cand.match_score || cand.score || 0;
      const skillsArr = Array.isArray(cand.skills) ? cand.skills : (cand.skills ? String(cand.skills).split(',').map(s => s.trim()) : []);
      const emailSkills = skillsArr.length > 0 ? skillsArr.slice(0, 3).join(', ') : 'your technical expertise';

      const inviteTemplate = `Hi ${emailName},\n\nThank you for applying to DriverAI! We were thoroughly impressed by your background in ${emailSkills} and believe your skills are a fantastic match for our ${emailRole} role (${emailScore}% fit).\n\nWe would love to invite you to an introductory technical conversation with our team next week. Please let us know your availability for a 45-minute video chat.\n\nBest regards,\nDriverAI Recruiting Team\napply@driverai.io`;
      
      const declineTemplate = `Hi ${emailName},\n\nThank you for your interest in DriverAI and for taking the time to share your background with us.\n\nAt this time, our open positions for ${emailRole} require permanent United States work authorization and residency. Because your profile indicates a non-USA location, we are unfortunately unable to advance your application for this cycle.\n\nWe truly appreciate your interest in DriverAI and wish you the best in your career search.\n\nSincerely,\nDriverAI Talent Team`;

      const probingTemplate = `Technical Probing Questions for ${emailName} (${emailRole}):\n\n1. "Can you walk us through the architecture of a high-throughput pipeline you built using ${emailSkills}?"\n2. "How did you measure and optimize p99 latency in production systems?"\n3. "Describe a time when a consensus or distributed state failure occurred—how did you debug it?"\n4. "What trade-offs did you evaluate when selecting your core framework vs alternatives?"\n5. "How do you approach end-to-end testing and CI/CD validation in a distributed microservices environment?"`;

      drawerBody.innerHTML = `
        <div class="outreach-card">
          <div class="outreach-header">
            <div class="outreach-title"><i class="fa-solid fa-circle-check text-emerald"></i> 1. Technical Interview Invitation</div>
          </div>
          <textarea class="outreach-textarea" id="email-invite-txt">${escapeHtml(inviteTemplate)}</textarea>
          <div class="outreach-actions">
            <button class="btn btn-secondary btn-sm" onclick="window.copyToClipboard(document.getElementById('email-invite-txt').value, 'Invite copied!')">
              <i class="fa-solid fa-copy"></i> Copy
            </button>
            <a class="btn btn-primary btn-sm" href="mailto:${escapeHtml(cand.email || '')}?subject=Interview%20Invitation%20%E2%80%94%20DriverAI%20${encodeURIComponent(emailRole)}&body=${encodeURIComponent(inviteTemplate)}">
              <i class="fa-solid fa-envelope"></i> Open in Mail
            </a>
          </div>
        </div>

        <div class="outreach-card">
          <div class="outreach-header">
            <div class="outreach-title"><i class="fa-solid fa-clipboard-question text-purple"></i> 2. 5 Tailored Technical Probing Questions</div>
          </div>
          <textarea class="outreach-textarea" id="email-probing-txt">${escapeHtml(probingTemplate)}</textarea>
          <div class="outreach-actions">
            <button class="btn btn-secondary btn-sm" onclick="window.copyToClipboard(document.getElementById('email-probing-txt').value, 'Questions copied!')">
              <i class="fa-solid fa-copy"></i> Copy
            </button>
          </div>
        </div>

        <div class="outreach-card">
          <div class="outreach-header">
            <div class="outreach-title"><i class="fa-solid fa-ban text-amber"></i> 3. Polite Location Decline (Non-USA)</div>
          </div>
          <textarea class="outreach-textarea" id="email-decline-txt">${escapeHtml(declineTemplate)}</textarea>
          <div class="outreach-actions">
            <button class="btn btn-secondary btn-sm" onclick="window.copyToClipboard(document.getElementById('email-decline-txt').value, 'Decline copied!')">
              <i class="fa-solid fa-copy"></i> Copy
            </button>
          </div>
        </div>
      `;
    } else if (tabName === 'export') {
      const exportRole = cand.suggested_role_1 || cand.role_1 || 'N/A';
      const exportScore = cand.match_score || cand.score || 0;
      const skillsStr = Array.isArray(cand.skills) ? cand.skills.join(', ') : cand.skills;
      const mdReport = `### 👤 Candidate Evaluation: ${cand.full_name}\n- **Match Score**: ${exportScore}%\n- **Target Role**: \`${exportRole}\`\n- **Location**: ${cand.location || 'USA'} (${cand.status})\n- **Experience**: ${cand.experience || 'N/A'}\n- **Education**: ${cand.education || 'N/A'}\n- **Technical Skills**: ${skillsStr}\n\n**DriverAI AI Fit Summary**:\n${cand.reasoning_1 || 'Strong candidate with verified engineering depth and alignment with 105 company benchmarks.'}`;

      drawerBody.innerHTML = `
        <div class="outreach-card">
          <div class="outreach-header">
            <div class="outreach-title"><i class="fa-brands fa-markdown text-purple"></i> GitHub Issue / Slack Markdown Format</div>
          </div>
          <textarea class="outreach-textarea" id="md-export-txt" style="height: 240px;">${escapeHtml(mdReport)}</textarea>
          <div class="outreach-actions">
            <button class="btn btn-primary btn-sm" onclick="window.copyToClipboard(document.getElementById('md-export-txt').value, 'Markdown copied!')">
              <i class="fa-solid fa-copy"></i> Copy GitHub Markdown
            </button>
          </div>
        </div>
      `;
    }
  }

  // Drawer Tab clicks
  document.querySelectorAll('.drawer-nav-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.getAttribute('data-dtab');
      renderDrawerTab(tab);
    });
  });

  if (btnCloseDrawer) {
    btnCloseDrawer.addEventListener('click', () => drawerOverlay.classList.add('hidden'));
  }
  if (drawerOverlay) {
    drawerOverlay.addEventListener('click', (e) => {
      if (e.target === drawerOverlay) drawerOverlay.classList.add('hidden');
    });
  }

  window.copyToClipboard = function(text, successMsg) {
    navigator.clipboard.writeText(text).then(() => {
      alert(successMsg || 'Copied to clipboard!');
    }).catch(() => {
      alert('Copied to clipboard!');
    });
  };

  // ADDED 2026-09-18: open the candidate's CV. The server answers with either a
  // SharePoint link (Developer mode) or the file itself (local), so ask first and
  // then open, rather than guessing which it will be.
  window.openResume = async function(appId) {
    try {
      const r = await fetch(`/api/candidates/${encodeURIComponent(appId)}/resume`);
      if (!r.ok) {
        const d = await r.json().catch(() => ({}));
        showToast(d.detail || 'Resume file not available', 'error');
        return;
      }
      const type = r.headers.get('content-type') || '';
      if (type.includes('application/json')) {
        const d = await r.json();
        window.open(d.url, '_blank', 'noopener');      // SharePoint-hosted CV
      } else {
        const blob = await r.blob();                    // local file
        const url = URL.createObjectURL(blob);
        window.open(url, '_blank', 'noopener');
        setTimeout(() => URL.revokeObjectURL(url), 60000);
      }
    } catch (e) {
      showToast('Could not open the resume', 'error');
    }
  };

  window.askCopilotAbout = function(name, role) {
    drawerOverlay.classList.add('hidden');
    switchView('copilot');
    copilotInput.value = `Tell me about candidate ${name} who applied for ${role}. What are their strengths and potential gaps?`;
    btnSendChat.click();
  };

  // ── Open Roles Explorer Loader ──────────────────────────────────────────────

  async function loadRoles() {
    try {
      const res = await fetch('/api/roles');
      if (!res.ok) return;
      const data = await res.json();
      state.roles = data.roles || [];
      state.categories = data.categories || [];

      renderRolesCards(state.roles);
    } catch (err) {
      console.error('Failed to load roles:', err);
    }
  }

  function renderRolesCards(roles) {
    if (!rolesCardsGrid) return;

    const query = filterRolesSearch ? filterRolesSearch.value.trim().toLowerCase() : '';
    const activeCategoryBtn = document.querySelector('.category-pill.active');
    const selectedCategory = activeCategoryBtn ? activeCategoryBtn.dataset.category : 'all';

    let filtered = roles;
    if (selectedCategory !== 'all') {
      filtered = filtered.filter((r) => r.category === selectedCategory);
    }
    if (query) {
      filtered = filtered.filter((r) =>
        r.title.toLowerCase().includes(query) ||
        (r.category || '').toLowerCase().includes(query) ||
        (r.skills || []).some((s) => (s || '').toLowerCase().includes(query))
      );
    }

    if (filtered.length === 0) {
      rolesCardsGrid.innerHTML = `
        <div class="glass-card text-center py-4" style="grid-column: 1 / -1;">
          <p class="text-muted">No open roles found matching "${query}".</p>
        </div>
      `;
      return;
    }

    rolesCardsGrid.innerHTML = filtered.map((r) => `
      <div class="glass-card role-card">
        <div class="role-card-header">
          <div>
            <h3 class="role-card-title">${escapeHtml(r.title)}</h3>
            <span class="role-card-dept">${escapeHtml(r.category || 'Uncategorised')}</span>
          </div>
          <span class="badge badge-source">${escapeHtml(r.experience_level || 'Mid-Senior')}</span>
        </div>

        <div class="role-card-skills">
          ${(r.skills || []).slice(0, 5).map((s) => `<span class="skill-chip">${escapeHtml(s)}</span>`).join('')}
        </div>

        <div class="role-card-actions">
          <span class="text-muted" style="font-size: 12px;"><i class="fa-solid fa-location-dot"></i> USA Only</span>
          <button class="btn btn-secondary btn-sm" onclick="filterCandidatesForRole('${escapeHtml(r.title)}')">View Candidates</button>
        </div>
      </div>
    `).join('');
  }

  if (filterRolesSearch) filterRolesSearch.addEventListener('input', () => renderRolesCards(state.roles));

  document.querySelectorAll('.category-pill').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.category-pill').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      renderRolesCards(state.roles);
    });
  });

  window.filterCandidatesForRole = function(roleTitle) {
    switchView('candidates');
    filterCandSearch.value = roleTitle;
    loadCandidates();
  };

  // ── Recruiter Copilot Chat Flow ─────────────────────────────────────────────

  promptChips.forEach((chip) => {
    chip.addEventListener('click', () => {
      copilotInput.value = chip.dataset.prompt;
      btnSendChat.click();
    });
  });

  if (btnSendChat) {
    btnSendChat.addEventListener('click', handleSendChatMessage);
  }

  if (copilotInput) {
    copilotInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSendChatMessage();
      }
    });
  }

  if (btnClearChat) {
    btnClearChat.addEventListener('click', () => {
      state.chatHistory = [];
      copilotMessages.innerHTML = `
        <div class="chat-bubble assistant-bubble">
          <div class="bubble-avatar"><img src="/static/driverai_logo.png" alt="DriverAI" class="copilot-bubble-logo" /></div>
          <div class="bubble-content">
            <p>Chat cleared. How can I assist your talent sourcing today?</p>
          </div>
        </div>
      `;
    });
  }

  async function handleSendChatMessage() {
    const text = copilotInput.value.trim();
    if (!text) return;

    // Append User Message
    appendChatBubble('user', text);
    copilotInput.value = '';

    // Append Thinking Indicator
    const thinkingId = `thinking-${Date.now()}`;
    const thinkingEl = document.createElement('div');
    thinkingEl.id = thinkingId;
    thinkingEl.className = 'chat-bubble assistant-bubble';
    thinkingEl.innerHTML = `
      <div class="bubble-avatar"><img src="/static/driverai_logo.png" alt="DriverAI" class="copilot-bubble-logo" /></div>
      <div class="bubble-content"><i class="fa-solid fa-circle-notch fa-spin me-2"></i> Thinking with Gemini 3.1...</div>
    `;
    copilotMessages.appendChild(thinkingEl);
    copilotMessages.scrollTop = copilotMessages.scrollHeight;

    try {
      const res = await fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: text,
          history: state.chatHistory,
        }),
      });

      if (!res.ok) throw new Error('Copilot response failed');
      const data = await res.json();

      // Remove thinking
      const el = document.getElementById(thinkingId);
      if (el) el.remove();

      // Append Assistant Message
      const replyText = data.reply || data.response || 'No response';
      appendChatBubble('assistant', replyText);

      // Save to history
      state.chatHistory.push({ role: 'user', content: text });
      state.chatHistory.push({ role: 'assistant', content: replyText });
    } catch (err) {
      const el = document.getElementById(thinkingId);
      if (el) el.remove();
      appendChatBubble('assistant', `⚠️ Sorry, I encountered an error: ${err.message}`);
    }
  }

  function appendChatBubble(role, content) {
    const bubble = document.createElement('div');
    bubble.className = `chat-bubble ${role}-bubble`;
    const avatarHtml = role === 'user' 
      ? '<i class="fa-solid fa-user"></i>' 
      : '<img src="/static/driverai_logo.png" alt="DriverAI" class="copilot-bubble-logo" />';
    bubble.innerHTML = `
      <div class="bubble-avatar">${avatarHtml}</div>
      <div class="bubble-content"><p>${formatMarkdownText(content)}</p></div>
    `;
    copilotMessages.appendChild(bubble);
    copilotMessages.scrollTop = copilotMessages.scrollHeight;
  }

  // ── Side-by-Side Candidate Comparison Matrix & Skill Radar ────────────────
  const compareRoleSelect = document.getElementById('compare-role-select');
  const compareCandidateChips = document.getElementById('compare-candidate-chips');
  const btnDemoCompare = document.getElementById('btn-demo-compare');
  const btnRunCompare = document.getElementById('btn-run-compare');
  const radarSvgContainer = document.getElementById('radar-svg-container');
  const radarLegend = document.getElementById('radar-legend');
  const compareVerdictContent = document.getElementById('compare-verdict-content');
  const matrixGridContainer = document.getElementById('matrix-grid-container');

  async function initCompareView() {
    // 1. Ensure candidates are loaded
    if (!state.candidates || state.candidates.length === 0) {
      try {
        const res = await fetch('/api/candidates');
        const data = await res.json();
        state.candidates = data.candidates || [];
      } catch (e) {
        console.error('Failed to load candidates for compare', e);
      }
    }

    // 2. Populate Target Role Select if empty
    if (compareRoleSelect && state.roles && state.roles.length > 0) {
      const currentVal = compareRoleSelect.value;
      compareRoleSelect.innerHTML = state.roles.slice(0, 30).map((r) =>
        `<option value="${escapeHtml(r.title)}">${escapeHtml(r.title)}</option>`
      ).join('');
      if (currentVal) compareRoleSelect.value = currentVal;
    }

    // 3. Render Candidate Select Chips
    renderCompareCandidateChips();

    // 4. Auto-trigger comparison if not cached
    if (!state.compareState.cachedComparison) {
      if (state.candidates && state.candidates.length >= 2) {
        state.compareState.selectedAppIds = [state.candidates[0].app_id, state.candidates[1].app_id];
        renderCompareCandidateChips();
        runComparison(state.compareState.selectedAppIds, compareRoleSelect ? compareRoleSelect.value : null);
      } else if (state.candidates && state.candidates.length === 1) {
        state.compareState.selectedAppIds = [state.candidates[0].app_id];
        renderCompareCandidateChips();
        runComparison(state.compareState.selectedAppIds, compareRoleSelect ? compareRoleSelect.value : null);
      }
    }
  }

  function renderCompareCandidateChips() {
    if (!compareCandidateChips) return;
    if (!state.candidates || state.candidates.length === 0) {
      compareCandidateChips.innerHTML = '<span class="text-dim text-sm">No candidates in pipeline yet. Upload a resume first!</span>';
      return;
    }

    compareCandidateChips.innerHTML = state.candidates.map((cand) => {
      const isSelected = state.compareState.selectedAppIds.includes(cand.app_id);
      const initials = (cand.full_name || 'C').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
      return `
        <div class="compare-chip ${isSelected ? 'active' : ''}" data-app-id="${cand.app_id}">
          <span class="compare-chip-avatar">${initials}</span>
          <span>${escapeHtml(cand.full_name || 'Candidate')}</span>
          <i class="fa-solid fa-check compare-chip-check"></i>
        </div>
      `;
    }).join('');

    // Attach click listeners to chips
    compareCandidateChips.querySelectorAll('.compare-chip').forEach((chip) => {
      chip.addEventListener('click', () => {
        const appId = chip.dataset.appId;
        const idx = state.compareState.selectedAppIds.indexOf(appId);
        if (idx > -1) {
          state.compareState.selectedAppIds.splice(idx, 1);
        } else {
          if (state.compareState.selectedAppIds.length >= 4) {
            alert('You can compare up to 4 candidates simultaneously.');
            return;
          }
          state.compareState.selectedAppIds.push(appId);
        }
        renderCompareCandidateChips();
      });
    });
  }

  async function runComparison(appIds, targetRole) {
    if (!appIds || appIds.length === 0) {
      alert('Please select at least 1 candidate to compare.');
      return;
    }

    if (!targetRole && compareRoleSelect) {
      targetRole = compareRoleSelect.value;
    }

    // Show loading state in verdict & radar
    if (compareVerdictContent) {
      compareVerdictContent.innerHTML = `
        <div class="verdict-placeholder text-center py-4">
          <i class="fa-solid fa-circle-notch fa-spin fa-2x text-cyan mb-2"></i>
          <p class="text-secondary">Generating AI Head-to-Head Comparative Verdict with Gemini 3.1...</p>
        </div>
      `;
    }

    try {
      const res = await fetch('/api/compare', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          app_ids: appIds,
          target_role: targetRole,
        }),
      });

      if (!res.ok) throw new Error('Comparison API request failed');
      const data = await res.json();
      state.compareState.cachedComparison = data;

      // Render Visual Components
      renderRadarChart(data.axes, data.series);
      renderVerdict(data.verdict, targetRole);
      renderMatrixGrid(data.candidates, data.axes, data.series);

    } catch (err) {
      console.error('Comparison error', err);
      if (compareVerdictContent) {
        compareVerdictContent.innerHTML = `
          <div class="verdict-placeholder text-center py-4">
            <i class="fa-solid fa-triangle-exclamation fa-2x text-amber mb-2"></i>
            <p class="text-secondary">Error loading comparison: ${escapeHtml(err.message)}</p>
          </div>
        `;
      }
    }
  }

  function renderRadarChart(axes, series) {
    if (!radarSvgContainer) return;
    if (!axes || !series || series.length === 0) {
      radarSvgContainer.innerHTML = '<p class="text-muted text-sm">No series data available</p>';
      return;
    }

    const size = 420;
    const cx = size / 2;
    const cy = size / 2;
    const maxRadius = 135;
    const numAxes = axes.length;
    const angleSlice = (Math.PI * 2) / numAxes;

    // Build Concentric Grid Rings (20%, 40%, 60%, 80%, 100%)
    const rings = [0.2, 0.4, 0.6, 0.8, 1.0];
    let gridPolygonsSvg = rings.map((factor) => {
      const r = maxRadius * factor;
      const pts = [];
      for (let i = 0; i < numAxes; i++) {
        const angle = angleSlice * i - Math.PI / 2;
        const x = cx + r * Math.cos(angle);
        const y = cy + r * Math.sin(angle);
        pts.push(`${x.toFixed(1)},${y.toFixed(1)}`);
      }
      return `<polygon class="radar-grid-polygon" points="${pts.join(' ')}" />`;
    }).join('\n');

    // Build Axis Spokes & Labels
    let spokesSvg = '';
    let labelsSvg = '';
    for (let i = 0; i < numAxes; i++) {
      const angle = angleSlice * i - Math.PI / 2;
      const xMax = cx + maxRadius * Math.cos(angle);
      const yMax = cy + maxRadius * Math.sin(angle);
      spokesSvg += `<line class="radar-grid-spoke" x1="${cx}" y1="${cy}" x2="${xMax.toFixed(1)}" y2="${yMax.toFixed(1)}" />`;

      // Label coordinate slightly beyond maxRadius
      const labelRadius = maxRadius + 32;
      const xLabel = cx + labelRadius * Math.cos(angle);
      const yLabel = cy + labelRadius * Math.sin(angle);
      
      labelsSvg += `
        <text class="radar-axis-label" x="${xLabel.toFixed(1)}" y="${yLabel.toFixed(1)}">${escapeHtml(axes[i])}</text>
      `;
    }

    // Build Candidate Data Polygons & Dots
    let seriesSvg = series.map((s, sIdx) => {
      const pts = [];
      const dots = [];
      const col = s.color || '#00f2fe';

      for (let i = 0; i < numAxes; i++) {
        const val = s.values[i] || 0;
        const r = (val / 100) * maxRadius;
        const angle = angleSlice * i - Math.PI / 2;
        const x = cx + r * Math.cos(angle);
        const y = cy + r * Math.sin(angle);
        pts.push(`${x.toFixed(1)},${y.toFixed(1)}`);
        dots.push(`
          <circle class="radar-data-point" cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" 
                  style="fill: ${col}; stroke: #fff;" 
                  data-name="${escapeHtml(s.name)}" 
                  data-axis="${escapeHtml(axes[i])}" 
                  data-val="${val}%">
            <title>${escapeHtml(s.name)} - ${escapeHtml(axes[i])}: ${val}%</title>
          </circle>
        `);
      }

      return `
        <g class="radar-data-series" data-series="${sIdx}">
          <polygon class="radar-data-polygon" points="${pts.join(' ')}" 
                   style="fill: ${col}; stroke: ${col};" />
          ${dots.join('\n')}
        </g>
      `;
    }).join('\n');

    radarSvgContainer.innerHTML = `
      <svg class="radar-svg" viewBox="0 0 ${size} ${size}">
        <g class="radar-grid">${gridPolygonsSvg}</g>
        <g class="radar-spokes">${spokesSvg}</g>
        <g class="radar-series-group">${seriesSvg}</g>
        <g class="radar-labels">${labelsSvg}</g>
      </svg>
    `;

    // Render Legend
    if (radarLegend) {
      radarLegend.innerHTML = series.map((s) => `
        <div class="radar-legend-item">
          <span class="radar-legend-color" style="background-color: ${s.color}; color: ${s.color};"></span>
          <span>${escapeHtml(s.name)}</span>
        </div>
      `).join('');
    }
  }

  function renderVerdict(verdictMarkdown, targetRole) {
    if (!compareVerdictContent) return;
    if (!verdictMarkdown) {
      compareVerdictContent.innerHTML = '<p class="text-dim">No verdict returned.</p>';
      return;
    }

    let html = formatMarkdownText(verdictMarkdown);
    // Wrap any recommendation in a glowing box
    if (html.includes('Final Hiring Recommendation') || html.includes('Recommendation:')) {
      html = html.replace(/(<strong>(?:Final Hiring Recommendation|Recommendation:?)<\/strong>[\s\S]*)/i,
        '<div class="verdict-highlight-box"><i class="fa-solid fa-trophy text-amber me-2"></i>$1</div>');
    }

    compareVerdictContent.innerHTML = html;
  }

  function renderMatrixGrid(candidates, axes, series) {
    if (!matrixGridContainer) return;
    if (!candidates || candidates.length === 0) {
      matrixGridContainer.innerHTML = '<p class="text-dim">No candidates to display.</p>';
      return;
    }

    const palette = ['#00f2fe', '#a855f7', '#10b981', '#f59e0b'];

    matrixGridContainer.innerHTML = candidates.map((cand, idx) => {
      const colColor = (series[idx] && series[idx].color) || palette[idx % palette.length];
      const initials = (cand.full_name || 'C').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
      const candSeries = series[idx] ? series[idx].values : [75, 75, 75, 75, 75, 75];

      const dimensionBarsHtml = axes.map((axisName, aIdx) => {
        const val = candSeries[aIdx] || 50;
        return `
          <div class="matrix-dim-row">
            <div class="matrix-dim-header">
              <span>${escapeHtml(axisName)}</span>
              <span style="color: ${colColor}; font-weight: 700;">${val}%</span>
            </div>
            <div class="matrix-dim-track">
              <div class="matrix-dim-fill" style="width: ${val}%; background-color: ${colColor}; color: ${colColor};"></div>
            </div>
          </div>
        `;
      }).join('');

      const topSkills = (cand.skills || []).slice(0, 8);

      return `
        <div class="matrix-col-card" style="border-top: 2px solid ${colColor};">
          <div class="matrix-col-header">
            <div class="matrix-candidate-meta">
              <div class="matrix-avatar" style="background: ${colColor}; color: #030712; box-shadow: 0 0 12px ${colColor};">
                ${initials}
              </div>
              <div class="matrix-name-group">
                <span class="matrix-candidate-name">${escapeHtml(cand.full_name)}</span>
                <span class="matrix-candidate-sub">${escapeHtml(cand.location || 'USA')}</span>
              </div>
            </div>
            <div class="matrix-overall-score">
              <span class="matrix-score-value" style="color: ${colColor};">${cand.score}%</span>
              <span class="matrix-score-sub">Match</span>
            </div>
          </div>

          <div class="matrix-dimensions-group">
            <div class="matrix-section-title"><i class="fa-solid fa-chart-simple text-cyan"></i> 6-Axis Dimension Breakdown</div>
            ${dimensionBarsHtml}
          </div>

          <div class="matrix-details-list">
            <div class="matrix-detail-item">
              <span class="matrix-detail-label"><i class="fa-solid fa-briefcase"></i> Suggested Role:</span>
              <span class="matrix-detail-value text-cyan" style="max-width: 160px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${escapeHtml(cand.role_1 || 'General')}</span>
            </div>
            <div class="matrix-detail-item">
              <span class="matrix-detail-label"><i class="fa-solid fa-graduation-cap"></i> Education:</span>
              <span class="matrix-detail-value">${escapeHtml(cand.education || 'Not extracted')}</span>
            </div>
            <div class="matrix-detail-item">
              <span class="matrix-detail-label"><i class="fa-solid fa-clock-rotate-left"></i> Experience:</span>
              <span class="matrix-detail-value">${escapeHtml(cand.experience || 'Not extracted')}</span>
            </div>
            <div class="matrix-detail-item">
              <span class="matrix-detail-label"><i class="fa-solid fa-flag-usa"></i> Status:</span>
              <span class="matrix-detail-value"><span class="badge ${getStatusBadgeClass(cand.status)}">${escapeHtml(cand.status)}</span></span>
            </div>
          </div>

          <div class="matrix-skills-group">
            <div class="matrix-section-title"><i class="fa-solid fa-code text-purple"></i> Key Technical Stack</div>
            <div class="matrix-skills-container">
              ${topSkills.map(s => `<span class="matrix-skill-pill">${escapeHtml(s)}</span>`).join('')}
            </div>
          </div>

          <div class="matrix-card-actions">
            <button class="btn btn-secondary btn-sm" onclick="window.viewCandidateDetails('${cand.app_id}')">
              <i class="fa-solid fa-id-card"></i> View Full Dossier
            </button>
            <button class="btn btn-secondary btn-sm" onclick="askCopilotAbout('${escapeHtml(cand.full_name)}', '${escapeHtml(cand.role_1 || '')}')">
              <i class="fa-solid fa-comment-dots"></i> Probing Questions
            </button>
          </div>
        </div>
      `;
    }).join('');
  }

  // Connect Comparison Action Buttons
  if (btnDemoCompare) {
    btnDemoCompare.addEventListener('click', () => {
      const alex = state.candidates.find(c => c.full_name && c.full_name.toLowerCase().includes('alex'));
      const sarah = state.candidates.find(c => c.full_name && c.full_name.toLowerCase().includes('sarah'));

      let appIds = [];
      if (alex && sarah) {
        appIds = [alex.app_id, sarah.app_id];
      } else if (state.candidates.length >= 2) {
        appIds = [state.candidates[0].app_id, state.candidates[1].app_id];
      } else if (state.candidates.length === 1) {
        appIds = [state.candidates[0].app_id];
      }

      state.compareState.selectedAppIds = appIds;
      renderCompareCandidateChips();
      runComparison(appIds, compareRoleSelect ? compareRoleSelect.value : null);
    });
  }

  if (btnRunCompare) {
    btnRunCompare.addEventListener('click', () => {
      if (state.compareState.selectedAppIds.length === 0) {
        alert('Please select at least 1 candidate to compare.');
        return;
      }
      runComparison(state.compareState.selectedAppIds, compareRoleSelect ? compareRoleSelect.value : null);
    });
  }

  if (compareRoleSelect) {
    compareRoleSelect.addEventListener('change', () => {
      if (state.compareState.selectedAppIds.length > 0) {
        runComparison(state.compareState.selectedAppIds, compareRoleSelect.value);
      }
    });
  }

  // ── Live Resume Parsing Stepper Controller ──────────────────────────────────

  window.runStepperProgress = function(duration = 2400) {
    const stepper = document.getElementById('dash-stepper');
    if (!stepper) return;
    stepper.classList.add('active');

    const steps = [
      document.getElementById('step-ocr'),
      document.getElementById('step-geo'),
      document.getElementById('step-match'),
      document.getElementById('step-radar')
    ];

    steps.forEach((s, idx) => {
      if (s) {
        s.className = 'stepper-step';
        const iconEl = s.querySelector('.step-icon');
        const defaultIcons = ['fa-file-lines', 'fa-earth-americas', 'fa-wand-magic-sparkles', 'fa-chart-pie'];
        if (iconEl) iconEl.innerHTML = `<i class="fa-solid ${defaultIcons[idx]}"></i>`;
      }
    });

    const stepInterval = duration / 4;
    let current = 0;

    function advance() {
      if (current > 0 && steps[current - 1]) {
        steps[current - 1].className = 'stepper-step step-done';
        const iconEl = steps[current - 1].querySelector('.step-icon');
        if (iconEl) iconEl.innerHTML = '<i class="fa-solid fa-check"></i>';
      }
      if (current < steps.length && steps[current]) {
        steps[current].className = 'stepper-step step-active';
        current++;
        setTimeout(advance, stepInterval);
      }
    }
    advance();
  };

  // ── Compare Matrix Markdown Export ──────────────────────────────────────────

  const btnExportCompareMd = document.getElementById('btn-export-compare-md');
  if (btnExportCompareMd) {
    btnExportCompareMd.addEventListener('click', async () => {
      const appIds = state.compareState.selectedAppIds;
      if (!appIds || appIds.length === 0) {
        alert('Please run a candidate comparison first before exporting.');
        return;
      }

      try {
        const res = await fetch('/api/candidates/export-markdown', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            app_ids: appIds,
            target_role: compareRoleSelect ? compareRoleSelect.value : null
          })
        });

        if (!res.ok) throw new Error('Failed to generate markdown report');
        const data = await res.json();
        const mdText = data.markdown;

        window.copyToClipboard(mdText, '📋 GitHub Markdown Comparison Scorecard copied to clipboard!');
      } catch (err) {
        alert(`Export failed: ${err.message}`);
      }
    });
  }

  // ── Local Directory Batch Ingestion ─────────────────────────────────────────

  const btnTriggerLocalIngest = document.getElementById('btn-trigger-local-ingest');
  const localIngestPath = document.getElementById('local-ingest-path');
  const localIngestResult = document.getElementById('local-ingest-result');

  if (btnTriggerLocalIngest) {
    btnTriggerLocalIngest.addEventListener('click', async () => {
      const dirPath = localIngestPath ? localIngestPath.value.trim() : './resumes';
      if (!dirPath) {
        alert('Please specify a directory path.');
        return;
      }

      btnTriggerLocalIngest.disabled = true;
      btnTriggerLocalIngest.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i> Ingesting...';
      if (localIngestResult) localIngestResult.textContent = `Scanning '${dirPath}' for PDF and DOCX files...`;

      try {
        const res = await fetch('/api/ingest/local-dir', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ directory_path: dirPath, save_to_db: true })
        });

        const data = await res.json();
        if (!res.ok) throw new Error(data.detail || 'Ingestion request failed');

        if (localIngestResult) {
          localIngestResult.innerHTML = `<span class="text-emerald"><i class="fa-solid fa-circle-check"></i> ${data.message}</span>`;
        }

        // Refresh stats and candidate list
        loadStats();
        loadCandidates();
      } catch (err) {
        if (localIngestResult) {
          localIngestResult.innerHTML = `<span class="text-rose"><i class="fa-solid fa-triangle-exclamation"></i> ${err.message}</span>`;
        }
      } finally {
        btnTriggerLocalIngest.disabled = false;
        btnTriggerLocalIngest.innerHTML = '<i class="fa-solid fa-play"></i> Scan & Ingest Folder';
      }
    });
  }

  // ── Command Palette (Cmd + K) Controller ────────────────────────────────────

  const paletteOverlay = document.getElementById('command-palette-overlay');
  const paletteInput = document.getElementById('palette-search-input');
  const paletteResults = document.getElementById('palette-results-list');
  const btnOpenPalette = document.getElementById('btn-open-palette');
  const globalSearchInput = document.getElementById('global-search');

  function openPalette() {
    if (!paletteOverlay) return;
    paletteOverlay.classList.remove('hidden');
    if (paletteInput) {
      paletteInput.value = '';
      paletteInput.focus();
      renderPaletteResults('');
    }
  }

  function closePalette() {
    if (!paletteOverlay) return;
    paletteOverlay.classList.add('hidden');
  }

  if (btnOpenPalette) btnOpenPalette.addEventListener('click', openPalette);
  if (globalSearchInput) {
    globalSearchInput.addEventListener('focus', () => {
      globalSearchInput.blur();
      openPalette();
    });
  }

  if (paletteOverlay) {
    paletteOverlay.addEventListener('click', (e) => {
      if (e.target === paletteOverlay) closePalette();
    });
  }

  function renderPaletteResults(query) {
    if (!paletteResults) return;
    const q = query.toLowerCase().trim();

    const quickActions = [
      { type: 'action', title: 'Go to Executive Dashboard', meta: 'Navigation', icon: 'fa-chart-pie', action: () => switchView('dashboard') },
      { type: 'action', title: 'Open Candidate Pipeline', meta: 'Navigation', icon: 'fa-users-viewfinder', action: () => switchView('candidates') },
      { type: 'action', title: 'Open Compare Matrix & Skill Radar', meta: 'Navigation', icon: 'fa-code-compare', action: () => switchView('compare') },
      { type: 'action', title: `Explore ${JD_COUNT.n} Open Roles`, meta: 'Navigation', icon: 'fa-briefcase', action: () => switchView('roles') },
      { type: 'action', title: 'Ask Recruiter AI Copilot', meta: 'Gemini 3.1', icon: 'fa-comments', action: () => switchView('copilot') },
      { type: 'action', title: '1-Click Demo (Alex vs Sarah)', meta: 'Compare', icon: 'fa-bolt', action: () => { switchView('compare'); btnDemoCompare.click(); } },
    ];

    let items = [];

    if (!q) {
      items = quickActions;
    } else {
      // Filter quick actions
      const matchedActions = quickActions.filter(a => a.title.toLowerCase().includes(q) || a.meta.toLowerCase().includes(q));
      
      // Filter candidates
      const matchedCandidates = (state.candidates || []).filter(c => {
        const name = (c.full_name || '').toLowerCase();
        const skills = (c.skills || '').toLowerCase();
        const role = (c.role_1 || '').toLowerCase();
        return name.includes(q) || skills.includes(q) || role.includes(q);
      }).map(c => ({
        type: 'candidate',
        title: `${c.full_name} (${c.score || c.match_score || 0}% match)`,
        meta: c.role_1 || 'Candidate',
        icon: 'fa-user',
        action: () => { window.viewCandidate(c.app_id); }
      }));

      // Filter roles
      const matchedRoles = (state.roles || []).filter(r => {
        const title = (r.title || '').toLowerCase();
        const dept = (r.category || '').toLowerCase();
        return title.includes(q) || dept.includes(q);
      }).slice(0, 5).map(r => ({
        type: 'role',
        title: r.title,
        meta: r.category || 'Roles Catalog',
        icon: 'fa-briefcase',
        action: () => { switchView('roles'); }
      }));

      items = [...matchedActions, ...matchedCandidates, ...matchedRoles];
    }

    if (items.length === 0) {
      paletteResults.innerHTML = `<div style="padding: 16px; text-align: center; color: var(--color-fg-muted); font-size: 13px;">No results found for "${escapeHtml(q)}"</div>`;
      return;
    }

    paletteResults.innerHTML = items.slice(0, 10).map((it, idx) => `
      <div class="palette-item ${idx === 0 ? 'active' : ''}" data-idx="${idx}">
        <div style="display: flex; align-items: center; gap: 10px;">
          <i class="fa-solid ${it.icon} text-cyan"></i>
          <span>${escapeHtml(it.title)}</span>
        </div>
        <span class="palette-item-meta">${escapeHtml(it.meta)}</span>
      </div>
    `).join('');

    // Attach click handlers
    paletteResults.querySelectorAll('.palette-item').forEach((el, idx) => {
      el.addEventListener('click', () => {
        closePalette();
        if (items[idx] && items[idx].action) items[idx].action();
      });
    });
  }

  if (paletteInput) {
    paletteInput.addEventListener('input', (e) => {
      renderPaletteResults(e.target.value);
    });

    paletteInput.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        closePalette();
      } else if (e.key === 'Enter') {
        const activeEl = paletteResults.querySelector('.palette-item.active') || paletteResults.querySelector('.palette-item');
        if (activeEl) {
          activeEl.click();
        }
      }
    });
  }

  // ── Global Keyboard Shortcuts ───────────────────────────────────────────────

  let lastKeyTime = 0;
  let lastKey = '';

  document.addEventListener('keydown', (e) => {
    // Cmd + K or Ctrl + K
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
      e.preventDefault();
      openPalette();
      return;
    }

    // Do not trigger shortcuts if typing inside inputs/textareas
    const tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || e.target.isContentEditable) return;

    if (e.key === 'Escape') {
      closePalette();
      if (drawerOverlay) drawerOverlay.classList.add('hidden');
      return;
    }

    const now = Date.now();
    const key = e.key.toLowerCase();

    if (key === '?') {
      openPalette();
      return;
    }

    // Sequence shortcuts starting with 'g'
    if (lastKey === 'g' && (now - lastKeyTime) < 1000) {
      if (key === 'd') { switchView('dashboard'); e.preventDefault(); }
      else if (key === 'c') { switchView('candidates'); e.preventDefault(); }
      else if (key === 'm') { switchView('compare'); e.preventDefault(); }
      else if (key === 'r') { switchView('roles'); e.preventDefault(); }
      else if (key === 'a') { switchView('copilot'); e.preventDefault(); }
      lastKey = '';
    } else {
      lastKey = key;
      lastKeyTime = now;
    }
  });

  // ── Utilities ───────────────────────────────────────────────────────────────

  function getScoreBadgeClass(score) {
    const s = Number(score) || 0;
    if (s >= 85) return 'score-emerald';
    if (s >= 70) return 'score-cyan';
    if (s >= 50) return 'score-purple';
    return 'score-amber';
  }

  function getStatusBadgeClass(status) {
    if (status === 'Scored') return 'badge-success';
    if (status && status.includes('Non-USA')) return 'badge-danger';
    if (status === 'Needs Review') return 'badge-warning';
    return 'badge-source';
  }

  // ADDED 2026-09-18: non-blocking notice. alert() freezes the tab, which is
  // rough in a tool someone is demoing. Deliberately dependency-free.
  function showToast(msg, kind) {
    let host = document.getElementById('toast-host');
    if (!host) {
      host = document.createElement('div');
      host.id = 'toast-host';
      host.style.cssText =
        'position:fixed;bottom:20px;right:20px;z-index:9999;display:flex;' +
        'flex-direction:column;gap:8px;max-width:380px;';
      document.body.appendChild(host);
    }
    const el = document.createElement('div');
    el.textContent = msg;
    el.style.cssText =
      'padding:10px 14px;border-radius:6px;font-size:13px;line-height:1.4;' +
      'color:#e6edf3;background:' + (kind === 'error' ? '#5a1e1e' : '#1f2d3d') +
      ';border:1px solid ' + (kind === 'error' ? '#8b3232' : '#30363d') +
      ';box-shadow:0 4px 14px rgba(0,0,0,.4);';
    host.appendChild(el);
    setTimeout(() => el.remove(), 4500);
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function formatMarkdownText(txt) {
    if (!txt) return '';
    let out = escapeHtml(txt);
    // Headings
    out = out.replace(/^#### (.*?)$/gm, '<h4 style="color: var(--accent-cyan-bright); font-size: 15px; font-weight: 700; margin-top: 14px; margin-bottom: 6px;">$1</h4>');
    out = out.replace(/^### (.*?)$/gm, '<h3 style="color: #fff; font-size: 17px; font-weight: 700; margin-top: 16px; margin-bottom: 8px;">$1</h3>');
    out = out.replace(/^## (.*?)$/gm, '<h2 style="color: #fff; font-size: 19px; font-weight: 700; margin-top: 18px; margin-bottom: 10px;">$1</h2>');
    // Horizontal rules
    out = out.replace(/^---$/gm, '<hr style="border: 0; border-top: 1px solid var(--glass-border-subtle); margin: 14px 0;">');
    // Bold & italic & code
    out = out.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    out = out.replace(/\*(.*?)\*/g, '<em>$1</em>');
    out = out.replace(/`([^`]+)`/g, '<code>$1</code>');
    // Bullet points
    out = out.replace(/^\s*\*\s+(.*?)$/gm, '<li style="margin-bottom: 4px;">$1</li>');
    out = out.replace(/(<li[\s\S]*?<\/li>)/g, '<ul style="padding-left: 20px; margin-bottom: 10px;">$1</ul>');
    out = out.replace(/<\/ul>\s*<ul[^>]*>/g, '');
    // Paragraphs
    out = out.replace(/\n\n+/g, '</p><p>');
    out = out.replace(/\n/g, '<br>');
    return out;
  }

  function debounce(func, wait) {
    let timeout;
    return function(...args) {
      clearTimeout(timeout);
      timeout = setTimeout(() => func.apply(this, args), wait);
    };
  }

  // ── Initial Boot ────────────────────────────────────────────────────────────
  loadStats();
  loadRoles();

  const initialHash = (window.location.hash || '').replace('#', '');
  if (initialHash && viewMetadata[initialHash]) {
    switchView(initialHash);
  }
});
