import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
import './styles.css';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

const MAX_FILE_SIZE = 25 * 1024 * 1024;
const MAX_IMAGE_FILES = 12;
const tools = {
  image: {
    label: 'JPG to PDF',
    shortLabel: 'Images',
    eyebrow: 'Image conversion',
    title: 'Turn a stack of images into one polished PDF.',
    description: 'Drop JPEGs in, arrange the pages, and export a clean PDF without sending your files anywhere.',
    accept: '.jpg,.jpeg,image/jpeg',
    formats: 'JPG or JPEG · up to 25 MB each',
    icon: 'image',
    emptyTitle: 'Drop your images here',
    emptyText: 'or choose JPG / JPEG files from your device',
    button: 'Choose images',
    helper: 'Up to 12 images · reorder before converting',
  },
  excel: {
    label: 'Excel to PDF',
    shortLabel: 'Spreadsheets',
    eyebrow: 'Spreadsheet conversion',
    title: 'Make spreadsheets easy to share and print.',
    description: 'Convert the first worksheet into a tidy PDF with simple page and orientation controls.',
    accept: '.xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv',
    formats: 'XLSX, XLS or CSV · up to 25 MB',
    icon: 'sheet',
    emptyTitle: 'Drop your spreadsheet here',
    emptyText: 'or choose an Excel or CSV file from your device',
    button: 'Choose spreadsheet',
    helper: 'First worksheet · values are kept in browser memory',
  },
  pdf: {
    label: 'PDF to Excel',
    shortLabel: 'PDF tables',
    eyebrow: 'Table extraction',
    title: 'Bring PDF tables back into a spreadsheet.',
    description: 'Extract selectable text from a PDF into a downloadable XLSX file, ready for cleanup or analysis.',
    accept: '.pdf,application/pdf',
    formats: 'PDF · up to 25 MB · selectable text works best',
    icon: 'pdf',
    emptyTitle: 'Drop your PDF here',
    emptyText: 'or choose a PDF with selectable table text',
    button: 'Choose PDF',
    helper: 'Browser-based extraction · scanned PDFs may need OCR',
  },
  compress: {
    label: 'Compress PDF', shortLabel: 'Compress', eyebrow: 'PDF optimization',
    title: 'Make a PDF lighter without leaving your browser.',
    description: 'Rebuild image-heavy PDFs at a smaller size for quicker sharing and storage.',
    accept: '.pdf,application/pdf', formats: 'PDF · up to 25 MB', icon: 'pdf',
    emptyTitle: 'Drop your PDF here', emptyText: 'or choose a PDF to compress', button: 'Choose PDF',
    helper: 'Best for image-heavy PDFs · text may be rasterized',
  },
  crop: {
    label: 'Crop PDF', shortLabel: 'Crop', eyebrow: 'PDF framing',
    title: 'Trim the edges and keep the page focused.',
    description: 'Remove generous margins from every page and export a cleaner, tighter PDF.',
    accept: '.pdf,application/pdf', formats: 'PDF · up to 25 MB', icon: 'pdf',
    emptyTitle: 'Drop your PDF here', emptyText: 'or choose a PDF to crop', button: 'Choose PDF',
    helper: 'Choose a consistent trim amount for every page',
  },
  merge: {
    label: 'Merge PDF', shortLabel: 'Merge', eyebrow: 'PDF assembly',
    title: 'Bring separate PDFs together in one flow.',
    description: 'Arrange multiple PDFs in order and combine them into one downloadable document.',
    accept: '.pdf,application/pdf', formats: 'PDF · up to 25 MB each · up to 12 files', icon: 'pdf', multiple: true,
    emptyTitle: 'Drop your PDFs here', emptyText: 'or choose multiple PDFs to merge', button: 'Choose PDFs',
    helper: 'Up to 12 PDFs · reorder before merging',
  },
  pdfjpg: {
    label: 'PDF to JPG', shortLabel: 'PDF images', eyebrow: 'Image export',
    title: 'Turn PDF pages into easy-to-share JPG images.',
    description: 'Export every page as a crisp JPG bundle, ready for sharing or visual review.',
    accept: '.pdf,application/pdf', formats: 'PDF · up to 25 MB each · up to 12 files', icon: 'image', multiple: true,
    emptyTitle: 'Drop your PDFs here', emptyText: 'or choose PDFs to export as JPGs', button: 'Choose PDFs',
    helper: 'Every page becomes a JPG inside one ZIP download',
  },
  protect: {
    label: 'Protect PDF', shortLabel: 'Protect', eyebrow: 'PDF security',
    title: 'Add a password before you share a PDF.',
    description: 'Encrypt a PDF locally with a password so only the people you choose can open it.',
    accept: '.pdf,application/pdf', formats: 'PDF · up to 25 MB', icon: 'lock',
    emptyTitle: 'Drop your PDF here', emptyText: 'or choose a PDF to protect', button: 'Choose PDF',
    helper: 'Password protection is applied locally in your browser',
  },
  unlock: {
    label: 'Unlock PDF', shortLabel: 'Unlock', eyebrow: 'PDF security',
    title: 'Remove a PDF password you already know.',
    description: 'Enter the document password to create an editable, password-free copy in your browser.',
    accept: '.pdf,application/pdf', formats: 'PDF · up to 25 MB', icon: 'lock',
    emptyTitle: 'Drop your protected PDF here', emptyText: 'or choose a password-protected PDF', button: 'Choose PDF',
    helper: 'You must provide the current password to unlock the file',
  },
  resume: {
    label: 'Resume Maker', shortLabel: 'Resume', eyebrow: 'Career toolkit',
    title: 'Build a resume that feels like you.',
    description: 'Fill in the essentials, choose a visual direction, and download a polished resume PDF in minutes.',
    icon: 'file', helper: 'Five curated templates · edit, preview, download',
  },
};

const state = {
  activeTool: 'image',
  files: [],
  status: 'idle',
  error: '',
  output: null,
  settings: { pageSize: 'a4', orientation: 'portrait', fit: 'contain', cropInset: '0', password: '', confirmPassword: '' },
  resume: { template: 'classic', name: 'Alex Morgan', role: 'Product Designer', email: 'alex@example.com', phone: '+1 555 012 3456', location: 'New York, NY', summary: 'Product designer focused on clear systems, thoughtful details, and useful digital experiences.', experience: 'Senior Product Designer — Northstar Studio\n2022 — Present\nLed end-to-end product design for a growing SaaS platform.\n\nProduct Designer — Fieldwork\n2019 — 2022\nShaped customer journeys and launched a new design system.', education: 'BFA, Communication Design — School of Visual Arts\n2015 — 2019', skills: 'Product strategy, UX research, Prototyping, Design systems, Figma' },
  resumeOutput: null,
  notice: '',
};

const app = document.querySelector('#app');

function icon(name, size = 20) {
  const common = `width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"`;
  const paths = {
    arrow: '<path d="M5 12h14M13 6l6 6-6 6"/>',
    check: '<path d="m5 12 4 4L19 6"/>',
    chevron: '<path d="m7 10 5 5 5-5"/>',
    download: '<path d="M12 3v12m0 0 5-5m-5 5-5-5M5 21h14"/>',
    file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/>',
    image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="8.5" cy="9" r="1.5"/><path d="m21 15-4.5-4.5L7 20"/>',
    lock: '<rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/>',
    menu: '<path d="M4 6h16M4 12h16M4 18h16"/>',
    pdf: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 15h2a2 2 0 1 0 0-4H8v6m6 0v-6h1.5a2.5 2.5 0 0 1 0 5H14M8 18h8"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    refresh: '<path d="M20 11a8.1 8.1 0 0 0-14.7-3L3 11m0-5v5h5M4 13a8.1 8.1 0 0 0 14.7 3L21 13m0 5v-5h-5"/>',
    shield: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/>',
    sheet: '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 7h8M8 11h8M8 15h3M14 15h2M8 18h3M14 18h2"/>',
    spark: '<path d="m12 3 1.2 4.8L18 9l-4.8 1.2L12 15l-1.2-4.8L6 9l4.8-1.2zM19 15l.6 2.4L22 18l-2.4.6L19 21l-.6-2.4L16 18l2.4-.6z"/>',
    trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 14h10l1-14M9 7V4h6v3"/>',
    upload: '<path d="M12 16V4m0 0L7 9m5-5 5 5M4 15v4a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-4"/>',
  };
  return `<svg ${common}>${paths[name] || paths.file}</svg>`;
}

function formatBytes(bytes) {
  if (!bytes) return '0 KB';
  const units = ['B', 'KB', 'MB', 'GB'];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / Math.pow(1024, index)).toFixed(index ? 1 : 0)} ${units[index]}`;
}

function escapeHtml(value = '') {
  return String(value).replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#039;', '"': '&quot;' }[char]));
}

function humanFileName(name) {
  return name.length > 34 ? `${name.slice(0, 27)}…${name.slice(-7)}` : name;
}

function setStatus(status, error = '') {
  state.status = status;
  state.error = error;
}

function render() {
  const route = window.location.pathname;
  if (route === '/privacy') {
    app.innerHTML = legalPage('privacy');
    bindShellEvents();
    return;
  }
  if (route === '/terms') {
    app.innerHTML = legalPage('terms');
    bindShellEvents();
    return;
  }
  if (route === '/about') {
    app.innerHTML = legalPage('about');
    bindShellEvents();
    return;
  }
  app.innerHTML = homePage();
  bindHomeEvents();
}

function header() {
  return `
    <header class="site-header">
      <a class="brand" href="/" data-route="/" aria-label="AlphaConverter home">
        <span class="brand-mark">${icon('file', 23)}</span>
        <span class="brand-name">Alpha<span>Converter</span></span>
      </a>
      <nav class="desktop-nav" aria-label="Main navigation">
        <div class="nav-dropdown"><button class="nav-dropdown-trigger" type="button" aria-haspopup="true">Tools ${icon('chevron', 14)}</button><div class="nav-dropdown-panel" role="menu">${Object.entries(tools).map(([key, tool]) => `<button type="button" data-tool-link="${key}" role="menuitem"><span>${icon(tool.icon, 15)}</span>${tool.label}</button>`).join('')}</div></div>
        <a href="#workflow">How it works</a>
        <a href="#trust">Privacy first</a>
      </nav>
      <div class="header-actions">
        <span class="secure-pill">${icon('lock', 14)} Files stay on device</span>
        <a class="header-cta" href="#tools">Open converter ${icon('arrow', 16)}</a>
      </div>
      <button class="mobile-menu" type="button" aria-label="Open navigation" aria-expanded="false">${icon('menu', 21)}</button>
    </header>`;
}

function homePage() {
  return `
    <div class="site-shell">
      ${header()}
      <main>
        <section class="hero" aria-labelledby="hero-title">
          <div class="hero-copy">
            <div class="eyebrow"><span class="eyebrow-dot"></span> Private conversion, beautifully simple</div>
            <h1 id="hero-title">Your files.<br /><em>Better formats.</em></h1>
            <p class="hero-lede">Convert the documents that keep work moving — without uploading them to a server or learning another complicated tool.</p>
            <div class="hero-proof">
              <span>${icon('check', 15)} No account required</span>
              <span>${icon('check', 15)} Works in your browser</span>
              <span>${icon('check', 15)} Free to try</span>
            </div>
          </div>
          <div class="hero-art" aria-label="Illustration of neatly stacked document formats">
            <div class="art-orbit orbit-one"></div><div class="art-orbit orbit-two"></div>
            <div class="paper paper-back"><span class="paper-line long"></span><span class="paper-line"></span><span class="paper-line short"></span></div>
            <div class="paper paper-mid"><span class="paper-tab">XLSX</span><span class="paper-grid"><i></i><i></i><i></i><i></i><i></i><i></i></span></div>
            <div class="paper paper-front"><span class="paper-fold"></span><span class="paper-type">PDF</span><span class="paper-line long"></span><span class="paper-line"></span></div>
            <div class="art-caption"><span class="caption-dot"></span> One calm place for file work</div>
          </div>
        </section>

        <section class="studio-section" id="tools" aria-labelledby="studio-title">
          <div class="section-intro">
            <div><p class="section-kicker">Conversion studio</p><h2 id="studio-title">Choose a format to begin.</h2></div>
            <p class="section-note">Local-first processing means your files are never sent away for conversion.</p>
          </div>
          <button class="feature-banner" type="button" data-tool-link="resume"><span class="feature-banner-icon">${icon('spark', 20)}</span><span><strong>New · Resume Maker</strong><small>Create a polished resume from five thoughtful templates.</small></span><span class="feature-banner-action">Open builder ${icon('arrow', 16)}</span></button>
          <div class="studio-card">
            <div class="tool-tabs" role="tablist" aria-label="Conversion tools">
              ${Object.entries(tools).map(([key, tool]) => `<button class="tool-tab ${state.activeTool === key ? 'active' : ''}" role="tab" aria-selected="${state.activeTool === key}" data-tool="${key}"><span class="tab-icon">${icon(tool.icon, 19)}</span><span>${tool.label}</span></button>`).join('')}
            </div>
            <div class="tool-panel" id="tool-panel">${toolPanel()}</div>
          </div>
        </section>

        <section class="workflow-section" id="workflow" aria-labelledby="workflow-title">
          <div class="section-intro compact"><div><p class="section-kicker">A lighter workflow</p><h2 id="workflow-title">From file to finished in three beats.</h2></div></div>
          <div class="workflow-grid">
            <div class="workflow-step"><span class="step-number">01</span><div class="step-icon">${icon('upload', 22)}</div><h3>Bring your file</h3><p>Drop it in or choose it from your device. We show what is supported before you begin.</p></div>
            <div class="workflow-step"><span class="step-number">02</span><div class="step-icon">${icon('spark', 22)}</div><h3>Make it yours</h3><p>Reorder images or choose page settings. Every control stays close to the work.</p></div>
            <div class="workflow-step"><span class="step-number">03</span><div class="step-icon">${icon('download', 22)}</div><h3>Take it with you</h3><p>Download a fresh file, then start another conversion whenever you need it.</p></div>
          </div>
        </section>

        <section class="trust-section" id="trust" aria-labelledby="trust-title">
          <div class="trust-copy"><p class="section-kicker">Quietly private</p><h2 id="trust-title">Your documents are yours.<br /><em>Full stop.</em></h2><p>AlphaConverter runs conversions in your browser. Nothing is uploaded, stored, or used to train anything. When you close the tab, the working files go with it.</p><a class="text-link" href="/privacy" data-route="/privacy">Read our privacy promise ${icon('arrow', 15)}</a></div>
          <div class="trust-seal"><div class="seal-ring">${icon('shield', 38)}</div><strong>Local-first</strong><span>conversion</span><small>Built for thoughtful file handling</small></div>
        </section>
      </main>
      ${footer()}
    </div>`;
}

function footer() {
  return `<footer class="site-footer"><div class="footer-main"><a class="brand" href="/" data-route="/"><span class="brand-mark">${icon('file', 20)}</span><span class="brand-name">Alpha<span>Converter</span></span></a><p>Useful file work, without the noise.</p><div class="footer-links"><a href="/about" data-route="/about">About Us</a><a href="/privacy" data-route="/privacy">Privacy</a><a href="/terms" data-route="/terms">Terms</a><a href="#tools">Tools</a></div></div><div class="footer-bottom"><span>© ${new Date().getFullYear()} AlphaConverter</span><span>Made for the files between moments.</span></div></footer>`;
}

function toolPanel() {
  const tool = tools[state.activeTool];
  const hasFiles = state.files.length > 0;
  if (state.activeTool === 'resume') return `<div class="panel-heading"><div><p class="panel-eyebrow">${tool.eyebrow}</p><h3>${tool.title}</h3><p>${tool.description}</p></div><span class="panel-counter">08 <i>/</i> 08</span></div>${resumePanel()}<div class="panel-footer"><span>${icon('shield', 15)} Your resume is generated locally in your browser</span><span>Download as PDF</span></div>`;
  return `<div class="panel-heading"><div><p class="panel-eyebrow">${tool.eyebrow}</p><h3>${tool.title}</h3><p>${tool.description}</p></div><span class="panel-counter">${String(Object.keys(tools).indexOf(state.activeTool) + 1).padStart(2, '0')} <i>/</i> ${String(Object.keys(tools).length).padStart(2, '0')}</span></div>
    ${state.status === 'processing' ? processingState() : state.status === 'success' ? successState() : state.status === 'error' ? errorState() : uploadState(hasFiles, tool)}
    <div class="panel-footer"><span>${icon('shield', 15)} Files are processed locally in your browser</span><span>Max file size: 25 MB</span></div>`;
}

function resumePanel() {
  const r = state.resume;
  const fields = [['name', 'Full name', 'text'], ['role', 'Target role', 'text'], ['email', 'Email', 'email'], ['phone', 'Phone', 'text'], ['location', 'Location', 'text']];
  const templates = [['classic', 'Classic', 'Structured and dependable'], ['editorial', 'Editorial', 'Warm and expressive'], ['modern', 'Modern', 'Crisp and energetic'], ['minimal', 'Minimal', 'Quiet and focused'], ['bold', 'Bold', 'Confident and distinctive']];
  return `<div class="resume-builder"><div class="resume-form"><div class="resume-field-grid">${fields.map(([key, label, type]) => `<label>${label}<input class="resume-field" data-resume-field="${key}" type="${type}" value="${escapeHtml(r[key])}" /></label>`).join('')}</div><label class="resume-wide">Professional summary<textarea class="resume-field" data-resume-field="summary" rows="3">${escapeHtml(r.summary)}</textarea></label><label class="resume-wide">Experience<textarea class="resume-field" data-resume-field="experience" rows="6">${escapeHtml(r.experience)}</textarea></label><label class="resume-wide">Education<textarea class="resume-field" data-resume-field="education" rows="4">${escapeHtml(r.education)}</textarea></label><label class="resume-wide">Skills<textarea class="resume-field" data-resume-field="skills" rows="2">${escapeHtml(r.skills)}</textarea></label><div class="template-picker"><span class="template-label">Choose a template</span><div class="template-options" role="tablist" aria-label="Resume templates">${templates.map(([key, label, note]) => `<button type="button" class="template-choice ${r.template === key ? 'active' : ''}" data-template="${key}" role="tab" aria-selected="${r.template === key}"><span class="template-swatch ${key}"></span><strong>${label}</strong><small>${note}</small></button>`).join('')}</div></div><div class="resume-actions"><button class="primary-button" id="resume-download" type="button">${icon('download', 17)} Download resume PDF</button>${state.resumeOutput ? `<a class="resume-ready" href="${state.resumeOutput.url}" download="${state.resumeOutput.name}">${icon('check', 15)} Resume ready — download again</a>` : ''}</div></div><div class="resume-preview-wrap"><span class="preview-label">Live preview</span><div class="resume-preview resume-${r.template}" id="resume-preview">${resumePreview()}</div></div></div>`;
}

function resumePreview() {
  const r = state.resume;
  const lines = (value) => escapeHtml(value).split(/\n+/).map((line) => line.trim()).filter(Boolean).map((line) => `<div>${line}</div>`).join('');
  return `<header><h4>${escapeHtml(r.name || 'Your Name')}</h4><p>${escapeHtml(r.role || 'Your target role')}</p><small>${escapeHtml([r.email, r.phone, r.location].filter(Boolean).join(' · '))}</small></header><section><h5>Profile</h5><p>${escapeHtml(r.summary || 'Add a short professional summary.')}</p></section><section><h5>Experience</h5><div class="resume-lines">${lines(r.experience)}</div></section><section><h5>Education</h5><div class="resume-lines">${lines(r.education)}</div></section><section><h5>Skills</h5><p>${escapeHtml(r.skills || 'Add your key skills.')}</p></section>`;
}

function uploadState(hasFiles, tool) {
  return `<div class="upload-layout"><div class="drop-zone ${hasFiles ? 'has-files' : ''}" id="drop-zone" tabindex="0" role="button" aria-label="${tool.button}. ${tool.formats}"><input id="file-input" type="file" accept="${tool.accept}" ${tool.multiple ? 'multiple' : ''} hidden /><div class="drop-symbol">${icon(hasFiles ? 'check' : 'upload', 26)}</div><strong>${hasFiles ? `${state.files.length} file${state.files.length > 1 ? 's' : ''} ready` : tool.emptyTitle}</strong><span>${hasFiles ? 'Add more or continue below' : tool.emptyText}</span><button class="secondary-button" type="button" id="choose-files">${icon('plus', 16)} ${tool.button}</button><small>${tool.formats}</small></div>${hasFiles ? fileQueue() : helperCard(tool)}</div>${hasFiles ? settingsCard() + actionBar() : ''}`;
}

function helperCard(tool) {
  return `<div class="helper-card"><div class="helper-icon">${icon(tool.icon, 19)}</div><div><strong>Good to know</strong><p>${tool.helper}</p></div><span class="helper-arrow">${icon('arrow', 17)}</span></div>`;
}

function fileQueue() {
  const reorderable = state.activeTool === 'image' || state.activeTool === 'merge' || state.activeTool === 'pdfjpg';
  return `<div class="file-queue" aria-live="polite"><div class="queue-head"><strong>Your file${state.files.length > 1 ? 's' : ''}</strong><button class="quiet-button" id="clear-files" type="button">Clear all</button></div><div class="queue-list">${state.files.map((file, index) => `<div class="file-row" data-index="${index}"><span class="file-type ${state.activeTool}">${state.activeTool === 'image' || state.activeTool === 'pdfjpg' ? icon('image', 16) : state.activeTool === 'excel' ? icon('sheet', 16) : icon('pdf', 16)}</span><span class="file-name" title="${escapeHtml(file.name)}">${escapeHtml(humanFileName(file.name))}</span><span class="file-size">${formatBytes(file.size)}</span>${reorderable ? `<button class="icon-button move-file" data-direction="up" data-index="${index}" aria-label="Move ${escapeHtml(file.name)} up" ${index === 0 ? 'disabled' : ''}>↑</button><button class="icon-button move-file" data-direction="down" data-index="${index}" aria-label="Move ${escapeHtml(file.name)} down" ${index === state.files.length - 1 ? 'disabled' : ''}>↓</button>` : ''}<button class="icon-button remove-file" data-index="${index}" aria-label="Remove ${escapeHtml(file.name)}">×</button></div>`).join('')}</div></div>`;
}

function settingsCard() {
  if (state.activeTool === 'pdf') return '';
  if (['compress', 'merge', 'pdfjpg'].includes(state.activeTool)) return '';
  if (state.activeTool === 'crop') return `<div class="settings-row"><label>Trim amount<select id="crop-inset"><option value="0">No trim</option><option value="5">5% edge trim</option><option value="10">10% edge trim</option></select></label><span class="settings-hint">The same trim is applied to every page.</span></div>`;
  if (state.activeTool === 'protect') return `<div class="settings-row security-settings"><label>Password<input id="pdf-password" type="password" autocomplete="new-password" placeholder="Create a password" value="${escapeHtml(state.settings.password)}" /></label><label>Confirm password<input id="pdf-confirm-password" type="password" autocomplete="new-password" placeholder="Repeat the password" value="${escapeHtml(state.settings.confirmPassword)}" /></label><span class="settings-hint">Use at least 8 characters. Keep a copy of the password — it cannot be recovered.</span></div>`;
  if (state.activeTool === 'unlock') return `<div class="settings-row security-settings"><label>Current PDF password<input id="pdf-password" type="password" autocomplete="current-password" placeholder="Enter the password" value="${escapeHtml(state.settings.password)}" /></label><span class="settings-hint">The password is used only in this browser session and is never uploaded.</span></div>`;
  return `<div class="settings-row"><label>Page size<select id="page-size"><option value="a4" ${state.settings.pageSize === 'a4' ? 'selected' : ''}>A4</option><option value="letter" ${state.settings.pageSize === 'letter' ? 'selected' : ''}>US Letter</option><option value="legal" ${state.settings.pageSize === 'legal' ? 'selected' : ''}>Legal</option></select></label><label>Orientation<select id="orientation"><option value="portrait" ${state.settings.orientation === 'portrait' ? 'selected' : ''}>Portrait</option><option value="landscape" ${state.settings.orientation === 'landscape' ? 'selected' : ''}>Landscape</option></select></label>${state.activeTool === 'image' ? `<label>Image fit<select id="image-fit"><option value="contain" ${state.settings.fit === 'contain' ? 'selected' : ''}>Fit to page</option><option value="cover" ${state.settings.fit === 'cover' ? 'selected' : ''}>Fill page</option></select></label>` : ''}</div>`;
}

function actionBar() {
  const label = state.activeTool === 'image' ? 'Create PDF' : state.activeTool === 'excel' ? 'Convert to PDF' : state.activeTool === 'pdf' ? 'Extract to Excel' : state.activeTool === 'protect' ? 'Protect PDF' : state.activeTool === 'unlock' ? 'Unlock PDF' : tools[state.activeTool].label;
  return `<div class="action-bar"><span>${icon('check', 16)} Ready when you are</span><button class="primary-button" id="convert-button" type="button">${label} ${icon('arrow', 17)}</button></div>`;
}

function processingState() {
  return `<div class="state-card processing-card"><div class="processing-visual"><span></span><span></span><span></span></div><div><strong>Working on your file</strong><p>Everything is happening locally. Keep this tab open for a moment.</p></div><span class="processing-label">PROCESSING</span></div>`;
}

function successState() {
  return `<div class="state-card success-card"><div class="state-icon">${icon('check', 27)}</div><div><strong>Your file is ready.</strong><p>${escapeHtml(state.output?.description || 'The conversion finished successfully.')}</p></div><a class="primary-button download-button" href="${state.output?.url || '#'}" download="${escapeHtml(state.output?.name || 'alphaconverter-output')}">${icon('download', 17)} Download file</a></div><button class="again-button" type="button" id="start-again">${icon('refresh', 15)} Convert another file</button>`;
}

function errorState() {
  return `<div class="state-card error-card"><div class="state-icon">!</div><div><strong>That one needs another look.</strong><p>${escapeHtml(state.error || 'The file could not be converted. Please try again.')}</p></div></div><button class="again-button" type="button" id="try-again">${icon('refresh', 15)} Try again</button>`;
}

function legalPage(kind) {
  const privacy = kind === 'privacy';
  const about = kind === 'about';
  return `<div class="site-shell legal-shell">${header()}<main class="legal-main"><a class="back-link" href="/" data-route="/">${icon('arrow', 15)} Back to AlphaConverter</a><p class="section-kicker">${about ? 'About AlphaConverter' : privacy ? 'Privacy' : 'Terms of use'}</p><h1>${about ? 'Useful file work, without the noise.' : privacy ? 'A clearer way to handle files.' : 'Simple terms for a simple tool.'}</h1><p class="legal-intro">${about ? 'AlphaConverter is a calm, browser-first toolkit for turning everyday documents into the formats people need.' : privacy ? 'AlphaConverter is designed around a straightforward promise: your files stay with you.' : 'AlphaConverter is a browser-based utility for everyday document conversion.'}</p><div class="legal-content">${about ? `<h2>Built for the moment between formats</h2><p>AlphaConverter helps people convert, organize, protect, and prepare files without a complicated workflow or an account wall.</p><h2>Private by default</h2><p>Our tools are designed to process files in your browser. That means your documents stay on your device while you work, with clear controls and simple downloads when you are done.</p><h2>Thoughtful utility</h2><p>From JPG and Excel conversion to PDF security and resume creation, AlphaConverter focuses on practical tools that feel straightforward, respectful, and easy to return to.</p>` : privacy ? `<h2>Local-first by design</h2><p>Conversions run in your browser. AlphaConverter does not upload your selected files to an AlphaConverter server, create an account record, or use your documents for training.</p><h2>What your browser keeps</h2><p>Your browser temporarily reads the files you select so it can convert them. Generated downloads are created in your browser and are not stored by AlphaConverter. Clear your browser tab or close the page to release working memory.</p><h2>Third-party libraries</h2><p>The converter uses open-source document libraries loaded with the site. They process the file data inside the current browser session.</p><h2>Questions</h2><p>If this product evolves to use server-side processing, this page will be updated before that change takes effect.</p>` : `<h2>Use the tool responsibly</h2><p>You are responsible for the files you choose and for reviewing converted outputs before sharing or relying on them.</p><h2>Conversion quality</h2><p>Different source files can produce different results. PDF table extraction works best with selectable text and clear tabular structure. AlphaConverter does not guarantee that every conversion will preserve every visual or structural detail.</p><h2>Availability</h2><p>We aim to keep the tool useful and available, but browser capabilities, file complexity, and future changes can affect results.</p><h2>Updates</h2><p>We may refine the interface and conversion libraries over time. The current terms are always available on this page.</p>`}</div></main>${footer()}</div>`;
}

function bindShellEvents() {
  document.querySelectorAll('[data-route]').forEach((link) => link.addEventListener('click', (event) => {
    const target = link.getAttribute('href');
    if (!target || target.startsWith('#')) return;
    event.preventDefault();
    window.history.pushState({}, '', target);
    window.scrollTo({ top: 0, behavior: 'smooth' });
    render();
  }));
  const menu = document.querySelector('.mobile-menu');
  if (menu) menu.addEventListener('click', () => {
    const nav = document.querySelector('.desktop-nav');
    const open = nav.classList.toggle('mobile-open');
    menu.setAttribute('aria-expanded', String(open));
  });
}

function bindHomeEvents() {
  bindShellEvents();
  document.querySelectorAll('[data-tool-link]').forEach((link) => link.addEventListener('click', () => {
    state.activeTool = link.dataset.toolLink;
    state.files = [];
    state.output = null;
    state.settings.password = '';
    state.settings.confirmPassword = '';
    setStatus('idle');
    render();
    document.querySelector('#tools')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }));
  document.querySelectorAll('.resume-field').forEach((field) => field.addEventListener('input', () => {
    state.resume[field.dataset.resumeField] = field.value;
    const preview = document.querySelector('#resume-preview');
    if (preview) preview.innerHTML = resumePreview();
  }));
  document.querySelectorAll('.template-choice').forEach((choice) => choice.addEventListener('click', () => {
    state.resume.template = choice.dataset.template;
    render();
  }));
  document.querySelector('#resume-download')?.addEventListener('click', downloadResume);
  document.querySelectorAll('.tool-tab').forEach((tab) => tab.addEventListener('click', () => {
    state.activeTool = tab.dataset.tool;
    state.files = [];
    state.output = null;
    state.settings.password = '';
    state.settings.confirmPassword = '';
    state.notice = '';
    setStatus('idle');
    render();
    document.querySelector('#tools')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }));

  const dropZone = document.querySelector('#drop-zone');
  const input = document.querySelector('#file-input');
  if (dropZone && input) {
    document.querySelector('#choose-files')?.addEventListener('click', (event) => { event.stopPropagation(); input.click(); });
    dropZone.addEventListener('click', (event) => { if (event.target.closest('button')) return; input.click(); });
    dropZone.addEventListener('keydown', (event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); input.click(); } });
    input.addEventListener('change', (event) => acceptFiles([...event.target.files]));
    ['dragenter', 'dragover'].forEach((eventName) => dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.add('dragging'); }));
    ['dragleave', 'drop'].forEach((eventName) => dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.remove('dragging'); }));
    dropZone.addEventListener('drop', (event) => acceptFiles([...event.dataTransfer.files]));
  }
  document.querySelector('#clear-files')?.addEventListener('click', () => { state.files = []; setStatus('idle'); render(); });
  document.querySelectorAll('.remove-file').forEach((button) => button.addEventListener('click', () => { state.files.splice(Number(button.dataset.index), 1); setStatus(state.files.length ? 'ready' : 'idle'); render(); }));
  document.querySelectorAll('.move-file').forEach((button) => button.addEventListener('click', () => {
    const from = Number(button.dataset.index); const to = button.dataset.direction === 'up' ? from - 1 : from + 1;
    if (to < 0 || to >= state.files.length) return;
    [state.files[from], state.files[to]] = [state.files[to], state.files[from]]; render();
  }));
  document.querySelector('#page-size')?.addEventListener('change', (event) => { state.settings.pageSize = event.target.value; });
  document.querySelector('#orientation')?.addEventListener('change', (event) => { state.settings.orientation = event.target.value; });
  document.querySelector('#image-fit')?.addEventListener('change', (event) => { state.settings.fit = event.target.value; });
  document.querySelector('#crop-inset')?.addEventListener('change', (event) => { state.settings.cropInset = event.target.value; });
  document.querySelector('#pdf-password')?.addEventListener('input', (event) => { state.settings.password = event.target.value; });
  document.querySelector('#pdf-confirm-password')?.addEventListener('input', (event) => { state.settings.confirmPassword = event.target.value; });
  document.querySelector('#convert-button')?.addEventListener('click', convert);
  document.querySelector('#start-again')?.addEventListener('click', () => { state.files = []; state.output = null; setStatus('idle'); render(); });
  document.querySelector('#try-again')?.addEventListener('click', () => { state.output = null; setStatus('ready'); render(); });
}

function downloadResume() {
  const r = state.resume;
  const colors = { classic: ['#245C4F', '#F3F7EF'], editorial: ['#B7654F', '#FBF0E8'], modern: ['#234D73', '#EAF2F8'], minimal: ['#3D4843', '#F4F4F1'], bold: ['#7B4D2E', '#F6EBDD'] };
  const [accent, tint] = colors[r.template] || colors.classic;
  const doc = new jsPDF({ unit: 'pt', format: 'a4' }); const margin = 42; const width = doc.internal.pageSize.getWidth(); let y = 52;
  doc.setFillColor(tint); doc.rect(0, 0, width, 128, 'F'); doc.setTextColor(accent); doc.setFont('helvetica', 'bold'); doc.setFontSize(27); doc.text(r.name || 'Your Name', margin, y); y += 25; doc.setFont('helvetica', 'normal'); doc.setFontSize(12); doc.text(r.role || 'Your target role', margin, y); y += 21; doc.setTextColor('#68756d'); doc.setFontSize(9); doc.text([r.email, r.phone, r.location].filter(Boolean).join('  ·  '), margin, y); y = 165;
  const section = (title, text, size = 10) => { doc.setTextColor(accent); doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.text(title.toUpperCase(), margin, y); y += 15; doc.setTextColor('#3b4841'); doc.setFont('helvetica', 'normal'); doc.setFontSize(size); const lines = doc.splitTextToSize(text || '', width - margin * 2); doc.text(lines, margin, y); y += lines.length * (size + 4) + 17; };
  section('Profile', r.summary); section('Experience', r.experience); section('Education', r.education); section('Skills', r.skills);
  const url = URL.createObjectURL(doc.output('blob')); state.resumeOutput = { url, name: `${(r.name || 'resume').toLowerCase().replace(/[^a-z0-9]+/g, '-')}-resume.pdf` }; render();
}

function validateFile(file) {
  const name = file.name.toLowerCase();
  const valid = state.activeTool === 'image' ? /\.(jpe?g)$/i.test(name) || file.type === 'image/jpeg' : state.activeTool === 'excel' ? /\.(xlsx?|csv)$/i.test(name) || /spreadsheet|excel|csv/.test(file.type) : /\.pdf$/i.test(name) || file.type === 'application/pdf';
  if (!valid) return `“${file.name}” is not a supported ${tools[state.activeTool].label} file.`;
  if (file.size > MAX_FILE_SIZE) return `“${file.name}” is larger than the 25 MB limit.`;
  return '';
}

function acceptFiles(incoming) {
  const allowsMultiple = state.activeTool === 'image' || tools[state.activeTool].multiple;
  const files = allowsMultiple ? incoming : incoming.slice(0, 1);
  const issue = files.map(validateFile).find(Boolean);
  if (issue) { setStatus('error', issue); render(); return; }
  if (allowsMultiple && state.files.length + files.length > MAX_IMAGE_FILES) { setStatus('error', `You can combine up to ${MAX_IMAGE_FILES} files at once.`); render(); return; }
  state.files = allowsMultiple ? [...state.files, ...files] : files;
  state.output = null;
  setStatus('ready');
  render();
}

function readDataUrl(file) {
  return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(reader.result); reader.onerror = () => reject(new Error('Could not read the selected file.')); reader.readAsDataURL(file); });
}

function loadImage(src) {
  return new Promise((resolve, reject) => { const image = new Image(); image.onload = () => resolve(image); image.onerror = () => reject(new Error('One of the images could not be decoded.')); image.src = src; });
}

function pageFormat() {
  return state.settings.pageSize === 'letter' ? 'letter' : state.settings.pageSize === 'legal' ? 'legal' : 'a4';
}

async function imagesToPdf() {
  const doc = new jsPDF({ orientation: state.settings.orientation, unit: 'pt', format: pageFormat() });
  for (let index = 0; index < state.files.length; index += 1) {
    if (index > 0) doc.addPage(pageFormat(), state.settings.orientation);
    const data = await readDataUrl(state.files[index]);
    const image = await loadImage(data);
    const pageWidth = doc.internal.pageSize.getWidth(); const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 30; const maxWidth = pageWidth - margin * 2; const maxHeight = pageHeight - margin * 2;
    const ratio = state.settings.fit === 'cover' ? Math.max(maxWidth / image.width, maxHeight / image.height) : Math.min(maxWidth / image.width, maxHeight / image.height);
    const width = image.width * ratio; const height = image.height * ratio;
    const x = (pageWidth - width) / 2; const y = (pageHeight - height) / 2;
    doc.addImage(data, 'JPEG', x, y, width, height, undefined, 'FAST');
  }
  return { blob: doc.output('blob'), name: 'alphaconverter-images.pdf', description: `${state.files.length} image${state.files.length > 1 ? 's' : ''} combined into one PDF.` };
}

async function excelToPdf() {
  const workbook = XLSX.read(await state.files[0].arrayBuffer(), { type: 'array', cellText: true, cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' }).filter((row) => row.some((cell) => String(cell).trim() !== ''));
  if (!rows.length) throw new Error('This spreadsheet does not contain any visible rows.');
  const doc = new jsPDF({ orientation: state.settings.orientation === 'portrait' ? 'landscape' : 'portrait', unit: 'pt', format: pageFormat() });
  doc.setFontSize(15); doc.setTextColor('#245C4F'); doc.text(sheetName || 'Spreadsheet', 30, 34);
  autoTable(doc, { head: [rows[0].map((cell) => String(cell))], body: rows.slice(1).map((row) => row.map((cell) => String(cell))), startY: 50, margin: { left: 30, right: 30 }, theme: 'grid', styles: { font: 'helvetica', fontSize: 8, cellPadding: 5, overflow: 'linebreak' }, headStyles: { fillColor: [36, 92, 79], textColor: [255, 250, 240], fontStyle: 'bold' }, alternateRowStyles: { fillColor: [248, 245, 237] } });
  return { blob: doc.output('blob'), name: `${state.files[0].name.replace(/\.[^/.]+$/, '') || 'spreadsheet'}.pdf`, description: `${rows.length} rows from “${sheetName}” formatted as a PDF.` };
}

async function pdfToExcel() {
  const pdf = await pdfjsLib.getDocument({ data: await state.files[0].arrayBuffer() }).promise;
  const outputRows = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const grouped = new Map();
    content.items.forEach((item) => {
      const x = item.transform[4]; const y = Math.round(item.transform[5] / 4) * 4;
      if (!grouped.has(y)) grouped.set(y, []);
      grouped.get(y).push({ x, text: item.str.trim() });
    });
    [...grouped.entries()].sort((a, b) => b[0] - a[0]).forEach(([, items]) => {
      const cells = items.filter((item) => item.text).sort((a, b) => a.x - b.x).map((item) => item.text);
      if (cells.length) outputRows.push(cells);
    });
    if (pageNumber < pdf.numPages) outputRows.push([]);
  }
  if (!outputRows.length) throw new Error('No selectable text was found. A scanned PDF may need OCR before it can be extracted.');
  const sheet = XLSX.utils.aoa_to_sheet(outputRows);
  sheet['!cols'] = Array.from({ length: Math.max(...outputRows.map((row) => row.length)) }, () => ({ wch: 22 }));
  const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, sheet, 'Extracted tables');
  const bytes = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  return { blob: new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name: `${state.files[0].name.replace(/\.[^/.]+$/, '') || 'extracted'}.xlsx`, description: `${pdf.numPages} page${pdf.numPages > 1 ? 's' : ''} extracted into an Excel workbook.` };
}

async function renderPdfPages(file, scale = 1.25, password = '') {
  const loadingTask = pdfjsLib.getDocument({ data: await file.arrayBuffer(), password: password || undefined });
  const pdf = await loadingTask.promise;
  const pages = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
    pages.push({ canvas, width: viewport.width, height: viewport.height });
  }
  return pages;
}

async function pagesToPdf(pages, options = {}) {
  let doc;
  pages.forEach((page, index) => {
    const width = page.canvas.width; const height = page.canvas.height;
    const orientation = width > height ? 'landscape' : 'portrait';
    if (index === 0) doc = new jsPDF({ orientation, unit: 'pt', format: [width, height], ...(options.encryption ? { encryption: options.encryption } : {}) });
    else doc.addPage([width, height], orientation);
    doc.addImage(page.canvas.toDataURL('image/jpeg', options.quality || 0.92), 'JPEG', 0, 0, width, height, undefined, 'FAST');
  });
  return doc;
}

async function protectPdf() {
  const password = state.settings.password.trim();
  const confirmPassword = state.settings.confirmPassword.trim();
  if (password.length < 8) throw new Error('Choose a password with at least 8 characters.');
  if (password !== confirmPassword) throw new Error('The password and confirmation do not match.');
  const pages = await renderPdfPages(state.files[0], 1.25);
  if (!pages.length) throw new Error('No pages were found in the selected PDF.');
  const doc = await pagesToPdf(pages, { encryption: { userPassword: password, ownerPassword: `${password}-owner`, userPermissions: ['print', 'copy'] } });
  return { blob: doc.output('blob'), name: `${state.files[0].name.replace(/\.[^/.]+$/, '') || 'document'}-protected.pdf`, description: `${pages.length} page${pages.length > 1 ? 's' : ''} encrypted with a password locally in your browser.` };
}

async function unlockPdf() {
  const password = state.settings.password.trim();
  if (!password) throw new Error('Enter the current PDF password to unlock this file.');
  const pages = await renderPdfPages(state.files[0], 1.25, password);
  if (!pages.length) throw new Error('No pages were found in the selected PDF.');
  const doc = await pagesToPdf(pages, { quality: 0.94 });
  return { blob: doc.output('blob'), name: `${state.files[0].name.replace(/\.[^/.]+$/, '') || 'document'}-unlocked.pdf`, description: `${pages.length} page${pages.length > 1 ? 's' : ''} exported as a password-free PDF.` };
}

function pageCanvas(page, inset = 0) {
  if (!inset) return page.canvas;
  const crop = document.createElement('canvas');
  const x = Math.round(page.canvas.width * inset); const y = Math.round(page.canvas.height * inset);
  crop.width = Math.max(1, page.canvas.width - x * 2); crop.height = Math.max(1, page.canvas.height - y * 2);
  crop.getContext('2d').drawImage(page.canvas, x, y, crop.width, crop.height, 0, 0, crop.width, crop.height);
  return crop;
}

async function pdfsToPdf() {
  const isCompress = state.activeTool === 'compress';
  const isCrop = state.activeTool === 'crop';
  const quality = isCompress ? 0.55 : state.activeTool === 'merge' ? 0.78 : 0.86;
  const scale = isCompress ? 0.95 : 1.25;
  const inset = isCrop ? Number(state.settings.cropInset || 0) / 100 : 0;
  const pages = [];
  for (const file of state.files) pages.push(...await renderPdfPages(file, scale));
  if (!pages.length) throw new Error('No pages were found in the selected PDF.');
  let doc;
  const croppedPages = pages.map((page) => ({ ...page, canvas: pageCanvas(page, inset) }));
  doc = await pagesToPdf(croppedPages, { quality });
  const name = isCompress ? 'alphaconverter-compressed.pdf' : isCrop ? 'alphaconverter-cropped.pdf' : 'alphaconverter-merged.pdf';
  const description = isCompress ? `${pages.length} page${pages.length > 1 ? 's' : ''} rebuilt at a lighter image weight.` : isCrop ? `${pages.length} page${pages.length > 1 ? 's' : ''} cropped with a ${state.settings.cropInset || 0}% edge trim.` : `${pages.length} page${pages.length > 1 ? 's' : ''} combined from ${state.files.length} PDF${state.files.length > 1 ? 's' : ''}.`;
  return { blob: doc.output('blob'), name, description };
}

async function pdfToJpg() {
  const zip = new JSZip(); let pageCount = 0;
  for (const file of state.files) {
    const pages = await renderPdfPages(file, 1.5); const base = file.name.replace(/\.[^/.]+$/, '') || 'document';
    pages.forEach((page, index) => { const data = page.canvas.toDataURL('image/jpeg', 0.9).split(',')[1]; zip.file(`${base}-${String(index + 1).padStart(2, '0')}.jpg`, data, { base64: true }); pageCount += 1; });
  }
  return { blob: await zip.generateAsync({ type: 'blob' }), name: 'alphaconverter-jpg-pages.zip', description: `${pageCount} PDF page${pageCount > 1 ? 's' : ''} exported as JPG images inside a ZIP.` };
}

async function convert() {
  if (!state.files.length) return;
  setStatus('processing'); state.output = null; render();
  try {
    const result = state.activeTool === 'image' ? await imagesToPdf() : state.activeTool === 'excel' ? await excelToPdf() : state.activeTool === 'pdf' ? await pdfToExcel() : state.activeTool === 'pdfjpg' ? await pdfToJpg() : state.activeTool === 'protect' ? await protectPdf() : state.activeTool === 'unlock' ? await unlockPdf() : await pdfsToPdf();
    result.url = URL.createObjectURL(result.blob);
    state.output = result; setStatus('success'); render();
  } catch (error) {
    console.error(error); setStatus('error', error?.message || 'The conversion could not be completed.'); render();
  }
}

window.addEventListener('popstate', render);
render();
