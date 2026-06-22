// ─────────────────────────────────────────────────────────────
// PDF Assignment Renamer — app.js
// ─────────────────────────────────────────────────────────────

// pdfjsLib is initialised lazily in initPdfJs() so a failed import
// does NOT prevent the folder/Excel pickers from working.
let pdfjsLib = null;

async function initPdfJs() {
  if (pdfjsLib) return;
  pdfjsLib = await import('./lib/pdf.min.js');
  pdfjsLib.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL('lib/pdf.worker.min.mjs');
}

// ─────────────────────────────────────────────────────────────
// STATE
// ─────────────────────────────────────────────────────────────
const state = {
  dirHandle: null,        // FileSystemDirectoryHandle for the PDF folder
  pdfFiles: [],           // [{ name, handle }]
  students: [],           // [{ name, atsId }] from Excel
  results: [],            // one entry per PDF after OCR
  tesseractWorker: null,
  courseCode: '',         // confirmed by user after auto-detection
  criteriaType:  '',      // 'GC' or 'PC' — confirmed by user
  criteriaStart: '',      // e.g. '5.1'
  criteriaEnd:   '',      // e.g. '5.3'
  finalReport: [],        // populated after apply
};

// ─────────────────────────────────────────────────────────────
// STEP NAVIGATION
// ─────────────────────────────────────────────────────────────
function showStep(n) {
  [1, 2, 3, 4].forEach(i => {
    document.getElementById(`step-${i}`).classList.toggle('hidden', i !== n);
    const tab = document.getElementById(`tab-${i}`);
    tab.classList.remove('active', 'done');
    if (i === n) tab.classList.add('active');
    if (i < n) tab.classList.add('done');
  });
}

// ─────────────────────────────────────────────────────────────
// STEP 1 — FILE LOADING
// ─────────────────────────────────────────────────────────────
async function pickFolder() {
  try {
    state.dirHandle = await window.showDirectoryPicker();
    state.pdfFiles = [];
    for await (const [name, handle] of state.dirHandle.entries()) {
      if (handle.kind === 'file' && name.toLowerCase().endsWith('.pdf')) {
        state.pdfFiles.push({ name, handle });
      }
    }
    state.pdfFiles.sort((a, b) => a.name.localeCompare(b.name));

    const zone = document.getElementById('zone-pdf');
    zone.classList.add('loaded');
    zone.removeEventListener('click', pickFolder);
    zone.querySelector('h3').textContent = `${state.pdfFiles.length} PDF files loaded`;
    zone.querySelector('p').textContent = state.dirHandle.name;
    const tag = zone.querySelector('.status-tag');
    tag.textContent = `${state.pdfFiles.length} files`;
    tag.classList.remove('hidden');
    updateStartBtn();
  } catch (e) {
    if (e.name !== 'AbortError') alert('Could not open folder: ' + e.message);
  }
};

async function pickExcel() {
  try {
    const [fh] = await window.showOpenFilePicker({
      types: [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls'] } }],
    });
    const file = await fh.getFile();
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Find the header row (contains "Student Name" or "Name")
    let headerRow = -1;
    let nameCol = 1, idCol = 2;
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const rowStr = rows[r].map(c => String(c).toLowerCase()).join('|');
      if (rowStr.includes('student name') || rowStr.includes('student id')) {
        headerRow = r;
        // Find column indexes
        rows[r].forEach((cell, ci) => {
          const s = String(cell).toLowerCase();
          if (s.includes('name')) nameCol = ci;
          if (s.includes('id')) idCol = ci;
        });
        break;
      }
    }

    state.students = [];
    for (let r = headerRow + 1; r < rows.length; r++) {
      const row = rows[r];
      const name = String(row[nameCol] || '').trim();
      const atsId = String(row[idCol] || '').trim();
      if (name.length > 3 && !/^\d+$/.test(name)) {
        state.students.push({ name, atsId });
      }
    }

    const zone = document.getElementById('zone-excel');
    zone.classList.add('loaded');
    zone.removeEventListener('click', pickExcel);
    zone.querySelector('h3').textContent = `${state.students.length} students loaded`;
    zone.querySelector('p').textContent = file.name;
    const tag = zone.querySelector('.status-tag');
    tag.textContent = `${state.students.length} students`;
    tag.classList.remove('hidden');
    updateStartBtn();
  } catch (e) {
    if (e.name !== 'AbortError') alert('Could not read Excel file: ' + e.message);
  }
};

function updateStartBtn() {
  const btn   = document.getElementById('start-btn');
  const ready = state.pdfFiles.length > 0 && state.students.length > 0;
  btn.disabled = !ready;
  if (ready) {
    btn.innerHTML = `<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
    Start OCR Processing (${state.pdfFiles.length} PDFs × ${state.students.length} students)`;
  }
}

// ─────────────────────────────────────────────────────────────
// COURSE CODE CONFIRMATION
// ─────────────────────────────────────────────────────────────
function confirmCourseCode(detected) {
  return new Promise(resolve => {
    const panel         = document.getElementById('course-code-confirm');
    const panelDetected = document.getElementById('cc-panel-detected');
    const panelManual   = document.getElementById('cc-panel-manual');
    const detectedValue = document.getElementById('cc-detected-value');
    const yesBtn        = document.getElementById('cc-yes-btn');
    const noBtn         = document.getElementById('cc-no-btn');
    const input         = document.getElementById('confirm-code-input');
    const confirmBtn    = document.getElementById('confirm-code-btn');
    const manualLabel   = document.getElementById('cc-manual-label');

    panel.classList.remove('hidden');

    const finish = code => {
      panel.classList.add('hidden');
      panelDetected.classList.add('hidden');
      panelManual.classList.add('hidden');
      resolve(code.toUpperCase());
    };

    if (detected && detected.length >= 2) {
      // Show Yes/No panel with detected value
      detectedValue.textContent = detected.toUpperCase();
      panelDetected.classList.remove('hidden');
      panelManual.classList.add('hidden');

      yesBtn.onclick = () => finish(detected);

      noBtn.onclick = () => {
        // Switch to manual input panel
        panelDetected.classList.add('hidden');
        manualLabel.textContent = 'The detected code was wrong — enter the correct one:';
        input.value = detected.toUpperCase();
        input.style.borderColor = '';
        panelManual.classList.remove('hidden');
        input.focus();
        input.select();
      };
    } else {
      // Nothing detected — go straight to manual input
      panelDetected.classList.add('hidden');
      manualLabel.textContent = 'Course code not found by OCR — please type it:';
      input.value = '';
      panelManual.classList.remove('hidden');
      input.focus();
    }

    // Manual input confirm handler
    input.oninput = () => { input.value = input.value.toUpperCase(); };

    const submitManual = () => {
      const code = input.value.trim().toUpperCase();
      if (code.length < 2) {
        input.style.borderColor = '#ef4444';
        input.focus();
        return;
      }
      finish(code);
    };

    confirmBtn.onclick = submitManual;
    input.onkeydown = e => { if (e.key === 'Enter') submitManual(); };
  });
}

// ─────────────────────────────────────────────────────────────
// CRITERIA DETECTION (pages 1 and 2) + CONFIRMATION
// ─────────────────────────────────────────────────────────────
async function detectCriteria(fileHandle) {
  const file        = await fileHandle.getFile();
  const arrayBuffer = await file.arrayBuffer();
  const pdf         = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) }).promise;
  const pagesToScan = Math.min(pdf.numPages, 2);

  let combinedText = '';
  for (let p = 1; p <= pagesToScan; p++) {
    const page     = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: 3.0 });
    const canvas   = document.createElement('canvas');
    canvas.width   = viewport.width;
    canvas.height  = viewport.height;
    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
    const { data: { text } } = await state.tesseractWorker.recognize(canvas);
    combinedText += '\n' + text;
  }
  await pdf.destroy();

  // Collect ALL GC/PC criterion numbers from both pages
  const allMatches = [...combinedText.matchAll(/\b(GC|PC)\s*(\d+\.\d+)/gi)];
  if (allMatches.length === 0) return null;

  const type    = allMatches[0][1].toUpperCase();
  const numbers = allMatches
    .map(m => parseFloat(m[2]))
    .filter(n => !isNaN(n))
    .sort((a, b) => a - b);

  return {
    type,
    start: numbers[0].toFixed(1),
    end:   numbers[numbers.length - 1].toFixed(1),
    count: numbers.length,
  };
}

function confirmCriteria(detected) {
  return new Promise(resolve => {
    const panel        = document.getElementById('criteria-confirm');
    const panelDet     = document.getElementById('crit-panel-detected');
    const panelMan     = document.getElementById('crit-panel-manual');
    const detValue     = document.getElementById('crit-detected-value');
    const yesBtn       = document.getElementById('crit-yes-btn');
    const noBtn        = document.getElementById('crit-no-btn');
    const typeSelect   = document.getElementById('crit-type-select');
    const startInput   = document.getElementById('crit-start-input');
    const endInput     = document.getElementById('crit-end-input');
    const confirmBtn   = document.getElementById('crit-confirm-btn');
    const errorDiv     = document.getElementById('crit-manual-error');
    const manualLabel  = document.getElementById('crit-manual-label');

    panel.classList.remove('hidden');

    const finish = (type, start, end) => {
      panel.classList.add('hidden');
      panelDet.classList.add('hidden');
      panelMan.classList.add('hidden');
      resolve({ type: type.toUpperCase(), start, end });
    };

    const showManual = (prefill) => {
      panelDet.classList.add('hidden');
      if (prefill) {
        typeSelect.value  = prefill.type;
        startInput.value  = prefill.start;
        endInput.value    = prefill.end;
        manualLabel.textContent = 'Change the criteria range:';
      } else {
        manualLabel.textContent = 'No criteria detected — please enter the range:';
      }
      errorDiv.style.display = 'none';
      panelMan.classList.remove('hidden');
      startInput.focus();
    };

    if (detected) {
      detValue.textContent = `${detected.type} ${detected.start} – ${detected.end}`;
      panelDet.classList.remove('hidden');
      panelMan.classList.add('hidden');

      yesBtn.onclick = () => finish(detected.type, detected.start, detected.end);
      noBtn.onclick  = () => showManual(detected);
    } else {
      showManual(null);
    }

    confirmBtn.onclick = () => {
      const type  = typeSelect.value;
      const start = startInput.value.trim();
      const end   = endInput.value.trim();
      if (!start || !end) {
        errorDiv.style.display = 'block';
        return;
      }
      errorDiv.style.display = 'none';
      finish(type, start, end);
    };
  });
}

// ─────────────────────────────────────────────────────────────
// STEP 2 — OCR PIPELINE
// ─────────────────────────────────────────────────────────────
async function startProcessing() {
  showStep(2);
  addLog('Initialising Tesseract OCR engine…', 'ok');

  try {
    addLog('Loading PDF renderer…', 'ok');
    await initPdfJs();
    addLog('PDF renderer ready ✓', 'ok');

    // Verify Tesseract UMD global is available
    if (typeof Tesseract === 'undefined') {
      throw new Error('Tesseract global not found — lib/tesseract.min.js may have failed to load');
    }
    addLog('Tesseract global found ✓', 'ok');

    const workerPath = chrome.runtime.getURL('lib/worker.min.js');
    const corePath   = chrome.runtime.getURL('lib/');
    const langPath   = 'https://tessdata.projectnaptha.com/4.0.0'; // no trailing slash per API docs

    addLog(`Worker: ${workerPath}`, 'ok');

    // ── Worker constructor intercept ────────────────────────────
    // Tesseract v7 does: new Worker(URL.createObjectURL(blob))
    // where the blob contains: importScripts('chrome-extension://...lib/worker.min.js')
    // Blob-URL workers cannot importScripts from chrome-extension:// origins.
    // Fix: when Tesseract creates a blob: URL worker, swap it for the extension URL directly.
    const _OrigWorker = window.Worker;
    window.Worker = function(url, opts) {
      return new _OrigWorker(
        (typeof url === 'string' && url.startsWith('blob:')) ? workerPath : url,
        opts
      );
    };
    window.Worker.prototype = _OrigWorker.prototype;

    try {
      state.tesseractWorker = await Tesseract.createWorker('eng', Tesseract.OEM?.LSTM_ONLY ?? 1, {
        workerPath,
        corePath,
        langPath,
        logger: m => {
          if (m.status === 'loading tesseract core') updateProgress(null, 'Loading OCR engine…');
          if (m.status === 'loading language traineddata') updateProgress(null, 'Downloading English language model (~10 MB, first run only)…');
          if (m.status === 'initializing api') updateProgress(null, 'Initialising OCR API…');
        },
      });
    } finally {
      window.Worker = _OrigWorker; // restore regardless of success/failure
    }

    addLog('OCR engine ready ✓', 'ok');

    // ── Auto-detect course code from the first PDF ──
    addLog('Scanning first PDF to detect course code…', 'ok');
    try {
      const sampleOcr    = await ocrFirstPage(state.pdfFiles[0].handle);
      const sampleParsed = parseOcrText(sampleOcr);
      // Debug: show raw OCR snippet so we can see what Tesseract actually read
      addLog('OCR raw (first 300 chars): ' + sampleOcr.replace(/\n/g, ' ').slice(0, 300), 'ok');

      const detected = sampleParsed.courseCode || '';
      if (detected) {
        addLog(`Course code detected: ${detected} — awaiting confirmation…`, 'ok');
      } else {
        addLog('Course code not detected by OCR — please enter it manually.', 'warn');
      }
      state.courseCode = await confirmCourseCode(detected);
      addLog(`Course code set to: ${state.courseCode}`, 'ok');
    } catch (e) {
      addLog('Course code scan failed — please enter it manually.', 'warn');
      state.courseCode = await confirmCourseCode('');
    }

    // ── Detect GC/PC criteria range from pages 1–2 ──
    addLog('Scanning pages 1–2 to detect GC/PC criteria range…', 'ok');
    try {
      const critDetected = await detectCriteria(state.pdfFiles[0].handle);
      if (critDetected) {
        addLog(`Criteria detected: ${critDetected.type} ${critDetected.start}–${critDetected.end} (${critDetected.count} criteria found)`, 'ok');
      } else {
        addLog('No GC/PC criteria detected — please enter manually.', 'warn');
      }
      const confirmed = await confirmCriteria(critDetected);
      state.criteriaType  = confirmed.type;
      state.criteriaStart = confirmed.start;
      state.criteriaEnd   = confirmed.end;
      addLog(`Criteria confirmed: ${state.criteriaType} ${state.criteriaStart}–${state.criteriaEnd}`, 'ok');
    } catch (e) {
      addLog('Criteria detection failed — please enter manually.', 'warn');
      const confirmed = await confirmCriteria(null);
      state.criteriaType  = confirmed.type;
      state.criteriaStart = confirmed.start;
      state.criteriaEnd   = confirmed.end;
    }

    state.results = [];

    for (let i = 0; i < state.pdfFiles.length; i++) {
      const pdf = state.pdfFiles[i];
      const pct = Math.round((i / state.pdfFiles.length) * 100);
      updateProgress(pct, `Processing ${i + 1} of ${state.pdfFiles.length}: ${pdf.name}`);

      try {
        const ocr = await ocrFirstPage(pdf.handle);
        const parsed = parseOcrText(ocr);
        const filenameHint = extractNameFromFilename(pdf.name);
        const match = matchStudent(parsed.learnerName, filenameHint, parsed.atsId, state.students);

        const newName = match.student
          ? buildFilename(parsed.courseCode, parsed.criteriaType, parsed.criteriaStart, parsed.criteriaEnd, match.student.atsId)
          : null;

        state.results.push({ pdf, parsed, filenameHint, match, newName, selected: !!match.student });

        const statusText = match.status === 'matched' ? '✓ matched' : match.status === 'low-confidence' ? '⚠ low confidence' : '✗ needs review';
        const logClass = match.status === 'matched' ? 'ok' : match.status === 'low-confidence' ? 'warn' : 'err';
        addLog(`[${i + 1}/${state.pdfFiles.length}] ${pdf.name} → ${statusText}`, logClass);

      } catch (e) {
        addLog(`[${i + 1}/${state.pdfFiles.length}] ${pdf.name} → ERROR: ${e.message}`, 'err');
        state.results.push({
          pdf, parsed: {}, filenameHint: extractNameFromFilename(pdf.name),
          match: { status: 'error', candidates: state.students.slice(0, 3) },
          newName: null, selected: false,
        });
      }
    }

    updateProgress(100, 'Processing complete!');
    await state.tesseractWorker.terminate();
    state.tesseractWorker = null;
    addLog('All PDFs processed. Building results…', 'ok');

    setTimeout(() => {
      buildResultsTable();
      showStep(3);
    }, 600);

  } catch (e) {
    const msg = (e instanceof Error)
      ? (e.message || e.toString())
      : (typeof e === 'string' ? e : JSON.stringify(e) ?? String(e));
    addLog('Fatal error: ' + msg, 'err');
    if (e?.stack) addLog(e.stack.split('\n').slice(0,3).join(' | '), 'err');
    console.error('startProcessing error:', e);
  }
};

function updateProgress(pct, label) {
  if (pct !== null) {
    document.getElementById('prog-fill').style.width = pct + '%';
    document.getElementById('prog-pct').textContent = pct + '%';
  }
  if (label) {
    document.getElementById('prog-label').textContent = label;
    document.getElementById('prog-current').textContent = label;
  }
}

function addLog(msg, cls = '') {
  const log = document.getElementById('ocr-log');
  const line = document.createElement('div');
  line.className = 'ocr-log-line' + (cls ? ' ' + cls : '');
  line.textContent = msg;
  log.appendChild(line);
  log.scrollTop = log.scrollHeight;
}

// ─────────────────────────────────────────────────────────────
// PDF → CANVAS → OCR
// ─────────────────────────────────────────────────────────────
async function ocrFirstPage(fileHandle) {
  await initPdfJs(); // no-op if already loaded
  const file = await fileHandle.getFile();
  const arrayBuffer = await file.arrayBuffer();

  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) }).promise;
  const page = await pdf.getPage(1);

  // 3× scale ≈ 216 DPI effective — higher scale improves small-text OCR
  const viewport = page.getViewport({ scale: 3.0 });
  const canvas = document.createElement('canvas');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  const ctx = canvas.getContext('2d');
  await page.render({ canvasContext: ctx, viewport }).promise;
  await pdf.destroy();

  const { data: { text } } = await state.tesseractWorker.recognize(canvas);
  return text;
}

// ─────────────────────────────────────────────────────────────
// OCR TEXT PARSING
// ─────────────────────────────────────────────────────────────
function parseOcrText(text) {
  const result = {
    learnerName: '',
    atsId: '',
    courseCode: '',
    criteriaType: '',
    criteriaStart: '',
    criteriaEnd: '',
  };

  // ── Learner Name ──
  // Pattern: "Learner Name [name] Course Code" (all on one line from OCR)
  const nameLineMatch = text.match(/Learner\s*Name\s+([A-Za-z][^\n|]{2,40}?)(?:\s{2,}|\s+Course|\s+\|)/i);
  if (nameLineMatch) {
    result.learnerName = nameLineMatch[1].trim();
  } else {
    // Fallback: any text after "Learner Name" up to end of line
    const m = text.match(/Learner\s*Name[:\s]+([A-Za-z][^\n]{2,40})/i);
    if (m) result.learnerName = m[1].trim().replace(/\s*Course.*$/i, '').trim();
  }

  // ── ATS ID ──
  // Pattern: "ATS ID [digits+noise] Subject" — extract digit sequence
  const atsLineMatch = text.match(/ATS\s*ID\s+([A-Za-z0-9\s|!@#]{2,30}?)(?:\s{2,}|\s+Subject|\s*\n)/i);
  if (atsLineMatch) {
    result.atsId = atsLineMatch[1].replace(/\D/g, ''); // keep digits only
  } else {
    const m = text.match(/ATS\s*ID[:\s]+([^\n]{1,20})/i);
    if (m) result.atsId = m[1].replace(/\D/g, '');
  }

  // ── Course Code ── e.g. ENE61, EEE61, ETE61
  result.courseCode = (() => {
    // 1. "Course Code" label — tolerates any separator and optional space between letters and digits
    let m = text.match(/[Cc]ourse\s*[Cc]ode[\s|:\-]*([A-Za-z]{2,4})\s*(\d{1,4})/i);
    if (m) return (m[1] + m[2]).toUpperCase();

    // 2. Just "Code" label (handles OCR garbling "Course")
    m = text.match(/\bcode[\s|:\-]+([A-Za-z]{2,4})\s*(\d{1,4})/i);
    if (m) return (m[1] + m[2]).toUpperCase();

    // 3. Scan all letter+digit codes, exclude GC/PC/PS/IT/IS/AS/AD/ID patterns
    const excluded = /^(GC|PC|PS|IT|IS|AS|AD|ID|NO|PO)/i;
    const allCodes = [...text.matchAll(/\b([A-Za-z]{2,4})\s*(\d{2,3})\b/g)]
      .map(x => (x[1] + x[2]).toUpperCase())
      .filter(c => !excluded.test(c));
    if (allCodes.length > 0) return allCodes[0];

    return '';
  })();

  // ── Criteria type and range ── e.g. "GC 5.1-5.3" or "PC 2.1-2.4"
  const rangeMatch = text.match(/\b(GC|PC)\s*(\d+\.\d+)\s*[-–]\s*(\d+\.\d+)/i);
  if (rangeMatch) {
    result.criteriaType  = rangeMatch[1].toUpperCase();
    result.criteriaStart = rangeMatch[2];
    result.criteriaEnd   = rangeMatch[3];
  } else {
    // Single criterion or range written without dash
    const singleMatch = text.match(/\b(GC|PC)\s*(\d+\.\d+)/i);
    if (singleMatch) {
      result.criteriaType  = singleMatch[1].toUpperCase();
      result.criteriaStart = singleMatch[2];
      result.criteriaEnd   = singleMatch[2];
      // Try to find the last criterion number in the criteria table
      const allMatches = [...text.matchAll(/\b(?:GC|PC)\s*(\d+\.\d+)/gi)];
      if (allMatches.length > 1) {
        result.criteriaEnd = allMatches[allMatches.length - 1][1];
      }
    }
  }

  return result;
}

// ─────────────────────────────────────────────────────────────
// FILENAME HELPERS
// ─────────────────────────────────────────────────────────────
function extractNameFromFilename(filename) {
  // "Milestone 1 - Aysha  Alrashdi.pdf" → "Aysha Alrashdi"
  const m = filename.replace(/\.pdf$/i, '').match(/[-–]\s+(.+)$/);
  return m ? m[1].trim().replace(/\s+/g, ' ') : '';
}

function buildFilename(courseCode, criteriaType, start, end, atsId) {
  // Confirmed state values always take priority over per-file OCR results
  const code = (state.courseCode  || courseCode  || 'UNKNOWN').toUpperCase();
  const type = (state.criteriaType  || criteriaType  || 'GC').toUpperCase();
  const s    =  state.criteriaStart || start || '?';
  const e    =  state.criteriaEnd   || end   || s;
  const id   = atsId || 'UNKNOWN';
  if (s === e) return `${code}_${type}${s}_P_${id}.pdf`;
  return `${code}_${type}${s}_${e}_P_${id}.pdf`;
}

// ─────────────────────────────────────────────────────────────
// MATCHING ALGORITHM
// ─────────────────────────────────────────────────────────────
function normalise(str) {
  return (str || '')
    .toLowerCase()
    .replace(/\b(al|el)\s+(?=[a-z])/g, '$1') // "Al Ktebi" → "Alktebi", "Al Rashdi" → "Alrashdi"
    .replace(/[^a-z\s]/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function tokenize(str) {
  return normalise(str).split(' ').filter(t => t.length > 1);
}

function jaccardTokens(a, b) {
  if (!a || !b) return 0;
  const sa = new Set(tokenize(a));
  const sb = new Set(tokenize(b));
  if (sa.size === 0 || sb.size === 0) return 0;
  const inter = [...sa].filter(t => sb.has(t)).length;
  const union  = new Set([...sa, ...sb]).size;
  return inter / union;
}

// Character trigram Jaccard — tolerates OCR misreadings
function trigramJaccard(a, b) {
  if (!a || !b) return 0;
  const norm = s => s.toLowerCase().replace(/[^a-z]/g, '');
  const tg   = s => {
    const n = norm(s);
    const g = new Set();
    for (let i = 0; i <= n.length - 3; i++) g.add(n.slice(i, i + 3));
    return g;
  };
  const ga = tg(a);
  const gb = tg(b);
  if (ga.size === 0 || gb.size === 0) return 0;
  const inter = [...ga].filter(g => gb.has(g)).length;
  const union  = new Set([...ga, ...gb]).size;
  return inter / union;
}

// Extract first name token
function getFirstName(nameStr) {
  const tokens = tokenize(nameStr);
  return tokens.length > 0 ? tokens[0] : '';
}

// Extract second name token (middle name)
function getSecondName(nameStr) {
  const tokens = tokenize(nameStr);
  return tokens.length > 1 ? tokens[1] : '';
}

// Extract last name — merges Arabic "Al/El" prefix with the following word
// e.g. "Rauda Hamad Al Ktebi" → "alktebi"  (not just "ktebi")
function getLastName(nameStr) {
  const tokens = tokenize(nameStr); // all tokens, no length filter
  if (tokens.length === 0) return '';
  const last = tokens[tokens.length - 1];
  if (tokens.length >= 2) {
    const prev = tokens[tokens.length - 2];
    if (/^(al|el)$/i.test(prev)) {
      return prev + last; // "al" + "ktebi" → "alktebi"
    }
  }
  // Skip very short tokens when looking for the family name
  const long = tokens.filter(t => t.length >= 4);
  return long.length > 0 ? long[long.length - 1] : last;
}

// Consonant skeleton — strips vowels so Arabic transliteration variants match
// "Alketbi" → "lktb",  "Alktebi" → "lktb"  ← identical ✓
// "Rouda"   → "rd",    "Rauda"   → "rd"     ← identical ✓
function consonantSkeleton(str) {
  return (str || '').toLowerCase().replace(/[aeiou\s]/g, '');
}

// Last-name similarity: consonant skeleton first, trigram as fallback
function lastNameSimilarity(a, b) {
  if (!a || !b) return 0;
  const ca = consonantSkeleton(a);
  const cb = consonantSkeleton(b);
  if (ca && cb && ca === cb) return 1.0;           // exact consonant match
  const skelScore = trigramJaccard(ca, cb);        // trigram on skeletons
  const rawScore  = trigramJaccard(a, b);          // trigram on full strings
  return Math.max(skelScore, rawScore);
}

// Digit sequence overlap — for ATS ID tiebreaker
function digitOverlap(ocrDigits, excelId) {
  if (!ocrDigits || ocrDigits.length < 2) return 0;
  const ed = (excelId || '').replace(/\D/g, '');
  if (!ed) return 0;
  let hits = 0, total = 0;
  for (let len = 2; len <= Math.min(ocrDigits.length, 5); len++) {
    for (let i = 0; i <= ocrDigits.length - len; i++) {
      const sub = ocrDigits.slice(i, i + len);
      if (ed.includes(sub)) hits++;
      total++;
    }
  }
  return total ? hits / total : 0;
}

function scoreAgainstExcel(nameStr, students) {
  const ocrLastName = getLastName(nameStr); // last meaningful token e.g. "alrashdi"

  return students.map(s => {
    const jaccard       = jaccardTokens(nameStr, s.name);
    const trigram       = trigramJaccard(nameStr, s.name);
    const excelLastName = getLastName(s.name);

    // Last-name similarity — uses consonant skeleton so "Alktebi" ≡ "Alketbi"
    const lastNameScore = ocrLastName.length >= 3
      ? lastNameSimilarity(ocrLastName, excelLastName)
      : 0;

    // Weights: 50% exact token match + 30% full-name trigram + 20% last-name trigram
    const score = jaccard * 0.5 + trigram * 0.3 + lastNameScore * 0.2;

    return { student: s, score, jaccard, trigram, lastNameScore };
  }).sort((a, b) => b.score - a.score);
}

function matchStudent(ocrName, filenameHint, ocrId, students) {

  // ── Stage 0: First name + Last name direct check ──────────────────────────
  // Use BOTH the OCR name and the filename hint as name sources.
  // If first name AND last name both match a student clearly:
  //   → Check ATS ID: if match → high confidence (matched)
  //   → If ID doesn't match → check second/middle name to choose between candidates
  const nameSources = [ocrName, filenameHint].filter(n => n && n.length > 2);

  for (const src of nameSources) {
    const srcFirst = getFirstName(src);
    const srcLast  = getLastName(src);
    if (!srcFirst || !srcLast || srcFirst === srcLast) continue; // need both

    const firstLastMatches = students.filter(s => {
      const exFirst = getFirstName(s.name);
      const exLast  = getLastName(s.name);
      // Use consonant skeletons for both first and last names
      const firstOk = trigramJaccard(srcFirst, exFirst) >= 0.45
                   || consonantSkeleton(srcFirst) === consonantSkeleton(exFirst);
      const lastOk  = lastNameSimilarity(srcLast, exLast) >= 0.5;
      return firstOk && lastOk;
    });

    if (firstLastMatches.length === 1) {
      // Unique first+last match — check ATS ID for confidence level
      const student = firstLastMatches[0];
      const idScore = digitOverlap(ocrId, student.atsId);
      const conf    = idScore >= 0.3 ? 0.92 : 0.72;
      return { status: 'matched', student, confidence: conf, candidates: [] };
    }

    if (firstLastMatches.length > 1) {
      // Multiple students share first+last name pattern — use ATS ID to pick
      const withId = firstLastMatches.map(s => ({
        student: s,
        score: 0.7,
        idScore: digitOverlap(ocrId, s.atsId),
      })).sort((a, b) => b.idScore - a.idScore);

      const top  = withId[0];
      const next = withId[1];

      if (top.idScore > 0.25 && top.idScore - next.idScore > 0.15) {
        return { status: 'matched', student: top.student, confidence: 0.85, candidates: withId };
      }

      // ID didn't resolve it — check second (middle) name
      const srcSecond = getSecondName(src);
      if (srcSecond) {
        const withSecond = firstLastMatches.map(s => ({
          student: s,
          score: 0.7,
          secondScore: trigramJaccard(srcSecond, getSecondName(s.name)),
        })).sort((a, b) => b.secondScore - a.secondScore);

        const s0 = withSecond[0], s1 = withSecond[1];
        if (s0.secondScore >= 0.5 && s0.secondScore - s1.secondScore > 0.2) {
          return { status: 'matched', student: s0.student, confidence: 0.80, candidates: withSecond };
        }
      }

      // Still ambiguous — flag for manual
      return {
        status: 'manual',
        student: null,
        confidence: 0.65,
        candidates: firstLastMatches.map(s => ({ student: s, score: 0.65 })),
        bestGuess: firstLastMatches[0],
      };
    }
  }

  // ── Score using OCR name ──
  const ocrScored = ocrName.length > 2 ? scoreAgainstExcel(ocrName, students) : [];

  // ── Score using filename hint ──
  const fnScored  = filenameHint.length > 2 ? scoreAgainstExcel(filenameHint, students) : [];

  // ── Merge: best score for each student from either source ──
  const merged = students.map(s => {
    const o = ocrScored.find(r => r.student === s);
    const f = fnScored.find(r => r.student === s);
    const score = Math.max(o?.score || 0, f?.score || 0);
    const ocrScore = o?.score || 0;
    const fnScore  = f?.score || 0;
    return { student: s, score, ocrScore, fnScore };
  }).sort((a, b) => b.score - a.score);

  const top    = merged[0];
  const second = merged[1];
  const gap    = second ? top.score - second.score : 1;

  // ── Stage 1: Clear winner ──
  if (top.score >= 0.5 && gap >= 0.2) {
    return { status: 'matched', student: top.student, confidence: top.score, candidates: [] };
  }

  // ── Too low to work with ──
  if (top.score < 0.2) {
    return { status: 'unreadable', student: null, confidence: 0, candidates: merged.slice(0, 3) };
  }

  // ── Stage 2: Narrowing via more tokens (already in merged score) — check top candidates ──
  const candidates = merged.filter(r => r.score >= 0.2);

  if (candidates.length === 1) {
    const conf = candidates[0].score;
    const status = conf >= 0.5 ? 'matched' : 'low-confidence';
    return { status, student: candidates[0].student, confidence: conf, candidates };
  }

  // ── Stage 2b: Last name tiebreaker ──
  // When two candidates are close in score, compare the last name from
  // both the OCR text and the filename hint against each candidate's family name.
  if (candidates.length > 1) {
    const ocrLastName = getLastName(ocrName);
    const fnLastName  = getLastName(filenameHint);

    const withLastName = candidates.map(c => {
      const excelLast    = getLastName(c.student.name);
      const ocrLnScore   = ocrLastName.length >= 3 ? lastNameSimilarity(ocrLastName, excelLast) : 0;
      const fnLnScore    = fnLastName.length  >= 3 ? lastNameSimilarity(fnLastName,  excelLast) : 0;
      const lnScore      = Math.max(ocrLnScore, fnLnScore);
      return { ...c, lnScore, combinedScore: c.score * 0.7 + lnScore * 0.3 };
    }).sort((a, b) => b.combinedScore - a.combinedScore);

    const lnTop  = withLastName[0];
    const lnNext = withLastName[1];
    const lnGap  = lnNext ? lnTop.combinedScore - lnNext.combinedScore : 1;

    if (lnTop.lnScore >= 0.4 && lnGap >= 0.15) {
      const conf = lnTop.combinedScore;
      return {
        status: conf >= 0.45 ? 'matched' : 'low-confidence',
        student: lnTop.student,
        confidence: conf,
        candidates: withLastName,
      };
    }
  }

  // ── Stage 3: ATS ID digit tiebreaker ──
  if (ocrId && ocrId.length >= 3) {
    const withId = candidates.map(c => ({
      ...c,
      idScore: digitOverlap(ocrId, c.student.atsId),
    })).sort((a, b) => {
      const diff = b.idScore - a.idScore;
      return diff !== 0 ? diff : b.score - a.score;
    });

    const idTop  = withId[0];
    const idNext = withId[1];
    const idGap  = idNext ? idTop.idScore - idNext.idScore : 1;

    if (idTop.idScore > 0.25 && idGap >= 0.15) {
      const conf = (idTop.score + idTop.idScore) / 2;
      return { status: conf >= 0.4 ? 'matched' : 'low-confidence', student: idTop.student, confidence: conf, candidates: withId };
    }
  }

  // ── Stage 4: Manual assignment needed ──
  const bestConf = top.score;
  return {
    status: 'manual',
    student: null,
    confidence: bestConf,
    candidates: candidates.slice(0, 5),
    bestGuess: top.student,
  };
}

// ─────────────────────────────────────────────────────────────
// STEP 3 — RESULTS TABLE
// ─────────────────────────────────────────────────────────────
function buildResultsTable() {
  const tbody = document.getElementById('results-tbody');
  tbody.innerHTML = '';

  let nMatched = 0, nLow = 0, nManual = 0, nUnread = 0;
  state.results.forEach(r => {
    const s = r.match.status;
    if (s === 'matched') nMatched++;
    else if (s === 'low-confidence') nLow++;
    else nManual++;
    if (s === 'unreadable' || s === 'error') nUnread++;
  });

  // Summary pills
  const pills = document.getElementById('summary-pills');
  pills.innerHTML = `
    <span class="pill pill-green">✅ ${nMatched} matched</span>
    ${nLow > 0 ? `<span class="pill pill-yellow">⚠️ ${nLow} low-confidence</span>` : ''}
    ${(nManual - nUnread) > 0 ? `<span class="pill pill-red">⚠️ ${nManual - nUnread} need manual</span>` : ''}
    ${nUnread > 0 ? `<span class="pill pill-red">❌ ${nUnread} unreadable</span>` : ''}
  `;

  // Update apply button label
  const nSelected = state.results.filter(r => r.selected).length;
  document.getElementById('apply-label').textContent = `Apply Renaming (${nSelected} files)`;

  state.results.forEach((r, idx) => {
    const tr = document.createElement('tr');
    const st = r.match.status;
    tr.className = st === 'matched' ? 'row-matched'
      : st === 'low-confidence'    ? 'row-low'
      : 'row-manual';

    // Confidence bar colour
    const conf = r.match.confidence || 0;
    const confPct = Math.round(conf * 100);
    const confColor = conf >= 0.6 ? '#22c55e' : conf >= 0.4 ? '#f59e0b' : '#ef4444';

    // Badge
    const badgeHtml = st === 'matched'
      ? '<span class="badge badge-green">✅ Matched</span>'
      : st === 'low-confidence'
      ? '<span class="badge badge-yellow">⚠ Low confidence</span>'
      : st === 'unreadable' || st === 'error'
      ? '<span class="badge badge-red">❌ Unreadable</span>'
      : '<span class="badge badge-red">⚠ Manual needed</span>';

    // Student cell — dropdown for non-matched, text for matched
    let studentCell = '';
    const currentStudent = r.match.student || r.match.bestGuess;
    if (st === 'matched' && currentStudent) {
      studentCell = `<div class="student-cell truncate" title="${escHtml(currentStudent.name)}">${escHtml(currentStudent.name)}</div>`;
    } else {
      const opts = (r.match.candidates || state.students.slice(0, 5))
        .map(c => {
          const s = c.student || c;
          return `<option value="${escHtml(s.atsId)}" data-name="${escHtml(s.name)}">${escHtml(s.name)}</option>`;
        }).join('');
      const selectedVal = currentStudent ? `value="${escHtml(currentStudent.atsId)}"` : '';
      studentCell = `<select class="assign-select" data-idx="${idx}">
        <option value="">— Select student —</option>
        ${opts}
      </select>`;
    }

    // New filename cell
    const nameCell = r.newName
      ? `<span class="new-name-cell">${escHtml(r.newName)}</span>`
      : '<span style="color:var(--gray-400);font-size:12px;">— pending assignment —</span>';

    // OCR name + filename hint
    const ocrDisplay = r.parsed.learnerName
      ? escHtml(r.parsed.learnerName)
      : `<em style="color:var(--gray-400)">${escHtml(r.filenameHint || 'not read')}</em>`;

    tr.innerHTML = `
      <td><input type="checkbox" class="cb row-cb" data-idx="${idx}" ${r.selected ? 'checked' : ''}></td>
      <td><div class="filename-cell truncate" title="${escHtml(r.pdf.name)}">${escHtml(r.pdf.name)}</div></td>
      <td><div class="ocr-name-cell" title="${escHtml(r.parsed.learnerName || '')}">${ocrDisplay}</div></td>
      <td>${studentCell}</td>
      <td style="font-size:12px;color:var(--gray-600);font-family:monospace;">${escHtml((r.match.student || currentStudent)?.atsId || '—')}</td>
      <td>
        <div class="conf-bar">
          <span class="conf-track"><span class="conf-fill" style="width:${confPct}%;background:${confColor}"></span></span>
          <span style="font-size:11px;color:var(--gray-500)">${confPct}%</span>
        </div>
      </td>
      <td>${nameCell}</td>
      <td>${badgeHtml}</td>
    `;
    tbody.appendChild(tr);
  });

  // Pre-fill dropdowns for low-confidence best guesses
  state.results.forEach((r, idx) => {
    if ((r.match.status === 'low-confidence' || r.match.status === 'manual') && r.match.bestGuess) {
      const sel = document.querySelector(`select[data-idx="${idx}"]`);
      if (sel) {
        sel.value = r.match.bestGuess.atsId;
        // Trigger the assignment
        const ev = new Event('change');
        sel.dispatchEvent(ev);
      }
    }
  });
}

function toggleRow(idx, checked) {
  state.results[idx].selected = checked;
  const n = state.results.filter(r => r.selected).length;
  document.getElementById('apply-label').textContent = `Apply Renaming (${n} files)`;
  // Sync select-all checkbox
  document.getElementById('select-all').checked = state.results.every(r => r.selected || !r.newName);
};

function toggleAll(checked) {
  state.results.forEach((r, idx) => {
    if (r.newName) {
      r.selected = checked;
      const cb = document.querySelector(`.row-cb[data-idx="${idx}"]`);
      if (cb) cb.checked = checked;
    }
  });
  const n = state.results.filter(r => r.selected).length;
  document.getElementById('apply-label').textContent = `Apply Renaming (${n} files)`;
};

function handleManualAssign(select) {
  const idx  = parseInt(select.dataset.idx, 10);
  const atsId = select.value;
  const r    = state.results[idx];

  if (!atsId) {
    r.match.student = null;
    r.newName = null;
    r.selected = false;
    // Update cells
    const row = select.closest('tr');
    row.querySelector('td:nth-child(5)').textContent = '—';
    row.querySelector('td:nth-child(7)').innerHTML = '<span style="color:var(--gray-400);font-size:12px;">— pending assignment —</span>';
    row.querySelector('.row-cb').checked = false;
    toggleRow(idx, false);
    return;
  }

  const student = state.students.find(s => s.atsId === atsId);
  if (!student) return;

  r.match.student = student;
  r.match.status  = 'low-confidence';
  const p = r.parsed;
  r.newName = buildFilename(p.courseCode, p.criteriaType, p.criteriaStart, p.criteriaEnd, student.atsId);
  r.selected = true;

  // Update the row cells without full re-render
  const row = select.closest('tr');
  row.querySelector('td:nth-child(5)').textContent = student.atsId;
  row.querySelector('td:nth-child(7)').innerHTML = `<span class="new-name-cell">${escHtml(r.newName)}</span>`;
  const cb = row.querySelector('.row-cb');
  cb.checked = true;
  toggleRow(idx, true);
};

// ─────────────────────────────────────────────────────────────
// STEP 4 — APPLY RENAMING
// ─────────────────────────────────────────────────────────────
async function applyRenaming() {
  const toProcess = state.results.filter(r => r.selected && r.newName);
  if (toProcess.length === 0) { alert('No files selected for renaming.'); return; }

  const applyBtn = document.getElementById('apply-btn');
  applyBtn.disabled = true;
  applyBtn.innerHTML = '⏳ Copying files…';

  try {
    // Request write permission at apply time (not at folder selection)
    const perm = await state.dirHandle.requestPermission({ mode: 'readwrite' });
    if (perm !== 'granted') {
      alert('Write permission is required to create the output subfolder. Please allow it.');
      applyBtn.disabled = false;
      applyBtn.innerHTML = '⚠ Retry — write permission needed';
      return;
    }

    const ts  = new Date().toISOString().replace(/[-T:.Z]/g, '').slice(0, 14);
    const outDir = await state.dirHandle.getDirectoryHandle(`Renamed_${ts}`, { create: true });
    state.finalReport = [];
    let done = 0, errors = 0;

    for (const r of toProcess) {
      try {
        const srcHandle  = await state.dirHandle.getFileHandle(r.pdf.name);
        const file       = await srcHandle.getFile();
        const destHandle = await outDir.getFileHandle(r.newName, { create: true });
        const writable   = await destHandle.createWritable();
        await writable.write(file);
        await writable.close();
        state.finalReport.push({ ...r, applyStatus: 'ok' });
        done++;
      } catch (e) {
        state.finalReport.push({ ...r, applyStatus: 'error', applyError: e.message });
        errors++;
      }
    }

    // Also add unselected/unmatched to report
    state.results.filter(r => !r.selected || !r.newName).forEach(r => {
      state.finalReport.push({ ...r, applyStatus: 'skipped' });
    });

    // Show done step
    const folderName = `Renamed_${ts}`;
    document.getElementById('done-title').textContent = `${done} file${done !== 1 ? 's' : ''} renamed!`;
    document.getElementById('done-sub').textContent = `Copied to subfolder "${folderName}" inside your PDF folder.`;
    document.getElementById('done-stats').innerHTML = `
      <span class="pill pill-green">✅ ${done} renamed</span>
      ${errors > 0 ? `<span class="pill pill-red">❌ ${errors} errors</span>` : ''}
      <span class="pill pill-blue">⏭ ${state.results.length - toProcess.length} skipped</span>
    `;
    showStep(4);

  } catch (e) {
    applyBtn.disabled = false;
    applyBtn.innerHTML = '⚠ Error — retry';
    alert('Error during renaming: ' + e.message);
    console.error(e);
  }
};

// ─────────────────────────────────────────────────────────────
// CSV REPORT
// ─────────────────────────────────────────────────────────────
function downloadCSV() {
  const report = state.finalReport.length ? state.finalReport : state.results;
  const rows = [
    ['Original Filename', 'OCR Name Read', 'Filename Hint', 'Matched Student', 'ATS ID', 'Course Code', 'Criteria', 'New Filename', 'Status'],
    ...report.map(r => [
      r.pdf.name,
      r.parsed?.learnerName || '',
      r.filenameHint || '',
      r.match?.student?.name || '',
      r.match?.student?.atsId || '',
      r.parsed?.courseCode || '',
      r.parsed?.criteriaType && r.parsed?.criteriaStart
        ? `${r.parsed.criteriaType}${r.parsed.criteriaStart}–${r.parsed.criteriaEnd}` : '',
      r.newName || '',
      r.applyStatus || r.match?.status || '',
    ]),
  ];

  const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = `renaming_report_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
};

// ─────────────────────────────────────────────────────────────
// RESTART
// ─────────────────────────────────────────────────────────────
function restart() {
  state.dirHandle     = null;
  state.pdfFiles      = [];
  state.students      = [];
  state.results       = [];
  state.courseCode    = '';
  state.criteriaType  = '';
  state.criteriaStart = '';
  state.criteriaEnd   = '';
  state.finalReport   = [];

  // Reset step 1 UI
  ['zone-pdf', 'zone-excel'].forEach(id => {
    const z = document.getElementById(id);
    z.classList.remove('loaded');
  });
  document.getElementById('zone-pdf').addEventListener('click', pickFolder);
  document.getElementById('zone-excel').addEventListener('click', pickExcel);
  document.getElementById('pdf-tag').classList.add('hidden');
  document.getElementById('excel-tag').classList.add('hidden');
  document.getElementById('zone-pdf').querySelector('h3').textContent = 'Student PDF Folder';
  document.getElementById('zone-pdf').querySelector('p').textContent = 'Click to select the folder containing all student assignment PDFs';
  document.getElementById('zone-excel').querySelector('h3').textContent = 'Gradebook (Excel)';
  document.getElementById('zone-excel').querySelector('p').textContent = 'Click to select the .xlsx file with student names and ATS IDs';
  document.getElementById('ocr-log').innerHTML = '';
  document.getElementById('start-btn').disabled = true;
  document.getElementById('start-btn').innerHTML = `<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg> Start OCR Processing`;
  showStep(1);
};

// ─────────────────────────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────────────────────────
function escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Surface any unhandled module errors to the page console
window.addEventListener('error', e => console.error('app.js error:', e.message, e.filename, e.lineno));
window.addEventListener('unhandledrejection', e => console.error('app.js unhandled promise:', e.reason));

// ─────────────────────────────────────────────────────────────
// WIRE UP ALL EVENT LISTENERS (replaces all removed inline handlers)
// ─────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  // Step 1 — file loaders
  document.getElementById('zone-pdf').addEventListener('click', pickFolder);
  document.getElementById('zone-excel').addEventListener('click', pickExcel);
  document.getElementById('start-btn').addEventListener('click', startProcessing);

  // Step 3 — select-all checkbox
  document.getElementById('select-all').addEventListener('change', e => toggleAll(e.target.checked));

  // Step 3 — apply + CSV buttons
  document.getElementById('apply-btn').addEventListener('click', applyRenaming);
  document.getElementById('csv-btn').addEventListener('click', downloadCSV);

  // Step 4 — done-page buttons
  document.getElementById('done-csv-btn').addEventListener('click', downloadCSV);
  document.getElementById('restart-btn').addEventListener('click', restart);

  // Step 3 — event delegation for dynamic table rows (checkboxes + dropdowns)
  document.getElementById('results-tbody').addEventListener('change', e => {
    if (e.target.classList.contains('row-cb')) {
      const idx = parseInt(e.target.dataset.idx, 10);
      toggleRow(idx, e.target.checked);
    }
    if (e.target.classList.contains('assign-select')) {
      handleManualAssign(e.target);
    }
  });
});
