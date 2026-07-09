const api = window.postaMassivaApi;

const state = {
  inputPath: '',
  outputDir: '',
  sheetInfoByName: new Map(),
  lastResult: null
};

const acceptedExtensions = ['.xlsx', '.xls', '.xlsm'];

const els = {
  inputDropzone: document.getElementById('inputDropzone'),
  inputPath: document.getElementById('inputPath'),
  outputDir: document.getElementById('outputDir'),
  sheetName: document.getElementById('sheetName'),
  capColumn: document.getElementById('capColumn'),
  boxCapacity: document.getElementById('boxCapacity'),
  postingDate: document.getElementById('postingDate'),
  productHomologated: document.getElementById('productHomologated'),
  boxTare: document.getElementById('boxTare'),
  shipper: document.getElementById('shipper'),
  center: document.getElementById('center'),
  shippingCode: document.getElementById('shippingCode'),
  productCode: document.getElementById('productCode'),
  defaultWeight: document.getElementById('defaultWeight'),
  minBancaleWeight: document.getElementById('minBancaleWeight'),
  maxBancaleWeight: document.getElementById('maxBancaleWeight'),
  status: document.getElementById('status'),
  results: document.getElementById('results'),
  btnInput: document.getElementById('btnInput'),
  btnOutput: document.getElementById('btnOutput'),
  btnProcess: document.getElementById('btnProcess'),
  btnOpenSorted: document.getElementById('btnOpenSorted'),
  btnOpenChiudi: document.getElementById('btnOpenChiudi'),
  btnOpenBancali: document.getElementById('btnOpenBancali')
};

function dirname(filePath) {
  const idx = Math.max(filePath.lastIndexOf('\\'), filePath.lastIndexOf('/'));
  if (idx < 0) {
    return '';
  }
  return filePath.slice(0, idx);
}

function normalizeHeader(input) {
  return String(input || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function detectHeader(headers, patterns) {
  for (const header of headers) {
    const normalized = normalizeHeader(header);
    if (patterns.some((pattern) => normalized.includes(pattern))) {
      return header;
    }
  }
  return '';
}

function setStatus(message, type = '') {
  els.status.className = `status ${type}`.trim();
  els.status.textContent = message;
}

function isExcelFilePath(filePath) {
  const normalized = String(filePath || '').toLowerCase();
  return acceptedExtensions.some((ext) => normalized.endsWith(ext));
}

function getDroppedFilePath(event) {
  const files = event.dataTransfer?.files;
  if (!files || files.length === 0) {
    return '';
  }

  const droppedFile = files[0];
  if (droppedFile && typeof droppedFile.path === 'string') {
    return droppedFile.path;
  }

  return '';
}

function fillSelect(selectElement, values, includeEmpty = false) {
  selectElement.innerHTML = '';

  if (includeEmpty) {
    const emptyOption = document.createElement('option');
    emptyOption.value = '';
    emptyOption.textContent = '-- non impostata --';
    selectElement.appendChild(emptyOption);
  }

  for (const value of values) {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    selectElement.appendChild(option);
  }
}

function updateColumnSelectors() {
  const selectedSheet = els.sheetName.value;
  const sheetInfo = state.sheetInfoByName.get(selectedSheet);
  const headers = sheetInfo ? sheetInfo.headers : [];

  fillSelect(els.capColumn, headers, false);

  const capDetected = detectHeader(headers, ['cap', 'capdest', 'cap_dest']);

  if (capDetected) {
    els.capColumn.value = capDetected;
  }
}

function renderResult(result) {
  const warnings = Array.isArray(result.ctaWarnings) ? result.ctaWarnings : [];
  const warningsBlock =
    warnings.length > 0
      ? `\nAvvisi conformità CTA:\n- ${warnings.join('\n- ')}`
      : '\nAvvisi conformità CTA: nessuno';

  els.results.textContent = [
    `File input: ${result.inputPath}`,
    `Foglio elaborato: ${result.sheetName}`,
    `Record totali: ${result.totalRows}`,
    `Record classificati: ${result.matchedRows}`,
    `Record non classificati: ${result.unmatchedRows}`,
    `Plichi generati: ${result.totalPlichi}`,
    `Plichi MIX: ${result.mixPlichi ?? 0}`,
    `Limite massimo record per plico: ${result.boxCapacity}`,
    `Bancali generati: ${result.totalBancali ?? 0}`,
    `Bancali MIX: ${result.mixBancali ?? 0}`,
    `Peso minimo bancale: ${result.minBancaleWeightKg ?? ''} kg`,
    `Peso massimo bancale: ${result.maxBancaleWeightKg ?? ''} kg`,
    `Regole CAP caricate da manuale: ${result.rulesCount}`,
    warningsBlock,
    `\nOutput record ordinati: ${result.sortedOutputPath}`,
    `Output plichi: ${result.chiudiOutputPath}`,
    `Output bancali: ${result.bancaliOutputPath}`
  ].join('\n');
}

async function loadWorkbookInfo(filePath) {
  const info = await api.getExcelInfo(filePath);
  state.sheetInfoByName.clear();

  for (const sheet of info.sheets) {
    state.sheetInfoByName.set(sheet.name, sheet);
  }

  fillSelect(els.sheetName, info.sheetNames, false);
  if (info.sheetNames.length > 0) {
    els.sheetName.value = info.sheetNames[0];
  }

  updateColumnSelectors();
}

async function applyInputFile(filePath) {
  if (!filePath) {
    throw new Error('Percorso file non valido.');
  }

  if (!isExcelFilePath(filePath)) {
    throw new Error('Formato non supportato. Usa un file .xlsx, .xls o .xlsm.');
  }

  state.inputPath = filePath;
  els.inputPath.value = filePath;

  if (!els.outputDir.value) {
    const defaultOutputDir = dirname(filePath);
    els.outputDir.value = defaultOutputDir;
    state.outputDir = defaultOutputDir;
  }

  await loadWorkbookInfo(filePath);
  setStatus('File caricato. Verifica CAP e avvia l\'elaborazione.', 'ok');
}

els.sheetName.addEventListener('change', updateColumnSelectors);

els.btnInput.addEventListener('click', async () => {
  try {
    const result = await api.selectInputFile();
    if (result.canceled) {
      return;
    }

    await applyInputFile(result.filePath);
  } catch (error) {
    setStatus(`Errore caricamento file: ${error.message}`, 'error');
  }
});

let dragDepth = 0;

window.addEventListener('dragover', (event) => {
  event.preventDefault();
});

window.addEventListener('drop', (event) => {
  event.preventDefault();
});

els.inputDropzone.addEventListener('dragenter', (event) => {
  event.preventDefault();
  event.stopPropagation();
  dragDepth += 1;
  els.inputDropzone.classList.add('drop-active');
});

els.inputDropzone.addEventListener('dragover', (event) => {
  event.preventDefault();
  event.stopPropagation();
});

els.inputDropzone.addEventListener('dragleave', (event) => {
  event.preventDefault();
  event.stopPropagation();
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) {
    els.inputDropzone.classList.remove('drop-active');
  }
});

els.inputDropzone.addEventListener('drop', async (event) => {
  event.preventDefault();
  event.stopPropagation();
  dragDepth = 0;
  els.inputDropzone.classList.remove('drop-active');

  try {
    const droppedPath = getDroppedFilePath(event);
    await applyInputFile(droppedPath);
  } catch (error) {
    setStatus(`Errore caricamento file: ${error.message}`, 'error');
  }
});

els.btnOutput.addEventListener('click', async () => {
  try {
    const result = await api.selectOutputDirectory();
    if (result.canceled) {
      return;
    }

    state.outputDir = result.directoryPath;
    els.outputDir.value = result.directoryPath;
  } catch (error) {
    setStatus(`Errore selezione cartella: ${error.message}`, 'error');
  }
});

els.btnProcess.addEventListener('click', async () => {
  if (!state.inputPath) {
    setStatus('Seleziona prima un file Excel input.', 'error');
    return;
  }

  if (!els.capColumn.value) {
    setStatus('Seleziona la colonna CAP.', 'error');
    return;
  }

  if (!String(els.defaultWeight.value || '').trim()) {
    setStatus('Inserisci il peso unitario di default in grammi.', 'error');
    return;
  }

  if (!String(els.maxBancaleWeight.value || '').trim()) {
    setStatus('Inserisci il peso massimo bancale in kg.', 'error');
    return;
  }

  if (!String(els.minBancaleWeight.value || '').trim()) {
    setStatus('Inserisci il peso minimo bancale in kg.', 'error');
    return;
  }

  if (Number(els.minBancaleWeight.value) > Number(els.maxBancaleWeight.value)) {
    setStatus('Il peso minimo bancale non puo superare il massimo.', 'error');
    return;
  }

  const options = {
    inputPath: state.inputPath,
    outputDir: els.outputDir.value || dirname(state.inputPath),
    sheetName: els.sheetName.value,
    capColumn: els.capColumn.value,
    boxCapacity: els.boxCapacity.value,
    postingDate: els.postingDate.value,
    productHomologated: els.productHomologated.value,
    boxTare: els.boxTare.value,
    shipper: els.shipper.value,
    center: els.center.value,
    shippingCode: els.shippingCode.value,
    productCode: els.productCode.value,
    defaultWeight: els.defaultWeight.value,
    minBancaleWeight: els.minBancaleWeight.value,
    maxBancaleWeight: els.maxBancaleWeight.value
  };

  try {
    els.btnProcess.disabled = true;
    setStatus('Elaborazione in corso...', '');

    const result = await api.processWorkbook(options);
    state.lastResult = result;

    renderResult(result);
    setStatus('Elaborazione completata.', 'ok');

    els.btnOpenSorted.disabled = false;
    els.btnOpenChiudi.disabled = false;
    els.btnOpenBancali.disabled = false;
  } catch (error) {
    setStatus(`Errore elaborazione: ${error.message}`, 'error');
  } finally {
    els.btnProcess.disabled = false;
  }
});

els.btnOpenSorted.addEventListener('click', async () => {
  if (!state.lastResult) {
    return;
  }
  await api.openPath(state.lastResult.sortedOutputPath);
});

els.btnOpenChiudi.addEventListener('click', async () => {
  if (!state.lastResult) {
    return;
  }
  await api.openPath(state.lastResult.chiudiOutputPath);
});

els.btnOpenBancali.addEventListener('click', async () => {
  if (!state.lastResult) {
    return;
  }
  await api.openPath(state.lastResult.bancaliOutputPath);
});
