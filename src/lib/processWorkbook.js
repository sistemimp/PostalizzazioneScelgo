const path = require('node:path');
const fs = require('node:fs');
const XLSX = require('xlsx');
const {
  parseBaciniRulesFromPdf,
  findRuleByCap,
  normalizeCap,
  BACINI_ORDER
} = require('./baciniRules');

const TARIFF_ORDER = { AM: 0, CP: 1, EU: 2, MIX: 3 };
const MIX_BACINO = 'MIX BACINI DESTINAZIONI VARIE';
const MIX_PROVINCIA = 'BACINI VARI';
const MIX_TARIFF = 'MIX';
const MIX_CLOSURE_TYPE = 'CHIUDI PLICO MIX BACINI DESTINAZIONI VARIE';
const DEFAULT_MAX_INVII_PER_PLICO = 500;
const DEFAULT_PLICO_TARE_GRAMS = 203;
const MIN_INVII_PER_BACINO_GROUP = 10;
const DEFAULT_BANCALI_MIN_WEIGHT_KG = 200;
const DEFAULT_BANCALI_MAX_WEIGHT_KG = 500;
const BACINO_INDEX = new Map(BACINI_ORDER.map((name, index) => [name, index]));

function toNumber(value, fallbackValue = 0) {
  if (value === null || value === undefined || value === '') {
    return fallbackValue;
  }

  const normalized = String(value).replace(',', '.');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : fallbackValue;
}

function parseWeightGrams(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  const normalized = String(value).trim().replace(',', '.').replace(/[^0-9.-]/g, '');
  if (!normalized) {
    return null;
  }

  const parsed = Number(normalized);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return null;
  }

  return parsed;
}

function formatPostingDate(dateInput) {
  const now = new Date();

  if (!dateInput) {
    const dd = String(now.getDate()).padStart(2, '0');
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const yyyy = String(now.getFullYear());
    return `${dd}/${mm}/${yyyy}`;
  }

  const match = String(dateInput).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return String(dateInput);
  }

  return `${match[3]}/${match[2]}/${match[1]}`;
}

function makeGroupKey(parts) {
  return parts.map((part) => String(part ?? '')).join('||');
}

function appendProcessingNote(currentNote, message) {
  if (!currentNote) {
    return message;
  }

  return `${currentNote} | ${message}`;
}

function isMixBacinoValue(value) {
  const normalized = String(value || '').trim().toUpperCase();
  return normalized === MIX_BACINO || normalized === 'MIX BACINI';
}

function getBacinoSortOrder(bacino) {
  return BACINO_INDEX.has(bacino) ? BACINO_INDEX.get(bacino) : 999;
}

function getTariffSortOrder(tariff) {
  return TARIFF_ORDER[tariff] ?? 999;
}

function createSorter() {
  return (left, right) => {
    const leftBacino = getBacinoSortOrder(left.BACINO_DESTINAZIONE);
    const rightBacino = getBacinoSortOrder(right.BACINO_DESTINAZIONE);
    if (leftBacino !== rightBacino) {
      return leftBacino - rightBacino;
    }

    const leftTariff = getTariffSortOrder(left.DESTINAZIONE_TARIFFARIA);
    const rightTariff = getTariffSortOrder(right.DESTINAZIONE_TARIFFARIA);
    if (leftTariff !== rightTariff) {
      return leftTariff - rightTariff;
    }

    const leftProvince = String(left.PROVINCIA_BACINO || '').localeCompare(String(right.PROVINCIA_BACINO || ''), 'it');
    if (leftProvince !== 0) {
      return leftProvince;
    }

    const leftCap = Number(left.CAP_NORMALIZZATO || 99999);
    const rightCap = Number(right.CAP_NORMALIZZATO || 99999);
    if (leftCap !== rightCap) {
      return leftCap - rightCap;
    }

    return left.__sourceIndex - right.__sourceIndex;
  };
}

function applyMinimumInviiForBacinoGrouping(rows, minInvii) {
  const threshold = Math.max(1, Math.floor(toNumber(minInvii, MIN_INVII_PER_BACINO_GROUP)));
  const groupedCounts = new Map();

  for (const row of rows) {
    const bacino = String(row.BACINO_DESTINAZIONE || '').trim().toUpperCase();
    const tariff = String(row.DESTINAZIONE_TARIFFARIA || '').trim().toUpperCase();
    const provincia = String(row.PROVINCIA_BACINO || '').trim().toUpperCase();

    if (!bacino || !tariff || isMixBacinoValue(bacino) || tariff === MIX_TARIFF) {
      continue;
    }

    const key = makeGroupKey([bacino, tariff, provincia]);
    groupedCounts.set(key, (groupedCounts.get(key) || 0) + 1);
  }

  let reroutedRows = 0;

  for (const row of rows) {
    const bacino = String(row.BACINO_DESTINAZIONE || '').trim().toUpperCase();
    const tariff = String(row.DESTINAZIONE_TARIFFARIA || '').trim().toUpperCase();
    const provincia = String(row.PROVINCIA_BACINO || '').trim().toUpperCase();

    if (!bacino || !tariff || isMixBacinoValue(bacino) || tariff === MIX_TARIFF) {
      continue;
    }

    const key = makeGroupKey([bacino, tariff, provincia]);
    const groupCount = groupedCounts.get(key) || 0;
    if (groupCount >= threshold) {
      continue;
    }

    row.BACINO_DESTINAZIONE = MIX_BACINO;
    row.DESTINAZIONE_TARIFFARIA = MIX_TARIFF;
    row.PROVINCIA_BACINO = MIX_PROVINCIA;
    row.NOTE_ELABORAZIONE = appendProcessingNote(
      row.NOTE_ELABORAZIONE,
      `Riassegnato a MIX: gruppo ${bacino}/${tariff}/${provincia || 'N.D.'} con ${groupCount} invii (< ${threshold}).`
    );
    reroutedRows += 1;
  }

  return {
    threshold,
    reroutedRows
  };
}

function buildFallbackItemsByPlico(totalItems, plicoCapacity) {
  const totalPlichi = Math.ceil(totalItems / plicoCapacity);
  const itemsByPlico = [];

  for (let plico = 1; plico <= totalPlichi; plico += 1) {
    const remaining = totalItems - (plico - 1) * plicoCapacity;
    itemsByPlico.push(Math.min(plicoCapacity, remaining));
  }

  return itemsByPlico;
}

function getProvinciaBacinoForPlico(provinceNames, tariffs, bacini, fallbackValue = '') {
  const normalizedProvinces = (Array.isArray(provinceNames) ? provinceNames : [])
    .map((value) => String(value || '').trim().toUpperCase())
    .filter(Boolean);
  const normalizedBacini = (Array.isArray(bacini) ? bacini : [])
    .map((value) => String(value || '').trim().toUpperCase())
    .filter(Boolean);
  const uniqueBacini = new Set(normalizedBacini);
  const plicoBacino = uniqueBacini.size === 1 ? [...uniqueBacini][0] : '';
  const hasDefinedBacino = plicoBacino && !isMixBacinoValue(plicoBacino);
  const nonHomogeneousValue = hasDefinedBacino ? `${plicoBacino} MIX PROVINCIE` : MIX_PROVINCIA;

  if (normalizedProvinces.length === 0) {
    return hasDefinedBacino ? nonHomogeneousValue : fallbackValue || '';
  }

  const uniqueProvinces = new Set(normalizedProvinces);
  if (uniqueProvinces.size !== 1) {
    return nonHomogeneousValue;
  }

  const provinceName = [...uniqueProvinces][0];
  const normalizedTariffs = (Array.isArray(tariffs) ? tariffs : [])
    .map((value) => String(value || '').trim().toUpperCase())
    .filter(Boolean);
  if (normalizedTariffs.length === 0) {
    return hasDefinedBacino ? nonHomogeneousValue : fallbackValue || provinceName;
  }

  const allCityDestinations = normalizedTariffs.every((value) => value === 'AM' || value === 'CP');
  if (allCityDestinations) {
    return `${provinceName} CITTA'`;
  }

  const allProvinceDestinations = normalizedTariffs.every((value) => value === 'EU');
  if (allProvinceDestinations) {
    return `${provinceName} PROVINCIA`;
  }

  return hasDefinedBacino ? nonHomogeneousValue : fallbackValue || provinceName;
}

function mapRowToPlicoGroup(row) {
  const bacino = row.BACINO_DESTINAZIONE || MIX_BACINO;
  const tariff = row.DESTINAZIONE_TARIFFARIA || MIX_TARIFF;
  const isMix = isMixBacinoValue(bacino) || tariff === MIX_TARIFF;

  return {
    bacino,
    tariff,
    provinciaBacino: row.PROVINCIA_BACINO || MIX_PROVINCIA,
    closureType: isMix ? MIX_CLOSURE_TYPE : 'CHIUDI PLICO'
  };
}

function buildPlicoGroups(sortedRows) {
  const groups = [];
  const groupsByKey = new Map();

  for (const [rowIndex, row] of sortedRows.entries()) {
    const mappedGroup = mapRowToPlicoGroup(row);
    const key = makeGroupKey([mappedGroup.bacino, mappedGroup.tariff, mappedGroup.provinciaBacino, mappedGroup.closureType]);

    if (!groupsByKey.has(key)) {
      groupsByKey.set(key, {
        ...mappedGroup,
        count: 0,
        rowIndexes: [],
        rowWeights: [],
        rowTariffs: [],
        rowProvinceNames: [],
        rowBacini: []
      });
      groups.push(groupsByKey.get(key));
    }

    const group = groupsByKey.get(key);
    group.count += 1;
    group.rowIndexes.push(rowIndex);
    group.rowWeights.push(Math.max(0, toNumber(row.PESO_UNITARIO_GR, 0)));
    group.rowTariffs.push(mappedGroup.tariff || MIX_TARIFF);
    group.rowProvinceNames.push(String(row.PROVINCIA_BACINO || '').trim().toUpperCase());
    group.rowBacini.push(mappedGroup.bacino || MIX_BACINO);
  }

  return groups;
}

function buildPlichiRows(sortedRows, options) {
  const plicoCapacity = Math.max(1, Math.floor(toNumber(options.boxCapacity, DEFAULT_MAX_INVII_PER_PLICO)));
  const plicoTareGrams = Math.max(0, toNumber(options.boxTare, DEFAULT_PLICO_TARE_GRAMS));
  const postingDate = formatPostingDate(options.postingDate);
  const rowProgressivi = Array(sortedRows.length).fill('');
  const rowAssignmentOrder = Array(sortedRows.length).fill(Number.MAX_SAFE_INTEGER);
  const warnings = [];
  const plichiRows = [];
  let nextProgressivo = 1;
  let nextRowSequence = 1;
  let mixPlichi = 0;

  for (const group of buildPlicoGroups(sortedRows)) {
    const itemsByPlico = buildFallbackItemsByPlico(group.count, plicoCapacity);
    const totalPlichi = itemsByPlico.length;
    let rowOffset = 0;

    for (let plicoNumber = 1; plicoNumber <= totalPlichi; plicoNumber += 1) {
      const itemsInThisPlico = itemsByPlico[plicoNumber - 1];
      const firstLetterIndex = nextRowSequence;
      let netWeightGrams = 0;
      const plicoTariffs = [];
      const plicoProvinceNames = [];
      const plicoBacini = [];
      const assignedRowIndexes = [];

      for (let i = 0; i < itemsInThisPlico && rowOffset < group.rowIndexes.length; i += 1) {
        const rowIndex = group.rowIndexes[rowOffset];
        const rowWeight = toNumber(group.rowWeights[rowOffset], 0);
        const rowTariff = String(group.rowTariffs[rowOffset] || '').trim().toUpperCase();
        const rowProvinceName = String(group.rowProvinceNames[rowOffset] || '').trim().toUpperCase();
        const rowBacino = String(group.rowBacini[rowOffset] || '').trim().toUpperCase();

        netWeightGrams += rowWeight;
        if (rowTariff) {
          plicoTariffs.push(rowTariff);
        }
        if (rowProvinceName) {
          plicoProvinceNames.push(rowProvinceName);
        }
        if (rowBacino) {
          plicoBacini.push(rowBacino);
        }

        rowProgressivi[rowIndex] = nextProgressivo;
        rowAssignmentOrder[rowIndex] = nextRowSequence;
        nextRowSequence += 1;
        assignedRowIndexes.push(rowIndex);
        rowOffset += 1;
      }

      const lastLetterIndex = nextRowSequence - 1;
      const lettersRange = lastLetterIndex >= firstLetterIndex ? `da ${firstLetterIndex} a ${lastLetterIndex}` : '';
      const provinciaBacino = getProvinciaBacinoForPlico(
        plicoProvinceNames,
        plicoTariffs,
        plicoBacini,
        group.provinciaBacino || ''
      );
      const netWeightKg = netWeightGrams / 1000;
      const grossWeightKg = (netWeightGrams + plicoTareGrams) / 1000;
      const isMix = group.tariff === MIX_TARIFF || isMixBacinoValue(group.bacino) || group.closureType === MIX_CLOSURE_TYPE;

      if (isMix) {
        mixPlichi += 1;
      }

      plichiRows.push({
        __progressivoPlico: nextProgressivo,
        __assignedRowIndexes: assignedRowIndexes,
        __grossWeightKg: grossWeightKg,
        __netWeightKg: netWeightKg,
        __groupOrderKey: makeGroupKey([group.bacino, group.tariff, group.provinciaBacino]),
        __schemaProvince: group.provinciaBacino,
        'Data di Impostazione (gg/mm/aaaa)': postingDate,
        'Progressivo chiudi plico': nextProgressivo,
        'Azienda Speditrice (ID SAP + Nome Cliente)': options.shipper || '',
        'Centro di Impostazione (AGE + Nome Centro)': options.center || '',
        'Tipo chiusura': group.closureType,
        'CODICE SPEDIZIONE': options.shippingCode || '',
        'Prodotto (Cod. Materiale)': options.productCode || '',
        'Prodotto Omologato (SI/NO)': options.productHomologated || 'SI',
        'Destinazione tariffaria': group.tariff,
        'Peso unitario (in grammi)': toNumber(options.defaultWeight, 0),
        'Tara plico (grammi)': plicoTareGrams,
        'Numero invii totale plico': itemsInThisPlico,
        'Range lettere contenute': lettersRange,
        'Peso netto invii plico (kg)': Number(netWeightKg.toFixed(3)),
        'Peso totale plico (kg)': Number(grossWeightKg.toFixed(3)),
        'Provincia bacino': provinciaBacino,
        'BACINO DESTINAZIONE': group.bacino,
        'Numero plico nel gruppo': plicoNumber,
        'Totale plichi nel gruppo': totalPlichi,
        'Gruppo bancale': '',
        'Progressivo bancale': '',
        'Tipo bancale': ''
      });

      nextProgressivo += 1;
    }
  }

  return {
    rows: plichiRows,
    rowProgressivi,
    rowAssignmentOrder,
    warnings,
    plicoCapacity,
    mixPlichi
  };
}

function normalizeProvinceForSchema(value) {
  return String(value || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[’']/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function getBancaleSchemaGroup(bacino, provinciaBacino) {
  const normalizedBacino = String(bacino || '').trim().toUpperCase();
  const normalizedProvince = normalizeProvinceForSchema(provinciaBacino);

  if (normalizedBacino === 'ANCONA') {
    if (['ANCONA', 'PESARO', 'PESARO E URBINO', 'URBINO', 'MACERATA'].includes(normalizedProvince)) {
      return 'ANCONA -> ANCONA + PESARO + MACERATA';
    }

    if (['ASCOLI PICENO', 'FERMO'].includes(normalizedProvince)) {
      return 'ANCONA -> ASCOLI PICENO + FERMO';
    }
  }

  if (normalizedBacino === 'PESCARA') {
    if (normalizedProvince === 'TERAMO') {
      return 'PESCARA -> TERAMO';
    }

    if (normalizedProvince === 'PESCARA') {
      return 'PESCARA -> PESCARA';
    }

    if (['CHIETI', 'L AQUILA', 'AQUILA', 'CAMPOBASSO', 'ISERNIA'].includes(normalizedProvince)) {
      return "PESCARA -> CHIETI + L'AQUILA + CAMPOBASSO + ISERNIA";
    }
  }

  return MIX_BACINO;
}

function splitPlichiIntoBancali(plichi, maxWeightKg, warnings, groupLabel) {
  const bancali = [];
  let current = [];
  let currentWeight = 0;

  for (const plico of plichi) {
    const plicoWeightKg = Math.max(0, toNumber(plico.__grossWeightKg, 0));

    if (plicoWeightKg > maxWeightKg) {
      warnings.push(
        `Plico ${plico['Progressivo chiudi plico']} nel gruppo ${groupLabel} supera il massimo bancale (${plicoWeightKg.toFixed(3)} kg > ${maxWeightKg} kg).`
      );
    }

    if (current.length > 0 && currentWeight + plicoWeightKg > maxWeightKg) {
      bancali.push({ plichi: current, totalWeightKg: currentWeight });
      current = [];
      currentWeight = 0;
    }

    current.push(plico);
    currentWeight += plicoWeightKg;
  }

  if (current.length > 0) {
    bancali.push({ plichi: current, totalWeightKg: currentWeight });
  }

  return bancali;
}

function summarizeDistinct(values) {
  return [...new Set(values.map((value) => String(value || '').trim()).filter(Boolean))].join(' + ');
}

function toExportBancaleRow(progressivoBancale, groupLabel, typeLabel, bancale, minWeightKg, maxWeightKg, reason = '') {
  const firstPlico = bancale.plichi[0];
  const lastPlico = bancale.plichi[bancale.plichi.length - 1];
  const bacini = summarizeDistinct(bancale.plichi.map((plico) => plico['BACINO DESTINAZIONE']));
  const province = summarizeDistinct(bancale.plichi.map((plico) => plico['Provincia bacino']));
  const tariffs = summarizeDistinct(bancale.plichi.map((plico) => plico['Destinazione tariffaria']));

  return {
    'Progressivo bancale': progressivoBancale,
    'Tipo bancale': typeLabel,
    'Gruppo bancale': groupLabel,
    'Numero plichi': bancale.plichi.length,
    'Da progressivo plico': firstPlico ? firstPlico['Progressivo chiudi plico'] : '',
    'A progressivo plico': lastPlico ? lastPlico['Progressivo chiudi plico'] : '',
    'Peso totale bancale (kg)': Number(bancale.totalWeightKg.toFixed(3)),
    'Peso minimo bancale (kg)': minWeightKg,
    'Peso massimo bancale (kg)': maxWeightKg,
    'Bacini contenuti': bacini,
    'Province contenute': province,
    'Destinazioni tariffarie': tariffs,
    Note: reason
  };
}

function buildBancaliRows(plichiRows, minBancaleWeightKg, maxBancaleWeightKg) {
  const warnings = [];
  const buckets = new Map();

  for (const plico of plichiRows) {
    const groupLabel = getBancaleSchemaGroup(plico['BACINO DESTINAZIONE'], plico.__schemaProvince);
    if (!buckets.has(groupLabel)) {
      buckets.set(groupLabel, []);
    }
    buckets.get(groupLabel).push(plico);
  }

  const mixPool = [];
  const exportRows = [];
  let progressivoBancale = 1;
  let mixBancali = 0;

  for (const [groupLabel, plichi] of buckets.entries()) {
    if (groupLabel === MIX_BACINO) {
      mixPool.push(...plichi);
      continue;
    }

    const split = splitPlichiIntoBancali(plichi, maxBancaleWeightKg, warnings, groupLabel);
    for (const bancale of split) {
      if (bancale.totalWeightKg < minBancaleWeightKg) {
        mixPool.push(...bancale.plichi);
        warnings.push(
          `Gruppo bancale ${groupLabel} sotto minimo ${minBancaleWeightKg} kg (${bancale.totalWeightKg.toFixed(3)} kg). Riassegnato a MIX.`
        );
        continue;
      }

      const exportRow = toExportBancaleRow(progressivoBancale, groupLabel, 'Bancale', bancale, minBancaleWeightKg, maxBancaleWeightKg);
      exportRows.push(exportRow);
      for (const plico of bancale.plichi) {
        plico['Gruppo bancale'] = groupLabel;
        plico['Progressivo bancale'] = progressivoBancale;
        plico['Tipo bancale'] = 'Bancale';
      }
      progressivoBancale += 1;
    }
  }

  if (mixPool.length > 0) {
    const splitMix = splitPlichiIntoBancali(mixPool, maxBancaleWeightKg, warnings, MIX_BACINO);
    for (const bancale of splitMix) {
      const underMinimum = bancale.totalWeightKg < minBancaleWeightKg;
      if (underMinimum) {
        warnings.push(
          `Bancale MIX sotto minimo ${minBancaleWeightKg} kg (${bancale.totalWeightKg.toFixed(3)} kg). Generato comunque per esaurire i plichi residui.`
        );
      }

      const typeLabel = 'Bancale MIX';
      const note = underMinimum ? 'Residuo MIX sotto soglia minima.' : '';
      const exportRow = toExportBancaleRow(progressivoBancale, MIX_BACINO, typeLabel, bancale, minBancaleWeightKg, maxBancaleWeightKg, note);
      exportRows.push(exportRow);
      mixBancali += 1;

      for (const plico of bancale.plichi) {
        plico['Gruppo bancale'] = MIX_BACINO;
        plico['Progressivo bancale'] = progressivoBancale;
        plico['Tipo bancale'] = typeLabel;
      }

      progressivoBancale += 1;
    }
  }

  for (const plico of plichiRows) {
    delete plico.__progressivoPlico;
    delete plico.__assignedRowIndexes;
    delete plico.__grossWeightKg;
    delete plico.__netWeightKg;
    delete plico.__groupOrderKey;
    delete plico.__schemaProvince;
  }

  return {
    rows: exportRows,
    warnings,
    mixBancali
  };
}

function addSheet(workbook, rows, name) {
  const worksheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, worksheet, name);
}

function createTimestamp() {
  const now = new Date();
  const y = String(now.getFullYear());
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');
  return `${y}${m}${d}_${hh}${mm}${ss}`;
}

async function processWorkbook(options) {
  if (!options || !options.inputPath) {
    throw new Error('Percorso file Excel non valido.');
  }

  if (!options.capColumn) {
    throw new Error('Seleziona la colonna CAP.');
  }

  const defaultWeightGrams = parseWeightGrams(options.defaultWeight);
  if (!defaultWeightGrams) {
    throw new Error('Peso unitario di default non valido.');
  }

  const minBancaleWeightKg = Math.max(1, toNumber(options.minBancaleWeight, DEFAULT_BANCALI_MIN_WEIGHT_KG));
  const maxBancaleWeightKg = Math.max(
    minBancaleWeightKg,
    toNumber(options.maxBancaleWeight, DEFAULT_BANCALI_MAX_WEIGHT_KG)
  );

  const inputPath = path.resolve(options.inputPath);
  const outputDir = options.outputDir ? path.resolve(options.outputDir) : path.dirname(inputPath);
  fs.mkdirSync(outputDir, { recursive: true });

  const docsDir = options.docsDir ? path.resolve(options.docsDir) : path.resolve(process.cwd(), 'docs');
  const baciniPdfPath = path.join(docsDir, 'postamassiva-elenco-bacini-destinazione.pdf');
  const rules = await parseBaciniRulesFromPdf(baciniPdfPath);

  const workbook = XLSX.readFile(inputPath, { cellDates: true });
  const sheetName = options.sheetName || workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    throw new Error(`Foglio "${sheetName}" non trovato.`);
  }

  const inputRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

  const enrichedRows = inputRows.map((row, index) => {
    const capValue = row[options.capColumn];
    const capNormalized = normalizeCap(capValue);
    const rule = findRuleByCap(capNormalized, rules);

    let status = 'OK';
    let note = '';

    if (!rule) {
      status = 'CAP_NON_CLASSIFICATO';
      note = 'CAP non trovato nel manuale bacini.';
    }

    return {
      ...row,
      CAP_ORIGINALE: capValue,
      CAP_NORMALIZZATO: capNormalized || '',
      BACINO_DESTINAZIONE: rule ? rule.bacino : MIX_BACINO,
      DESTINAZIONE_TARIFFARIA: rule ? rule.tariff : MIX_TARIFF,
      PROVINCIA_BACINO: rule ? rule.province : MIX_PROVINCIA,
      PESO_UNITARIO_GR: defaultWeightGrams,
      STATO_ELABORAZIONE: status,
      NOTE_ELABORAZIONE: note,
      __sourceIndex: index
    };
  });

  const groupingResult = applyMinimumInviiForBacinoGrouping(enrichedRows, MIN_INVII_PER_BACINO_GROUP);
  enrichedRows.sort(createSorter());

  const rowsForExportBase = enrichedRows.map((row) => {
    const { __sourceIndex, ...cleaned } = row;
    return cleaned;
  });

  const plichiBuild = buildPlichiRows(rowsForExportBase, options);
  const bancaliBuild = buildBancaliRows(plichiBuild.rows, minBancaleWeightKg, maxBancaleWeightKg);
  const rowsForExport = rowsForExportBase
    .map((row, idx) => ({
      ...row,
      PROGRESSIVO_CHIUDI_PLICO: plichiBuild.rowProgressivi[idx] || '',
      __ASSIGNMENT_ORDER: Number.isFinite(plichiBuild.rowAssignmentOrder[idx])
        ? plichiBuild.rowAssignmentOrder[idx]
        : Number.MAX_SAFE_INTEGER,
      __ORIGINAL_ORDER: idx
    }))
    .sort((left, right) => {
      if (left.__ASSIGNMENT_ORDER !== right.__ASSIGNMENT_ORDER) {
        return left.__ASSIGNMENT_ORDER - right.__ASSIGNMENT_ORDER;
      }
      return left.__ORIGINAL_ORDER - right.__ORIGINAL_ORDER;
    })
    .map((row, idx) => {
      const { __ASSIGNMENT_ORDER, __ORIGINAL_ORDER, ...cleaned } = row;
      return {
        ORDINE_ELABORAZIONE: idx + 1,
        ...cleaned
      };
    });

  const unmatchedRows = rowsForExport.filter((row) => row.STATO_ELABORAZIONE !== 'OK');
  const plichiRows = plichiBuild.rows;
  const bancaliRows = bancaliBuild.rows;
  const warnings = [...plichiBuild.warnings, ...bancaliBuild.warnings];
  const timestamp = createTimestamp();
  const baseName = path.parse(inputPath).name;

  const sortedOutputPath = path.join(outputDir, `${baseName}_record_ordinati_${timestamp}.xlsx`);
  const chiudiOutputPath = path.join(outputDir, `${baseName}_plichi_${timestamp}.xlsx`);
  const bancaliOutputPath = path.join(outputDir, `${baseName}_bancali_${timestamp}.xlsx`);

  const sortedWorkbook = XLSX.utils.book_new();
  addSheet(sortedWorkbook, rowsForExport, 'Record_Ordinati');
  if (unmatchedRows.length > 0) {
    addSheet(sortedWorkbook, unmatchedRows, 'Scarti_Elaborazione');
  }
  XLSX.writeFile(sortedWorkbook, sortedOutputPath);

  const plichiWorkbook = XLSX.utils.book_new();
  addSheet(plichiWorkbook, plichiRows, 'Plichi');
  XLSX.writeFile(plichiWorkbook, chiudiOutputPath);

  const bancaliWorkbook = XLSX.utils.book_new();
  addSheet(bancaliWorkbook, bancaliRows, 'Bancali');
  XLSX.writeFile(bancaliWorkbook, bancaliOutputPath);

  return {
    inputPath,
    outputDir,
    sheetName,
    sortedOutputPath,
    chiudiOutputPath,
    bancaliOutputPath,
    totalRows: rowsForExport.length,
    matchedRows: rowsForExport.length - unmatchedRows.length,
    unmatchedRows: unmatchedRows.length,
    totalPlichi: plichiRows.length,
    mixPlichi: plichiBuild.mixPlichi,
    totalBancali: bancaliRows.length,
    mixBancali: bancaliBuild.mixBancali,
    minBancaleWeightKg,
    maxBancaleWeightKg,
    ctaWarnings: warnings,
    boxCapacity: plichiBuild.plicoCapacity,
    rulesCount: rules.ranges.length,
    minInviiPerBacinoGroup: groupingResult.threshold,
    reroutedToMixRows: groupingResult.reroutedRows
  };
}

module.exports = {
  processWorkbook
};
