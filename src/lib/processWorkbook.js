const path = require('node:path');
const fs = require('node:fs');
const XLSX = require('xlsx');
const {
  parseBaciniRulesFromPdf,
  findRuleByCap,
  normalizeCap,
  BACINI_ORDER
} = require('./baciniRules');

const TARIFF_ORDER = { AM: 0, CP: 1, EU: 2 };
const CTA_WEIGHT_TOLERANCE = 0.02;
const MIX_BACINO = 'MIX BACINI DESTINAZIONI VARIE';
const MIX_CLOSURE_TYPE = 'CHIUDI PLICO MIX BACINI DESTINAZIONI VARIE';
const MIX_PROVINCIA = 'BACINI VARI';
const MIX_TARIFF = 'MIX';
const MIN_INVII_PER_BACINO_GROUP = 10;
const DEFAULT_BOX_HEIGHT_CM = 46;
const DEFAULT_BOX_REACTION = 1;
const DEFAULT_MAX_INVII_PER_BOX = 500;
const ENVELOPE_PROFILE_BY_ELEMENTS = new Map([
  [1, { weightGrams: 9.7, thicknessMm: 0.04 }],
  [2, { weightGrams: 14.7, thicknessMm: 0.09 }],
  [3, { weightGrams: 19.7, thicknessMm: 0.13 }],
  [4, { weightGrams: 24.7, thicknessMm: 0.17 }],
  [5, { weightGrams: 29.7, thicknessMm: 0.21 }],
  [6, { weightGrams: 34.7, thicknessMm: 0.26 }],
  [7, { weightGrams: 39.7, thicknessMm: 0.3 }]
]);
const BACINO_INDEX = new Map(BACINI_ORDER.map((name, index) => [name, index]));
const CTA_FORMAT_RULES = {
  P: {
    minStandardKg: 4,
    maxStandardKg: 7,
    minPartialKg: 2.5
  },
  M: {
    minStandardKg: 6,
    maxStandardKg: 9,
    minPartialKg: 4
  }
};

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

function parseElementsCount(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  const normalized = String(value).trim().replace(',', '.').replace(/[^0-9.-]/g, '');
  if (!normalized) {
    return null;
  }

  const parsed = Number(normalized);
  if (!Number.isFinite(parsed) || !Number.isInteger(parsed)) {
    return null;
  }

  return ENVELOPE_PROFILE_BY_ELEMENTS.has(parsed) ? parsed : null;
}

function normalizeFormat(value) {
  const normalized = String(value || '').trim().toUpperCase();
  if (!normalized) {
    return '';
  }

  if (normalized.startsWith('P')) {
    return 'P';
  }

  if (normalized.startsWith('M')) {
    return 'M';
  }

  return '';
}

function formatPostingDate(dateInput) {
  const now = new Date();

  if (!dateInput) {
    const dd = String(now.getDate()).padStart(2, '0');
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const yyyy = String(now.getFullYear());
    return `${dd}/${mm}/${yyyy}`;
  }

  // input da HTML date: YYYY-MM-DD
  const match = String(dateInput).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return String(dateInput);
  }

  return `${match[3]}/${match[2]}/${match[1]}`;
}

function formatProvinciaBacino(provincia, tariff, bacino = '') {
  const normalized = String(provincia || '').trim().toUpperCase();
  if (!normalized) {
    return '';
  }

  if (String(bacino || '').toUpperCase() === MIX_BACINO || tariff === 'MIX') {
    return normalized;
  }

  if (tariff === 'AM' || tariff === 'CP') {
    return `${normalized} BACINO`;
  }

  if (tariff === 'EU') {
    return `${normalized} PROVINCIA`;
  }

  return normalized;
}

function getProvinciaBacinoForBox(provinceNames, tariffs, bacini, fallbackValue = '') {
  const normalizedProvinces = (Array.isArray(provinceNames) ? provinceNames : [])
    .map((value) => String(value || '').trim().toUpperCase())
    .filter(Boolean);
  const normalizedBacini = (Array.isArray(bacini) ? bacini : [])
    .map((value) => String(value || '').trim().toUpperCase())
    .filter(Boolean);
  const uniqueBacini = new Set(normalizedBacini);
  const boxBacino = uniqueBacini.size === 1 ? [...uniqueBacini][0] : '';
  const hasDefinedBacino = boxBacino && !isMixBacinoValue(boxBacino);
  const nonHomogeneousValue = hasDefinedBacino ? `${boxBacino} MIX PROVINCIE` : MIX_PROVINCIA;

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

function getRowFormatAndWeight(row, options) {
  const rawFormat = (options.formatColumn ? row[options.formatColumn] : '') || options.defaultFormat || '';
  const rawWeight = (options.weightColumn ? row[options.weightColumn] : '') || options.defaultWeight || '';
  const rawElements = options.elementsColumn ? row[options.elementsColumn] : '';

  const format = normalizeFormat(rawFormat);
  const elementsCount = parseElementsCount(rawElements);
  const elementsProfile = elementsCount ? ENVELOPE_PROFILE_BY_ELEMENTS.get(elementsCount) : null;
  const weightGrams = elementsProfile ? elementsProfile.weightGrams : parseWeightGrams(rawWeight);
  const thicknessMm = elementsProfile ? elementsProfile.thicknessMm : null;

  return {
    rawFormat,
    format,
    rawWeight,
    weightGrams,
    rawElements,
    elementsCount,
    thicknessMm
  };
}

function getCtaLimits(format) {
  const formatRule = CTA_FORMAT_RULES[format];
  if (!formatRule) {
    return null;
  }

  return {
    minStandardKg: formatRule.minStandardKg * (1 - CTA_WEIGHT_TOLERANCE),
    maxStandardKg: formatRule.maxStandardKg * (1 + CTA_WEIGHT_TOLERANCE),
    minPartialKg: formatRule.minPartialKg
  };
}

function calculatePieceThresholds(unitWeightGrams, thicknessMm, limits, boxCapacity, boxHeightCm, boxReaction) {
  const boxHeightMm = Math.max(1, toNumber(boxHeightCm, DEFAULT_BOX_HEIGHT_CM) * 10);
  const reaction = Math.max(0.01, toNumber(boxReaction, DEFAULT_BOX_REACTION));
  const maxByWeight = Math.floor((limits.maxStandardKg * 1000) / unitWeightGrams);
  const maxByPolicy = Math.max(1, Math.floor(boxCapacity));
  const maxByThickness =
    thicknessMm && Number.isFinite(thicknessMm) && thicknessMm > 0
      ? Math.max(1, Math.floor((boxHeightMm * reaction) / thicknessMm))
      : Number.POSITIVE_INFINITY;
  const maxPerBox = Math.max(1, Math.min(maxByPolicy, maxByWeight, maxByThickness));

  return {
    maxPerBox,
    maxByThickness
  };
}

function allocateItemsAcrossBoxes(totalItems, thresholds) {
  if (totalItems <= 0) {
    return [];
  }

  const maxPerBox = Math.max(1, Math.floor(toNumber(thresholds.maxPerBox, 1)));
  const totalBoxes = Math.ceil(totalItems / maxPerBox);
  const baseItems = Math.floor(totalItems / totalBoxes);
  let remainder = totalItems % totalBoxes;
  const boxes = Array(totalBoxes).fill(baseItems);

  for (let i = 0; i < boxes.length && remainder > 0; i += 1) {
    boxes[i] += 1;
    remainder -= 1;
  }

  return boxes;
}

function buildFallbackItemsByBox(totalItems, boxCapacity) {
  const totalBoxes = Math.ceil(totalItems / boxCapacity);
  const itemsByBox = [];
  for (let box = 1; box <= totalBoxes; box += 1) {
    const remaining = totalItems - (box - 1) * boxCapacity;
    itemsByBox.push(Math.min(boxCapacity, remaining));
  }
  return itemsByBox;
}

function getBacinoSortOrder(bacino) {
  return BACINO_INDEX.has(bacino) ? BACINO_INDEX.get(bacino) : 999;
}

function getTariffSortOrder(tariff) {
  return TARIFF_ORDER[tariff] ?? 999;
}

function getFormatSortOrder(format) {
  const normalized = normalizeFormat(format);
  if (normalized === 'P') {
    return 0;
  }
  if (normalized === 'M') {
    return 1;
  }
  return 999;
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

    const leftFormat = getFormatSortOrder(left.FORMATO_NORMALIZZATO);
    const rightFormat = getFormatSortOrder(right.FORMATO_NORMALIZZATO);
    if (leftFormat !== rightFormat) {
      return leftFormat - rightFormat;
    }

    const leftElements = Math.floor(toNumber(left.NUMERO_ELEMENTI_BUSTA, 0));
    const rightElements = Math.floor(toNumber(right.NUMERO_ELEMENTI_BUSTA, 0));
    if (leftElements !== rightElements) {
      return leftElements - rightElements;
    }

    const leftWeight = toNumber(left.PESO_UNITARIO_GR, 0);
    const rightWeight = toNumber(right.PESO_UNITARIO_GR, 0);
    if (leftWeight !== rightWeight) {
      return leftWeight - rightWeight;
    }

    const leftThickness = toNumber(left.SPESSORE_BUSTA_MM, 0);
    const rightThickness = toNumber(right.SPESSORE_BUSTA_MM, 0);
    if (leftThickness !== rightThickness) {
      return leftThickness - rightThickness;
    }

    const leftCap = Number(left.CAP_NORMALIZZATO || 99999);
    const rightCap = Number(right.CAP_NORMALIZZATO || 99999);
    if (leftCap !== rightCap) {
      return leftCap - rightCap;
    }

    return left.__sourceIndex - right.__sourceIndex;
  };
}

function makeGroupKey(parts) {
  return parts.map((part) => String(part ?? '')).join('||');
}

function isMixBacinoValue(value) {
  const normalized = String(value || '').trim().toUpperCase();
  return normalized === MIX_BACINO || normalized === 'MIX BACINI';
}

function appendProcessingNote(currentNote, message) {
  if (!currentNote) {
    return message;
  }

  return `${currentNote} | ${message}`;
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

function appendBoxes(chiudiRows, group, itemsByBox, options, stats, limits, assignmentState = null) {
  const boxTare = toNumber(options.boxTare, 0);
  const postingDate = formatPostingDate(options.postingDate);
  const totalBoxes = itemsByBox.length;
  const pesoUnitarioNumero = parseWeightGrams(group.pesoUnitario);
  const rowIndexes = Array.isArray(group.rowIndexes) ? group.rowIndexes : [];
  const rowWeights = Array.isArray(group.rowWeights) ? group.rowWeights : [];
  const rowTariffs = Array.isArray(group.rowTariffs) ? group.rowTariffs : [];
  const rowProvinceNames = Array.isArray(group.rowProvinceNames) ? group.rowProvinceNames : [];
  const rowBacini = Array.isArray(group.rowBacini) ? group.rowBacini : [];
  let rowOffset = 0;

  for (let boxNumber = 1; boxNumber <= totalBoxes; boxNumber += 1) {
    const itemsInThisBox = itemsByBox[boxNumber - 1];
    let boxWeightGrams = 0;
    const boxTariffs = [];
    const boxProvinceNames = [];
    const boxBacini = [];
    const firstLetterIndex = assignmentState ? assignmentState.nextRowSequence : '';

    for (let i = 0; i < itemsInThisBox && rowOffset < rowIndexes.length; i += 1) {
      const rowIndex = rowIndexes[rowOffset];
      const rowWeight = toNumber(rowWeights[rowOffset], 0);
      const rowTariff = String(rowTariffs[rowOffset] || '').trim().toUpperCase();
      const rowProvinceName = String(rowProvinceNames[rowOffset] || '').trim().toUpperCase();
      const rowBacino = String(rowBacini[rowOffset] || '').trim().toUpperCase();
      if (rowWeight > 0) {
        boxWeightGrams += rowWeight;
      }
      if (rowTariff) {
        boxTariffs.push(rowTariff);
      }
      if (rowProvinceName) {
        boxProvinceNames.push(rowProvinceName);
      }
      if (rowBacino) {
        boxBacini.push(rowBacino);
      }
      if (assignmentState) {
        const progressivo = assignmentState.nextProgressivo;
        assignmentState.rowProgressivi[rowIndex] = progressivo;
        assignmentState.rowAssignmentOrder[rowIndex] = assignmentState.nextRowSequence;
        assignmentState.nextRowSequence += 1;
      }
      rowOffset += 1;
    }
    const lastLetterIndex = assignmentState ? assignmentState.nextRowSequence - 1 : '';
    const lettersRange =
      assignmentState && Number.isFinite(firstLetterIndex) && Number.isFinite(lastLetterIndex)
      && lastLetterIndex >= firstLetterIndex
        ? `da ${firstLetterIndex} a ${lastLetterIndex}`
        : '';
    const provinciaBacino = getProvinciaBacinoForBox(
      boxProvinceNames,
      boxTariffs,
      boxBacini,
      group.provinciaBacino || ''
    );

    if (boxWeightGrams <= 0 && pesoUnitarioNumero) {
      boxWeightGrams = itemsInThisBox * pesoUnitarioNumero;
    }

    const netWeightKg = boxWeightGrams > 0 ? boxWeightGrams / 1000 : null;
    const grossWeightKg = netWeightKg !== null ? netWeightKg + boxTare / 1000 : null;
    const partialByWeight = limits && netWeightKg !== null ? netWeightKg < limits.minStandardKg : false;
    const progressivo = assignmentState ? assignmentState.nextProgressivo : '';

    if (assignmentState) {
      assignmentState.nextProgressivo += 1;
    }

    if (group.tariff === 'MIX' || isMixBacinoValue(group.bacino) || group.closureType === MIX_CLOSURE_TYPE) {
      stats.mixBoxes += 1;
    }
    if (partialByWeight) {
      stats.partialBoxes += 1;
    } else {
      stats.standardBoxes += 1;
    }

    chiudiRows.push({
      'Data di Impostazione (gg/mm/aaaa)': postingDate,
      'Progressivo chiudi scatola': progressivo,
      'Azienda Speditrice (ID SAP + Nome Cliente)': options.shipper || '',
      'Centro di Impostazione (AGE + Nome Centro)': options.center || '',
      'Tipo chiusura': group.closureType || 'CHIUDI SCATOLA',
      'CODICE SPEDIZIONE': options.shippingCode || '',
      'Prodotto (Cod. Materiale)': options.productCode || '',
      'Prodotto Omologato (SI/NO)': options.productHomologated || 'SI',
      Formato: group.formato,
      'Destinazione tariffaria': group.tariff,
      'Numero elementi per busta': group.elementsCount || '',
      'Peso unitario (in grammi)': group.pesoUnitario,
      'Spessore busta (mm)': group.thicknessMm || '',
      'Scatola riempita parzialmente (SI/NO)': partialByWeight ? 'SI' : 'NO',
      'Tara scatola': boxTare,
      'Numero invii totale scatola': itemsInThisBox,
      'Range lettere contenute': lettersRange,
      'Peso netto invii scatola (kg)': netWeightKg !== null ? Number(netWeightKg.toFixed(3)) : '',
      'Peso totale scatola (kg)': grossWeightKg !== null ? Number(grossWeightKg.toFixed(3)) : '',
      'Provincia bacino': provinciaBacino,
      'BACINO DESTINAZIONE': group.bacino,
      'Numero scatola nel gruppo': boxNumber,
      'Totale scatole nel gruppo': totalBoxes
    });
  }
}

function mapRowToChiudiGroup(row, options) {
  const mapped = getRowFormatAndWeight(row, options);
  const format = row.FORMATO_NORMALIZZATO || mapped.format || normalizeFormat(options.defaultFormat) || 'MIX';
  const weight = row.PESO_UNITARIO_GR || mapped.weightGrams || '';
  const thicknessMm = toNumber(row.SPESSORE_BUSTA_MM, mapped.thicknessMm || 0) || '';
  const elementsCount = Math.max(0, Math.floor(toNumber(row.NUMERO_ELEMENTI_BUSTA, mapped.elementsCount || 0))) || '';
  const bacino = row.BACINO_DESTINAZIONE || MIX_BACINO;
  const tariff = row.DESTINAZIONE_TARIFFARIA || MIX_TARIFF;
  const isMix = isMixBacinoValue(bacino) || tariff === MIX_TARIFF;
  const provinciaBacino = formatProvinciaBacino(row.PROVINCIA_BACINO || '', tariff, bacino);

  return {
    bacino,
    provinciaBacino,
    tariff,
    formato: format,
    pesoUnitario: weight,
    thicknessMm,
    elementsCount,
    closureType: isMix ? MIX_CLOSURE_TYPE : 'CHIUDI SCATOLA'
  };
}

function getGroupFirstRowIndex(group) {
  if (!group || !Array.isArray(group.rowIndexes) || group.rowIndexes.length === 0) {
    return Number.MAX_SAFE_INTEGER;
  }
  return group.rowIndexes[0];
}

function getGroupCapacityAnalysis(group, boxCapacity, boxHeightCm, boxReaction) {
  const limits = getCtaLimits(group.formato);
  if (!limits) {
    return null;
  }

  let unitWeightGrams = 0;
  for (const value of Array.isArray(group.rowWeights) ? group.rowWeights : []) {
    const numeric = toNumber(value, 0);
    if (numeric > unitWeightGrams) {
      unitWeightGrams = numeric;
    }
  }
  if (unitWeightGrams <= 0) {
    unitWeightGrams = toNumber(group.pesoUnitario, 0);
  }
  if (unitWeightGrams <= 0) {
    return null;
  }

  let thicknessMm = 0;
  for (const value of Array.isArray(group.rowThicknesses) ? group.rowThicknesses : []) {
    const numeric = toNumber(value, 0);
    if (numeric > thicknessMm) {
      thicknessMm = numeric;
    }
  }
  if (thicknessMm <= 0) {
    thicknessMm = toNumber(group.thicknessMm, 0);
  }

  const thresholds = calculatePieceThresholds(
    unitWeightGrams,
    thicknessMm > 0 ? thicknessMm : null,
    limits,
    boxCapacity,
    boxHeightCm,
    boxReaction
  );

  return {
    limits,
    unitWeightGrams,
    thresholds
  };
}

function getGroupNetWeightKg(group, fallbackUnitWeightGrams = 0) {
  let totalWeightGrams = 0;
  for (const value of Array.isArray(group.rowWeights) ? group.rowWeights : []) {
    const numeric = toNumber(value, 0);
    if (numeric > 0) {
      totalWeightGrams += numeric;
    }
  }

  if (totalWeightGrams > 0) {
    return totalWeightGrams / 1000;
  }

  const unitWeightGrams = fallbackUnitWeightGrams > 0 ? fallbackUnitWeightGrams : toNumber(group.pesoUnitario, 0);
  if (unitWeightGrams <= 0) {
    return 0;
  }

  return (group.count * unitWeightGrams) / 1000;
}

function getUniformPositiveNumber(values) {
  const sanitized = values
    .map((value) => toNumber(value, 0))
    .filter((value) => Number.isFinite(value) && value > 0);
  if (sanitized.length === 0) {
    return null;
  }

  const reference = sanitized[0];
  for (const value of sanitized) {
    if (value !== reference) {
      return null;
    }
  }

  return reference;
}

function mergePackingGroups(groups, mixRule = '') {
  const orderedGroups = [...groups].sort((left, right) => getGroupFirstRowIndex(left) - getGroupFirstRowIndex(right));
  const rowIndexes = [];
  const rowWeights = [];
  const rowThicknesses = [];
  const rowElements = [];
  const rowTariffs = [];
  const rowProvinceNames = [];
  const rowBacini = [];

  for (const group of orderedGroups) {
    rowIndexes.push(...(group.rowIndexes || []));
    rowWeights.push(...(group.rowWeights || []));
    rowThicknesses.push(...(group.rowThicknesses || []));
    rowElements.push(...(group.rowElements || []));
    rowTariffs.push(...(group.rowTariffs || []));
    rowProvinceNames.push(...(group.rowProvinceNames || []));
    rowBacini.push(...(group.rowBacini || []));
  }

  const uniqueBacini = new Set(rowBacini.filter(Boolean));
  const uniqueTariffs = new Set(rowTariffs.filter(Boolean));
  const uniqueFormats = new Set(orderedGroups.map((group) => group.formato).filter(Boolean));
  const formato = uniqueFormats.size === 1 ? [...uniqueFormats][0] : orderedGroups[0].formato || 'MIX';
  const bacino = uniqueBacini.size === 1 ? [...uniqueBacini][0] : MIX_BACINO;
  const tariff = uniqueTariffs.size === 1 ? [...uniqueTariffs][0] : MIX_TARIFF;
  const isMix = uniqueBacini.size > 1 || uniqueTariffs.size > 1 || isMixBacinoValue(bacino) || tariff === MIX_TARIFF;

  const uniqueWeightLabels = new Set(
    orderedGroups.map((group) => String(group.pesoUnitario ?? '').trim()).filter((value) => value !== '')
  );
  let pesoUnitario = '';
  if (uniqueWeightLabels.size === 1) {
    pesoUnitario = [...uniqueWeightLabels][0];
  } else if (uniqueWeightLabels.size > 1) {
    pesoUnitario = 'MIX';
  }

  const uniformThickness = getUniformPositiveNumber(rowThicknesses);
  const thicknessMm = uniformThickness !== null ? uniformThickness : '';
  const uniformElements = getUniformPositiveNumber(rowElements);
  const elementsCount = uniformElements !== null ? Math.floor(uniformElements) : '';
  const defaultClosureType = orderedGroups[0].closureType || 'CHIUDI SCATOLA';

  return {
    bacino,
    provinciaBacino: isMix ? MIX_PROVINCIA : orderedGroups[0].provinciaBacino || '',
    tariff,
    formato,
    pesoUnitario,
    thicknessMm,
    elementsCount,
    closureType: isMix || mixRule ? MIX_CLOSURE_TYPE : defaultClosureType,
    count: rowIndexes.length,
    rowIndexes,
    rowWeights,
    rowThicknesses,
    rowElements,
    rowTariffs,
    rowProvinceNames,
    rowBacini,
    sequenceKey: makeGroupKey([bacino, tariff, formato, pesoUnitario || 'NA', thicknessMm || 'NA', elementsCount || 'NA']),
    mixRule
  };
}

function runMixingStage(groups, keyBuilder, mixRule, boxCapacity, boxHeightCm, boxReaction) {
  const buckets = new Map();
  for (const group of groups) {
    const key = keyBuilder(group);
    if (!buckets.has(key)) {
      buckets.set(key, []);
    }
    buckets.get(key).push(group);
  }

  const resolved = [];
  const pending = [];

  for (const bucket of buckets.values()) {
    if (bucket.length === 1) {
      pending.push(bucket[0]);
      continue;
    }

    const merged = mergePackingGroups(bucket, mixRule);
    const analysis = getGroupCapacityAnalysis(merged, boxCapacity, boxHeightCm, boxReaction);
    const minHalfBoxItems = analysis ? Math.max(1, Math.ceil(analysis.thresholds.maxPerBox / 2)) : Number.MAX_SAFE_INTEGER;
    if (merged.count >= minHalfBoxItems) {
      resolved.push(merged);
    } else {
      pending.push(merged);
    }
  }

  return { resolved, pending };
}

function mixUnderfilledPackingGroups(groups, boxCapacity, boxHeightCm, boxReaction, warnings) {
  if (!groups || groups.length === 0) {
    return [];
  }

  let pending = [...groups].sort((left, right) => getGroupFirstRowIndex(left) - getGroupFirstRowIndex(right));
  const mixedGroups = [];

  const stage1 = runMixingStage(
    pending,
    (group) => makeGroupKey([group.formato || 'NA', group.pesoUnitario || 'NA']),
    'REGOLA_1_STESSO_FORMATO_STESSO_PESO',
    boxCapacity,
    boxHeightCm,
    boxReaction
  );
  mixedGroups.push(...stage1.resolved);
  pending = stage1.pending;

  const stage2 = runMixingStage(
    pending,
    (group) => makeGroupKey([group.formato || 'NA', group.tariff || 'NA']),
    'REGOLA_2_STESSO_FORMATO_STESSA_TARIFFA',
    boxCapacity,
    boxHeightCm,
    boxReaction
  );
  mixedGroups.push(...stage2.resolved);
  pending = stage2.pending;

  const stage3Buckets = new Map();
  for (const group of pending) {
    const key = makeGroupKey([group.formato || 'NA']);
    if (!stage3Buckets.has(key)) {
      stage3Buckets.set(key, []);
    }
    stage3Buckets.get(key).push(group);
  }

  for (const bucket of stage3Buckets.values()) {
    if (bucket.length === 1) {
      mixedGroups.push(bucket[0]);
      continue;
    }

    const merged = mergePackingGroups(bucket, 'REGOLA_3_STESSO_FORMATO_TARIFFE_DIVERSE');
    const analysis = getGroupCapacityAnalysis(merged, boxCapacity, boxHeightCm, boxReaction);
    const maxPerBox = analysis ? Math.max(1, Math.floor(toNumber(analysis.thresholds.maxPerBox, 1))) : Math.max(1, boxCapacity);
    const predictedBoxes = Math.ceil(merged.count / maxPerBox);
    const bacinoCount = new Set(
      merged.rowBacini.filter((value) => value && !isMixBacinoValue(value))
    ).size;

    if (bacinoCount > 0 && predictedBoxes > bacinoCount) {
      warnings.push(
        `Regola 3 non applicata per formato ${merged.formato}: scatole previste ${predictedBoxes} > bacini ${bacinoCount}.`
      );
      mixedGroups.push(...bucket);
      continue;
    }

    mixedGroups.push(merged);
  }

  return mixedGroups.sort((left, right) => getGroupFirstRowIndex(left) - getGroupFirstRowIndex(right));
}

function buildSequenceGroups(sortedRows, options, boxCapacity, boxHeightCm, boxReaction) {
  const groups = [];
  const groupsByKey = new Map();

  for (const [rowIndex, row] of sortedRows.entries()) {
    const mappedGroup = mapRowToChiudiGroup(row, options);
    const key = makeGroupKey([
      mappedGroup.bacino,
      mappedGroup.tariff,
      mappedGroup.formato,
      mappedGroup.pesoUnitario || 'NA',
      mappedGroup.thicknessMm || 'NA',
      mappedGroup.elementsCount || 'NA',
      mappedGroup.closureType
    ]);

    if (!groupsByKey.has(key)) {
      groupsByKey.set(key, {
        ...mappedGroup,
        count: 0,
        rowIndexes: [],
        rowWeights: [],
        rowThicknesses: [],
        rowElements: [],
        rowTariffs: [],
        rowProvinceNames: [],
        rowBacini: [],
        sequenceKey: key,
        mixRule: ''
      });
      groups.push(groupsByKey.get(key));
    }

    const group = groupsByKey.get(key);
    group.count += 1;
    group.rowIndexes.push(rowIndex);
    group.rowWeights.push(Math.max(0, toNumber(row.PESO_UNITARIO_GR, toNumber(mappedGroup.pesoUnitario, 0))));
    group.rowThicknesses.push(Math.max(0, toNumber(row.SPESSORE_BUSTA_MM, toNumber(mappedGroup.thicknessMm, 0))));
    group.rowElements.push(Math.max(0, Math.floor(toNumber(row.NUMERO_ELEMENTI_BUSTA, toNumber(mappedGroup.elementsCount, 0)))));
    group.rowTariffs.push(mappedGroup.tariff || MIX_TARIFF);
    group.rowProvinceNames.push(String(row.PROVINCIA_BACINO || '').trim().toUpperCase());
    group.rowBacini.push(mappedGroup.bacino || MIX_BACINO);
  }

  const warnings = [];
  const stableGroups = [];
  const underfilledGroups = [];

  for (const group of groups) {
    const analysis = getGroupCapacityAnalysis(group, boxCapacity, boxHeightCm, boxReaction);
    if (!analysis) {
      stableGroups.push(group);
      continue;
    }

    const minHalfBoxItems = Math.max(1, Math.ceil(analysis.thresholds.maxPerBox / 2));
    if (group.count < minHalfBoxItems) {
      underfilledGroups.push(group);
    } else {
      stableGroups.push(group);
    }
  }

  const mixedUnderfilledGroups = mixUnderfilledPackingGroups(
    underfilledGroups,
    boxCapacity,
    boxHeightCm,
    boxReaction,
    warnings
  );

  const finalGroups = [...stableGroups, ...mixedUnderfilledGroups].sort(
    (left, right) => getGroupFirstRowIndex(left) - getGroupFirstRowIndex(right)
  );

  return { groups: finalGroups, warnings };
}

function buildChiudiScatolaRows(sortedRows, options) {
  const boxCapacity = Math.max(1, Math.floor(toNumber(options.boxCapacity, DEFAULT_MAX_INVII_PER_BOX)));
  const boxHeightCm = Math.max(1, toNumber(options.boxHeightCm, DEFAULT_BOX_HEIGHT_CM));
  const boxReaction = Math.max(0.01, toNumber(options.boxReaction, DEFAULT_BOX_REACTION));
  const sequenceBuild = buildSequenceGroups(sortedRows, options, boxCapacity, boxHeightCm, boxReaction);
  const sequenceGroups = sequenceBuild.groups;
  const rowProgressivi = Array(sortedRows.length).fill('');
  const rowAssignmentOrder = Array(sortedRows.length).fill(Number.MAX_SAFE_INTEGER);
  const warnings = [...sequenceBuild.warnings];
  const stats = {
    standardBoxes: 0,
    partialBoxes: 0,
    mixBoxes: 0
  };
  const assignmentState = {
    nextProgressivo: 1,
    nextRowSequence: 1,
    rowAssignmentOrder,
    rowProgressivi
  };

  const chiudiRows = [];

  for (const sequenceGroup of sequenceGroups) {
    const group = sequenceGroup;
    const analysis = getGroupCapacityAnalysis(group, boxCapacity, boxHeightCm, boxReaction);
    if (!analysis) {
      const fallback = buildFallbackItemsByBox(group.count, boxCapacity);
      warnings.push(`Formato o peso non valido per il gruppo ${group.bacino}/${group.tariff}. Generato non conforme.`);
      appendBoxes(chiudiRows, group, fallback, options, stats, null, assignmentState);
      continue;
    }

    const groupNetKg = getGroupNetWeightKg(group, analysis.unitWeightGrams);
    if (groupNetKg < analysis.limits.minPartialKg) {
      warnings.push(
        `Gruppo ${group.bacino}/${group.tariff}/${group.formato} sotto minimo parziale CTA (${groupNetKg.toFixed(3)} kg). Generato non conforme.`
      );
      const fallback = buildFallbackItemsByBox(group.count, boxCapacity);
      appendBoxes(chiudiRows, group, fallback, options, stats, null, assignmentState);
      continue;
    }

    const itemsByBox = allocateItemsAcrossBoxes(group.count, analysis.thresholds);
    if (!itemsByBox || itemsByBox.length === 0) {
      warnings.push(`Impossibile allocare in modo conforme il gruppo ${group.bacino}/${group.tariff}. Generato non conforme.`);
      const fallback = buildFallbackItemsByBox(group.count, boxCapacity);
      appendBoxes(chiudiRows, group, fallback, options, stats, null, assignmentState);
      continue;
    }

    appendBoxes(chiudiRows, group, itemsByBox, options, stats, analysis.limits, assignmentState);
  }

  return {
    rows: chiudiRows,
    rowProgressivi,
    rowAssignmentOrder,
    warnings,
    stats,
    boxCapacity,
    boxHeightCm,
    boxReaction,
    capacityMode: options.elementsColumn ? 'ELEMENTI_AUTOMATICO' : 'LIMITE_MANUALE'
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
    const mapped = getRowFormatAndWeight(row, options);

    let status = 'OK';
    let note = '';

    if (!rule) {
      status = 'CAP_NON_CLASSIFICATO';
      note = 'CAP non trovato nel manuale bacini.';
    } else if (!mapped.format) {
      status = 'FORMATO_NON_VALIDO';
      note = 'Formato assente o non riconosciuto (atteso P o M).';
    } else if (options.elementsColumn && !mapped.elementsCount) {
      status = 'ELEMENTI_NON_VALIDI';
      note = 'Numero elementi busta non valido (atteso intero da 1 a 7).';
    } else if (!mapped.weightGrams) {
      status = 'PESO_NON_VALIDO';
      note = 'Peso unitario assente o non numerico.';
    }

    return {
      ...row,
      CAP_ORIGINALE: capValue,
      CAP_NORMALIZZATO: capNormalized || '',
      BACINO_DESTINAZIONE: rule ? rule.bacino : MIX_BACINO,
      DESTINAZIONE_TARIFFARIA: rule ? rule.tariff : '',
      PROVINCIA_BACINO: rule ? rule.province : MIX_PROVINCIA,
      FORMATO_NORMALIZZATO: mapped.format || '',
      NUMERO_ELEMENTI_BUSTA: mapped.elementsCount ?? '',
      PESO_UNITARIO_GR: mapped.weightGrams ?? '',
      SPESSORE_BUSTA_MM: mapped.thicknessMm ?? '',
      STATO_ELABORAZIONE: status,
      NOTE_ELABORAZIONE: note,
      __sourceIndex: index
    };
  });

  const groupingResult = applyMinimumInviiForBacinoGrouping(enrichedRows, MIN_INVII_PER_BACINO_GROUP);

  const sorter = createSorter();
  enrichedRows.sort(sorter);

  const rowsForExportBase = enrichedRows.map((row) => {
    const { __sourceIndex, ...cleaned } = row;
    return cleaned;
  });

  const chiudiBuild = buildChiudiScatolaRows(rowsForExportBase, options);
  const rowsForExport = rowsForExportBase
    .map((row, idx) => ({
      ...row,
      PROGRESSIVO_CHIUDI_SCATOLA: chiudiBuild.rowProgressivi[idx] || '',
      __ASSIGNMENT_ORDER: Number.isFinite(chiudiBuild.rowAssignmentOrder[idx])
        ? chiudiBuild.rowAssignmentOrder[idx]
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
  const chiudiRows = chiudiBuild.rows;

  const timestamp = createTimestamp();
  const baseName = path.parse(inputPath).name;

  const sortedOutputPath = path.join(outputDir, `${baseName}_record_ordinati_${timestamp}.xlsx`);
  const chiudiOutputPath = path.join(outputDir, `${baseName}_chiudi_scatola_${timestamp}.xlsx`);

  const sortedWorkbook = XLSX.utils.book_new();
  addSheet(sortedWorkbook, rowsForExport, 'Record_Ordinati');
  if (unmatchedRows.length > 0) {
    addSheet(sortedWorkbook, unmatchedRows, 'Scarti_Elaborazione');
  }
  XLSX.writeFile(sortedWorkbook, sortedOutputPath);

  const chiudiWorkbook = XLSX.utils.book_new();
  addSheet(chiudiWorkbook, chiudiRows, 'Chiudi_Scatola');
  XLSX.writeFile(chiudiWorkbook, chiudiOutputPath);

  return {
    inputPath,
    outputDir,
    sheetName,
    sortedOutputPath,
    chiudiOutputPath,
    totalRows: rowsForExport.length,
    matchedRows: rowsForExport.length - unmatchedRows.length,
    unmatchedRows: unmatchedRows.length,
    totalBoxes: chiudiRows.length,
    standardBoxes: chiudiBuild.stats.standardBoxes,
    partialBoxes: chiudiBuild.stats.partialBoxes,
    mixBoxes: chiudiBuild.stats.mixBoxes,
    ctaWarnings: chiudiBuild.warnings,
    boxCapacity: chiudiBuild.boxCapacity,
    boxHeightCm: chiudiBuild.boxHeightCm,
    boxReaction: chiudiBuild.boxReaction,
    capacityMode: chiudiBuild.capacityMode,
    elementsColumnUsed: Boolean(options.elementsColumn),
    rulesCount: rules.ranges.length,
    minInviiPerBacinoGroup: groupingResult.threshold,
    reroutedToMixRows: groupingResult.reroutedRows
  };
}

module.exports = {
  processWorkbook
};
