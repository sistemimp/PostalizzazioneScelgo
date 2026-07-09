const fs = require('node:fs');
const path = require('node:path');
const { PDFParse } = require('pdf-parse');

const BACINI_ORDER = [
  'ANCONA',
  'BARI',
  'BOLOGNA',
  'CAGLIARI',
  'CATANIA',
  'FIRENZE',
  'GENOVA',
  'LAMEZIA TERME',
  'MILANO 1',
  'MILANO 2',
  'NAPOLI',
  'PADOVA 1',
  'PADOVA 2',
  'PALERMO',
  'PESCARA',
  'PISA',
  'ROMA',
  'TORINO',
  'VERONA'
];

const BACINI_SORTED = [...BACINI_ORDER].sort((a, b) => b.length - a.length);
const TARIFF_CODES = new Set(['AM', 'CP', 'EU', '-']);
const cache = new Map();

function normalizeLine(rawLine) {
  return String(rawLine || '').trim().replace(/\s+/g, ' ');
}

function isIgnorableLine(line) {
  if (!line) {
    return true;
  }

  const canon = line
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[’']/g, "'");

  if (
    line.startsWith('PRS.INDES') ||
    /\d+ di 4$/.test(line) ||
    canon.startsWith('condizioni') ||
    canon.startsWith('del servizio') ||
    canon.startsWith('elenco bacini') ||
    line.startsWith('* Le seguenti') ||
    /^[1-9 ]+$/.test(line) ||
    canon.startsWith('bacino centro di') ||
    canon.startsWith('riferimento provincia') ||
    canon.startsWith('citta provincia') ||
    canon.startsWith('dest.') ||
    canon.startsWith('tariff') ||
    canon.startsWith('primo') ||
    canon.startsWith('ultimo') ||
    canon.startsWith('allegato') ||
    canon.startsWith('legenda:') ||
    canon.startsWith('colonna ') ||
    canon.startsWith('nota') ||
    canon.startsWith("l'aggregazione") ||
    canon.startsWith('aggregati cap') ||
    canon.startsWith('da a') ||
    canon.startsWith('* da utilizzare') ||
    canon.startsWith('*  da utilizzare') ||
    canon.startsWith('tutti gli intervalli')
  ) {
    return true;
  }

  return false;
}

function readField(tokens, startIndex) {
  if (startIndex >= tokens.length) {
    return [null, startIndex];
  }

  if (tokens[startIndex] === 'VEDI' && tokens[startIndex + 1] === 'NOTA') {
    return ['VEDI NOTA', startIndex + 2];
  }

  return [tokens[startIndex], startIndex + 1];
}

function toNumber(value) {
  return /^\d{5}$/.test(String(value)) ? Number(value) : null;
}

async function parseBaciniRulesFromPdf(pdfPath) {
  const absolutePath = path.resolve(pdfPath);
  if (cache.has(absolutePath)) {
    return cache.get(absolutePath);
  }

  const buffer = fs.readFileSync(absolutePath);
  const parser = new PDFParse({ data: buffer });
  const data = await parser.getText();
  await parser.destroy();
  let text = data.text;

  text = text
    .replace(/76017ANDRIA/g, '76017 ANDRIA')
    .replace(/20099MILANO/g, '20099 MILANO');

  const lines = text
    .split(/\r?\n/)
    .map(normalizeLine)
    .filter((line) => !isIgnorableLine(line));

  let currentBacino = null;
  let lastProvince = null;
  const ranges = [];

  for (const line of lines) {
    let bacinoMatched = false;

    for (const bacino of BACINI_SORTED) {
      if (line === bacino || line.startsWith(`${bacino} `)) {
        const rest = line.slice(bacino.length).trim();
        const containsTariffCode = [' AM ', ' CP ', ' EU ', ' - '].some((snippet) => ` ${rest} `.includes(snippet));

        if (!rest || !containsTariffCode) {
          currentBacino = bacino;
          bacinoMatched = true;
          break;
        }
      }
    }

    if (bacinoMatched) {
      continue;
    }

    const tokens = line.split(' ');
    if (!tokens.some((token) => TARIFF_CODES.has(token))) {
      continue;
    }

    let index = 0;
    let province = lastProvince;

    if (!TARIFF_CODES.has(tokens[0])) {
      const provinceTokens = [];
      while (index < tokens.length && !TARIFF_CODES.has(tokens[index])) {
        provinceTokens.push(tokens[index]);
        index += 1;
      }
      province = provinceTokens.join(' ');
      lastProvince = province;
    }

    while (index < tokens.length) {
      if (!TARIFF_CODES.has(tokens[index])) {
        const provinceTokens = [];
        while (index < tokens.length && !TARIFF_CODES.has(tokens[index])) {
          provinceTokens.push(tokens[index]);
          index += 1;
        }
        if (provinceTokens.length > 0) {
          province = provinceTokens.join(' ');
          lastProvince = province;
        }
        continue;
      }

      const tariff = tokens[index];
      index += 1;

      const [fromRaw, nextIndex] = readField(tokens, index);
      index = nextIndex;
      const [toRaw, finalIndex] = readField(tokens, index);
      index = finalIndex;

      const from = toNumber(fromRaw);
      const to = toNumber(toRaw);

      if (tariff !== '-' && from !== null && to !== null && currentBacino) {
        ranges.push({
          bacino: currentBacino,
          province,
          tariff,
          from,
          to,
          special: false
        });
      }
    }
  }

  // Nei manuali Roma/Milano hanno dettaglio per distretti in nota (VEDI NOTA).
  // Per la generazione chiudi-scatola e l'ordinamento, classifichiamo AM su CAP municipali.
  ranges.push({ bacino: 'ROMA', province: 'ROMA', tariff: 'AM', from: 100, to: 199, special: true });
  ranges.push({ bacino: 'MILANO 1', province: 'MILANO', tariff: 'AM', from: 20100, to: 20199, special: true });

  const result = {
    generatedAt: new Date().toISOString(),
    sourcePdf: absolutePath,
    baciniOrder: BACINI_ORDER,
    ranges
  };

  cache.set(absolutePath, result);
  return result;
}

function normalizeCap(value) {
  if (value === null || value === undefined) {
    return null;
  }

  const digits = String(value).replace(/\D/g, '');
  if (!digits) {
    return null;
  }

  let normalized = digits;
  if (digits.length > 5) {
    normalized = digits.slice(0, 5);
  }

  if (normalized.length < 5) {
    normalized = normalized.padStart(5, '0');
  }

  return normalized;
}

function findRuleByCap(cap, rules) {
  const normalized = normalizeCap(cap);
  if (!normalized) {
    return null;
  }

  const capNumber = Number(normalized);
  const matches = rules.ranges.filter((range) => capNumber >= range.from && capNumber <= range.to);
  if (matches.length === 0) {
    return null;
  }

  matches.sort((a, b) => {
    const aSpan = a.to - a.from;
    const bSpan = b.to - b.from;
    if (aSpan !== bSpan) {
      return aSpan - bSpan;
    }

    const tariffPriority = { AM: 0, CP: 1, EU: 2 };
    const ta = tariffPriority[a.tariff] ?? 99;
    const tb = tariffPriority[b.tariff] ?? 99;
    if (ta !== tb) {
      return ta - tb;
    }

    return 0;
  });

  return {
    ...matches[0],
    capNormalized: normalized
  };
}

module.exports = {
  parseBaciniRulesFromPdf,
  findRuleByCap,
  normalizeCap,
  BACINI_ORDER
};
