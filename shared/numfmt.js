// Excel-style number and date formatting.
//
// A cell is stored as a raw value plus a format code. A date is a number with a
// date format, exactly as Excel stores it, which is what makes date arithmetic
// work: the engine only ever does arithmetic on numbers.
//
// shortcut: this implements the format codes that appear in practice (numbers,
// decimals, grouping, percent, currency, dates, times, text, sections) and
// ignores colour sections. Upgrade to a full format-code state machine if
// `[Red]`-style rendering is ever needed.

export const GENERAL = 'General';

const EPOCH_1900 = Date.UTC(1899, 11, 30);
const DAY = 86400000;
const DAYS_1904_OFFSET = 1462;

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Built-in format ids, so a reader that only has an id still resolves a code. */
export const BUILTIN_FORMATS = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
  18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};

// --------------------------------------------------------------- date helpers

function shiftParts(serial) {
  // Excel counts 1900 as a leap year, so serials below 60 sit one day out.
  return serial < 60 ? serial + 1 : serial;
}

/** @returns the calendar parts a serial number represents (always UTC). */
export function serialToParts(serial, date1904 = false) {
  const days = shiftParts(date1904 ? serial + DAYS_1904_OFFSET : serial);
  const date = new Date(EPOCH_1900 + Math.round(days * DAY) + Math.round((days - Math.floor(days)) * DAY));
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    weekday: date.getUTCDay(),
    hours: date.getUTCHours(),
    minutes: date.getUTCMinutes(),
    seconds: date.getUTCSeconds(),
    ms: date.getUTCMilliseconds(),
    days: Math.floor(days),
  };
}

export function partsToSerial(parts, date1904 = false) {
  const date = Date.UTC(
    parts.year, parts.month - 1, parts.day,
    parts.hours || 0, parts.minutes || 0, parts.seconds || 0, parts.ms || 0,
  );
  const days = (date - EPOCH_1900) / DAY;
  const serial = days < 60 ? days - 1 : days;
  return date1904 ? serial - DAYS_1904_OFFSET : serial;
}

export function dateToSerial(date, date1904 = false) {
  return partsToSerial({
    year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate(),
    hours: date.getUTCHours(), minutes: date.getUTCMinutes(), seconds: date.getUTCSeconds(),
    ms: date.getUTCMilliseconds(),
  }, date1904);
}

// ------------------------------------------------------------ code inspection

/** Drop quoted literals, escapes, [bracketed] sections and `_` padding. */
function bareCode(code) {
  return String(code)
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/_./g, '');
}

export function isDateCode(code) {
  if (!code || code === GENERAL || code === '@') return false;
  const bare = bareCode(code);
  return /[ymdhs]/i.test(bare) && !/[#0?]/.test(bare);
}

// ---------------------------------------------------------------- date format

const DATE_TOKENS = /(\[h\]|\[m\]|\[s\]|am\/pm|a\/p|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|\.0+)/gi;

const pad = (value, width) => String(Math.abs(value)).padStart(width, '0');

function renderDate(parts, code) {
  const bare = bareCode(code);
  const lower = bare.toLowerCase();
  const hour12 = /am\/pm|a\/p/.test(lower);
  const hour24 = /\bh\b|\bhh\b/.test(lower) || /\[h\]/.test(lower);

  let out = '';
  let last = 0;
  let previous = '';

  DATE_TOKENS.lastIndex = 0;
  let match;
  while ((match = DATE_TOKENS.exec(lower))) {
    out += lower.slice(last, match.index);
    last = DATE_TOKENS.lastIndex;
    const token = match[0];

    if (/^\.0+$/.test(token)) { out += `.${pad(parts.ms, 3).slice(0, token.length - 1)}`; previous = token; continue; }

    // `m` means minutes when it follows an hour token or precedes a seconds one.
    const asMinutes = token === 'm' && (previous === 'h' || previous === 'hh' || /^:?s|^s/.test(lower.slice(last)));

    switch (token) {
      case 'yyyy': out += pad(parts.year, 4); break;
      case 'yy': out += pad(parts.year % 100, 2); break;
      case 'mmmm': out += MONTHS[parts.month - 1]; break;
      case 'mmm': out += MONTHS[parts.month - 1].slice(0, 3); break;
      case 'mm': out += pad(asMinutes ? parts.minutes : parts.month, 2); break;
      case 'm': out += String(asMinutes ? parts.minutes : parts.month); break;
      case 'dddd': out += WEEKDAYS[parts.weekday]; break;
      case 'ddd': out += WEEKDAYS[parts.weekday].slice(0, 3); break;
      case 'dd': out += pad(parts.day, 2); break;
      case 'd': out += String(parts.day); break;
      case 'hh': out += pad(hour12 ? (parts.hours % 12 || 12) : parts.hours, 2); break;
      case 'h': out += String(hour12 ? (parts.hours % 12 || 12) : parts.hours); break;
      case 'ss': out += pad(parts.seconds, 2); break;
      case 's': out += String(parts.seconds); break;
      case '[h]': out += String(Math.floor(parts.any * 24)); break;
      case '[m]': out += String(Math.floor(parts.any * 1440)); break;
      case '[s]': out += String(Math.floor(parts.any * 86400)); break;
      case 'am/pm': out += parts.hours < 12 ? 'AM' : 'PM'; break;
      case 'a/p': out += parts.hours < 12 ? 'A' : 'P'; break;
      default: out += token;
    }
    previous = token;
  }
  void hour24;
  return out + lower.slice(last);
}

// ------------------------------------------------------------- number format

function renderNumber(value, code) {
  const sections = String(code).split(';');

  // Section order is positive;negative;zero;text. With no negative section the
  // positive one is reused and we add the minus ourselves.
  const chosen = value < 0 ? sections[1] : (value === 0 ? (sections[2] ?? sections[0]) : sections[0]);
  const useSign = chosen === undefined || chosen === null;
  const section = useSign ? (sections[0] || GENERAL) : chosen;
  if (section === GENERAL || section === '') return String(value);

  const percent = /%/.test(section);
  let scaled = percent ? Math.abs(value) * 100 : Math.abs(value);

  // Each comma trailing the last placeholder divides by a thousand.
  const trailing = /[0#?](,+(?=[^0#?]*$))/.exec(section);
  if (trailing) scaled /= Math.pow(1000, trailing[1].length);

  const numberMatch = /[#0?][#0?,]*(?:\.[#0?]+)?/.exec(section);
  if (!numberMatch) return section.replace(/"/g, '');

  const [wholePattern, fractionPattern = ''] = numberMatch[0].split('.');
  const decimals = fractionPattern.length;
  const optionalDigits = (fractionPattern.match(/#/g) || []).length;
  const minIntegers = Math.max(1, (wholePattern.match(/[0?]/g) || []).length);

  let [whole, fraction] = scaled.toFixed(decimals).split('.');

  // `#` placeholders drop trailing zeros, `0` placeholders keep them.
  if (optionalDigits && fraction) {
    let left = optionalDigits;
    while (left > 0 && fraction.endsWith('0')) { fraction = fraction.slice(0, -1); left -= 1; }
  }
  if (/,/.test(wholePattern)) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  while (whole.replace(/\D/g, '').length < minIntegers) whole = `0${whole}`;

  const body = fraction ? `${whole}.${fraction}` : whole;
  const prefix = section.slice(0, numberMatch.index);
  const suffix = section.slice(numberMatch.index + numberMatch[0].length);
  return `${prefix}${useSign ? '-' : ''}${body}${suffix}`;
}

// ---------------------------------------------------------------- public API

/** Format a computed cell value for display. */
export function formatCell(value, code, date1904 = false) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') return value;
  if (!Number.isFinite(value)) return String(value);

  const format = code && code !== GENERAL ? code : null;
  if (!format) {
    // Excel shows at most 15 significant digits, which also hides float noise
    // such as 0.30000000000000004.
    const rounded = Number(value.toPrecision(15));
    if (rounded !== 0 && (Math.abs(rounded) >= 1e11 || Math.abs(rounded) < 1e-9)) return rounded.toExponential();
    return String(rounded);
  }
  if (format === '@') return String(value);
  if (isDateCode(format)) return renderDate(serialToParts(value, date1904), format);
  return renderNumber(value, format);
}

/**
 * Date literal -> `{serial, code}`, or null when the text is not a date.
 * Accepts ISO and unambiguous day/month forms only, so a date is never silently
 * reinterpreted when the month and day are both plausible.
 */
export function parseDateLiteral(text) {
  const source = String(text).trim();

  let match = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(source);
  if (!match) match = /^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(source);

  if (!match) {
    const loose = /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/.exec(source);
    if (loose) {
      const [, a, b, c] = loose;
      if (Number(a) > 12 && Number(b) <= 12) match = [source, c, b, a];
      else if (Number(b) > 12 && Number(a) <= 12) match = [source, c, a, b];
    }
  }

  if (!match) {
    const timeOnly = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(source);
    if (!timeOnly) return null;
    const [, h, mi, s] = timeOnly;
    if (Number(h) > 23 || Number(mi) > 59 || Number(s || 0) > 59) return null;
    return {
      serial: partsToSerial({ year: 1900, month: 1, day: 0, hours: Number(h), minutes: Number(mi), seconds: Number(s || 0) }),
      code: s === undefined ? 'h:mm' : 'h:mm:ss',
    };
  }

  const [, y, mo, d, h, mi, s] = match;
  const month = Number(mo);
  const day = Number(d);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  if (Number(h || 0) > 23 || Number(mi || 0) > 59 || Number(s || 0) > 59) return null;

  // Reject 2024-02-31 and friends.
  const check = new Date(Date.UTC(Number(y), month - 1, day));
  if (check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) return null;

  const hasTime = h !== undefined;
  return {
    serial: partsToSerial({
      year: Number(y), month, day,
      hours: Number(h || 0), minutes: Number(mi || 0), seconds: Number(s || 0),
    }),
    code: hasTime ? (s === undefined ? 'yyyy-mm-dd h:mm' : 'yyyy-mm-dd h:mm:ss') : 'yyyy-mm-dd',
  };
}

/** The formats offered in the spreadsheet toolbar. */
export const PRESETS = [
  { label: 'General', code: GENERAL },
  { label: 'Number', code: '0.00' },
  { label: 'Thousands', code: '#,##0' },
  { label: 'Thousands, 2dp', code: '#,##0.00' },
  { label: 'Percent', code: '0%' },
  { label: 'Percent, 2dp', code: '0.00%' },
  { label: 'Currency', code: '$#,##0.00' },
  { label: 'Currency, 0dp', code: '$#,##0' },
  { label: 'Scientific', code: '0.00E+00' },
  { label: 'Text', code: '@' },
  { label: 'Date', code: 'yyyy-mm-dd' },
  { label: 'Date, short', code: 'd-mmm-yy' },
  { label: 'Date & time', code: 'yyyy-mm-dd h:mm' },
  { label: 'Time', code: 'h:mm:ss' },
];
