/**
 * Pure scheduling primitives for the calendar.
 * No Foundry dependencies — easy to unit test.
 *
 * Dates are { year, month, day } objects. Months are 1-indexed.
 * `calendarDef` is { monthsPerYear, daysPerMonth[], weekdays[] }.
 */

export const DEFAULT_CALENDAR = {
  monthsPerYear: 12,
  daysPerMonth:  [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31],
  weekdays:      ['Moonday', 'Toilday', 'Wealday', 'Oathday', 'Fireday', 'Starday', 'Sunday'],
  monthNames:    ['Abadius', 'Calistril', 'Pharast', 'Gozran', 'Desnus', 'Sarenith',
                  'Erastus', 'Arodus', 'Rova', 'Lamashan', 'Neth', 'Kuthona'],
};

export const GREGORIAN_CALENDAR = {
  monthsPerYear: 12,
  daysPerMonth:  [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31],
  weekdays:      ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
  monthNames:    ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December'],
};

export const GENERIC_FANTASY_CALENDAR = {
  monthsPerYear: 10,
  daysPerMonth:  [36, 36, 36, 36, 36, 36, 36, 36, 36, 36],
  weekdays:      ['Firstday', 'Secondday', 'Thirdday', 'Fourthday', 'Fifthday'],
  monthNames:    ['Frostmonth', 'Coldmonth', 'Thawmonth', 'Rainmonth', 'Seedmonth',
                  'Greenmonth', 'Sunmonth', 'Harvestmonth', 'Leafmonth', 'Darkmonth'],
};

export const CALENDAR_PRESETS = [
  { key: 'golarion',        label: 'Golarion (Inner Sea)',    def: DEFAULT_CALENDAR },
  { key: 'gregorian',       label: 'Gregorian (Earth)',       def: GREGORIAN_CALENDAR },
  { key: 'generic-fantasy', label: 'Generic Fantasy (10 months, 5-day week)', def: GENERIC_FANTASY_CALENDAR },
];

export function totalDaysInYear(cal = DEFAULT_CALENDAR) {
  return cal.daysPerMonth.reduce((a, b) => a + b, 0);
}

/** Compare two dates: returns -1/0/1. */
export function cmpDate(a, b) {
  if (a.year  !== b.year)  return a.year  < b.year  ? -1 : 1;
  if (a.month !== b.month) return a.month < b.month ? -1 : 1;
  if (a.day   !== b.day)   return a.day   < b.day   ? -1 : 1;
  return 0;
}

/** Add N days to a date, normalising overflow into months/years. */
export function addDays(date, n, cal = DEFAULT_CALENDAR) {
  let { year, month, day } = date;
  day += n;
  while (true) {
    if (day < 1) {
      month -= 1;
      if (month < 1) { month = cal.monthsPerYear; year -= 1; }
      day += cal.daysPerMonth[month - 1];
      continue;
    }
    const dim = cal.daysPerMonth[month - 1];
    if (day > dim) {
      day -= dim;
      month += 1;
      if (month > cal.monthsPerYear) { month = 1; year += 1; }
      continue;
    }
    break;
  }
  return { year, month, day };
}

/** Linear day-of-history index (lets us compute weekday & differences). */
export function dayOrdinal(date, cal = DEFAULT_CALENDAR) {
  const dpy = totalDaysInYear(cal);
  let d = date.year * dpy;
  for (let m = 0; m < date.month - 1; m++) d += cal.daysPerMonth[m];
  d += date.day - 1;
  return d;
}

export function weekdayIndex(date, cal = DEFAULT_CALENDAR) {
  return ((dayOrdinal(date, cal) % cal.weekdays.length) + cal.weekdays.length) % cal.weekdays.length;
}

/**
 * Given a rule and the current date, compute the next time this rule fires
 * strictly AFTER `fromDate`.
 *
 * Supported rule shapes:
 *   { every: 7, unit: 'day' }
 *   { every: 2, unit: 'week' }
 *   { every: 1, unit: 'month', dayOfMonth: 1 }
 *   { every: 1, unit: 'year',  month: 6, day: 21 }
 */
export function computeNext(rule, fromDate, cal = DEFAULT_CALENDAR) {
  const r = rule || {};
  const every = Math.max(1, Number(r.every) || 1);
  switch (r.unit) {
    case 'day':   return addDays(fromDate, every, cal);
    case 'week':  return addDays(fromDate, every * cal.weekdays.length, cal);
    case 'month': {
      let year = fromDate.year, month = fromDate.month + every;
      while (month > cal.monthsPerYear) { month -= cal.monthsPerYear; year += 1; }
      const dom = Math.max(1, Math.min(cal.daysPerMonth[month - 1], Number(r.dayOfMonth) || fromDate.day));
      return { year, month, day: dom };
    }
    case 'year': {
      const month = Math.max(1, Math.min(cal.monthsPerYear, Number(r.month) || fromDate.month));
      const day   = Math.max(1, Math.min(cal.daysPerMonth[month - 1], Number(r.day) || fromDate.day));
      let year = fromDate.year + every;
      // If the rule date is later in the current year, fire later this year first.
      const thisYearTarget = { year: fromDate.year, month, day };
      if (cmpDate(thisYearTarget, fromDate) > 0 && every === 1) return thisYearTarget;
      return { year, month, day };
    }
    default: return addDays(fromDate, every, cal);
  }
}

export function describeRule(rule, cal = DEFAULT_CALENDAR) {
  const r = rule || {};
  const every = Math.max(1, Number(r.every) || 1);
  switch (r.unit) {
    case 'day':   return every === 1 ? 'every day' : `every ${every} days`;
    case 'week':  return every === 1 ? 'every week' : `every ${every} weeks`;
    case 'month': return `every ${every === 1 ? '' : every + ' '}month${every === 1 ? '' : 's'} on day ${r.dayOfMonth || 1}`;
    case 'year': {
      const m = cal.monthNames?.[(r.month || 1) - 1] || `m${r.month}`;
      return `yearly on ${m} ${r.day || 1}`;
    }
    default: return 'unknown rule';
  }
}

export function formatDate(date, cal = DEFAULT_CALENDAR) {
  const m = cal.monthNames?.[date.month - 1] || `M${date.month}`;
  return `${m} ${date.day}, ${date.year}`;
}

export const SEASONS = ['spring', 'summer', 'autumn', 'winter'];

/** Quarter-of-the-year season, scaled to any monthsPerYear so custom calendars still get 4 seasons. */
export function seasonForMonth(month, cal = DEFAULT_CALENDAR) {
  const idx = Math.floor(((month - 1) / cal.monthsPerYear) * SEASONS.length) % SEASONS.length;
  return SEASONS[idx];
}

export const BIOMES = ['temperate', 'arid', 'coastal', 'arctic'];

const WEATHER_BY_BIOME = {
  temperate: {
    spring: ['clear', 'overcast', 'rain', 'rain', 'overcast'],
    summer: ['clear', 'clear', 'overcast', 'storm'],
    autumn: ['overcast', 'rain', 'clear', 'storm'],
    winter: ['overcast', 'snow', 'snow', 'clear'],
  },
  arid: {
    spring: ['clear', 'clear', 'overcast'],
    summer: ['clear', 'clear', 'clear', 'storm'],
    autumn: ['clear', 'overcast', 'clear'],
    winter: ['clear', 'overcast', 'clear'],
  },
  coastal: {
    spring: ['overcast', 'rain', 'clear'],
    summer: ['clear', 'overcast', 'storm'],
    autumn: ['rain', 'storm', 'overcast'],
    winter: ['rain', 'overcast', 'storm', 'snow'],
  },
  arctic: {
    spring: ['snow', 'overcast', 'clear'],
    summer: ['clear', 'overcast', 'rain'],
    autumn: ['snow', 'overcast', 'clear'],
    winter: ['snow', 'snow', 'storm', 'clear'],
  },
};

/** Weighted-by-repetition weather roll for a biome + season. `rng` is injectable for tests. */
export function rollWeather(biome, season, rng = Math.random) {
  const table = WEATHER_BY_BIOME[biome] || WEATHER_BY_BIOME.temperate;
  const options = table[season] || table.spring;
  return options[Math.floor(rng() * options.length)];
}
