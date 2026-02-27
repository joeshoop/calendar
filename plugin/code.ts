// Calendar Updater — Figma Plugin
// Updates a 12-month calendar design for a given year.

// =============================================================================
// CONSTANTS
// =============================================================================

// US Pacific timezone: UTC offset in hours (standard / DST)
const TZ_STANDARD_OFFSET = -8; // PST
const TZ_DST_OFFSET = -7;      // PDT

// Check if a UTC date falls within US DST (2nd Sunday Mar to 1st Sunday Nov)
function isDST(utcMs: number): boolean {
  const d = new Date(utcMs);
  const year = d.getUTCFullYear();
  const month = d.getUTCMonth();

  if (month < 2 || month > 10) return false;  // Jan, Feb, Dec = no DST
  if (month > 2 && month < 10) return true;    // Apr-Oct = DST

  // March: DST starts 2nd Sunday at 2am local (10am UTC)
  if (month === 2) {
    const firstDay = new Date(Date.UTC(year, 2, 1)).getUTCDay();
    const secondSunday = 8 + ((7 - firstDay) % 7);
    const dstStartUTC = Date.UTC(year, 2, secondSunday, 10, 0, 0); // 2am PST = 10am UTC
    return utcMs >= dstStartUTC;
  }

  // November: DST ends 1st Sunday at 2am local (9am UTC, since still PDT)
  if (month === 10) {
    const firstDay = new Date(Date.UTC(year, 10, 1)).getUTCDay();
    const firstSunday = 1 + ((7 - firstDay) % 7);
    const dstEndUTC = Date.UTC(year, 10, firstSunday, 9, 0, 0); // 2am PDT = 9am UTC
    return utcMs < dstEndUTC;
  }

  return false;
}

// Convert a UTC timestamp (ms) to the local date in Pacific Time
function utcToLocalDate(utcMs: number): { year: number; month: number; day: number } {
  const offsetHours = isDST(utcMs) ? TZ_DST_OFFSET : TZ_STANDARD_OFFSET;
  const localMs = utcMs + offsetHours * 3600000;
  const d = new Date(localMs);
  return {
    year: d.getUTCFullYear(),
    month: d.getUTCMonth(),
    day: d.getUTCDate(),
  };
}

const MONTHS = [
  { name: "January",   frameId: "0:3"   },
  { name: "February",  frameId: "0:47"  },
  { name: "March",     frameId: "0:91"  },
  { name: "April",     frameId: "0:135" },
  { name: "May",       frameId: "0:187" },
  { name: "June",      frameId: "0:231" },
  { name: "July",      frameId: "0:275" },
  { name: "August",    frameId: "0:327" },
  { name: "September", frameId: "0:371" },
  { name: "October",   frameId: "0:415" },
  { name: "November",  frameId: "0:467" },
  { name: "December",  frameId: "0:514" },
];

const MOON_NAMES: { [key: number]: string } = {
  0: "Wolf Moon",
  1: "Snow Moon",
  2: "Crow Moon",
  3: "Pink Moon",
  4: "Flower Moon",
  5: "Strawberry Moon",
  6: "Raspberry Moon",
  7: "Blackberry Moon",
  8: "Harvest Moon",
  9: "Falling Leaf Moon",
  10: "Frost Moon",
  11: "Cold Moon",
};

const DAY_LETTERS = ["S", "M", "T", "W", "T", "F", "S"];

// Floating holiday images: instance name → target date per month
const HOLIDAY_IMAGES: Array<{
  monthIndex: number;
  imageName: string;
  getDay: (year: number) => number;
}> = [
  { monthIndex: 9,  imageName: "pumpkin", getDay: () => 31 },                                      // Halloween
  { monthIndex: 10, imageName: "turkey",  getDay: (year) => nthWeekdayOfMonth(year, 10, 4, 4) },   // Thanksgiving
  { monthIndex: 11, imageName: "santa",   getDay: () => 25 },                                      // Christmas
];

// Lunar New Year dates (month is 0-indexed)
const LUNAR_NEW_YEAR: { [key: number]: { month: number; day: number } } = {
  2024: { month: 1, day: 10 },
  2025: { month: 0, day: 29 },
  2026: { month: 1, day: 17 },
  2027: { month: 1, day: 6 },
  2028: { month: 0, day: 26 },
  2029: { month: 1, day: 13 },
  2030: { month: 1, day: 3 },
  2031: { month: 0, day: 23 },
  2032: { month: 1, day: 11 },
  2033: { month: 0, day: 31 },
  2034: { month: 1, day: 19 },
  2035: { month: 1, day: 8 },
  2036: { month: 0, day: 28 },
  2037: { month: 1, day: 15 },
  2038: { month: 1, day: 4 },
  2039: { month: 0, day: 24 },
  2040: { month: 1, day: 12 },
};

// =============================================================================
// DATE COMPUTATION
// =============================================================================

function daysInMonth(year: number, month: number): number {
  return new Date(year, month + 1, 0).getDate();
}

function dayOfWeek(year: number, month: number, day: number): number {
  return new Date(year, month, day).getDay();
}

// Compute Easter using the Anonymous Gregorian algorithm
function computeEaster(year: number): { month: number; day: number } {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31) - 1; // 0-indexed
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return { month, day };
}

// Supermoon lookup table: maps year to array of 1-indexed full moon ordinals
// that are supermoons. Source: astronomical almanac data.
// Supermoon = full moon within 90% of closest perigee for the year.
const SUPERMOONS: { [key: number]: number[] } = {
  2024: [10, 11, 12, 13], // Sep, Oct, Nov (x2)
  2025: [10, 11, 12],     // Oct, Nov, Dec
  2026: [1, 12, 13],      // Jan, Nov, Dec
  2027: [1, 2, 12, 13],   // Jan, Feb, Nov, Dec
  2028: [1, 2, 13],       // Jan, Feb, Dec
  2029: [2, 3],            // Feb, Mar
  2030: [2, 3, 4],         // Mar, Apr, May
  2031: [3, 4, 5],         // Apr, May, Jun
  2032: [5, 6, 7],         // May, Jun, Jul
  2033: [6, 7, 8],         // Jun, Jul, Aug
  2034: [7, 8, 9],         // Jul, Aug, Sep
  2035: [8, 9, 10],        // Aug, Sep, Oct
  2036: [9, 10, 11],       // Sep, Oct, Nov
  2037: [10, 11, 12],      // Oct, Nov, Dec
  2038: [10, 11, 12],      // Oct, Nov, Dec
  2039: [1, 11, 12, 13],   // Jan, Nov, Dec (x2)
  2040: [1, 2, 12, 13],    // Jan, Feb, Dec (x2)
};

// Compute full moon dates with timezone-aware local dates.
// Returns local dates in Pacific Time.
interface FullMoonInfo {
  month: number;
  day: number;
  isSuper: boolean;
  ordinal: number;  // 1-indexed position in the year's full moons
}

function computeFullMoonDates(year: number): FullMoonInfo[] {
  // Reference new moon: January 6, 2000 at 18:14 UTC
  const KNOWN_NEW_MOON_MS = Date.UTC(2000, 0, 6, 18, 14, 0);
  const SYNODIC_MONTH = 29.53058770576; // days
  const MS_PER_DAY = 86400000;

  const rawMoons: Array<{ utcMs: number; month: number; day: number }> = [];

  // Scan from a bit before the year to a bit after to catch boundary moons
  const scanStart = Date.UTC(year, 0, 1) - 35 * MS_PER_DAY;
  const scanEnd = Date.UTC(year + 1, 0, 1) + 5 * MS_PER_DAY;

  // Step through in 6-hour increments for higher precision
  const STEP_MS = 6 * 3600000;
  let prevPhase = -1;
  let prevMs = scanStart;

  for (let ms = scanStart; ms <= scanEnd; ms += STEP_MS) {
    const diffDays = (ms - KNOWN_NEW_MOON_MS) / MS_PER_DAY;
    const cyclePos = ((diffDays % SYNODIC_MONTH) + SYNODIC_MONTH) % SYNODIC_MONTH;
    const phase = cyclePos / SYNODIC_MONTH;

    if (prevPhase >= 0 && prevPhase < 0.5 && phase >= 0.5) {
      // Interpolate to find the approximate UTC moment of the full moon
      const prevDist = 0.5 - prevPhase;
      const curDist = phase - 0.5;
      const ratio = prevDist / (prevDist + curDist);
      const fullMoonMs = prevMs + ratio * STEP_MS;

      // Convert to local date in Pacific Time
      const localDate = utcToLocalDate(fullMoonMs);

      // Only include moons that fall in our target year (in local time)
      if (localDate.year === year) {
        rawMoons.push({
          utcMs: fullMoonMs,
          month: localDate.month,
          day: localDate.day,
        });
      }
    }
    prevPhase = phase;
    prevMs = ms;
  }

  // Look up which ordinals are supermoons for this year
  const superOrdinals = new Set(SUPERMOONS[year] || []);

  return rawMoons.map((m, i) => ({
    month: m.month,
    day: m.day,
    isSuper: superOrdinals.has(i + 1),
    ordinal: i + 1,
  }));
}

// Compute equinoxes and solstices using Meeus polynomial formulas.
// Returns local dates in Pacific Time.
function computeEquinoxesAndSolstices(year: number): Array<{ month: number; day: number; name: string }> {
  const Y = (year - 2000) / 1000;

  const springJDE = 2451623.80984 + 365242.37404 * Y + 0.05169 * Y * Y
    - 0.00411 * Y * Y * Y - 0.00057 * Y * Y * Y * Y;
  const summerJDE = 2451716.56767 + 365241.62603 * Y + 0.00325 * Y * Y
    + 0.00888 * Y * Y * Y - 0.00030 * Y * Y * Y * Y;
  const fallJDE = 2451810.21715 + 365242.01767 * Y - 0.11575 * Y * Y
    + 0.00337 * Y * Y * Y + 0.00078 * Y * Y * Y * Y;
  const winterJDE = 2451900.05952 + 365242.74049 * Y - 0.06223 * Y * Y
    - 0.00823 * Y * Y * Y + 0.00032 * Y * Y * Y * Y;

  // Convert Julian Ephemeris Day to UTC milliseconds
  function jdeToUtcMs(jde: number): number {
    // JDE 2440587.5 = Unix epoch (Jan 1, 1970 00:00 UTC)
    return (jde - 2440587.5) * 86400000;
  }

  // Convert JDE to local Pacific Time date
  function jdeToLocalDate(jde: number): { month: number; day: number } {
    const utcMs = jdeToUtcMs(jde);
    const local = utcToLocalDate(utcMs);
    return { month: local.month, day: local.day };
  }

  return [
    { ...jdeToLocalDate(springJDE), name: "Spring Equinox" },
    { ...jdeToLocalDate(summerJDE), name: "Summer Solstice" },
    { ...jdeToLocalDate(fallJDE), name: "Fall Equinox" },
    { ...jdeToLocalDate(winterJDE), name: "Winter Solstice" },
  ];
}

// Get the nth occurrence of a weekday in a month
function nthWeekdayOfMonth(year: number, month: number, weekday: number, n: number): number {
  const first = new Date(year, month, 1).getDay();
  return 1 + ((weekday - first + 7) % 7) + (n - 1) * 7;
}

// Get the last occurrence of a weekday in a month
function lastWeekdayOfMonth(year: number, month: number, weekday: number): number {
  const last = daysInMonth(year, month);
  const lastDow = new Date(year, month, last).getDay();
  return last - ((lastDow - weekday + 7) % 7);
}

// Compute sunrise or sunset time for a given date.
// Uses simplified solar position algorithm with atmospheric refraction correction.
// Returns a string like "5:23 PM" or "7:54 AM".
function computeSunEvent(
  year: number, month: number, day: number,
  lat: number, lng: number,
  type: "sunrise" | "sunset",
): string {
  // Day of year
  const date = new Date(year, month, day);
  const startOfYear = new Date(year, 0, 1);
  const doy = Math.floor((date.getTime() - startOfYear.getTime()) / 86400000) + 1;

  // Solar declination (simplified)
  const declRad = Math.asin(0.39779 * Math.sin(
    (Math.PI / 180) * (280.46646 + 0.9856474 * (doy - 1))
    - (Math.PI / 180) * 2.0 // rough equation of center correction
  ));

  const latRad = lat * Math.PI / 180;

  // Hour angle (sun at -0.833° below horizon for atmospheric refraction)
  const sunAngle = -0.833 * Math.PI / 180;
  const cosHA = (Math.sin(sunAngle) - Math.sin(latRad) * Math.sin(declRad))
    / (Math.cos(latRad) * Math.cos(declRad));

  // Clamp for polar edge cases
  if (cosHA > 1 || cosHA < -1) return type === "sunrise" ? "No sunrise" : "No sunset";

  const hourAngle = Math.acos(cosHA);

  // Solar noon approximation (hours UTC) via equation of time
  const B = (2 * Math.PI * (doy - 81)) / 365;
  const eqTime = 9.87 * Math.sin(2 * B) - 7.53 * Math.cos(B) - 1.5 * Math.sin(B);
  const solarNoonUTC = 12 - (lng / 15) - (eqTime / 60);

  // Add hour angle for sunset, subtract for sunrise
  const haDeg = (hourAngle * 180 / Math.PI) / 15;
  const eventUTC = type === "sunset" ? solarNoonUTC + haDeg : solarNoonUTC - haDeg;

  // Convert to local time
  const noonUtcMs = Date.UTC(year, month, day, 12, 0, 0);
  const offsetHours = isDST(noonUtcMs) ? TZ_DST_OFFSET : TZ_STANDARD_OFFSET;
  const localTime = eventUTC + offsetHours;

  // Format as "H:MM AM/PM"
  let hours = Math.floor(localTime);
  let minutes = Math.round((localTime - hours) * 60);
  if (minutes === 60) { hours++; minutes = 0; }

  const ampm = hours >= 12 ? "PM" : "AM";
  const displayHour = hours > 12 ? hours - 12 : hours === 0 ? 12 : hours;
  const displayMin = minutes.toString().padStart(2, "0");

  return `${displayHour}:${displayMin} ${ampm}`;
}

function computeSunrise(year: number, month: number, day: number, lat: number, lng: number): string {
  return computeSunEvent(year, month, day, lat, lng, "sunrise");
}

function computeSunset(year: number, month: number, day: number, lat: number, lng: number): string {
  return computeSunEvent(year, month, day, lat, lng, "sunset");
}

// =============================================================================
// GRID COMPUTATION
// =============================================================================

interface DayCell {
  date: number;
  month: number;
  year: number;
  isCurrentMonth: boolean;
}

function computeMonthGrid(year: number, month: number): DayCell[][] {
  const firstDay = dayOfWeek(year, month, 1);
  const numDays = daysInMonth(year, month);
  const totalCells = firstDay + numDays;
  const numRows = totalCells <= 28 ? 4 : totalCells <= 35 ? 5 : 6;

  const grid: DayCell[][] = [];
  let dayCounter = 1 - firstDay;

  for (let row = 0; row < numRows; row++) {
    const rowCells: DayCell[] = [];
    for (let col = 0; col < 7; col++) {
      if (dayCounter < 1) {
        const prevMonth = month === 0 ? 11 : month - 1;
        const prevYear = month === 0 ? year - 1 : year;
        const prevDays = daysInMonth(prevYear, prevMonth);
        rowCells.push({
          date: prevDays + dayCounter,
          month: prevMonth,
          year: prevYear,
          isCurrentMonth: false,
        });
      } else if (dayCounter > numDays) {
        const nextMonth = month === 11 ? 0 : month + 1;
        const nextYear = month === 11 ? year + 1 : year;
        rowCells.push({
          date: dayCounter - numDays,
          month: nextMonth,
          year: nextYear,
          isCurrentMonth: false,
        });
      } else {
        rowCells.push({
          date: dayCounter,
          month,
          year,
          isCurrentMonth: true,
        });
      }
      dayCounter++;
    }
    grid.push(rowCells);
  }

  return grid;
}

// =============================================================================
// HOLIDAY CATEGORIES
// =============================================================================

interface HolidayEntry {
  name: string;
  compute: (year: number) => { month: number; day: number };
}

const FEDERAL_HOLIDAYS: HolidayEntry[] = [
  { name: "New Year's Day",          compute: () => ({ month: 0, day: 1 }) },
  { name: "MLK Jr. Day",             compute: (y) => ({ month: 0, day: nthWeekdayOfMonth(y, 0, 1, 3) }) },
  { name: "Presidents' Day",         compute: (y) => ({ month: 1, day: nthWeekdayOfMonth(y, 1, 1, 3) }) },
  { name: "Memorial Day",            compute: (y) => ({ month: 4, day: lastWeekdayOfMonth(y, 4, 1) }) },
  { name: "Juneteenth",              compute: () => ({ month: 5, day: 19 }) },
  { name: "Independence Day",        compute: () => ({ month: 6, day: 4 }) },
  { name: "Labor Day",               compute: (y) => ({ month: 8, day: nthWeekdayOfMonth(y, 8, 1, 1) }) },
  { name: "Indigenous People's Day",  compute: (y) => ({ month: 9, day: nthWeekdayOfMonth(y, 9, 1, 2) }) },
  { name: "Veteran's Day",           compute: () => ({ month: 10, day: 11 }) },
  { name: "Thanksgiving",            compute: (y) => ({ month: 10, day: nthWeekdayOfMonth(y, 10, 4, 4) }) },
  { name: "Christmas Day",           compute: () => ({ month: 11, day: 25 }) },
];

// Fixed-date federal holidays that get observed-day adjustments (Sat→Fri, Sun→Mon)
const FEDERAL_OBSERVED = new Set([
  "New Year's Day", "Juneteenth", "Independence Day", "Veteran's Day", "Christmas Day",
]);

const OBSERVANCES: HolidayEntry[] = [
  { name: "Groundhog Day",     compute: () => ({ month: 1, day: 2 }) },
  { name: "Valentine's Day",   compute: () => ({ month: 1, day: 14 }) },
  { name: "St. Patrick's Day", compute: () => ({ month: 2, day: 17 }) },
  { name: "April Fool's Day",  compute: () => ({ month: 3, day: 1 }) },
  { name: "Earth Day",         compute: () => ({ month: 3, day: 22 }) },
  { name: "Cinco de Mayo",     compute: () => ({ month: 4, day: 5 }) },
  { name: "Mother's Day",      compute: (y) => ({ month: 4, day: nthWeekdayOfMonth(y, 4, 0, 2) }) },
  { name: "Father's Day",      compute: (y) => ({ month: 5, day: nthWeekdayOfMonth(y, 5, 0, 3) }) },
  { name: "Halloween",         compute: () => ({ month: 9, day: 31 }) },
  { name: "Christmas Eve",     compute: () => ({ month: 11, day: 24 }) },
  { name: "New Year's Eve",    compute: () => ({ month: 11, day: 31 }) },
  { name: "Easter",            compute: (y) => computeEaster(y) },
  { name: "Lunar New Year",    compute: (y) => LUNAR_NEW_YEAR[y] || { month: -1, day: -1 } },
  { name: "Election Day",      compute: (y) => ({ month: 10, day: nthWeekdayOfMonth(y, 10, 1, 1) + 1 }) },
  { name: "DST Starts",        compute: (y) => ({ month: 2, day: nthWeekdayOfMonth(y, 2, 0, 2) }) },
  { name: "DST Ends",          compute: (y) => ({ month: 10, day: nthWeekdayOfMonth(y, 10, 0, 1) }) },
];

// =============================================================================
// OPTIONS & BIRTHDAY PARSER
// =============================================================================

interface CalendarOptions {
  federalHolidays: boolean;
  observances: boolean;
  sunriseSunset: boolean;
  fullMoons: boolean;
  equinoxesSolstices: boolean;
  birthdays: Array<{ month: number; day: number; birthYear: number | null; name: string }>;
}

function parseBirthdays(text: string): Array<{ month: number; day: number; birthYear: number | null; name: string }> {
  const MONTH_ABBREVS: { [key: string]: number } = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
    jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
  };
  const result: Array<{ month: number; day: number; birthYear: number | null; name: string }> = [];
  for (const rawLine of text.split("\n")) {
    let line = rawLine.trim();
    if (!line) continue;
    // Strip legacy "birthday" prefix if present
    line = line.replace(/^birthday\s+/i, "");
    // With year: Jun 12 1984 Joe
    const withYear = line.match(/^(\w+)\s+(\d+)\s+(\d{4})\s+(.+)$/i);
    if (withYear) {
      const month = MONTH_ABBREVS[withYear[1].toLowerCase().substring(0, 3)];
      if (month !== undefined) {
        result.push({
          month,
          day: parseInt(withYear[2]),
          birthYear: parseInt(withYear[3]),
          name: withYear[4].trim(),
        });
      }
      continue;
    }
    // Without year: Jun 12 Joe
    const noYear = line.match(/^(\w+)\s+(\d+)\s+(.+)$/i);
    if (noYear) {
      const month = MONTH_ABBREVS[noYear[1].toLowerCase().substring(0, 3)];
      if (month !== undefined) {
        result.push({
          month,
          day: parseInt(noYear[2]),
          birthYear: null,
          name: noYear[3].trim(),
        });
      }
    }
  }
  return result;
}

// =============================================================================
// EVENTS MAP
// =============================================================================

interface CalendarEvent {
  label: string;
  isMoon?: boolean;
}

function buildEventsMap(year: number, lat: number, lng: number, options: CalendarOptions): Map<string, CalendarEvent[]> {
  const events = new Map<string, CalendarEvent[]>();

  function addEvent(eventYear: number, month: number, day: number, event: CalendarEvent) {
    const key = `${eventYear}-${month}-${day}`;
    if (!events.has(key)) events.set(key, []);
    events.get(key)!.push(event);
  }

  function addHolidayList(entries: HolidayEntry[], withObserved: boolean) {
    for (const entry of entries) {
      const { month, day } = entry.compute(year);
      if (month < 0 || day < 0) continue;
      addEvent(year, month, day, { label: entry.name });

      if (withObserved && FEDERAL_OBSERVED.has(entry.name)) {
        const dow = dayOfWeek(year, month, day);
        if (dow === 6) addEvent(year, month, day - 1, { label: `${entry.name} (Observed)` });
        else if (dow === 0) addEvent(year, month, day + 1, { label: `${entry.name} (Observed)` });
      }
    }
  }

  // --- Federal Holidays ---
  if (options.federalHolidays) {
    addHolidayList(FEDERAL_HOLIDAYS, true);
  }

  // --- Observances ---
  if (options.observances) {
    addHolidayList(OBSERVANCES, false);
  }

  // --- Equinoxes & Solstices ---
  if (options.equinoxesSolstices) {
    for (const es of computeEquinoxesAndSolstices(year)) {
      addEvent(year, es.month, es.day, { label: es.name });
    }
  }

  // --- Full Moons ---
  if (options.fullMoons) {
    const fullMoons = computeFullMoonDates(year);
    const moonCountByMonth: { [key: number]: number } = {};
    for (const fm of fullMoons) {
      moonCountByMonth[fm.month] = (moonCountByMonth[fm.month] || 0) + 1;
    }
    const moonSeenByMonth: { [key: number]: number } = {};
    for (const fm of fullMoons) {
      moonSeenByMonth[fm.month] = (moonSeenByMonth[fm.month] || 0) + 1;
      const isBlue = moonCountByMonth[fm.month] > 1 && moonSeenByMonth[fm.month] === 2;
      const baseName = MOON_NAMES[fm.month] || "Full Moon";
      let label = baseName;
      if (fm.isSuper && isBlue) label = `Super Blue ${baseName}`;
      else if (fm.isSuper) label = `Super ${baseName}`;
      else if (isBlue) label = `Blue ${baseName}`;
      addEvent(year, fm.month, fm.day, { label, isMoon: true });
    }
  }

  // --- Birthdays ---
  for (const bday of options.birthdays) {
    let label: string;
    if (bday.birthYear !== null) {
      const age = year - bday.birthYear;
      const suffix =
        age % 100 >= 11 && age % 100 <= 13 ? "th" :
        age % 10 === 1 ? "st" :
        age % 10 === 2 ? "nd" :
        age % 10 === 3 ? "rd" : "th";
      label = `${bday.name}'s ${age}${suffix} Birthday`;
    } else {
      label = `${bday.name}'s Birthday`;
    }
    addEvent(year, bday.month, bday.day, { label });
  }

  // --- Sunrise/sunset at the four corners of each month's displayed grid ---
  if (options.sunriseSunset) {
    const sunriseDone = new Set<string>();
    const sunsetDone = new Set<string>();
    for (let m = 0; m < 12; m++) {
      const grid = computeMonthGrid(year, m);
      const lastRow = grid.length - 1;
      const corners: Array<{ cell: DayCell; type: "sunrise" | "sunset" }> = [
        { cell: grid[0][0],        type: "sunrise" },
        { cell: grid[0][6],        type: "sunset"  },
        { cell: grid[lastRow][0],  type: "sunrise" },
        { cell: grid[lastRow][6],  type: "sunset"  },
      ];
      for (const { cell, type } of corners) {
        const key = `${cell.year}-${cell.month}-${cell.date}`;
        const done = type === "sunrise" ? sunriseDone : sunsetDone;
        if (done.has(key)) continue;
        done.add(key);
        if (type === "sunrise") {
          addEvent(cell.year, cell.month, cell.date, {
            label: `Sunrise ${computeSunrise(cell.year, cell.month, cell.date, lat, lng)}`,
          });
        } else {
          addEvent(cell.year, cell.month, cell.date, {
            label: `Sunset ${computeSunset(cell.year, cell.month, cell.date, lat, lng)}`,
          });
        }
      }
    }
  }

  return events;
}

// =============================================================================
// FIGMA NODE HELPERS
// =============================================================================

// Day component child node names (set in Figma's main Day component).
const NODE_DATE_NUMBER = "date-number";
const NODE_DAY_LETTER = "day-letter";
const NODE_DESCRIPTION = "day-description";
const NODE_MOON_DOT = "moon-dot";

// The main Day component, used for replacing broken instances.
let MAIN_DAY_COMPONENT: ComponentNode | null = null;

// Get the main Day component from any instance.
async function initMainComponent(sampleDay: InstanceNode): Promise<void> {
  MAIN_DAY_COMPONENT = await sampleDay.getMainComponentAsync();
  if (!MAIN_DAY_COMPONENT) {
    throw new Error("Could not get main component from Day instance");
  }
}

// Recursively collect all fonts from a node tree
function collectAllFonts(node: SceneNode, fontsToLoad: Set<string>) {
  if (node.type === "TEXT") {
    const textNode = node as TextNode;
    if (textNode.fontName === figma.mixed) {
      const len = textNode.characters.length;
      for (let i = 0; i < len; i++) {
        const font = textNode.getRangeFontName(i, i + 1) as FontName;
        fontsToLoad.add(JSON.stringify(font));
      }
    } else {
      fontsToLoad.add(JSON.stringify(textNode.fontName as FontName));
    }
  }
  if ("children" in node) {
    for (const child of (node as FrameNode).children) {
      collectAllFonts(child, fontsToLoad);
    }
  }
}

// Load fonts needed for nodes the plugin writes to (day cells)
async function loadRequiredFonts(monthFrames: FrameNode[]): Promise<void> {
  const fontsToLoad = new Set<string>();

  // Collect from the main Day component (has ALL children including hidden ones)
  const firstMonth = monthFrames[0];
  const frame1 = firstMonth.findOne(n => n.name === "Frame 1") as FrameNode;
  if (frame1) {
    const firstRow = frame1.children[0] as FrameNode;
    if (firstRow && firstRow.children.length > 0) {
      const sampleDay = firstRow.children[0] as InstanceNode;
      const mainComp = await sampleDay.getMainComponentAsync();
      if (mainComp) {
        collectAllFonts(mainComp, fontsToLoad);
      }
      // Also collect from the instance itself (in case of overrides)
      collectAllFonts(sampleDay, fontsToLoad);
    }
  }

  for (const fontStr of fontsToLoad) {
    await figma.loadFontAsync(JSON.parse(fontStr) as FontName);
  }
}

// =============================================================================
// CELL + MONTH UPDATE
// =============================================================================

// Opacity for overflow (prev/next month) day cells
const OVERFLOW_OPACITY = 0.9;
const CURRENT_MONTH_OPACITY = 1.0;

// Set text content and apply the month's foreground color
function setCharacters(textNode: TextNode, newText: string, fills: Paint[]): void {
  textNode.characters = newText;
  textNode.fills = fills;
}

// Sample the foreground color from a current-month date-number node.
// Searches rows for a day instance with opacity 1.0 (current month) and reads its fills.
function sampleForegroundColor(rows: FrameNode[]): Paint[] | null {
  for (const row of rows) {
    for (const child of row.children) {
      if (child.type !== "INSTANCE" || child.opacity !== CURRENT_MONTH_OPACITY) continue;
      const dateNode = (child as InstanceNode).findOne(n => n.name === NODE_DATE_NUMBER) as TextNode | null;
      if (dateNode && dateNode.fills !== figma.mixed) {
        return dateNode.fills as Paint[];
      }
    }
  }
  return null;
}

function updateDayCell(
  dayInstance: InstanceNode,
  cell: DayCell,
  events: CalendarEvent[],
  colIndex: number,
  rowIndex: number,
  foregroundFills: Paint[],
): void {
  dayInstance.opacity = cell.isCurrentMonth ? CURRENT_MONTH_OPACITY : OVERFLOW_OPACITY;

  const hasEvents = events.length > 0;

  // Find child nodes by their stable names set in the Day component
  const dateNode = dayInstance.findOne(n => n.name === NODE_DATE_NUMBER) as TextNode | null;
  const dayLetterNode = dayInstance.findOne(n => n.name === NODE_DAY_LETTER) as TextNode | null;
  const descNode = dayInstance.findOne(n => n.name === NODE_DESCRIPTION) as TextNode | null;
  const moonDotNode = dayInstance.findOne(n => n.name === NODE_MOON_DOT);

  // Update date number
  if (dateNode) {
    setCharacters(dateNode, cell.date.toString(), foregroundFills);
  }

  // Update day-of-week letter — only visible in the first row
  if (dayLetterNode) {
    if (rowIndex === 0) {
      setCharacters(dayLetterNode, DAY_LETTERS[colIndex], foregroundFills);
      dayLetterNode.visible = true;
    } else {
      dayLetterNode.visible = false;
    }
  }

  // Update description (holidays, events, sunrise/sunset)
  if (descNode) {
    if (hasEvents) {
      setCharacters(descNode, events.map(e => e.label).join("\n"), foregroundFills);
      descNode.visible = true;
    } else {
      setCharacters(descNode, " ", foregroundFills);
      descNode.visible = false;
    }
  }

  // Update moon dot visibility and color
  if (moonDotNode) {
    moonDotNode.visible = events.some(e => e.isMoon);
    if ("fills" in moonDotNode) {
      (moonDotNode as GeometryMixin & SceneNode).fills = foregroundFills;
    }
  }
}

function updateMonth(
  monthFrame: FrameNode,
  monthIndex: number,
  year: number,
  eventsMap: Map<string, CalendarEvent[]>,
): void {
  // Get the grid container
  const frame1 = monthFrame.findOne(n => n.name === "Frame 1") as FrameNode | null;
  if (!frame1) {
    console.warn(`No "Frame 1" found in ${MONTHS[monthIndex].name}`);
    return;
  }

  // Sort rows by their name ("Row 1", "Row 2", etc.)
  const rows = (frame1.children.filter(
    n => n.name.startsWith("Row ")
  ) as FrameNode[]).sort((a, b) => {
    const aNum = parseInt(a.name.replace("Row ", ""), 10);
    const bNum = parseInt(b.name.replace("Row ", ""), 10);
    return aNum - bNum;
  });

  // Compute the grid
  const grid = computeMonthGrid(year, monthIndex);
  const neededRows = grid.length; // 4, 5, or 6

  console.log(`${MONTHS[monthIndex].name}: ${rows.length} rows in Figma, grid needs ${neededRows} rows`);

  // Sample the month's foreground color before updating any cells
  const foregroundFills = sampleForegroundColor(rows);
  if (!foregroundFills) {
    console.warn(`Could not sample foreground color for ${MONTHS[monthIndex].name}, using default`);
  }
  const fills = foregroundFills || [{ type: "SOLID" as const, color: { r: 0, g: 0, b: 0 } }];

  // Show/hide rows based on how many the grid needs
  for (let r = 0; r < rows.length; r++) {
    rows[r].visible = r < neededRows;
  }

  // Update each cell
  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];
    const dayInstances = row.children as SceneNode[];

    if (rowIdx < grid.length) {
      // This row has grid data
      for (let colIdx = 0; colIdx < Math.min(7, dayInstances.length); colIdx++) {
        const instance = dayInstances[colIdx];
        if (instance.type !== "INSTANCE") continue;

        const cell = grid[rowIdx][colIdx];
        // Show events for all cells (including overflow days from prev/next month)
        const key = `${cell.year}-${cell.month}-${cell.date}`;
        const cellEvents = eventsMap.get(key) || [];

        updateDayCell(instance as InstanceNode, cell, cellEvents, colIdx, rowIdx, fills);
      }
    } else {
      // Extra row (Row 6 when only 5 rows needed) — clear it
      for (let colIdx = 0; colIdx < dayInstances.length; colIdx++) {
        const instance = dayInstances[colIdx];
        if (instance.type !== "INSTANCE") continue;

        // Use same empty-events path as updateDayCell
        updateDayCell(instance as InstanceNode, {
          date: 1, month: monthIndex, year, isCurrentMonth: false,
        }, [], colIdx, rowIdx, fills);
      }
    }
  }

  // Reposition any floating holiday images for this month
  for (const hi of HOLIDAY_IMAGES) {
    if (hi.monthIndex !== monthIndex) continue;

    const image = monthFrame.children.find(n => n.name === hi.imageName);
    if (!image) continue;

    const targetDay = hi.getDay(year);

    // Find the day cell for the target date in the grid
    for (let rowIdx = 0; rowIdx < grid.length; rowIdx++) {
      for (let colIdx = 0; colIdx < 7; colIdx++) {
        const cell = grid[rowIdx][colIdx];
        if (!cell.isCurrentMonth || cell.date !== targetDay) continue;

        const row = rows[rowIdx];
        if (!row || colIdx >= row.children.length) break;
        const dayInstance = row.children[colIdx];

        // Resize to 48px wide (maintain aspect ratio)
        if ("resize" in image) {
          const aspect = image.height / image.width;
          (image as FrameNode).resize(48, 48 * aspect);
        }

        // Position offset from the day cell's top-left corner
        const cellAbsX = dayInstance.absoluteTransform[0][2];
        const cellAbsY = dayInstance.absoluteTransform[1][2];
        const parentAbsX = monthFrame.absoluteTransform[0][2];
        const parentAbsY = monthFrame.absoluteTransform[1][2];

        image.x = (cellAbsX - parentAbsX) + 4;
        image.y = (cellAbsY - parentAbsY) + 4;

        console.log(`Repositioned "${hi.imageName}" to day ${targetDay} in ${MONTHS[monthIndex].name}`);
        break;
      }
    }
  }
}

// =============================================================================
// PLUGIN ENTRY POINT
// =============================================================================

figma.showUI(__html__, { width: 360, height: 460 });

// Send saved config to the UI on launch, and clean up legacy storage key
(async () => {
  figma.clientStorage.deleteAsync("eventsConfig");
  const saved = await figma.clientStorage.getAsync("calendarConfig");
  if (saved) {
    figma.ui.postMessage({ type: "loadConfig", config: saved });
  }
})();

interface UpdateMessage {
  type: string;
  year?: number;
  lat?: number;
  lng?: number;
  federalHolidays?: boolean;
  observances?: boolean;
  sunriseSunset?: boolean;
  fullMoons?: boolean;
  equinoxesSolstices?: boolean;
  birthdayText?: string;
}

figma.ui.onmessage = async (msg: UpdateMessage) => {
  if (msg.type !== "update" || !msg.year) return;

  const year = msg.year;
  const lat = msg.lat ?? 47.67;
  const lng = msg.lng ?? -122.38;
  const birthdayText = msg.birthdayText ?? "";

  try {
    // Step 1: Load all pages so we can find nodes
    figma.ui.postMessage({ type: "progress", message: "Loading pages..." });
    await figma.loadAllPagesAsync();

    // Step 2: Find all month frames (must use async with dynamic-page access)
    figma.ui.postMessage({ type: "progress", message: "Finding month frames..." });
    const monthFrames: FrameNode[] = [];
    for (const m of MONTHS) {
      const node = await figma.getNodeByIdAsync(m.frameId);
      if (!node || node.type !== "FRAME") {
        throw new Error(`Month frame not found: ${m.name} (ID: ${m.frameId}). ` +
          `Make sure the calendar file is open and node IDs haven't changed.`);
      }
      monthFrames.push(node as FrameNode);
    }

    // Step 3: Load fonts
    figma.ui.postMessage({ type: "progress", message: "Loading fonts..." });
    await loadRequiredFonts(monthFrames);

    // Step 4: Get the main Day component
    figma.ui.postMessage({ type: "progress", message: "Analyzing component structure..." });
    const discoverFrame1 = monthFrames[0].findOne(n => n.name === "Frame 1") as FrameNode;
    const discoverRow = discoverFrame1.children[0] as FrameNode;
    const sampleDay = discoverRow.children[0] as InstanceNode;
    await initMainComponent(sampleDay);

    // Step 5: Fix broken Day instances by replacing them with fresh ones.
    // Some instances lost child nodes (textframe text, dayframe text, moon dot) when
    // the Day component structure was modified. These children are permanently missing
    // and can't be restored via swapComponent, findOne, or getNodeByIdAsync.
    // The fix: detect broken instances, create fresh ones from the component, and swap them in.
    figma.ui.postMessage({ type: "progress", message: "Fixing broken day instances..." });
    if (MAIN_DAY_COMPONENT) {
      let fixedCount = 0;
      for (const monthFrame of monthFrames) {
        const frame1 = monthFrame.findOne(n => n.name === "Frame 1") as FrameNode | null;
        if (!frame1) continue;
        const rows = frame1.children.filter(n => n.name.startsWith("Row ")) as FrameNode[];
        for (const row of rows) {
          // Iterate by index so we can replace in place
          for (let i = 0; i < row.children.length; i++) {
            const child = row.children[i];
            if (child.type !== "INSTANCE") continue;
            const inst = child as InstanceNode;

            // Check if this instance is broken: can we find the day-description node?
            const descNode = inst.findOne(n => n.name === NODE_DESCRIPTION);
            const isBroken = !descNode;

            if (isBroken) {
              // Save properties from old instance
              const oldName = inst.name;
              const oldWidth = inst.width;
              const oldHeight = inst.height;

              // Save layout sizing if this is inside an auto-layout frame
              let oldLayoutH: "FIXED" | "HUG" | "FILL" | null = null;
              let oldLayoutV: "FIXED" | "HUG" | "FILL" | null = null;
              try {
                oldLayoutH = inst.layoutSizingHorizontal;
                oldLayoutV = inst.layoutSizingVertical;
              } catch (_) { /* not in auto-layout */ }

              // Create a fresh instance from the main component
              const newInst = MAIN_DAY_COMPONENT.createInstance();
              newInst.name = oldName;

              // Insert new instance at the same position in parent, THEN remove old
              row.insertChild(i, newInst);
              inst.remove();
              // After insertChild+remove, newInst is at index i

              // Now set size (after it's in the parent)
              newInst.resize(oldWidth, oldHeight);

              // Copy layout-related properties (now that it's a child of the row)
              try {
                if (oldLayoutH) newInst.layoutSizingHorizontal = oldLayoutH;
                if (oldLayoutV) newInst.layoutSizingVertical = oldLayoutV;
              } catch (_) { /* ignore if not applicable */ }

              fixedCount++;
            }
          }
        }
      }
      console.log(`Fixed ${fixedCount} broken Day instances (replaced with fresh ones)`);
    }

    // Step 6: Build options and events map
    figma.ui.postMessage({ type: "progress", message: "Computing calendar events..." });
    const options: CalendarOptions = {
      federalHolidays: msg.federalHolidays ?? true,
      observances: msg.observances ?? true,
      sunriseSunset: msg.sunriseSunset ?? true,
      fullMoons: msg.fullMoons ?? true,
      equinoxesSolstices: msg.equinoxesSolstices ?? true,
      birthdays: parseBirthdays(birthdayText),
    };
    await figma.clientStorage.setAsync("calendarConfig", {
      federalHolidays: options.federalHolidays,
      observances: options.observances,
      sunriseSunset: options.sunriseSunset,
      fullMoons: options.fullMoons,
      equinoxesSolstices: options.equinoxesSolstices,
      birthdayText,
    });
    const eventsMap = buildEventsMap(year, lat, lng, options);

    // Log events for verification
    console.log(`Events for ${year}:`);
    for (const [key, evts] of eventsMap) {
      console.log(`  ${key}: ${evts.map(e => e.label).join(", ")}`);
    }

    // Step 7: Update each month
    for (let i = 0; i < 12; i++) {
      figma.ui.postMessage({
        type: "progress",
        message: `Updating ${MONTHS[i].name}... (${i + 1}/12)`,
      });
      updateMonth(monthFrames[i], i, year, eventsMap);
    }

    // Done!
    figma.ui.postMessage({
      type: "done",
      message: `Calendar updated to ${year}!`,
    });
    figma.notify(`Calendar updated to ${year}! 🗓`);

  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("Calendar Updater error:", message);
    figma.ui.postMessage({ type: "error", message: `Error: ${message}` });
  }
};
