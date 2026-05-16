const SPREADSHEET_ID = '1v5Czq5A28AyGo3gzPbaf77S1LT6Ai7SK93XzDl1piSM';
const SHEET_NAME = 'Sheet2';
const DAY_CELL = 'B1';
const PATIENT_COUNT_CELL = 'B2';
const SCHEDULE_RANGE = 'B3:B19';

const DAY_CYCLE = ['Monday', 'Tuesday', 'Wednesday', 'Thursday'];

const COLOR_CODES = {
  orange: 'o',
  blue: 'b',
  green: 'g',
  purple: 'p',
  o: 'o',
  b: 'b',
  g: 'g',
  p: 'p'
};

const WORD_NUMBERS = {
  one: 1,
  two: 2,
  three: 3,
  four: 4,
  five: 5,
  six: 6,
  seven: 7,
  eight: 8,
  nine: 9,
  ten: 10,
  eleven: 11,
  twelve: 12
};

const TIME_TO_OFFSET = {
  '9:00': 0,
  '9:30': 1,
  '10:00': 2,
  '10:30': 3,
  '11:00': 4,
  '11:30': 5,
  '12:00': 6,
  '12:30': 7,
  '1:00': 8,
  '1:30': 9,
  '2:00': 10,
  '2:30': 11,
  '3:00': 12,
  '3:30': 13,
  '4:00': 14,
  '4:30': 15
};

function doGet(event) {
  return handleRequest_(event);
}

function doPost(event) {
  return handleRequest_(event);
}

function handleRequest_(event) {
  try {
    const text = getText_(event);
    if (!text) {
      return json_({ ok: false, error: 'Missing text. Send ?text=9 orange, 11 blue or POST {"text":"..."}.' });
    }

    const result = updateScheduleFromText_(text);
    return json_({ ok: true, ...result });
  } catch (error) {
    return json_({ ok: false, error: error.message });
  }
}

function updateScheduleFromText_(text) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet not found: ${SHEET_NAME}`);

  const currentDay = String(sheet.getRange(DAY_CELL).getDisplayValue() || DAY_CYCLE[0]).trim();
  const requestedDay = findRequestedDay_(text);
  const sameDay = /\b(same day|same|fix|update|correction|edit)\b/i.test(text);
  const addOnly = /\b(add|append)\b/i.test(text);
  const nextDay = requestedDay || (sameDay ? normalizeDay_(currentDay) : advanceDay_(currentDay));
  const appointments = parseAppointments_(text);

  if (!appointments.length) {
    throw new Error('No appointments found. Try: 9 orange, 10:30 orange, 11 blue');
  }

  const values = addOnly ? sheet.getRange(SCHEDULE_RANGE).getValues() : Array.from({ length: 17 }, () => ['']);

  appointments.forEach((appointment) => {
    const offset = TIME_TO_OFFSET[appointment.time];
    if (offset === undefined) {
      throw new Error(`Time is outside the schedule: ${appointment.time}`);
    }
    values[offset][0] = appointment.code;
  });

  sheet.getRange(DAY_CELL).setValue(nextDay);
  sheet.getRange(PATIENT_COUNT_CELL).setFormula('=COUNTA(B3:B19)');
  sheet.getRange(SCHEDULE_RANGE).setValues(values);
  SpreadsheetApp.flush();

  return {
    day: nextDay,
    mode: addOnly ? 'add' : sameDay ? 'same-day' : requestedDay ? 'day-override' : 'advance',
    appointments,
    patientCount: appointments.length
  };
}

function getText_(event) {
  const parameterText = event?.parameter?.text;
  if (parameterText) return parameterText;

  const postData = event?.postData?.contents;
  if (!postData) return '';

  try {
    const body = JSON.parse(postData);
    return body.text || body.input || body.schedule || '';
  } catch (error) {
    return postData;
  }
}

function findRequestedDay_(text) {
  const lower = text.toLowerCase();
  return DAY_CYCLE.find((day) => lower.includes(day.toLowerCase())) || '';
}

function normalizeDay_(day) {
  const match = DAY_CYCLE.find((candidate) => candidate.toLowerCase() === String(day).toLowerCase());
  return match || DAY_CYCLE[0];
}

function advanceDay_(day) {
  const normalized = normalizeDay_(day);
  const index = DAY_CYCLE.indexOf(normalized);
  return DAY_CYCLE[(index + 1) % DAY_CYCLE.length];
}

function parseAppointments_(text) {
  const tokens = tokenize_(text);
  const appointments = [];

  for (let index = 0; index < tokens.length; index += 1) {
    const parsedTime = parseTime_(tokens, index);
    if (!parsedTime) continue;

    const color = findColor_(tokens, index + parsedTime.consumed);
    if (!color) continue;

    appointments.push({
      time: parsedTime.time,
      color: color.color,
      code: color.code
    });

    index = color.index;
  }

  return dedupeAppointments_(appointments);
}

function tokenize_(text) {
  return String(text)
    .toLowerCase()
    .replace(/[.,;:]/g, (match) => (match === ':' ? ':' : ' '))
    .replace(/\btelehealth\b/g, 'tele')
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .filter(Boolean);
}

function parseTime_(tokens, index) {
  const token = tokens[index];
  const next = tokens[index + 1];

  const colonMatch = token.match(/^(\d{1,2}):(\d{2})(am|pm)?$/);
  if (colonMatch) return formatTime_(Number(colonMatch[1]), Number(colonMatch[2]), 1);

  const compactMatch = token.match(/^(\d{3,4})(am|pm)?$/);
  if (compactMatch) {
    const digits = compactMatch[1];
    return formatTime_(Number(digits.slice(0, -2)), Number(digits.slice(-2)), 1);
  }

  const hourMatch = token.match(/^(\d{1,2})(am|pm)?$/);
  if (hourMatch) {
    const minuteInfo = parseMinuteToken_(next);
    return formatTime_(Number(hourMatch[1]), minuteInfo.minute, 1 + minuteInfo.consumed);
  }

  if (WORD_NUMBERS[token]) {
    const minuteInfo = parseMinuteToken_(next);
    return formatTime_(WORD_NUMBERS[token], minuteInfo.minute, 1 + minuteInfo.consumed);
  }

  return null;
}

function parseMinuteToken_(token) {
  if (token === '30' || token === 'thirty') return { minute: 30, consumed: 1 };
  return { minute: 0, consumed: 0 };
}

function formatTime_(hour, minute, consumed) {
  if (minute !== 0 && minute !== 30) return null;

  let normalizedHour = hour;
  if (normalizedHour >= 13) normalizedHour -= 12;
  if (normalizedHour === 0) normalizedHour = 12;

  const time = `${normalizedHour}:${String(minute).padStart(2, '0')}`;
  return { time, consumed };
}

function findColor_(tokens, startIndex) {
  const stopIndex = Math.min(tokens.length, startIndex + 4);
  for (let index = startIndex; index < stopIndex; index += 1) {
    const token = tokens[index];
    if (COLOR_CODES[token]) return { color: token, code: COLOR_CODES[token], index };
  }
  return null;
}

function dedupeAppointments_(appointments) {
  const byTime = {};
  appointments.forEach((appointment) => {
    byTime[appointment.time] = appointment;
  });
  return Object.keys(byTime)
    .sort((left, right) => TIME_TO_OFFSET[left] - TIME_TO_OFFSET[right])
    .map((time) => byTime[time]);
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function testParser_() {
  Logger.log(parseAppointments_('Tuesday 9 orange, 10:30 orange, eleven blue, 1 orange, three orange, 4 30 orange'));
}
