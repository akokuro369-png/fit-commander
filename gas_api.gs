const SHEET_ID = 'PASTE_YOUR_FITCOMMANDER_SHEET_ID_HERE';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let result;

  try {
    switch (action) {
      case 'getExercises':
        result = ok({ exercises: getExercises(ss) });
        break;
      case 'getLastWorkout':
        result = ok(getLastWorkout(ss, e.parameter.exercise_id));
        break;
      case 'getCalendar':
        result = ok({ days: getCalendar(ss, Number(e.parameter.year), Number(e.parameter.month)) });
        break;
      case 'getToday':
        result = ok(getToday(ss, e.parameter.date));
        break;
      default:
        result = fail('unknown action: ' + action);
    }
  } catch (err) {
    result = fail(err.message || String(err));
  }

  return json(result);
}

function doPost(e) {
  const body = parseBody(e);
  const action = body.action || '';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const lock = LockService.getScriptLock();
  let result;

  try {
    lock.waitLock(10000);
    switch (action) {
      case 'addExercise':
        result = ok({ exercise: addExercise(ss, body) });
        break;
      case 'addWorkout':
        result = ok({ rows: addWorkout(ss, body) });
        break;
      case 'addBodyLog':
        result = ok({ row: addBodyLog(ss, body) });
        break;
      case 'addMeditation':
        result = ok({ row: addMeditation(ss, body) });
        break;
      case 'receiveHealth':
        result = ok(receiveHealth(ss, body));
        break;
      case 'aggregateDaily':
        result = ok({ rows: aggregateDaily(ss) });
        break;
      case 'setAdvice':
        result = ok({ advice: setAdvice(ss, body) });
        break;
      default:
        result = fail('unknown action: ' + action);
    }
  } catch (err) {
    result = fail(err.message || String(err));
  } finally {
    try {
      lock.releaseLock();
    } catch (err) {
      // Lock was not acquired or was already released.
    }
  }

  return json(result);
}

function ok(payload) {
  const base = { ok: true };
  return payload ? Object.assign(base, payload) : base;
}

function fail(message) {
  return { ok: false, error: message };
}

function json(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseBody(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  const text = e.postData.contents;
  try {
    return JSON.parse(text);
  } catch (err) {
    return e.parameter || {};
  }
}

function getExercises(ss) {
  return getSheet(ss, 'exercise_master')
    .sort((a, b) => Number(a.sort || 0) - Number(b.sort || 0));
}

function addExercise(ss, body) {
  const sheet = ss.getSheetByName('exercise_master');
  const exercises = getSheet(ss, 'exercise_master');
  const id = body.id || nextExerciseId(exercises);
  const sort = body.sort || nextSort(exercises, body.body_part);
  const row = [
    id,
    body.body_part || '',
    body.name || '',
    body.equipment || 'その他',
    normalizeBool(body.is_cardio),
    Number(sort) || 0
  ];
  sheet.appendRow(row);
  return rowToObject(getHeaders(sheet), row);
}

function getLastWorkout(ss, exerciseId) {
  if (!exerciseId) throw new Error('exercise_id is required');
  const rows = getSheet(ss, 'workout_log')
    .filter(row => row.exercise_id === exerciseId)
    .sort((a, b) => dateKey(b.date) - dateKey(a.date) || Number(a.set_no || 0) - Number(b.set_no || 0));

  if (!rows.length) {
    return { exercise_id: exerciseId, date: '', weight_kg: null, reps: null, sets: [] };
  }

  const latestDate = normalizeDate(rows[0].date);
  const sets = rows
    .filter(row => normalizeDate(row.date) === latestDate)
    .sort((a, b) => Number(a.set_no || 0) - Number(b.set_no || 0))
    .map(row => ({
      set_no: Number(row.set_no) || 0,
      weight_kg: nullableNumber(row.weight_kg),
      reps: nullableNumber(row.reps),
      duration_min: nullableNumber(row.duration_min),
      distance_km: nullableNumber(row.distance_km),
      memo: row.memo || ''
    }));

  const firstStrengthSet = sets.find(set => set.weight_kg !== null || set.reps !== null) || sets[0];
  return {
    exercise_id: exerciseId,
    date: latestDate,
    weight_kg: firstStrengthSet ? firstStrengthSet.weight_kg : null,
    reps: firstStrengthSet ? firstStrengthSet.reps : null,
    sets: sets
  };
}

function addWorkout(ss, body) {
  if (!body.exercise_id) throw new Error('exercise_id is required');
  if (!Array.isArray(body.sets) || !body.sets.length) throw new Error('sets must be a non-empty array');

  const date = normalizeDate(body.date || new Date());
  const rows = body.sets.map((set, i) => [
    date,
    body.exercise_id,
    Number(set.set_no || i + 1),
    blankIfNull(set.weight_kg),
    blankIfNull(set.reps),
    blankIfNull(set.duration_min),
    blankIfNull(set.distance_km),
    set.memo || body.memo || ''
  ]);

  const sheet = ss.getSheetByName('workout_log');
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  aggregateDaily(ss);
  return rows.length;
}

function addBodyLog(ss, body) {
  const row = [
    normalizeDate(body.date || new Date()),
    blankIfNull(body.weight_kg),
    blankIfNull(body.body_fat_pct),
    body.source || 'manual'
  ];
  ss.getSheetByName('body_log').appendRow(row);
  aggregateDaily(ss);
  return rowToObject(['date', 'weight_kg', 'body_fat_pct', 'source'], row);
}

function addMeditation(ss, body) {
  const row = [
    normalizeDate(body.date || new Date()),
    Number(body.duration_min) || 0,
    body.memo || ''
  ];
  ss.getSheetByName('meditation_log').appendRow(row);
  aggregateDaily(ss);
  return rowToObject(['date', 'duration_min', 'memo'], row);
}

function receiveHealth(ss, body) {
  const date = normalizeDate(body.date || new Date());
  const steps = Number(body.steps || 0);
  upsertByDate(ss.getSheetByName('health_log'), date, [date, steps]);

  let bodyLogUpdated = false;
  if (body.weight_kg !== undefined || body.body_fat_pct !== undefined) {
    ss.getSheetByName('body_log').appendRow([
      date,
      blankIfNull(body.weight_kg),
      blankIfNull(body.body_fat_pct),
      body.source || 'shortcut'
    ]);
    bodyLogUpdated = true;
  }

  aggregateDaily(ss);
  return { date: date, steps: steps, bodyLogUpdated: bodyLogUpdated };
}

function getCalendar(ss, year, month) {
  if (!year || !month) throw new Error('year and month are required');
  const prefix = year + '-' + pad2(month) + '-';
  return getSheet(ss, 'daily_summary').filter(row => String(row.date).indexOf(prefix) === 0);
}

function getToday(ss, dateValue) {
  const date = normalizeDate(dateValue || new Date());
  const summary = getSheet(ss, 'daily_summary').find(row => normalizeDate(row.date) === date) || emptySummary(date);
  const adviceText = getAdviceText(ss, date) || buildAdvice(summary);
  return {
    date: date,
    summary: summary,
    advice: adviceText
  };
}

function setAdvice(ss, body) {
  const date = normalizeDate(body.date || new Date());
  if (!date) throw new Error('date is required');
  if (body.text === undefined || body.text === null) throw new Error('text is required');
  const text = String(body.text);
  const sheet = getOrCreateAdviceSheet(ss);
  upsertByDate(sheet, date, [date, text]);
  return { date: date, text: text };
}

function getAdviceText(ss, date) {
  const row = getSheet(ss, 'advice').find(item => normalizeDate(item.date) === date);
  return row && row.text ? String(row.text) : '';
}

function getOrCreateAdviceSheet(ss) {
  let sheet = ss.getSheetByName('advice');
  if (!sheet) {
    sheet = ss.insertSheet('advice');
    sheet.getRange(1, 1, 1, 2).setValues([['date', 'text']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#E0F7FA');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 2);
  }
  return sheet;
}

function aggregateDailyJob() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return aggregateDaily(ss);
}

function aggregateDaily(ss) {
  const exercises = indexBy(getSheet(ss, 'exercise_master'), 'id');
  const daily = {};

  getSheet(ss, 'workout_log').forEach(row => {
    const date = normalizeDate(row.date);
    if (!date) return;
    const d = daily[date] || emptyDaily();
    const exercise = exercises[row.exercise_id] || {};
    if (exercise.body_part) d.bodyParts[exercise.body_part] = true;
    d.totalVolume += (Number(row.weight_kg) || 0) * (Number(row.reps) || 0);
    if (normalizeBool(exercise.is_cardio) || exercise.body_part === '有酸素') {
      d.cardioMin += Number(row.duration_min) || 0;
    }
    daily[date] = d;
  });

  getSheet(ss, 'health_log').forEach(row => {
    const date = normalizeDate(row.date);
    if (!date) return;
    const d = daily[date] || emptyDaily();
    d.steps = Number(row.steps) || 0;
    daily[date] = d;
  });

  getSheet(ss, 'body_log').forEach(row => {
    const date = normalizeDate(row.date);
    if (!date) return;
    const d = daily[date] || emptyDaily();
    if (row.weight_kg !== '') d.weightKg = nullableNumber(row.weight_kg);
    if (row.body_fat_pct !== '') d.bodyFatPct = nullableNumber(row.body_fat_pct);
    daily[date] = d;
  });

  getSheet(ss, 'meditation_log').forEach(row => {
    const date = normalizeDate(row.date);
    if (!date) return;
    const d = daily[date] || emptyDaily();
    d.meditationMin += Number(row.duration_min) || 0;
    daily[date] = d;
  });

  const rows = Object.keys(daily).sort().map(date => {
    const d = daily[date];
    return [
      date,
      Object.keys(d.bodyParts).join(','),
      d.totalVolume,
      d.cardioMin,
      d.steps,
      blankIfNull(d.weightKg),
      blankIfNull(d.bodyFatPct),
      d.meditationMin
    ];
  });

  const sheet = ss.getSheetByName('daily_summary');
  const headers = getHeaders(sheet);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  return rows.length;
}

function getSheet(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).filter(row => row.some(cell => cell !== '')).map(row => rowToObject(headers, row));
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((header, i) => {
    let value = row[i];
    if (value instanceof Date) value = normalizeDate(value);
    obj[header] = value;
  });
  return obj;
}

function upsertByDate(sheet, date, row) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (normalizeDate(values[i][0]) === date) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  sheet.appendRow(row);
}

function nextExerciseId(exercises) {
  const max = exercises.reduce((acc, row) => {
    const match = String(row.id || '').match(/^ex_(\d+)$/);
    return match ? Math.max(acc, Number(match[1])) : acc;
  }, 0);
  return 'ex_' + String(max + 1).padStart(3, '0');
}

function nextSort(exercises, bodyPart) {
  const samePart = exercises.filter(row => row.body_part === bodyPart);
  if (!samePart.length) return exercises.length + 1;
  return Math.max.apply(null, samePart.map(row => Number(row.sort) || 0)) + 1;
}

function indexBy(rows, key) {
  const result = {};
  rows.forEach(row => { result[row[key]] = row; });
  return result;
}

function emptyDaily() {
  return {
    bodyParts: {},
    totalVolume: 0,
    cardioMin: 0,
    steps: 0,
    weightKg: null,
    bodyFatPct: null,
    meditationMin: 0
  };
}

function emptySummary(date) {
  return {
    date: date,
    body_parts: '',
    total_volume: 0,
    cardio_min: 0,
    steps: 0,
    weight_kg: '',
    body_fat_pct: '',
    meditation_min: 0
  };
}

function buildAdvice(summary) {
  const parts = [];
  if (Number(summary.total_volume || 0) > 0) parts.push('筋トレ記録あり。次回は前回値を基準に小さく上積み。');
  if (Number(summary.cardio_min || 0) > 0) parts.push('有酸素を実施済み。回復を見ながら継続。');
  if (Number(summary.meditation_min || 0) > 0) parts.push('瞑想記録あり。心身の継続ログとして良い状態。');
  if (!parts.length) parts.push('今日は未記録。短時間の運動か瞑想から開始。');
  return parts.join('\n');
}

function normalizeDate(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value).slice(0, 10);
}

function dateKey(value) {
  const date = normalizeDate(value);
  return date ? Number(date.replace(/-/g, '')) : 0;
}

function normalizeBool(value) {
  return value === true || value === 'TRUE' || value === 'true' || value === 1 || value === '1';
}

function nullableNumber(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = Number(value);
  return Number.isNaN(n) ? null : n;
}

function blankIfNull(value) {
  return value === null || value === undefined ? '' : value;
}

function pad2(value) {
  return String(value).padStart(2, '0');
}


