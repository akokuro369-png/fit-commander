function setupFitCommanderSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const definitions = [
    { name: 'exercise_master', headers: ['id', 'body_part', 'name', 'equipment', 'is_cardio', 'sort'] },
    { name: 'workout_log', headers: ['date', 'exercise_id', 'set_no', 'weight_kg', 'reps', 'duration_min', 'distance_km', 'memo'] },
    { name: 'body_log', headers: ['date', 'weight_kg', 'body_fat_pct', 'source'] },
    { name: 'health_log', headers: ['date', 'steps'] },
    { name: 'meditation_log', headers: ['date', 'duration_min', 'memo'] },
    { name: 'daily_summary', headers: ['date', 'body_parts', 'total_volume', 'cardio_min', 'steps', 'weight_kg', 'body_fat_pct', 'meditation_min'] },
    { name: 'advice', headers: ['date', 'text'] }
  ];

  definitions.forEach(def => {
    const sheet = ss.getSheetByName(def.name) || ss.insertSheet(def.name);
    sheet.clear();
    sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
    sheet.getRange(1, 1, 1, def.headers.length).setFontWeight('bold').setBackground('#E0F7FA');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, def.headers.length);
  });

  const exercises = [
    ['ex_001', '胸', 'ベンチプレス', 'バーベル', false, 1],
    ['ex_002', '胸', 'チェストプレス', 'マシン', false, 2],
    ['ex_003', '胸', 'ダンベルフライ', 'ダンベル', false, 3],
    ['ex_004', '胸', 'ペックフライ', 'マシン', false, 4],
    ['ex_005', '背中', 'ラットプルダウン', 'マシン', false, 10],
    ['ex_006', '背中', 'シーテッドロー', 'マシン', false, 11],
    ['ex_007', '背中', 'デッドリフト', 'バーベル', false, 12],
    ['ex_008', '背中', 'ベントオーバーロー', 'バーベル', false, 13],
    ['ex_009', '肩', 'ショルダープレス', 'マシン/ダンベル', false, 20],
    ['ex_010', '肩', 'サイドレイズ', 'ダンベル', false, 21],
    ['ex_011', '腕', 'アームカール', 'ダンベル/ケーブル', false, 30],
    ['ex_012', '腕', 'トライセプスプレスダウン', 'ケーブル', false, 31],
    ['ex_013', '脚', 'レッグプレス', 'マシン', false, 40],
    ['ex_014', '脚', 'レッグエクステンション', 'マシン', false, 41],
    ['ex_015', '脚', 'レッグカール', 'マシン', false, 42],
    ['ex_016', '脚', 'スクワット', 'スミス/バーベル', false, 43],
    ['ex_017', '腹', 'アブドミナルクランチ', 'マシン', false, 50],
    ['ex_018', '腹', 'プランク', '自重', false, 51],
    ['ex_019', '有酸素', 'トレッドミル', 'マシン', true, 60],
    ['ex_020', '有酸素', 'エアロバイク', 'マシン', true, 61],
    ['ex_021', '有酸素', 'クロストレーナー', 'マシン', true, 62]
  ];

  const master = ss.getSheetByName('exercise_master');
  master.getRange(2, 1, exercises.length, 6).setValues(exercises);
  definitions.forEach(def => ss.getSheetByName(def.name).autoResizeColumns(1, def.headers.length));

  SpreadsheetApp.getUi().alert('FitCommander sheets setup complete.\n\n7 sheets created and exercise_master initialized.');
}

