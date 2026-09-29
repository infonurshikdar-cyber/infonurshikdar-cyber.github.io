/*************************************************
 * CRICKET STATS - GOOGLE APPS SCRIPT BACKEND
 * Supports:
 * - Stats
 * - History
 * - New Match
 * - Add Player
 * - Adjust Stats
 * - Delete Player
 * - Reset Player
 * - Delete History
 * - Hat-trick
 * - Highest Run
 *************************************************/

const CONFIG = {
  STATS_SHEET: 'Stats',
  HISTORY_SHEET: 'History',

  // তোমার Admin PIN এখানে রাখা হয়েছে।
  // আগের PIN যদি 1234 না হয়, শুধু এই সংখ্যাটি বদলাবে।
  ADMIN_PIN: '44990'
};


/* =========================
   MAIN
========================= */

function doGet(e) {
  e = e || {};
  const p = e.parameter || {};

  const action = String(p.action || 'get');
  const callback = String(p.callback || '');

  let result;

  try {
    switch (action) {

      case 'get':
        result = getData();
        break;

      case 'checkPin':
        result = checkPin(p.pin);
        break;

      case 'save':
        result = saveMatch(p);
        break;

      case 'addPlayer':
        result = addPlayer(p);
        break;

      case 'updatePhoto':
        result = updatePhoto(p);
        break;

      case 'adjustStats':
        result = adjustStats(p);
        break;

      case 'deletePlayer':
        result = deletePlayer(p);
        break;

      case 'resetPlayer':
        result = resetPlayer(p);
        break;

      case 'deleteHistory':
        result = deleteHistory(p);
        break;

      default:
        result = {
          ok: false,
          error: 'Unknown action'
        };
    }

  } catch (err) {
    result = {
      ok: false,
      error: String(err && err.message ? err.message : err)
    };
  }

  return sendResponse(result, callback);
}


/* =========================
   RESPONSE
========================= */

function sendResponse(data, callback) {
  const json = JSON.stringify(data);

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}


/* =========================
   SHEETS
========================= */

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}


function getStatsSheet() {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.STATS_SHEET);

  if (!sh) {
    sh = ss.insertSheet(CONFIG.STATS_SHEET);
  }

  ensureStatsHeaders(sh);

  return sh;
}


function getHistorySheet() {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.HISTORY_SHEET);

  if (!sh) {
    sh = ss.insertSheet(CONFIG.HISTORY_SHEET);
  }

  ensureHistoryHeaders(sh);

  return sh;
}


/* =========================
   HEADERS
========================= */

function ensureStatsHeaders(sh) {

  const required = [
    'Player',
    'Total Match',
    'Won',
    'Lost',
    'Total Run',
    'Total Wicket',
    '50',
    '100',
    'Photo',
    'Hat-trick',
    'Highest Run'
  ];

  const lastColumn = Math.max(sh.getLastColumn(), 1);
  const firstRow = sh.getRange(1, 1, 1, lastColumn).getValues()[0];

  let headers = firstRow.map(function(x) {
    return String(x || '').trim();
  });

  if (headers.every(function(x) { return x === ''; })) {
    sh.getRange(1, 1, 1, required.length).setValues([required]);
    return;
  }

  // পুরোনো Header থাকলে নতুন কলাম যোগ করবে
  required.forEach(function(header) {

    if (headers.indexOf(header) === -1) {

      const newColumn = headers.length + 1;

      sh.getRange(1, newColumn).setValue(header);

      headers.push(header);
    }
  });
}


function ensureHistoryHeaders(sh) {

  const required = [
    'ID',
    'Player',
    'Match',
    'Won',
    'Lost',
    'Run',
    'Wicket',
    '50',
    '100',
    'Hat-trick',
    'Highest Run',
    'CreatedAt'
  ];

  const lastColumn = Math.max(sh.getLastColumn(), 1);
  const firstRow = sh.getRange(1, 1, 1, lastColumn).getValues()[0];

  let headers = firstRow.map(function(x) {
    return String(x || '').trim();
  });

  if (headers.every(function(x) { return x === ''; })) {
    sh.getRange(1, 1, 1, required.length).setValues([required]);
    return;
  }

  required.forEach(function(header) {

    if (headers.indexOf(header) === -1) {

      const newColumn = headers.length + 1;

      sh.getRange(1, newColumn).setValue(header);

      headers.push(header);
    }
  });
}


/* =========================
   HEADER MAP
========================= */

function getHeaderMap(sh) {

  const lastColumn = sh.getLastColumn();

  const headers = sh
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(function(x) {
      return String(x || '').trim();
    });

  const map = {};

  headers.forEach(function(header, index) {
    if (header) {
      map[header] = index + 1;
    }
  });

  return map;
}


/* =========================
   NUMBER HELPERS
========================= */

function num(value) {

  if (value === undefined || value === null || value === '') {
    return 0;
  }

  const n = Number(value);

  if (isNaN(n)) {
    return 0;
  }

  return n;
}


function nonNegative(value) {
  return Math.max(0, num(value));
}


/* =========================
   PIN
========================= */

function validPin(pin) {
  return String(pin || '') === String(CONFIG.ADMIN_PIN);
}


function checkPin(pin) {

  if (validPin(pin)) {
    return {
      ok: true,
      message: 'PIN correct'
    };
  }

  return {
    ok: false,
    error: 'ভুল PIN'
  };
}


/* =========================
   GET DATA
========================= */

function getData() {

  const sh = getStatsSheet();

  const lastRow = sh.getLastRow();
  const lastColumn = sh.getLastColumn();

  if (lastRow < 2) {
    return {
      ok: true,
      players: []
    };
  }

  const values = sh
    .getRange(1, 1, lastRow, lastColumn)
    .getValues();

  const headers = values[0].map(function(x) {
    return String(x || '').trim();
  });

  const players = [];

  for (let r = 1; r < values.length; r++) {

    const row = values[r];

    if (!row.length) continue;

    const obj = {};

    headers.forEach(function(header, i) {

      if (header) {
        obj[header] = row[i];
      }

    });

    if (!String(obj['Player'] || '').trim()) {
      continue;
    }

    players.push({
      player: String(obj['Player'] || ''),
      name: String(obj['Player'] || ''),
      match: nonNegative(obj['Total Match']),
      won: nonNegative(obj['Won']),
      lost: nonNegative(obj['Lost']),
      run: nonNegative(obj['Total Run']),
      wicket: nonNegative(obj['Total Wicket']),
      fifty: nonNegative(obj['50']),
      hundred: nonNegative(obj['100']),
      photo: String(obj['Photo'] || ''),
      hattrick: nonNegative(obj['Hat-trick']),
      highestRun: nonNegative(obj['Highest Run'])
    });
  }

  return {
    ok: true,
    players: players
  };
}


/* =========================
   FIND PLAYER
========================= */

function findPlayerRow(sh, player) {

  const map = getHeaderMap(sh);
  const playerColumn = map['Player'];

  if (!playerColumn) {
    return 0;
  }

  const lastRow = sh.getLastRow();

  if (lastRow < 2) {
    return 0;
  }

  const values = sh
    .getRange(2, playerColumn, lastRow - 1, 1)
    .getValues();

  const target = String(player || '').trim().toLowerCase();

  for (let i = 0; i < values.length; i++) {

    const name = String(values[i][0] || '')
      .trim()
      .toLowerCase();

    if (name === target) {
      return i + 2;
    }
  }

  return 0;
}


/* =========================
   SAVE NEW MATCH
========================= */

function saveMatch(p) {

  if (!validPin(p.pin)) {
    return {
      ok: false,
      error: 'ভুল PIN'
    };
  }

  const player = String(p.player || '').trim();

  if (!player) {
    return {
      ok: false,
      error: 'Player name missing'
    };
  }

  const sh = getStatsSheet();
  const map = getHeaderMap(sh);

  const row = findPlayerRow(sh, player);

  if (!row) {
    return {
      ok: false,
      error: 'Player পাওয়া যায়নি'
    };
  }

  const match = nonNegative(p.match);
  const won = nonNegative(p.won);
  const lost = nonNegative(p.lost);
  const run = nonNegative(p.run);
  const wicket = nonNegative(p.wicket);
  const fifty = nonNegative(p.fifty);
  const hundred = nonNegative(p.hundred);
  const hattrick = nonNegative(p.hattrick);

  // Highest Run = এই ম্যাচের রান
  const enteredHighestRun = nonNegative(
    p.highestRun !== undefined && p.highestRun !== ''
      ? p.highestRun
      : run
  );

  // পুরোনো data
  const oldMatch = getCellNumber(sh, row, map['Total Match']);
  const oldWon = getCellNumber(sh, row, map['Won']);
  const oldLost = getCellNumber(sh, row, map['Lost']);
  const oldRun = getCellNumber(sh, row, map['Total Run']);
  const oldWicket = getCellNumber(sh, row, map['Total Wicket']);
  const oldFifty = getCellNumber(sh, row, map['50']);
  const oldHundred = getCellNumber(sh, row, map['100']);
  const oldHattrick = getCellNumber(sh, row, map['Hat-trick']);
  const oldHighestRun = getCellNumber(sh, row, map['Highest Run']);

  // Update totals
  setCellNumber(
    sh,
    row,
    map['Total Match'],
    oldMatch + match
  );

  setCellNumber(
    sh,
    row,
    map['Won'],
    oldWon + won
  );

  setCellNumber(
    sh,
    row,
    map['Lost'],
    oldLost + lost
  );

  setCellNumber(
    sh,
    row,
    map['Total Run'],
    oldRun + run
  );

  setCellNumber(
    sh,
    row,
    map['Total Wicket'],
    oldWicket + wicket
  );

  setCellNumber(
    sh,
    row,
    map['50'],
    oldFifty + fifty
  );

  setCellNumber(
    sh,
    row,
    map['100'],
    oldHundred + hundred
  );

  setCellNumber(
    sh,
    row,
    map['Hat-trick'],
    oldHattrick + hattrick
  );

  // Highest Run কখনো মোট রান নয়।
  // শুধু এক ম্যাচের সর্বোচ্চ রান থাকবে।
  if (enteredHighestRun > oldHighestRun) {

    setCellNumber(
      sh,
      row,
      map['Highest Run'],
      enteredHighestRun
    );
  }

  // History
  const history = getHistorySheet();

  const id =
    new Date().getTime() +
    '_' +
    Math.floor(Math.random() * 100000);

  const historyMap = getHeaderMap(history);

  const historyRow = [];

  const historyHeaders = history
    .getRange(
      1,
      1,
      1,
      history.getLastColumn()
    )
    .getValues()[0];

  historyHeaders.forEach(function(header) {

    header = String(header || '').trim();

    switch (header) {

      case 'ID':
        historyRow.push(id);
        break;

      case 'Player':
        historyRow.push(player);
        break;

      case 'Match':
        historyRow.push(match);
        break;

      case 'Won':
        historyRow.push(won);
        break;

      case 'Lost':
        historyRow.push(lost);
        break;

      case 'Run':
        historyRow.push(run);
        break;

      case 'Wicket':
        historyRow.push(wicket);
        break;

      case '50':
        historyRow.push(fifty);
        break;

      case '100':
        historyRow.push(hundred);
        break;

      case 'Hat-trick':
        historyRow.push(hattrick);
        break;

      case 'Highest Run':
        historyRow.push(enteredHighestRun);
        break;

      case 'CreatedAt':
        historyRow.push(new Date());
        break;

      default:
        historyRow.push('');
    }
  });

  history.appendRow(historyRow);

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Stats saved successfully',
    player: player
  };
}


/* =========================
   ADD PLAYER
========================= */

function addPlayer(p) {

  if (!validPin(p.pin)) {
    return {
      ok: false,
      error: 'ভুল PIN'
    };
  }

  const player = String(p.player || '').trim();

  if (!player) {
    return {
      ok: false,
      error: 'Player name missing'
    };
  }

  const sh = getStatsSheet();

  const existing = findPlayerRow(sh, player);

  if (existing) {
    return {
      ok: false,
      error: 'এই Player আগে থেকেই আছে'
    };
  }

  const map = getHeaderMap(sh);

  const row = new Array(sh.getLastColumn()).fill('');

  row[map['Player'] - 1] = player;
  row[map['Total Match'] - 1] = 0;
  row[map['Won'] - 1] = 0;
  row[map['Lost'] - 1] = 0;
  row[map['Total Run'] - 1] = 0;
  row[map['Total Wicket'] - 1] = 0;
  row[map['50'] - 1] = 0;
  row[map['100'] - 1] = 0;
  row[map['Photo'] - 1] = '';
  row[map['Hat-trick'] - 1] = 0;
  row[map['Highest Run'] - 1] = 0;

  sh.appendRow(row);

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Player added'
  };
}


/* =========================
   UPDATE PHOTO
========================= */

function updatePhoto(p) {

  if (!validPin(p.pin)) {
    return {
      ok: false,
      error: 'ভুল PIN'
    };
  }

  const player = String(p.player || '').trim();
  const photo = String(p.photo || '');

  if (!player) {
    return {
      ok: false,
      error: 'Player name missing'
    };
  }

  const sh = getStatsSheet();
  const map = getHeaderMap(sh);
  const row = findPlayerRow(sh, player);

  if (!row) {
    return {
      ok: false,
      error: 'Player পাওয়া যায়নি'
    };
  }

  if (!map['Photo']) {
    return {
      ok: false,
      error: 'Photo column পাওয়া যায়নি'
    };
  }

  sh.getRange(row, map['Photo']).setValue(photo);

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Photo updated'
  };
}


/* =========================
   ADJUST STATS
========================= */

function adjustStats(p) {

  if (!validPin(p.pin)) {
    return {
      ok: false,
      error: 'ভুল PIN'
    };
  }

  const player = String(p.player || '').trim();
  const mode = String(p.mode || 'add').toLowerCase();

  if (!player) {
    return {
      ok: false,
      error: 'Player name missing'
    };
  }

  if (mode !== 'add' && mode !== 'subtract') {
    return {
      ok: false,
      error: 'Invalid mode'
    };
  }

  const sh = getStatsSheet();
  const map = getHeaderMap(sh);
  const row = findPlayerRow(sh, player);

  if (!row) {
    return {
      ok: false,
      error: 'Player পাওয়া যায়নি'
    };
  }

  const fields = [
    {
      key: 'match',
      column: 'Total Match'
    },
    {
      key: 'won',
      column: 'Won'
    },
    {
      key: 'lost',
      column: 'Lost'
    },
    {
      key: 'run',
      column: 'Total Run'
    },
    {
      key: 'wicket',
      column: 'Total Wicket'
    },
    {
      key: 'fifty',
      column: '50'
    },
    {
      key: 'hundred',
      column: '100'
    },
    {
      key: 'hattrick',
      column: 'Hat-trick'
    }
  ];

  fields.forEach(function(field) {

    const amount = nonNegative(p[field.key]);

    if (!amount) {
      return;
    }

    const column = map[field.column];

    if (!column) {
      return;
    }

    const oldValue = getCellNumber(
      sh,
      row,
      column
    );

    let newValue;

    if (mode === 'add') {
      newValue = oldValue + amount;
    } else {
      newValue = Math.max(
        0,
        oldValue - amount
      );
    }

    setCellNumber(
      sh,
      row,
      column,
      newValue
    );
  });


  /*
   * Highest Run:
   *
   * এটা মোট রান নয়।
   * Add করলে নতুন Highest Run value
   * যদি বর্তমানের চেয়ে বেশি হয়, সেট হবে।
   *
   * Subtract করলে:
   * বর্তমান Highest Run থেকে কমানো হবে।
   * 0-এর নিচে যাবে না।
   */

  const highestRunAmount = nonNegative(p.highestRun);

  if (
    highestRunAmount &&
    map['Highest Run']
  ) {

    const oldHighestRun = getCellNumber(
      sh,
      row,
      map['Highest Run']
    );

    let newHighestRun;

    if (mode === 'add') {

      newHighestRun = Math.max(
        oldHighestRun,
        highestRunAmount
      );

    } else {

      newHighestRun = Math.max(
        0,
        oldHighestRun - highestRunAmount
      );
    }

    setCellNumber(
      sh,
      row,
      map['Highest Run'],
      newHighestRun
    );
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    message:
      mode === 'add'
        ? 'Stats increased'
        : 'Stats decreased',
    player: player
  };
}


/* =========================
   DELETE PLAYER
========================= */

function deletePlayer(p) {

  if (!validPin(p.pin)) {
    return {
      ok: false,
      error: 'ভুল PIN'
    };
  }

  const player = String(p.player || '').trim();

  const sh = getStatsSheet();
  const row = findPlayerRow(sh, player);

  if (!row) {
    return {
      ok: false,
      error: 'Player পাওয়া যায়নি'
    };
  }

  sh.deleteRow(row);

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Player deleted'
  };
}


/* =========================
   RESET PLAYER
========================= */

function resetPlayer(p) {

  if (!validPin(p.pin)) {
    return {
      ok: false,
      error: 'ভুল PIN'
    };
  }

  const player = String(p.player || '').trim();

  const sh = getStatsSheet();
  const map = getHeaderMap(sh);
  const row = findPlayerRow(sh, player);

  if (!row) {
    return {
      ok: false,
      error: 'Player পাওয়া যায়নি'
    };
  }

  const resetColumns = [
    'Total Match',
    'Won',
    'Lost',
    'Total Run',
    'Total Wicket',
    '50',
    '100',
    'Hat-trick',
    'Highest Run'
  ];

  resetColumns.forEach(function(column) {

    if (map[column]) {
      sh.getRange(row, map[column]).setValue(0);
    }

  });

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Player stats reset'
  };
}


/* =========================
   DELETE HISTORY
========================= */

function deleteHistory(p) {

  if (!validPin(p.pin)) {
    return {
      ok: false,
      error: 'ভুল PIN'
    };
  }

  const history = getHistorySheet();

  const lastRow = history.getLastRow();

  if (lastRow < 2) {
    return {
      ok: false,
      error: 'History empty'
    };
  }

  const id = String(p.id || '').trim();

  if (!id) {
    return {
      ok: false,
      error: 'History ID missing'
    };
  }

  const map = getHeaderMap(history);

  if (!map['ID']) {
    return {
      ok: false,
      error: 'ID column পাওয়া যায়নি'
    };
  }

  const values = history
    .getRange(
      2,
      map['ID'],
      lastRow - 1,
      1
    )
    .getValues();

  let foundRow = 0;

  for (let i = 0; i < values.length; i++) {

    if (
      String(values[i][0] || '').trim() === id
    ) {
      foundRow = i + 2;
      break;
    }
  }

  if (!foundRow) {
    return {
      ok: false,
      error: 'History পাওয়া যায়নি'
    };
  }

  history.deleteRow(foundRow);

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'History deleted'
  };
}


/* =========================
   CELL HELPERS
========================= */

function getCellNumber(sh, row, column) {

  if (!column) {
    return 0;
  }

  return nonNegative(
    sh.getRange(row, column).getValue()
  );
}


function setCellNumber(sh, row, column, value) {

  if (!column) {
    return;
  }

  sh.getRange(row, column)
    .setValue(nonNegative(value));
}