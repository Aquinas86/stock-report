const CONFIG = {
  SHEET_NAME: '',
  MARKETS: ['KOSPI', 'KOSDAQ'],
  PAGE_SIZE: 50,
  MAX_PAGES: 80,
  MIN_RISE_PERCENT: 15,
  VOLUME_10M: 10000000,
  AMOUNT_500EOK: 50000000000,
  AMOUNT_1000EOK: 100000000000,
  HEADERS: [
    '월',
    '일',
    '종목',
    '상승률',
    '거래량',
    '거래대금',
    '유보율',
    '테마',
    '상승 이유',
    '천만주 이상 터짐',
    '500억 이상 터짐',
    '1000억 이상'
  ]
};

function installDailyStockTrigger() {
  deleteDailyStockTriggers_();
  ScriptApp.newTrigger('updateDailyStockRows')
    .timeBased()
    .everyDays(1)
    .atHour(16)
    .nearMinute(10)
    .create();
}

function deleteDailyStockTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'updateDailyStockRows') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function updateDailyStockRows() {
  const sheet = getTargetSheet_();
  ensureHeaders_(sheet);

  const rows = collectRisingStocks_();
  if (!rows.length) {
    Logger.log('No stocks matched today, or market data was unavailable.');
    return;
  }

  const tradeDate = rows[0].tradeDate;
  if (isWeekendOrHoliday_(tradeDate)) {
    Logger.log('Skip non-business day: ' + tradeDate);
    return;
  }

  const existing = getExistingKeys_(sheet);
  const values = rows
    .filter(function(row) {
      return !existing[row.month + '|' + row.day + '|' + row.name];
    })
    .map(function(row) {
      return [
        row.month,
        row.day,
        row.name,
        row.risePercent,
        row.volume,
        row.amountMillionKrw,
        row.reserveRatio,
        row.theme,
        row.reason,
        row.over10mVolume ? 'Yes' : 'No',
        row.over500eok ? 'Yes' : 'No',
        row.over1000eok ? 'Yes' : 'No'
      ];
    });

  if (!values.length) {
    Logger.log('Rows already exist for ' + tradeDate);
    return;
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, CONFIG.HEADERS.length).setValues(values);
  Logger.log('Inserted ' + values.length + ' rows for ' + tradeDate);
}

function collectRisingStocks_() {
  const result = [];
  const seen = {};

  CONFIG.MARKETS.forEach(function(market) {
    for (let page = 1; page <= CONFIG.MAX_PAGES; page++) {
      const data = fetchNaverStockPage_(market, page);
      const stocks = data.stocks || [];
      if (!stocks.length) break;

      stocks.forEach(function(stock) {
        const code = String(stock.itemCode || '').trim();
        const name = String(stock.stockName || '').trim();
        if (!code || !name || seen[code]) return;
        seen[code] = true;

        const risePercent = toNumber_(stock.fluctuationsRatio);
        if (risePercent < CONFIG.MIN_RISE_PERCENT) return;

        const tradeDate = parseTradeDate_(stock.localTradedAt);
        const valueRaw = toNumber_(stock.accumulatedTradingValueRaw);
        const volumeRaw = toNumber_(stock.accumulatedTradingVolumeRaw || stock.accumulatedTradingVolume);
        const reserveRatio = fetchReserveRatio_(code);

        result.push({
          tradeDate: tradeDate,
          month: Number(tradeDate.slice(5, 7)),
          day: Number(tradeDate.slice(8, 10)),
          code: code,
          name: name,
          risePercent: risePercent,
          volume: volumeRaw,
          amountMillionKrw: Math.round(valueRaw / 1000000),
          reserveRatio: reserveRatio,
          theme: '자동수집',
          reason: '네이버 종가 기준 15% 이상 상승. 상세 사유 확인 필요: https://m.stock.naver.com/domestic/stock/' + code + '/total',
          over10mVolume: volumeRaw >= CONFIG.VOLUME_10M,
          over500eok: valueRaw >= CONFIG.AMOUNT_500EOK,
          over1000eok: valueRaw >= CONFIG.AMOUNT_1000EOK
        });
      });

      if (data.totalCount && page * CONFIG.PAGE_SIZE >= Number(data.totalCount)) break;
      Utilities.sleep(120);
    }
  });

  result.sort(function(a, b) {
    return b.risePercent - a.risePercent;
  });
  return result;
}

function fetchNaverStockPage_(market, page) {
  const url = 'https://m.stock.naver.com/api/stocks/marketValue/' + market +
    '?page=' + page + '&pageSize=' + CONFIG.PAGE_SIZE;
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0',
      'Accept': 'application/json',
      'Referer': 'https://m.stock.naver.com/'
    }
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Naver stock page failed: ' + market + ' page ' + page + ' / ' + response.getResponseCode());
  }
  return JSON.parse(response.getContentText());
}

function fetchReserveRatio_(code) {
  try {
    const url = 'https://m.stock.naver.com/api/stock/' + code + '/finance/annual';
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0',
        'Accept': 'application/json',
        'Referer': 'https://m.stock.naver.com/'
      }
    });
    if (response.getResponseCode() !== 200) return '';

    const data = JSON.parse(response.getContentText());
    const finance = data.financeInfo || {};
    const titles = finance.trTitleList || [];
    const reserve = (finance.rowList || []).find(function(row) {
      return row.title === '유보율';
    });
    if (!reserve || !reserve.columns) return '';

    const keys = titles
      .filter(function(item) { return item.isConsensus !== 'Y'; })
      .map(function(item) { return item.key; })
      .reverse();

    for (let i = 0; i < keys.length; i++) {
      const cell = reserve.columns[keys[i]];
      if (cell && cell.value && cell.value !== '-') return toNumber_(cell.value);
    }
    return '';
  } catch (err) {
    return '';
  }
}

function getTargetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CONFIG.SHEET_NAME) {
    return ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  }
  return ss.getSheets()[0];
}

function ensureHeaders_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).getValues()[0];
  const hasHeaders = firstRow.some(function(value) {
    return String(value || '').trim() !== '';
  });
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  }
}

function getExistingKeys_(sheet) {
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  values.forEach(function(row) {
    if (row[0] && row[1] && row[2]) {
      map[Number(row[0]) + '|' + Number(row[1]) + '|' + String(row[2]).trim()] = true;
    }
  });
  return map;
}

function parseTradeDate_(value) {
  const text = String(value || '');
  const match = text.match(/\d{4}-\d{2}-\d{2}/);
  if (match) return match[0];
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}

function isWeekendOrHoliday_(yyyyMmDd) {
  const date = new Date(yyyyMmDd + 'T00:00:00+09:00');
  const day = date.getDay();
  if (day === 0 || day === 6) return true;

  const year = yyyyMmDd.slice(0, 4);
  const holidays = getKoreanHolidays_(year);
  return holidays.indexOf(yyyyMmDd) >= 0;
}

function getKoreanHolidays_(year) {
  try {
    const url = 'https://date.nager.at/api/v3/PublicHolidays/' + year + '/KR';
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return [year + '-12-31'];

    const items = JSON.parse(response.getContentText());
    const dates = items
      .filter(function(item) {
        return item.global === true && item.name !== 'Constitution Day';
      })
      .map(function(item) { return item.date; });
    dates.push(year + '-12-31');
    return dates;
  } catch (err) {
    return [year + '-12-31'];
  }
}

function toNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  return Number(String(value).replace(/,/g, '').replace(/[^\d.-]/g, '')) || 0;
}
