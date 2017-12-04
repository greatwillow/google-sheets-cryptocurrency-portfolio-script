var portfolioSheetName = "Portfolio";
var historySheetName = "History";

function UPDATE() {
  var sheet = getSheetWithName(portfolioSheetName);

  range = sheet.getRange(3, 1, 16, 1);
  coins = range.getDisplayValues();

  // get secondary currency.
  range2 = sheet.getRange(2, 7, 1, 1);
  currency2 = range2.getDisplayValue();

  lastIdx = 0;

  Logger.log(range.getDisplayValues());
  for (var idx in coins) {
    coin = coins[idx][0];
    if (coin.length == 0) continue;
    coin = coin.trim().toLowerCase();
    coin = coin.replace(/ /g, "-");
    Logger.log(coin);
    if (coin == "-") {
      break;
    }

    r = getCoinMarketCapPrice(coin, currency2);

    outRange = sheet.getRange(3 + parseInt(idx), 2, 1, r.length);
    outRange.setValues([r]);
    lastIdx = parseInt(idx);
  }

  outRange = sheet.getRange(3 + parseInt(lastIdx) + 4, 2, 1, 1);
  outRange.setValues([["Updated: " + getDateTimeString()]]);
}

function RECORD_HISTORY() {
  var historySheet = getSheetWithName(historySheetName);
  if (historySheet == null) {
    historySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(
      historySheetName,
      1
    );
  }

  setupHistoryHeader(historySheet);

  range = historySheet.getRange(2, 1, 1, 100);
  var values = range.getValues();
  values[0][0] = new Date();
  historySheet.appendRow(values[0]);
}

function getSheetWithName(name) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var idx in sheets) {
    if (sheets[idx].getName() == name) {
      return sheets[idx];
    }
  }
  return null;
}

function getCoinMarketCapPrice(coin, convertCurrency) {
  var url =
    "https://api.coinmarketcap.com/v1/ticker/" +
    coin +
    "?convert=" +
    convertCurrency;
  var response = UrlFetchApp.fetch(url);
  var json = JSON.parse(response.getContentText());

  symbol = json[0]["symbol"];
  btc = parseFloat(json[0]["price_btc"]);
  usd = parseFloat(json[0]["price_usd"]);
  krw = parseFloat(json[0]["price_" + convertCurrency.toLowerCase()]);

  change_1h = parseFloat(json[0]["percent_change_1h"]) + "%";
  change_1d = parseFloat(json[0]["percent_change_24h"]) + "%";
  change_7d = parseFloat(json[0]["percent_change_7d"]) + "%";
  rank = parseInt(json[0]["rank"]);

  return [change_1h, change_1d, change_7d, btc, usd, krw];
}

function getDateTimeString(timestamp) {
  var date = new Date();

  if (0) {
    // Korean format
    var month = date.getUTCMonth() + 1;
    var day = date.getUTCDate();
    var hours = date.getHours();
    var minutes = "0" + date.getMinutes();
    var seconds = "0" + date.getSeconds();

    var formattedTime =
      month +
      "월 " +
      day +
      "일" +
      " " +
      hours +
      ":" +
      minutes.substr(-2) +
      ":" +
      seconds.substr(-2);

    return formattedTime;
  }
  return String(date);
}

function setupHistoryHeader(hsheet) {
  range = hsheet.getRange(1, 1, 1, 200);

  arr = range.getDisplayValues();
  //Logger.log(arr[0][0]);
  header = ["DATE", "USD TOTAL"];

  for (var idx in arr[0]) {
    h = arr[0][idx];
    if (idx < 2) continue;
    if (h.length == 0) break;
    if (h == "-") {
      break;
    }
    header.push(h);
  }

  var sheet = getSheetWithName(portfolioSheetName);
  {
    // 1. find new symbol and add to new header (out)
    range = sheet.getRange(3, 1, 100, 2);
    coins = range.getDisplayValues();

    for (var idx in coins) {
      coin = coins[idx][0];
      if (coin.length == 0) continue;
      if (coin == "-") {
        break;
      }
      coinSymbol = coins[idx][1];
      if (header.indexOf(coinSymbol) == -1) {
        header.push(coinSymbol);
      }
      //Logger.log(coinSymbol);
    }
  }
  header.push("-");

  //Logger.log(header);

  outRange = hsheet.getRange(1, 1, 1, header.length);
  outRange.setValues([header]);

  outRange = hsheet.getRange(4, 1, 1, header.length);
  outRange.setValues([header]);

  outRange = hsheet.getRange(2, 1, 1, 1);
  outRange.setValues([["*CURRENT*"]]);

  {
    // 2. find USD (out)
    range = sheet.getRange(2, 1, 1, 100);
    arr = range.getDisplayValues();
    c = arr[0].indexOf("Rank");
    if (c == -1) {
      c = arr[0].indexOf("24h");
    }
    usdIdx = arr[0].indexOf("USD", c) + 1;

    // 3. find new symbol and add to new header (out)
    range = sheet.getRange(3, 1, 100, 2);
    coins = range.getDisplayValues();
    //Logger.log(coins);
    for (var idx in coins) {
      coin = coins[idx][0];
      if (coin.length == 0) continue;
      if (coin == "-") {
        break;
      }
      coinSymbol = coins[idx][1];
      //Logger.log(coinSymbol);
      headerIdx = header.indexOf(coinSymbol);
      if (headerIdx == -1) {
        Logger.log("WHAT???");
      }
      range2 = sheet.getRange(3 + parseInt(idx), usdIdx, 1, 1);
      a1n = range2.getA1Notation();
      outRange = hsheet.getRange(2, headerIdx + 1, 1, 1);
      outRange.setFormula("=" + portfolioSheetName + "!" + a1n);

      lastIdx = parseInt(idx);
    }

    range2 = sheet.getRange(3 + parseInt(lastIdx) + 2, usdIdx, 1, 1);
    a1n = range2.getA1Notation();

    outRange = hsheet.getRange(2, 2, 1, 1);
    outRange.setFormula("=" + portfolioSheetName + "!" + a1n);
  }
}
