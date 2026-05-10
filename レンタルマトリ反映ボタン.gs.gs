/*******************************************************
 * レンタルマトリックス反映ボタン V1
 *
 * 対象：
 * - レンタル受け入れ
 * - 貸出累計
 * - その他出庫
 *
 * 対象外：
 * - 棚卸
 *
 * 重要：
 * - 貸出累計は在庫を減らさない
 * - 現在庫列には直接書き込まない
 *******************************************************/

function submitRentalMatrixV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SHEET_RENTAL = "レンタル";
  const SHEET_SKU = "SKU";
  const SHEET_SKU_LOG = "SKUログ";

  const HEADER_ROW = 6;
  const DATA_START_ROW = 7;

  const PROCESS_SOURCE = "RT";

  // SKUシート固定列
  const SKU_COL_CODE = 1;             // A列：061_SKUコード

  // 現在庫列：読むだけ。直接書き込まない
  const SKU_COL_HQ_CURRENT = 34;      // AH列：本部現在庫
  const SKU_COL_RENTAL_CURRENT = 45;  // AS列：レンタル現在庫

  // 累計列：ここに加算する
  const SKU_COL_HQ_PAYOUT = 32;       // AF列：本部払い出し
  const SKU_COL_RENTAL_RECEIVE = 43;  // AQ列：レンタル受け入れ
  const SKU_COL_RENTAL_OUT = 44;      // AR列：レンタルその他出庫
  const SKU_COL_RENTAL_LEND = 46;     // AT列：レンタル貸出累計

  const SIZE_LIST = ["XS", "S", "M", "L", "XL", "F"];

  // レンタルシート入力列
  const INPUT_COLS = {
    receive: { XS: 41, S: 42, M: 43, L: 44, XL: 45, F: 46 }, // AO:AT 受け入れ
    lend:    { XS: 48, S: 49, M: 50, L: 51, XL: 52, F: 53 }, // AV:BA 貸出累計
    out:     { XS: 55, S: 56, M: 57, L: 58, XL: 59, F: 60 }  // BC:BH その他出庫
  };

  const rentalSheet = ss.getSheetByName(SHEET_RENTAL);
  const skuSheet = ss.getSheetByName(SHEET_SKU);
  const logSheet = ss.getSheetByName(SHEET_SKU_LOG);

  if (!rentalSheet || !skuSheet || !logSheet) {
    Browser.msgBox("レンタル / SKU / SKUログ のいずれかのシートが見つかりません。");
    return;
  }

  const confirm = Browser.msgBox(
    "レンタルマトリックス反映",
    "受け入れ・貸出累計・その他出庫をSKUへ反映します。\n棚卸は処理しません。\n\n実行しますか？",
    Browser.Buttons.YES_NO
  );

  if (confirm !== "yes") return;

  const rentalColMap = getRentalMatrixColMap_(rentalSheet, HEADER_ROW);
  const logColMap = getRentalMatrixColMap_(logSheet, HEADER_ROW);

  if (!rentalColMap["064"] || !rentalColMap["17"]) {
    Browser.msgBox("レンタルシートに 064_ または 17_ が見つかりません。");
    return;
  }

  const requiredLogIds = ["01", "02", "03", "04", "05", "061", "064", "09", "17", "10", "11", "12"];
  const missingLogIds = requiredLogIds.filter(id => !logColMap[id]);

  if (missingLogIds.length > 0) {
    Browser.msgBox("SKUログに必要な項目IDがありません：\n" + missingLogIds.join(", "));
    return;
  }

  const lastRentalRow = rentalSheet.getLastRow();
  if (lastRentalRow < DATA_START_ROW) {
    Browser.msgBox("レンタルシートに処理対象データがありません。");
    return;
  }

  const rentalValues = rentalSheet.getRange(
    DATA_START_ROW,
    1,
    lastRentalRow - DATA_START_ROW + 1,
    rentalSheet.getLastColumn()
  ).getValues();

  const lastSkuRow = skuSheet.getLastRow();
  const skuValues = skuSheet.getRange(
    DATA_START_ROW,
    1,
    lastSkuRow - DATA_START_ROW + 1,
    skuSheet.getLastColumn()
  ).getValues();

  const skuMap = new Map();

  skuValues.forEach((row, i) => {
    const skuCode = String(row[SKU_COL_CODE - 1] || "").trim();
    if (!skuCode) return;

    skuMap.set(skuCode, {
      rowNumber: DATA_START_ROW + i,

      hqCurrent: Number(row[SKU_COL_HQ_CURRENT - 1] || 0),
      rentalCurrent: Number(row[SKU_COL_RENTAL_CURRENT - 1] || 0),

      hqPayout: Number(row[SKU_COL_HQ_PAYOUT - 1] || 0),
      rentalReceive: Number(row[SKU_COL_RENTAL_RECEIVE - 1] || 0),
      rentalOut: Number(row[SKU_COL_RENTAL_OUT - 1] || 0),
      rentalLend: Number(row[SKU_COL_RENTAL_LEND - 1] || 0),

      changed: false
    });
  });

  const processId = createRentalMatrixProcessId_(logSheet, logColMap["01"], PROCESS_SOURCE);
  const now = new Date();

  const logRows = [];
  const clearRanges = [];
  const errors = [];

  rentalValues.forEach((row, rowIndex) => {
    const sheetRow = DATA_START_ROW + rowIndex;

    const code064 = String(row[rentalColMap["064"] - 1] || "").trim();
    const displayName = String(row[rentalColMap["17"] - 1] || "").trim();

    if (!code064) return;

    SIZE_LIST.forEach(size => {
      const skuCode = `${code064}-${size}`;
      const skuInfo = skuMap.get(skuCode);

      const receiveQty = toRentalMatrixNumber_(row[INPUT_COLS.receive[size] - 1]);
      const lendQty = toRentalMatrixNumber_(row[INPUT_COLS.lend[size] - 1]);
      const outQty = toRentalMatrixNumber_(row[INPUT_COLS.out[size] - 1]);

      if ((receiveQty > 0 || lendQty > 0 || outQty > 0) && !skuInfo) {
        errors.push(`SKU未発見：${skuCode}`);
        return;
      }

      // ① レンタル受け入れ：本部払い出し + / レンタル受け入れ +
      if (receiveQty > 0) {
        if (skuInfo.hqCurrent < receiveQty) {
          errors.push(`本部在庫不足：${skuCode} / 現在 ${skuInfo.hqCurrent} / 受け入れ ${receiveQty}`);
        } else {
          skuInfo.hqPayout += receiveQty;
          skuInfo.rentalReceive += receiveQty;

          // 同一実行内の判定用に現在庫も仮更新
          skuInfo.hqCurrent -= receiveQty;
          skuInfo.rentalCurrent += receiveQty;

          skuInfo.changed = true;

          logRows.push(buildRentalMatrixLogRow_(logColMap, {
            processId, now, source: PROCESS_SOURCE,
            processType: "出庫",
            adjustmentType: "レンタル払い出し",
            skuCode, code064, size, displayName,
            changeQty: -receiveQty
          }));

          logRows.push(buildRentalMatrixLogRow_(logColMap, {
            processId, now, source: PROCESS_SOURCE,
            processType: "入庫",
            adjustmentType: "レンタル受け入れ",
            skuCode, code064, size, displayName,
            changeQty: receiveQty
          }));

          clearRanges.push(rentalSheet.getRange(sheetRow, INPUT_COLS.receive[size]));
        }
      }

      // ② 貸出累計：レンタル貸出累計 + ※在庫は減らさない
      if (lendQty > 0) {
        skuInfo.rentalLend += lendQty;
        skuInfo.changed = true;

        logRows.push(buildRentalMatrixLogRow_(logColMap, {
          processId, now, source: PROCESS_SOURCE,
          processType: "貸出",
          adjustmentType: "レンタル貸出",
          skuCode, code064, size, displayName,
          changeQty: lendQty
        }));

        clearRanges.push(rentalSheet.getRange(sheetRow, INPUT_COLS.lend[size]));
      }

      // ③ その他出庫：レンタルその他出庫 +
      if (outQty > 0) {
        if (skuInfo.rentalCurrent < outQty) {
          errors.push(`レンタル在庫不足：${skuCode} / 現在 ${skuInfo.rentalCurrent} / その他出庫 ${outQty}`);
        } else {
          skuInfo.rentalOut += outQty;

          // 同一実行内の判定用に現在庫も仮更新
          skuInfo.rentalCurrent -= outQty;

          skuInfo.changed = true;

          logRows.push(buildRentalMatrixLogRow_(logColMap, {
            processId, now, source: PROCESS_SOURCE,
            processType: "出庫",
            adjustmentType: "レンタルその他出庫",
            skuCode, code064, size, displayName,
            changeQty: -outQty
          }));

          clearRanges.push(rentalSheet.getRange(sheetRow, INPUT_COLS.out[size]));
        }
      }
    });
  });

  if (errors.length > 0) {
    Browser.msgBox(
      "反映を中止しました。\n\n" +
      errors.slice(0, 20).join("\n") +
      (errors.length > 20 ? `\n...他 ${errors.length - 20} 件` : "")
    );
    return;
  }

  if (logRows.length === 0) {
    Browser.msgBox("反映する入力値がありませんでした。");
    return;
  }

  // SKUシート更新：現在庫列には書かない。累計列だけ更新。
  skuMap.forEach(info => {
    if (!info.changed) return;

    skuSheet.getRange(info.rowNumber, SKU_COL_HQ_PAYOUT).setValue(info.hqPayout);
    skuSheet.getRange(info.rowNumber, SKU_COL_RENTAL_RECEIVE).setValue(info.rentalReceive);
    skuSheet.getRange(info.rowNumber, SKU_COL_RENTAL_OUT).setValue(info.rentalOut);
    skuSheet.getRange(info.rowNumber, SKU_COL_RENTAL_LEND).setValue(info.rentalLend);
  });

  // SKUログ追記
  const logStartRow = getRentalMatrixNextDataRow_(logSheet, DATA_START_ROW);

  logSheet.getRange(
    logStartRow,
    1,
    logRows.length,
    logSheet.getLastColumn()
  ).setValues(logRows);

  // 入力欄クリア
  clearRanges.forEach(range => range.clearContent());

  try {
    SpreadsheetApp.getUi().alert(
      "レンタルマトリックス反映完了\n\n" +
      "処理番号：" + processId + "\n" +
      "ログ件数：" + logRows.length + "件"
    );
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "レンタルマトリックス反映完了",
      "完了",
      10
    );
  }
}


/*******************************************************
 * ヘルパー
 *******************************************************/

function getRentalMatrixColMap_(sheet, headerRow) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};

  headers.forEach((header, i) => {
    const text = String(header || "")
      .trim()
      .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
      .replace(/＿/g, "_");

    const match = text.match(/^(\d{2,4})_/);
    if (match) map[match[1]] = i + 1;
  });

  return map;
}

function toRentalMatrixNumber_(value) {
  if (value === "" || value === null || value === undefined) return 0;
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

function createRentalMatrixProcessId_(logSheet, processIdCol, prefix) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const base = `${prefix}-${today}-`;

  const lastRow = logSheet.getLastRow();
  let maxNo = 0;

  if (lastRow >= 7) {
    const values = logSheet.getRange(7, processIdCol, lastRow - 6, 1).getValues();

    values.forEach(row => {
      const id = String(row[0] || "").trim();
      if (!id.startsWith(base)) return;

      const num = Number(id.replace(base, ""));
      if (!isNaN(num)) maxNo = Math.max(maxNo, num);
    });
  }

  return `${base}${String(maxNo + 1).padStart(4, "0")}`;
}

function getRentalMatrixNextDataRow_(sheet, dataStartRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return dataStartRow;

  const values = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || "").trim() !== "") {
      return dataStartRow + i + 1;
    }
  }

  return dataStartRow;
}

function buildRentalMatrixLogRow_(logColMap, data) {
  const row = [];

  Object.values(logColMap).forEach(col => {
    row[col - 1] = "";
  });

  row[logColMap["01"] - 1] = data.processId;
  row[logColMap["02"] - 1] = data.now;
  row[logColMap["03"] - 1] = data.source;
  row[logColMap["04"] - 1] = data.processType;
  row[logColMap["05"] - 1] = data.adjustmentType;
  row[logColMap["061"] - 1] = data.skuCode;
  row[logColMap["064"] - 1] = data.code064;
  row[logColMap["09"] - 1] = data.size;
  row[logColMap["17"] - 1] = data.displayName;
  row[logColMap["10"] - 1] = data.changeQty;
  row[logColMap["11"] - 1] = "";
  row[logColMap["12"] - 1] = "";

  return row;
}