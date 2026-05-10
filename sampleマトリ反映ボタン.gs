/*******************************************************
 * サンプルマトリックス反映ボタン V1
 *
 * 対象：
 * - サンプル受け入れ
 * - サンプルその他出庫
 *
 * 対象外：
 * - 棚卸
 *
 * 重要：
 * - サンプル現在庫 AW列には直接書かない
 * - AW列はSKUシート側のARRAYFORMULAで計算する
 *******************************************************/

function submitSampleMatrixV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SHEET_SAMPLE = "sample"; // 日本語シート名なら "サンプル"
  const SHEET_SKU = "SKU";
  const SHEET_SKU_LOG = "SKUログ";

  const HEADER_ROW = 6;
  const DATA_START_ROW = 7;

  const PROCESS_SOURCE = "SP";

  // SKUシート固定列
  const SKU_COL_CODE = 1;             // A列：061_SKUコード

  // 現在庫列：読むだけ。直接書き込まない
  const SKU_COL_HQ_CURRENT = 34;      // AH列：本部現在庫
  const SKU_COL_SAMPLE_CURRENT = 49;  // AW列：サンプル現在庫

  // 累計列：ここに加算する
  const SKU_COL_HQ_PAYOUT = 32;       // AF列：本部払い出し
  const SKU_COL_SAMPLE_RECEIVE = 47;  // AU列：サンプル累計入庫
  const SKU_COL_SAMPLE_OUT = 48;      // AV列：サンプル累計出庫

  const SIZE_LIST = ["XS", "S", "M", "L", "XL", "F"];

  // sampleシート入力列
  const INPUT_COLS = {
    receive: { XS: 41, S: 42, M: 43, L: 44, XL: 45, F: 46 }, // AO:AT サンプル受け入れ
    out:     { XS: 48, S: 49, M: 50, L: 51, XL: 52, F: 53 }  // AV:BA サンプルその他出庫
  };

  const sampleSheet = ss.getSheetByName(SHEET_SAMPLE);
  const skuSheet = ss.getSheetByName(SHEET_SKU);
  const logSheet = ss.getSheetByName(SHEET_SKU_LOG);

  if (!sampleSheet || !skuSheet || !logSheet) {
    Browser.msgBox("sample / SKU / SKUログ のいずれかのシートが見つかりません。");
    return;
  }

  const confirm = Browser.msgBox(
    "サンプルマトリックス反映",
    "サンプル受け入れ・その他出庫をSKUへ反映します。\n棚卸は処理しません。\n\n実行しますか？",
    Browser.Buttons.YES_NO
  );

  if (confirm !== "yes") return;

  const sampleColMap = getSampleMatrixColMap_(sampleSheet, HEADER_ROW);
  const logColMap = getSampleMatrixColMap_(logSheet, HEADER_ROW);

  if (!sampleColMap["064"] || !sampleColMap["17"]) {
    Browser.msgBox("sampleシートに 064_ または 17_ が見つかりません。");
    return;
  }

  const requiredLogIds = ["01", "02", "03", "04", "05", "061", "064", "09", "17", "10", "11", "12"];
  const missingLogIds = requiredLogIds.filter(id => !logColMap[id]);

  if (missingLogIds.length > 0) {
    Browser.msgBox("SKUログに必要な項目IDがありません：\n" + missingLogIds.join(", "));
    return;
  }

  const lastSampleRow = sampleSheet.getLastRow();
  if (lastSampleRow < DATA_START_ROW) {
    Browser.msgBox("sampleシートに処理対象データがありません。");
    return;
  }

  const sampleValues = sampleSheet.getRange(
    DATA_START_ROW,
    1,
    lastSampleRow - DATA_START_ROW + 1,
    sampleSheet.getLastColumn()
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
      sampleCurrent: Number(row[SKU_COL_SAMPLE_CURRENT - 1] || 0),

      hqPayout: Number(row[SKU_COL_HQ_PAYOUT - 1] || 0),
      sampleReceive: Number(row[SKU_COL_SAMPLE_RECEIVE - 1] || 0),
      sampleOut: Number(row[SKU_COL_SAMPLE_OUT - 1] || 0),

      changed: false
    });
  });

  const processId = createSampleMatrixProcessId_(logSheet, logColMap["01"], PROCESS_SOURCE);
  const now = new Date();

  const logRows = [];
  const clearRanges = [];
  const errors = [];

  sampleValues.forEach((row, rowIndex) => {
    const sheetRow = DATA_START_ROW + rowIndex;

    const code064 = String(row[sampleColMap["064"] - 1] || "").trim();
    const displayName = String(row[sampleColMap["17"] - 1] || "").trim();

    if (!code064) return;

    SIZE_LIST.forEach(size => {
      const skuCode = `${code064}-${size}`;
      const skuInfo = skuMap.get(skuCode);

      const receiveQty = toSampleMatrixNumber_(row[INPUT_COLS.receive[size] - 1]);
      const outQty = toSampleMatrixNumber_(row[INPUT_COLS.out[size] - 1]);

      if ((receiveQty > 0 || outQty > 0) && !skuInfo) {
        errors.push(`SKU未発見：${skuCode}`);
        return;
      }

      // ① サンプル受け入れ：本部払い出し + / サンプル受け入れ +
      if (receiveQty > 0) {
        if (skuInfo.hqCurrent < receiveQty) {
          errors.push(`本部在庫不足：${skuCode} / 現在 ${skuInfo.hqCurrent} / 受け入れ ${receiveQty}`);
        } else {
          skuInfo.hqPayout += receiveQty;
          skuInfo.sampleReceive += receiveQty;

          // 同一実行内の判定用に仮更新
          skuInfo.hqCurrent -= receiveQty;
          skuInfo.sampleCurrent += receiveQty;

          skuInfo.changed = true;

          logRows.push(buildSampleMatrixLogRow_(logColMap, {
            processId, now, source: PROCESS_SOURCE,
            processType: "出庫",
            adjustmentType: "サンプル払い出し",
            skuCode, code064, size, displayName,
            changeQty: -receiveQty
          }));

          logRows.push(buildSampleMatrixLogRow_(logColMap, {
            processId, now, source: PROCESS_SOURCE,
            processType: "入庫",
            adjustmentType: "サンプル受け入れ",
            skuCode, code064, size, displayName,
            changeQty: receiveQty
          }));

          clearRanges.push(sampleSheet.getRange(sheetRow, INPUT_COLS.receive[size]));
        }
      }

      // ② サンプルその他出庫：サンプル出庫 +
      if (outQty > 0) {
        if (skuInfo.sampleCurrent < outQty) {
          errors.push(`サンプル在庫不足：${skuCode} / 現在 ${skuInfo.sampleCurrent} / その他出庫 ${outQty}`);
        } else {
          skuInfo.sampleOut += outQty;

          // 同一実行内の判定用に仮更新
          skuInfo.sampleCurrent -= outQty;

          skuInfo.changed = true;

          logRows.push(buildSampleMatrixLogRow_(logColMap, {
            processId, now, source: PROCESS_SOURCE,
            processType: "出庫",
            adjustmentType: "サンプルその他出庫",
            skuCode, code064, size, displayName,
            changeQty: -outQty
          }));

          clearRanges.push(sampleSheet.getRange(sheetRow, INPUT_COLS.out[size]));
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
    skuSheet.getRange(info.rowNumber, SKU_COL_SAMPLE_RECEIVE).setValue(info.sampleReceive);
    skuSheet.getRange(info.rowNumber, SKU_COL_SAMPLE_OUT).setValue(info.sampleOut);
  });

  // SKUログ追記
  const logStartRow = getSampleMatrixNextDataRow_(logSheet, DATA_START_ROW);

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
      "サンプルマトリックス反映完了\n\n" +
      "処理番号：" + processId + "\n" +
      "ログ件数：" + logRows.length + "件"
    );
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "サンプルマトリックス反映完了",
      "完了",
      10
    );
  }
}


/*******************************************************
 * ヘルパー
 *******************************************************/

function getSampleMatrixColMap_(sheet, headerRow) {
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

function toSampleMatrixNumber_(value) {
  if (value === "" || value === null || value === undefined) return 0;
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

function createSampleMatrixProcessId_(logSheet, processIdCol, prefix) {
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

function getSampleMatrixNextDataRow_(sheet, dataStartRow) {
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

function buildSampleMatrixLogRow_(logColMap, data) {
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