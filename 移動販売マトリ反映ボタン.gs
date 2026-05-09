/*******************************************************
 * 移動販売マトリックス反映ボタン V1
 *
 * 対象：
 * - 持出：倉庫在庫 - / 移動販売在庫 +
 * - 販売：移動販売在庫 -
 * - 不良廃棄その他：移動販売在庫 -
 *
 * 対象外：
 * - 棚卸
 *******************************************************/

function submitMobileSalesMatrixV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SHEET_MOBILE = "移動販売";
  const SHEET_SKU = "SKU";
  const SHEET_SKU_LOG = "SKUログ";

  const HEADER_ROW = 6;
  const DATA_START_ROW = 7;

  const PROCESS_SOURCE = "MV";

  // SKUシート固定列
  const SKU_COL_CODE = 1;      // 061_SKUコードがA列想定
  const SKU_COL_WH_STOCK = 34; // 倉庫在庫
  const SKU_COL_MV_STOCK = 42; // 移動販売在庫 AP列

  const SIZE_LIST = ["XS", "S", "M", "L", "XL", "F"];

  // 移動販売シート入力列
  const INPUT_COLS = {
    carryOut: { XS: 41, S: 42, M: 43, L: 44, XL: 45, F: 46 }, // AO:AT 持出
    sale:     { XS: 48, S: 49, M: 50, L: 51, XL: 52, F: 53 }, // AV:BA 販売
    defect:   { XS: 55, S: 56, M: 57, L: 58, XL: 59, F: 60 }  // BC:BH 不良廃棄その他
  };

  const mobileSheet = ss.getSheetByName(SHEET_MOBILE);
  const skuSheet = ss.getSheetByName(SHEET_SKU);
  const logSheet = ss.getSheetByName(SHEET_SKU_LOG);

  if (!mobileSheet || !skuSheet || !logSheet) {
    Browser.msgBox("移動販売 / SKU / SKUログ のいずれかのシートが見つかりません。");
    return;
  }

  const result = Browser.msgBox(
    "移動販売マトリックス反映",
    "持出・販売・不良廃棄その他をSKUへ反映します。\n棚卸は処理しません。\n\n実行しますか？",
    Browser.Buttons.YES_NO
  );

  if (result !== "yes") return;

  const mvColMap = getMvColMap_(mobileSheet, HEADER_ROW);
  const logColMap = getMvColMap_(logSheet, HEADER_ROW);

  if (!mvColMap["064"] || !mvColMap["17"]) {
    Browser.msgBox("移動販売シートに 064_ または 17_ が見つかりません。");
    return;
  }

  const requiredLogIds = ["01", "02", "03", "04", "05", "061", "064", "09", "17", "10", "11", "12"];
  const missingLogIds = requiredLogIds.filter(id => !logColMap[id]);
  if (missingLogIds.length > 0) {
    Browser.msgBox("SKUログに必要な項目IDがありません：\n" + missingLogIds.join(", "));
    return;
  }

  const lastMvRow = mobileSheet.getLastRow();
  if (lastMvRow < DATA_START_ROW) {
    Browser.msgBox("移動販売シートに処理対象データがありません。");
    return;
  }

  const mvValues = mobileSheet.getRange(
    DATA_START_ROW,
    1,
    lastMvRow - DATA_START_ROW + 1,
    mobileSheet.getLastColumn()
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
      warehouseStock: Number(row[SKU_COL_WH_STOCK - 1] || 0),
      mobileStock: Number(row[SKU_COL_MV_STOCK - 1] || 0)
    });
  });

  const processId = createMvProcessId_(logSheet, logColMap["01"], PROCESS_SOURCE);
  const now = new Date();

  const logRows = [];
  const clearRanges = [];
  const errors = [];

  mvValues.forEach((row, rowIndex) => {
    const sheetRow = DATA_START_ROW + rowIndex;

    const code064 = String(row[mvColMap["064"] - 1] || "").trim();
    const displayName = String(row[mvColMap["17"] - 1] || "").trim();

    if (!code064) return;

    SIZE_LIST.forEach(size => {
      const skuCode = `${code064}-${size}`;
      const skuInfo = skuMap.get(skuCode);

      // ① 持出：倉庫 - / 移動販売 +
      const carryQty = mvToNumber_(row[INPUT_COLS.carryOut[size] - 1]);
      if (carryQty > 0) {
        if (!skuInfo) {
          errors.push(`SKU未発見：${skuCode}`);
        } else if (skuInfo.warehouseStock < carryQty) {
          errors.push(`倉庫在庫不足：${skuCode} / 現在 ${skuInfo.warehouseStock} / 持出 ${carryQty}`);
        } else {
          skuInfo.warehouseStock -= carryQty;
          skuInfo.mobileStock += carryQty;

          logRows.push(buildMvLogRow_(logColMap, {
            processId, now,
            source: PROCESS_SOURCE,
            processType: "出庫",
            adjustmentType: "移動販売持出_倉庫",
            skuCode, code064, size, displayName,
            changeQty: -carryQty
          }));

          logRows.push(buildMvLogRow_(logColMap, {
            processId, now,
            source: PROCESS_SOURCE,
            processType: "入庫",
            adjustmentType: "移動販売持出_移動販売",
            skuCode, code064, size, displayName,
            changeQty: carryQty
          }));

          clearRanges.push(mobileSheet.getRange(sheetRow, INPUT_COLS.carryOut[size]));
        }
      }

      // ② 販売：移動販売 -
      const saleQty = mvToNumber_(row[INPUT_COLS.sale[size] - 1]);
      if (saleQty > 0) {
        if (!skuInfo) {
          errors.push(`SKU未発見：${skuCode}`);
        } else if (skuInfo.mobileStock < saleQty) {
          errors.push(`移動販売在庫不足：${skuCode} / 現在 ${skuInfo.mobileStock} / 販売 ${saleQty}`);
        } else {
          skuInfo.mobileStock -= saleQty;

          logRows.push(buildMvLogRow_(logColMap, {
            processId, now,
            source: PROCESS_SOURCE,
            processType: "出庫",
            adjustmentType: "移動販売販売",
            skuCode, code064, size, displayName,
            changeQty: -saleQty
          }));

          clearRanges.push(mobileSheet.getRange(sheetRow, INPUT_COLS.sale[size]));
        }
      }

      // ③ 不良廃棄その他：移動販売 -
      const defectQty = mvToNumber_(row[INPUT_COLS.defect[size] - 1]);
      if (defectQty > 0) {
        if (!skuInfo) {
          errors.push(`SKU未発見：${skuCode}`);
        } else if (skuInfo.mobileStock < defectQty) {
          errors.push(`移動販売在庫不足：${skuCode} / 現在 ${skuInfo.mobileStock} / 不良廃棄その他 ${defectQty}`);
        } else {
          skuInfo.mobileStock -= defectQty;

          logRows.push(buildMvLogRow_(logColMap, {
            processId, now,
            source: PROCESS_SOURCE,
            processType: "出庫",
            adjustmentType: "移動販売不良廃棄その他",
            skuCode, code064, size, displayName,
            changeQty: -defectQty
          }));

          clearRanges.push(mobileSheet.getRange(sheetRow, INPUT_COLS.defect[size]));
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

  // SKU更新
  skuMap.forEach(info => {
    skuSheet.getRange(info.rowNumber, SKU_COL_WH_STOCK).setValue(info.warehouseStock);
    skuSheet.getRange(info.rowNumber, SKU_COL_MV_STOCK).setValue(info.mobileStock);
  });

  // SKUログ追記
  const logStartRow = getMvNextDataRow_(logSheet, DATA_START_ROW);
  logSheet.getRange(
    logStartRow,
    1,
    logRows.length,
    logSheet.getLastColumn()
  ).setValues(logRows);

  // 入力欄クリア
  clearRanges.forEach(r => r.clearContent());

  Browser.msgBox(
    "移動販売マトリックス反映完了！\n\n" +
    "処理番号：" + processId + "\n" +
    "ログ件数：" + logRows.length + "件"
  );
}


/*******************************************************
 * ヘルパー
 *******************************************************/

function getMvColMap_(sheet, headerRow) {
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

function mvToNumber_(value) {
  if (value === "" || value === null || value === undefined) return 0;
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

function createMvProcessId_(logSheet, processIdCol, prefix) {
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

function getMvNextDataRow_(sheet, dataStartRow) {
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

function buildMvLogRow_(logColMap, data) {
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