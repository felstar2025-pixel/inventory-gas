/*******************************************************
 * 倉庫 → SKU 反映ボタン V1
 * 
 * 対象：
 * - 入庫
 * - 出庫
 * - 不良廃棄その他
 *
 * 対象外：
 * - 棚卸
 *
 * SKUコード：
 * - 064_P+V+Sコード + "-" + サイズ
 *
 * SKUログ：
 * - 6行目：項目ID
 * - 7行目以降：ログデータ
 *******************************************************/

function submitWarehouseStockDataV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const SHEET_WAREHOUSE = "倉庫";
  const SHEET_SKU = "SKU";
  const SHEET_SKU_LOG = "SKUログ";

  const HEADER_ROW = 6;
  const DATA_START_ROW = 7;

  const PROCESS_SOURCE = "WH"; // Warehouse
  const ADJUSTMENT_TYPE = "通常反映";

  const SIZE_LIST = ["XS", "S", "M", "L", "XL", "F"];

  // 倉庫シート側：各サイズの入力列
  // ※ここは倉庫シートの実際の列位置に合わせて調整してください
  const INPUT_COLS = {
    XS: { in: 34, out: 41, ng: 48 },
    S:  { in: 35, out: 42, ng: 49 },
    M:  { in: 36, out: 43, ng: 50 },
    L:  { in: 37, out: 44, ng: 51 },
    XL: { in: 38, out: 45, ng: 52 },
    F:  { in: 39, out: 46, ng: 53 }
  };

  const whSheet = ss.getSheetByName(SHEET_WAREHOUSE);
  const skuSheet = ss.getSheetByName(SHEET_SKU);
  const logSheet = ss.getSheetByName(SHEET_SKU_LOG);

  if (!whSheet || !skuSheet || !logSheet) {
    ui.alert("必要なシートが見つかりません。倉庫 / SKU / SKUログ を確認してください。");
    return;
  }

  const result = ui.alert(
    "倉庫在庫をSKUへ反映します",
    "入庫・出庫・不良廃棄その他の入力値をSKUシートへ反映し、SKUログへ記録します。\n棚卸はこのボタンでは処理しません。\n\n実行しますか？",
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  const whColMap = getColumnMapByItemId_(whSheet, HEADER_ROW);
  const skuColMap = getColumnMapByItemId_(skuSheet, HEADER_ROW);
  const logColMap = getColumnMapByItemId_(logSheet, HEADER_ROW);

  const requiredWhIds = ["064", "17"];
  const requiredSkuIds = ["061", "50"];
  const requiredLogIds = ["01", "02", "03", "04", "05", "061", "064", "09", "17", "10", "11", "12"];

  checkRequiredIds_(whColMap, requiredWhIds, "倉庫");
  checkRequiredIds_(skuColMap, requiredSkuIds, "SKU");
  checkRequiredIds_(logColMap, requiredLogIds, "SKUログ");

  const lastWhRow = whSheet.getLastRow();
  if (lastWhRow < DATA_START_ROW) {
    ui.alert("倉庫シートに処理対象データがありません。");
    return;
  }

  const whValues = whSheet
    .getRange(DATA_START_ROW, 1, lastWhRow - DATA_START_ROW + 1, whSheet.getLastColumn())
    .getValues();

  const lastSkuRow = skuSheet.getLastRow();
  if (lastSkuRow < DATA_START_ROW) {
    ui.alert("SKUシートにデータがありません。");
    return;
  }

  const skuValues = skuSheet
    .getRange(DATA_START_ROW, 1, lastSkuRow - DATA_START_ROW + 1, skuSheet.getLastColumn())
    .getValues();

  // SKUコード → SKUシート上の行番号・現在在庫
  const skuMap = new Map();

  skuValues.forEach((row, i) => {
    const skuCode = String(row[skuColMap["061"] - 1] || "").trim();
    if (!skuCode) return;

    skuMap.set(skuCode, {
      rowNumber: DATA_START_ROW + i,
      stock: Number(row[skuColMap["50"] - 1] || 0)
    });
  });

  const processId = createProcessId_(logSheet, logColMap["01"], PROCESS_SOURCE);
  const now = new Date();

  const logRows = [];
  const clearRanges = [];
  const stockUpdates = [];

  whValues.forEach((row, rowIndex) => {
    const sheetRow = DATA_START_ROW + rowIndex;

    const code064 = String(row[whColMap["064"] - 1] || "").trim();
    const displayName = String(row[whColMap["17"] - 1] || "").trim();

    if (!code064) return;

    SIZE_LIST.forEach(size => {
      const cols = INPUT_COLS[size];
      if (!cols) return;

      const inQty = toNumber_(row[cols.in - 1]);
      const outQty = toNumber_(row[cols.out - 1]);
      const ngQty = toNumber_(row[cols.ng - 1]);

      const entries = [
        { qty: inQty, type: "入庫", sign: 1, col: cols.in },
        { qty: outQty, type: "出庫", sign: -1, col: cols.out },
        { qty: ngQty, type: "不良廃棄その他", sign: -1, col: cols.ng }
      ];

      entries.forEach(entry => {
        if (!entry.qty || entry.qty <= 0) return;

        const skuCode = `${code064}-${size}`;
        const skuInfo = skuMap.get(skuCode);

        if (!skuInfo) {
          throw new Error(`SKUシートにSKUコードが見つかりません：${skuCode}`);
        }

        const changeQty = entry.qty * entry.sign;
        const beforeStock = Number(skuInfo.stock || 0);
        const afterStock = beforeStock + changeQty;

        skuInfo.stock = afterStock;

        stockUpdates.push({
          rowNumber: skuInfo.rowNumber,
          value: afterStock
        });

        logRows.push(buildSkuLogRow_(logColMap, {
          processId,
          dateTime: now,
          source: PROCESS_SOURCE,
          processType: entry.type,
          adjustmentType: ADJUSTMENT_TYPE,
          skuCode,
          code064,
          size,
          displayName,
          changeQty,
          staff: "",
          memo: ""
        }));

        clearRanges.push(whSheet.getRange(sheetRow, entry.col));
      });
    });
  });

  if (logRows.length === 0) {
    ui.alert("反映する入力値がありませんでした。");
    return;
  }

  // SKU在庫を反映
  stockUpdates.forEach(update => {
    skuSheet.getRange(update.rowNumber, skuColMap["50"]).setValue(update.value);
  });

  // SKUログへ追記
  const logStartRow = getNextDataRow_(logSheet, DATA_START_ROW);
  logSheet
    .getRange(logStartRow, 1, logRows.length, logSheet.getLastColumn())
    .setValues(logRows);

  // 入力セルをクリア
  clearRanges.forEach(range => range.clearContent());

  ui.alert(`完了しました。\n処理番号：${processId}\nログ件数：${logRows.length}件`);
}


/*******************************************************
 * 共通関数
 *******************************************************/

function getColumnMapByItemId_(sheet, headerRow) {
  const headers = sheet
    .getRange(headerRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const map = {};

  headers.forEach((header, index) => {
    const text = String(header || "").trim();
    if (!text) return;

    const id = text.split("_")[0].trim();
    if (id) {
      map[id] = index + 1;
    }
  });

  return map;
}


function checkRequiredIds_(colMap, requiredIds, sheetName) {
  const missing = requiredIds.filter(id => !colMap[id]);

  if (missing.length > 0) {
    throw new Error(`${sheetName}シートに必要な項目IDがありません：${missing.join(", ")}`);
  }
}


function toNumber_(value) {
  if (value === "" || value === null || value === undefined) return 0;

  const num = Number(value);
  if (isNaN(num)) return 0;

  return num;
}


function getNextDataRow_(sheet, dataStartRow) {
  const lastRow = sheet.getLastRow();

  if (lastRow < dataStartRow) return dataStartRow;

  const values = sheet
    .getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 1)
    .getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || "").trim() !== "") {
      return dataStartRow + i + 1;
    }
  }

  return dataStartRow;
}


function createProcessId_(logSheet, processIdCol, prefix) {
  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd"
  );

  const base = `${prefix}-${today}-`;

  const lastRow = logSheet.getLastRow();
  let maxNo = 0;

  if (lastRow >= 7) {
    const values = logSheet
      .getRange(7, processIdCol, lastRow - 6, 1)
      .getValues();

    values.forEach(row => {
      const id = String(row[0] || "").trim();
      if (!id.startsWith(base)) return;

      const num = Number(id.replace(base, ""));
      if (!isNaN(num)) {
        maxNo = Math.max(maxNo, num);
      }
    });
  }

  const nextNo = String(maxNo + 1).padStart(4, "0");

  return `${base}${nextNo}`;
}


function buildSkuLogRow_(logColMap, data) {
  const row = [];

  Object.values(logColMap).forEach(col => {
    row[col - 1] = "";
  });

  row[logColMap["01"] - 1] = data.processId;
  row[logColMap["02"] - 1] = data.dateTime;
  row[logColMap["03"] - 1] = data.source;
  row[logColMap["04"] - 1] = data.processType;
  row[logColMap["05"] - 1] = data.adjustmentType;
  row[logColMap["061"] - 1] = data.skuCode;
  row[logColMap["064"] - 1] = data.code064;
  row[logColMap["09"] - 1] = data.size;
  row[logColMap["17"] - 1] = data.displayName;
  row[logColMap["10"] - 1] = data.changeQty;
  row[logColMap["11"] - 1] = data.staff;
  row[logColMap["12"] - 1] = data.memo;

  return row;
}
