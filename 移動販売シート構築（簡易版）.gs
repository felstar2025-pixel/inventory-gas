/*******************************************************
 * 4.3.1 移動販売シートの構築【PrMASTER連動・6行目固定・商品リスト順】
 *
 * 重要：
 * - 6行目は手動管理。GASは6行目を一切上書きしない。
 * - GASは6行目の項目IDを読んで、7行目以降だけを再構築する。
 *
 * 取得元：
 * - VaMASTER：元の登録情報
 * - PrMASTER：商品リストで登録・加工した販売情報
 * - SKU：数量・在庫
 *
 * 今回の在庫運用：
 * - 50 = 表示対象判定。一度でも会社として受け入れた商品
 * - 52 = 倉庫現在庫。移動販売はこの倉庫在庫を見て販売する
 * - 51 = 倉庫累計払い出し。今回は使わない
 * - 58 = 移動販売累計受け入れ。今回は使わない
 * - 60 = 移動販売現在庫。今回は使わない
 *
 * 並び順：
 * - 50_倉庫累計入庫 と 52_倉庫現在庫 で状態を作る
 * - 1: 倉庫現在庫あり（52 > 0）
 * - 2: 倉庫入庫済み・現在庫なし（50 > 0 かつ 52 = 0）
 * - その中で VN/CN → 型番 → BC優先 → バリエーション → 064
 *
 * 主関数：
 * - generateMobileSalesMatrix_PrMASTER_FixedHeader()
 *******************************************************/

const MOBILE_PR_FIXED_CONFIG = {
  SHEETS: {
    SKU: "SKU",
    VAMASTER: "VaMASTER",
    PRMASTER: "PrMASTER",
    PRODUCT_LIST: "商品リスト",
    TARGET: "移動販売"
  },

  HEADER_ROW: 6,
  DATA_START_ROW: 7,

  SIZE_ORDER: ["XS", "S", "M", "L", "XL", "F"],

  LEFT_ITEM_IDS: [
    "064", "01", "02", "04", "05", "07", "10", "12",
    "09", "13", "17", "23", "24", "4021", "4022", "1001"
  ],

  PR_SAVE_ALLOWED_IDS: [
    "05", "17", "1000", "1001", "1002", "21", "20",
    "23", "24", "4021", "4022", "080", "081",
    "4108", "4109", "4110", "4111", "4112", "4113", "4114", "4115"
  ],

  SKU_IDS: {
    SKU_CODE: ["061"],
    SKU_064: ["064"],
    RECEIVED: ["50"], // 倉庫累計入庫。表示対象判定・並び順判定に使う
    WAREHOUSE_STOCK: ["52"],
    WAREHOUSE_PAYOUT: ["51"],
    MOBILE_RECEIVED: ["58"],
    MOBILE_STOCK: ["60"]
  },

  MATRIX: {
    WAREHOUSE: { name: "倉庫現在庫",     startCol: 35, totalCol: 41, startLetter: "AI", endLetter: "AN", totalLetter: "AO", input: false },
    MOBILE:    { name: "移動販売現在庫", startCol: 42, totalCol: 48, startLetter: "AP", endLetter: "AU", totalLetter: "AV", input: false },
    RECEIVE:   { name: "受け入れ",       startCol: 49, totalCol: 55, startLetter: "AW", endLetter: "BB", totalLetter: "BC", input: true  },
    SALES:     { name: "販売",           startCol: 56, totalCol: 62, startLetter: "BD", endLetter: "BI", totalLetter: "BJ", input: true  },
    LOSS:      { name: "不良廃棄その他", startCol: 63, totalCol: 69, startLetter: "BK", endLetter: "BP", totalLetter: "BQ", input: true  },
    STOCKTAKE: { name: "棚卸",           startCol: 70, totalCol: 76, startLetter: "BR", endLetter: "BW", totalLetter: "BX", input: true  }
  },

  TOTAL_COLS: 76
};


/*******************************************************
 * メイン
 *******************************************************/
function generateMobileSalesMatrix_PrMASTER_FixedHeader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const skuSheet = ss.getSheetByName(MOBILE_PR_FIXED_CONFIG.SHEETS.SKU);
  const vaSheet = ss.getSheetByName(MOBILE_PR_FIXED_CONFIG.SHEETS.VAMASTER);
  const prSheet = ss.getSheetByName(MOBILE_PR_FIXED_CONFIG.SHEETS.PRMASTER);
  const targetSheet = ss.getSheetByName(MOBILE_PR_FIXED_CONFIG.SHEETS.TARGET);

  if (!skuSheet || !vaSheet || !prSheet || !targetSheet) {
    mobilePrFixedAlert_("SKU / VaMASTER / PrMASTER / 移動販売 のいずれかのシートが見つかりません。");
    return;
  }

  const productListSheet = ss.getSheetByName(MOBILE_PR_FIXED_CONFIG.SHEETS.PRODUCT_LIST);
  let savedCount = 0;
  if (productListSheet) {
    savedCount = syncProductListToPrMaster_MobilePrFixed_(productListSheet, prSheet);
  }

  const skuColMap = getMobilePrFixedColMap_(skuSheet);
  const vaColMap = getMobilePrFixedColMap_(vaSheet);
  const prColMap = getMobilePrFixedColMap_(prSheet);
  const targetColMap = getMobilePrFixedColMap_(targetSheet);

  const errors = [];
  if (!targetColMap["064"]) errors.push("移動販売シート6行目に 064_ がありません。");
  if (!vaColMap["064"]) errors.push("VaMASTERに 064_ がありません。");
  if (!findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.RECEIVED)) errors.push("SKUに 50_ がありません。");
  if (
    !findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.SKU_CODE) &&
    !findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.SKU_064)
  ) {
    errors.push("SKUに 061_完全SKUコード または 064_ がありません。");
  }

  if (errors.length) {
    mobilePrFixedAlert_("必要な項目IDが不足しています。\n\n" + errors.join("\n"));
    return;
  }

  const inputBackup = backupMobilePrFixedInputs_(targetSheet, targetColMap);
  const prMap = readPrMasterMap_MobilePrFixed_(prSheet, prColMap);
  const received064Set = buildReceived064SetFromSku_MobilePrFixed_(skuSheet, skuColMap);
  const warehouseIntakeTotalMap = buildWarehouseIntakeTotalMap_MobilePrFixed_(skuSheet, skuColMap);
  const warehouseStockTotalMap = buildWarehouseStockTotalMap_MobilePrFixed_(skuSheet, skuColMap);
  const sizeMap = buildVaSizeMap_MobilePrFixed_(vaSheet, vaColMap);

  const output = buildMobileOutput_MobilePrFixed_(
    vaSheet,
    vaColMap,
    prMap,
    received064Set,
    warehouseIntakeTotalMap,
    warehouseStockTotalMap,
    sizeMap,
    targetColMap,
    inputBackup
  );

  if (output.values.length === 0) {
    mobilePrFixedAlert_("移動販売へ展開する対象がありませんでした。");
    return;
  }

  writeMobilePrFixedBody_(targetSheet, output);
  applyMobilePrFixedMatrixFormulas_(targetSheet, skuSheet, skuColMap, targetColMap, output.values.length);
  applyMobilePrFixedMatrixFormatting_(targetSheet, output.backgrounds, output.values.length);
  applyMobilePrFixedGroups_(targetSheet);

  mobilePrFixedAlert_(
    "移動販売シート構築完了。\n\n" +
    "構築行数：" + output.values.length + "行\n" +
    "PrMASTER保存：" + savedCount + "行\n\n" +
    "6行目は上書きせず、7行目以降だけ更新しました。\n" +
    "並び順は 50_倉庫累計入庫 / 52_倉庫現在庫 を基準にしています。"
  );
}


/*******************************************************
 * 商品リスト → PrMASTER 保存
 *******************************************************/
function syncProductListToPrMaster_MobilePrFixed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productListSheet = ss.getSheetByName(MOBILE_PR_FIXED_CONFIG.SHEETS.PRODUCT_LIST);
  const prSheet = ss.getSheetByName(MOBILE_PR_FIXED_CONFIG.SHEETS.PRMASTER);

  if (!productListSheet || !prSheet) {
    mobilePrFixedAlert_("商品リスト または PrMASTER が見つかりません。");
    return;
  }

  const count = syncProductListToPrMaster_MobilePrFixed_(productListSheet, prSheet);
  mobilePrFixedAlert_("PrMASTERへ保存しました：" + count + "行");
}

function syncProductListToPrMaster_MobilePrFixed_(productListSheet, prSheet) {
  const plColMap = getMobilePrFixedColMap_(productListSheet);
  const prColMap = getMobilePrFixedColMap_(prSheet);

  if (!plColMap["064"] || !prColMap["064"]) return 0;

  const allowed = new Set(MOBILE_PR_FIXED_CONFIG.PR_SAVE_ALLOWED_IDS);
  const plLastRow = productListSheet.getLastRow();
  if (plLastRow < MOBILE_PR_FIXED_CONFIG.DATA_START_ROW) return 0;

  const plValues = productListSheet.getRange(
    MOBILE_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    plLastRow - MOBILE_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    productListSheet.getLastColumn()
  ).getValues();

  const prLastRow = prSheet.getLastRow();
  const prLastCol = Math.max(prSheet.getLastColumn(), 1);

  const prRows = prLastRow >= MOBILE_PR_FIXED_CONFIG.DATA_START_ROW
    ? prSheet.getRange(
        MOBILE_PR_FIXED_CONFIG.DATA_START_ROW,
        1,
        prLastRow - MOBILE_PR_FIXED_CONFIG.DATA_START_ROW + 1,
        prLastCol
      ).getValues()
    : [];

  const rowIndexBy064 = new Map();

  prRows.forEach((row, idx) => {
    const key = String(row[prColMap["064"] - 1] || "").trim();
    if (key) rowIndexBy064.set(key, idx);
  });

  let saved = 0;

  plValues.forEach(plRow => {
    const key064 = String(plRow[plColMap["064"] - 1] || "").trim();
    if (!key064) return;

    let idx;
    if (rowIndexBy064.has(key064)) {
      idx = rowIndexBy064.get(key064);
    } else {
      prRows.push(new Array(prLastCol).fill(""));
      idx = prRows.length - 1;
      rowIndexBy064.set(key064, idx);
    }

    const prRow = prRows[idx];
    prRow[prColMap["064"] - 1] = key064;

    Object.keys(prColMap).forEach(id => {
      if (id === "064") return;
      if (!allowed.has(id)) return;
      if (!plColMap[id]) return;
      prRow[prColMap[id] - 1] = plRow[plColMap[id] - 1];
    });

    saved++;
  });

  if (prRows.length > 0) {
    ensureMobilePrFixedRows_(prSheet, MOBILE_PR_FIXED_CONFIG.DATA_START_ROW + prRows.length - 1);
    prSheet
      .getRange(MOBILE_PR_FIXED_CONFIG.DATA_START_ROW, 1, prRows.length, prLastCol)
      .setValues(prRows);
  }

  return saved;
}


/*******************************************************
 * 出力データ作成
 *******************************************************/
function buildMobileOutput_MobilePrFixed_(vaSheet, vaColMap, prMap, received064Set, warehouseIntakeTotalMap, warehouseStockTotalMap, sizeMap, targetColMap, inputBackup) {
  const lastRow = vaSheet.getLastRow();
  const lastCol = vaSheet.getLastColumn();

  if (lastRow < MOBILE_PR_FIXED_CONFIG.DATA_START_ROW) {
    return { values: [], backgrounds: [] };
  }

  const values = vaSheet.getRange(
    MOBILE_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - MOBILE_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    lastCol
  ).getValues();

  const rows = [];
  const seen = new Set();

  values.forEach(vaRow => {
    const key064 = String(vaRow[vaColMap["064"] - 1] || "").trim();
    if (!key064) return;
    if (seen.has(key064)) return;
    seen.add(key064);

    if (!received064Set.has(key064)) return;

    rows.push({
      key064,
      vaRow,
      pr: prMap.get(key064) || {},
      sizes: sizeMap.get(key064) || new Set(),
      intakeTotal: warehouseIntakeTotalMap.get(key064) || 0,
      whTotal: warehouseStockTotalMap.get(key064) || 0
    });
  });

  // 移動販売用の並び順
  // 1. 倉庫現在庫あり：52 > 0
  // 2. 倉庫入庫済み・現在庫なし：50 > 0 かつ 52 = 0
  // その中で、商品リスト方式の VN/CN → 型番 → BC優先 → バリエーション → 064。
  rows.sort((a, b) => {
    const stockSortA = getMobileWarehouseStatusSortNo_(a);
    const stockSortB = getMobileWarehouseStatusSortNo_(b);
    if (stockSortA !== stockSortB) return stockSortA - stockSortB;

    const countryA = getMobilePrFixedCountrySortNo_(getVaValueById_MobilePrFixed_(a.vaRow, vaColMap, "13"));
    const countryB = getMobilePrFixedCountrySortNo_(getVaValueById_MobilePrFixed_(b.vaRow, vaColMap, "13"));
    if (countryA !== countryB) return countryA - countryB;

    const modelA = getMobilePrFixedModelCode_(a.key064, a.vaRow, vaColMap);
    const modelB = getMobilePrFixedModelCode_(b.key064, b.vaRow, vaColMap);
    if (modelA !== modelB) return modelA.localeCompare(modelB, "ja");

    const supplierA = getMobilePrFixedSupplierCodeFrom064_(a.key064);
    const supplierB = getMobilePrFixedSupplierCodeFrom064_(b.key064);
    const supplierSortA = getMobilePrFixedSupplierSortNo_(supplierA);
    const supplierSortB = getMobilePrFixedSupplierSortNo_(supplierB);
    if (supplierSortA !== supplierSortB) return supplierSortA - supplierSortB;
    if (supplierA !== supplierB) return supplierA.localeCompare(supplierB, "ja");

    const variationA = getMobilePrFixedVariationCodeFrom064_(a.key064);
    const variationB = getMobilePrFixedVariationCodeFrom064_(b.key064);
    const variationSortA = getMobilePrFixedVariationSortNo_(variationA);
    const variationSortB = getMobilePrFixedVariationSortNo_(variationB);
    if (variationSortA !== variationSortB) return variationSortA - variationSortB;
    if (variationA !== variationB) return variationA.localeCompare(variationB, "ja");

    return a.key064.localeCompare(b.key064, "ja");
  });

  const totalCols = Math.max(
    MOBILE_PR_FIXED_CONFIG.TOTAL_COLS,
    Math.max.apply(null, Object.values(targetColMap))
  );

  const outputValues = [];
  const outputBackgrounds = [];

  rows.forEach(item => {
    const row = new Array(totalCols).fill("");
    const bg = new Array(totalCols).fill(null);

    MOBILE_PR_FIXED_CONFIG.LEFT_ITEM_IDS.forEach(id => {
      const col = targetColMap[id];
      if (!col) return;
      row[col - 1] = getMobileLeftValue_MobilePrFixed_(id, item, vaColMap);
    });

    if (inputBackup.has(item.key064)) {
      const saved = inputBackup.get(item.key064);
      Object.keys(saved).forEach(colStr => {
        const col = Number(colStr);
        if (col >= 1 && col <= row.length) row[col - 1] = saved[colStr];
      });
    }

    Object.values(MOBILE_PR_FIXED_CONFIG.MATRIX).forEach(group => {
      MOBILE_PR_FIXED_CONFIG.SIZE_ORDER.forEach((size, idx) => {
        const col = group.startCol + idx;
        if (!item.sizes.has(size)) bg[col - 1] = "#999999";
      });
      bg[group.totalCol - 1] = "#fff2cc";
    });

    [MOBILE_PR_FIXED_CONFIG.MATRIX.WAREHOUSE, MOBILE_PR_FIXED_CONFIG.MATRIX.MOBILE].forEach(group => {
      for (let c = group.startCol; c <= group.totalCol; c++) {
        if (!bg[c - 1]) bg[c - 1] = "#f3f3f3";
      }
    });

    outputValues.push(row);
    outputBackgrounds.push(bg);
  });

  return { values: outputValues, backgrounds: outputBackgrounds };
}

function getMobileLeftValue_MobilePrFixed_(id, item, vaColMap) {
  if (id === "064") return item.key064;

  if (id === "04") {
    const photoUrl = getPhotoUrlForMobilePrFixed_(item.vaRow, vaColMap, item.pr);
    if (!photoUrl) return "";
    return '=IMAGE("' + escapeMobilePrFixedFormulaText_(getMobilePrFixedImageUrl_(photoUrl)) + '")';
  }

  if (id === "05") {
    return getPhotoUrlForMobilePrFixed_(item.vaRow, vaColMap, item.pr);
  }

  if (["17", "23", "24", "4021", "4022", "1001"].includes(id)) {
    return item.pr[id] !== undefined && item.pr[id] !== "" ? item.pr[id] : "";
  }

  return getVaValueById_MobilePrFixed_(item.vaRow, vaColMap, id);
}


/*******************************************************
 * 書き込み
 *******************************************************/
function writeMobilePrFixedBody_(targetSheet, output) {
  const startRow = MOBILE_PR_FIXED_CONFIG.DATA_START_ROW;
  const lastRow = targetSheet.getLastRow();
  const clearCols = Math.max(targetSheet.getLastColumn(), MOBILE_PR_FIXED_CONFIG.TOTAL_COLS, output.values[0].length);

  if (lastRow >= startRow) {
    const rowsToClear = lastRow - startRow + 1;
    targetSheet.getRange(startRow, 1, rowsToClear, clearCols).clearContent();
    targetSheet.getRange(startRow, 1, rowsToClear, clearCols).setBackground(null);
  }

  ensureMobilePrFixedRows_(targetSheet, startRow + output.values.length - 1);

  targetSheet
    .getRange(startRow, 1, output.values.length, output.values[0].length)
    .setValues(output.values);
}


/*******************************************************
 * 在庫式・合計式
 *******************************************************/
function applyMobilePrFixedMatrixFormulas_(targetSheet, skuSheet, skuColMap, targetColMap, rowCount) {
  const startRow = MOBILE_PR_FIXED_CONFIG.DATA_START_ROW;
  const endRow = startRow + rowCount - 1;
  const keyCol = targetColMap["064"];
  if (!keyCol) return;

  const keyColLetter = getMobilePrFixedColumnLetter_(keyCol);

  Object.values(MOBILE_PR_FIXED_CONFIG.MATRIX).forEach(group => {
    if (group === MOBILE_PR_FIXED_CONFIG.MATRIX.MOBILE) return;

    targetSheet
      .getRange(startRow, group.totalCol)
      .setFormula(
        '=BYROW(' + group.startLetter + startRow + ':' + group.endLetter + endRow +
        ', LAMBDA(row, SUM(row)))'
      );
  });

  const skuKeyCol = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.SKU_CODE);
  const whStockCol = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.WAREHOUSE_STOCK);

  if (skuKeyCol && whStockCol) {
    setStockIndexMatchFormulas_MobilePrFixed_(
      targetSheet,
      MOBILE_PR_FIXED_CONFIG.MATRIX.WAREHOUSE,
      keyColLetter,
      skuKeyCol,
      whStockCol,
      rowCount
    );
  }
}

function setStockIndexMatchFormulas_MobilePrFixed_(targetSheet, group, keyColLetter, skuKeyCol, stockCol, rowCount) {
  const startRow = MOBILE_PR_FIXED_CONFIG.DATA_START_ROW;
  const skuKeyLetter = getMobilePrFixedColumnLetter_(skuKeyCol);
  const stockLetter = getMobilePrFixedColumnLetter_(stockCol);

  const formulas = [];

  for (let i = 0; i < rowCount; i++) {
    const r = startRow + i;
    const row = [];

    MOBILE_PR_FIXED_CONFIG.SIZE_ORDER.forEach(size => {
      row.push(
        '=IF($' + keyColLetter + r + '="","",' +
        'IFERROR(INDEX(SKU!$' + stockLetter + ':$' + stockLetter + ',' +
        'MATCH($' + keyColLetter + r + '&"-' + size + '",SKU!$' + skuKeyLetter + ':$' + skuKeyLetter + ',0)),""))'
      );
    });

    formulas.push(row);
  }

  targetSheet
    .getRange(startRow, group.startCol, rowCount, MOBILE_PR_FIXED_CONFIG.SIZE_ORDER.length)
    .setFormulas(formulas);
}


/*******************************************************
 * 入力中データ退避
 *******************************************************/
function backupMobilePrFixedInputs_(targetSheet, targetColMap) {
  const backup = new Map();
  const lastRow = targetSheet.getLastRow();
  if (lastRow < MOBILE_PR_FIXED_CONFIG.DATA_START_ROW || !targetColMap["064"]) return backup;

  const totalCols = Math.max(targetSheet.getLastColumn(), MOBILE_PR_FIXED_CONFIG.TOTAL_COLS);

  const values = targetSheet.getRange(
    MOBILE_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - MOBILE_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    totalCols
  ).getValues();

  const inputGroups = [
    MOBILE_PR_FIXED_CONFIG.MATRIX.RECEIVE,
    MOBILE_PR_FIXED_CONFIG.MATRIX.SALES,
    MOBILE_PR_FIXED_CONFIG.MATRIX.LOSS,
    MOBILE_PR_FIXED_CONFIG.MATRIX.STOCKTAKE
  ];

  values.forEach(row => {
    const key064 = String(row[targetColMap["064"] - 1] || "").trim();
    if (!key064) return;

    const saved = {};

    inputGroups.forEach(group => {
      for (let c = group.startCol; c < group.totalCol; c++) {
        const value = row[c - 1];
        if (value !== "" && value !== null) saved[c] = value;
      }
    });

    if (Object.keys(saved).length > 0) backup.set(key064, saved);
  });

  return backup;
}


/*******************************************************
 * PrMASTER / SKU / VaMASTER読込
 *******************************************************/
function readPrMasterMap_MobilePrFixed_(prSheet, prColMap) {
  const map = new Map();
  const lastRow = prSheet.getLastRow();
  if (lastRow < MOBILE_PR_FIXED_CONFIG.DATA_START_ROW || !prColMap["064"]) return map;

  const values = prSheet.getRange(
    MOBILE_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - MOBILE_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    prSheet.getLastColumn()
  ).getValues();

  values.forEach(row => {
    const key064 = String(row[prColMap["064"] - 1] || "").trim();
    if (!key064) return;

    const obj = {};
    Object.keys(prColMap).forEach(id => {
      obj[id] = row[prColMap[id] - 1];
    });

    map.set(key064, obj);
  });

  return map;
}

function buildReceived064SetFromSku_MobilePrFixed_(skuSheet, skuColMap) {
  const set = new Set();
  const lastRow = skuSheet.getLastRow();
  if (lastRow < MOBILE_PR_FIXED_CONFIG.DATA_START_ROW) return set;

  const values = skuSheet.getRange(
    MOBILE_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - MOBILE_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    skuSheet.getLastColumn()
  ).getValues();

  const receivedCol = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.RECEIVED);
  const skuKeyCol = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.SKU_CODE);
  const sku064Col = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.SKU_064);

  values.forEach(row => {
    const received = toMobilePrFixedNumber_(row[receivedCol - 1]);
    if (!received || received <= 0) return;

    let key064 = "";
    if (sku064Col) key064 = String(row[sku064Col - 1] || "").trim();
    if (!key064 && skuKeyCol) key064 = convertSku061To064_MobilePrFixed_(row[skuKeyCol - 1]);

    if (key064) set.add(key064);
  });

  return set;
}


function buildWarehouseIntakeTotalMap_MobilePrFixed_(skuSheet, skuColMap) {
  const map = new Map();
  const lastRow = skuSheet.getLastRow();
  if (lastRow < MOBILE_PR_FIXED_CONFIG.DATA_START_ROW) return map;

  const values = skuSheet.getRange(
    MOBILE_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - MOBILE_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    skuSheet.getLastColumn()
  ).getValues();

  const skuKeyCol = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.SKU_CODE);
  const sku064Col = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.SKU_064);
  const intakeCol = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.RECEIVED);

  if (!intakeCol) return map;

  values.forEach(row => {
    let key064 = "";
    if (sku064Col) key064 = String(row[sku064Col - 1] || "").trim();
    if (!key064 && skuKeyCol) key064 = convertSku061To064_MobilePrFixed_(row[skuKeyCol - 1]);
    if (!key064) return;

    const qty = toMobilePrFixedNumber_(row[intakeCol - 1]);
    map.set(key064, (map.get(key064) || 0) + qty);
  });

  return map;
}

function buildWarehouseStockTotalMap_MobilePrFixed_(skuSheet, skuColMap) {
  const map = new Map();
  const lastRow = skuSheet.getLastRow();
  if (lastRow < MOBILE_PR_FIXED_CONFIG.DATA_START_ROW) return map;

  const values = skuSheet.getRange(
    MOBILE_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - MOBILE_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    skuSheet.getLastColumn()
  ).getValues();

  const skuKeyCol = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.SKU_CODE);
  const sku064Col = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.SKU_064);
  const whStockCol = findFirstCol_MobilePrFixed_(skuColMap, MOBILE_PR_FIXED_CONFIG.SKU_IDS.WAREHOUSE_STOCK);

  if (!whStockCol) return map;

  values.forEach(row => {
    let key064 = "";
    if (sku064Col) key064 = String(row[sku064Col - 1] || "").trim();
    if (!key064 && skuKeyCol) key064 = convertSku061To064_MobilePrFixed_(row[skuKeyCol - 1]);
    if (!key064) return;

    const qty = toMobilePrFixedNumber_(row[whStockCol - 1]);
    map.set(key064, (map.get(key064) || 0) + qty);
  });

  return map;
}

function buildVaSizeMap_MobilePrFixed_(vaSheet, vaColMap) {
  const map = new Map();
  const lastRow = vaSheet.getLastRow();
  if (lastRow < MOBILE_PR_FIXED_CONFIG.DATA_START_ROW || !vaColMap["064"]) return map;

  const values = vaSheet.getRange(
    MOBILE_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - MOBILE_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    vaSheet.getLastColumn()
  ).getValues();

  values.forEach(row => {
    const key064 = String(row[vaColMap["064"] - 1] || "").trim();
    if (!key064) return;

    if (!map.has(key064)) map.set(key064, new Set());

    const sizeSet = map.get(key064);
    const raw = vaColMap["09"] ? String(row[vaColMap["09"] - 1] || "") : "";

    raw
      .split(/[,、\n]/)
      .map(v => normalizeSize_MobilePrFixed_(v))
      .filter(Boolean)
      .forEach(size => sizeSet.add(size));
  });

  return map;
}


/*******************************************************
 * 書式・グループ
 *******************************************************/
function applyMobilePrFixedMatrixFormatting_(targetSheet, backgrounds, rowCount) {
  const startRow = MOBILE_PR_FIXED_CONFIG.DATA_START_ROW;

  targetSheet
    .getRange(startRow, 1, rowCount, backgrounds[0].length)
    .setBackgrounds(backgrounds);

  try {
    targetSheet.setRowHeights(startRow, rowCount, 92);
  } catch (e) {}

  const ranges = [
    targetSheet.getRange(startRow, MOBILE_PR_FIXED_CONFIG.MATRIX.WAREHOUSE.startCol, rowCount, 7),
    targetSheet.getRange(startRow, MOBILE_PR_FIXED_CONFIG.MATRIX.MOBILE.startCol, rowCount, 7)
  ];

  Object.values(MOBILE_PR_FIXED_CONFIG.MATRIX).forEach(group => {
    ranges.push(targetSheet.getRange(startRow, group.totalCol, rowCount, 1));
  });

  ranges.forEach(range => {
    try {
      range.protect().setWarningOnly(true);
    } catch (e) {}
  });
}

function applyMobilePrFixedGroups_(targetSheet) {
  const groups = [
    MOBILE_PR_FIXED_CONFIG.MATRIX.MOBILE,
    MOBILE_PR_FIXED_CONFIG.MATRIX.RECEIVE,
    MOBILE_PR_FIXED_CONFIG.MATRIX.STOCKTAKE
  ];

  groups.forEach(group => {
    try {
      targetSheet
        .getRange(1, group.startCol, targetSheet.getMaxRows(), group.totalCol - group.startCol + 1)
        .shiftColumnGroupDepth(1);
      targetSheet.getColumnGroup(group.startCol, 1).collapse();
    } catch (e) {}
  });
}


/*******************************************************
 * ヘルパー
 *******************************************************/
function getMobilePrFixedColMap_(sheet) {
  const map = {};
  const lastCol = Math.max(sheet.getLastColumn(), 1);

  const headers = sheet
    .getRange(MOBILE_PR_FIXED_CONFIG.HEADER_ROW, 1, 1, lastCol)
    .getValues()[0];

  headers.forEach((header, idx) => {
    const text = normalizeMobilePrFixedHeader_(header);
    const match = text.match(/^(\d{2,4})_/);
    if (match) map[match[1]] = idx + 1;
  });

  return map;
}

function normalizeMobilePrFixedHeader_(value) {
  return String(value || "")
    .trim()
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/＿/g, "_")
    .replace(/\s+/g, "")
    .replace(/ｺｰﾄﾞ/g, "コード")
    .replace(/サプライヤ$/g, "サプライヤー")
    .replace(/写真\/URL/g, "写真URL")
    .replace(/SIZE/g, "Size");
}

function findFirstCol_MobilePrFixed_(colMap, ids) {
  for (let i = 0; i < ids.length; i++) {
    if (colMap[ids[i]]) return colMap[ids[i]];
  }
  return null;
}

function getVaValueById_MobilePrFixed_(vaRow, vaColMap, id) {
  if (!vaColMap[id]) return "";
  return vaRow[vaColMap[id] - 1];
}

function getPhotoUrlForMobilePrFixed_(vaRow, vaColMap, pr) {
  if (pr["05"]) return pr["05"];
  if (vaColMap["05"]) return vaRow[vaColMap["05"] - 1];
  return "";
}

function convertSku061To064_MobilePrFixed_(sku) {
  const text = String(sku || "").trim();
  if (!text) return "";

  const parts = text.split("-");
  if (parts.length >= 4) {
    parts.pop();
    return parts.join("-");
  }

  return text;
}

function normalizeSize_MobilePrFixed_(value) {
  const text = String(value || "").trim().toUpperCase();
  if (!text) return "";
  if (["FREE", "FREES", "FREE SIZE", "フリー", "フリーサイズ"].includes(text)) return "F";
  return text;
}

function getMobilePrFixedImageUrl_(url) {
  const text = String(url || "").trim();
  if (!text) return "";

  const match = text.match(/(?:id=|\/d\/)([\w-]+)/);
  if (match) return "https://drive.google.com/uc?export=download&id=" + match[1];

  return text;
}

function escapeMobilePrFixedFormulaText_(value) {
  return String(value || "").replace(/"/g, '""');
}

function getMobilePrFixedColumnLetter_(column) {
  let temp;
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

function ensureMobilePrFixedRows_(sheet, neededLastRow) {
  const maxRows = sheet.getMaxRows();
  if (neededLastRow > maxRows) {
    sheet.insertRowsAfter(maxRows, neededLastRow - maxRows);
  }
}

function toMobilePrFixedNumber_(value) {
  const num = Number(String(value || "").replace(/,/g, ""));
  return isNaN(num) ? 0 : num;
}

function getMobilePrFixedCountrySortNo_(value) {
  const text = String(value || "").trim().toUpperCase();
  if (text === "VN" || text === "VND" || text.includes("ベトナム")) return 1;
  if (text === "CN" || text === "CNY" || text === "RMB" || text.includes("中国")) return 2;
  return 9;
}

function getMobilePrFixedModelCode_(key064, vaRow, vaColMap) {
  if (vaColMap["06"]) {
    const model = String(vaRow[vaColMap["06"] - 1] || "").trim();
    if (model) return model;
  }
  return getMobilePrFixedCodePart_(key064, 0);
}

function getMobilePrFixedCodePart_(key064, index) {
  return String(key064 || "").split("-")[index] || "";
}

function getMobilePrFixedSupplierCodeFrom064_(key064) {
  const parts = String(key064 || "").split("-");
  return parts.length >= 3 ? parts[2] : "";
}

function getMobilePrFixedVariationCodeFrom064_(key064) {
  const parts = String(key064 || "").split("-");
  return parts.length >= 2 ? parts[1] : "";
}

function getMobilePrFixedSupplierSortNo_(supplier) {
  const text = String(supplier || "").trim().toUpperCase();
  if (text === "BC") return 1;
  return 9;
}

function getMobilePrFixedVariationSortNo_(variation) {
  const text = String(variation || "").trim().toUpperCase();

  if (text === "N") return 1;
  if (text === "A") return 2;
  if (text === "B") return 3;
  if (text === "C") return 4;
  if (text === "D") return 5;

  return 99;
}


function getMobileWarehouseStatusSortNo_(item) {
  // 1: 倉庫現在庫あり
  if ((item.whTotal || 0) > 0) return 1;

  // 2: 倉庫入庫済みだが現在庫なし
  if ((item.intakeTotal || 0) > 0) return 2;

  // 通常は表示対象外だが、念のため最後へ
  return 9;
}

function mobilePrFixedAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "移動販売", 8);
  }
}
