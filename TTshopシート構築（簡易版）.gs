/*******************************************************
 * 4.4.1 TTshopシートの構築【PrMASTER連動・6行目固定・商品リスト順】
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
 * - 52 = 倉庫現在庫。TTshopはこの倉庫在庫を見て販売する
 * - 51 = 倉庫累計払い出し。今回は使わない
 * - 54 = TTshop累計受け入れ。今回は使わない
 * - 56 = TTshop現在庫。今回は使わない
 *
 * 並び順：
 * - 50_倉庫累計入庫 と 52_倉庫現在庫 で状態を作る
 * - 1: 倉庫現在庫あり（52 > 0）
 * - 2: 倉庫入庫済み・現在庫なし（50 > 0 かつ 52 = 0）
 * - その中で VN/CN → 型番 → BC優先 → バリエーション → 064
 *
 * 主関数：
 * - generateTTshopMatrix_PrMASTER_FixedHeader()
 *******************************************************/

const TTSHOP_PR_FIXED_CONFIG = {
  SHEETS: {
    SKU: "SKU",
    VAMASTER: "VaMASTER",
    PRMASTER: "PrMASTER",
    PRODUCT_LIST: "商品リスト",
    TARGET: "TTshop"
  },

  HEADER_ROW: 6,
  DATA_START_ROW: 7,

  SIZE_ORDER: ["XS", "S", "M", "L", "XL", "F"],

  LEFT_ITEM_IDS: [
    "064", "01", "02", "04", "05", "07", "10", "12",
    "09", "13", "17", "23", "24", "4021", "4022", "1000"
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
    TTSHOP_RECEIVED: ["54"],
    TTSHOP_STOCK: ["56"]
  },

  MATRIX: {
    WAREHOUSE: { name: "倉庫現在庫",     startCol: 35, totalCol: 41, startLetter: "AI", endLetter: "AN", totalLetter: "AO", input: false },
    MOBILE:    { name: "TTshop現在庫", startCol: 42, totalCol: 48, startLetter: "AP", endLetter: "AU", totalLetter: "AV", input: false },
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
function generateTTshopMatrix_PrMASTER_FixedHeader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const skuSheet = ss.getSheetByName(TTSHOP_PR_FIXED_CONFIG.SHEETS.SKU);
  const vaSheet = ss.getSheetByName(TTSHOP_PR_FIXED_CONFIG.SHEETS.VAMASTER);
  const prSheet = ss.getSheetByName(TTSHOP_PR_FIXED_CONFIG.SHEETS.PRMASTER);
  const targetSheet = ss.getSheetByName(TTSHOP_PR_FIXED_CONFIG.SHEETS.TARGET);

  if (!skuSheet || !vaSheet || !prSheet || !targetSheet) {
    ttshopPrFixedAlert_("SKU / VaMASTER / PrMASTER / TTshop のいずれかのシートが見つかりません。");
    return;
  }

  const productListSheet = ss.getSheetByName(TTSHOP_PR_FIXED_CONFIG.SHEETS.PRODUCT_LIST);
  let savedCount = 0;
  if (productListSheet) {
    savedCount = syncProductListToPrMaster_TTshopPrFixed_(productListSheet, prSheet);
  }

  const skuColMap = getTTshopPrFixedColMap_(skuSheet);
  const vaColMap = getTTshopPrFixedColMap_(vaSheet);
  const prColMap = getTTshopPrFixedColMap_(prSheet);
  const targetColMap = getTTshopPrFixedColMap_(targetSheet);

  const errors = [];
  if (!targetColMap["064"]) errors.push("TTshopシート6行目に 064_ がありません。");
  if (!vaColMap["064"]) errors.push("VaMASTERに 064_ がありません。");
  if (!findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.RECEIVED)) errors.push("SKUに 50_ がありません。");
  if (
    !findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.SKU_CODE) &&
    !findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.SKU_064)
  ) {
    errors.push("SKUに 061_完全SKUコード または 064_ がありません。");
  }

  if (errors.length) {
    ttshopPrFixedAlert_("必要な項目IDが不足しています。\n\n" + errors.join("\n"));
    return;
  }

  const inputBackup = backupTTshopPrFixedInputs_(targetSheet, targetColMap);
  const prMap = readPrMasterMap_TTshopPrFixed_(prSheet, prColMap);
  const received064Set = buildReceived064SetFromSku_TTshopPrFixed_(skuSheet, skuColMap);
  const warehouseIntakeTotalMap = buildWarehouseIntakeTotalMap_TTshopPrFixed_(skuSheet, skuColMap);
  const warehouseStockTotalMap = buildWarehouseStockTotalMap_TTshopPrFixed_(skuSheet, skuColMap);
  const sizeMap = buildVaSizeMap_TTshopPrFixed_(vaSheet, vaColMap);

  const output = buildTTshopOutput_TTshopPrFixed_(
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
    ttshopPrFixedAlert_("TTshopへ展開する対象がありませんでした。");
    return;
  }

  writeTTshopPrFixedBody_(targetSheet, output);
  applyTTshopPrFixedMatrixFormulas_(targetSheet, skuSheet, skuColMap, targetColMap, output.values.length);
  applyTTshopPrFixedMatrixFormatting_(targetSheet, output.backgrounds, output.values.length);
  applyTTshopPrFixedGroups_(targetSheet);

  ttshopPrFixedAlert_(
    "TTshopシート構築完了。\n\n" +
    "構築行数：" + output.values.length + "行\n" +
    "PrMASTER保存：" + savedCount + "行\n\n" +
    "6行目は上書きせず、7行目以降だけ更新しました。\n" +
    "並び順は 50_倉庫累計入庫 / 52_倉庫現在庫 を基準にしています。"
  );
}


/*******************************************************
 * 商品リスト → PrMASTER 保存
 *******************************************************/
function syncProductListToPrMaster_TTshopPrFixed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productListSheet = ss.getSheetByName(TTSHOP_PR_FIXED_CONFIG.SHEETS.PRODUCT_LIST);
  const prSheet = ss.getSheetByName(TTSHOP_PR_FIXED_CONFIG.SHEETS.PRMASTER);

  if (!productListSheet || !prSheet) {
    ttshopPrFixedAlert_("商品リスト または PrMASTER が見つかりません。");
    return;
  }

  const count = syncProductListToPrMaster_TTshopPrFixed_(productListSheet, prSheet);
  ttshopPrFixedAlert_("PrMASTERへ保存しました：" + count + "行");
}

function syncProductListToPrMaster_TTshopPrFixed_(productListSheet, prSheet) {
  const plColMap = getTTshopPrFixedColMap_(productListSheet);
  const prColMap = getTTshopPrFixedColMap_(prSheet);

  if (!plColMap["064"] || !prColMap["064"]) return 0;

  const allowed = new Set(TTSHOP_PR_FIXED_CONFIG.PR_SAVE_ALLOWED_IDS);
  const plLastRow = productListSheet.getLastRow();
  if (plLastRow < TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW) return 0;

  const plValues = productListSheet.getRange(
    TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    plLastRow - TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    productListSheet.getLastColumn()
  ).getValues();

  const prLastRow = prSheet.getLastRow();
  const prLastCol = Math.max(prSheet.getLastColumn(), 1);

  const prRows = prLastRow >= TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW
    ? prSheet.getRange(
        TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW,
        1,
        prLastRow - TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW + 1,
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
    ensureTTshopPrFixedRows_(prSheet, TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW + prRows.length - 1);
    prSheet
      .getRange(TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW, 1, prRows.length, prLastCol)
      .setValues(prRows);
  }

  return saved;
}


/*******************************************************
 * 出力データ作成
 *******************************************************/
function buildTTshopOutput_TTshopPrFixed_(vaSheet, vaColMap, prMap, received064Set, warehouseIntakeTotalMap, warehouseStockTotalMap, sizeMap, targetColMap, inputBackup) {
  const lastRow = vaSheet.getLastRow();
  const lastCol = vaSheet.getLastColumn();

  if (lastRow < TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW) {
    return { values: [], backgrounds: [] };
  }

  const values = vaSheet.getRange(
    TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW + 1,
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

  // TTshop用の並び順
  // 1. 倉庫現在庫あり：52 > 0
  // 2. 倉庫入庫済み・現在庫なし：50 > 0 かつ 52 = 0
  // その中で、商品リスト方式の VN/CN → 型番 → BC優先 → バリエーション → 064。
  rows.sort((a, b) => {
    const stockSortA = getTTshopWarehouseStatusSortNo_(a);
    const stockSortB = getTTshopWarehouseStatusSortNo_(b);
    if (stockSortA !== stockSortB) return stockSortA - stockSortB;

    const countryA = getTTshopPrFixedCountrySortNo_(getVaValueById_TTshopPrFixed_(a.vaRow, vaColMap, "13"));
    const countryB = getTTshopPrFixedCountrySortNo_(getVaValueById_TTshopPrFixed_(b.vaRow, vaColMap, "13"));
    if (countryA !== countryB) return countryA - countryB;

    const modelA = getTTshopPrFixedModelCode_(a.key064, a.vaRow, vaColMap);
    const modelB = getTTshopPrFixedModelCode_(b.key064, b.vaRow, vaColMap);
    if (modelA !== modelB) return modelA.localeCompare(modelB, "ja");

    const supplierA = getTTshopPrFixedSupplierCodeFrom064_(a.key064);
    const supplierB = getTTshopPrFixedSupplierCodeFrom064_(b.key064);
    const supplierSortA = getTTshopPrFixedSupplierSortNo_(supplierA);
    const supplierSortB = getTTshopPrFixedSupplierSortNo_(supplierB);
    if (supplierSortA !== supplierSortB) return supplierSortA - supplierSortB;
    if (supplierA !== supplierB) return supplierA.localeCompare(supplierB, "ja");

    const variationA = getTTshopPrFixedVariationCodeFrom064_(a.key064);
    const variationB = getTTshopPrFixedVariationCodeFrom064_(b.key064);
    const variationSortA = getTTshopPrFixedVariationSortNo_(variationA);
    const variationSortB = getTTshopPrFixedVariationSortNo_(variationB);
    if (variationSortA !== variationSortB) return variationSortA - variationSortB;
    if (variationA !== variationB) return variationA.localeCompare(variationB, "ja");

    return a.key064.localeCompare(b.key064, "ja");
  });

  const totalCols = Math.max(
    TTSHOP_PR_FIXED_CONFIG.TOTAL_COLS,
    Math.max.apply(null, Object.values(targetColMap))
  );

  const outputValues = [];
  const outputBackgrounds = [];

  rows.forEach(item => {
    const row = new Array(totalCols).fill("");
    const bg = new Array(totalCols).fill(null);

    TTSHOP_PR_FIXED_CONFIG.LEFT_ITEM_IDS.forEach(id => {
      const col = targetColMap[id];
      if (!col) return;
      row[col - 1] = getTTshopLeftValue_TTshopPrFixed_(id, item, vaColMap);
    });

    if (inputBackup.has(item.key064)) {
      const saved = inputBackup.get(item.key064);
      Object.keys(saved).forEach(colStr => {
        const col = Number(colStr);
        if (col >= 1 && col <= row.length) row[col - 1] = saved[colStr];
      });
    }

    Object.values(TTSHOP_PR_FIXED_CONFIG.MATRIX).forEach(group => {
      TTSHOP_PR_FIXED_CONFIG.SIZE_ORDER.forEach((size, idx) => {
        const col = group.startCol + idx;
        if (!item.sizes.has(size)) bg[col - 1] = "#999999";
      });
      bg[group.totalCol - 1] = "#fff2cc";
    });

    [TTSHOP_PR_FIXED_CONFIG.MATRIX.WAREHOUSE, TTSHOP_PR_FIXED_CONFIG.MATRIX.MOBILE].forEach(group => {
      for (let c = group.startCol; c <= group.totalCol; c++) {
        if (!bg[c - 1]) bg[c - 1] = "#f3f3f3";
      }
    });

    outputValues.push(row);
    outputBackgrounds.push(bg);
  });

  return { values: outputValues, backgrounds: outputBackgrounds };
}

function getTTshopLeftValue_TTshopPrFixed_(id, item, vaColMap) {
  if (id === "064") return item.key064;

  if (id === "04") {
    const photoUrl = getPhotoUrlForTTshopPrFixed_(item.vaRow, vaColMap, item.pr);
    if (!photoUrl) return "";
    return '=IMAGE("' + escapeTTshopPrFixedFormulaText_(getTTshopPrFixedImageUrl_(photoUrl)) + '")';
  }

  if (id === "05") {
    return getPhotoUrlForTTshopPrFixed_(item.vaRow, vaColMap, item.pr);
  }

  if (["17", "23", "24", "4021", "4022", "1000"].includes(id)) {
    return item.pr[id] !== undefined && item.pr[id] !== "" ? item.pr[id] : "";
  }

  return getVaValueById_TTshopPrFixed_(item.vaRow, vaColMap, id);
}


/*******************************************************
 * 書き込み
 *******************************************************/
function writeTTshopPrFixedBody_(targetSheet, output) {
  const startRow = TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW;
  const lastRow = targetSheet.getLastRow();
  const clearCols = Math.max(targetSheet.getLastColumn(), TTSHOP_PR_FIXED_CONFIG.TOTAL_COLS, output.values[0].length);

  if (lastRow >= startRow) {
    const rowsToClear = lastRow - startRow + 1;
    targetSheet.getRange(startRow, 1, rowsToClear, clearCols).clearContent();
    targetSheet.getRange(startRow, 1, rowsToClear, clearCols).setBackground(null);
  }

  ensureTTshopPrFixedRows_(targetSheet, startRow + output.values.length - 1);

  targetSheet
    .getRange(startRow, 1, output.values.length, output.values[0].length)
    .setValues(output.values);
}


/*******************************************************
 * 在庫式・合計式
 *******************************************************/
function applyTTshopPrFixedMatrixFormulas_(targetSheet, skuSheet, skuColMap, targetColMap, rowCount) {
  const startRow = TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW;
  const endRow = startRow + rowCount - 1;
  const keyCol = targetColMap["064"];
  if (!keyCol) return;

  const keyColLetter = getTTshopPrFixedColumnLetter_(keyCol);

  Object.values(TTSHOP_PR_FIXED_CONFIG.MATRIX).forEach(group => {
    if (group === TTSHOP_PR_FIXED_CONFIG.MATRIX.MOBILE) return;

    targetSheet
      .getRange(startRow, group.totalCol)
      .setFormula(
        '=BYROW(' + group.startLetter + startRow + ':' + group.endLetter + endRow +
        ', LAMBDA(row, SUM(row)))'
      );
  });

  const skuKeyCol = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.SKU_CODE);
  const whStockCol = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.WAREHOUSE_STOCK);

  if (skuKeyCol && whStockCol) {
    setStockIndexMatchFormulas_TTshopPrFixed_(
      targetSheet,
      TTSHOP_PR_FIXED_CONFIG.MATRIX.WAREHOUSE,
      keyColLetter,
      skuKeyCol,
      whStockCol,
      rowCount
    );
  }
}

function setStockIndexMatchFormulas_TTshopPrFixed_(targetSheet, group, keyColLetter, skuKeyCol, stockCol, rowCount) {
  const startRow = TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW;
  const skuKeyLetter = getTTshopPrFixedColumnLetter_(skuKeyCol);
  const stockLetter = getTTshopPrFixedColumnLetter_(stockCol);

  const formulas = [];

  for (let i = 0; i < rowCount; i++) {
    const r = startRow + i;
    const row = [];

    TTSHOP_PR_FIXED_CONFIG.SIZE_ORDER.forEach(size => {
      row.push(
        '=IF($' + keyColLetter + r + '="","",' +
        'IFERROR(INDEX(SKU!$' + stockLetter + ':$' + stockLetter + ',' +
        'MATCH($' + keyColLetter + r + '&"-' + size + '",SKU!$' + skuKeyLetter + ':$' + skuKeyLetter + ',0)),""))'
      );
    });

    formulas.push(row);
  }

  targetSheet
    .getRange(startRow, group.startCol, rowCount, TTSHOP_PR_FIXED_CONFIG.SIZE_ORDER.length)
    .setFormulas(formulas);
}


/*******************************************************
 * 入力中データ退避
 *******************************************************/
function backupTTshopPrFixedInputs_(targetSheet, targetColMap) {
  const backup = new Map();
  const lastRow = targetSheet.getLastRow();
  if (lastRow < TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW || !targetColMap["064"]) return backup;

  const totalCols = Math.max(targetSheet.getLastColumn(), TTSHOP_PR_FIXED_CONFIG.TOTAL_COLS);

  const values = targetSheet.getRange(
    TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    totalCols
  ).getValues();

  const inputGroups = [
    TTSHOP_PR_FIXED_CONFIG.MATRIX.RECEIVE,
    TTSHOP_PR_FIXED_CONFIG.MATRIX.SALES,
    TTSHOP_PR_FIXED_CONFIG.MATRIX.LOSS,
    TTSHOP_PR_FIXED_CONFIG.MATRIX.STOCKTAKE
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
function readPrMasterMap_TTshopPrFixed_(prSheet, prColMap) {
  const map = new Map();
  const lastRow = prSheet.getLastRow();
  if (lastRow < TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW || !prColMap["064"]) return map;

  const values = prSheet.getRange(
    TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW + 1,
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

function buildReceived064SetFromSku_TTshopPrFixed_(skuSheet, skuColMap) {
  const set = new Set();
  const lastRow = skuSheet.getLastRow();
  if (lastRow < TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW) return set;

  const values = skuSheet.getRange(
    TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    skuSheet.getLastColumn()
  ).getValues();

  const receivedCol = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.RECEIVED);
  const skuKeyCol = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.SKU_CODE);
  const sku064Col = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.SKU_064);

  values.forEach(row => {
    const received = toTTshopPrFixedNumber_(row[receivedCol - 1]);
    if (!received || received <= 0) return;

    let key064 = "";
    if (sku064Col) key064 = String(row[sku064Col - 1] || "").trim();
    if (!key064 && skuKeyCol) key064 = convertSku061To064_TTshopPrFixed_(row[skuKeyCol - 1]);

    if (key064) set.add(key064);
  });

  return set;
}


function buildWarehouseIntakeTotalMap_TTshopPrFixed_(skuSheet, skuColMap) {
  const map = new Map();
  const lastRow = skuSheet.getLastRow();
  if (lastRow < TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW) return map;

  const values = skuSheet.getRange(
    TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    skuSheet.getLastColumn()
  ).getValues();

  const skuKeyCol = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.SKU_CODE);
  const sku064Col = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.SKU_064);
  const intakeCol = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.RECEIVED);

  if (!intakeCol) return map;

  values.forEach(row => {
    let key064 = "";
    if (sku064Col) key064 = String(row[sku064Col - 1] || "").trim();
    if (!key064 && skuKeyCol) key064 = convertSku061To064_TTshopPrFixed_(row[skuKeyCol - 1]);
    if (!key064) return;

    const qty = toTTshopPrFixedNumber_(row[intakeCol - 1]);
    map.set(key064, (map.get(key064) || 0) + qty);
  });

  return map;
}

function buildWarehouseStockTotalMap_TTshopPrFixed_(skuSheet, skuColMap) {
  const map = new Map();
  const lastRow = skuSheet.getLastRow();
  if (lastRow < TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW) return map;

  const values = skuSheet.getRange(
    TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW + 1,
    skuSheet.getLastColumn()
  ).getValues();

  const skuKeyCol = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.SKU_CODE);
  const sku064Col = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.SKU_064);
  const whStockCol = findFirstCol_TTshopPrFixed_(skuColMap, TTSHOP_PR_FIXED_CONFIG.SKU_IDS.WAREHOUSE_STOCK);

  if (!whStockCol) return map;

  values.forEach(row => {
    let key064 = "";
    if (sku064Col) key064 = String(row[sku064Col - 1] || "").trim();
    if (!key064 && skuKeyCol) key064 = convertSku061To064_TTshopPrFixed_(row[skuKeyCol - 1]);
    if (!key064) return;

    const qty = toTTshopPrFixedNumber_(row[whStockCol - 1]);
    map.set(key064, (map.get(key064) || 0) + qty);
  });

  return map;
}

function buildVaSizeMap_TTshopPrFixed_(vaSheet, vaColMap) {
  const map = new Map();
  const lastRow = vaSheet.getLastRow();
  if (lastRow < TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW || !vaColMap["064"]) return map;

  const values = vaSheet.getRange(
    TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW,
    1,
    lastRow - TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW + 1,
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
      .map(v => normalizeSize_TTshopPrFixed_(v))
      .filter(Boolean)
      .forEach(size => sizeSet.add(size));
  });

  return map;
}


/*******************************************************
 * 書式・グループ
 *******************************************************/
function applyTTshopPrFixedMatrixFormatting_(targetSheet, backgrounds, rowCount) {
  const startRow = TTSHOP_PR_FIXED_CONFIG.DATA_START_ROW;

  targetSheet
    .getRange(startRow, 1, rowCount, backgrounds[0].length)
    .setBackgrounds(backgrounds);

  try {
    targetSheet.setRowHeights(startRow, rowCount, 92);
  } catch (e) {}

  const ranges = [
    targetSheet.getRange(startRow, TTSHOP_PR_FIXED_CONFIG.MATRIX.WAREHOUSE.startCol, rowCount, 7),
    targetSheet.getRange(startRow, TTSHOP_PR_FIXED_CONFIG.MATRIX.MOBILE.startCol, rowCount, 7)
  ];

  Object.values(TTSHOP_PR_FIXED_CONFIG.MATRIX).forEach(group => {
    ranges.push(targetSheet.getRange(startRow, group.totalCol, rowCount, 1));
  });

  ranges.forEach(range => {
    try {
      range.protect().setWarningOnly(true);
    } catch (e) {}
  });
}

function applyTTshopPrFixedGroups_(targetSheet) {
  const groups = [
    TTSHOP_PR_FIXED_CONFIG.MATRIX.MOBILE,
    TTSHOP_PR_FIXED_CONFIG.MATRIX.RECEIVE,
    TTSHOP_PR_FIXED_CONFIG.MATRIX.STOCKTAKE
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
function getTTshopPrFixedColMap_(sheet) {
  const map = {};
  const lastCol = Math.max(sheet.getLastColumn(), 1);

  const headers = sheet
    .getRange(TTSHOP_PR_FIXED_CONFIG.HEADER_ROW, 1, 1, lastCol)
    .getValues()[0];

  headers.forEach((header, idx) => {
    const text = normalizeTTshopPrFixedHeader_(header);
    const match = text.match(/^(\d{2,4})_/);
    if (match) map[match[1]] = idx + 1;
  });

  return map;
}

function normalizeTTshopPrFixedHeader_(value) {
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

function findFirstCol_TTshopPrFixed_(colMap, ids) {
  for (let i = 0; i < ids.length; i++) {
    if (colMap[ids[i]]) return colMap[ids[i]];
  }
  return null;
}

function getVaValueById_TTshopPrFixed_(vaRow, vaColMap, id) {
  if (!vaColMap[id]) return "";
  return vaRow[vaColMap[id] - 1];
}

function getPhotoUrlForTTshopPrFixed_(vaRow, vaColMap, pr) {
  if (pr["05"]) return pr["05"];
  if (vaColMap["05"]) return vaRow[vaColMap["05"] - 1];
  return "";
}

function convertSku061To064_TTshopPrFixed_(sku) {
  const text = String(sku || "").trim();
  if (!text) return "";

  const parts = text.split("-");
  if (parts.length >= 4) {
    parts.pop();
    return parts.join("-");
  }

  return text;
}

function normalizeSize_TTshopPrFixed_(value) {
  const text = String(value || "").trim().toUpperCase();
  if (!text) return "";
  if (["FREE", "FREES", "FREE SIZE", "フリー", "フリーサイズ"].includes(text)) return "F";
  return text;
}

function getTTshopPrFixedImageUrl_(url) {
  const text = String(url || "").trim();
  if (!text) return "";

  const match = text.match(/(?:id=|\/d\/)([\w-]+)/);
  if (match) return "https://drive.google.com/uc?export=download&id=" + match[1];

  return text;
}

function escapeTTshopPrFixedFormulaText_(value) {
  return String(value || "").replace(/"/g, '""');
}

function getTTshopPrFixedColumnLetter_(column) {
  let temp;
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

function ensureTTshopPrFixedRows_(sheet, neededLastRow) {
  const maxRows = sheet.getMaxRows();
  if (neededLastRow > maxRows) {
    sheet.insertRowsAfter(maxRows, neededLastRow - maxRows);
  }
}

function toTTshopPrFixedNumber_(value) {
  const num = Number(String(value || "").replace(/,/g, ""));
  return isNaN(num) ? 0 : num;
}

function getTTshopPrFixedCountrySortNo_(value) {
  const text = String(value || "").trim().toUpperCase();
  if (text === "VN" || text === "VND" || text.includes("ベトナム")) return 1;
  if (text === "CN" || text === "CNY" || text === "RMB" || text.includes("中国")) return 2;
  return 9;
}

function getTTshopPrFixedModelCode_(key064, vaRow, vaColMap) {
  if (vaColMap["06"]) {
    const model = String(vaRow[vaColMap["06"] - 1] || "").trim();
    if (model) return model;
  }
  return getTTshopPrFixedCodePart_(key064, 0);
}

function getTTshopPrFixedCodePart_(key064, index) {
  return String(key064 || "").split("-")[index] || "";
}

function getTTshopPrFixedSupplierCodeFrom064_(key064) {
  const parts = String(key064 || "").split("-");
  return parts.length >= 3 ? parts[2] : "";
}

function getTTshopPrFixedVariationCodeFrom064_(key064) {
  const parts = String(key064 || "").split("-");
  return parts.length >= 2 ? parts[1] : "";
}

function getTTshopPrFixedSupplierSortNo_(supplier) {
  const text = String(supplier || "").trim().toUpperCase();
  if (text === "BC") return 1;
  return 9;
}

function getTTshopPrFixedVariationSortNo_(variation) {
  const text = String(variation || "").trim().toUpperCase();

  if (text === "N") return 1;
  if (text === "A") return 2;
  if (text === "B") return 3;
  if (text === "C") return 4;
  if (text === "D") return 5;

  return 99;
}


function getTTshopWarehouseStatusSortNo_(item) {
  // 1: 倉庫現在庫あり
  if ((item.whTotal || 0) > 0) return 1;

  // 2: 倉庫入庫済みだが現在庫なし
  if ((item.intakeTotal || 0) > 0) return 2;

  // 通常は表示対象外だが、念のため最後へ
  return 9;
}

function ttshopPrFixedAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "TTshop", 8);
  }
}
