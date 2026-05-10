/*******************************************************
 * TTshopシート構築GAS 
 *
 * ベース：
 * - VaMASTER起点
 * - 064_P+V+Sコード起点
 * - サイズ存在判定はVaMASTERの09_
 * - 在庫参照はSKUシート
 *
 * 表示：
 * - 本部倉庫現在庫
 * - TTshop現在庫
 *
 * 入力：
 * - 持出(受入)
 * - 販売
 * - 不良廃棄その他
 * - 棚卸
 *******************************************************/

const TTSHOP_MATRIX_CONFIG = {
  SOURCE_SKU: "SKU",
  SOURCE_VAMASTER: "VaMASTER",
  TARGET_SHEET: "TTshop",

  HEADER_ROW: 6,
  DATA_START_ROW: 7,

  MASTER_PULL_IDS: ["20", "21", "22"],

  // SKUシート側の在庫項目ID
  SKU_STOCK_ID_WAREHOUSE: "50", // 倉庫現在庫
  SKU_STOCK_ID_MOBILE: "56",    // TTshop現在庫

  TARGET: {
    TOTAL_COLS: 68, // BP列まで

    SIZE_COLS: {
      XS: [27, 34, 41, 48, 55, 62],
      S:  [28, 35, 42, 49, 56, 63],
      M:  [29, 36, 43, 50, 57, 64],
      L:  [30, 37, 44, 51, 58, 65],
      XL: [31, 38, 45, 52, 59, 66],
      F:  [32, 39, 46, 53, 60, 67]
    },

    SUM_COLS: [
      { col: 33, startRange: "AA", endRange: "AF" }, // 倉庫現在庫合計
      { col: 40, startRange: "AH", endRange: "AM" }, // 移動販売現在庫合計
      { col: 47, startRange: "AO", endRange: "AT" }, // 持出合計
      { col: 54, startRange: "AV", endRange: "BA" }, // 販売合計
      { col: 61, startRange: "BC", endRange: "BH" }, // 不良廃棄その他合計
      { col: 68, startRange: "BJ", endRange: "BO" }  // 棚卸合計
    ]
  },

  SIZE_ORDER: ["XS", "S", "M", "L", "XL", "F"]
};


/*******************************************************
 * メイン関数
 *******************************************************/

function generateTTshopMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const skuSheet = ss.getSheetByName(TTSHOP_MATRIX_CONFIG.SOURCE_SKU);
  const vaMasterSheet = ss.getSheetByName(TTSHOP_MATRIX_CONFIG.SOURCE_VAMASTER);
  const targetSheet = ss.getSheetByName(TTSHOP_MATRIX_CONFIG.TARGET_SHEET);

  if (!skuSheet || !vaMasterSheet || !targetSheet) {
    Browser.msgBox("SKU / VaMASTER / TTshop のいずれかのシートが見つかりません。");
    return;
  }

  const skuColMap = getMobileColMap_(skuSheet);
  const vaColMap = getMobileColMap_(vaMasterSheet);
  const tgtColMap = getMobileColMap_(targetSheet);

  if (!tgtColMap["064"] || !vaColMap["064"]) {
    Browser.msgBox("064_列が見つかりません。");
    return;
  }

  if (!skuColMap["061"]) {
    Browser.msgBox("SKUシートに061_SKUコード列が見つかりません。");
    return;
  }

  if (!skuColMap[TTSHOP_MATRIX_CONFIG.SKU_STOCK_ID_WAREHOUSE]) {
    Browser.msgBox("SKUシートに50_倉庫現在庫列が見つかりません。");
    return;
  }

  if (!skuColMap[TTSHOP_MATRIX_CONFIG.SKU_STOCK_ID_MOBILE]) {
    Browser.msgBox("SKUシートに51_移動販売現在庫列が見つかりません。");
    return;
  }

  // ==========================================
  // Step 1: 手入力データのバックアップ
  // ==========================================

  const backupMap = new Map();
  const targetLastRow = targetSheet.getLastRow();

  if (targetLastRow >= TTSHOP_MATRIX_CONFIG.DATA_START_ROW) {
    const existingData = targetSheet.getRange(
      TTSHOP_MATRIX_CONFIG.DATA_START_ROW,
      1,
      targetLastRow - TTSHOP_MATRIX_CONFIG.DATA_START_ROW + 1,
      TTSHOP_MATRIX_CONFIG.TARGET.TOTAL_COLS
    ).getValues();

    existingData.forEach(row => {
      const key064 = String(row[tgtColMap["064"] - 1] || "").trim();
      if (!key064) return;

      const savedInput = {};

      for (let c = 27; c <= TTSHOP_MATRIX_CONFIG.TARGET.TOTAL_COLS; c++) {
        // 表示専用：倉庫現在庫 AA:AF
        if (c >= 27 && c <= 32) continue;

        // 表示専用：移動販売現在庫 AH:AM
        if (c >= 34 && c <= 39) continue;

        // 合計列
        if ([33, 40, 47, 54, 61, 68].includes(c)) continue;

        if (row[c - 1] !== "" && row[c - 1] !== null) {
          savedInput[c] = row[c - 1];
        }
      }

      backupMap.set(key064, savedInput);
    });
  }

  // ==========================================
  // Step 2: 初期化
  // ==========================================

  if (targetLastRow >= TTSHOP_MATRIX_CONFIG.DATA_START_ROW) {
    clearMobileContentAndProtections_(
      targetSheet.getRange(
        TTSHOP_MATRIX_CONFIG.DATA_START_ROW,
        1,
        targetLastRow - TTSHOP_MATRIX_CONFIG.DATA_START_ROW + 1,
        TTSHOP_MATRIX_CONFIG.TARGET.TOTAL_COLS
      )
    );
  }

  // ==========================================
  // Step 3: VaMASTERからサイズ存在判定
  // ==========================================

  const sizeExistMap = new Map();

  const vaSizeData = vaMasterSheet.getRange(
    TTSHOP_MATRIX_CONFIG.DATA_START_ROW,
    1,
    Math.max(vaMasterSheet.getLastRow() - TTSHOP_MATRIX_CONFIG.DATA_START_ROW + 1, 1),
    vaMasterSheet.getLastColumn()
  ).getValues();

  vaSizeData.forEach(row => {
    const key064 = String(row[vaColMap["064"] - 1] || "").trim();
    if (!key064) return;

    const rawSizeText = String(row[vaColMap["09"] - 1] || "").trim();

    if (!sizeExistMap.has(key064)) {
      sizeExistMap.set(key064, new Set());
    }

    const sizeSet = sizeExistMap.get(key064);

    rawSizeText
      .split(/[,、\n]/)
      .map(size => String(size).trim().toUpperCase())
      .filter(Boolean)
      .forEach(size => {
        if (
          size === "FREE" ||
          size === "FREES" ||
          size === "FREE SIZE" ||
          size === "フリー"
        ) {
          sizeSet.add("F");
        } else {
          sizeSet.add(size);
        }
      });
  });

  // ==========================================
  // Step 4: VaMASTERから縦軸フレーム作成
  // ==========================================

  const vaData = vaMasterSheet.getRange(
    TTSHOP_MATRIX_CONFIG.DATA_START_ROW,
    1,
    Math.max(vaMasterSheet.getLastRow() - TTSHOP_MATRIX_CONFIG.DATA_START_ROW + 1, 1),
    vaMasterSheet.getLastColumn()
  ).getValues();

  const outputData = [];
  const outputBackgrounds = [];
  const smartChipCopyTasks = [];

  vaData.forEach((vaRow, idx) => {
    const key064 = String(vaRow[vaColMap["064"] - 1] || "").trim();
    if (!key064) return;

    const rowData = new Array(TTSHOP_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill("");
    const rowBg = new Array(TTSHOP_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill(null);

    const actualSizes = sizeExistMap.get(key064) || new Set();

    // 左側A〜ZはVaMASTERから作る
    for (let id in tgtColMap) {
      const tgtCol = tgtColMap[id];

      if (tgtCol > 26) continue;

      if (id === "15") {
        rowData[tgtCol - 1] = "";
      } else if (id === "09") {
        rowData[tgtCol - 1] = TTSHOP_MATRIX_CONFIG.SIZE_ORDER
          .filter(size => actualSizes.has(size))
          .join(", ");
      } else if (id === "04" && vaColMap["05"]) {
        const photoUrl = vaRow[vaColMap["05"] - 1];
        if (photoUrl) {
          rowData[tgtCol - 1] = `=IMAGE("${getMobileDirectImageUrl_(photoUrl)}")`;
        }
      } else if (TTSHOP_MATRIX_CONFIG.MASTER_PULL_IDS.includes(id) && vaColMap[id]) {
        smartChipCopyTasks.push({
          srcRow: idx + TTSHOP_MATRIX_CONFIG.DATA_START_ROW,
          srcCol: vaColMap[id],
          dstRow: TTSHOP_MATRIX_CONFIG.DATA_START_ROW + outputData.length,
          dstCol: tgtCol
        });
      } else if (vaColMap[id]) {
        rowData[tgtCol - 1] = vaRow[vaColMap[id] - 1];
      }
    }

    // 手入力復元
    if (backupMap.has(key064)) {
      const savedInput = backupMap.get(key064);
      for (let c in savedInput) {
        rowData[c - 1] = savedInput[c];
      }
    }

    // 色設定
    [27, 28, 29, 30, 31, 32].forEach(c => rowBg[c - 1] = "#f3f3f3"); // 倉庫現在庫
    [34, 35, 36, 37, 38, 39].forEach(c => rowBg[c - 1] = "#f3f3f3"); // TTshop在庫

    [33, 40, 47, 54, 61, 68].forEach(c => rowBg[c - 1] = "#fff2cc"); // 合計

    TTSHOP_MATRIX_CONFIG.SIZE_ORDER.forEach(size => {
      if (!actualSizes.has(size)) {
        TTSHOP_MATRIX_CONFIG.TARGET.SIZE_COLS[size].forEach(c => {
          rowBg[c - 1] = "#999999";
        });
      }
    });

    outputData.push(rowData);
    outputBackgrounds.push(rowBg);
  });

  // ==========================================
  // Step 5: 書き込み
  // ==========================================

  if (outputData.length === 0) {
    Browser.msgBox("展開するデータがありませんでした。");
    return;
  }

  const currentMaxRows = targetSheet.getMaxRows();
  const neededMaxRows =
    TTSHOP_MATRIX_CONFIG.DATA_START_ROW + outputData.length - 1;

  if (neededMaxRows > currentMaxRows) {
    targetSheet.insertRowsAfter(
      currentMaxRows,
      neededMaxRows - currentMaxRows
    );
  }

  const targetRange = targetSheet.getRange(
    TTSHOP_MATRIX_CONFIG.DATA_START_ROW,
    1,
    outputData.length,
    TTSHOP_MATRIX_CONFIG.TARGET.TOTAL_COLS
  );

  targetRange.setValues(outputData);
  targetRange.setBackgrounds(outputBackgrounds);

  // スマートチップコピー
  smartChipCopyTasks.forEach(task => {
    vaMasterSheet
      .getRange(task.srcRow, task.srcCol)
      .copyTo(targetSheet.getRange(task.dstRow, task.dstCol));

    targetSheet
      .getRange(task.dstRow, task.dstCol)
      .setBackground(null);
  });

  const fStart = TTSHOP_MATRIX_CONFIG.DATA_START_ROW;
  const finalLastRow = neededMaxRows;
  const col064Str = getMobileColumnLetter_(tgtColMap["064"]);

  // ==========================================
  // Step 6: 日本円数式
  // ==========================================

  const colJpy = tgtColMap["15"];

  if (colJpy && tgtColMap["13"] && tgtColMap["14"]) {
    const formula =
      `=BYROW(${getMobileColumnLetter_(tgtColMap["13"])}${fStart}:` +
      `${getMobileColumnLetter_(tgtColMap["14"])}${finalLastRow}, ` +
      `LAMBDA(row, IF(INDEX(row, 1, 2)="", "", ` +
      `IF(INDEX(row, 1, 1)="VN", INDEX(row, 1, 2) * $Q$2, ` +
      `IF(INDEX(row, 1, 1)="CN", INDEX(row, 1, 2) * $Q$3, "")))))`;

    targetSheet.getRange(fStart, colJpy).setFormula(formula);
  }

  // ==========================================
  // Step 7: 合計数式
  // ==========================================

  TTSHOP_MATRIX_CONFIG.TARGET.SUM_COLS.forEach(sumConfig => {
    targetSheet
      .getRange(fStart, sumConfig.col)
      .setFormula(
        `=BYROW(${sumConfig.startRange}${fStart}:` +
        `${sumConfig.endRange}${finalLastRow}, LAMBDA(row, SUM(row)))`
      );
  });

  // ==========================================
  // Step 8: SKUから現在庫を引く
  // ==========================================

  const warehouseStockIndex = 34; // 本部/倉庫現在庫
  const ttStockIndex = 38;        // TTshop現在庫 AL列

  const warehouseStockFormulas = [];
  const ttshopStockFormulas = [];

  for (let r = fStart; r <= finalLastRow; r++) {
    const whRow = new Array(6).fill("");
    const mvRow = new Array(6).fill("");

    TTSHOP_MATRIX_CONFIG.SIZE_ORDER.forEach((size, idx) => {
      whRow[idx] =
        `=IF($${col064Str}${r}="", "", ` +
        `IFERROR(VLOOKUP($${col064Str}${r} & "-${size}", 'SKU'!$A:$BE, ${warehouseStockIndex}, FALSE), ""))`;

      mvRow[idx] =
        `=IF($${col064Str}${r}="", "", ` +
        `IFERROR(VLOOKUP($${col064Str}${r} & "-${size}", 'SKU'!$A:$BE, ${mobileStockIndex}, FALSE), ""))`;
    });

    warehouseStockFormulas.push(whRow);
    ttshopStockFormulas.push(mvRow);
  }

  // 倉庫現在庫 AA:AF
  targetSheet
    .getRange(fStart, 27, warehouseStockFormulas.length, 6)
    .setFormulas(warehouseStockFormulas);

  // TTshop現在庫 AH:AM
  targetSheet
    .getRange(fStart, 34, mobileStockFormulas.length, 6)
    .setFormulas(ttshopStockFormulas);

  // ==========================================
  // Step 9: 棚卸アラート
  // ==========================================

  const rules = targetSheet.getConditionalFormatRules();

  const inventoryRange = targetSheet.getRange(`BJ${fStart}:BO${finalLastRow}`);

  const inventoryAlertRule = SpreadsheetApp
    .newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(BJ${fStart}<>"", BJ${fStart}<>AH${fStart})`)
    .setBackground("#f8cecc")
    .setFontColor("#cc0000")
    .setRanges([inventoryRange])
    .build();

  rules.push(inventoryAlertRule);
  targetSheet.setConditionalFormatRules(rules);

  // ==========================================
  // Step 10: 警告保護
  // ==========================================

  const protectRanges = [
    // 倉庫現在庫
    targetSheet.getRange(`AA${fStart}:AF${finalLastRow}`),

    // TTshop現在庫
    targetSheet.getRange(`AH${fStart}:AM${finalLastRow}`),

    // 卸価格系：項目ID 14 / 15
    targetSheet.getRange(
      fStart,
      tgtColMap["14"],
      finalLastRow - fStart + 1,
      1
    ),

    targetSheet.getRange(
      fStart,
      tgtColMap["15"],
      finalLastRow - fStart + 1,
      1
    ),

    // 合計列
    targetSheet.getRange(`AG${fStart}:AG${finalLastRow}`),
    targetSheet.getRange(`AN${fStart}:AN${finalLastRow}`),
    targetSheet.getRange(`AU${fStart}:AU${finalLastRow}`),
    targetSheet.getRange(`BB${fStart}:BB${finalLastRow}`),
    targetSheet.getRange(`BI${fStart}:BI${finalLastRow}`),
    targetSheet.getRange(`BP${fStart}:BP${finalLastRow}`)
  ];

  protectRanges.forEach(range => {
    range.protect().setWarningOnly(true);
  });

  Browser.msgBox(
    "TTshopマトリックス生成完了！\n" +
    "構築行数：" + outputData.length + " 行\n" +
    "最終行：" + finalLastRow + "\n\n" +
    "倉庫現在庫・TTshop現在庫・入力欄・棚卸欄をセットしました。"
  );
}


/*******************************************************
 * ヘルパー
 *******************************************************/

function getMobileColMap_(sheet) {
  const map = {};

  const headers = sheet
    .getRange(
      TTSHOP_MATRIX_CONFIG.HEADER_ROW,
      1,
      1,
      Math.max(sheet.getLastColumn(), 1)
    )
    .getValues()[0];

  headers.forEach((header, idx) => {
    const text = normalizeMobileId_(header);
    const match = text.match(/^(\d{2,4})_/);

    if (match) {
      map[match[1]] = idx + 1;
    }
  });

  return map;
}


function normalizeMobileId_(value) {
  return String(value || "")
    .trim()
    .replace(/[０-９]/g, s =>
      String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
    )
    .replace(/＿/g, "_");
}


function getMobileDirectImageUrl_(url) {
  if (!url) return "";

  const match = String(url).match(/(?:id=|d\/)([\w-]+)/);

  if (match) {
    return `https://drive.google.com/uc?export=download&id=${match[1]}`;
  }

  return url;
}


function getMobileColumnLetter_(column) {
  let temp;
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}


function clearMobileContentAndProtections_(range) {
  const sheet = range.getSheet();

  range.clearContent();
  range.setBackground(null);

  const protections = sheet.getProtections(
    SpreadsheetApp.ProtectionType.RANGE
  );

  protections.forEach(p => {
    p.remove();
  });

  sheet.clearConditionalFormatRules();
}