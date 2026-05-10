/*******************************************************
 * レンタルマトリックス生成GAS V1
 *
 * ベース：
 * - TTshopマトリックス構築
 *
 * 表示：
 * - 本部現在庫
 * - レンタル現在庫
 *
 * 入力：
 * - レンタル受け入れ
 * - 貸出累計
 * - その他出庫
 * - 棚卸
 *
 * 重要：
 * - 貸出累計は在庫から引かない
 * - レンタル現在庫 = レンタル受け入れ - レンタルその他出庫
 *******************************************************/

const RENTAL_MATRIX_CONFIG = {
  SOURCE_SKU: "SKU",
  SOURCE_VAMASTER: "VaMASTER",
  TARGET_SHEET: "レンタル",

  HEADER_ROW: 6,
  DATA_START_ROW: 7,

  MASTER_PULL_IDS: ["20", "21", "22"],

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
      { col: 33, startRange: "AA", endRange: "AF" }, // 本部現在庫合計
      { col: 40, startRange: "AH", endRange: "AM" }, // レンタル現在庫合計
      { col: 47, startRange: "AO", endRange: "AT" }, // レンタル受け入れ合計
      { col: 54, startRange: "AV", endRange: "BA" }, // 貸出累計合計
      { col: 61, startRange: "BC", endRange: "BH" }, // その他出庫合計
      { col: 68, startRange: "BJ", endRange: "BO" }  // 棚卸合計
    ]
  },

  SIZE_ORDER: ["XS", "S", "M", "L", "XL", "F"]
};


/*******************************************************
 * メイン関数
 *******************************************************/

function generateRentalMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const skuSheet = ss.getSheetByName(RENTAL_MATRIX_CONFIG.SOURCE_SKU);
  const vaMasterSheet = ss.getSheetByName(RENTAL_MATRIX_CONFIG.SOURCE_VAMASTER);
  const targetSheet = ss.getSheetByName(RENTAL_MATRIX_CONFIG.TARGET_SHEET);

  if (!skuSheet || !vaMasterSheet || !targetSheet) {
    Browser.msgBox("SKU / VaMASTER / レンタル のいずれかのシートが見つかりません。");
    return;
  }

  const skuColMap = getRentalColMap_(skuSheet);
  const vaColMap = getRentalColMap_(vaMasterSheet);
  const tgtColMap = getRentalColMap_(targetSheet);

  if (!tgtColMap["064"] || !vaColMap["064"]) {
    Browser.msgBox("064_列が見つかりません。");
    return;
  }

  if (!skuColMap["061"]) {
    Browser.msgBox("SKUシートに061_SKUコード列が見つかりません。");
    return;
  }

  // ==========================================
  // Step 1: 手入力データのバックアップ
  // ==========================================

  const backupMap = new Map();
  const targetLastRow = targetSheet.getLastRow();

  if (targetLastRow >= RENTAL_MATRIX_CONFIG.DATA_START_ROW) {
    const existingData = targetSheet.getRange(
      RENTAL_MATRIX_CONFIG.DATA_START_ROW,
      1,
      targetLastRow - RENTAL_MATRIX_CONFIG.DATA_START_ROW + 1,
      RENTAL_MATRIX_CONFIG.TARGET.TOTAL_COLS
    ).getValues();

    existingData.forEach(row => {
      const key064 = String(row[tgtColMap["064"] - 1] || "").trim();
      if (!key064) return;

      const savedInput = {};

      for (let c = 27; c <= RENTAL_MATRIX_CONFIG.TARGET.TOTAL_COLS; c++) {
        // 表示専用：本部現在庫 AA:AF
        if (c >= 27 && c <= 32) continue;

        // 表示専用：レンタル現在庫 AH:AM
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

  if (targetLastRow >= RENTAL_MATRIX_CONFIG.DATA_START_ROW) {
    clearRentalContentAndProtections_(
      targetSheet.getRange(
        RENTAL_MATRIX_CONFIG.DATA_START_ROW,
        1,
        targetLastRow - RENTAL_MATRIX_CONFIG.DATA_START_ROW + 1,
        RENTAL_MATRIX_CONFIG.TARGET.TOTAL_COLS
      )
    );
  }

  // ==========================================
  // Step 3: VaMASTERからサイズ存在判定
  // ==========================================

  const sizeExistMap = new Map();

  const vaSizeData = vaMasterSheet.getRange(
    RENTAL_MATRIX_CONFIG.DATA_START_ROW,
    1,
    Math.max(vaMasterSheet.getLastRow() - RENTAL_MATRIX_CONFIG.DATA_START_ROW + 1, 1),
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
    RENTAL_MATRIX_CONFIG.DATA_START_ROW,
    1,
    Math.max(vaMasterSheet.getLastRow() - RENTAL_MATRIX_CONFIG.DATA_START_ROW + 1, 1),
    vaMasterSheet.getLastColumn()
  ).getValues();

  const outputData = [];
  const outputBackgrounds = [];
  const smartChipCopyTasks = [];

  vaData.forEach((vaRow, idx) => {
    const key064 = String(vaRow[vaColMap["064"] - 1] || "").trim();
    if (!key064) return;

    const rowData = new Array(RENTAL_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill("");
    const rowBg = new Array(RENTAL_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill(null);

    const actualSizes = sizeExistMap.get(key064) || new Set();

    // 左側A〜ZはVaMASTERから作る
    for (let id in tgtColMap) {
      const tgtCol = tgtColMap[id];

      if (tgtCol > 26) continue;

      if (id === "15") {
        rowData[tgtCol - 1] = "";
      } else if (id === "09") {
        rowData[tgtCol - 1] = RENTAL_MATRIX_CONFIG.SIZE_ORDER
          .filter(size => actualSizes.has(size))
          .join(", ");
      } else if (id === "04" && vaColMap["05"]) {
        const photoUrl = vaRow[vaColMap["05"] - 1];
        if (photoUrl) {
          rowData[tgtCol - 1] = `=IMAGE("${getRentalDirectImageUrl_(photoUrl)}")`;
        }
      } else if (RENTAL_MATRIX_CONFIG.MASTER_PULL_IDS.includes(id) && vaColMap[id]) {
        smartChipCopyTasks.push({
          srcRow: idx + RENTAL_MATRIX_CONFIG.DATA_START_ROW,
          srcCol: vaColMap[id],
          dstRow: RENTAL_MATRIX_CONFIG.DATA_START_ROW + outputData.length,
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
    [27, 28, 29, 30, 31, 32].forEach(c => rowBg[c - 1] = "#f3f3f3"); // 本部現在庫
    [34, 35, 36, 37, 38, 39].forEach(c => rowBg[c - 1] = "#f3f3f3"); // レンタル現在庫

    [33, 40, 47, 54, 61, 68].forEach(c => rowBg[c - 1] = "#fff2cc"); // 合計

    RENTAL_MATRIX_CONFIG.SIZE_ORDER.forEach(size => {
      if (!actualSizes.has(size)) {
        RENTAL_MATRIX_CONFIG.TARGET.SIZE_COLS[size].forEach(c => {
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
    RENTAL_MATRIX_CONFIG.DATA_START_ROW + outputData.length - 1;

  if (neededMaxRows > currentMaxRows) {
    targetSheet.insertRowsAfter(
      currentMaxRows,
      neededMaxRows - currentMaxRows
    );
  }

  const targetRange = targetSheet.getRange(
    RENTAL_MATRIX_CONFIG.DATA_START_ROW,
    1,
    outputData.length,
    RENTAL_MATRIX_CONFIG.TARGET.TOTAL_COLS
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

  const fStart = RENTAL_MATRIX_CONFIG.DATA_START_ROW;
  const finalLastRow = neededMaxRows;
  const col064Str = getRentalColumnLetter_(tgtColMap["064"]);

  // ==========================================
  // Step 6: 日本円数式
  // ==========================================

  const colJpy = tgtColMap["15"];

  if (colJpy && tgtColMap["13"] && tgtColMap["14"]) {
    const formula =
      `=BYROW(${getRentalColumnLetter_(tgtColMap["13"])}${fStart}:` +
      `${getRentalColumnLetter_(tgtColMap["14"])}${finalLastRow}, ` +
      `LAMBDA(row, IF(INDEX(row, 1, 2)="", "", ` +
      `IF(INDEX(row, 1, 1)="VN", INDEX(row, 1, 2) * $Q$2, ` +
      `IF(INDEX(row, 1, 1)="CN", INDEX(row, 1, 2) * $Q$3, "")))))`;

    targetSheet.getRange(fStart, colJpy).setFormula(formula);
  }

  // ==========================================
  // Step 7: 合計数式
  // ==========================================

  RENTAL_MATRIX_CONFIG.TARGET.SUM_COLS.forEach(sumConfig => {
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

  const warehouseStockIndex = 34; // 本部現在庫
  const rentalStockIndex = 45;    // AS列：レンタル現在庫

  const warehouseStockFormulas = [];
  const rentalStockFormulas = [];

  for (let r = fStart; r <= finalLastRow; r++) {
    const whRow = new Array(6).fill("");
    const rentalRow = new Array(6).fill("");

    RENTAL_MATRIX_CONFIG.SIZE_ORDER.forEach((size, idx) => {
      whRow[idx] =
        `=IF($${col064Str}${r}="", "", ` +
        `IFERROR(VLOOKUP($${col064Str}${r} & "-${size}", 'SKU'!$A:$BE, ${warehouseStockIndex}, FALSE), ""))`;

      rentalRow[idx] =
        `=IF($${col064Str}${r}="", "", ` +
        `IFERROR(VLOOKUP($${col064Str}${r} & "-${size}", 'SKU'!$A:$BE, ${rentalStockIndex}, FALSE), ""))`;
    });

    warehouseStockFormulas.push(whRow);
    rentalStockFormulas.push(rentalRow);
  }

  // 本部現在庫 AA:AF
  targetSheet
    .getRange(fStart, 27, warehouseStockFormulas.length, 6)
    .setFormulas(warehouseStockFormulas);

  // レンタル現在庫 AH:AM
  targetSheet
    .getRange(fStart, 34, rentalStockFormulas.length, 6)
    .setFormulas(rentalStockFormulas);

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
    // 本部現在庫
    targetSheet.getRange(`AA${fStart}:AF${finalLastRow}`),

    // レンタル現在庫
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
    "レンタルマトリックス生成完了！\n" +
    "構築行数：" + outputData.length + " 行\n" +
    "最終行：" + finalLastRow + "\n\n" +
    "本部現在庫・レンタル現在庫・受け入れ・貸出累計・その他出庫・棚卸欄をセットしました。"
  );
}


/*******************************************************
 * ヘルパー
 *******************************************************/

function getRentalColMap_(sheet) {
  const map = {};

  const headers = sheet
    .getRange(
      RENTAL_MATRIX_CONFIG.HEADER_ROW,
      1,
      1,
      Math.max(sheet.getLastColumn(), 1)
    )
    .getValues()[0];

  headers.forEach((header, idx) => {
    const text = normalizeRentalId_(header);
    const match = text.match(/^(\d{2,4})_/);

    if (match) {
      map[match[1]] = idx + 1;
    }
  });

  return map;
}

function normalizeRentalId_(value) {
  return String(value || "")
    .trim()
    .replace(/[０-９]/g, s =>
      String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
    )
    .replace(/＿/g, "_");
}

function getRentalDirectImageUrl_(url) {
  if (!url) return "";

  const match = String(url).match(/(?:id=|d\/)([\w-]+)/);

  if (match) {
    return `https://drive.google.com/uc?export=download&id=${match[1]}`;
  }

  return url;
}

function getRentalColumnLetter_(column) {
  let temp;
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

function clearRentalContentAndProtections_(range) {
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