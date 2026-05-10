/*******************************************************
 * sampleシート生成GAS V1
 *
 * 表示：
 * - 本部現在庫
 * - sample在庫
 *
 * 入力：
 * - sample受け入れ
 * - sampleその他出庫
 * - 棚卸
 *******************************************************/

const SAMPLE_MATRIX_CONFIG = {
  SOURCE_SKU: "SKU",
  SOURCE_VAMASTER: "VaMASTER",
  TARGET_SHEET: "sample",

  HEADER_ROW: 6,
  DATA_START_ROW: 7,

  MASTER_PULL_IDS: ["20", "21", "22"],

  TARGET: {
    TOTAL_COLS: 61, // BI列まで

    SIZE_COLS: {
      XS: [27, 34, 41, 48, 55],
      S:  [28, 35, 42, 49, 56],
      M:  [29, 36, 43, 50, 57],
      L:  [30, 37, 44, 51, 58],
      XL: [31, 38, 45, 52, 59],
      F:  [32, 39, 46, 53, 60]
    },

    SUM_COLS: [
      { col: 33, startRange: "AA", endRange: "AF" }, // 本部現在庫合計
      { col: 40, startRange: "AH", endRange: "AM" }, // sample在庫合計
      { col: 47, startRange: "AO", endRange: "AT" }, // sample受け入れ合計
      { col: 54, startRange: "AV", endRange: "BA" }, // sampleその他出庫合計
      { col: 61, startRange: "BC", endRange: "BH" }  // 棚卸合計
    ]
  },

  SIZE_ORDER: ["XS", "S", "M", "L", "XL", "F"]
};


function generateSampleMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const confirm = Browser.msgBox(
  "サンプルマトリックス構築",
  "サンプルシートを再構築します。\n既存の入力欄は一時保存して復元します。\n\n実行しますか？",
  Browser.Buttons.YES_NO
);

if (confirm !== "yes") return;

  const skuSheet = ss.getSheetByName(SAMPLE_MATRIX_CONFIG.SOURCE_SKU);
  const vaMasterSheet = ss.getSheetByName(SAMPLE_MATRIX_CONFIG.SOURCE_VAMASTER);
  const targetSheet = ss.getSheetByName(SAMPLE_MATRIX_CONFIG.TARGET_SHEET);

  if (!skuSheet || !vaMasterSheet || !targetSheet) {
    Browser.msgBox("SKU / VaMASTER / sample のいずれかのシートが見つかりません。");
    return;
  }

  const skuColMap = getSampleColMap_(skuSheet);
  const vaColMap = getSampleColMap_(vaMasterSheet);
  const tgtColMap = getSampleColMap_(targetSheet);

  if (!tgtColMap["064"] || !vaColMap["064"]) {
    Browser.msgBox("064_列が見つかりません。");
    return;
  }

  if (!skuColMap["061"]) {
    Browser.msgBox("SKUシートに061_SKUコード列が見つかりません。");
    return;
  }

  // 既存入力バックアップ
  const backupMap = new Map();
  const targetLastRow = targetSheet.getLastRow();

  if (targetLastRow >= SAMPLE_MATRIX_CONFIG.DATA_START_ROW) {
    const existingData = targetSheet.getRange(
      SAMPLE_MATRIX_CONFIG.DATA_START_ROW,
      1,
      targetLastRow - SAMPLE_MATRIX_CONFIG.DATA_START_ROW + 1,
      SAMPLE_MATRIX_CONFIG.TARGET.TOTAL_COLS
    ).getValues();

    existingData.forEach(row => {
      const key064 = String(row[tgtColMap["064"] - 1] || "").trim();
      if (!key064) return;

      const savedInput = {};

      for (let c = 27; c <= SAMPLE_MATRIX_CONFIG.TARGET.TOTAL_COLS; c++) {
        if (c >= 27 && c <= 32) continue; // 本部現在庫
        if (c >= 34 && c <= 39) continue; // sample在庫
        if ([33, 40, 47, 54, 61].includes(c)) continue; // 合計列

        if (row[c - 1] !== "" && row[c - 1] !== null) {
          savedInput[c] = row[c - 1];
        }
      }

      backupMap.set(key064, savedInput);
    });
  }

  // 初期化
  if (targetLastRow >= SAMPLE_MATRIX_CONFIG.DATA_START_ROW) {
    clearSampleContentAndProtections_(
      targetSheet.getRange(
        SAMPLE_MATRIX_CONFIG.DATA_START_ROW,
        1,
        targetLastRow - SAMPLE_MATRIX_CONFIG.DATA_START_ROW + 1,
        SAMPLE_MATRIX_CONFIG.TARGET.TOTAL_COLS
      )
    );
  }

  // VaMASTERからサイズ判定
  const sizeExistMap = new Map();

  const vaDataAll = vaMasterSheet.getRange(
    SAMPLE_MATRIX_CONFIG.DATA_START_ROW,
    1,
    Math.max(vaMasterSheet.getLastRow() - SAMPLE_MATRIX_CONFIG.DATA_START_ROW + 1, 1),
    vaMasterSheet.getLastColumn()
  ).getValues();

  vaDataAll.forEach(row => {
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

  const outputData = [];
  const outputBackgrounds = [];
  const smartChipCopyTasks = [];

  vaDataAll.forEach((vaRow, idx) => {
    const key064 = String(vaRow[vaColMap["064"] - 1] || "").trim();
    if (!key064) return;

    const rowData = new Array(SAMPLE_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill("");
    const rowBg = new Array(SAMPLE_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill(null);

    const actualSizes = sizeExistMap.get(key064) || new Set();

    // A〜ZはVaMASTERから作る
    for (let id in tgtColMap) {
      const tgtCol = tgtColMap[id];

      if (tgtCol > 26) continue;

      if (id === "15") {
        rowData[tgtCol - 1] = "";
      } else if (id === "09") {
        rowData[tgtCol - 1] = SAMPLE_MATRIX_CONFIG.SIZE_ORDER
          .filter(size => actualSizes.has(size))
          .join(", ");
      } else if (id === "04" && vaColMap["05"]) {
        const photoUrl = vaRow[vaColMap["05"] - 1];
        if (photoUrl) {
          rowData[tgtCol - 1] = `=IMAGE("${getSampleDirectImageUrl_(photoUrl)}")`;
        }
      } else if (SAMPLE_MATRIX_CONFIG.MASTER_PULL_IDS.includes(id) && vaColMap[id]) {
        smartChipCopyTasks.push({
          srcRow: idx + SAMPLE_MATRIX_CONFIG.DATA_START_ROW,
          srcCol: vaColMap[id],
          dstRow: SAMPLE_MATRIX_CONFIG.DATA_START_ROW + outputData.length,
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
    [34, 35, 36, 37, 38, 39].forEach(c => rowBg[c - 1] = "#f3f3f3"); // sample在庫

    [33, 40, 47, 54, 61].forEach(c => rowBg[c - 1] = "#fff2cc"); // 合計

    SAMPLE_MATRIX_CONFIG.SIZE_ORDER.forEach(size => {
      if (!actualSizes.has(size)) {
        SAMPLE_MATRIX_CONFIG.TARGET.SIZE_COLS[size].forEach(c => {
          rowBg[c - 1] = "#999999";
        });
      }
    });

    outputData.push(rowData);
    outputBackgrounds.push(rowBg);
  });

  if (outputData.length === 0) {
    Browser.msgBox("展開するデータがありませんでした。");
    return;
  }

  const currentMaxRows = targetSheet.getMaxRows();
  const neededMaxRows = SAMPLE_MATRIX_CONFIG.DATA_START_ROW + outputData.length - 1;

  if (neededMaxRows > currentMaxRows) {
    targetSheet.insertRowsAfter(currentMaxRows, neededMaxRows - currentMaxRows);
  }

  const targetRange = targetSheet.getRange(
    SAMPLE_MATRIX_CONFIG.DATA_START_ROW,
    1,
    outputData.length,
    SAMPLE_MATRIX_CONFIG.TARGET.TOTAL_COLS
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

  const fStart = SAMPLE_MATRIX_CONFIG.DATA_START_ROW;
  const finalLastRow = neededMaxRows;
  const col064Str = getSampleColumnLetter_(tgtColMap["064"]);

  // 日本円数式
  const colJpy = tgtColMap["15"];

  if (colJpy && tgtColMap["13"] && tgtColMap["14"]) {
    const formula =
      `=BYROW(${getSampleColumnLetter_(tgtColMap["13"])}${fStart}:` +
      `${getSampleColumnLetter_(tgtColMap["14"])}${finalLastRow}, ` +
      `LAMBDA(row, IF(INDEX(row, 1, 2)="", "", ` +
      `IF(INDEX(row, 1, 1)="VN", INDEX(row, 1, 2) * $Q$2, ` +
      `IF(INDEX(row, 1, 1)="CN", INDEX(row, 1, 2) * $Q$3, "")))))`;

    targetSheet.getRange(fStart, colJpy).setFormula(formula);
  }

  // 合計数式
  SAMPLE_MATRIX_CONFIG.TARGET.SUM_COLS.forEach(sumConfig => {
    targetSheet
      .getRange(fStart, sumConfig.col)
      .setFormula(
        `=BYROW(${sumConfig.startRange}${fStart}:` +
        `${sumConfig.endRange}${finalLastRow}, LAMBDA(row, SUM(row)))`
      );
  });

  // SKUから現在庫を引く
  const warehouseStockIndex = 34; // 本部現在庫
  const sampleStockIndex = 49;    // AW列：sample在庫 ※違う場合はここを変更

  const warehouseStockFormulas = [];
  const sampleStockFormulas = [];

  for (let r = fStart; r <= finalLastRow; r++) {
    const whRow = new Array(6).fill("");
    const sampleRow = new Array(6).fill("");

    SAMPLE_MATRIX_CONFIG.SIZE_ORDER.forEach((size, idx) => {
      whRow[idx] =
        `=IF($${col064Str}${r}="", "", ` +
        `IFERROR(VLOOKUP($${col064Str}${r} & "-${size}", 'SKU'!$A:$BE, ${warehouseStockIndex}, FALSE), ""))`;

      sampleRow[idx] =
        `=IF($${col064Str}${r}="", "", ` +
        `IFERROR(VLOOKUP($${col064Str}${r} & "-${size}", 'SKU'!$A:$BE, ${sampleStockIndex}, FALSE), ""))`;
    });

    warehouseStockFormulas.push(whRow);
    sampleStockFormulas.push(sampleRow);
  }

  // 本部現在庫 AA:AF
  targetSheet
    .getRange(fStart, 27, warehouseStockFormulas.length, 6)
    .setFormulas(warehouseStockFormulas);

  // sample在庫 AH:AM
  targetSheet
    .getRange(fStart, 34, sampleStockFormulas.length, 6)
    .setFormulas(sampleStockFormulas);

  // 棚卸アラート
  const rules = targetSheet.getConditionalFormatRules();

  const inventoryRange = targetSheet.getRange(`BC${fStart}:BH${finalLastRow}`);

  const inventoryAlertRule = SpreadsheetApp
    .newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(BC${fStart}<>"", BC${fStart}<>AH${fStart})`)
    .setBackground("#f8cecc")
    .setFontColor("#cc0000")
    .setRanges([inventoryRange])
    .build();

  rules.push(inventoryAlertRule);
  targetSheet.setConditionalFormatRules(rules);

  // 警告保護
  const protectRanges = [
    targetSheet.getRange(`AA${fStart}:AF${finalLastRow}`),
    targetSheet.getRange(`AH${fStart}:AM${finalLastRow}`),

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

    targetSheet.getRange(`AG${fStart}:AG${finalLastRow}`),
    targetSheet.getRange(`AN${fStart}:AN${finalLastRow}`),
    targetSheet.getRange(`AU${fStart}:AU${finalLastRow}`),
    targetSheet.getRange(`BB${fStart}:BB${finalLastRow}`),
    targetSheet.getRange(`BI${fStart}:BI${finalLastRow}`)
  ];

  protectRanges.forEach(range => {
    range.protect().setWarningOnly(true);
  });

  try {
    SpreadsheetApp.getUi().alert(
      "sampleマトリックス生成完了\n\n" +
      "構築行数：" + outputData.length + " 行\n" +
      "最終行：" + finalLastRow
    );
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "sampleマトリックス生成完了",
      "完了",
      10
    );
  }
}


/*******************************************************
 * ヘルパー
 *******************************************************/

function getSampleColMap_(sheet) {
  const map = {};

  const headers = sheet
    .getRange(
      SAMPLE_MATRIX_CONFIG.HEADER_ROW,
      1,
      1,
      Math.max(sheet.getLastColumn(), 1)
    )
    .getValues()[0];

  headers.forEach((header, idx) => {
    const text = normalizeSampleId_(header);
    const match = text.match(/^(\d{2,4})_/);

    if (match) {
      map[match[1]] = idx + 1;
    }
  });

  return map;
}

function normalizeSampleId_(value) {
  return String(value || "")
    .trim()
    .replace(/[０-９]/g, s =>
      String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
    )
    .replace(/＿/g, "_");
}

function getSampleDirectImageUrl_(url) {
  if (!url) return "";

  const match = String(url).match(/(?:id=|d\/)([\w-]+)/);

  if (match) {
    return `https://drive.google.com/uc?export=download&id=${match[1]}`;
  }

  return url;
}

function getSampleColumnLetter_(column) {
  let temp;
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

function clearSampleContentAndProtections_(range) {
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