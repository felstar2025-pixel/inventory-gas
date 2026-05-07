/**
 * 倉庫マトリックス生成GAS
 * ★アキラさん究極仕様（スマートチップコピペ ＆ 棚卸アラート機能搭載）
 */

const WAREHOUSE_MATRIX_CONFIG = {
  SOURCE_SKU: "SKU",       
  SOURCE_MASTER: "MASTER",  
  TARGET_SHEET: "倉庫",      // ★ターゲットを倉庫シートに変更
  HEADER_ROW: 6,
  DATA_START_ROW: 7,            
  MASTER_PULL_IDS: ["20", "21", "22"], 

  TARGET: {
    TOTAL_COLS: 61, // BI列まで
    SIZE_COLS: {
      "XS": [27, 34, 41, 48, 55], // 在庫, 入庫, 出庫, 不良, 棚卸
      "S":  [28, 35, 42, 49, 56], 
      "M":  [29, 36, 43, 50, 57], 
      "L":  [30, 37, 44, 51, 58], 
      "XL": [31, 38, 45, 52, 59], 
      "F":  [32, 39, 46, 53, 60]  
    },
    SUM_COLS: [
      { col: 33, startRange: "AA", endRange: "AF" }, // 在庫合計
      { col: 40, startRange: "AH", endRange: "AM" }, // 入庫合計
      { col: 47, startRange: "AO", endRange: "AT" }, // 出庫合計
      { col: 54, startRange: "AV", endRange: "BA" }, // 不良合計
      { col: 61, startRange: "BC", endRange: "BH" }  // 棚卸合計
    ]
  },
  SIZE_ORDER: ["XS", "S", "M", "L", "XL", "F"]
};

// ヘルパー関数群（マトリックス生成用）
function getDirectImageUrl(url) {
  if (!url) return "";
  const match = url.match(/(?:id=|d\/)([\w-]+)/);
  if (match) return `https://drive.google.com/uc?export=download&id=${match[1]}`;
  return url;
}
function getColumnLetter(column) {
  let temp, letter = '';
  while (column > 0) { temp = (column - 1) % 26; letter = String.fromCharCode(temp + 65) + letter; column = (column - temp - 1) / 26; }
  return letter;
}
function clearContentAndProtections(range) {
  range.clearContent();
  range.setBackground(null);
  const sheet = range.getSheet();
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let p of protections) { if (p.isWarningOnly()) p.remove(); }
  sheet.clearConditionalFormatRules(); // ★条件付き書式も一度リセット
}

function generateWarehouseMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName(WAREHOUSE_MATRIX_CONFIG.SOURCE_SKU);
  const masterSheet = ss.getSheetByName(WAREHOUSE_MATRIX_CONFIG.SOURCE_MASTER); 
  const targetSheet = ss.getSheetByName(WAREHOUSE_MATRIX_CONFIG.TARGET_SHEET);
  
  if (!skuSheet || !masterSheet || !targetSheet) return Browser.msgBox("シートが見つかりません。");

  const normalizeId = (h) => String(h).trim().replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/＿/g, "_");
  const getColMap = (sheet) => {
    const map = {};
    sheet.getRange(WAREHOUSE_MATRIX_CONFIG.HEADER_ROW, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].forEach((h, idx) => {
      const match = normalizeId(h).match(/^(\d{2,3})_/);
      if (match) map[match[1]] = idx + 1;
    });
    return map;
  };

  const skuColMap = getColMap(skuSheet);
  const masterColMap = getColMap(masterSheet); 
  const tgtColMap = getColMap(targetSheet);
  
  if (!tgtColMap["062"] || !masterColMap["06"] || !skuColMap["06"]) return Browser.msgBox("必須IDが見つかりません。");

  const skuLastRow = skuSheet.getLastRow();
  const skuData = skuSheet.getRange(WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW, 1, skuLastRow - WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW + 1, skuSheet.getLastColumn()).getValues();

  const masterLastRow = masterSheet.getLastRow();
  const masterData = masterLastRow < WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW ? [] : masterSheet.getRange(WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW, 1, masterLastRow - WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW + 1, masterSheet.getLastColumn()).getValues();

  const masterMap = new Map();
  masterData.forEach((row, idx) => {
    const parentCode = String(row[masterColMap["06"] - 1] || "").trim();
    if (parentCode && !masterMap.has(parentCode)) masterMap.set(parentCode, { values: row, rowIndex: idx + WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW });
  });

  const productMap = new Map();
  for (let i = 0; i < skuData.length; i++) {
    const row = skuData[i];
    const tagCode = String(row[skuColMap["062"] - 1] || "").trim();
    const parentCode = String(row[skuColMap["06"] - 1] || "").trim();
    if (!tagCode) continue; 
    if (!productMap.has(tagCode)) productMap.set(tagCode, { skuFirstRow: row, parentCode: parentCode, actualSizes: new Set() });
    const sizeUnit = String(row[skuColMap["09"] - 1] || "").trim().toUpperCase();
    if (sizeUnit) productMap.get(tagCode).actualSizes.add((sizeUnit === "FREE" || sizeUnit === "FREES") ? "F" : sizeUnit);
  }
  
  let outputData = [];
  let outputBackgrounds = [];
  let smartChipCopyTasks = [];
  let currentWriteRow = WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW;

  productMap.forEach((info, tagCode) => {
    let rowData = new Array(WAREHOUSE_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill("");
    let rowBg = new Array(WAREHOUSE_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill(null);
    const mData = masterMap.get(info.parentCode); 

    for (let id in tgtColMap) {
      const tgtCol = tgtColMap[id];
      if (tgtCol > 26) continue; 
      
      if (id === "15") rowData[tgtCol - 1] = ""; 
      else if (id === "09") rowData[tgtCol - 1] = WAREHOUSE_MATRIX_CONFIG.SIZE_ORDER.filter(s => info.actualSizes.has(s)).join(", ");
      else if (id === "04" && skuColMap["05"]) {
        const photoUrl = info.skuFirstRow[skuColMap["05"] - 1];
        if (photoUrl) rowData[tgtCol - 1] = `=IMAGE("${getDirectImageUrl(photoUrl)}")`;
      } else if (WAREHOUSE_MATRIX_CONFIG.MASTER_PULL_IDS.includes(id) && masterColMap[id] && mData) {
        smartChipCopyTasks.push({ srcRow: mData.rowIndex, srcCol: masterColMap[id], dstRow: currentWriteRow, dstCol: tgtCol });
      } else if (skuColMap[id]) {
        rowData[tgtCol - 1] = info.skuFirstRow[skuColMap[id] - 1];
      }
    }
    
    // カラー設定：在庫ブロック(AA~AF)は薄いグレー（自動計算エリアのため）
    [27, 28, 29, 30, 31, 32].forEach(c => rowBg[c - 1] = "#f3f3f3");
    [33, 40, 47, 54, 61].forEach(c => rowBg[c - 1] = "#fff2cc"); // 合計列は黄色
    WAREHOUSE_MATRIX_CONFIG.SIZE_ORDER.forEach(size => {
      if (!info.actualSizes.has(size)) WAREHOUSE_MATRIX_CONFIG.TARGET.SIZE_COLS[size].forEach(c => rowBg[c - 1] = "#999999");
    });

    outputData.push(rowData);
    outputBackgrounds.push(rowBg);
    currentWriteRow++;
  });
  
  if (outputData.length > 0) {
    const targetLastRow = targetSheet.getLastRow();
    if (targetLastRow >= WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW) {
      clearContentAndProtections(targetSheet.getRange(WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW, 1, targetLastRow - WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW + 1, WAREHOUSE_MATRIX_CONFIG.TARGET.TOTAL_COLS));
    }
    
    const targetRange = targetSheet.getRange(WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW, 1, outputData.length, WAREHOUSE_MATRIX_CONFIG.TARGET.TOTAL_COLS);
    targetRange.setValues(outputData);
    targetRange.setBackgrounds(outputBackgrounds);
    
    smartChipCopyTasks.forEach(task => {
      masterSheet.getRange(task.srcRow, task.srcCol).copyTo(targetSheet.getRange(task.dstRow, task.dstCol));
      targetSheet.getRange(task.dstRow, task.dstCol).setBackground(null);
    });

    const fStart = WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW;
    const finalLastRow = targetSheet.getLastRow();
    const suppColStr = getColumnLetter(tgtColMap["01"] || 2); 
    const colCodeStr = getColumnLetter(tgtColMap["062"]); 
    // --- ★ここから追加（BYROW数式の自動セット） ---
    const colJpy = tgtColMap["15"];
    if (colJpy) {
      const formula = `=BYROW(${getColumnLetter(tgtColMap["13"])}${fStart}:${getColumnLetter(tgtColMap["14"])}${finalLastRow}, LAMBDA(row, IF(INDEX(row, 1, 2)="", "", IF(INDEX(row, 1, 1)="VN", INDEX(row, 1, 2) * $Q$2, IF(INDEX(row, 1, 1)="CN", INDEX(row, 1, 2) * $Q$3, "")))))`;
      targetSheet.getRange(fStart, colJpy).setFormula(formula);
    }
       // 合計列セット
    WAREHOUSE_MATRIX_CONFIG.TARGET.SUM_COLS.forEach(sumConfig => {
      targetSheet.getRange(fStart, sumConfig.col).setFormula(`=BYROW(${sumConfig.startRange}${fStart}:${sumConfig.endRange}${finalLastRow}, LAMBDA(row, SUM(row)))`);
    });

    // 1ブロック目（在庫）のVLOOKUPセット
    let stockFormulas = [];
    for (let r = fStart; r <= finalLastRow; r++) {
      let rowF = new Array(6).fill("");
      Object.keys(WAREHOUSE_MATRIX_CONFIG.TARGET.SIZE_COLS).forEach((s, idx) => {
        // ★SKUシートから「52_実行時在庫」をVLOOKUPで引っ張る（34は仮列番号なので、SKUの実際の在庫列に合わせる）
        rowF[idx] = `=IF($${colCodeStr}${r}="", "", IFERROR(VLOOKUP($${colCodeStr}${r} & "-${s}-" & $${suppColStr}${r}, 'SKU'!$A:$BE, 34, FALSE), 0))`;
      });
      stockFormulas.push(rowF);
    }
    targetSheet.getRange(fStart, 27, stockFormulas.length, 6).setFormulas(stockFormulas); // AA(27)からAF(32)まで

    // ★アキラさん特製：棚卸しブロックの「赤く光る」条件付き書式を自動セット
    const rules = targetSheet.getConditionalFormatRules();
    const rangeToFormat = targetSheet.getRange(`BC${fStart}:BH${finalLastRow}`);
    // ルール：棚卸セル(BC)が空ではなく、かつ在庫セル(AA)と等しくない場合、背景を赤に
    const redAlertRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(BC${fStart}<>"", BC${fStart}<>AA${fStart})`)
      .setBackground("#f8cecc") // 薄い赤
      .setFontColor("#cc0000")
      .setRanges([rangeToFormat])
      .build();
    rules.push(redAlertRule);
    targetSheet.setConditionalFormatRules(rules);
    // --- ★ここから追加：自動計算エリアの「警告」保護セット ---
    // 保護したい範囲を配列で指定（AA列〜AF列の在庫エリア、および各ブロックの合計列）
    const protectRanges = [
      targetSheet.getRange(`AA${fStart}:AF${finalLastRow}`), // 在庫表示ブロック
      targetSheet.getRange(`O${fStart}:O${finalLastRow}`),   // 15_価格などの数式列（あれば）
      targetSheet.getRange(`AG${fStart}:AG${finalLastRow}`), // 各合計列（33列目）
      targetSheet.getRange(`AN${fStart}:AN${finalLastRow}`), // 各合計列（40列目）
      targetSheet.getRange(`AU${fStart}:AU${finalLastRow}`), // 各合計列（47列目）
      targetSheet.getRange(`BB${fStart}:BB${finalLastRow}`), // 各合計列（54列目）
      targetSheet.getRange(`BI${fStart}:BI${finalLastRow}`)  // 各合計列（61列目）
    ];

    protectRanges.forEach(rng => {
      const protection = rng.protect();
      protection.setWarningOnly(true); // ★「警告を表示する」に設定
    });
    // --- ★ここまで追加 ---

    Browser.msgBox("倉庫マトリックス生成完了！棚卸しアラートもセットしました。");
  }
}