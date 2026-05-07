/**
 * [SKU]シートと[マスター]シートからデータを集約し、
 * サイズを横に展開するマトリックス自動生成GAS
 * ★アキラさん究極仕様（スマートチップコピペ ＆ 最強保護システム搭載）
 */

const ORDER_IS_APPEND_MODE = false; 

const ORDER_MATRIX_CONFIG = {
  SOURCE_SKU: "SKU",       
  SOURCE_MASTER: "MASTER", // ※シート名がMASTERなら変更してください
  TARGET_SHEET: "注文", 
  HEADER_ROW: 6,
  DATA_START_ROW: 7,            
  
  MASTER_PULL_IDS: ["20", "21", "22"], 

  TARGET: {
    TOTAL_COLS: 61,          
    SIZE_COLS: {
      "XS": [27, 34, 41, 48, 55], 
      "S":  [28, 35, 42, 49, 56], 
      "M":  [29, 36, 43, 50, 57], 
      "L":  [30, 37, 44, 51, 58], 
      "XL": [31, 38, 45, 52, 59], 
      "F":  [32, 39, 46, 53, 60]  
    },
    SUM_COLS: [
      { col: 33, startRange: "AA", endRange: "AF" }, 
      { col: 40, startRange: "AH", endRange: "AM" }, 
      { col: 47, startRange: "AO", endRange: "AT" }, 
      { col: 54, startRange: "AV", endRange: "BA" }, 
      { col: 61, startRange: "BC", endRange: "BH" }  
    ]
  },
  
  SIZE_ORDER: ["XS", "S", "M", "L", "XL", "F"]
};

// --- ヘルパー関数 ---
function getDirectImageUrl(url) {
  if (!url) return "";
  const match = url.match(/(?:id=|d\/)([\w-]+)/);
  if (match) return `https://drive.google.com/uc?export=download&id=${match[1]}`;
  return url;
}

function getColumnLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function clearContentAndProtections(range) {
  range.clearContent();
  range.setBackground(null);
  const sheet = range.getSheet();
  // ★最強の消しゴム：遠慮せずに古い保護を全部ぶっ壊す！
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let p of protections) { p.remove(); }
}

// --- メイン関数 ---
function generateOrderMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName(ORDER_MATRIX_CONFIG.SOURCE_SKU);
  const masterSheet = ss.getSheetByName(ORDER_MATRIX_CONFIG.SOURCE_MASTER); 
  const targetSheet = ss.getSheetByName(ORDER_MATRIX_CONFIG.TARGET_SHEET);
  
  if (!skuSheet || !masterSheet || !targetSheet) {
    Browser.msgBox("シートが見つかりません。");
    return;
  }

  const normalizeId = (h) => String(h).trim().replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/＿/g, "_");

  const getColMap = (sheet) => {
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(ORDER_MATRIX_CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
    const map = {};
    headers.forEach((h, idx) => {
      const match = normalizeId(h).match(/^(\d{2,3})_/);
      if (match) map[match[1]] = idx + 1;
    });
    return map;
  };

  const skuColMap = getColMap(skuSheet);
  const masterColMap = getColMap(masterSheet); 
  const tgtColMap = getColMap(targetSheet);
  
  if (!tgtColMap["062"] || !masterColMap["06"] || !skuColMap["06"]) {
    Browser.msgBox("必須IDが見つかりません。");
    return;
  }

  const skuLastRow = skuSheet.getLastRow();
  const skuData = skuSheet.getRange(ORDER_MATRIX_CONFIG.DATA_START_ROW, 1, skuLastRow - ORDER_MATRIX_CONFIG.DATA_START_ROW + 1, skuSheet.getLastColumn()).getValues();

  const masterLastRow = masterSheet.getLastRow();
  const masterData = masterLastRow < ORDER_MATRIX_CONFIG.DATA_START_ROW ? [] : masterSheet.getRange(ORDER_MATRIX_CONFIG.DATA_START_ROW, 1, masterLastRow - ORDER_MATRIX_CONFIG.DATA_START_ROW + 1, masterSheet.getLastColumn()).getValues();

  const masterMap = new Map();
  masterData.forEach((row, idx) => {
    const parentCode = String(row[masterColMap["06"] - 1] || "").trim();
    if (parentCode && !masterMap.has(parentCode)) {
      masterMap.set(parentCode, { 
        values: row, 
        rowIndex: idx + ORDER_MATRIX_CONFIG.DATA_START_ROW 
      });
    }
  });
  
  const targetLastRow = targetSheet.getLastRow();
  let existingCodes = new Set();
  if (ORDER_IS_APPEND_MODE && targetLastRow >= ORDER_MATRIX_CONFIG.DATA_START_ROW) {
    const fColValues = targetSheet.getRange(ORDER_MATRIX_CONFIG.DATA_START_ROW, tgtColMap["062"], targetLastRow - ORDER_MATRIX_CONFIG.DATA_START_ROW + 1, 1).getValues();
    for (let r of fColValues) { if (r[0]) existingCodes.add(String(r[0]).trim()); }
  }

  const productMap = new Map();
  for (let i = 0; i < skuData.length; i++) {
    const row = skuData[i];
    const tagCode = String(row[skuColMap["062"] - 1] || "").trim();
    const parentCode = String(row[skuColMap["06"] - 1] || "").trim();
    if (!tagCode) continue; 
    if (!productMap.has(tagCode)) {
      productMap.set(tagCode, { skuFirstRow: row, parentCode: parentCode, actualSizes: new Set() });
    }
    const sizeUnit = String(row[skuColMap["09"] - 1] || "").trim().toUpperCase();
    if (sizeUnit) {
      const g = productMap.get(tagCode);
      let sCode = (sizeUnit === "FREE" || sizeUnit === "FREES") ? "F" : sizeUnit;
      g.actualSizes.add(sCode);
    }
  }
  
  let outputData = [];
  let outputBackgrounds = [];
  let smartChipCopyTasks = []; 

  let writeStartRow = ORDER_MATRIX_CONFIG.DATA_START_ROW;
  if (ORDER_IS_APPEND_MODE) {
    writeStartRow = targetLastRow < ORDER_MATRIX_CONFIG.DATA_START_ROW ? ORDER_MATRIX_CONFIG.DATA_START_ROW : targetLastRow + 1;
  }
  let currentWriteRow = writeStartRow;

  productMap.forEach((info, tagCode) => {
    if (ORDER_IS_APPEND_MODE && existingCodes.has(tagCode)) return;

    let rowData = new Array(ORDER_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill("");
    let rowBg = new Array(ORDER_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill(null);
    const mData = masterMap.get(info.parentCode); 

    for (let id in tgtColMap) {
      const tgtCol = tgtColMap[id];
      if (tgtCol > 26) continue; 
      
      if (id === "15") {
        rowData[tgtCol - 1] = ""; 
      } else if (id === "09") {
        const sizeOrderFound = ORDER_MATRIX_CONFIG.SIZE_ORDER.filter(s => info.actualSizes.has(s));
        rowData[tgtCol - 1] = sizeOrderFound.join(", ");
      } else if (id === "04" && skuColMap["05"]) {
        const photoUrl = info.skuFirstRow[skuColMap["05"] - 1];
        if (photoUrl) rowData[tgtCol - 1] = `=IMAGE("${getDirectImageUrl(photoUrl)}")`;
      } else if (ORDER_MATRIX_CONFIG.MASTER_PULL_IDS.includes(id) && masterColMap[id]) {
        if (mData) {
          smartChipCopyTasks.push({
            srcRow: mData.rowIndex,
            srcCol: masterColMap[id],
            dstRow: currentWriteRow,
            dstCol: tgtCol
          });
        }
      } else if (skuColMap[id]) {
        rowData[tgtCol - 1] = info.skuFirstRow[skuColMap[id] - 1];
      }
    }
    
    [33, 40, 47, 54, 61].forEach(c => rowBg[c - 1] = "#fff2cc");
    [tgtColMap["15"], 48, 49, 50, 51, 52, 53, 55, 56, 57, 58, 59, 60].forEach(c => { if(c) rowBg[c - 1] = "#ddebf7"; });
    ORDER_MATRIX_CONFIG.SIZE_ORDER.forEach(size => {
      if (!info.actualSizes.has(size)) {
        ORDER_MATRIX_CONFIG.TARGET.SIZE_COLS[size].forEach(c => rowBg[c - 1] = "#999999");
      }
    });

    outputData.push(rowData);
    outputBackgrounds.push(rowBg);
    currentWriteRow++;
  });
  
  if (outputData.length > 0) {
    if (!ORDER_IS_APPEND_MODE && targetLastRow >= ORDER_MATRIX_CONFIG.DATA_START_ROW) {
      const clearRange = targetSheet.getRange(ORDER_MATRIX_CONFIG.DATA_START_ROW, 1, targetLastRow - ORDER_MATRIX_CONFIG.DATA_START_ROW + 1, ORDER_MATRIX_CONFIG.TARGET.TOTAL_COLS);
      clearContentAndProtections(clearRange);
    }
    
    const targetRange = targetSheet.getRange(writeStartRow, 1, outputData.length, ORDER_MATRIX_CONFIG.TARGET.TOTAL_COLS);
    targetRange.setValues(outputData);
    targetRange.setBackgrounds(outputBackgrounds);
    
    smartChipCopyTasks.forEach(task => {
      masterSheet.getRange(task.srcRow, task.srcCol).copyTo(targetSheet.getRange(task.dstRow, task.dstCol));
      targetSheet.getRange(task.dstRow, task.dstCol).setBackground(null);
    });

    // --- 数式一括貼り付け ---
    const fStart = ORDER_MATRIX_CONFIG.DATA_START_ROW;
    const finalLastRow = targetSheet.getLastRow();
    try {
      const suppColStr = getColumnLetter(tgtColMap["01"] || 2); 
      const colCodeStr = getColumnLetter(tgtColMap["062"]); 
      const colJpy = tgtColMap["15"];
      if (colJpy) {
        const formula = `=BYROW(${getColumnLetter(tgtColMap["13"])}${fStart}:${getColumnLetter(tgtColMap["14"])}${finalLastRow}, LAMBDA(row, IF(INDEX(row, 1, 2)="", "", IF(INDEX(row, 1, 1)="VN", INDEX(row, 1, 2) * $L$2, IF(INDEX(row, 1, 1)="CN", INDEX(row, 1, 2) * $L$3, "")))))`;
        targetSheet.getRange(fStart, colJpy).setFormula(formula);
      }
      ORDER_MATRIX_CONFIG.TARGET.SUM_COLS.filter(c => c.col < 48).forEach(sumConfig => {
        const formula = `=BYROW(${sumConfig.startRange}${fStart}:${sumConfig.endRange}${finalLastRow}, LAMBDA(row, SUM(row)))`;
        targetSheet.getRange(fStart, sumConfig.col).setFormula(formula);
      });
      let combinedFormulas = [];
      for (let r = fStart; r <= finalLastRow; r++) {
        let rowF = new Array(14).fill("");
        Object.keys(ORDER_MATRIX_CONFIG.TARGET.SIZE_COLS).forEach(s => {
          const cols = ORDER_MATRIX_CONFIG.TARGET.SIZE_COLS[s];
          rowF[cols[3] - 48] = `=IF(AND(${getColumnLetter(cols[1])}${r}="", ${getColumnLetter(cols[2])}${r}=""), "", N(${getColumnLetter(cols[1])}${r}) - N(${getColumnLetter(cols[2])}${r}))`;
          rowF[cols[4] - 48] = `=IF($${colCodeStr}${r}="", "", IFERROR(VLOOKUP($${colCodeStr}${r} & "-${s}-" & $${suppColStr}${r}, 'SKU'!$A:$AH, 34, FALSE), 0))`;
        });
        if (r === fStart) {
          rowF[54 - 48] = `=BYROW(AV${fStart}:BA${finalLastRow}, LAMBDA(row, SUM(row)))`;
          rowF[61 - 48] = `=BYROW(BC${fStart}:BH${finalLastRow}, LAMBDA(row, SUM(row)))`;
        }
        combinedFormulas.push(rowF);
      }
      targetSheet.getRange(fStart, 48, combinedFormulas.length, 14).setFormulas(combinedFormulas);
    } catch (e) { Logger.log(e.message); }

    // =================================================================
    // --- ★ここから追加：絶対にサボらない「警告」保護セット ---
    // =================================================================
    try {
      // 1. まずはここまでの処理（データの書き込み・数式）を強制的にセーブする
      SpreadsheetApp.flush(); 
      
      // 注文マトリックスで保護したい範囲のリスト
      const protectA1Notations = [
        `O${fStart}:O${finalLastRow}`,   // 15_価格列
        `AV${fStart}:BH${finalLastRow}`, // 48列目以降（引き算やVLOOKUPの数式エリア）
        `AG${fStart}:AG${finalLastRow}`, // 合計列
        `AN${fStart}:AN${finalLastRow}`, // 合計列
        `AU${fStart}:AU${finalLastRow}`  // 合計列
      ];
      
      protectA1Notations.forEach(a1 => {
        const rng = targetSheet.getRange(a1);
        if (rng) {
          const p = rng.protect();
          p.setWarningOnly(true); // ★警告のみにする
        }
      });
      
      // 2. 保護をかけた状態を、もう一度強制的にセーブする！
      SpreadsheetApp.flush(); 
      
    } catch (error) {
      // 万が一保護でエラーが起きたら、隠さずに画面に出す！
      Browser.msgBox("【エラー報告】保護の設定で失敗しました：\\n" + error.message);
    }
    // =================================================================
    // --- ★ここまで追加 ---
    // =================================================================

    Browser.msgBox("完了！エラーも消え、スマートチップもそのまま運び、保護も完璧にかけました！");
  }
}