/**
 * 移動販売マトリックス生成GAS
 * ★6ブロック構成 ＆ 棚卸アラート ＆ 最強保護システム搭載
 */

const MOBILE_MATRIX_CONFIG = {
  SOURCE_SKU: "SKU",       
  SOURCE_MASTER: "MASTER",  
  TARGET_SHEET: "移動販売", // ★ターゲット
  HEADER_ROW: 6,
  DATA_START_ROW: 7,            
  MASTER_PULL_IDS: ["20", "21", "22"], 

  TARGET: {
    TOTAL_COLS: 68, // BP列まで（6ブロック）
    SIZE_COLS: {
      "XS": [27, 34, 41, 48, 55, 62], // 倉庫, 販売在庫, 取寄せ, 販売, 不良, 棚卸
      "S":  [28, 35, 42, 49, 56, 63], 
      "M":  [29, 36, 43, 50, 57, 64], 
      "L":  [30, 37, 44, 51, 58, 65], 
      "XL": [31, 38, 45, 52, 59, 66], 
      "F":  [32, 39, 46, 53, 60, 67]  
    },
    SUM_COLS: [
      { col: 33, startRange: "AA", endRange: "AF" }, // B1: 倉庫在庫 合計
      { col: 40, startRange: "AH", endRange: "AM" }, // B2: 販売在庫 合計
      { col: 47, startRange: "AO", endRange: "AT" }, // B3: 取寄せ 合計
      { col: 54, startRange: "AV", endRange: "BA" }, // B4: 販売 合計
      { col: 61, startRange: "BC", endRange: "BH" }, // B5: 不良 合計
      { col: 68, startRange: "BJ", endRange: "BO" }  // B6: 棚卸 合計
    ]
  },
  SIZE_ORDER: ["XS", "S", "M", "L", "XL", "F"]
};

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
  for (let p of protections) { p.remove(); } // ★最強消しゴム
  sheet.clearConditionalFormatRules(); 
}

function generateMobileSalesMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName(MOBILE_MATRIX_CONFIG.SOURCE_SKU);
  const masterSheet = ss.getSheetByName(MOBILE_MATRIX_CONFIG.SOURCE_MASTER); 
  const targetSheet = ss.getSheetByName(MOBILE_MATRIX_CONFIG.TARGET_SHEET);
  
  if (!skuSheet || !masterSheet || !targetSheet) return Browser.msgBox("シートが見つかりません。");

  const normalizeId = (h) => String(h).trim().replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/＿/g, "_");
  const getColMap = (sheet) => {
    const map = {};
    sheet.getRange(MOBILE_MATRIX_CONFIG.HEADER_ROW, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].forEach((h, idx) => {
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
  const skuData = skuSheet.getRange(MOBILE_MATRIX_CONFIG.DATA_START_ROW, 1, skuLastRow - MOBILE_MATRIX_CONFIG.DATA_START_ROW + 1, skuSheet.getLastColumn()).getValues();

  const masterLastRow = masterSheet.getLastRow();
  const masterData = masterLastRow < MOBILE_MATRIX_CONFIG.DATA_START_ROW ? [] : masterSheet.getRange(MOBILE_MATRIX_CONFIG.DATA_START_ROW, 1, masterLastRow - MOBILE_MATRIX_CONFIG.DATA_START_ROW + 1, masterSheet.getLastColumn()).getValues();

  const masterMap = new Map();
  masterData.forEach((row, idx) => {
    const parentCode = String(row[masterColMap["06"] - 1] || "").trim();
    if (parentCode && !masterMap.has(parentCode)) masterMap.set(parentCode, { values: row, rowIndex: idx + MOBILE_MATRIX_CONFIG.DATA_START_ROW });
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
  let currentWriteRow = MOBILE_MATRIX_CONFIG.DATA_START_ROW;

  productMap.forEach((info, tagCode) => {
    let rowData = new Array(MOBILE_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill("");
    let rowBg = new Array(MOBILE_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill(null);
    const mData = masterMap.get(info.parentCode); 

    for (let id in tgtColMap) {
      const tgtCol = tgtColMap[id];
      if (tgtCol > 26) continue; 
      
      if (id === "15") rowData[tgtCol - 1] = ""; 
      else if (id === "09") rowData[tgtCol - 1] = MOBILE_MATRIX_CONFIG.SIZE_ORDER.filter(s => info.actualSizes.has(s)).join(", ");
      else if (id === "04" && skuColMap["05"]) {
        const photoUrl = info.skuFirstRow[skuColMap["05"] - 1];
        if (photoUrl) rowData[tgtCol - 1] = `=IMAGE("${getDirectImageUrl(photoUrl)}")`;
      } else if (MOBILE_MATRIX_CONFIG.MASTER_PULL_IDS.includes(id) && masterColMap[id] && mData) {
        smartChipCopyTasks.push({ srcRow: mData.rowIndex, srcCol: masterColMap[id], dstRow: currentWriteRow, dstCol: tgtCol });
      } else if (skuColMap[id]) {
        rowData[tgtCol - 1] = info.skuFirstRow[skuColMap[id] - 1];
      }
    }
    
    // カラー設定：B1(倉庫在庫)とB2(販売在庫)は自動表示なのでグレー
    [27, 28, 29, 30, 31, 32].forEach(c => rowBg[c - 1] = "#f3f3f3");
    [34, 35, 36, 37, 38, 39].forEach(c => rowBg[c - 1] = "#e8eaed");
    [33, 40, 47, 54, 61, 68].forEach(c => rowBg[c - 1] = "#fff2cc"); // 合計列
    MOBILE_MATRIX_CONFIG.SIZE_ORDER.forEach(size => {
      if (!info.actualSizes.has(size)) MOBILE_MATRIX_CONFIG.TARGET.SIZE_COLS[size].forEach(c => rowBg[c - 1] = "#999999");
    });

    outputData.push(rowData);
    outputBackgrounds.push(rowBg);
    currentWriteRow++;
  });
  
  if (outputData.length > 0) {
    const targetLastRow = targetSheet.getLastRow();
    if (targetLastRow >= MOBILE_MATRIX_CONFIG.DATA_START_ROW) {
      clearContentAndProtections(targetSheet.getRange(MOBILE_MATRIX_CONFIG.DATA_START_ROW, 1, targetLastRow - MOBILE_MATRIX_CONFIG.DATA_START_ROW + 1, MOBILE_MATRIX_CONFIG.TARGET.TOTAL_COLS));
    }
    
    const targetRange = targetSheet.getRange(MOBILE_MATRIX_CONFIG.DATA_START_ROW, 1, outputData.length, MOBILE_MATRIX_CONFIG.TARGET.TOTAL_COLS);
    targetRange.setValues(outputData);
    targetRange.setBackgrounds(outputBackgrounds);
    
    smartChipCopyTasks.forEach(task => {
      masterSheet.getRange(task.srcRow, task.srcCol).copyTo(targetSheet.getRange(task.dstRow, task.dstCol));
      targetSheet.getRange(task.dstRow, task.dstCol).setBackground(null);
    });

    const fStart = MOBILE_MATRIX_CONFIG.DATA_START_ROW;
    const finalLastRow = targetSheet.getLastRow();
    const suppColStr = getColumnLetter(tgtColMap["01"] || 2); 
    const colCodeStr = getColumnLetter(tgtColMap["062"]); 
    const colJpy = tgtColMap["15"];

    try {
      if (colJpy) {
        const formula = `=BYROW(${getColumnLetter(tgtColMap["13"])}${fStart}:${getColumnLetter(tgtColMap["14"])}${finalLastRow}, LAMBDA(row, IF(INDEX(row, 1, 2)="", "", IF(INDEX(row, 1, 1)="VN", INDEX(row, 1, 2) * $Q$2, IF(INDEX(row, 1, 1)="CN", INDEX(row, 1, 2) * $Q$3, "")))))`;
        targetSheet.getRange(fStart, colJpy).setFormula(formula);
      }

      MOBILE_MATRIX_CONFIG.TARGET.SUM_COLS.forEach(sumConfig => {
        targetSheet.getRange(fStart, sumConfig.col).setFormula(`=BYROW(${sumConfig.startRange}${fStart}:${sumConfig.endRange}${finalLastRow}, LAMBDA(row, SUM(row)))`);
      });

      // 1ブロック目（倉庫在庫: 34_）と 2ブロック目（販売在庫: 60_）のVLOOKUPセット
      let stockFormulas = [];
      for (let r = fStart; r <= finalLastRow; r++) {
        let rowF = new Array(13).fill(""); // AA(27)からAM(39)まで（合計列AG(33)を含む）
        Object.keys(MOBILE_MATRIX_CONFIG.TARGET.SIZE_COLS).forEach((s, idx) => {
          // B1: 倉庫在庫（SKUの左から34列目と仮定）
          rowF[idx] = `=IF($${colCodeStr}${r}="", "", IFERROR(VLOOKUP($${colCodeStr}${r} & "-${s}-" & $${suppColStr}${r}, 'SKU'!$A:$BP, 34, FALSE), 0))`;
          // B2: 販売在庫（SKUの60_の列：SKUシートの列配置に合わせてVLOOKUPの列番号「42」等は調整が必要です ※一旦ダミーで列ID検索できる関数を入れます）
          // ※VLOOKUPだと列がズレるため、今回は販売在庫も確実に引っ張れるように INDEX & MATCH に似せた構成がベストですが、
          // シンプルにSKUの「AP列」と指定いただいているので、AP列は「42列目」として設定します。
          rowF[idx + 7] = `=IF($${colCodeStr}${r}="", "", IFERROR(VLOOKUP($${colCodeStr}${r} & "-${s}-" & $${suppColStr}${r}, 'SKU'!$A:$BP, 42, FALSE), 0))`; 
        });
        rowF[6] = `=BYROW(AA${r}:AF${r}, LAMBDA(row, SUM(row)))`; // 合計列AG
        stockFormulas.push(rowF);
      }
      targetSheet.getRange(fStart, 27, stockFormulas.length, 13).setFormulas(stockFormulas); 

      // 赤アラート（B6:棚卸 BJ列 と B2:販売在庫 AH列 の比較）
      const rules = targetSheet.getConditionalFormatRules();
      const rangeToFormat = targetSheet.getRange(`BJ${fStart}:BO${finalLastRow}`);
      const redAlertRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(BJ${fStart}<>"", BJ${fStart}<>AH${fStart})`)
        .setBackground("#f8cecc") 
        .setFontColor("#cc0000")
        .setRanges([rangeToFormat])
        .build();
      rules.push(redAlertRule);
      targetSheet.setConditionalFormatRules(rules);
    } catch (e) { Logger.log(e.message); }

    // =================================================================
    // --- ★絶対にサボらない「警告」保護セット ---
    // =================================================================
    try {
      SpreadsheetApp.flush(); 
      const protectA1Notations = [
        `O${fStart}:O${finalLastRow}`,   // 15_価格列
        `AA${fStart}:AM${finalLastRow}`, // B1(倉庫在庫) ＆ B2(販売在庫) の表示エリア全て
        `AG${fStart}:AG${finalLastRow}`, // 合計
        `AN${fStart}:AN${finalLastRow}`, // 合計
        `AU${fStart}:AU${finalLastRow}`, // 合計
        `BB${fStart}:BB${finalLastRow}`, // 合計
        `BI${fStart}:BI${finalLastRow}`, // 合計
        `BP${fStart}:BP${finalLastRow}`  // 合計
      ];
      protectA1Notations.forEach(a1 => {
        const rng = targetSheet.getRange(a1);
        if (rng) rng.protect().setWarningOnly(true);
      });
      SpreadsheetApp.flush(); 
    } catch (error) {
      Browser.msgBox("【エラー報告】保護の設定で失敗しました：\\n" + error.message);
    }

    Browser.msgBox("移動販売マトリックス生成完了！6ブロックと保護をセットしました。");
  }
}