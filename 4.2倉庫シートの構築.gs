/**
 * 倉庫マトリックス生成GAS
 * ★アキラさん究極仕様【VaMASTERベース・完全統合版】
 * - 064合鍵による手入力データの完全バックアップ＆復元（消えない！）
 * - VaMASTERからの爆速フレーム構築（SKUはサイズ判定のみに使用）
 * - スマートチップ維持 ＆ 警告アラート保護搭載
 * VaMASTER基点: 型番・バリエーション・サプライヤー単位で、写真付きの正確なマトリックスを構築
 * ID:064合鍵: 再構築前の手入力データ（入出庫数・但し書き等）を自動退避し、正しい行へ完全に復元
 * 自動制御: 存在しないサイズのグレーアウト、在庫参照関数、棚卸しアラート、保護枠を一括で再セット
 * スマート維持: copyTo方式で写真等のチップを保護し、行の追加・削除やソートが発生してもリンクを維持
 */

const WAREHOUSE_MATRIX_CONFIG = {
  SOURCE_SKU: "SKU",       
  SOURCE_VAMASTER: "VaMASTER", // ★MASTERではなく、完成されたVaMASTERから縦軸を作る！
  TARGET_SHEET: "倉庫",
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

// --- ヘルパー関数群 ---
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
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let p of protections) { p.remove(); } // ★今回は警告のみの保護も完全に剥がして新しく作り直す
  sheet.clearConditionalFormatRules();
}

// --- メイン関数 ---
function generateWarehouseMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName(WAREHOUSE_MATRIX_CONFIG.SOURCE_SKU);
  const vaMasterSheet = ss.getSheetByName(WAREHOUSE_MATRIX_CONFIG.SOURCE_VAMASTER);
  const targetSheet = ss.getSheetByName(WAREHOUSE_MATRIX_CONFIG.TARGET_SHEET);
  
  if (!skuSheet || !vaMasterSheet || !targetSheet) return Browser.msgBox("シートが見つかりません。");

  const normalizeId = (h) => String(h).trim().replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/＿/g, "_");
  const getColMap = (sheet) => {
    const map = {};
    sheet.getRange(WAREHOUSE_MATRIX_CONFIG.HEADER_ROW, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].forEach((h, idx) => {
      const match = normalizeId(h).match(/^(\d{2,4})_/);
      if (match) map[match[1]] = idx + 1;
    });
    return map;
  };

  const skuColMap = getColMap(skuSheet);
  const vaColMap = getColMap(vaMasterSheet); 
  const tgtColMap = getColMap(targetSheet);

  if (!tgtColMap["064"] || !vaColMap["064"]) return Browser.msgBox("合鍵となる「064_」列が見つかりません。");

  // ==========================================
  // Step 1: 手入力データの「記憶」（バックアップ）
  // ==========================================
  const backupMap = new Map();
  const targetLastRow = targetSheet.getLastRow();
  
  if (targetLastRow >= WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW) {
    const existingData = targetSheet.getRange(WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW, 1, targetLastRow - WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW + 1, WAREHOUSE_MATRIX_CONFIG.TARGET.TOTAL_COLS).getValues();
    
    existingData.forEach(row => {
      const key064 = String(row[tgtColMap["064"] - 1] || "").trim(); // ★最強の合鍵
      if (key064) {
        const savedInput = {};
        // 関数が入っているブロック以外（手入力部分）をすべて保存する
        for(let c = 27; c <= WAREHOUSE_MATRIX_CONFIG.TARGET.TOTAL_COLS; c++) {
          if (c >= 27 && c <= 32) continue; // 在庫ブロックは関数なので無視
          if ([33, 40, 47, 54, 61].includes(c)) continue; // 合計列も関数なので無視
          
          if (row[c - 1] !== "" && row[c - 1] !== null) {
             savedInput[c] = row[c - 1];
          }
        }
        backupMap.set(key064, savedInput); // 鍵付きのロッカーにしまう
      }
    });
  }

  // ==========================================
  // Step 2: シートの「完全初期化」
  // ==========================================
  if (targetLastRow >= WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW) {
    clearContentAndProtections(targetSheet.getRange(WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW, 1, targetLastRow - WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW + 1, WAREHOUSE_MATRIX_CONFIG.TARGET.TOTAL_COLS));
  }

  // ==========================================
  // Step 3 & 4: VaMASTERから縦軸構築 サイズ判定
  // ==========================================
 // ① VaMASTERを回して「その064コードには、どのサイズが存在するか」を判定する
const sizeExistMap = new Map();

const vaSizeData = vaMasterSheet.getRange(
  WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW,
  1,
  Math.max(vaMasterSheet.getLastRow() - WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW + 1, 1),
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
      if (size === "FREE" || size === "FREES" || size === "FREE SIZE" || size === "フリー") {
        sizeSet.add("F");
      } else {
        sizeSet.add(size);
      }
    });

  sizeExistMap.set(key064, sizeSet);
});

  // ② VaMASTERのデータを読み込んでフレームを作る
  const vaData = vaMasterSheet.getRange(WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW, 1, Math.max(vaMasterSheet.getLastRow() - WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW + 1, 1), vaMasterSheet.getLastColumn()).getValues();
  
  let outputData = [];
  let outputBackgrounds = [];
  let smartChipCopyTasks = [];
  
  vaData.forEach((vaRow, idx) => {
    const key064 = String(vaRow[vaColMap["064"] - 1] || "").trim();
    if (!key064) return;

    let rowData = new Array(WAREHOUSE_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill("");
    let rowBg = new Array(WAREHOUSE_MATRIX_CONFIG.TARGET.TOTAL_COLS).fill(null);
    const actualSizes = sizeExistMap.get(key064) || new Set();

    // 縦軸（左側の基本情報）をVaMASTERから丸写しする
    for (let id in tgtColMap) {
      const tgtCol = tgtColMap[id];
      if (tgtCol > 26) continue; 
      
      if (id === "15") {
         rowData[tgtCol - 1] = ""; // 日本円数式の場所は空ける
      } else if (id === "09") {
         rowData[tgtCol - 1] = WAREHOUSE_MATRIX_CONFIG.SIZE_ORDER.filter(s => actualSizes.has(s)).join(", "); // サイズ
      } else if (id === "04" && vaColMap["05"]) {
         const photoUrl = vaRow[vaColMap["05"] - 1];
         if (photoUrl) rowData[tgtCol - 1] = `=IMAGE("${getDirectImageUrl(photoUrl)}")`;
      } else if (WAREHOUSE_MATRIX_CONFIG.MASTER_PULL_IDS.includes(id) && vaColMap[id]) {
         // ★スマートチップは後で「copyTo」するためにタスクに積む
         smartChipCopyTasks.push({ 
           srcRow: idx + WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW, 
           srcCol: vaColMap[id], 
           dstRow: WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW + outputData.length, 
           dstCol: tgtCol 
         });
      } else if (vaColMap[id]) {
         rowData[tgtCol - 1] = vaRow[vaColMap[id] - 1];
      }
    }

    // ==========================================
    // Step 6: 記憶した手入力データの「復元」
    // ==========================================
    if (backupMap.has(key064)) {
      const savedInput = backupMap.get(key064);
      for (let c in savedInput) {
        rowData[c - 1] = savedInput[c]; // 鍵が一致した行の正しい場所に、入力データを戻す！
      }
    }

    // --- 色塗り設定 ---
    [27, 28, 29, 30, 31, 32].forEach(c => rowBg[c - 1] = "#f3f3f3");
    [33, 40, 47, 54, 61].forEach(c => rowBg[c - 1] = "#fff2cc");
    WAREHOUSE_MATRIX_CONFIG.SIZE_ORDER.forEach(size => {
      // その商品に存在しないサイズ列はグレーで塗りつぶす
      if (!actualSizes.has(size)) WAREHOUSE_MATRIX_CONFIG.TARGET.SIZE_COLS[size].forEach(c => rowBg[c - 1] = "#999999");
    });

    outputData.push(rowData);
    outputBackgrounds.push(rowBg);
  });

  // ==========================================
  // Step 5: 書き込み ＆ 関数・アラート・保護の設定
  // ==========================================
  if (outputData.length > 0) {
    // ★自動行追加機能（行が足りなければ足す）
    const currentMaxRows = targetSheet.getMaxRows();
    const neededMaxRows = WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW + outputData.length - 1;
    if (neededMaxRows > currentMaxRows) {
      targetSheet.insertRowsAfter(currentMaxRows, neededMaxRows - currentMaxRows);
    }

    const targetRange = targetSheet.getRange(WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW, 1, outputData.length, WAREHOUSE_MATRIX_CONFIG.TARGET.TOTAL_COLS);
    targetRange.setValues(outputData);
    targetRange.setBackgrounds(outputBackgrounds);

    // ★VaMASTERからスマートチップをコピー！
    smartChipCopyTasks.forEach(task => {
      vaMasterSheet.getRange(task.srcRow, task.srcCol).copyTo(targetSheet.getRange(task.dstRow, task.dstCol));
      targetSheet.getRange(task.dstRow, task.dstCol).setBackground(null);
    });

    const fStart = WAREHOUSE_MATRIX_CONFIG.DATA_START_ROW;
    const finalLastRow = neededMaxRows;
    const col064Str = getColumnLetter(tgtColMap["064"]); 

    // ① 日本円数式 (BYROW)
    const colJpy = tgtColMap["15"];
    if (colJpy) {
      const formula = `=BYROW(${getColumnLetter(tgtColMap["13"])}${fStart}:${getColumnLetter(tgtColMap["14"])}${finalLastRow}, LAMBDA(row, IF(INDEX(row, 1, 2)="", "", IF(INDEX(row, 1, 1)="VN", INDEX(row, 1, 2) * $Q$2, IF(INDEX(row, 1, 1)="CN", INDEX(row, 1, 2) * $Q$3, "")))))`;
      targetSheet.getRange(fStart, colJpy).setFormula(formula);
    }

    // ② 各ブロックの合計 (BYROW)
    WAREHOUSE_MATRIX_CONFIG.TARGET.SUM_COLS.forEach(sumConfig => {
      targetSheet.getRange(fStart, sumConfig.col).setFormula(`=BYROW(${sumConfig.startRange}${fStart}:${sumConfig.endRange}${finalLastRow}, LAMBDA(row, SUM(row)))`);
    });

    // ③ 在庫のVLOOKUP（SKUから引いてくる）
    let stockFormulas = [];
    for (let r = fStart; r <= finalLastRow; r++) {
      let rowF = new Array(6).fill("");
      Object.keys(WAREHOUSE_MATRIX_CONFIG.TARGET.SIZE_COLS).forEach((s, idx) => {
       rowF[idx] = `=IF($${col064Str}${r}="", "", IFERROR(VLOOKUP($${col064Str}${r} & "-${s}", 'SKU'!$A:$BE, 34, FALSE), ""))`;
      });
      stockFormulas.push(rowF);
    }
    targetSheet.getRange(fStart, 27, stockFormulas.length, 6).setFormulas(stockFormulas); 

    // ④ 棚卸し赤光りアラート（条件付き書式）
    const rules = targetSheet.getConditionalFormatRules();
    const rangeToFormat = targetSheet.getRange(`BC${fStart}:BH${finalLastRow}`);
    const redAlertRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(BC${fStart}<>"", BC${fStart}<>AA${fStart})`)
      .setBackground("#f8cecc")
      .setFontColor("#cc0000")
      .setRanges([rangeToFormat])
      .build();
    rules.push(redAlertRule);
    targetSheet.setConditionalFormatRules(rules);

    const INPUT_PROTECT_COLS = {
      XS: [34, 41, 48, 55],
      S:  [35, 42, 49, 56],
      M:  [36, 43, 50, 57],
      L:  [37, 44, 51, 58],
      XL: [38, 45, 52, 59],
      F:  [39, 46, 53, 60]
    };

          // ⑤ 絶対にサボらない「警告」保護
    const protectRanges = [
      targetSheet.getRange(`AA${fStart}:AF${finalLastRow}`), 

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
              // 存在しないサイズのグレーセルにも警告保護を追加
      outputData.forEach((row, i) => {
        const sheetRow = fStart + i;
        const key064 = String(row[tgtColMap["064"] - 1] || "").trim();
        const actualSizes = sizeExistMap.get(key064) || new Set();

      WAREHOUSE_MATRIX_CONFIG.SIZE_ORDER.forEach(size => {
        if (!actualSizes.has(size)) {
      INPUT_PROTECT_COLS[size].forEach(col => {
        protectRanges.push(targetSheet.getRange(sheetRow, col));
      });
    }
  });
});
    protectRanges.forEach(rng => {
      rng.protect().setWarningOnly(true);
    });

    Browser.msgBox("倉庫マトリックス生成完了！\\n手入力データの復元、棚卸しアラート、保護、すべて完璧にセットしました！");
  } else {
    Browser.msgBox("展開するデータがありませんでした。");
  }
}