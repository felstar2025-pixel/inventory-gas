/**
 * [MASTER]から[SKU]シートを作成・追記するGAS
 * ★【Ver_21列対応_列指定ソート版】
 * - B列:「サプライヤー(01_)」
 * - I列:「TikTok商品名(17_)」を追加した全21列仕様
 */

const ACTUAL_IS_APPEND_MODE_SKU = false;

const ACTUAL_SKU_GEN_CONFIG = {
  MASTER: {
    SHEET_NAME: "MASTER",       
    HEADER_ROW: 6,              
    START_ROW: 7,               
  },
  
  // ★ I列にTikTok商品名を追加した、全21列の構成
  SKU: {
    SHEET_NAME: "SKU",   
    START_ROW: 7,               
    COL_FULL_CODE: 1,           // A列 (061) 完全SKUコード
    COL_SUPPLIER: 2,            // B列 (01)  サプライヤー
    COL_BRAND: 3,               // C列 (02)  ブランド
    COL_SITE_URL: 4,            // D列 (03)  サイトURL
    COL_PHOTO: 5,               // E列 (04)  写真表示
    COL_PHOTO_URL: 6,           // F列 (05)  写真URL
    COL_BASE_CODE: 7,           // G列 (06)  ベース型番
    COL_TAG_CODE: 8,            // H列 (062) タグ用コード
    COL_TT_NAME: 9,             // I列 (17)  TikTok商品名 ★NEW!!
    COL_ORIGINAL_NAME: 10,      // J列 (07)  元名
    COL_ENGLISH_NAME: 11,       // K列 (08)  英名
    COL_JAPANESE_NAME: 12,      // L列 (16)  日名
    COL_ORIGINAL_FULL: 13,      // M列 (101) 元名フル
    COL_ENGLISH_FULL: 14,       // N列 (111) 英名フル
    COL_JAPANESE_FULL: 15,      // O列 (121) 日名フル
    COL_VAR_ONLY: 16,           // P列 (11)  英バリエ
    COL_VAR_JP_ONLY: 17,        // Q列 (12)  日バリエ
    COL_SIZE: 18,               // R列 (09)  サイズ
    COL_COUNTRY: 19,            // S列 (13)  国
    COL_LOCAL_PRICE: 20,        // T列 (14)  現地価格
    COL_JPY_PRICE: 21,          // U列 (15)  日本円
    COL_TT_PRICE: 22,           // V列 (22) TikTok価格
    COL_MOBILE_PRICE: 23,       // W列 (23) 移動販売価格
    COL_RENTAL_PRICE: 24,       // X列 (24) レンタル価格
    COL_DIFFERENCE: 25,         // Y列 (2000) 入庫ｰ全個数 差異チェック
    COL_TOTAL_PIECES: 26,       // Z列 (2001) 全在庫合計
    COL_CUMULATIVE_SALES: 27    // AA列 (2002) 累計販売数

  }
};

function getDirectImageUrl(url) {
  if (!url) return "";
  const match = url.match(/(?:id=|d\/)([\w-]+)/);
  if (match) {
    return `https://drive.google.com/uc?export=download&id=${match[1]}`;
  }
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

function importMasterDataToSKU_V4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(ACTUAL_SKU_GEN_CONFIG.MASTER.SHEET_NAME);
  const skuSheet = ss.getSheetByName(ACTUAL_SKU_GEN_CONFIG.SKU.SHEET_NAME);
  
  if (!masterSheet || !skuSheet) {
    Browser.msgBox("エラー：シート名が見つかりません。");
    return;
  }

  const lastRowMaster = masterSheet.getLastRow();
  if (lastRowMaster < ACTUAL_SKU_GEN_CONFIG.MASTER.START_ROW) {
    Browser.msgBox("エラー：MASTERシートにデータがありません。");
    return;
  }

  const normalizeId = (h) => {
    return String(h).trim()
      .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
      .replace(/＿/g, "_");
  };

  const lastColMaster = masterSheet.getLastColumn();
  const headerValues = masterSheet.getRange(ACTUAL_SKU_GEN_CONFIG.MASTER.HEADER_ROW, 1, 1, lastColMaster).getValues()[0];
  const colMap = {};
  
  for (let i = 0; i < headerValues.length; i++) {
    const cellValue = normalizeId(headerValues[i]);
    const match = cellValue.match(/^(\d{2,3})_/); 
    if (match) {
      const id = match[1]; 
      colMap[id] = i + 1;
    }
  }

  // ★ 17番(TikTok商品名)も必須チェックに追加しました
  const requiredIds = ["01", "02", "03", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "16", "17"];
  for (let id of requiredIds) {
    if (!colMap[id]) {
      Browser.msgBox("エラー：MASTERシートの見出しに必須ID「" + id + "_」が見つかりません。");
      return;
    }
  }
  
  const skuLastRow = skuSheet.getLastRow();
  const skuLastCol = skuSheet.getLastColumn();
  let existingStockMap = new Map();
  let existingCodesInSheet = new Set();
  
  if (skuLastRow >= ACTUAL_SKU_GEN_CONFIG.SKU.START_ROW) {
    const skuExistingData = skuSheet.getRange(ACTUAL_SKU_GEN_CONFIG.SKU.START_ROW, 1, skuLastRow - ACTUAL_SKU_GEN_CONFIG.SKU.START_ROW + 1, Math.max(skuLastCol, 1)).getValues();
    for (let r of skuExistingData) {
      const code = String(r[ACTUAL_SKU_GEN_CONFIG.SKU.COL_FULL_CODE - 1] || "").trim(); 
      if (code !== "") {
        const stockData = r.slice(ACTUAL_SKU_GEN_CONFIG.SKU.COL_CUMULATIVE_SALES); 
        existingStockMap.set(code, stockData);
        existingCodesInSheet.add(code);
      }
    }
  }

  const masterData = masterSheet.getRange(1, 1, lastRowMaster, lastColMaster).getValues();
  let outputData = [];
  
  for (let i = ACTUAL_SKU_GEN_CONFIG.MASTER.START_ROW - 1; i < masterData.length; i++) {
    const row = masterData[i];
    
    const supplierData = String(row[colMap["01"] - 1] || "");
    const brandName    = String(row[colMap["02"] - 1] || "");
    const siteUrl      = String(row[colMap["03"] - 1] || "");
    const rawPhotoUrl  = String(row[colMap["05"] - 1] || "");
    const baseCode     = String(row[colMap["06"] - 1] || "");
    const originalName = String(row[colMap["07"] - 1] || "");
    const englishName  = String(row[colMap["08"] - 1] || "");
    const jpName       = String(row[colMap["16"] - 1] || "");
    const ttName       = String(row[colMap["17"] - 1] || ""); // ★ NEW: 17番の取得
    const rawSizes     = String(row[colMap["09"] - 1] || "");
    const originalVar  = String(row[colMap["10"] - 1] || "");
    const englishVar   = String(row[colMap["11"] - 1] || "");
    const jpVar        = String(row[colMap["12"] - 1] || "");
    const country      = String(row[colMap["13"] - 1] || "");
    const price        = row[colMap["14"] - 1] || "";
    const ttPrice      = row[colMap["1000"] - 1] || ""; // TikTok価格(ID:1000)
    const mobilePrice  = row[colMap["1001"] - 1] || ""; // 移動販売価格(ID:1001)
    const rentalPrice  = row[colMap["1002"] - 1] || ""; // レンタル価格(ID:1002)
    const difference   = row[colMap["2000"] - 1] || ""; // 入庫ｰ全個数 差異チェック(2000) 
    const totalpieces  = row[colMap["2001"] - 1] || ""; // 全在庫合計(2001)
    const cumulativesales  = row[colMap["2002"] - 1] || ""; // 累計販売数(2002) 

    if (!baseCode) continue;

    let suppCode = "";
    if (supplierData.includes(":") || supplierData.includes("：")) {
      suppCode = supplierData.split(/[:：]/)[0].trim();
    } else if (supplierData.includes(",") || supplierData.includes("、")) {
      suppCode = supplierData.split(/[,、]/)[0].trim(); 
    } else {
      suppCode = supplierData.trim();
    }

    const photoUrl = rawPhotoUrl.split(/[,、\n\s]+/)[0] || ""; 
    
    let engVars = englishVar.split(/[,、\n]+/).map(v => v.trim()).filter(v => v !== "");
    if (engVars.length === 0) engVars.push("");
    let origVars = originalVar.split(/[,、\n]+/).map(v => v.trim()).filter(v => v !== "");
    if (origVars.length === 0) origVars.push("");
    let jpnVars = jpVar.split(/[,、\n]+/).map(v => v.trim()).filter(v => v !== "");
    if (jpnVars.length === 0) jpnVars.push("");
    
    let sizes = rawSizes.split(/[,、\s\n]+/).map(s => s.trim()).filter(s => s !== "");
    if (sizes.length === 0) sizes.push(""); 
    
    for (let vIndex = 0; vIndex < engVars.length; vIndex++) {
      let evar = engVars[vIndex] || "";
      let ovar = origVars[vIndex] || ""; 
      let jvar = jpnVars[vIndex] || ""; 
      
      let vCode = evar;
      let evName = evar;
      let ovName = ovar;
      let jvName = jvar;

      let matchE = evar.match(/^([A-ZNS])[:：]\s*(.*)/i);
      if (matchE) {
        vCode = matchE[1].trim();
        evName = matchE[2].trim();
      }

      let matchO = ovar.match(/^([A-ZNS])[:：]\s*(.*)/i);
      if (matchO) {
        ovName = matchO[2].trim();
      }

      let matchJ = jvar.match(/^([A-ZNS])[:：]\s*(.*)/i);
      if (matchJ) {
        jvName = matchJ[2].trim();
      }

      let tagCode = baseCode;
      if (vCode) tagCode += "-" + vCode;

      let fullOName = originalName + (ovName && ovName !== "N" ? " - " + ovName : "");
      let fullEName = englishName + (evName && evName !== "N" ? " - " + evName : "");
      let fullJName = jpName + (jvName && jvName !== "N" ? " - " + jvName : "");

      for (let s of sizes) {
        let sCode = s;
        if (sCode) {
          const sl = sCode.toLowerCase();
          if (sl === "f" || sl === "free" || sl === "frees") sCode = "F";
        }

        let fullCodeParts = [baseCode];
        if (vCode) fullCodeParts.push(vCode); 
        if (sCode) fullCodeParts.push(sCode);
        if (suppCode) fullCodeParts.push(suppCode);
        
        let fullCode = fullCodeParts.join("-");
        
        if (ACTUAL_IS_APPEND_MODE_SKU && existingCodesInSheet.has(fullCode)) continue;
        
        const displayUrl = getDirectImageUrl(photoUrl);

        // ★ 全27列の配列を作成
        let outRow = new Array(ACTUAL_SKU_GEN_CONFIG.SKU.COL_CUMULATIVE_SALES).fill("");
        
        // ★ 各列への流し込み（I列の追加に伴い、J列以降が自動で右にズレます）
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_FULL_CODE - 1]     = fullCode;   
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_SUPPLIER - 1]      = suppCode;    // B列
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_BRAND - 1]         = brandName;  
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_SITE_URL - 1]      = siteUrl;    
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_PHOTO - 1]         = displayUrl ? '=IMAGE("' + displayUrl + '")' : ""; 
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_PHOTO_URL - 1]     = photoUrl;   
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_BASE_CODE - 1]     = baseCode;    // G列
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_TAG_CODE - 1]      = tagCode;     // H列   
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_TT_NAME - 1]       = ttName;      // I列 (17) ★NEW
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_ORIGINAL_NAME - 1] = originalName;// J列
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_ENGLISH_NAME - 1]  = englishName; // K列   
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_JAPANESE_NAME - 1] = jpName;      // L列   
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_ORIGINAL_FULL - 1] = fullOName;   // M列   
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_ENGLISH_FULL - 1]  = fullEName;   // N列   
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_JAPANESE_FULL - 1] = fullJName;   // O列   
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_VAR_ONLY - 1]      = evar;        // P列
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_VAR_JP_ONLY - 1]   = jvar;        // Q列   
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_SIZE - 1]          = sCode;       // R列
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_COUNTRY - 1]       = country;     // S列
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_LOCAL_PRICE - 1]   = price;       // T列   
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_JPY_PRICE - 1]     = "";          // U列
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_TT_PRICE - 1]      = ttPrice;     // V列 (TikTok価格)
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_MOBILE_PRICE - 1]  = mobilePrice; // W列 (移動販売価格)
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_RENTAL_PRICE - 1]  = rentalPrice; // X列 (レンタル価格)
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_DIFFERENCE - 1]        = ""; 
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_TOTAL_PIECES - 1]      = "";
        outRow[ACTUAL_SKU_GEN_CONFIG.SKU.COL_CUMULATIVE_SALES - 1]  = "";
        
        if (existingStockMap.has(fullCode)) {
          outRow = outRow.concat(existingStockMap.get(fullCode));
          
          // 【Akiraさんロジック】関数が入る列の「過去の数値」を上書きで消す
          if (outRow.length > 33) outRow[33] = ""; // AH列（52_倉庫現在庫）
          if (outRow.length > 37) outRow[37] = ""; // AL列（56_[TT Shop] 現在庫数）
          if (outRow.length > 41) outRow[41] = ""; // AP列（60_[移動販売] 在庫数）
          if (outRow.length > 45) outRow[45] = ""; // AT列（73_[レンタル] 在庫数）
        }
             
        outputData.push(outRow);
      }
    }
  }
  
  if (outputData.length > 0) {
    if (!ACTUAL_IS_APPEND_MODE_SKU) {
      
      // =========================================================
      // ★ 21列仕様に合わせたソートのインデックス調整！
      // =========================================================
      outputData.sort((a, b) => {
        
        // 1. 国 (S列 / Index: 18) VN > CN 降順
        const countryA = String(a[18] || "").toUpperCase();
        const countryB = String(b[18] || "").toUpperCase();
        if (countryA > countryB) return -1;
        if (countryA < countryB) return 1;

        // 2. 型番 (G列 / Index: 6) 昇順
        const codeA = String(a[6] || "").toUpperCase();
        const codeB = String(b[6] || "").toUpperCase();
        if (codeA < codeB) return -1;
        if (codeA > codeB) return 1;

        // 3. サプライヤー (B列 / Index: 1) 「BC」を最優先、その他は昇順
        const suppA = String(a[1] || "").toUpperCase();
        const suppB = String(b[1] || "").toUpperCase();
        if (suppA === "BC" && suppB !== "BC") return -1;
        if (suppA !== "BC" && suppB === "BC") return 1;
        if (suppA < suppB) return -1;
        if (suppA > suppB) return 1;

        // 4. バリエーション英字 (P列 / Index: 15) 昇順
        const varA = String(a[15] || "").toUpperCase();
        const varB = String(b[15] || "").toUpperCase();
        if (varA < varB) return -1;
        if (varA > varB) return 1;

        // 5. サイズ (R列 / Index: 17) カスタム順
        const sizeA = String(a[17] || "").toUpperCase();
        const sizeB = String(b[17] || "").toUpperCase();
        const sizeOrder = { "XS":1, "S":2, "M":3, "L":4, "XL":5, "F":6, "FREE":6 };
        const orderA = sizeOrder[sizeA] || 99;
        const orderB = sizeOrder[sizeB] || 99;
        return orderA - orderB;
      });
    }

    const maxCol = Math.max(...outputData.map(r => r.length), skuLastCol);
    let outputColors = []; let colorFlag = true; let previousBase = "";

    for (let row of outputData) {
      const currentTagCode = String(row[ACTUAL_SKU_GEN_CONFIG.SKU.COL_TAG_CODE - 1] || "");
      let currentBase = currentTagCode.split("-")[0];
      const match = currentTagCode.match(/^([A-Z]+\d+)/i);
      if (match) currentBase = match[1];

      if (currentBase !== previousBase) { colorFlag = !colorFlag; previousBase = currentBase; }
      const rowColor = colorFlag ? "#ffffff" : "#f3f3f3";
      outputColors.push(new Array(maxCol).fill(rowColor));
    }

    let writeStartRow = ACTUAL_SKU_GEN_CONFIG.SKU.START_ROW;
    if (ACTUAL_IS_APPEND_MODE_SKU) {
      writeStartRow = skuLastRow < ACTUAL_SKU_GEN_CONFIG.SKU.START_ROW ? ACTUAL_SKU_GEN_CONFIG.SKU.START_ROW : skuLastRow + 1;
    } else {
      if (skuLastRow >= ACTUAL_SKU_GEN_CONFIG.SKU.START_ROW) {
        skuSheet.getRange(ACTUAL_SKU_GEN_CONFIG.SKU.START_ROW, 1, skuLastRow - ACTUAL_SKU_GEN_CONFIG.SKU.START_ROW + 1, skuLastCol).clearContent().setBackground("#ffffff");
      }
    }
    // ★ 行が足りない場合のみ、一番下に追加する
    const currentMaxRows = skuSheet.getMaxRows();
    const neededMaxRows = writeStartRow + outputData.length - 1;
    if (neededMaxRows > currentMaxRows) {
      skuSheet.insertRowsAfter(currentMaxRows, neededMaxRows - currentMaxRows);
    }
    const target = skuSheet.getRange(writeStartRow, 1, outputData.length, maxCol);
    const finalizedData = outputData.map(r => { while (r.length < maxCol) r.push(""); return r; });
    target.setValues(finalizedData).setBackgrounds(outputColors);

    const finalRow = writeStartRow + outputData.length - 1;
    const colCountry = getColumnLetter(ACTUAL_SKU_GEN_CONFIG.SKU.COL_COUNTRY); // S列
    const colPrice = getColumnLetter(ACTUAL_SKU_GEN_CONFIG.SKU.COL_LOCAL_PRICE); // T列
    const colJpy = getColumnLetter(ACTUAL_SKU_GEN_CONFIG.SKU.COL_JPY_PRICE); // U列
    
    
    // ▼▼▼ ここから新しく貼り付ける ▼▼▼
    try {
      // U列 (21列目): 15_卸価格(￥)
      skuSheet.getRange(6, 21).setFormula('={"15_卸価格(￥)"; BYROW(S7:T, LAMBDA(row, IF(INDEX(row, 1, 2)="", "", IF(INDEX(row, 1, 1)="VN", INDEX(row, 1, 2) * $U$2, IF(INDEX(row, 1, 1)="CN", INDEX(row, 1, 2) * $U$3, "")))))}');

      // AH列 (34列目): 52_倉庫現在庫
      skuSheet.getRange(6, 34).setFormula('={"52_倉庫現在庫"; ARRAYFORMULA(IF(A7:A="", "", AE7:AE - AF7:AF - AG7:AG))}');

      // AL列 (38列目): 56_[TT Shop] 現在庫数
      skuSheet.getRange(6, 38).setFormula('={"56_[TT Shop] 現在庫数"; ARRAYFORMULA(IF(A7:A="", "", AI7:AI - AJ7:AJ - AK7:AK))}');

      // AP列 (42列目): 60_[移動販売] 在庫数
      skuSheet.getRange(6, 42).setFormula('={"60_[移動販売] 在庫数"; ARRAYFORMULA(IF(A7:A="", "", AM7:AM - AN7:AN - AO7:AO))}');

      // AT列 (46列目): 73_[レンタル] 在庫数
      skuSheet.getRange(6, 46).setFormula('={"73_[レンタル] 在庫数"; ARRAYFORMULA(IF(A7:A="", "", AQ7:AQ - AR7:AR - AS7:AS))}');

    } catch (e) {
      Browser.msgBox("⚠️数式セット時にエラーが起きました。\n詳細: " + e.message);
    }
    // ▲▲▲ ここまで ▲▲▲
        
    Browser.msgBox(`完了！\nI列にTikTok商品名を追加した全21列のデータを展開しました！`);
  } else {
    Browser.msgBox("⚠️ 展開するデータが0件でした。");
  }
}