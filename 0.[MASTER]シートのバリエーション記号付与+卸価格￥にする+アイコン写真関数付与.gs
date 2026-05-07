/**
 * [MASTER]シートのIDに基づき、以下の下ごしらえを行う専用スクリプト
 * 1. サプライヤー(01_)が「B」「BC」「BD」の時のみ06番を自動採番する！
 * 2. バリエーションへの記号(A, B..)付与 (10_)
 * 3. 【NEW】写真表示(05->04)：スマートチップ化されたURLにも完全対応！
 * 4. 日本円計算(15_)の関数セット
 */
function generateCodesAndSymbols() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MASTER");
  if (!sheet) {
    Browser.msgBox("MASTERシートが見つかりません。");
    return;
  }

  const headerRow = 6;  
  const startRow = 7;   
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const lastCol = sheet.getLastColumn();
  const headerValues = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const colMap = {};
  
  for (let i = 0; i < headerValues.length; i++) {
    const cellValue = String(headerValues[i]).trim();
    const match = cellValue.match(/^(\d{2,3})_/); 
    if (match) colMap[match[1]] = i + 1;
  }

  const requiredIds = ["01", "02", "04", "05", "06", "07", "10"];
  for (let id of requiredIds) {
    if (!colMap[id]) {
      Browser.msgBox(`エラー：見出しに「${id}_」が見つかりません。`);
      return;
    }
  }

  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol);
  const data = dataRange.getValues();
  
  // ★追加：スマートチップのURLや数式を高速で読み取るための準備
  const richData = dataRange.getRichTextValues(); 
  const formulas = dataRange.getFormulas();       
  let updated = false;

  // =========================================================
  // ★文字抜き出しフィルター
  // =========================================================
  const getBrandPrefix = (str) => {
    if (!str) return null;
    let s = String(str).replace(/[Ａ-Ｚａ-ｚ]/g, m => String.fromCharCode(m.charCodeAt(0) - 0xFEE0))
                       .toUpperCase()
                       .replace(/\s+/g, ""); 
    let match = s.match(/[A-Z]{2}/); 
    return match ? match[0] : null; 
  };

  const getSuppPrefix = (str) => {
    if (!str) return "";
    let s = String(str).toUpperCase().replace(/\s+/g, "");
    let match = s.match(/[A-Z]+/); 
    return match ? match[0] : "";
  };

  // =========================================================
  // 事前準備：現在の最大連番の把握
  // =========================================================
  const brandCounters = new Map();

  for (let i = 0; i < data.length; i++) {
    const code = String(data[i][colMap["06"] - 1] || "").trim().replace(/\s+/g, "");
    if (code) {
      const match = code.match(/^([A-Z]{2})(\d{5})$/i);
      if (match) {
        const brand = match[1].toUpperCase();
        const num = parseInt(match[2], 10);
        if (!brandCounters.has(brand) || brandCounters.get(brand) < num) {
          brandCounters.set(brand, num);
        }
      }
    }
  }

  // =========================================================
  // 各行の処理
  // =========================================================
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const currentRowNum = startRow + i; 

    // --- 1. 自動採番 ---
    let baseModel = String(row[colMap["06"] - 1] || "").trim();
    const suppRaw  = String(row[colMap["01"] - 1] || "").trim();
    const brandRaw = String(row[colMap["02"] - 1] || "").trim();
    
    const suppCode = getSuppPrefix(suppRaw);
    const brandPrefix = getBrandPrefix(brandRaw);

    if (!baseModel && brandPrefix && suppCode && (suppCode === "B" || suppCode === "BC" || suppCode === "BD")) {
      let currentNum = brandCounters.has(brandPrefix) ? brandCounters.get(brandPrefix) : 10000;
      currentNum++; 
      brandCounters.set(brandPrefix, currentNum); 
      
      baseModel = brandPrefix + String(currentNum).padStart(5, '0'); 
      
      sheet.getRange(currentRowNum, colMap["06"]).setValue(baseModel);
      updated = true;
    }

    // --- 2. バリエーションの記号付与 (10_) ---
    const origVar = String(row[colMap["10"] - 1] || "").trim();
    let varCodes = []; 

    if (origVar !== "") {
      let vars = origVar.split(/[,、\n]+/).map(v => v.trim()).filter(v => v !== "");
      if (vars.length > 0) {
        let newOrigParts = [];
        const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        for (let j = 0; j < vars.length; j++) {
          let v = vars[j];
          let prefix, cleanO;
          let match = v.match(/^([A-ZNS])[:：]\s*(.*)/i);
          if (match) {
            prefix = match[1].toUpperCase();
            cleanO = match[2].trim();
          } else {
            prefix = (vars.length === 1) ? "N" : alphabet[j] || (j + 1);
            cleanO = v;
          }
          newOrigParts.push(prefix + ":" + cleanO);
          varCodes.push(prefix);
        }

        const newOrigStr = newOrigParts.join(", ");
        if (origVar !== newOrigStr) {
          sheet.getRange(currentRowNum, colMap["10"]).setValue(newOrigStr);
          updated = true;
        }
      }
    } else {
      varCodes = [""]; 
    }

    // --- 3. 写真表示 (05 -> 04) 【スマートチップ対応版】 ---
    if (colMap["05"] && colMap["04"]) {
      // セルの裏側からリンクURLを引っ張り出す（なければ普通の文字として取得）
      const richCell = richData[i][colMap["05"] - 1];
      let rawPhotoUrl = richCell.getLinkUrl() || richCell.getText();
      rawPhotoUrl = String(rawPhotoUrl || "").trim();
      
      // 既存の数式（=IMAGE...）がすでに入っているかチェック
      const currentFormula = String(formulas[i][colMap["04"] - 1] || "").trim();
      const currentVal = String(row[colMap["04"] - 1] || "").trim();
      
      if (rawPhotoUrl !== "" && currentFormula === "" && currentVal === "") {
        const photoUrl = rawPhotoUrl.split(/[,、\n\s]+/)[0];
        const directUrl = getDirectImageUrl(photoUrl);
        if (directUrl) {
          sheet.getRange(currentRowNum, colMap["04"]).setFormula('=IMAGE("' + directUrl + '")');
          updated = true;
        }
      }
    }
  }

  // --- 4. 日本円計算のBYROW数式セット (15) ---
  if (colMap["15"] && colMap["13"] && colMap["14"]) {
    const colCur = getColumnLetter(colMap["13"]);
    const colPrc = getColumnLetter(colMap["14"]);
    const formula = `=BYROW(${colCur}${startRow}:${colPrc}${lastRow}, LAMBDA(row, IF(INDEX(row, 1, 2)="", "", IF(INDEX(row, 1, 1)="VN", INDEX(row, 1, 2) * $L$2, IF(INDEX(row, 1, 1)="CN", INDEX(row, 1, 2) * $L$3, "")))))`;
    const currentFormula = sheet.getRange(startRow, colMap["15"]).getFormula();
    if (currentFormula !== formula) {
      sheet.getRange(startRow, colMap["15"]).setFormula(formula);
      updated = true;
    }
  }

  if (updated) {
    // 最後に強制セーブして、サイドバーがすぐに読み込めるようにする
    SpreadsheetApp.flush(); 
    Browser.msgBox("完了！\n自動採番、記号付与、スマートチップ画像の変換を行いました。");
  } else {
    Browser.msgBox("更新が必要な項目はありませんでした。");
  }
}

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