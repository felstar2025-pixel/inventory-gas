/**
 * [MASTER]シートのIDに基づき、以下の処理を行うSTEP2専用スクリプト
 * 1. 現地語の商品名(07_)を英語(08_)に翻訳
 * 2. 記号付きのバリエーション(10_)の記号を維持したまま英語(11_)に翻訳
 * 3. 英語商品名と英語バリエーション名を結合し、英語フルネーム(111_)を生成
 */
function translateToEnglishStep2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MASTER");
  if (!sheet) {
    Browser.msgBox("MASTERシートが見つかりません。");
    return;
  }

  const headerRow = 6;  // 見出し行
  const startRow = 7;   // データ開始行
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  // =========================================================
  // ① 見出し行からIDを検索して列番号を特定する
  // =========================================================
  const lastCol = sheet.getLastColumn();
  const headerValues = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const colMap = {};
  
  for (let i = 0; i < headerValues.length; i++) {
    const cellValue = String(headerValues[i]).trim();
    const match = cellValue.match(/^(\d{2,3})_/); // 2桁または3桁のIDを取得
    if (match) {
      const id = match[1];
      colMap[id] = i + 1;
    }
  }

  // 今回必須となる列のチェック
  const requiredIds = ["07", "08", "10", "11"];
  for (let id of requiredIds) {
    if (!colMap[id]) {
      Browser.msgBox("エラー：見出しに「" + id + "_」から始まるIDが見つかりません。中断します。");
      return;
    }
  }

  // =========================================================
  // ② 辞書の設定（完全一致・部分一致）
  // =========================================================
  const apparelDict = {
    "Trắng": "White", "Da sáng": "Pearl", "Kem mỡ gà": "Pale Yellow",
    "Kem da sáng": "Yellow Pale", "Hồng ruốc": "Ash Pink", "Đen": "Black",
    "dryross": "dryross", "灰尘邓邓": "DustdipClub", "oshea": "oshea",
    "Đỏ đô": "Red Wine", "Xanh mint": "Mint Green", "Xanh ngọc": "Turquoise"
  };

  const nameDict = {
    "灰尘邓邓": "DustdipClub", "dryross": "dryross", "oshea": "oshea"
  };

  // データの読み込み
  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol);
  const data = dataRange.getValues();
  let updated = false;

  // =========================================================
  // ③ 各行の処理（1行ずつシートに直接書き込む方式）
  // =========================================================
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const currentRowNum = startRow + i; // 現在の行番号

    // --- 1. 商品名の翻訳 (07_元名 -> 08_英名) ---
    const origName = String(row[colMap["07"] - 1] || "").trim();
    let engName = String(row[colMap["08"] - 1] || "").trim();

    if (origName !== "" && engName === "") {
      let tempName = origName;
      // 辞書による置き換え
      for (let key in nameDict) {
        if (tempName.includes(key)) {
          tempName = tempName.split(key).join(nameDict[key]);
        }
      }
      try {
        // Google翻訳と頭文字の大文字化
        engName = toTitleCase(LanguageApp.translate(tempName, '', 'en'));
        sheet.getRange(currentRowNum, colMap["08"]).setValue(engName);
        updated = true;
      } catch (e) {}
    }

    // --- 2. バリエーションの翻訳 (10_元バリエ -> 11_英バリエ) ---
    const origVarStr = String(row[colMap["10"] - 1] || "").trim();
    let engVarStr = String(row[colMap["11"] - 1] || "").trim();

    if (origVarStr !== "") {
      let vars = origVarStr.split(/[,、\n]+/).map(v => v.trim()).filter(v => v !== "");
      
      if (vars.length > 0) {
        let newEngParts = [];

        for (let j = 0; j < vars.length; j++) {
          let v = vars[j];
          let prefix = "";
          let cleanOrig = v;
          
          // STEP1で付けた「A:」などの記号を分離
          let match = v.match(/^([A-ZNS])[:：]\s*(.*)/i);
          if (match) {
            prefix = match[1].toUpperCase() + ":";
            cleanOrig = match[2].trim();
          }

          // 英語翻訳
          let transEn = apparelDict[cleanOrig];
          if (!transEn) {
            try {
              transEn = LanguageApp.translate(cleanOrig, '', 'en');
            } catch(e) {
              transEn = cleanOrig; // 翻訳失敗時は元の文字をキープ
            }
          }
          
          transEn = toTitleCase(transEn);
          newEngParts.push(prefix + transEn); // 例: A:White
        }

        const newEngStr = newEngParts.join(", ");

        // 変更があれば書き込み
        if (engVarStr === "" || engVarStr !== newEngStr) {
          sheet.getRange(currentRowNum, colMap["11"]).setValue(newEngStr);
          updated = true;
        }
      }
    }

  }

  if (updated) {
    Browser.msgBox("STEP2 完了！\\n英語への翻訳が終わりました。");
  } else {
    Browser.msgBox("更新が必要な項目はありませんでした。");
  }
}

/**
 * 頭文字を大文字にする
 */
function toTitleCase(str) {
  if (!str) return "";
  return String(str).replace(/\w\S*/g, function(txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}