/**
 * [MASTER]シートのIDに基づき、英語から日本語への翻訳を行うスクリプト
 * 【STEP 3】日本語翻訳専用（色名辞書によるカタカナ固定） ＋ 最後に全体ソートを実行
 */
function translateToJapaneseStep3() {
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
  // ★ アキラさん発案：全角/半角の揺れを吸収するフィルター関数
  // =========================================================
  const normalizeId = (h) => {
    return String(h).trim()
      .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)) // 全角数字を半角に
      .replace(/＿/g, "_"); // 全角アンダーバーを半角に
  };

  // =========================================================
  // ① 見出し行からIDを検索して列番号を特定する
  // =========================================================
  const lastCol = sheet.getLastColumn();
  const headerValues = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const colMap = {};
  
  for (let i = 0; i < headerValues.length; i++) {
    const cellValue = normalizeId(headerValues[i]);
    const match = cellValue.match(/^(\d{2,3})_/); // 2桁または3桁のIDを正確に取得
    if (match) {
      const id = match[1]; 
      colMap[id] = i + 1;
    }
  }

  // 必須列のチェック (08:英名, 11:英バリエ, 16:日名, 12:日バリエ)
  // ※ソートに使う13と06もチェック対象に入れておくとより安全です
  const requiredIds = ["08", "11", "16", "12"];
  for (let id of requiredIds) {
    if (!colMap[id]) {
      Browser.msgBox("エラー：見出しに「" + id + "_」から始まるIDが見つかりません。中断します。");
      return;
    }
  }

  // =========================================================
  // ② 英語から日本語へのカタカナ辞書（アキラさん指定・完全一致用）
  // =========================================================
  const colorDictJp = {
    "Red Wine": "レッドワイン",  
    "Black": "ブラック",
    "White": "ホワイト",
    "Mint Green": "ミントグリーン",
    "Ash Pink": "アッシュピンク",
    "Yellow Pale": "ペールイエロー",
    "Yellow Pearl": "イエローパール",
    "Pearl Yellow": "イエローパール",
    "Pearl": "パール",
    "Yellow": "イエロー",
    "Gray": "グレー",
    "Grey": "グレー",
    "Turquoise": "ターコイズ",
    "Champagne": "シャンパン",
    "Champagne Beige": "シャンパンベージュ",
    "Apricot": "アプリコット",
    "Gold": "ゴールド",
    "Brown": "ブラウン",
    "Blue": "ブルー",
    "Purple": "パープル",
    "Cream": "クリーム",
    "Oatmeal": "オートミール",
    "Stripe": "ストライプ"
  };

  // データの読み込み
  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol);
  const data = dataRange.getValues();
  let updated = false;

  // =========================================================
  // ③ 各行の処理
  // =========================================================
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const currentRowNum = startRow + i;

    // --- 1. 商品名の翻訳 (08_英名 -> 16_日名) ---
    const engName = String(row[colMap["08"] - 1] || "").trim();
    let jpName = String(row[colMap["16"] - 1] || "").trim();

    if (engName !== "" && jpName === "") {
      try {
        jpName = LanguageApp.translate(engName, 'en', 'ja');
        sheet.getRange(currentRowNum, colMap["16"]).setValue(jpName);
        updated = true;
      } catch (e) {}
    }

    // --- 2. バリエーションの翻訳 (11_英バリエ -> 12_日バリエ) ---
    const engVarStr = String(row[colMap["11"] - 1] || "").trim();
    let jpVar = String(row[colMap["12"] - 1] || "").trim();

    if (engVarStr !== "" && jpVar === "") {
      let vars = engVarStr.split(/[,、\n]+/).map(v => v.trim()).filter(v => v !== "");
      let translatedVars = [];

      for (let v of vars) {
        let prefix = "";
        let colorPart = v;
        
        let match = v.match(/^([A-ZNS])[:：]\s*(.*)/i);
        if (match) {
          prefix = match[1].toUpperCase() + ":";
          colorPart = match[2].trim();
        }

        // 辞書引き
        let lookupKey = toTitleCase(colorPart);
        let translatedColor = colorDictJp[lookupKey];

        if (!translatedColor) {
          try {
            translatedColor = LanguageApp.translate(colorPart, 'en', 'ja');
          } catch (e) {
            translatedColor = colorPart; // エラー時は英語をそのまま残す
          }
        }

        translatedVars.push(prefix + translatedColor);
      }

      const finalJpVar = translatedVars.join(", ");
      
      // 変更があれば書き込み
      if (jpVar !== finalJpVar) {
        sheet.getRange(currentRowNum, colMap["12"]).setValue(finalJpVar);
        updated = true;
      }
    }
  }

  // =========================================================
  // ④ 処理完了後のソート（並べ替え） ★ここに復活！
  // =========================================================
  if (colMap["13"] && colMap["06"]) {
    const fullDataRange = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, lastCol);
    
    // ★ 以前のソート基準を一旦入れています
    fullDataRange.sort([
      { column: colMap["13"], ascending: false }, // 優先順位1: 13_(国)の降順 (VN > CN)
      { column: colMap["06"], ascending: true }   // 優先順位2: 06_(型番)の昇順
    ]);
  }

  if (updated) {
    Browser.msgBox("完了！\\n英語から日本語への翻訳が完了し、最後にシートをソートしました。");
  } else {
    Browser.msgBox("翻訳が必要な項目はありませんでしたが、シートのソートは実行しました。");
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