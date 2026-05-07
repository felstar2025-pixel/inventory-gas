/**
 * バリエーションマスター作成・更新プログラム
 * [MASTER]シートの分解ロジック（記号抽出・画像変換）を完全踏襲
 * ※エラー通知機能付き（エラーをすっ飛ばさない）
 */

function createVariationMaster() {
  // ★ ここでシート名を管理します（後で変更してもすぐに直せます）
  const SHEET_NAME_MASTER = "MASTER";
  const SHEET_NAME_VAR = "VaMASTER"; 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(SHEET_NAME_MASTER);
    const varMasterSheet = ss.getSheetByName(SHEET_NAME_VAR);
    
    // エラー検知①：シートが存在しない場合は明確にエラーを出す
    if (!masterSheet) {
      throw new Error(`シート「${SHEET_NAME_MASTER}」が見つかりません。名前が正しいか確認してください。`);
    }
    if (!varMasterSheet) {
      throw new Error(`シート「${SHEET_NAME_VAR}」が見つかりません。名前が正しいか確認してください。`);
    }

    const HEADER_ROW = 6;
    const DATA_START_ROW = 7;

    // --- 1. ヘッダーマッピング（ID_で列を特定） ---
    function getColumnMapping(sheet) {
      const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
      const mapping = {};
      let hasAnyId = false;
      headers.forEach((header, index) => {
        if (!header) return;
        const cleanHeader = String(header).replace(/＿/g, "_").replace(/\s/g, "");
        const match = cleanHeader.match(/^(\d+_)/);
        if (match) {
          mapping[match[1]] = index + 1;
          hasAnyId = true;
        }
      });
      return { mapping, hasAnyId };
    }

    const masterMapResult = getColumnMapping(masterSheet);
    const varMapResult = getColumnMapping(varMasterSheet);

    const masterMap = masterMapResult.mapping;
    const varMap = varMapResult.mapping;

    // エラー検知②：見出し(6行目)にIDが一つもない場合
    if (!masterMapResult.hasAnyId) {
      throw new Error(`「${SHEET_NAME_MASTER}」の6行目に項目ID（例: 01_ など）が見つかりません。`);
    }
    if (!varMapResult.hasAnyId) {
      throw new Error(`「${SHEET_NAME_VAR}」の6行目に項目ID（例: 01_ など）が見つかりません。`);
    }

    // --- 2. マスターデータ取得 ---
    const masterLastRow = masterSheet.getLastRow();
    
    // エラー検知③：マスターシートにデータがない場合
    if (masterLastRow < DATA_START_ROW) {
      throw new Error(`「${SHEET_NAME_MASTER}」の ${DATA_START_ROW} 行目以降にデータが登録されていません。`);
    }

    const masterData = masterSheet.getRange(DATA_START_ROW, 1, masterLastRow - DATA_START_ROW + 1, masterSheet.getLastColumn()).getValues();
    
    // --- 3. 既存データの記憶（追記・上書き判定用） ---
    const varLastRow = varMasterSheet.getLastRow();
    const existingRowsMap = {};
    const maxCols = varMasterSheet.getLastColumn() || 20;
    
    if (varLastRow >= DATA_START_ROW) {
      const varData = varMasterSheet.getRange(DATA_START_ROW, 1, varLastRow - DATA_START_ROW + 1, maxCols).getValues();
      const keyColIdx = (varMap["064_"] || 1) - 1;
      varData.forEach(row => { if (row[keyColIdx]) existingRowsMap[row[keyColIdx]] = row; });
    }

    const newRows = [];

    // --- 4. データ変換ロジック ---
    masterData.forEach((row, rowIndex) => {
      const realRowNumber = DATA_START_ROW + rowIndex; // エラー報告用に行番号を把握
      
      const modelNum = row[masterMap["06_"] - 1]; // 06_商品コード
      if (!modelNum) return; // 空行はスキップ

      // サプライヤー記号の抽出
      const rawSupplier = String(row[masterMap["01_"] - 1] || "");
      const suppCode = rawSupplier.split(/[:：,，]/)[0].trim();

      // バリエーションの分解
      const rawVariations = String(row[masterMap["10_"] - 1] || row[masterMap["11_"] - 1] || row[masterMap["12_"] - 1] || "").split(/[,、\n]/);

      rawVariations.forEach(v => {
        const vRaw = v.trim();
        if (!vRaw) return;

        const vMatch = vRaw.match(/^([A-ZNS0-9]+)[:：]\s*(.*)/i);
        const vCode = vMatch ? vMatch[1] : vRaw;
        
        const vKey = `${modelNum}-${vCode}-${suppCode}`;
        let newRow = existingRowsMap[vKey] ? [...existingRowsMap[vKey]] : new Array(maxCols).fill("");

        // A列付近: 064_商品コード
        if (varMap["064_"]) newRow[varMap["064_"] - 1] = vKey;
        if (varMap["06_"]) newRow[varMap["06_"] - 1] = modelNum;
        if (varMap["01_"]) newRow[varMap["01_"] - 1] = rawSupplier;
        if (varMap["09_"]) newRow[varMap["09_"] - 1] = row[masterMap["09_"] - 1];

        const langIDs = ["10_", "11_", "12_"];
        langIDs.forEach(id => {
          if (varMap[id] && masterMap[id]) {
            const mCells = String(row[masterMap[id] - 1]).split(/[,、\n]/);
            const matched = mCells.find(c => c.trim().startsWith(vCode + ":") || c.trim().startsWith(vCode + "："));
            newRow[varMap[id] - 1] = matched ? matched.trim() : "";
          }
        });

        const combineRules = [
          { target: "101_", base: "07_", vari: "10_" },
          { target: "111_", base: "08_", vari: "11_" },
          { target: "121_", base: "16_", vari: "12_" }
        ];
        combineRules.forEach(rule => {
          if (varMap[rule.target] && masterMap[rule.base]) {
            const baseName = row[masterMap[rule.base] - 1] || "";
            const mCells = String(row[masterMap[rule.vari] - 1]).split(/[,、\n]/);
            const matchedVari = mCells.find(c => c.trim().startsWith(vCode + ":") || c.trim().startsWith(vCode + "："));
            const namePart = matchedVari ? matchedVari.split(/[:：]/)[1].trim() : "";
            newRow[varMap[rule.target] - 1] = namePart ? `${baseName} - ${namePart}` : baseName;
          }
        });

        // 写真URLの変換
        if (varMap["04_"] && masterMap["05_"]) {
          const photoUrl = row[masterMap["05_"] - 1];
          if (photoUrl) {
            const directUrl = getDirectImageUrl(photoUrl);
            newRow[varMap["04_"] - 1] = `=IMAGE("${directUrl}")`;
          }
        }
        if (varMap["05_"] && masterMap["05_"]) newRow[varMap["05_"] - 1] = row[masterMap["05_"] - 1];

        // その他の項目
        Object.keys(varMap).forEach(id => {
          const skipIds = ["064_","06_","01_","09_","10_","11_","12_","101_","111_","121_","04_","05_"];
          if (!skipIds.includes(id) && masterMap[id]) {
            newRow[varMap[id] - 1] = row[masterMap[id] - 1];
          }
        });

        newRows.push(newRow);
      });
    });

    // --- 5. 書き出しとソート ---
    // エラー検知④：書き出すデータがない場合
    if (newRows.length === 0) {
      throw new Error("バリエーションとして展開できるデータが1件も見つかりませんでした。\nMASTERシートのバリエーション欄（10_, 11_, 12_）に正しく入力されているか確認してください。");
    }

    varMasterSheet.getRange(DATA_START_ROW, 1, newRows.length, maxCols).setValues(newRows);
    
    // 並び替え
    if (varMap["06_"] && varMap["12_"] && varMap["01_"]) {
      const sortSpecs = [
        {column: varMap["06_"], ascending: true},
        {column: varMap["12_"], ascending: true},
        {column: varMap["01_"], ascending: true}
      ];
      varMasterSheet.getRange(DATA_START_ROW, 1, newRows.length, maxCols).sort(sortSpecs);
    }

    SpreadsheetApp.getUi().alert(`${SHEET_NAME_VAR} の更新が完了しました！\n処理件数: ${newRows.length} 件`);

  } catch (e) {
    // ★ ここで全てのエラーをキャッチし、絶対に画面に表示させる
    SpreadsheetApp.getUi().alert("処理がエラーで中断されました。\n\n【原因】\n" + e.message);
  }
}

/**
 * GoogleドライブURLをIMAGE関数用URLに変換（既存ロジック）
 */
function getDirectImageUrl(url) {
  if (!url) return "";
  const idMatch = url.match(/[-\w]{25,}/);
  if (idMatch) {
    return "https://lh3.googleusercontent.com/d/" + idMatch[0];
  }
  return url;
}