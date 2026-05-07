/**
 * VA マスター作成・更新プログラム【完全統合・最終進化版】
 * - 064_完全SKUコードを「合鍵」にして写真URLを絶対保護
 * - BC最優先＋バリエーション順の爆速ソート
 */

function createVariationMaster() {
  const SHEET_NAME_MASTER = "MASTER";
  const SHEET_NAME_VAR = "VaMASTER"; 
  const COPY_TO_IDS = ["20", "21", "22"];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(SHEET_NAME_MASTER);
    const varMasterSheet = ss.getSheetByName(SHEET_NAME_VAR);

    if (!masterSheet) throw new Error(`シート「${SHEET_NAME_MASTER}」が見つかりません。`);
    if (!varMasterSheet) throw new Error(`シート「${SHEET_NAME_VAR}」が見つかりません。`);

    const HEADER_ROW = 6;
    const DATA_START_ROW = 7;

    // --- 1. ヘッダーを読み込む ---
    function getColumnMapping(sheet) {
      const headers = sheet.getRange(HEADER_ROW, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
      const mapping = {};
      let hasAnyId = false;
      headers.forEach((header, index) => {
        if (!header) return;
        const cleanHeader = String(header).replace(/＿/g, "_").replace(/\s/g, "");
        const match = cleanHeader.match(/^(\d+_)/);
        if (match) {
          mapping[match[1].replace("_", "")] = index + 1;
          hasAnyId = true;
        }
      });
      return { mapping, hasAnyId };
    }

    const masterMapRes = getColumnMapping(masterSheet);
    const varMapRes = getColumnMapping(varMasterSheet);

    if (!masterMapRes.hasAnyId) throw new Error(`MASTERシートに項目IDが見つかりません。`);
    if (!varMapRes.hasAnyId) throw new Error(`VaMASTERシートに項目IDが見つかりません。`);

    const mCols = masterMapRes.mapping;
    const vCols = varMapRes.mapping;

    // --- 2. 【記憶】Akiraさん提案：064コードを合鍵にして写真URLを保存！ ---
    const existingPhotoUrls = new Map();
    const varLastRow = varMasterSheet.getLastRow();
    
    // 064_列と05_列が存在する場合のみ記憶を実行
    if (vCols["064"] && vCols["05"] && varLastRow >= DATA_START_ROW) {
      const currentVarData = varMasterSheet.getRange(DATA_START_ROW, 1, varLastRow - DATA_START_ROW + 1, varMasterSheet.getLastColumn()).getValues();
      currentVarData.forEach(row => {
        const vKey = String(row[vCols["064"]-1] || "").trim(); // 合鍵（064_）
        const url = row[vCols["05"]-1];                       // 守りたい写真URL
        if (vKey && url) existingPhotoUrls.set(vKey, url);
      });
    }

    // --- 3. MASTERデータの読み込みと【超高速地図（Map）】の作成 ---
    const masterLastRow = masterSheet.getLastRow();
    if (masterLastRow < DATA_START_ROW) throw new Error("MASTERシートにデータが1件もありません。");
    const masterData = masterSheet.getRange(DATA_START_ROW, 1, masterLastRow - DATA_START_ROW + 1, masterSheet.getLastColumn()).getValues();

    const masterLookup = new Map();
    masterData.forEach((mRow, mIdx) => {
      const mKey = String(mRow[mCols["06"]-1] || "") + "|" + String(mRow[mCols["01"]-1] || "");
      if (!masterLookup.has(mKey)) masterLookup.set(mKey, DATA_START_ROW + mIdx);
    });

    // --- 4. 展開 ＆ 【合体】 ---
    const newRows = [];
    const copyTasks = [];

    masterData.forEach((row, rowIndex) => {
      const masterRowPos = DATA_START_ROW + rowIndex;
      const modelNum = row[mCols["06"] - 1]; 
      if (!modelNum) return;

      const rawSupplier = String(row[mCols["01"] - 1] || "");
      const suppCode = rawSupplier.split(/[:：,，]/)[0].trim();
      const rawVariations = String(row[mCols["10"] - 1] || row[mCols["11"] - 1] || row[mCols["12"] - 1] || "").split(/[,、\n]/);

      rawVariations.forEach(v => {
        const vRaw = v.trim();
        if (!vRaw) return;

        const vMatch = vRaw.match(/^([A-ZNS0-9]+)[:：]\s*(.*)/i);
        const vCode = vMatch ? vMatch[1] : vRaw;
        
        // ★これが最強の合鍵（DC10001-A-BC のような形になる）
        const vKey = `${modelNum}-${vCode}-${suppCode}`;
        
        let newRow = new Array(varMasterSheet.getLastColumn() || 25).fill("");

        // 【基本項目セット】
        if (vCols["064"]) newRow[vCols["064"] - 1] = vKey;
        if (vCols["06"])  newRow[vCols["06"] - 1]  = modelNum;
        if (vCols["01"])  newRow[vCols["01"] - 1]  = rawSupplier;
        if (vCols["09"])  newRow[vCols["09"] - 1]  = row[mCols["09"] - 1];
        if (vCols["13"])  newRow[vCols["13"] - 1]  = row[mCols["13"] - 1];
        if (vCols["14"])  newRow[vCols["14"] - 1]  = row[mCols["14"] - 1];

        // 各言語のバリエーション名抽出
        ["10", "11", "12"].forEach(id => {
          if (vCols[id] && mCols[id]) {
            const mCells = String(row[mCols[id] - 1]).split(/[,、\n]/);
            const matched = mCells.find(c => c.trim().startsWith(vCode + ":") || c.trim().startsWith(vCode + "："));
            newRow[vCols[id] - 1] = matched ? matched.trim() : "";
          }
        });

        // 商品名結合
        const nameConfigs = [{t:"101", b:"07", v:"10"}, {t:"111", b:"08", v:"11"}, {t:"121", b:"16", v:"12"}];
        nameConfigs.forEach(conf => {
          if (vCols[conf.t] && mCols[conf.b]) {
            const base = row[mCols[conf.b] - 1] || "";
            const variCells = String(row[mCols[conf.v] - 1]).split(/[,、\n]/);
            const matched = variCells.find(c => c.trim().startsWith(vCode + ":") || c.trim().startsWith(vCode + "："));
            const namePart = matched ? matched.split(/[:：]/)[1].trim() : "";
            newRow[vCols[conf.t] - 1] = namePart ? `${base} - ${namePart}` : base;
          }
        });

        // 【写真URLの合体！】
        let finalPhotoUrl = row[mCols["05"] - 1] || ""; // デフォルトはMASTERの写真
        
        // メモ帳（Map）の中にこの合鍵（vKey）があれば、手入力されたURLに差し替える
        if (existingPhotoUrls.has(vKey)) {
          finalPhotoUrl = existingPhotoUrls.get(vKey);
        }

        if (vCols["05"]) newRow[vCols["05"] - 1] = finalPhotoUrl;
        if (vCols["04"] && finalPhotoUrl) {
          newRow[vCols["04"] - 1] = `=IMAGE("${getDirectImageUrl(finalPhotoUrl)}")`;
        }

        // コピペ対象外の項目を埋める
        Object.keys(vCols).forEach(id => {
          const skip = ["064","06","01","09","10","11","12","101","111","121","04","05","15", ...COPY_TO_IDS];
          if (!skip.includes(id) && mCols[id]) {
            newRow[vCols[id] - 1] = row[mCols[id] - 1];
          }
        });

        newRows.push(newRow);
      });
    });

    if (newRows.length === 0) throw new Error("展開するデータが0件でした。");

    // --- 5. 賢い並び替え（型番 ＞ BC優先 ＞ 色順） ---
    newRows.sort((a, b) => {
      // 国 (VN > CN)
      const cA = String(a[vCols["13"]-1] || "");
      const cB = String(b[vCols["13"]-1] || "");
      if (cA !== cB) return cB.localeCompare(cA);

      // 型番 (昇順)
      const codeA = String(a[vCols["06"]-1] || "");
      const codeB = String(b[vCols["06"]-1] || "");
      if (codeA !== codeB) return codeA.localeCompare(codeB);

      // サプライヤー（BCを絶対優先！）
      const getS = (val) => String(val || "").split(/[:：,，]/)[0].trim().toUpperCase();
      const sA = getS(a[vCols["01"]-1]);
      const sB = getS(b[vCols["01"]-1]);
      if (sA === "BC" && sB !== "BC") return -1;
      if (sA !== "BC" && sB === "BC") return 1;
      if (sA !== sB) return sA.localeCompare(sB);

      // バリエーション英字順
      const vA = String(a[vCols["11"]-1] || "");
      const vB = String(b[vCols["11"]-1] || "");
      return vA.localeCompare(vB);
    });

    // --- 6. 一気に書き出し（白紙に戻してから貼る） ---
    if (varLastRow >= DATA_START_ROW) {
      varMasterSheet.getRange(DATA_START_ROW, 1, varLastRow - DATA_START_ROW + 1, varMasterSheet.getLastColumn()).clearContent();
    }
    varMasterSheet.getRange(DATA_START_ROW, 1, newRows.length, newRows[0].length).setValues(newRows);

    // --- 7. スマートチップを【バケツで一括】復元（5分を数秒にします） ---
    COPY_TO_IDS.forEach(id => {
      if (vCols[id] && mCols[id]) {
        // 1. マスター側のチップと色を「一列まるごと」バケツに読み込む
        const mRange = masterSheet.getRange(DATA_START_ROW, mCols[id], masterLastRow - DATA_START_ROW + 1, 1);
        const mRichTexts = mRange.getRichTextValues();
        const mBackgrounds = mRange.getBackgrounds();

        // 2. 書き出し用の新しいバケツを用意
        const newRichTexts = [];
        const newBackgrounds = [];

        // 3. 並べ替えた行に合わせて、パソコンの頭の中で中身を移し替える（一瞬）
        newRows.forEach(row => {
          const mKey = String(row[vCols["06"]-1] || "") + "|" + String(row[vCols["01"]-1] || "");
          const srcR = masterLookup.get(mKey);
          
          if (srcR) {
            const rowIdx = srcR - DATA_START_ROW;
            newRichTexts.push([mRichTexts[rowIdx][0]]);
            newBackgrounds.push([mBackgrounds[rowIdx][0]]);
          } else {
            // 見つからない場合は空っぽ
            newRichTexts.push([SpreadsheetApp.newRichTextValue().setText("").build()]);
            newBackgrounds.push(["#ffffff"]);
          }
        });

        // 4. バリエーションマスター側に「一列まるごとドカン！」と書き出す
        const vRange = varMasterSheet.getRange(DATA_START_ROW, vCols[id], newRows.length, 1);
        vRange.setRichTextValues(newRichTexts);
        vRange.setBackgrounds(newBackgrounds);
      }
    });

    // --- 8. 日本円のBYROW関数をポツンと置く ---
    if (vCols["15"] && vCols["14"] && vCols["13"]) {
      const finalR = DATA_START_ROW + newRows.length - 1;
      const formula = `=BYROW(${getColumnLetter(vCols["13"])}${DATA_START_ROW}:${getColumnLetter(vCols["14"])}${finalR}, LAMBDA(row, IF(INDEX(row,1,2)="", "", ROUND(INDEX(row,1,2) * IF(TRIM(INDEX(row,1,1))="VN", $L$2, IF(TRIM(INDEX(row,1,1))="CN", $L$3, 1))))))`;
      varMasterSheet.getRange(DATA_START_ROW, vCols["15"]).setFormula(formula);
    }

    SpreadsheetApp.getUi().alert(`完了: ${newRows.length}件\n【064】このシートでの追記保護しつつ、最新商品導入と整列が成功しました！`);

  } catch (e) {
    SpreadsheetApp.getUi().alert("エラー原因: " + e.message);
  }
}

// （補助機能）写真URL変換
function getDirectImageUrl(url) {
  if (!url) return "";
  const match = url.match(/(?:id=|d\/)([\w-]+)/);
  if (match) return `https://drive.google.com/uc?export=download&id=${match[1]}`;
  return url;
}

// （補助機能）列番号をA, B, C...に変換
function getColumnLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}