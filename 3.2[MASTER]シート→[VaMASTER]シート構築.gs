/**
 * VA マスター作成・更新プログラム【完全統合・最終版】
 * * 1. エラーを絶対に隠さないガード処理（復活）
 * 2. 05_写真URL をそのまま残す処理（復活）
 * 3. 各バリエーション(10,11,12)の抽出処理（復活）
 * 4. スマートチップ等を保持する最強のコピペ(copyTo)
 * 5. VN/CN で正確に判定する日本円(15_)のBYROW関数
 */

function createVariationMaster() {
  const SHEET_NAME_MASTER = "MASTER";
  const SHEET_NAME_VAR = "VaMASTER"; 
  
  // コピペ（copyTo）でデザインごと丸写ししたい項目ID
  const COPY_TO_IDS = ["20", "21", "22"]; 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(SHEET_NAME_MASTER);
    const varMasterSheet = ss.getSheetByName(SHEET_NAME_VAR);
    
    // 【ガード1】シートが存在するか確認
    if (!masterSheet) throw new Error(`シート「${SHEET_NAME_MASTER}」が見つかりません。`);
    if (!varMasterSheet) throw new Error(`シート「${SHEET_NAME_VAR}」が見つかりません。`);

    const HEADER_ROW = 6;
    const DATA_START_ROW = 7;

    // --- 1. ヘッダーマッピング（列の特定） ---
    function getColumnMapping(sheet) {
      const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
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

    // 【ガード2】項目ID（01_等）が存在するか確認
    if (!masterMapRes.hasAnyId) throw new Error(`「${SHEET_NAME_MASTER}」の6行目に項目IDが見つかりません。`);
    if (!varMapRes.hasAnyId) throw new Error(`「${SHEET_NAME_VAR}」の6行目に項目IDが見つかりません。`);

    const mCols = masterMapRes.mapping;
    const vCols = varMapRes.mapping;

    // --- 2. マスターデータ取得 ---
    const masterLastRow = masterSheet.getLastRow();
    
    // 【ガード3】MASTERシートにデータが存在するか確認
    if (masterLastRow < DATA_START_ROW) throw new Error("MASTERシートにデータが1件もありません。");
    
    const masterData = masterSheet.getRange(DATA_START_ROW, 1, masterLastRow - DATA_START_ROW + 1, masterSheet.getLastColumn()).getValues();

    const newRows = [];
    const copyTasks = []; // copyToの予約リスト

    // --- 3. データ変換ロジック ---
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
        const vKey = `${modelNum}-${vCode}-${suppCode}`;
        
        let newRow = new Array(varMasterSheet.getLastColumn() || 25).fill("");
        const currentRowIndex = newRows.length;

        // 【基本項目】
        if (vCols["064"]) newRow[vCols["064"] - 1] = vKey;
        if (vCols["06"])  newRow[vCols["06"] - 1]  = modelNum;
        if (vCols["01"])  newRow[vCols["01"] - 1]  = rawSupplier;
        if (vCols["09"])  newRow[vCols["09"] - 1]  = row[mCols["09"] - 1];
        if (vCols["13"])  newRow[vCols["13"] - 1]  = row[mCols["13"] - 1];
        if (vCols["14"])  newRow[vCols["14"] - 1]  = row[mCols["14"] - 1];

        // 【復活】バリエーション各言語（10, 11, 12）
        ["10", "11", "12"].forEach(id => {
          if (vCols[id] && mCols[id]) {
            const mCells = String(row[mCols[id] - 1]).split(/[,、\n]/);
            const matched = mCells.find(c => c.trim().startsWith(vCode + ":") || c.trim().startsWith(vCode + "："));
            newRow[vCols[id] - 1] = matched ? matched.trim() : "";
          }
        });

        // 【商品名】半角スペース入りで結合
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

        // 【写真表示とURLの保持】
        if (vCols["04"] && mCols["05"]) {
          const url = row[mCols["05"] - 1];
          if (url) newRow[vCols["04"] - 1] = `=IMAGE("${getDirectImageUrl(url)}")`;
        }
        // ★ これで 05_写真URL をそのままテキストとして残します
        if (vCols["05"] && mCols["05"]) {
          newRow[vCols["05"] - 1] = row[mCols["05"] - 1];
        }

        // 【copyTo対象の予約】
        COPY_TO_IDS.forEach(id => {
          if (vCols[id] && mCols[id]) {
            copyTasks.push({ srcR: masterRowPos, srcC: mCols[id], dstR: DATA_START_ROW + currentRowIndex, dstC: vCols[id] });
          }
        });

        // 【その他の項目（コピペ対象外）のコピー】
        Object.keys(vCols).forEach(id => {
          const skip = ["064","06","01","09","10","11","12","101","111","121","04","05","15", ...COPY_TO_IDS];
          if (!skip.includes(id) && mCols[id]) {
            newRow[vCols[id] - 1] = row[mCols[id] - 1];
          }
        });

        newRows.push(newRow);
      });
    });

    // --- 4. 書き出し ---
    // 【ガード4】書き出すデータがない場合
    if (newRows.length === 0) throw new Error("書き出せるバリエーションデータが0件でした。");
    
    varMasterSheet.getRange(DATA_START_ROW, 1, newRows.length, newRows[0].length).setValues(newRows);
    
    // 予約したスマートチップ等のコピペ(copyTo)を一気に実行
    copyTasks.forEach(t => {
      masterSheet.getRange(t.srcR, t.srcC).copyTo(varMasterSheet.getRange(t.dstR, t.dstC));
    });

    // --- 5. 日本円(15_)のBYROW関数入力 ---
    if (vCols["15"] && vCols["14"] && vCols["13"]) {
      const finalRow = DATA_START_ROW + newRows.length - 1;
      const c13 = getColumnLetter(vCols["13"]); 
      const c14 = getColumnLetter(vCols["14"]); 
      
      // ★ 判定文字を VN と CN に修正しました
      const formula = `=BYROW(${c13}${DATA_START_ROW}:${c14}${finalRow}, LAMBDA(row, ` + 
                      `IF(INDEX(row,1,2)="", "", ` +
                      `ROUND(INDEX(row,1,2) * IF(TRIM(INDEX(row,1,1))="VN", $L$2, IF(TRIM(INDEX(row,1,1))="CN", $L$3, 1))))))`;
      
      varMasterSheet.getRange(DATA_START_ROW, vCols["15"]).setFormula(formula);
    }

    // --- 【SKUシートの並び順を完全再現】 ---
    if (vCols["13"] && vCols["06"] && vCols["01"] && vCols["11"]) {
      const sortSpecs = [
        {column: vCols["13"], ascending: false}, // 1. 国 (VNを上にするため降順)
        {column: vCols["06"], ascending: true},  // 2. ベース型番 (昇順)
        {column: vCols["01"], ascending: true},  // 3. サプライヤー (BCなどを上にするため昇順)
        {column: vCols["11"], ascending: true}   // 4. バリエーション (A,B,C...順に昇順)
      ];
      varMasterSheet.getRange(DATA_START_ROW, 1, newRows.length, newRows[0].length).sort(sortSpecs);
    }

    SpreadsheetApp.getUi().alert(`更新完了: ${newRows.length}件\n反映しました。`);

  } catch (e) {
    // エラーはここで確実にキャッチして表示します
    SpreadsheetApp.getUi().alert("【処理中断】\n原因: " + e.message);
  }
}

// 補助ツール（列番号アルファベット変換）
function getColumnLetter(col) {
  let letter = "";
  while (col > 0) {
    let t = (col - 1) % 26;
    letter = String.fromCharCode(65 + t) + letter;
    col = (col - t) / 26 | 0;
  }
  return letter;
}

// 補助ツール（画像URL変換）
function getDirectImageUrl(url) {
  if (!url) return "";
  const idMatch = url.match(/[-\w]{25,}/);
  return idMatch ? "https://lh3.googleusercontent.com/d/" + idMatch[0] : url;
}