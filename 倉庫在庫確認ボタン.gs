/**
 * 倉庫マトリックスの入力データ（入庫・出庫・不良・棚卸）を確定し、
 * SKUに加算・減算＆ 倉庫ログに記録するスクリプト
 */

function submitWarehouseDataV22() {
  const response = Browser.msgBox("「倉庫」数量確定の確認", "【注意！】これは「倉庫」シートのデータを確定ボタンです！\\nSKUシートへ反映し、入力欄をクリアします。\\n本当によろしいですか？", Browser.Buttons.OK_CANCEL);
  if (response !== "ok") return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const whSheet = ss.getSheetByName("倉庫");
  const skuSheet = ss.getSheetByName("SKU");
  const logSheet = ss.getSheetByName("倉庫ログ"); // ※ログシート名に合わせてください
  
  if (!whSheet || !skuSheet || !logSheet) return Browser.msgBox("シートが見つかりません。");

  const normalizeId = (h) => String(h).trim().replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/＿/g, "_");
  const getColMap = (sheet) => {
    const map = {};
    sheet.getRange(6, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].forEach((h, idx) => {
      const match = normalizeId(h).match(/^(\d{2,3})_/);
      if (match) map[match[1]] = idx + 1;
    });
    return map;
  };

  const whColMap = getColMap(whSheet);
  const skuColMap = getColMap(skuSheet);
  const logColMap = getColMap(logSheet);
  
  const skuLastRow = skuSheet.getLastRow();
  const skuValues = skuSheet.getRange(7, 1, skuLastRow - 6, skuSheet.getLastColumn()).getValues(); 

  const skuCodeToIndex = new Map();
  skuValues.forEach((row, idx) => {
    const code = String(row[skuColMap["061"] - 1]).trim();
    if (code) skuCodeToIndex.set(code, idx);
  });

  const whLastRow = whSheet.getLastRow();
  const whValues = whSheet.getRange(7, 1, whLastRow - 6, 61).getValues();

  let logsToAppend = [];
  let updatedCount = 0;
  let rangesToClear = []; 
  const logMaxCol = logSheet.getLastColumn();
  const SIZE_ORDER = ["XS", "S", "M", "L", "XL", "F"];

  for (let i = 0; i < whValues.length; i++) {
    const row = whValues[i];
    const tagCode = String(row[whColMap["062"] - 1] || "").trim(); 
    const supplierRaw = String(row[whColMap["01"] - 1] || "").trim(); 
    if (!tagCode) continue;

    const suppCode = supplierRaw.split(/[:：,、\s]+/)[0].trim().toUpperCase().match(/[A-Z]{2,3}/i)?.[0] || supplierRaw.toUpperCase();

    for (let sIdx = 0; sIdx < SIZE_ORDER.length; sIdx++) {
      const size = SIZE_ORDER[sIdx];
      const searchKey = `${tagCode}-${size}-${suppCode}`; 
      
      const v = {
        stock: row[27 - 1 + sIdx], // AA列(システム在庫)
        in: row[34 - 1 + sIdx],    // AH列(入庫)
        out: row[41 - 1 + sIdx],   // AO列(出庫)
        defect: row[48 - 1 + sIdx], // AV列(不良)
        check: row[55 - 1 + sIdx]  // BC列(棚卸)
      };

      const hasVal = (val) => val !== "" && !isNaN(val);

      if (!hasVal(v.in) && !hasVal(v.out) && !hasVal(v.defect) && !hasVal(v.check)) continue;

      if (skuCodeToIndex.has(searchKey)) {
        const skuIdx = skuCodeToIndex.get(searchKey); 
        const skuRowData = skuValues[skuIdx];
        const sysStock = Number(v.stock) || 0; // マトリックス上の実行時在庫

        const logBase = {
          sku: searchKey, size: size, 
          ttName: String(skuRowData[skuColMap["17"] - 1] || ""),
          engName: String(skuRowData[skuColMap["111"] - 1] || ""), 
          jpName: String(skuRowData[skuColMap["121"] - 1] || ""),
          snapshot: sysStock // 88_実行時在庫
        };

        // 1. 入庫処理
        if (hasVal(v.in) && Number(v.in) !== 0) {
          const qty = Number(v.in);
          skuValues[skuIdx][skuColMap["50"] - 1] = (Number(skuRowData[skuColMap["50"] - 1]) || 0) + qty;
          skuValues[skuIdx][skuColMap["52"] - 1] = (Number(skuRowData[skuColMap["52"] - 1]) || 0) + qty; // 在庫も増やす
          logsToAppend.push(createLogArray(logColMap, logMaxCol, { ...logBase, type: "入庫", val: qty, targetCol: "85" }));
          rangesToClear.push(whSheet.getRange(i + 7, 34 + sIdx).getA1Notation());
          updatedCount++;
        }

        // 2. 出庫処理
        if (hasVal(v.out) && Number(v.out) !== 0) {
          const qty = Number(v.out);
          skuValues[skuIdx][skuColMap["52"] - 1] = (Number(skuRowData[skuColMap["52"] - 1]) || 0) - qty; // 在庫から引く
          logsToAppend.push(createLogArray(logColMap, logMaxCol, { ...logBase, type: "出庫", val: qty, targetCol: "89" }));
          rangesToClear.push(whSheet.getRange(i + 7, 41 + sIdx).getA1Notation());
          updatedCount++;
        }

        // 3. 不良処理
        if (hasVal(v.defect) && Number(v.defect) !== 0) {
          const qty = Number(v.defect);
          skuValues[skuIdx][skuColMap["52"] - 1] = (Number(skuRowData[skuColMap["52"] - 1]) || 0) - qty; // 在庫から引く
          logsToAppend.push(createLogArray(logColMap, logMaxCol, { ...logBase, type: "不良", val: qty, targetCol: "84" }));
          rangesToClear.push(whSheet.getRange(i + 7, 48 + sIdx).getA1Notation());
          updatedCount++;
        }

        // 4. ★棚卸処理（アキラさん特製ロジック）
        if (hasVal(v.check)) {
          const actualQty = Number(v.check);
          const diff = actualQty - sysStock; // 差分計算（例: 実際8 - 帳簿10 = -2）

          if (diff !== 0) {
            // 不一致の場合：在庫を強制的にactualQtyに合わせる（差分を足す）
            skuValues[skuIdx][skuColMap["52"] - 1] = (Number(skuRowData[skuColMap["52"] - 1]) || 0) + diff; 
            logsToAppend.push(createLogArray(logColMap, logMaxCol, { 
              ...logBase, 
              type: "棚卸(不一致)", 
              val: diff,          // 差分（-2など）をそのまま記録
              targetCol: "84",    // 不良枠にマイナスとして入れる
              flag: "TRUE"        // 86_不一致フラグを立てる！
            }));
            updatedCount++;
          }
          rangesToClear.push(whSheet.getRange(i + 7, 55 + sIdx).getA1Notation());
        }
      }
    }
  }

  if (updatedCount > 0) {
    // 変更があった列（50:入庫, 52:実行時在庫など）をSKUシートへ書き戻す処理（※列IDはSKUの実態に合わせて調整）
    ["50", "52"].forEach(id => {
      if(skuColMap[id]){
        const colIdx = skuColMap[id] - 1;
        const data = skuValues.map(r => [r[colIdx]]);
        skuSheet.getRange(7, skuColMap[id], data.length, 1).setValues(data);
      }
    });

    if (logsToAppend.length > 0) logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToAppend.length, logMaxCol).setValues(logsToAppend);
    if (rangesToClear.length > 0) whSheet.getRangeList(rangesToClear).clearContent();

    Browser.msgBox("倉庫データの確定＆ログ記録が完了しました！");
  } else {
    Browser.msgBox("確定するデータがありませんでした。");
  }
}

function createLogArray(logColMap, maxCol, p) {
  let row = new Array(maxCol).fill("");
  const setV = (id, val) => { if (logColMap[id]) row[logColMap[id] - 1] = val; };
  setV("81", new Date()); setV("33", p.type); // ※操作種別が33か82か、ログシートのIDに合わせてください
  setV("061", p.sku); setV("09", p.size);
  setV("17", p.ttName); setV("111", p.engName); setV("121", p.jpName); 
  setV("88", p.snapshot);
  if (p.targetCol) setV(p.targetCol, p.val);
  if (p.flag) setV("86", p.flag); // 不一致フラグ
  return row;
}