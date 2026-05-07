/**
 * 移動販売マトリックスの入力データ（取寄せ・販売・不良・棚卸）を確定し、
 * SKUシートの各項目（58, 59, 60, 70, 34等）へ完全連動させるスクリプト
 */

function submitMobileSalesData() {
  const response = Browser.msgBox("「移動販売」数量確定の確認", "【注意！】これは「移動販売」シートのデータを確定ボタンです！\\確定すると数値が消えます。\\n\\n本当によろしいですか？", Browser.Buttons.OK_CANCEL);
  if (response !== "ok") return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const whSheet = ss.getSheetByName("移動販売");
  const skuSheet = ss.getSheetByName("SKU");
  const logSheet = ss.getSheetByName("移動販売ログ"); // ※ログシート名に合わせてください
  
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
  
  // ★更新対象のSKU列IDリスト
  const ID_WAREHOUSE = "34"; // 倉庫の在庫（※もしIDが34でなければ変更してください）
  const ID_MOBILE_STOCK = "60"; // 移動販売 在庫数
  const ID_MOBILE_PULL = "58";  // 累計取寄せ数
  const ID_MOBILE_SALES = "59"; // 累計販売数
  const ID_MOBILE_DEFECT = "70";// 累計不良/処分数

  const skuLastRow = skuSheet.getLastRow();
  const skuValues = skuSheet.getRange(7, 1, skuLastRow - 6, skuSheet.getLastColumn()).getValues(); 

  const skuCodeToIndex = new Map();
  skuValues.forEach((row, idx) => {
    const code = String(row[skuColMap["061"] - 1]).trim();
    if (code) skuCodeToIndex.set(code, idx);
  });

  const whLastRow = whSheet.getLastRow();
  const whValues = whSheet.getRange(7, 1, whLastRow - 6, 68).getValues();

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
        mobileStock: row[34 - 1 + sIdx], // AH列(販売在庫・システム値)
        pull: row[41 - 1 + sIdx],        // AO列(取寄せ)
        sales: row[48 - 1 + sIdx],       // AV列(販売)
        defect: row[55 - 1 + sIdx],      // BC列(不良)
        check: row[62 - 1 + sIdx]        // BJ列(棚卸)
      };

      const hasVal = (val) => val !== "" && !isNaN(val);
      if (!hasVal(v.pull) && !hasVal(v.sales) && !hasVal(v.defect) && !hasVal(v.check)) continue;

      if (skuCodeToIndex.has(searchKey)) {
        const skuIdx = skuCodeToIndex.get(searchKey); 
        const skuRowData = skuValues[skuIdx];
        const sysMobileStock = Number(v.mobileStock) || 0; 

        const logBase = {
          sku: searchKey, size: size, 
          ttName: String(skuRowData[skuColMap["17"] - 1] || ""),
          engName: String(skuRowData[skuColMap["111"] - 1] || ""), 
          jpName: String(skuRowData[skuColMap["121"] - 1] || ""),
          snapshot: sysMobileStock 
        };

        const getSkuVal = (id) => Number(skuValues[skuIdx][skuColMap[id] - 1]) || 0;
        const setSkuVal = (id, val) => skuValues[skuIdx][skuColMap[id] - 1] = val;

        // B3: 取寄せ処理 (移動在庫に＋, 累計に＋, 倉庫在庫から−)
        if (hasVal(v.pull) && Number(v.pull) !== 0) {
          const qty = Number(v.pull);
          setSkuVal(ID_MOBILE_STOCK, getSkuVal(ID_MOBILE_STOCK) + qty);
          setSkuVal(ID_MOBILE_PULL, getSkuVal(ID_MOBILE_PULL) + qty);
          if(skuColMap[ID_WAREHOUSE]) setSkuVal(ID_WAREHOUSE, getSkuVal(ID_WAREHOUSE) - qty); // 倉庫から引く

          logsToAppend.push(createLogArray(logColMap, logMaxCol, { ...logBase, type: "移動_取寄せ", val: qty }));
          rangesToClear.push(whSheet.getRange(i + 7, 41 + sIdx).getA1Notation());
          updatedCount++;
        }

        // B4: 販売処理 (移動在庫から−, 累計に＋)
        if (hasVal(v.sales) && Number(v.sales) !== 0) {
          const qty = Number(v.sales);
          setSkuVal(ID_MOBILE_STOCK, getSkuVal(ID_MOBILE_STOCK) - qty);
          setSkuVal(ID_MOBILE_SALES, getSkuVal(ID_MOBILE_SALES) + qty);

          logsToAppend.push(createLogArray(logColMap, logMaxCol, { ...logBase, type: "移動_販売", val: qty }));
          rangesToClear.push(whSheet.getRange(i + 7, 48 + sIdx).getA1Notation());
          updatedCount++;
        }

        // B5: 不良処理 (移動在庫から−, 累計に＋)
        if (hasVal(v.defect) && Number(v.defect) !== 0) {
          const qty = Number(v.defect);
          setSkuVal(ID_MOBILE_STOCK, getSkuVal(ID_MOBILE_STOCK) - qty);
          setSkuVal(ID_MOBILE_DEFECT, getSkuVal(ID_MOBILE_DEFECT) + qty);

          logsToAppend.push(createLogArray(logColMap, logMaxCol, { ...logBase, type: "移動_不良", val: qty }));
          rangesToClear.push(whSheet.getRange(i + 7, 55 + sIdx).getA1Notation());
          updatedCount++;
        }

        // B6: 棚卸処理 (差分を移動在庫に足す)
        if (hasVal(v.check)) {
          const actualQty = Number(v.check);
          const diff = actualQty - sysMobileStock; 

          if (diff !== 0) {
            setSkuVal(ID_MOBILE_STOCK, getSkuVal(ID_MOBILE_STOCK) + diff); 
            logsToAppend.push(createLogArray(logColMap, logMaxCol, { 
              ...logBase, 
              type: "移動_棚卸不一致", 
              val: diff,          
              flag: "TRUE"        
            }));
            updatedCount++;
          }
          rangesToClear.push(whSheet.getRange(i + 7, 62 + sIdx).getA1Notation());
        }
      }
    }
  }

  if (updatedCount > 0) {
    // 変更したSKU列をシートへ書き戻す
    [ID_WAREHOUSE, ID_MOBILE_STOCK, ID_MOBILE_PULL, ID_MOBILE_SALES, ID_MOBILE_DEFECT].forEach(id => {
      if(skuColMap[id]){
        const colIdx = skuColMap[id] - 1;
        const data = skuValues.map(r => [r[colIdx]]);
        skuSheet.getRange(7, skuColMap[id], data.length, 1).setValues(data);
      }
    });

    if (logsToAppend.length > 0) logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToAppend.length, logMaxCol).setValues(logsToAppend);
    if (rangesToClear.length > 0) whSheet.getRangeList(rangesToClear).clearContent();

    Browser.msgBox("移動販売データの確定＆SKU連動が完了しました！");
  } else {
    Browser.msgBox("確定するデータがありませんでした。");
  }
}

function createLogArray(logColMap, maxCol, p) {
  let row = new Array(maxCol).fill("");
  const setV = (id, val) => { if (logColMap[id]) row[logColMap[id] - 1] = val; };
  setV("81", new Date()); setV("33", p.type); // 操作種別
  setV("061", p.sku); setV("09", p.size);
  setV("17", p.ttName); setV("111", p.engName); setV("121", p.jpName); 
  setV("88", p.snapshot);
  setV("84", p.val); // 数量記録列（用途に合わせて変更してください）
  if (p.flag) setV("86", p.flag); 
  return row;
}