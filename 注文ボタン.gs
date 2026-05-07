/**
 * 注文マトリックスの入力データ（注文・入荷・不良・入庫）を確定し、
 * SKUに加算 ＆ 注文ログに記録するスクリプト
 * ★アキラさん流・誤爆防止ポップアップ ＆ AV列対応版
 */

const SUBMIT_ORDER_CONFIG = {
  SHEET_ORDER: "注文",      
  SHEET_SKU: "SKU",         
  SHEET_LOG: "注文ログ", 
  HEADER_ROW: 6,
  DATA_START_ROW: 7,
  
  BLOCK_ORDER: 27,       // AA列
  BLOCK_ARRIVE: 34,      // AH列
  BLOCK_DEFECT: 41,      // AO列
  BLOCK_WAREHOUSE_IN: 48, // AV列
  
  SIZE_ORDER: ["XS", "S", "M", "L", "XL", "F"]
};

function submitOrderData() {
  // --- 1. 誤爆防止の確認メッセージ ---
  const response = Browser.msgBox(
    "「注文～入荷」確定の確認", "【注意！】これは「注文」シートのデータを確定ボタンです！\\n入荷～検品→初期不良確定まで押さないでください。数値が消えてしまいます。\\n数量決定したら押してください。\\n\\n本当によろしいですか？", 
    Browser.Buttons.OK_CANCEL
  );

  if (response !== "ok") {
    return; // キャンセルならここで終了
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(SUBMIT_ORDER_CONFIG.SHEET_ORDER);
  const skuSheet = ss.getSheetByName(SUBMIT_ORDER_CONFIG.SHEET_SKU);
  const logSheet = ss.getSheetByName(SUBMIT_ORDER_CONFIG.SHEET_LOG);
  
  if (!orderSheet || !skuSheet || !logSheet) {
    Browser.msgBox("エラー：シートが見つかりません。");
    return;
  }

  const normalizeId = (h) => String(h).trim().replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/＿/g, "_");

  const getColMap = (sheet) => {
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(SUBMIT_ORDER_CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
    const map = {};
    headers.forEach((h, idx) => {
      const match = normalizeId(h).match(/^(\d{2,3})_/);
      if (match) map[match[1]] = idx + 1;
    });
    return map;
  };

  const orderColMap = getColMap(orderSheet);
  const skuColMap = getColMap(skuSheet);
  const logColMap = getColMap(logSheet);
  
  if (!skuColMap["061"] || !skuColMap["47"] || !skuColMap["48"] || !skuColMap["49"] || !skuColMap["50"]) {
    Browser.msgBox("エラー：SKUシートに必須IDが見つかりません。");
    return;
  }

  const skuLastRow = skuSheet.getLastRow();
  const skuLastCol = skuSheet.getLastColumn();
  if (skuLastRow < SUBMIT_ORDER_CONFIG.DATA_START_ROW) return;
  const skuValues = skuSheet.getRange(SUBMIT_ORDER_CONFIG.DATA_START_ROW, 1, skuLastRow - SUBMIT_ORDER_CONFIG.DATA_START_ROW + 1, skuLastCol).getValues(); 

  const skuCodeToIndex = new Map();
  skuValues.forEach((row, idx) => {
    const code = String(row[skuColMap["061"] - 1]).trim();
    if (code) skuCodeToIndex.set(code, idx);
  });

  const orderLastRow = orderSheet.getLastRow();
  const orderLastCol = Math.max(orderSheet.getLastColumn(), 61);
  if (orderLastRow < SUBMIT_ORDER_CONFIG.DATA_START_ROW) return;
  const orderValues = orderSheet.getRange(SUBMIT_ORDER_CONFIG.DATA_START_ROW, 1, orderLastRow - SUBMIT_ORDER_CONFIG.DATA_START_ROW + 1, orderLastCol).getValues();

  let logsToAppend = [];
  let updatedCount = 0;
  let rangesToClear = []; 
  const logMaxCol = logSheet.getLastColumn();

  for (let i = 0; i < orderValues.length; i++) {
    const row = orderValues[i];
    const tagCode = String(row[orderColMap["062"] - 1] || "").trim(); 
    const supplierRaw = String(row[orderColMap["01"] - 1] || "").trim(); 
    if (!tagCode) continue;

    const suppCode = supplierRaw.split(/[:：,、\s]+/)[0].trim().toUpperCase().match(/[A-Z]{2,3}/i)?.[0] || supplierRaw.toUpperCase();

    for (let sIdx = 0; sIdx < SUBMIT_ORDER_CONFIG.SIZE_ORDER.length; sIdx++) {
      const size = SUBMIT_ORDER_CONFIG.SIZE_ORDER[sIdx];
      const searchKey = `${tagCode}-${size}-${suppCode}`; 
      
      const v = {
        order: row[SUBMIT_ORDER_CONFIG.BLOCK_ORDER - 1 + sIdx],
        arrive: row[SUBMIT_ORDER_CONFIG.BLOCK_ARRIVE - 1 + sIdx],
        defect: row[SUBMIT_ORDER_CONFIG.BLOCK_DEFECT - 1 + sIdx],
        in: row[SUBMIT_ORDER_CONFIG.BLOCK_WAREHOUSE_IN - 1 + sIdx]
      };

      const hasVal = (val) => val !== "" && !isNaN(val) && Number(val) !== 0;

      if (!hasVal(v.order) && !hasVal(v.arrive) && !hasVal(v.defect) && !hasVal(v.in)) continue;

      if (skuCodeToIndex.has(searchKey)) {
        const skuIdx = skuCodeToIndex.get(searchKey); 
        const skuRowData = skuValues[skuIdx];
        const logBase = {
          sku: searchKey, size: size, 
          ttName: String(skuRowData[skuColMap["17"] - 1] || ""),
          engName: String(skuRowData[skuColMap["111"] - 1] || ""), 
          jpName: String(skuRowData[skuColMap["121"] - 1] || ""),
          snapshot: Number(skuRowData[skuColMap["52"] - 1]) || 0   
        };

        const process = (val, skuId, logType, logCol, blockStart) => {
          if (!hasVal(val)) return;
          const qty = Number(val);
          skuValues[skuIdx][skuColMap[skuId] - 1] = (Number(skuRowData[skuColMap[skuId] - 1]) || 0) + qty;
          logsToAppend.push(createLogArray(logColMap, logMaxCol, { ...logBase, type: logType, val: qty, targetCol: logCol }));
          rangesToClear.push(orderSheet.getRange(i + SUBMIT_ORDER_CONFIG.DATA_START_ROW, blockStart + sIdx).getA1Notation());
          updatedCount++;
        };

        process(v.order, "47", "注文", "30", SUBMIT_ORDER_CONFIG.BLOCK_ORDER);
        process(v.arrive, "48", "入荷", "31", SUBMIT_ORDER_CONFIG.BLOCK_ARRIVE);
        process(v.defect, "49", "不良", "32", SUBMIT_ORDER_CONFIG.BLOCK_DEFECT);
        process(v.in, "50", "入庫", "85", SUBMIT_ORDER_CONFIG.BLOCK_WAREHOUSE_IN);
      }
    }
  }

  if (updatedCount > 0) {
    ["47", "48", "49", "50"].forEach(id => {
      const colIdx = skuColMap[id] - 1;
      const data = skuValues.map(r => [r[colIdx]]);
      skuSheet.getRange(SUBMIT_ORDER_CONFIG.DATA_START_ROW, skuColMap[id], data.length, 1).setValues(data);
    });

    if (logsToAppend.length > 0) logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToAppend.length, logMaxCol).setValues(logsToAppend);
    if (rangesToClear.length > 0) orderSheet.getRangeList(rangesToClear).clearContent();

    Browser.msgBox("確定完了！");
  } else {
    Browser.msgBox("確定するデータがありませんでした。");
  }
}

// ログ配列作成関数（関数の外に独立させています）
function createLogArray(logColMap, maxCol, p) {
  let row = new Array(maxCol).fill("");
  const setV = (id, val) => { if (logColMap[id]) row[logColMap[id] - 1] = val; };
  setV("81", new Date()); setV("33", p.type); setV("061", p.sku); setV("09", p.size);
  setV("17", p.ttName); setV("111", p.engName); setV("121", p.jpName); setV("88", p.snapshot);
  if (p.targetCol) setV(p.targetCol, p.val);
  return row;
}