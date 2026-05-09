/*******************************************************
 * 棚卸反映ボタン V1
 *
 * 対象：
 * - 棚卸入力欄のみ
 *
 * 動作：
 * 倉庫マトリックスの棚卸入力値 と
 * SKUシート現在在庫 を比較し、
 * 差分のみをSKUへ反映する
 *
 * ログ：
 * SKUログへ記録
 *
 * 処理番号：
 * IV-YYYYMMDD-0001
 *******************************************************/

function submitInventoryCheckV1() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const SHEET_WAREHOUSE = "倉庫";
  const SHEET_SKU = "SKU";
  const SHEET_SKU_LOG = "SKUログ";

  const HEADER_ROW = 6;
  const DATA_START_ROW = 7;

  const PROCESS_SOURCE = "IV";
  const ADJUSTMENT_TYPE = "棚卸調整";

  const SIZE_LIST = ["XS", "S", "M", "L", "XL", "F"];

  // 棚卸入力列
  const INVENTORY_COLS = {
    XS: 55,
    S: 56,
    M: 57,
    L: 58,
    XL: 59,
    F: 60
  };

  const whSheet = ss.getSheetByName(SHEET_WAREHOUSE);
  const skuSheet = ss.getSheetByName(SHEET_SKU);
  const logSheet = ss.getSheetByName(SHEET_SKU_LOG);

  const result = Browser.msgBox(
    "棚卸反映",
    "棚卸入力値をSKU在庫へ反映します。\n実行しますか？",
    Browser.Buttons.YES_NO
  );

  if (result !== "yes") return;

  const whColMap = getColumnMapByItemId_(whSheet, HEADER_ROW);
  const skuColMap = getColumnMapByItemId_(skuSheet, HEADER_ROW);
  const logColMap = getColumnMapByItemId_(logSheet, HEADER_ROW);

  const whValues = whSheet.getRange(
    DATA_START_ROW,
    1,
    whSheet.getLastRow() - DATA_START_ROW + 1,
    whSheet.getLastColumn()
  ).getValues();

  const skuValues = skuSheet.getRange(
    DATA_START_ROW,
    1,
    skuSheet.getLastRow() - DATA_START_ROW + 1,
    skuSheet.getLastColumn()
  ).getValues();

  // SKUマップ
  const skuMap = new Map();

  skuValues.forEach((row, i) => {

    const skuCode = String(
      row[skuColMap["061"] - 1] || ""
    ).trim();

    if (!skuCode) return;

    skuMap.set(skuCode, {
      rowNumber: DATA_START_ROW + i,
      stock: Number(
        row[skuColMap["50"] - 1] || 0
      )
    });

  });

  const processId = createProcessId_(
    logSheet,
    logColMap["01"],
    PROCESS_SOURCE
  );

  const now = new Date();

  const logRows = [];
  const clearRanges = [];
  const stockUpdates = [];

  whValues.forEach((row, rowIndex) => {

    const sheetRow = DATA_START_ROW + rowIndex;

    const code064 = String(
      row[whColMap["064"] - 1] || ""
    ).trim();

    const displayName = String(
      row[whColMap["17"] - 1] || ""
    ).trim();

    if (!code064) return;

    SIZE_LIST.forEach(size => {

      const invCol = INVENTORY_COLS[size];

      const inventoryValue = row[invCol - 1];

      // 空白なら無視
      if (
        inventoryValue === "" ||
        inventoryValue === null
      ) {
        return;
      }

      const inventoryQty = Number(inventoryValue);

      if (isNaN(inventoryQty)) return;

      const skuCode = `${code064}-${size}`;

      const skuInfo = skuMap.get(skuCode);

      // SKU存在しない場合
      if (!skuInfo) return;

      const currentStock = Number(skuInfo.stock || 0);

      // 差分
      const diffQty = inventoryQty - currentStock;

      // 差分なし
      if (diffQty === 0) return;

      const afterStock = inventoryQty;

      // SKU更新値保存
      skuInfo.stock = afterStock;

      stockUpdates.push({
        rowNumber: skuInfo.rowNumber,
        value: afterStock
      });

      // ログ区分
      const processType =
        diffQty > 0 ? "入庫" : "出庫";

      // ログ追加
      logRows.push(
        buildSkuLogRow_(logColMap, {
          processId,
          dateTime: now,
          source: PROCESS_SOURCE,
          processType,
          adjustmentType: ADJUSTMENT_TYPE,
          skuCode,
          code064,
          size,
          displayName,
          changeQty: diffQty,
          staff: "",
          memo: ""
        })
      );

      // 棚卸入力欄クリア
      clearRanges.push(
        whSheet.getRange(sheetRow, invCol)
      );

    });

  });

  // 更新なし
  if (logRows.length === 0) {

    Browser.msgBox(
      "棚卸差分はありませんでした。"
    );

    return;
  }

  // SKU更新
  stockUpdates.forEach(update => {

    skuSheet
      .getRange(
        update.rowNumber,
        skuColMap["50"]
      )
      .setValue(update.value);

  });

  // SKUログ
  const logStartRow = getNextDataRow_(
    logSheet,
    DATA_START_ROW
  );

  logSheet.getRange(
    logStartRow,
    1,
    logRows.length,
    logSheet.getLastColumn()
  ).setValues(logRows);

  // 棚卸入力クリア
  clearRanges.forEach(rng => {
    rng.clearContent();
  });

  Browser.msgBox(
    "棚卸反映完了！\n\n" +
    "処理番号：" + processId + "\n" +
    "反映件数：" + logRows.length + "件"
  );

}